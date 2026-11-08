using System.Diagnostics;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using SoftwareInstaller.Packaging;

namespace SoftwareInstaller.Tests;

internal static partial class Program
{
    private static async Task TestRunnerPackagingAsync(string template)
    {
        RunnerTemplate.Validate(template);
        Assert(FileVersionInfo.GetVersionInfo(template).ProductVersion == RunnerTemplate.ExpectedRunnerVersion,
            "the generator accepts the lightweight Framework runner profile");
        await Reject(() => { RunnerTemplate.Validate(Environment.ProcessPath!); return Task.CompletedTask; },
            "a legacy or unrelated runtime template cannot silently produce a slow bundle");
        byte[] data = [77, 90, 1, 2, 3];
        var installer = Path.Combine(root, "setup.exe"); File.WriteAllBytes(installer, data);
        var item = new SoftwareInstaller.Models.SoftwareItem { Name = "打包开发样例", Version = "", Size = "", Description = "", Category = "测试", FilePath = installer };
        var output = Path.Combine(root, "packaging-fixture.exe");
        await BundleBuilder.BuildAsync(new BundleBuildRequest { Name = "打包检查（不安装）", RunnerPath = template, OutputPath = output,
            Software = [item], Schemes = new() { ["当前方案"] = [item.Id] } });
        using var package = await BundleReader.OpenAsync(output);
        var work = Path.Combine(root, "extracted"); await package.ExtractAsync(package.Manifest.Software, work);
        Assert(File.ReadAllBytes(Path.Combine(work, item.Id, "setup.exe")).SequenceEqual(data),
            "the real builder and Framework template produce readable resources with matching hashes");
        using var input = File.OpenRead(output); input.Position = input.Length - BundleReader.FooterSize;
        var footer = new byte[BundleReader.FooterSize]; input.ReadExactly(footer);
        long length = BitConverter.ToInt64(footer, 0); input.Position = input.Length - BundleReader.FooterSize - length;
        using var payload = new MemoryStream(); input.CopyTo(payload); payload.SetLength(length); payload.Position = 0;
        using var zip = new ZipArchive(payload, ZipArchiveMode.Read);
        var resource = zip.GetEntry("content/" + package.Manifest.Software[0].Files[0].Sha256)!;
        Assert(resource.CompressedLength == resource.Length, "installer bytes are stored without secondary compression");
    }

    // Explicit development-only comparison; uses synthetic resources and never launches an installer.
    private static async Task MeasureRunnerStartupAsync(string before, string after)
    {
        byte[] installer = [77, 90, 1, 2, 3];
        string hash = Convert.ToHexString(SHA256.HashData(installer));
        var software = Enumerable.Range(1, 120).Select(i => new BundleSoftware
        {
            Id = Guid.NewGuid().ToString("N"), Name = $"启动测量 {i}", Installer = "setup.exe",
            Files = Enumerable.Range(0, 32).Select(j => new BundleFile
            { Path = j == 0 ? "setup.exe" : $"data/{j}.bin", Sha256 = hash, Size = installer.Length }).ToList()
        }).ToList();
        var manifest = new BundleManifest { Name = "开发启动测量（不安装）", Software = software,
            Schemes = new() { ["当前方案"] = software.Select(s => s.Id).ToList() } };
        var payloadPath = Path.Combine(root, "startup-payload.zip");
        using (var payload = File.Create(payloadPath))
        using (var zip = new ZipArchive(payload, ZipArchiveMode.Create))
        {
            using (var content = zip.CreateEntry("content/" + hash, CompressionLevel.NoCompression).Open()) content.Write(installer);
            // Match the substantial overlay that triggers scanning of real generated executables.
            using (var padding = zip.CreateEntry("benchmark-padding", CompressionLevel.NoCompression).Open())
            {
                var block = new byte[1024 * 1024]; new Random(12345).NextBytes(block);
                for (int i = 0; i < 64; i++) padding.Write(block);
            }
            using var metadata = zip.CreateEntry("manifest.json", CompressionLevel.Fastest).Open();
            JsonSerializer.Serialize(metadata, manifest);
        }
        var measurements = new List<object>();
        foreach (var (name, template) in new[] { ("before", before), ("after", after) })
        {
            var directory = Path.Combine(root, name); Directory.CreateDirectory(directory);
            var bundle = Path.Combine(directory, "软件集合.exe");
            File.Copy(template, bundle);
            await using (var output = new FileStream(bundle, FileMode.Append))
            {
                await using var payload = File.OpenRead(payloadPath);
                await payload.CopyToAsync(output);
                await output.WriteAsync(BitConverter.GetBytes(payload.Length));
                await output.WriteAsync(new byte[32]); await output.WriteAsync(BundleReader.Magic);
            }
            var extraction = Path.Combine(directory, "runtime-cache");
            for (int run = 0; run < 3; run++)
            {
                var report = Path.Combine(directory, $"startup-{run}.json");
                var info = new ProcessStartInfo(bundle, $"--measure-startup \"{report}\"")
                { UseShellExecute = false, CreateNoWindow = true, WindowStyle = ProcessWindowStyle.Hidden };
                info.Environment["DOTNET_BUNDLE_EXTRACT_BASE_DIR"] = extraction;
                var watch = Stopwatch.StartNew();
                using var process = Process.Start(info)!;
                using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(75));
                try { await process.WaitForExitAsync(deadline.Token); }
                catch { process.Kill(entireProcessTree: true); throw; }
                watch.Stop();
                Assert(process.ExitCode == 0 && File.Exists(report), $"{name} startup run {run + 1} loads the real software list");
                using var json = JsonDocument.Parse(File.ReadAllText(report));
                var data = json.RootElement;
                Assert(data.GetProperty("ready").GetBoolean(), "startup measurement reached a usable install button");
                measurements.Add(new { name, run = run + 1, firstWithEmptyRuntimeCache = run == 0,
                    elapsedMs = watch.Elapsed.TotalMilliseconds, readyAfterMainMs = data.GetProperty("readyAfterMainMs").GetDouble(),
                    workingSetBytes = data.GetProperty("workingSetBytes").GetInt64(), privateBytes = data.GetProperty("privateBytes").GetInt64(),
                    templateBytes = new FileInfo(template).Length,
                    runtimeExtractedBytes = Directory.Exists(extraction) ? Directory.EnumerateFiles(extraction, "*", SearchOption.AllDirectories).Sum(f => new FileInfo(f).Length) : 0 });
            }
        }
        var result = Path.Combine(root, "startup-comparison.json");
        File.WriteAllText(result, JsonSerializer.Serialize(new { os = Environment.OSVersion.ToString(),
            caveat = "First run has an empty private runtime cache; OS disk cache is not flushed. No Windows 7 or slow-PC measurements.",
            softwareCount = software.Count, resourceReferences = software.Sum(s => s.Files.Count), payloadBytes = new FileInfo(payloadPath).Length,
            measurements }, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine($"Startup comparison: {result}");
    }
}
