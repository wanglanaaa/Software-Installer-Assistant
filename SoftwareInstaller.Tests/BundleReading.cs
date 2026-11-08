using System.IO.Compression;
using System.Text.Json;
using System.Security.Cryptography;
using SoftwareInstaller.Packaging;

namespace SoftwareInstaller.Tests;

internal static partial class Program
{
    private sealed class BundleUpdates(List<BundleProgress> items) : IProgress<BundleProgress>
    {
        public void Report(BundleProgress value) => items.Add(value);
    }

    private static async Task TestBundleReadingAsync()
    {
        // Synthetic resources are never executed. The legacy whole-payload digest is deliberately unused.
        var data = new byte[] { 77, 90, 1, 2, 3 };
        var resource = new BundleFile { Path = "setup.exe", Sha256 = Convert.ToHexString(SHA256.HashData(data)), Size = data.Length };
        var item = new BundleSoftware { Id = Guid.NewGuid().ToString("N"), Name = "读取检查", Installer = resource.Path, Files = [resource] };
        var manifest = new BundleManifest { Software = [item], Schemes = new() { ["当前方案"] = [item.Id] } };

        async Task<string> WritePackageAsync(string name, bool includeResource = true)
        {
            using var payload = new MemoryStream();
            using (var zip = new ZipArchive(payload, ZipArchiveMode.Create, true))
            {
                if (includeResource)
                {
                    await using var content = zip.CreateEntry("content/" + resource.Sha256).Open();
                    await content.WriteAsync(data);
                }
                await using var metadata = zip.CreateEntry("manifest.json").Open();
                await JsonSerializer.SerializeAsync(metadata, manifest);
            }
            var path = Path.Combine(root, name + ".bundle");
            await using var output = File.Create(path);
            await output.WriteAsync(payload.ToArray());
            await output.WriteAsync(BitConverter.GetBytes(payload.Length));
            await output.WriteAsync(new byte[32]);
            await output.WriteAsync(BundleReader.Magic);
            return path;
        }

        var path = await WritePackageAsync("unchecked-digests");
        using (var package = await BundleReader.OpenAsync(path))
        {
            Assert(package.Manifest.Software.Single().Id == item.Id, "package opens without checking the stored payload digest");
            var destination = Path.Combine(root, "unchecked-extracted");
            var updates = new List<BundleProgress>();
            await package.ExtractAsync(package.Manifest.Software, destination, new BundleUpdates(updates));
            Assert(File.ReadAllBytes(Path.Combine(destination, item.Id, resource.Path)).SequenceEqual(data),
                "resources extract successfully after their SHA-256 is checked in the copy pass");
            Assert(updates[0].BytesCompleted == 0 && updates[^1].BytesCompleted == data.Length && updates[^1].TotalBytes == data.Length
                && updates[^1].Completed == updates[^1].Total, "extraction progress reports actual bytes and always includes completion");
            using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
            await Reject(() => package.ExtractAsync(package.Manifest.Software, Path.Combine(root, "cancelled-extraction"), token: cancelled.Token),
                "extraction still honors cancellation");
        }
        data[4] ^= 1; // Same length, different content: CRC and ZIP structure remain valid.
        var damaged = await WritePackageAsync("damaged-content");
        using (var package = await BundleReader.OpenAsync(damaged))
            await Reject(() => package.ExtractAsync(package.Manifest.Software, Path.Combine(root, "damaged-extraction")),
                "same-length corrupted resources cannot reach installation");
        data[4] ^= 1;
        var missing = await WritePackageAsync("missing-resource", includeResource: false);
        using (var package = await BundleReader.OpenAsync(missing))
            await Reject(() => package.ExtractAsync(package.Manifest.Software, Path.Combine(root, "missing-extraction")), "missing resources are rejected on demand");
        resource.Size++;
        var wrongLength = await WritePackageAsync("wrong-length");
        using (var package = await BundleReader.OpenAsync(wrongLength))
            await Reject(() => package.ExtractAsync(package.Manifest.Software, Path.Combine(root, "wrong-length-extraction")), "resource length mismatches are rejected on demand");
        resource.Size--;
        resource.Path = item.Installer = "../escape.exe";
        var traversal = await WritePackageAsync("invalid-path");
        await Reject(async () => { using var _ = await BundleReader.OpenAsync(traversal); }, "unsafe manifest paths are still rejected");
        using (var file = new FileStream(path, FileMode.Open, FileAccess.Write)) { file.Position = file.Length - 1; file.WriteByte(0); }
        await Reject(async () => { using var _ = await BundleReader.OpenAsync(path); }, "invalid package footers are still rejected");
    }
}
