using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using SoftwareInstaller.Models;

namespace SoftwareInstaller.Packaging;

public sealed class BundleBuildRequest
{
    public string Name { get; init; } = "软件合集";
    public string RunnerPath { get; init; } = "";
    public string OutputPath { get; init; } = "";
    public string DefaultScheme { get; init; } = "当前方案";
    public List<SoftwareItem> Software { get; init; } = new();
    public List<string> Categories { get; init; } = new();
    public Dictionary<string, List<string>> Schemes { get; init; } = new();
}

public static class BundleBuilder
{
    // Leave headroom below the Windows executable file size boundary.
    public const long MaxBundleBytes = 3500L * 1024 * 1024;
    public static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    public static List<int> ParseCodes(string value)
    {
        return value.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(s => int.Parse(s, System.Globalization.CultureInfo.InvariantCulture)).Distinct().ToList();
    }

    public static List<string> CheckItem(SoftwareItem item)
    {
        var errors = new List<string>();
        if (string.IsNullOrWhiteSpace(item.Name)) errors.Add("软件名称为空");
        if (!File.Exists(item.FilePath)) errors.Add("安装文件不存在");
        else
        {
            var extension = Path.GetExtension(item.FilePath);
            if (!extension.Equals(".exe", StringComparison.OrdinalIgnoreCase) && !extension.Equals(".msi", StringComparison.OrdinalIgnoreCase))
                errors.Add("首版仅支持 EXE / MSI");
            if (PackagePaths.IsReparsePoint(item.FilePath!)) errors.Add("安装文件不能是链接");
        }
        if (!string.IsNullOrWhiteSpace(item.ResourceDirectory))
        {
            if (!Directory.Exists(item.ResourceDirectory)) errors.Add("关联资源目录不存在");
            else if (File.Exists(item.FilePath))
            {
                var root = Path.GetFullPath(item.ResourceDirectory).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
                if (!Path.GetFullPath(item.FilePath!).StartsWith(root, StringComparison.OrdinalIgnoreCase))
                    errors.Add("安装文件必须位于关联资源目录内");
            }
        }
        if (item.TimeoutMinutes < 1 || item.TimeoutMinutes > 1440) errors.Add("超时必须在 1–1440 分钟内");
        try
        {
            if (ParseCodes(item.SuccessCodes).Count == 0) errors.Add("至少设置一个成功返回码");
            ParseCodes(item.RebootCodes);
        }
        catch (Exception ex) when (ex is FormatException or OverflowException) { errors.Add("返回码须为逗号分隔的整数"); }
        if (!string.IsNullOrWhiteSpace(item.DetectionVersion) &&
            (!Version.TryParse(item.DetectionVersion, out _) || string.IsNullOrWhiteSpace(item.DetectionFile)))
            errors.Add("文件版本检测需填写文件路径和有效版本号");
        if (!string.IsNullOrWhiteSpace(item.MsiProductCode) && !Guid.TryParse(item.MsiProductCode, out _))
            errors.Add("MSI 产品码格式错误");
        if (item.DependencyIds.Contains(item.Id)) errors.Add("软件不能依赖自身");
        try
        {
            InstallDirectoryOptions.BuildArguments(item);
            if (!string.IsNullOrWhiteSpace(item.DetectionFile)) InstallDirectoryOptions.ExpandDetectionPath(item.DetectionFile, item.DefaultInstallDirectory);
        }
        catch (InvalidDataException ex) { errors.Add(ex.Message); }
        return errors;
    }

    public static async Task BuildAsync(BundleBuildRequest request, IProgress<BundleProgress>? progress = null, CancellationToken token = default)
    {
        if (string.IsNullOrWhiteSpace(request.Name)) throw new InvalidDataException("请输入包名称。");
        if (request.Software.Count == 0) throw new InvalidDataException("没有可打包的软件。");
        if (request.Software.Select(s => s.Id).Distinct().Count() != request.Software.Count)
            throw new InvalidDataException("软件标识重复，请重新添加重复条目。");
        foreach (var item in request.Software)
        {
            var errors = CheckItem(item);
            if (errors.Count > 0) throw new InvalidDataException($"{item.Name}：{string.Join("；", errors)}");
        }
        if (!File.Exists(request.RunnerPath)) throw new FileNotFoundException("未找到自包含执行器，请先运行发布脚本。", request.RunnerPath);
        var output = Path.GetFullPath(request.OutputPath);
        if (!output.EndsWith(".exe", StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("输出文件必须是 .exe。");
        if (output.Equals(Path.GetFullPath(request.RunnerPath), StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("不能覆盖安装执行器。");
        var manifest = new BundleManifest { Name = request.Name.Trim(), DefaultScheme = request.DefaultScheme,
            Categories = CategoryNames.Normalize(request.Categories.Concat(request.Software.Select(s => s.Category))),
            Schemes = request.Schemes.ToDictionary(p => p.Key, p => p.Value.ToList()) };
        var sources = new Dictionary<string, List<(string Source, string Relative)>>(StringComparer.Ordinal);
        foreach (var item in request.Software)
        {
            var source = Path.GetFullPath(item.FilePath!);
            var files = new List<(string Source, string Relative)>();
            if (!string.IsNullOrWhiteSpace(item.ResourceDirectory))
            {
                var root = Path.GetFullPath(item.ResourceDirectory);
                CollectFiles(root, root, files);
                if (output.StartsWith(root.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException("输出位置不能位于关联资源目录内。");
            }
            else files.Add((source, Path.GetFileName(source)));
            foreach (var file in files)
            {
                PackagePaths.Resolve(Path.GetTempPath(), file.Relative);
                if (file.Source.Equals(output, StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("不能覆盖源安装文件。");
            }
            var installer = string.IsNullOrWhiteSpace(item.ResourceDirectory) ? Path.GetFileName(source) :
                Path.GetRelativePath(Path.GetFullPath(item.ResourceDirectory), source).Replace('\\', '/');
            var description = Describe(item);
            description.Installer = installer;
            manifest.Software.Add(description);
            sources.Add(item.Id, files);
        }
        BundleReader.ValidateManifest(manifest);
        InstallationEngine.Sort(manifest.Software); // Reject missing dependencies and cycles before writing anything.
        Directory.CreateDirectory(Path.GetDirectoryName(output)!);
        var staging = Path.Combine(Path.GetDirectoryName(output)!, $".sia-{Guid.NewGuid():N}.tmp");
        var archivePath = staging + ".zip";
        try
        {
            var allFiles = sources.Values.SelectMany(v => v).ToList();
            // Conservative scratch-space estimate: original size + output + headroom.
            long totalBytes = allFiles.Sum(f => new FileInfo(f.Source).Length);
            var drive = new DriveInfo(Path.GetPathRoot(output)!);
            if (drive.AvailableFreeSpace < totalBytes * 2 + new FileInfo(request.RunnerPath).Length + 100L * 1024 * 1024)
                throw new IOException("输出磁盘空间不足，请选择其他位置。");
            var hashes = new HashSet<string>(StringComparer.Ordinal);
            int completed = 0;
            using (var zipFile = new FileStream(archivePath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None))
            {
                using var archive = new ZipArchive(zipFile, ZipArchiveMode.Create, true);
                foreach (var software in manifest.Software)
                foreach (var file in sources[software.Id])
                {
                    token.ThrowIfCancellationRequested();
                    progress?.Report(new($"正在打包：{software.Name} / {file.Relative}", completed, allFiles.Count));
                    await using var input = new FileStream(file.Source, FileMode.Open, FileAccess.Read, FileShare.Read, 131072, true);
                    var hash = Convert.ToHexString(await SHA256.HashDataAsync(input, token));
                    software.Files.Add(new() { Path = file.Relative, Sha256 = hash, Size = input.Length });
                    if (hashes.Add(hash))
                    {
                        input.Position = 0;
                        await using var entry = archive.CreateEntry("content/" + hash, CompressionLevel.Fastest).Open();
                        await input.CopyToAsync(entry, token);
                    }
                    completed++;
                    if (zipFile.Length + new FileInfo(request.RunnerPath).Length + BundleReader.FooterSize > MaxBundleBytes)
                        throw new InvalidDataException("超过首版单文件容量上限（约 3.4 GiB），请减少软件数量。");
                }
                var manifestBytes = JsonSerializer.SerializeToUtf8Bytes(manifest, JsonOptions);
                if (manifestBytes.Length > 4 * 1024 * 1024) throw new InvalidDataException("清单过大，请减少关联文件数量。");
                await using var metadata = archive.CreateEntry("manifest.json").Open();
                await metadata.WriteAsync(manifestBytes, token);
            }
            if (new FileInfo(archivePath).Length + new FileInfo(request.RunnerPath).Length + BundleReader.FooterSize > MaxBundleBytes)
                throw new InvalidDataException("超过首版单文件容量上限，请减少软件数量。");
            progress?.Report(new("正在组合执行器并校验资源…", completed, allFiles.Count));
            await using (var destination = new FileStream(staging, FileMode.CreateNew, FileAccess.Write, FileShare.None, 131072, true))
            {
                await using var runner = new FileStream(request.RunnerPath, FileMode.Open, FileAccess.Read, FileShare.Read, 131072, true);
                var magic = new byte[2];
                await runner.ReadExactlyAsync(magic, token);
                if (magic[0] != 'M' || magic[1] != 'Z') throw new InvalidDataException("执行器不是有效的 Windows 程序。");
                runner.Position = 0;
                await runner.CopyToAsync(destination, token);
                await using var payload = new FileStream(archivePath, FileMode.Open, FileAccess.Read, FileShare.Read, 131072, true);
                var digest = await SHA256.HashDataAsync(payload, token);
                payload.Position = 0;
                await payload.CopyToAsync(destination, token);
                await destination.WriteAsync(BitConverter.GetBytes(payload.Length), token);
                await destination.WriteAsync(digest, token);
                await destination.WriteAsync(BundleReader.Magic, token);
                await destination.FlushAsync(token);
            }
            token.ThrowIfCancellationRequested();
            File.Move(staging, output, true);
            progress?.Report(new("生成完成", completed, allFiles.Count));
        }
        finally
        {
            if (File.Exists(staging)) File.Delete(staging);
            if (File.Exists(archivePath)) File.Delete(archivePath);
        }
    }

    private static void CollectFiles(string root, string folder, List<(string Source, string Relative)> files)
    {
        if (PackagePaths.IsReparsePoint(folder)) throw new InvalidDataException("关联资源目录不能包含链接。");
        foreach (var file in Directory.EnumerateFiles(folder).OrderBy(p => p, StringComparer.OrdinalIgnoreCase))
        {
            if (PackagePaths.IsReparsePoint(file)) throw new InvalidDataException("关联资源不能包含链接。");
            files.Add((Path.GetFullPath(file), Path.GetRelativePath(root, file).Replace('\\', '/')));
        }
        foreach (var directory in Directory.EnumerateDirectories(folder)) CollectFiles(root, directory, files);
    }

    public static BundleSoftware Describe(SoftwareItem item)
    {
        var profile = InstallerInspector.Inspect(item.FilePath);
        // An optional manual override also identifies the installer family.
        if (profile.Kind == InstallerKind.Unknown && item.DirectoryMode == InstallDirectoryMode.Nsis)
            profile = new(InstallerKind.Nsis, "NSIS（手动指定）", "/S", InstallDirectoryMode.Nsis);
        if (profile.Kind == InstallerKind.Unknown && item.DirectoryMode == InstallDirectoryMode.InnoSetup)
            profile = new(InstallerKind.InnoSetup, "Inno Setup（手动指定）", "/VERYSILENT /NORESTART /SP-", InstallDirectoryMode.InnoSetup);
        var mode = item.DirectoryMode == InstallDirectoryMode.Automatic ? profile.DirectoryMode : item.DirectoryMode;
        var parameter = item.DirectoryMode == InstallDirectoryMode.Automatic ? profile.DirectoryParameter : item.DirectoryParameter;
        var arguments = string.IsNullOrWhiteSpace(item.SilentInstallArgs) ? profile.SilentArguments : item.SilentInstallArgs.Trim();
        bool isMsi = string.Equals(Path.GetExtension(item.FilePath), ".msi", StringComparison.OrdinalIgnoreCase);
        bool directoryNeedsUi = !string.IsNullOrWhiteSpace(item.DefaultInstallDirectory) && mode == InstallDirectoryMode.Automatic;
        return new() { Id = item.Id, Name = item.Name, Version = item.Version, Category = item.Category,
            Installer = Path.GetFileName(item.FilePath ?? ""), Arguments = arguments, IsMsi = isMsi,
            DefaultInstallDirectory = InstallDirectoryOptions.Normalize(item.DefaultInstallDirectory), DirectoryMode = mode, DirectoryParameter = parameter,
            StartInteractively = (!isMsi && string.IsNullOrWhiteSpace(arguments)) || directoryNeedsUi,
            AllowInteractiveFallback = item.AllowInteractiveFallback, RequiresNetwork = item.RequiresNetwork, InstallerType = profile.Name,
            Order = item.InstallOrder, TimeoutMinutes = item.TimeoutMinutes, SuccessCodes = ParseCodes(item.SuccessCodes),
            RebootCodes = ParseCodes(item.RebootCodes), DetectionFile = item.DetectionFile, DetectionVersion = item.DetectionVersion,
            MsiProductCode = item.MsiProductCode, DependencyIds = new(item.DependencyIds) };
    }
}
