using System.Diagnostics;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using SoftwareInstaller.Models;

namespace SoftwareInstaller.Packaging;

public sealed class BundleReader : IDisposable
{
    public static readonly byte[] Magic = Encoding.ASCII.GetBytes("SIABUNDLE0000001");
    public const int FooterSize = 56;
    private readonly FileStream file;
    private readonly SliceStream slice;
    private readonly ZipArchive archive;
    public BundleManifest Manifest { get; }
    public InstallationPlan Plan { get; }

    private BundleReader(FileStream file, SliceStream slice, ZipArchive archive, BundleManifest manifest, InstallationPlan plan)
    {
        this.file = file; this.slice = slice; this.archive = archive; Manifest = manifest; Plan = plan;
    }

    public static async Task<BundleReader> OpenAsync(string path, CancellationToken token = default)
    {
        var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 131072, true);
        SliceStream? slice = null;
        ZipArchive? archive = null;
        try
        {
            if (file.Length < FooterSize) throw new InvalidDataException("文件不包含安装资源。");
            file.Position = file.Length - FooterSize;
            var footer = new byte[FooterSize];
            int offset = 0;
            while (offset < footer.Length)
            {
                int read = await file.ReadAsync(footer, offset, footer.Length - offset, token).ConfigureAwait(false);
                if (read == 0) throw new EndOfStreamException("安装包尾部不完整。");
                offset += read;
            }
            if (!footer.Skip(40).SequenceEqual(Magic)) throw new InvalidDataException("未找到安装包资源，请使用生成的 EXE。");
            var length = BitConverter.ToInt64(footer, 0);
            if (length <= 0 || length > file.Length - FooterSize || file.Length > RunnerRuntime.MaxBundleBytes)
                throw new InvalidDataException("安装包长度无效。");
            slice = new SliceStream(file, file.Length - FooterSize - length, length);
            // Keep the legacy digest field in the footer format, but do not scan the payload.
            archive = new ZipArchive(slice, ZipArchiveMode.Read, true);
            var metadata = archive.GetEntry("manifest.json") ?? throw new InvalidDataException("缺少资源清单。");
            if (metadata.Length > 4 * 1024 * 1024) throw new InvalidDataException("资源清单过大。");
            using var input = metadata.Open();
            token.ThrowIfCancellationRequested();
            var manifest = RunnerRuntime.ReadJson<BundleManifest>(input)
                ?? throw new InvalidDataException("资源清单无效。");
            // Startup only needs software/selection metadata. Validate resources when selected for extraction.
            var plan = ValidateManifest(manifest, validateResources: false);
            return new(file, slice, archive, manifest, plan);
        }
        catch { archive?.Dispose(); slice?.Dispose(); file.Dispose(); throw; }
    }

    public static InstallationPlan ValidateManifest(BundleManifest manifest, bool requireFiles = false, bool validateResources = true)
    {
        if (manifest.FormatVersion != 1 || string.IsNullOrWhiteSpace(manifest.Name) || manifest.Software.Count == 0)
            throw new InvalidDataException("不支持的安装包格式或空清单。");
        var ids = manifest.Software.Select(s => s.Id).ToHashSet(StringComparer.Ordinal);
        if (ids.Count != manifest.Software.Count || ids.Any(id => !Guid.TryParseExact(id, "N", out _)))
            throw new InvalidDataException("软件标识无效或重复。");
        if (!manifest.Schemes.ContainsKey(manifest.DefaultScheme) || manifest.Schemes.Values.Any(s => s.Count == 0 || s.Any(id => !ids.Contains(id))))
            throw new InvalidDataException("方案为空或引用了未打包的软件。");
        foreach (var item in manifest.Software)
        {
            InstallDirectoryOptions.BuildArguments(item.Arguments, item.DefaultInstallDirectory, item.DirectoryMode, item.DirectoryParameter, item.IsMsi);
            PackagePaths.Resolve(Path.GetTempPath(), item.Installer);
            if (item.TimeoutMinutes is < 1 or > 1440 || item.SuccessCodes.Count == 0 ||
                (!item.Installer.EndsWith(item.IsMsi ? ".msi" : ".exe", StringComparison.OrdinalIgnoreCase)))
                throw new InvalidDataException("安装配置无效。");
            if (validateResources) ValidateResources(item, requireFiles);
        }
        return InstallationPlan.Create(manifest.Software);
    }

    private static void ValidateResources(BundleSoftware item, bool requireFiles)
    {
        if (requireFiles && !item.Files.Any(f => f.Path.Equals(item.Installer, StringComparison.OrdinalIgnoreCase)))
            throw new InvalidDataException("安装文件未关联。");
        var paths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var resource in item.Files)
        {
            PackagePaths.Resolve(Path.GetTempPath(), resource.Path);
            if (!paths.Add(resource.Path) || resource.Size < 0 || resource.Sha256.Length != 64 || !resource.Sha256.All(Uri.IsHexDigit))
                throw new InvalidDataException("资源路径重复或校验值无效。");
        }
    }

    public async Task ExtractAsync(IEnumerable<BundleSoftware> software, string destination, IProgress<BundleProgress>? progress = null, CancellationToken token = default)
    {
        var items = software.ToList();
        token.ThrowIfCancellationRequested();
        foreach (var item in items)
        {
            if (!Manifest.Software.Contains(item)) throw new InvalidDataException("选择了未知的软件资源。");
            ValidateResources(item, true);
        }
        long required = items.SelectMany(s => s.Files).Sum(f => f.Size);
        Directory.CreateDirectory(destination);
        if (PackagePaths.IsReparsePoint(destination)) throw new InvalidDataException("工作目录不能是链接。");
        if (new DriveInfo(Path.GetPathRoot(Path.GetFullPath(destination))!).AvailableFreeSpace < required + 100L * 1024 * 1024)
            throw new IOException("临时磁盘空间不足，请清理空间后重试。");
        int completed = 0, total = items.Sum(s => s.Files.Count);
        long copied = 0, nextUpdate = 0;
        var clock = Stopwatch.StartNew();
        void Report(string message, bool force = false)
        {
            if (progress == null || !force && clock.ElapsedMilliseconds < nextUpdate) return;
            nextUpdate = clock.ElapsedMilliseconds + 100;
            progress.Report(new(message, completed, total, copied, required));
        }
        var buffer = new byte[131072];
        foreach (var item in items)
        {
            var root = PackagePaths.Resolve(destination, item.Id);
            Directory.CreateDirectory(root);
            foreach (var resource in item.Files)
            {
                token.ThrowIfCancellationRequested();
                var message = $"正在解包：{item.Name} / {resource.Path}";
                Report(message);
                var target = PackagePaths.Resolve(root, resource.Path);
                Directory.CreateDirectory(Path.GetDirectoryName(target)!);
                // Destination is a newly created, private random directory. Never overwrite existing files.
                var entry = archive.GetEntry("content/" + resource.Sha256);
                if (entry == null || entry.Length != resource.Size) throw new InvalidDataException($"{item.Name} 的资源缺失或长度不符。");
                using var output = new FileStream(target, FileMode.CreateNew, FileAccess.Write, FileShare.None, 131072, true);
                using var input = entry.Open();
                using var hash = SHA256.Create();
                long written = 0;
                int read;
                while ((read = await input.ReadAsync(buffer, 0, buffer.Length, token).ConfigureAwait(false)) > 0)
                {
                    if (read > resource.Size - written) throw new InvalidDataException($"{item.Name} 的资源长度不符。");
                    hash.TransformBlock(buffer, 0, read, null, 0);
                    await output.WriteAsync(buffer, 0, read, token).ConfigureAwait(false);
                    written += read;
                    copied += read;
                    Report(message);
                }
                if (output.Length != resource.Size) throw new InvalidDataException($"{item.Name} 的资源长度不符。");
                hash.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
                string digest = BitConverter.ToString(hash.Hash!).Replace("-", "");
                if (!digest.Equals(resource.Sha256, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException($"{item.Name} 的资源已损坏，已阻止安装。请重新获取安装包。");
                completed++;
            }
        }
        Report("资源已准备就绪", force: true);
    }

    public void Dispose() { archive.Dispose(); slice.Dispose(); file.Dispose(); }

    private sealed class SliceStream(FileStream source, long offset, long length) : Stream
    {
        private long position;
        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => false;
        public override long Length => length;
        public override long Position { get => position; set => Seek(value, SeekOrigin.Begin); }
        public override int Read(byte[] buffer, int offsetInBuffer, int count)
        {
            source.Position = offset + position;
            int read = source.Read(buffer, offsetInBuffer, (int)Math.Min(count, length - position));
            position += read; return read;
        }
        public override async Task<int> ReadAsync(byte[] buffer, int offsetInBuffer, int count, CancellationToken cancellationToken)
        {
            source.Position = offset + position;
            int read = await source.ReadAsync(buffer, offsetInBuffer, (int)Math.Min(count, length - position), cancellationToken).ConfigureAwait(false);
            position += read; return read;
        }
        public override long Seek(long value, SeekOrigin origin)
        {
            long next = origin switch { SeekOrigin.Begin => value, SeekOrigin.Current => position + value, SeekOrigin.End => length + value, _ => throw new ArgumentOutOfRangeException(nameof(origin)) };
            if (next < 0 || next > length) throw new IOException("超出资源边界。");
            return position = next;
        }
        public override void Flush() { }
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
