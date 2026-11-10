using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32.SafeHandles;

namespace SoftwareInstaller.Packaging;

internal sealed class InstallerProcess : IDisposable
{
    private readonly SafeFileHandle? job;
    private readonly string installerDirectory;
    private readonly InstallerProcessTree? tree;
    public Process Process { get; }
    public bool CanTerminate => job != null || tree != null;
    public string? TrackingWarning { get; }
    public string? TrackingNotice { get; }

    private InstallerProcess(Process process, string installer, SafeFileHandle? job = null, string? warning = null, bool allowBreakaway = false,
        bool useProcessTree = false, string? notice = null)
    {
        Process = process; this.job = job; TrackingWarning = warning;
        installerDirectory = Path.GetDirectoryName(Path.GetFullPath(installer))!;
        TrackingNotice = notice;
        if (useProcessTree || allowBreakaway && job != null) tree = new(process);
    }

    public static InstallerProcess Start(ProcessStartInfo info, bool trackChildren, bool killOnClose = false, bool allowBreakaway = false,
        bool useProcessTree = false)
    {
        if (!trackChildren || info.UseShellExecute)
            return new(Process.Start(info) ?? throw new IOException("无法启动安装程序。"), info.FileName);

        var startup = new StartupInfo { Size = Marshal.SizeOf<StartupInfo>() };
        var command = new StringBuilder($"\"{info.FileName}\" {info.Arguments}");
        const uint suspended = 0x00000004, noWindow = 0x08000000;
        useProcessTree |= Environment.OSVersion.Version < new Version(6, 2); // Windows 7 cannot nest jobs.
        uint flags = suspended | (info.CreateNoWindow ? noWindow : 0);
        ProcessInformation native;
        bool created = false;
        if (useProcessTree)
            created = CreateProcess(info.FileName, command, IntPtr.Zero, IntPtr.Zero, false,
                flags | 0x01000000, IntPtr.Zero, info.WorkingDirectory, ref startup, out native); // Try to leave an inherited job.
        else native = default;
        if (!created)
        {
            command = new StringBuilder($"\"{info.FileName}\" {info.Arguments}");
            if (!CreateProcess(info.FileName, command, IntPtr.Zero, IntPtr.Zero, false,
                flags, IntPtr.Zero, info.WorkingDirectory, ref startup, out native))
                throw new Win32Exception(Marshal.GetLastWin32Error());
        }

        Process? process = null;
        SafeFileHandle? job = null;
        try
        {
            process = Process.GetProcessById((int)native.ProcessId);
            _ = process.SafeHandle; // Retain a handle before a short-lived launcher can exit.
            string? notice = useProcessTree ? "进程管理：使用父子进程跟踪（兼容 Win7），不要求安装器加入嵌套进程组。" : null;
            if (!useProcessTree)
            {
                job = CreateJobObject(IntPtr.Zero, null);
                if (job.IsInvalid || !AssignProcessToJobObject(job, native.Process))
                {
                    int error = Marshal.GetLastWin32Error();
                    notice = $"无法加入安装进程组（系统错误 {error}：{new Win32Exception(error).Message}），已切换父子进程跟踪。";
                    job.Dispose(); job = null; useProcessTree = true;
                }
            }
            if (job != null && (killOnClose || allowBreakaway))
            {
                var limits = new JobLimits();
                limits.Basic.Flags = (killOnClose ? 0x2000u : 0) | (allowBreakaway ? 0x800u : 0);
                if (!SetInformationJobObject(job!, 9, ref limits, Marshal.SizeOf<JobLimits>()))
                    throw new Win32Exception(Marshal.GetLastWin32Error());
            }
            var session = new InstallerProcess(process, info.FileName, job, allowBreakaway: allowBreakaway, useProcessTree: useProcessTree, notice: notice);
            if (ResumeThread(native.Thread) == uint.MaxValue) throw new Win32Exception(Marshal.GetLastWin32Error());
            return session;
        }
        catch
        {
            TerminateProcess(native.Process, 1); // This newly created process has not been resumed on setup failure.
            process?.Dispose(); job?.Dispose();
            throw;
        }
        finally { CloseHandle(native.Thread); CloseHandle(native.Process); }
    }

    public async Task WaitAsync(CancellationToken token, Func<string, Task>? report = null)
    {
        // Avoid WaitForExitAsync, which requires modern .NET. Polling does no process enumeration.
        if (!Process.HasExited && report != null)
            await report($"正在等待程序主进程退出：PID={Process.Id}；操作完成后请退出该程序（包括托盘）。");
        while (!Process.HasExited)
        {
            token.ThrowIfCancellationRequested();
            tree?.Capture();
            await Task.Delay(tree == null ? 100 : 250, token).ConfigureAwait(false);
        }
        if (report != null) await report($"程序主进程已退出：PID={Process.Id}；返回码={Process.ExitCode}。");
        if (job == null && tree == null) return;
        string? lastWaiting = null;
        while (true)
        {
            token.ThrowIfCancellationRequested();
            var waiting = new List<string>();
            var applications = new List<string>();
            tree?.Capture();
            var processIds = GetProcessIds().Concat(tree?.ActiveIds ?? Enumerable.Empty<int>()).Distinct();
            foreach (int pid in processIds)
            {
                if (pid == Process.Id) continue; // Job accounting may briefly still contain the exited root.
                using var handle = OpenProcess(0x1000, false, pid); // PROCESS_QUERY_LIMITED_INFORMATION (Windows 7+).
                if (handle.IsInvalid)
                {
                    if (Marshal.GetLastWin32Error() != 87) waiting.Add($"PID={pid}（无法读取进程路径）");
                    continue;
                }
                var path = new StringBuilder(32768); int length = path.Capacity;
                if (!QueryFullProcessImageName(handle, 0, path, ref length))
                { waiting.Add($"PID={pid}（无法读取进程路径）"); continue; }
                string file = path.ToString();
                // Follow staged installers and Windows installation helpers, but not a newly installed
                // application launched from its final directory. Never enumerate unrelated processes.
                bool installer = IsWithin(file, installerDirectory) || IsWithin(file, Path.GetTempPath()) ||
                    IsWithin(file, Environment.GetFolderPath(Environment.SpecialFolder.Windows));
                // A console host can outlive the installer while a launched console application runs.
                // It carries no installation work of its own.
                if (file.Equals(Path.Combine(Environment.SystemDirectory, "conhost.exe"), StringComparison.OrdinalIgnoreCase)) installer = false;
                (installer ? waiting : applications).Add($"PID={pid} {file}");
            }
            if (waiting.Count == 0)
            {
                // The installed application may stay open. Normal completion must not kill it when
                // the job handle closes; cancellation still terminates the whole managed job.
                var limits = new JobLimits();
                limits.Basic.Flags = tree == null ? 0 : 0x800u; // Keep breakaway compatibility; remove only kill-on-close.
                if (job != null && !SetInformationJobObject(job, 9, ref limits, Marshal.SizeOf<JobLimits>()))
                    throw new Win32Exception(Marshal.GetLastWin32Error());
                if (applications.Count > 0 && report != null)
                    await report("安装器已退出，以下程序位于安装临时目录之外，保留运行并继续队列：" + string.Join("；", applications));
                return;
            }
            string message = "安装器主进程已退出，正在等待安装子进程：" + string.Join("；", waiting);
            if (message != lastWaiting && report != null) await report(message);
            lastWaiting = message;
            await Task.Delay(500, token).ConfigureAwait(false);
        }
    }

    private static bool IsWithin(string file, string directory) => file.StartsWith(
        directory.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase);

    private List<int> GetProcessIds()
    {
        if (job == null) return new();
        int capacity = 32;
        while (true)
        {
            int size = 8 + capacity * IntPtr.Size;
            IntPtr buffer = Marshal.AllocHGlobal(size);
            try
            {
                if (QueryJobProcessIds(job!, 3, buffer, size, IntPtr.Zero))
                {
                    int count = Marshal.ReadInt32(buffer, 4);
                    var ids = new List<int>(count);
                    for (int i = 0; i < count; i++) ids.Add(unchecked((int)Marshal.ReadIntPtr(buffer, 8 + i * IntPtr.Size).ToInt64()));
                    return ids;
                }
                int error = Marshal.GetLastWin32Error();
                if (error != 234) throw new Win32Exception(error); // ERROR_MORE_DATA
                capacity = checked(capacity * 2);
            }
            finally { Marshal.FreeHGlobal(buffer); }
        }
    }

    public async Task TerminateAsync()
    {
        if (!CanTerminate) throw new InvalidOperationException("当前安装器不支持受控终止，请在原安装窗口取消。");
        tree?.Capture();
        if (job != null && !TerminateJobObject(job, 1223)) throw new Win32Exception(Marshal.GetLastWin32Error());
        using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(5));
        // Termination must wait for every child, including installed applications.
        while (true)
        {
            tree?.Capture();
            tree?.TerminateChildren(includeRoot: job == null);
            var accounting = new JobAccounting();
            if (job != null && !QueryInformationJobObject(job, 1, out accounting, Marshal.SizeOf<JobAccounting>(), IntPtr.Zero))
                throw new Win32Exception(Marshal.GetLastWin32Error());
            if (Process.HasExited && accounting.ActiveProcesses == 0 && tree?.HasActiveChildren != true) return;
            await Task.Delay(100, deadline.Token).ConfigureAwait(false);
        }
    }

    public void Dispose() { tree?.Dispose(); Process.Dispose(); job?.Dispose(); }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct StartupInfo
    {
        public int Size;
        public IntPtr Reserved, Desktop, Title;
        public uint X, Y, XSize, YSize, XCountChars, YCountChars, FillAttribute, Flags;
        public ushort ShowWindow, ReservedSize;
        public IntPtr ReservedBytes, StandardInput, StandardOutput, StandardError;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct ProcessInformation
    {
        public IntPtr Process, Thread;
        public uint ProcessId, ThreadId;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct JobAccounting
    {
        public long UserTime, KernelTime, PeriodUserTime, PeriodKernelTime;
        public uint PageFaultCount, TotalProcesses, ActiveProcesses, TerminatedProcesses;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct BasicLimits
    {
        public long ProcessTime, JobTime;
        public uint Flags;
        public UIntPtr MinimumWorkingSet, MaximumWorkingSet;
        public uint ActiveProcessLimit;
        public UIntPtr Affinity;
        public uint PriorityClass, SchedulingClass;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct JobLimits
    {
        public BasicLimits Basic;
        public ulong ReadOperations, WriteOperations, OtherOperations, ReadBytes, WriteBytes, OtherBytes;
        public UIntPtr ProcessMemoryLimit, JobMemoryLimit, PeakProcessMemory, PeakJobMemory;
    }

    [DllImport("kernel32.dll", EntryPoint = "CreateProcessW", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CreateProcess(string applicationName, StringBuilder commandLine, IntPtr processAttributes,
        IntPtr threadAttributes, bool inheritHandles, uint flags, IntPtr environment, string directory,
        ref StartupInfo startupInfo, out ProcessInformation processInformation);
    [DllImport("kernel32.dll", EntryPoint = "CreateJobObjectW", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern SafeFileHandle CreateJobObject(IntPtr attributes, string? name);
    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool AssignProcessToJobObject(SafeFileHandle job, IntPtr process);
    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetInformationJobObject(SafeFileHandle job, int infoClass, ref JobLimits information, int size);
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern uint ResumeThread(IntPtr thread);
    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool QueryInformationJobObject(SafeFileHandle job, int infoClass, out JobAccounting information, int size, IntPtr returnLength);
    [DllImport("kernel32.dll", EntryPoint = "QueryInformationJobObject", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool QueryJobProcessIds(SafeFileHandle job, int infoClass, IntPtr information, int size, IntPtr returnLength);
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern SafeFileHandle OpenProcess(uint access, bool inheritHandle, int processId);
    [DllImport("kernel32.dll", EntryPoint = "QueryFullProcessImageNameW", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool QueryFullProcessImageName(SafeFileHandle process, int flags, StringBuilder path, ref int size);
    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool TerminateJobObject(SafeFileHandle job, uint exitCode);
    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool TerminateProcess(IntPtr process, uint exitCode);
    [DllImport("kernel32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CloseHandle(IntPtr handle);
}
