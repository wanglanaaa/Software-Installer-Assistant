using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32.SafeHandles;

namespace SoftwareInstaller.Packaging;

internal sealed class InstallerProcess : IDisposable
{
    private readonly SafeFileHandle? job;
    public Process Process { get; }
    public bool CanTerminate => job != null;
    public string? TrackingWarning { get; }

    private InstallerProcess(Process process, SafeFileHandle? job = null, string? warning = null)
    {
        Process = process; this.job = job; TrackingWarning = warning;
    }

    public static InstallerProcess Start(ProcessStartInfo info, bool trackChildren, bool killOnClose = false)
    {
        if (!trackChildren || info.UseShellExecute)
            return new(Process.Start(info) ?? throw new IOException("无法启动安装程序。"));

        var startup = new StartupInfo { Size = Marshal.SizeOf<StartupInfo>() };
        var command = new StringBuilder($"\"{info.FileName}\" {info.Arguments}");
        const uint suspended = 0x00000004, noWindow = 0x08000000;
        if (!CreateProcess(info.FileName, command, IntPtr.Zero, IntPtr.Zero, false,
            suspended | (info.CreateNoWindow ? noWindow : 0), IntPtr.Zero, info.WorkingDirectory, ref startup, out var native))
            throw new Win32Exception(Marshal.GetLastWin32Error());

        Process? process = null;
        SafeFileHandle? job = null;
        try
        {
            process = Process.GetProcessById((int)native.ProcessId);
            _ = process.SafeHandle; // Retain a handle before a short-lived launcher can exit.
            job = CreateJobObject(IntPtr.Zero, null);
            string? warning = null;
            if (job.IsInvalid || !AssignProcessToJobObject(job, native.Process))
            {
                if (killOnClose) throw new Win32Exception(Marshal.GetLastWin32Error(), "无法建立安装进程管理，已阻止启动。");
                warning = "无法管理安装子进程：" + new Win32Exception(Marshal.GetLastWin32Error()).Message;
                job.Dispose(); job = null;
            }
            if (killOnClose)
            {
                var limits = new JobLimits();
                limits.Basic.Flags = 0x2000; // JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE, supported since Windows XP.
                if (!SetInformationJobObject(job!, 9, ref limits, Marshal.SizeOf<JobLimits>()))
                    throw new Win32Exception(Marshal.GetLastWin32Error());
            }
            if (ResumeThread(native.Thread) == uint.MaxValue) throw new Win32Exception(Marshal.GetLastWin32Error());
            return new(process, job, warning);
        }
        catch
        {
            TerminateProcess(native.Process, 1); // This newly created process has not been resumed on setup failure.
            process?.Dispose(); job?.Dispose();
            throw;
        }
        finally { CloseHandle(native.Thread); CloseHandle(native.Process); }
    }

    public async Task WaitAsync(CancellationToken token)
    {
        // Avoid WaitForExitAsync, which requires modern .NET. Polling does no process enumeration.
        while (!Process.HasExited) await Task.Delay(100, token).ConfigureAwait(false);
        if (job == null) return;
        while (true)
        {
            if (!QueryInformationJobObject(job, 1, out var accounting, Marshal.SizeOf<JobAccounting>(), IntPtr.Zero))
                throw new Win32Exception(Marshal.GetLastWin32Error());
            if (accounting.ActiveProcesses == 0) return;
            await Task.Delay(100, token);
        }
    }

    public async Task TerminateAsync()
    {
        if (job == null) throw new InvalidOperationException("当前安装器不支持受控终止，请在原安装窗口取消。");
        if (!TerminateJobObject(job, 1223)) throw new Win32Exception(Marshal.GetLastWin32Error());
        using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(5));
        await WaitAsync(deadline.Token);
    }

    public void Dispose() { Process.Dispose(); job?.Dispose(); }

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
