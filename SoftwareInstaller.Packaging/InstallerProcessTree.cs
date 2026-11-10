using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace SoftwareInstaller.Packaging;

// Some interactive self-extracting programs explicitly create children outside their parent's job.
// Observe ancestry only while that program runs; keep process handles so cancellation never targets
// an unrelated process merely because a PID has been reused.
internal sealed class InstallerProcessTree : IDisposable
{
    private sealed record Tracked(Process Process, DateTime Started);
    private readonly Dictionary<int, Tracked> processes = new();
    private readonly int rootId;

    public InstallerProcessTree(Process root)
    {
        rootId = root.Id;
        processes.Add(root.Id, new(root, root.StartTime.ToUniversalTime()));
    }

    public void Capture()
    {
        using var snapshot = CreateToolhelp32Snapshot(2, 0); // TH32CS_SNAPPROCESS
        if (snapshot.IsInvalid) throw new Win32Exception(Marshal.GetLastWin32Error());
        var entries = new List<(int Id, int Parent)>();
        var entry = new ProcessEntry { Size = (uint)Marshal.SizeOf<ProcessEntry>() };
        bool available = Process32First(snapshot, ref entry);
        while (available)
        {
            entries.Add(((int)entry.Id, (int)entry.ParentId));
            available = Process32Next(snapshot, ref entry);
        }
        int error = Marshal.GetLastWin32Error();
        if (error != 18) throw new Win32Exception(error); // ERROR_NO_MORE_FILES
        bool added;
        do
        {
            added = false;
            foreach (var candidate in entries)
            {
                if (processes.ContainsKey(candidate.Id) || !processes.TryGetValue(candidate.Parent, out var parent)) continue;
                Process? process = null;
                try
                {
                    process = Process.GetProcessById(candidate.Id);
                    _ = process.SafeHandle;
                    DateTime started = process.StartTime.ToUniversalTime();
                    if (started < parent.Started || parent.Process.HasExited && started > parent.Process.ExitTime.ToUniversalTime()) continue;
                    processes.Add(candidate.Id, new(process, started));
                    process = null; added = true;
                }
                catch (ArgumentException) { } // Exited between snapshot and opening the handle.
                catch (InvalidOperationException) { }
                catch (Win32Exception ex) when (ex.NativeErrorCode == 87) { }
                finally { process?.Dispose(); }
            }
        } while (added);
    }

    public IEnumerable<int> ActiveIds => processes.Where(p => !p.Value.Process.HasExited).Select(p => p.Key);
    public bool HasActiveChildren => processes.Any(p => p.Key != rootId && !p.Value.Process.HasExited);

    public void TerminateChildren(bool includeRoot = false)
    {
        foreach (var child in processes.Where(p => includeRoot || p.Key != rootId).Select(p => p.Value.Process))
        {
            if (child.HasExited) continue;
            try { child.Kill(); }
            catch (InvalidOperationException) when (child.HasExited) { }
            catch (Win32Exception) when (child.HasExited) { }
        }
    }

    public void Dispose()
    {
        foreach (var child in processes.Where(p => p.Key != rootId)) child.Value.Process.Dispose();
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct ProcessEntry
    {
        public uint Size, Usage, Id;
        public UIntPtr DefaultHeapId;
        public uint ModuleId, Threads, ParentId;
        public int Priority;
        public uint Flags;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)] public string FileName;
    }
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern SafeFileHandle CreateToolhelp32Snapshot(uint flags, uint processId);
    [DllImport("kernel32.dll", EntryPoint = "Process32FirstW", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool Process32First(SafeFileHandle snapshot, ref ProcessEntry entry);
    [DllImport("kernel32.dll", EntryPoint = "Process32NextW", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool Process32Next(SafeFileHandle snapshot, ref ProcessEntry entry);
}
