using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace JinoSupporter.Web.Services;

/// <summary>Windows Job Object with <c>KILL_ON_JOB_CLOSE</c> — every assigned
/// child process (and any process *they* spawn, e.g. Excel.exe from the helper)
/// dies the moment this process's handle table is freed: graceful exit, crash,
/// or Ctrl-C from the WPF launcher.
///
/// Why it matters here: a helper subprocess mid-Excel-conversion otherwise
/// outlives the web exe, keeps a file lock on its own binary, and blocks the
/// next build's <c>CopyExcelHelperToWebOutput</c> target — the user has to
/// hunt the orphan in Task Manager before they can rebuild.
///
/// One singleton job for the lifetime of the process. The handle is
/// intentionally never closed — letting the process exit close it is what
/// triggers the kill.</summary>
[SupportedOSPlatform("windows")]
internal static class ChildProcessJob
{
    private static readonly IntPtr _job;

    static ChildProcessJob()
    {
        _job = CreateJobObject(IntPtr.Zero, null);
        if (_job == IntPtr.Zero) return;

        var info = new JOBOBJECT_EXTENDED_LIMIT_INFORMATION
        {
            BasicLimitInformation = new JOBOBJECT_BASIC_LIMIT_INFORMATION
            {
                LimitFlags = JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE,
            },
        };
        int len = Marshal.SizeOf<JOBOBJECT_EXTENDED_LIMIT_INFORMATION>();
        IntPtr ptr = Marshal.AllocHGlobal(len);
        try
        {
            Marshal.StructureToPtr(info, ptr, false);
            SetInformationJobObject(
                _job, JobObjectInfoType.ExtendedLimitInformation, ptr, (uint)len);
        }
        finally { Marshal.FreeHGlobal(ptr); }
    }

    /// <summary>Best-effort: silently no-ops if the job couldn't be created or
    /// the assignment fails (e.g. the process already belongs to a job that
    /// disallows nesting on older Windows).</summary>
    public static void Assign(System.Diagnostics.Process p)
    {
        if (_job == IntPtr.Zero) return;
        try { AssignProcessToJobObject(_job, p.Handle); }
        catch { }
    }

    private const uint JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE = 0x2000;

    private enum JobObjectInfoType
    {
        ExtendedLimitInformation = 9,
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct IO_COUNTERS
    {
        public ulong ReadOperationCount, WriteOperationCount, OtherOperationCount;
        public ulong ReadTransferCount,  WriteTransferCount,  OtherTransferCount;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct JOBOBJECT_BASIC_LIMIT_INFORMATION
    {
        public long    PerProcessUserTimeLimit;
        public long    PerJobUserTimeLimit;
        public uint    LimitFlags;
        public UIntPtr MinimumWorkingSetSize;
        public UIntPtr MaximumWorkingSetSize;
        public uint    ActiveProcessLimit;
        public UIntPtr Affinity;
        public uint    PriorityClass;
        public uint    SchedulingClass;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct JOBOBJECT_EXTENDED_LIMIT_INFORMATION
    {
        public JOBOBJECT_BASIC_LIMIT_INFORMATION BasicLimitInformation;
        public IO_COUNTERS IoInfo;
        public UIntPtr ProcessMemoryLimit;
        public UIntPtr JobMemoryLimit;
        public UIntPtr PeakProcessMemoryUsed;
        public UIntPtr PeakJobMemoryUsed;
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr CreateJobObject(IntPtr lpJobAttributes, string? lpName);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetInformationJobObject(
        IntPtr hJob,
        JobObjectInfoType infoType,
        IntPtr lpJobObjectInfo,
        uint cbJobObjectInfoLength);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool AssignProcessToJobObject(IntPtr job, IntPtr process);
}
