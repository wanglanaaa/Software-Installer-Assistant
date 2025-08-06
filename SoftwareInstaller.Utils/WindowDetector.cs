using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace SoftwareInstaller.Utils
{
    public static class WindowDetector
    {
        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder strText, int maxCount);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        public static bool HasVisibleWindow(int processId)
        {
            bool found = false;
            EnumWindows((hWnd, lParam) =>
            {
                if (!IsWindowVisible(hWnd))
                    return true; // Continue enumeration

                GetWindowThreadProcessId(hWnd, out uint windowProcessId);
                if (windowProcessId == processId)
                {
                    int length = GetWindowTextLength(hWnd);
                    if (length > 0)
                    {
                        found = true;
                        return false; // Stop enumeration
                    }
                }
                return true; // Continue enumeration
            }, IntPtr.Zero);

            return found;
        }
    }
}
