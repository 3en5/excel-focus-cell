// Detects Excel's "Find and Replace" dialog (Ctrl+F / Ctrl+H) with Win32 only - no object-model calls.
// Measured 21/10/2026 on Excel LTSC 2024 (Hebrew UI): a visible top-level window of this process, class
// "bosa_sdm_XL9", title "חיפוש והחלפה". Other Excel dialogs share the class, so the title is what identifies it.
using System;
using System.Runtime.InteropServices;
using System.Text;

namespace Natanor.FocusOverlay
{
    internal static class FindDialog
    {
        private static readonly string[] Titles = { "חיפוש והחלפה", "Find and Replace" };

        public static bool IsOpen(int pid)
        {
            bool found = false;
            EnumWindows((h, l) =>
            {
                int p; GetWindowThreadProcessId(h, out p);
                if (p != pid || !IsWindowVisible(h)) return true;
                var cls = new StringBuilder(32); GetClassName(h, cls, 32);
                if (cls.ToString() != "bosa_sdm_XL9") return true;
                var title = new StringBuilder(128); GetWindowText(h, title, 128);
                foreach (var t in Titles) if (title.ToString() == t) { found = true; return false; }
                return true;
            }, IntPtr.Zero);
            return found;
        }

        private delegate bool EnumProc(IntPtr h, IntPtr l);
        [DllImport("user32.dll")] private static extern bool EnumWindows(EnumProc f, IntPtr l);
        [DllImport("user32.dll")] private static extern int GetWindowThreadProcessId(IntPtr h, out int pid);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr h);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr h, StringBuilder s, int n);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
    }
}
