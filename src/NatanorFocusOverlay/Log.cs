// Diagnostics: versions, timings, reason codes and rectangles only. Never cell values, formulas, paths or titles.
using System;
using System.IO;

namespace Natanor.FocusOverlay
{
    internal static class Log
    {
        private static readonly string PathName = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "NatanorFocusOverlay", "focus-overlay.log");
        private static readonly object Gate = new object();
        // Per-update geometry lines cost a file append each; off unless NATANOR_FOCUS_VERBOSE=1 is set for the user.
        public static bool Verbose = Environment.GetEnvironmentVariable("NATANOR_FOCUS_VERBOSE") == "1";

        public static void Write(string line)
        {
            try
            {
                lock (Gate)
                {
                    Directory.CreateDirectory(System.IO.Path.GetDirectoryName(PathName));
                    var fi = new FileInfo(PathName);
                    if (fi.Exists && fi.Length > 2 * 1024 * 1024) fi.Delete();   // POC: crude size cap
                    File.AppendAllText(PathName, DateTime.Now.ToString("HH:mm:ss.fff") + " " + line + Environment.NewLine);
                }
            }
            catch { }   // logging must never affect Excel
        }
    }
}
