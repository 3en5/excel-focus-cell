// Overlay renderer: ONE per-pixel-alpha layered window over the grid (UpdateLayeredWindow). No Office references.
// WS_EX_LAYERED | WS_EX_TRANSPARENT (mouse passes through) | WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW (no taskbar button).
// Owned by the workbook's XLMAIN and kept directly above it in z-order, so Excel dialogs stay on top.
//
// Why one window (measured 21/10/2026): moving and resizing 4-8 SetLayeredWindowAttributes windows per selection step
// cost ~9 ms per draw; a single ULW bitmap keeps the cost flat regardless of how many segments there are.
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace Natanor.FocusOverlay
{
    internal sealed unsafe class Win32Overlay : IDisposable
    {
        private const string ClassName = "NatanorFocusOverlayPoc";
        private IntPtr _owner, _hwnd, _memDc, _dib, _oldBmp;
        private uint* _bits;
        private PxRect _bounds;                       // screen rect the window and bitmap currently cover
        private readonly List<PxRect> _painted = new List<PxRect>();   // last painted rects, in bitmap coordinates
        private readonly uint _pixel;                 // premultiplied BGRA
        private bool _visible;
        private static bool _registered;

        public Win32Overlay(IntPtr owner, int bgr, byte alpha)
        {
            _owner = owner;
            uint b = (uint)(bgr >> 16 & 0xFF), g = (uint)(bgr >> 8 & 0xFF), r = (uint)(bgr & 0xFF);
            _pixel = ((uint)alpha << 24) | ((r * alpha / 255) << 16) | ((g * alpha / 255) << 8) | (b * alpha / 255);
            if (!_registered)
            {
                var wc = new N.WNDCLASSEX();
                wc.cbSize = Marshal.SizeOf(typeof(N.WNDCLASSEX));
                wc.lpfnWndProc = N.GetProcAddress(N.GetModuleHandle("user32.dll"), "DefWindowProcW");
                wc.hInstance = N.GetModuleHandle(null);
                wc.lpszClassName = ClassName;
                if (N.RegisterClassEx(ref wc) == 0) Log.Write("register-class-failed " + Marshal.GetLastWin32Error());
                _registered = true;
            }
            CreateOverlayWindow(_owner);
            _memDc = N.CreateCompatibleDC(IntPtr.Zero);
        }

        // Owned popup windows are destroyed by Windows when their owner is destroyed.
        // Recreate the overlay instead of trying to re-parent a stale HWND.
        private void CreateOverlayWindow(IntPtr owner)
        {
            _owner = owner;
            _hwnd = N.CreateWindowEx(N.WS_EX_LAYERED | N.WS_EX_TRANSPARENT | N.WS_EX_NOACTIVATE | N.WS_EX_TOOLWINDOW,
                ClassName, "", N.WS_POPUP, 0, 0, 0, 0, owner, IntPtr.Zero, N.GetModuleHandle(null), IntPtr.Zero);
            _visible = false;
            if (_hwnd == IntPtr.Zero) Log.Write("create-window-failed " + Marshal.GetLastWin32Error());
        }

        // rects are in screen pixels; bounds is the grid rect the overlay may cover
        public void Show(IList<PxRect> rects, IntPtr owner, PxRect bounds)
        {
            if (_hwnd == IntPtr.Zero || !N.IsWindow(_hwnd) || owner != _owner)
            {
                if (_hwnd != IntPtr.Zero && N.IsWindow(_hwnd))
                {
                    Hide();
                    N.DestroyWindow(_hwnd);
                }
                _hwnd = IntPtr.Zero;
                _visible = false;
                CreateOverlayWindow(owner);
            }
            if (_hwnd == IntPtr.Zero) return;
            if (_dib == IntPtr.Zero || bounds.L != _bounds.L || bounds.T != _bounds.T || bounds.R != _bounds.R || bounds.B != _bounds.B)
                Allocate(bounds);
            int w = _bounds.R - _bounds.L, h = _bounds.B - _bounds.T;

            foreach (var p in _painted) Fill(p, 0, w, h);          // clear only what was painted last time
            _painted.Clear();
            foreach (var s in rects)
            {
                var local = new PxRect(s.L - _bounds.L, s.T - _bounds.T, s.R - _bounds.L, s.B - _bounds.T).Intersect(new PxRect(0, 0, w, h));
                if (local.IsEmpty) continue;
                Fill(local, _pixel, w, h);
                _painted.Add(local);
            }

            var dst = new N.POINT { X = _bounds.L, Y = _bounds.T };
            var size = new N.SIZE { cx = w, cy = h };
            var src = new N.POINT();
            var blend = new N.BLENDFUNCTION { BlendOp = 0, BlendFlags = 0, SourceConstantAlpha = 255, AlphaFormat = 1 };
            N.UpdateLayeredWindow(_hwnd, IntPtr.Zero, ref dst, ref size, _memDc, ref src, 0, ref blend, N.ULW_ALPHA);

            if (!_visible)
            {
                IntPtr after = N.GetWindow(_owner, N.GW_HWNDPREV);     // the window just above Excel; ours goes below it
                if (after == IntPtr.Zero || after == _hwnd) after = N.HWND_TOP;
                N.SetWindowPos(_hwnd, after, 0, 0, 0, 0, N.SWP_NOACTIVATE | N.SWP_SHOWWINDOW | N.SWP_NOMOVE | N.SWP_NOSIZE | N.SWP_NOOWNERZORDER);
                _visible = true;
            }
        }

        public void Hide()
        {
            if (_visible) { N.ShowWindow(_hwnd, N.SW_HIDE); _visible = false; }
        }

        private void Allocate(PxRect bounds)
        {
            FreeBitmap();
            _bounds = bounds;
            var bmi = new N.BITMAPINFOHEADER();
            bmi.biSize = Marshal.SizeOf(typeof(N.BITMAPINFOHEADER));
            bmi.biWidth = bounds.R - bounds.L; bmi.biHeight = -(bounds.B - bounds.T);   // top-down
            bmi.biPlanes = 1; bmi.biBitCount = 32;
            IntPtr bits;
            _dib = N.CreateDIBSection(_memDc, ref bmi, 0, out bits, IntPtr.Zero, 0);
            _bits = (uint*)bits;                                       // zero-initialised by GDI = fully transparent
            _oldBmp = N.SelectObject(_memDc, _dib);
            _painted.Clear();
        }

        private void Fill(PxRect r, uint value, int w, int h)
        {
            for (int y = r.T; y < r.B; y++)
            {
                uint* row = _bits + (long)y * w;
                for (int x = r.L; x < r.R; x++) row[x] = value;
            }
        }

        private void FreeBitmap()
        {
            if (_dib != IntPtr.Zero) { N.SelectObject(_memDc, _oldBmp); N.DeleteObject(_dib); _dib = IntPtr.Zero; _bits = null; }
        }

        public void Dispose()
        {
            Hide();
            FreeBitmap();
            if (_memDc != IntPtr.Zero) { N.DeleteDC(_memDc); _memDc = IntPtr.Zero; }
            if (_hwnd != IntPtr.Zero && N.IsWindow(_hwnd)) N.DestroyWindow(_hwnd);
            _hwnd = IntPtr.Zero;
        }
    }

    internal static class N
    {
        public const int WS_EX_LAYERED = 0x80000, WS_EX_TRANSPARENT = 0x20, WS_EX_NOACTIVATE = 0x8000000, WS_EX_TOOLWINDOW = 0x80;
        public const uint WS_POPUP = 0x80000000;
        public const int SW_HIDE = 0, GW_HWNDPREV = 3, GWLP_HWNDPARENT = -8, ULW_ALPHA = 2;
        public const uint SWP_NOSIZE = 0x1, SWP_NOMOVE = 0x2, SWP_NOACTIVATE = 0x10, SWP_SHOWWINDOW = 0x40, SWP_NOOWNERZORDER = 0x200;
        public static readonly IntPtr HWND_TOP = IntPtr.Zero;

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct WNDCLASSEX
        {
            public int cbSize; public uint style; public IntPtr lpfnWndProc; public int cbClsExtra; public int cbWndExtra;
            public IntPtr hInstance; public IntPtr hIcon; public IntPtr hCursor; public IntPtr hbrBackground;
            public string lpszMenuName; public string lpszClassName; public IntPtr hIconSm;
        }
        [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
        [StructLayout(LayoutKind.Sequential)] public struct POINT { public int X, Y; }
        [StructLayout(LayoutKind.Sequential)] public struct SIZE { public int cx, cy; }
        [StructLayout(LayoutKind.Sequential, Pack = 1)] public struct BLENDFUNCTION { public byte BlendOp, BlendFlags, SourceConstantAlpha, AlphaFormat; }
        [StructLayout(LayoutKind.Sequential)]
        public struct BITMAPINFOHEADER
        {
            public int biSize, biWidth, biHeight; public short biPlanes, biBitCount; public int biCompression, biSizeImage,
            biXPelsPerMeter, biYPelsPerMeter, biClrUsed, biClrImportant;
        }
        [StructLayout(LayoutKind.Sequential)] public struct LASTINPUTINFO { public uint cbSize; public uint dwTime; }

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)] public static extern ushort RegisterClassEx(ref WNDCLASSEX wc);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern IntPtr CreateWindowEx(int ex, string cls, string name, uint style, int x, int y, int w, int h, IntPtr parent, IntPtr menu, IntPtr inst, IntPtr param);
        [DllImport("user32.dll")] public static extern bool DestroyWindow(IntPtr h);
        [DllImport("user32.dll")] public static extern bool UpdateLayeredWindow(IntPtr h, IntPtr dstDc, ref POINT dst, ref SIZE size, IntPtr srcDc, ref POINT src, int key, ref BLENDFUNCTION blend, int flags);
        [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
        [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
        [DllImport("user32.dll")] public static extern IntPtr GetWindow(IntPtr h, int cmd);
        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtrW")] public static extern IntPtr SetWindowLongPtr(IntPtr h, int i, IntPtr v);
        [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] public static extern int GetWindowThreadProcessId(IntPtr h, out int pid);
        [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr h);
        [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr h);
        [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
        [DllImport("user32.dll")] public static extern bool GetClientRect(IntPtr h, out RECT r);
        [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
        [DllImport("user32.dll")] public static extern bool ClientToScreen(IntPtr h, ref POINT p);
        [DllImport("user32.dll")] public static extern bool GetLastInputInfo(ref LASTINPUTINFO i);
        public delegate bool EnumProc(IntPtr h, IntPtr l);
        [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, EnumProc f, IntPtr l);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
        [DllImport("gdi32.dll")] public static extern IntPtr CreateCompatibleDC(IntPtr dc);
        [DllImport("gdi32.dll")] public static extern bool DeleteDC(IntPtr dc);
        [DllImport("gdi32.dll")] public static extern IntPtr CreateDIBSection(IntPtr dc, ref BITMAPINFOHEADER bmi, uint usage, out IntPtr bits, IntPtr section, uint offset);
        [DllImport("gdi32.dll")] public static extern IntPtr SelectObject(IntPtr dc, IntPtr o);
        [DllImport("gdi32.dll")] public static extern bool DeleteObject(IntPtr o);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr GetModuleHandle(string name);
        [DllImport("kernel32.dll", CharSet = CharSet.Ansi)] public static extern IntPtr GetProcAddress(IntPtr mod, string name);

        // Tick of the last keyboard/mouse input in this session (cheap, no COM)
        public static int LastInputTick()
        {
            var i = new LASTINPUTINFO(); i.cbSize = (uint)Marshal.SizeOf(typeof(LASTINPUTINFO));
            return GetLastInputInfo(ref i) ? (int)i.dwTime : Environment.TickCount;
        }

        public static PxRect ClientRectOnScreen(IntPtr h)
        {
            RECT r; GetClientRect(h, out r);
            var p = new POINT(); ClientToScreen(h, ref p);
            return new PxRect(p.X, p.Y, p.X + r.R, p.Y + r.B);
        }

        public static PxRect WindowRect(IntPtr h) { RECT r; GetWindowRect(h, out r); return new PxRect(r.L, r.T, r.R, r.B); }

        // EXCEL7 grid window and NUIScrollbar windows under a workbook's XLMAIN (one enumeration; callers cache the result)
        public static void FindGridParts(IntPtr topLevel, out IntPtr grid, List<IntPtr> scrollbars)
        {
            IntPtr g = IntPtr.Zero;
            scrollbars.Clear();
            EnumChildWindows(topLevel, (h, l) =>
            {
                var sb = new StringBuilder(32); GetClassName(h, sb, 32);
                string c = sb.ToString();
                if (c == "EXCEL7" && g == IntPtr.Zero) g = h;
                else if (c == "NUIScrollbar") scrollbars.Add(h);
                return true;
            }, IntPtr.Zero);
            grid = g;
        }
    }
}
