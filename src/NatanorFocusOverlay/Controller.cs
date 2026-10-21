// Wires Excel events and a light poll to the adapter, the visibility policy and the renderer.
// Everything runs on Excel's main thread (WinForms timer = WM_TIMER on that thread).
using System;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

namespace Natanor.FocusOverlay
{
    internal sealed class Controller : IDisposable
    {
        private readonly Excel.Application _app;
        private readonly IntPtr _main;
        private readonly ExcelReadOnlyAdapter _reader;
        private readonly Win32Overlay _overlay;
        private readonly System.Windows.Forms.Timer _timer;
        private readonly int _pid = Process.GetCurrentProcess().Id;
        private bool _dirty = true;

        // ---- Mode: user preference and administrator policy (Settings.cs) -----------------------------------------
        // Always  - highlight all the time.  Search - only while Excel's "Find and Replace" dialog is open.  Off.
        // Settings are re-read every 5 s (cheap registry reads), so an administrator's policy reaches running sessions
        // without restarting Excel, and a choice made in one Excel window reaches the user's other Excel processes.
        // While the highlight is not engaged the add-in makes NO object-model calls: Off/policy -> a 5 s settings
        // check only; Search with the dialog closed -> a 250 ms Win32 window check only.
        public event Action StateChanged;
        public bool LockedByPolicy { get { return _policy == PolicyState.Disabled; } }
        public FocusMode Mode { get { return _mode; } }
        private PolicyState _policy;
        private FocusMode _mode;
        private bool _findOpen, _engaged;
        private int _lastSettingsCheck, _lastFindCheck;
        private const int SettingsIntervalMs = 5000, FindPollMs = 250;

        public void SetMode(FocusMode mode)
        {
            Settings.UserMode = mode;
            _mode = mode;
            _findOpen = mode == FocusMode.Search && FindDialog.IsOpen(_pid);
            Reconcile();
        }

        private void ReadSettings()
        {
            var policy = Settings.Policy();
            var mode = Settings.UserMode;
            bool changed = policy != _policy || mode != _mode;
            _policy = policy; _mode = mode;
            Reconcile();
            if (changed && StateChanged != null) StateChanged();
        }

        private void Reconcile()
        {
            bool on = _policy != PolicyState.Disabled && _mode != FocusMode.Off;
            bool engaged = on && (_mode == FocusMode.Always || _findOpen);
            int idleInterval = on && _mode == FocusMode.Search ? FindPollMs : SettingsIntervalMs;
            if (engaged == _engaged)
            {
                if (!engaged && _timer.Interval != idleInterval) _timer.Interval = idleInterval;
                return;
            }
            _engaged = engaged;
            if (engaged) { _dirty = true; _fullCheck = true; _timer.Interval = 100; }
            else
            {
                Hide(_policy == PolicyState.Disabled ? "disabled-by-policy" : _mode == FocusMode.Off ? "disabled-by-user" : "search-closed");
                _timer.Interval = idleInterval;
            }
            Log.Write("state engaged=" + _engaged + " mode=" + _mode + " policy=" + _policy + " find=" + _findOpen);
        }

        public Controller(Excel.Application app)
        {
            _app = app;
            _main = new IntPtr(app.Hwnd);
            Log.Write("ctor main=" + _main);
            _reader = new ExcelReadOnlyAdapter(app);
            _overlay = new Win32Overlay(_main, 0xE0A040 /* BGR: soft blue */, 46 /* ~18% */);
            Log.Write("ctor overlay ok");
            _app.SheetSelectionChange += OnSelection;
            _app.SheetChange += OnContent;
            _app.SheetActivate += OnSheet;
            _app.WindowActivate += OnWindow;
            _app.WindowResize += OnWindow;
            Log.Write("ctor events ok");
            _timer = new System.Windows.Forms.Timer();
            _timer.Interval = 100;
            _timer.Tick += OnTick;
            _policy = Settings.Policy(); _mode = Settings.UserMode;
            _findOpen = _mode == FocusMode.Search && FindDialog.IsOpen(_pid);
            _engaged = !(_policy != PolicyState.Disabled && (_mode == FocusMode.Always || _findOpen && _mode == FocusMode.Search));
            Reconcile();                                    // forces the first transition, sets the interval and logs
            _lastSettingsCheck = Environment.TickCount;
            _timer.Start();
        }

        private void OnSelection(object sh, Excel.Range target) { _dirty = true; TimedUpdate(); }   // cheap path
        private void OnContent(object sh, Excel.Range target) { _reader.ForgetSizes(); _dirty = true; }   // autofit may resize rows
        private void OnSheet(object sh) { _dirty = true; _fullCheck = true; TimedUpdate(); }
        private void OnWindow(Excel.Workbook wb, Excel.Window wn) { _dirty = true; _fullCheck = true; TimedUpdate(); }

        private void TimedUpdate()
        {
            if (!_engaged) return;                          // not engaged: never touch Excel
            Update();
        }
        // Adaptive poll. Scroll, zoom and freeze raise no Excel event but always follow user input, so poll at full
        // rate only for a short while after input, and once a second otherwise. Measured 21/10/2026: a flat 60 ms
        // poll cost ~37 core-ms/s at idle against a 1 ms/s baseline.
        private const int FastWindowMs = 1500, SlowIntervalMs = 1000;
        private int _lastSlowPoll;
        private void OnTick(object sender, EventArgs e)
        {
            int now = Environment.TickCount;
            if (now - _lastSettingsCheck >= SettingsIntervalMs) { _lastSettingsCheck = now; ReadSettings(); }
            if (_mode == FocusMode.Search && _policy != PolicyState.Disabled && now - _lastFindCheck >= FindPollMs)
            {
                _lastFindCheck = now;
                bool open = FindDialog.IsOpen(_pid);       // Win32 only
                if (open != _findOpen) { _findOpen = open; Reconcile(); }
            }
            if (!_engaged) return;
            bool fast = now - N.LastInputTick() <= FastWindowMs;
            int wanted = fast ? 100 : SlowIntervalMs;                // every WM_TIMER wakes Excel's own idle work too
            if (_timer.Interval != wanted) _timer.Interval = wanted;
            if (!fast && !_dirty)
            {
                // Idle: no object-model calls at all (each one costs Excel extra work after it returns, measured
                // 21/10/2026). Without input the sheet cannot move; only the window can, and Win32 sees that.
                if (now - _lastSlowPoll < SlowIntervalMs) return;
                _lastSlowPoll = now;
                if (!WindowStateChanged()) return;
                _dirty = true;
            }
            else if (!_dirty)
            {
                // Fast poll: a 4-call probe; a full read only when the probe changes, plus one full check a second
                // for split panes scrolled through their own scrollbars
                if (now - _lastProbe < 250 && now - _lastFullCheck < 1000) return;
                _lastProbe = now;
                Excel.Window w = null;
                try { w = _app.ActiveWindow; } catch { }
                string pk = w == null ? null : _reader.PollKey(w);
                bool due = now - _lastFullCheck >= 1000;
                if (pk != null && pk == _lastPollKey && !due) return;
                _lastPollKey = pk;
            }
            _fullCheck = true;
            _lastFullCheck = now;
            TimedUpdate();
        }

        private string _lastPollKey;
        private int _lastProbe;
        private int _lastFullCheck;
        private bool _fullCheck = true;
        private IntPtr _topHwnd;

        // EXCEL7 and scrollbar handles are looked up once per workbook window (an enumeration cost ~1.5 ms each time)
        private IntPtr _partsTop, _gridHwnd;
        private readonly System.Collections.Generic.List<IntPtr> _bars = new System.Collections.Generic.List<IntPtr>();
        private PxRect? GridRect(IntPtr top)
        {
            if (top != _partsTop || !N.IsWindow(_gridHwnd)) { N.FindGridParts(top, out _gridHwnd, _bars); _partsTop = top; }
            if (_gridHwnd == IntPtr.Zero) return null;
            return N.ClientRectOnScreen(_gridHwnd);
        }
        private System.Collections.Generic.List<PxRect> Occlusions(IntPtr top)
        {
            if (top != _partsTop) GridRect(top);
            var res = new System.Collections.Generic.List<PxRect>();
            foreach (var b in _bars) if (N.IsWindowVisible(b)) res.Add(N.WindowRect(b));
            return res;
        }

        private IntPtr _lastTop;
        private PxRect? _lastGrid;
        private bool _lastOursInFront, _lastIconic;
        private bool WindowStateChanged()
        {
            int fgPid; N.GetWindowThreadProcessId(N.GetForegroundWindow(), out fgPid);
            bool ours = fgPid == _pid;
            bool iconic = _lastTop != IntPtr.Zero && N.IsIconic(_lastTop);
            var grid = _lastTop == IntPtr.Zero ? null : GridRect(_lastTop);
            bool changed = ours != _lastOursInFront || iconic != _lastIconic || !Same(grid, _lastGrid);
            _lastOursInFront = ours; _lastIconic = iconic;
            return changed;
        }

        private static bool Same(PxRect? a, PxRect? b)
        {
            if (a.HasValue != b.HasValue) return false;
            if (!a.HasValue) return true;
            return a.Value.L == b.Value.L && a.Value.T == b.Value.T && a.Value.R == b.Value.R && a.Value.B == b.Value.B;
        }

        private void Update()
        {
            try
            {
                // Visibility policy first: cheap Win32 checks, no COM
                int fgPid; N.GetWindowThreadProcessId(N.GetForegroundWindow(), out fgPid);
                if (fgPid != _pid) { Hide("foreign-foreground"); return; }

                _dirty = false;
                Excel.Window w = _app.ActiveWindow;
                if (w == null) { Hide("no-window"); return; }
                // w.Hwnd costs ~44 us; it can only change with a window event or on the full check
                if (_fullCheck || _topHwnd == IntPtr.Zero) _topHwnd = new IntPtr(w.Hwnd);
                IntPtr top = _topHwnd;                          // the workbook's own XLMAIN
                if (N.IsIconic(top)) { Hide("minimized"); return; }
                var grid = GridRect(top);
                _lastTop = top; _lastGrid = grid; _lastOursInFront = true; _lastIconic = false;
                if (!grid.HasValue) { Hide("no-grid"); return; }
                var snap = _reader.Read(w, grid.Value, _fullCheck);
                _fullCheck = false;
                if (snap.Reason == Reason.Unchanged) return;    // nothing moved: the overlay on screen is current
                if (snap.Reason == Reason.Busy) return;         // editing or busy: leave the overlay as it is
                if (snap.Reason != Reason.Ok) { Hide(snap.Reason.ToString()); return; }
                if (snap.LayoutChanged) _partsTop = IntPtr.Zero;   // panes or splits may have added scrollbars
                var occ = Occlusions(top);
                var segs = GeometryEngine.SubtractAll(GeometryEngine.Segments(snap.Panes, true, true),
                                                      GeometryEngine.ExtendBars(occ, grid.Value));
                _overlay.Show(segs, top, grid.Value);
                _lastHide = null;
                if (Log.Verbose) Log.Write("seq=" + snap.Sequence + " segs=" + segs.Count + " " + snap.Debug);
            }
            catch (Exception ex)
            {
                Hide("fault");                                  // fail safe: hide, never guess
                Log.Write("update-fault " + ex.GetType().Name + " 0x" + ex.HResult.ToString("X8"));
            }
        }

        private string _lastHide;
        private void Hide(string why)
        {
            _overlay.Hide();
            if (why != _lastHide) { Log.Write("hide " + why); _lastHide = why; }
            _reader.Invalidate();                               // next update must redraw even if nothing moved
        }

        public void Dispose()
        {
            _timer.Stop(); _timer.Dispose();
            try
            {
                _app.SheetSelectionChange -= OnSelection;
                _app.SheetChange -= OnContent;
                _app.SheetActivate -= OnSheet;
                _app.WindowActivate -= OnWindow;
                _app.WindowResize -= OnWindow;
            }
            catch { }
            _overlay.Dispose();
        }
    }
}
