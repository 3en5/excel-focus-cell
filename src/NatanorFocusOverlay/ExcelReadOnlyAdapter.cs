// The only component that touches the Excel object model. READ ONLY: positions and geometry, never
// Value/Formula/Text, never Select/Activate/scroll, never any setter. Runs on Excel's main STA thread only.
//
// Cost model (measured 21/10/2026): every object-model call also costs Excel work after it returns, so calls are
// budgeted. Pane layout (visible ranges, pane rectangles) is cached and re-read only when the layout key changes;
// a plain selection move reads the active cell and maps only the edges it needs.
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace Natanor.FocusOverlay
{
    internal enum Reason { Ok, NoWindow, NotWorksheet, NotRange, Busy, Minimized, Foreign, NoGrid, Unchanged }

    internal sealed class GeometrySnapshot
    {
        public Reason Reason;
        public long Sequence;
        public List<PaneGeometry> Panes = new List<PaneGeometry>();
        public bool LayoutChanged;
        public string Debug;
    }

    internal sealed class ExcelReadOnlyAdapter
    {
        private sealed class PaneCache
        {
            public Excel.Pane Pane;
            public int Row1, Row2, Col1, Col2;
            public PxRect Cells;
            public double Ax, Bx, Ay, By;   // pixels = a * points + b, calibrated on the pane's own edges
            public int MapX(double pts) { return (int)Math.Round(Ax * pts + Bx); }
            public int MapY(double pts) { return (int)Math.Round(Ay * pts + By); }
        }

        // Points geometry of rows and columns already visited on this sheet: top/height, left/width.
        // Saves two object-model calls per axis per selection step once a row or column has been seen.
        private readonly Dictionary<int, double[]> _rowPts = new Dictionary<int, double[]>();
        private readonly Dictionary<int, double[]> _colPts = new Dictionary<int, double[]>();
        public void ForgetSizes() { _rowPts.Clear(); _colPts.Clear(); }

        private readonly Excel.Application _app;
        private long _seq;
        private string _layoutKey, _stateKey, _staticKey;
        private readonly List<PaneCache> _panes = new List<PaneCache>();
        private int _mirror;   // 0 = LTR; otherwise gridL + gridR (see X())
   // verbose only: worst linear-vs-exact edge difference seen, in pixels

        public ExcelReadOnlyAdapter(Excel.Application app) { _app = app; }

        public void Invalidate() { _layoutKey = null; _stateKey = null; _staticKey = null; ForgetSizes(); }

        // Cheap change probe for poll ticks, FAST calls only (measured 21/10/2026: ScrollRow, Zoom, Selection and Top
        // cost ~110 us each and their presence kept Excel's own CPU high; VisibleRange.Row/Column cost ~16 us).
        // Zoom and shape selection are left to the once-a-second full check.
        public string PollKey(Excel.Window w)
        {
            Excel.Range vr = null;
            try { vr = w.VisibleRange; return vr.Row + "," + vr.Column; }
            catch (COMException) { return null; }
            finally { Release(vr); }
        }

        // Active pane's scroll position through the fast VisibleRange call
        private static string ActiveScroll(Excel.Window w)
        {
            Excel.Range vr = null;
            try { vr = w.VisibleRange; return vr.Row + "," + vr.Column; }
            finally { Release(vr); }
        }

        // full=true re-reads the static layout (window, sheet, zoom, direction, panes, every pane's scroll);
        // full=false is the selection-event path: the sheet is a worksheet and the selection a range by definition,
        // and a selection move can scroll only the active pane
        public GeometrySnapshot Read(Excel.Window w, PxRect grid, bool full)
        {
            var s = new GeometrySnapshot { Sequence = ++_seq };
            Excel.Range ac = null, ma = null;
            try
            {
                if (full || _staticKey == null)
                {
                    if (!(_app.ActiveSheet is Excel.Worksheet)) { s.Reason = Reason.NotWorksheet; Invalidate(); return s; }
                    if (!(_app.Selection is Excel.Range)) { s.Reason = Reason.NotRange; Invalidate(); return s; }
                }
                ac = _app.ActiveCell;
                if (ac == null) { s.Reason = Reason.NotRange; Invalidate(); return s; }
                // Range.MergeCells costs ~5.8 ms per call (measured 21/10/2026); MergeArea.Count ~18 us
                ma = ac.MergeArea;
                bool merged = ma.Count > 1;
                Excel.Range area = merged ? ma : ac;
                int aRow1 = area.Row, aCol1 = area.Column, aRow2 = aRow1, aCol2 = aCol1;
                if (merged) { aRow2 = aRow1 + area.Rows.Count - 1; aCol2 = aCol1 + area.Columns.Count - 1; }

                // Layout key: anything that moves cells on screen without moving the selection
                Excel.Panes panes = null;
                int paneCount = _panes.Count;
                if (full || _staticKey == null)
                {
                    panes = w.Panes; paneCount = panes.Count;
                    var sk = new StringBuilder();
                    sk.Append(w.Hwnd).Append('|').Append(((Excel.Worksheet)_app.ActiveSheet).Index).Append('|').Append(w.Zoom)
                      .Append('|').Append(w.DisplayRightToLeft).Append('|').Append(grid).Append('|').Append(paneCount);
                    if (paneCount > 1) for (int i = 1; i <= paneCount; i++) { Excel.Pane p = panes[i]; sk.Append('|').Append(p.ScrollRow).Append(',').Append(p.ScrollColumn); }
                    string newStatic = sk.ToString();
                    if (newStatic != _staticKey) ForgetSizes();
                    _staticKey = newStatic;
                }
                string layoutKey = _staticKey + "|" + grid + "|" + ActiveScroll(w);
                string stateKey = layoutKey + "#" + aRow1 + "," + aCol1 + "," + aRow2 + "," + aCol2;

                // A drag-resize or a row-height/column-width change raises no event. Once a second (full check)
                // compare the active cell's real position with the caches; on a mismatch drop the caches AND the pane
                // layout so everything is re-read and redrawn. Stale geometry lives at most one second.
                // Must run before the "unchanged" exit: a resize moves cells without moving the selection.
                // (Clearing the caches unconditionally every second cost ~80 core-ms/s while navigating, 21/10/2026.)
                if (full && !merged && (_rowPts.ContainsKey(aRow1) || _colPts.ContainsKey(aCol1)))
                {
                    double[] vr0, vc0;
                    bool stale = (_rowPts.TryGetValue(aRow1, out vr0) && (Math.Abs(vr0[0] - (double)ac.Top) > 0.01 || Math.Abs(vr0[1] - (double)ac.Height) > 0.01))
                              || (_colPts.TryGetValue(aCol1, out vc0) && (Math.Abs(vc0[0] - (double)ac.Left) > 0.01 || Math.Abs(vc0[1] - (double)ac.Width) > 0.01));
                    if (stale)
                    {
                        ForgetSizes(); _layoutKey = null; _stateKey = null;
                        if (panes == null) { panes = w.Panes; paneCount = panes.Count; }
                        Log.Write("sizes-changed: caches and layout dropped");
                    }
                }
                if (stateKey == _stateKey) { s.Reason = Reason.Unchanged; return s; }
                if (panes == null && layoutKey != _layoutKey) { panes = w.Panes; paneCount = panes.Count; }

                if (layoutKey != _layoutKey)
                {
                    s.LayoutChanged = true;
                    _mirror = (bool)w.DisplayRightToLeft ? grid.L + grid.R : 0;
                    _panes.Clear();
                    for (int i = 1; i <= paneCount; i++)
                    {
                        Excel.Pane p = panes[i];
                        Excel.Range vr = p.VisibleRange;
                        double vL = (double)vr.Left, vT = (double)vr.Top, vW = (double)vr.Width, vH = (double)vr.Height;
                        var pc = new PaneCache { Pane = p, Row1 = vr.Row, Col1 = vr.Column };
                        pc.Row2 = pc.Row1 + vr.Rows.Count - 1; pc.Col2 = pc.Col1 + vr.Columns.Count - 1;
                        int pL = (int)Math.Round(vL), pR = (int)Math.Round(vL + vW), pT = (int)Math.Round(vT), pB = (int)Math.Round(vT + vH);
                        int xL = X(p, pL), xR = X(p, pR), yT = Y(p, pT), yB = Y(p, pB);
                        pc.Ax = pR == pL ? 0 : (double)(xR - xL) / (pR - pL); pc.Bx = xL - pc.Ax * pL;
                        pc.Ay = pB == pT ? 0 : (double)(yB - yT) / (pB - pT); pc.By = yT - pc.Ay * pT;
                        pc.Cells = new PxRect(xL, yT, xR, yB).Intersect(grid);
                        _panes.Add(pc);
                        Release(vr);
                    }
                    _layoutKey = layoutKey;
                }

                // Points of the active area: merged areas are read directly, single cells come from the size caches
                double aL, aT, aW, aH;
                double[] rp, cp;
                if (merged) { aL = (double)area.Left; aT = (double)area.Top; aW = (double)area.Width; aH = (double)area.Height; }
                else
                {

                    if (!_rowPts.TryGetValue(aRow1, out rp)) { rp = new[] { (double)ac.Top, (double)ac.Height }; _rowPts[aRow1] = rp; }
                    if (!_colPts.TryGetValue(aCol1, out cp)) { cp = new[] { (double)ac.Left, (double)ac.Width }; _colPts[aCol1] = cp; }
                    aT = rp[0]; aH = rp[1]; aL = cp[0]; aW = cp[1];
                }
                foreach (var pc in _panes)
                {
                    bool rowHere = aRow2 >= pc.Row1 && aRow1 <= pc.Row2;
                    bool colHere = aCol2 >= pc.Col1 && aCol1 <= pc.Col2;
                    var pg = new PaneGeometry { Cells = pc.Cells };
                    if (rowHere || colHere)
                    {
                        int cy1 = 0, cy2 = 0, cx1 = 0, cx2 = 0;
                        if (rowHere) { cy1 = pc.MapY(aT); cy2 = pc.MapY(aT + aH); pg.RowBand = new PxRect(pc.Cells.L, cy1, pc.Cells.R, cy2); }
                        if (colHere) { cx1 = pc.MapX(aL); cx2 = pc.MapX(aL + aW); pg.ColBand = new PxRect(cx1, pc.Cells.T, cx2, pc.Cells.B); }
                        if (rowHere && colHere) pg.ActiveCell = new PxRect(cx1, cy1, cx2, cy2);
                    }
                    s.Panes.Add(pg);
                }
                _stateKey = stateKey;
                if (Log.Verbose) s.Debug = "layout=" + layoutKey + " active=r" + aRow1 + "-" + aRow2 + "c" + aCol1 + "-" + aCol2;
                s.Reason = Reason.Ok;
            }
            catch (COMException ex) { s.Reason = Reason.Busy; s.Debug = "com 0x" + ex.HResult.ToString("X8"); }
            catch (InvalidCastException) { s.Reason = Reason.NotRange; Invalidate(); }
            finally { Release(ma); Release(ac); }
            return s;
        }

        // Release the fresh Range objects Excel hands us for each read (a new COM identity per call, so the RCW is ours
        // alone), on Excel's own thread. Left to the GC they are released from the finalizer thread, which has to
        // marshal back to Excel's STA. Never used on Application, Window, Worksheet or cached Pane objects, whose RCWs
        // can be shared with other code in the AppDomain.
        private static void Release(object o)
        {
            if (o != null && Marshal.IsComObject(o)) Marshal.ReleaseComObject(o);
        }

        // PointsToScreenPixelsX answers as if the sheet were left-to-right (measured 21/10/2026: P0 = grid left +
        // row header). For a right-to-left sheet, mirror around the grid window.
        private int X(Excel.Pane p, double pts)
        {
            int x = p.PointsToScreenPixelsX((int)Math.Round(pts));
            return _mirror == 0 ? x : _mirror - x;
        }
        private static int Y(Excel.Pane p, double pts) { return p.PointsToScreenPixelsY((int)Math.Round(pts)); }
    }
}
