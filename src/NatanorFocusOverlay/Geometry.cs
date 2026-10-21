// Geometry engine: pure functions over rectangles in physical screen pixels. No Office references.
using System;
using System.Collections.Generic;

namespace Natanor.FocusOverlay
{
    internal struct PxRect
    {
        public int L, T, R, B;
        public PxRect(int l, int t, int r, int b) { L = Math.Min(l, r); R = Math.Max(l, r); T = Math.Min(t, b); B = Math.Max(t, b); }
        public bool IsEmpty { get { return R <= L || B <= T; } }
        public PxRect Intersect(PxRect o) { var x = new PxRect(); x.L = Math.Max(L, o.L); x.T = Math.Max(T, o.T); x.R = Math.Min(R, o.R); x.B = Math.Min(B, o.B); return x; }
        public override string ToString() { return "[" + L + "," + T + "," + R + "," + B + "]"; }
    }

    internal sealed class PaneGeometry
    {
        public PxRect Cells;          // visible cell area of the pane, already clipped to the grid window
        public PxRect? ActiveCell;    // active cell (or whole merge area) in this pane's coordinates, if it maps here
        public PxRect? RowBand;       // active row span (T..B) mapped in this pane, if the row is visible here
        public PxRect? ColBand;       // active column span (L..R) mapped in this pane, if the column is visible here
    }

    internal static class GeometryEngine
    {
        // segments = (rowBand UNION colBand) INTERSECT pane MINUS activeCell, per pane; no double blend at the crossing
        public static List<PxRect> Segments(IList<PaneGeometry> panes, bool row, bool col)
        {
            var result = new List<PxRect>();
            foreach (var p in panes)
            {
                var pieces = new List<PxRect>();
                if (row && p.RowBand.HasValue)
                {
                    var band = new PxRect(p.Cells.L, p.RowBand.Value.T, p.Cells.R, p.RowBand.Value.B).Intersect(p.Cells);
                    if (!band.IsEmpty) pieces.Add(band);
                }
                if (col && p.ColBand.HasValue)
                {
                    var band = new PxRect(p.ColBand.Value.L, p.Cells.T, p.ColBand.Value.R, p.Cells.B).Intersect(p.Cells);
                    // remove the crossing with the row band so it is not painted twice
                    if (!band.IsEmpty)
                        foreach (var s in pieces.Count > 0 ? Subtract(band, pieces[0]) : new List<PxRect> { band }) pieces.Add(s);
                }
                foreach (var piece in pieces)
                {
                    if (p.ActiveCell.HasValue) result.AddRange(Subtract(piece, p.ActiveCell.Value));
                    else result.Add(piece);
                }
            }
            return result;
        }

        // A horizontal scrollbar shares its strip with the sheet tabs, which are not a separate window: the whole strip
        // across the grid is off limits. Same for a vertical bar along the full grid height.
        public static List<PxRect> ExtendBars(IList<PxRect> bars, PxRect grid)
        {
            var res = new List<PxRect>();
            foreach (var b in bars)
            {
                // Excel keeps zero-size "visible" scrollbars for unused split panes (seen at [10,10,10,10], 125% DPI);
                // extending one of those would wipe the whole grid
                if (b.IsEmpty || b.Intersect(grid).IsEmpty) continue;
                if (b.R - b.L >= b.B - b.T) res.Add(new PxRect(grid.L, b.T, grid.R, Math.Max(b.B, grid.B)));   // horizontal: to the bottom
                else if (b.L - grid.L < grid.R - b.R) res.Add(new PxRect(grid.L, grid.T, b.R, grid.B));         // vertical on the left (RTL)
                else res.Add(new PxRect(b.L, grid.T, grid.R, grid.B));                                         // vertical on the right
            }
            return res;
        }

        public static List<PxRect> SubtractAll(List<PxRect> segs, IList<PxRect> occlusions)
        {
            foreach (var o in occlusions)
            {
                var next = new List<PxRect>();
                foreach (var s in segs) next.AddRange(Subtract(s, o));
                segs = next;
            }
            return segs;
        }

        // a MINUS b, as up to four rectangles
        public static List<PxRect> Subtract(PxRect a, PxRect b)
        {
            var res = new List<PxRect>();
            var i = a.Intersect(b);
            if (i.IsEmpty) { res.Add(a); return res; }
            if (a.T < i.T) res.Add(new PxRect(a.L, a.T, a.R, i.T));
            if (i.B < a.B) res.Add(new PxRect(a.L, i.B, a.R, a.B));
            if (a.L < i.L) res.Add(new PxRect(a.L, i.T, i.L, i.B));
            if (i.R < a.R) res.Add(new PxRect(i.R, i.T, a.R, i.B));
            return res;
        }
    }
}
