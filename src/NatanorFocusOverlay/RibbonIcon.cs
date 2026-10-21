// The ribbon icon, drawn in code (no image files to ship): a small grid with the active row and column tinted and
// the active cell left white with a dark border - the same picture the overlay draws on the sheet.
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace Natanor.FocusOverlay
{
    internal static class RibbonIcon
    {
        public static object Picture(int size, bool lens)
        {
            using (var bmp = new Bitmap(size, size, System.Drawing.Imaging.PixelFormat.Format32bppArgb))
            {
                using (var g = Graphics.FromImage(bmp))
                {
                    g.SmoothingMode = SmoothingMode.None;
                    g.Clear(Color.Transparent);
                    int n = 4, cell = (size - 2) / n, origin = (size - cell * n) / 2, active = 1;
                    using (var tint = new SolidBrush(Color.FromArgb(255, 160, 200, 235)))
                    using (var white = new SolidBrush(Color.White))
                    using (var grid = new Pen(Color.FromArgb(255, 150, 150, 150), 1))
                    using (var border = new Pen(Color.FromArgb(255, 33, 115, 70), size >= 32 ? 2 : 1))
                    {
                        g.FillRectangle(white, origin, origin, cell * n, cell * n);
                        g.FillRectangle(tint, origin, origin + active * cell, cell * n, cell);   // active row
                        g.FillRectangle(tint, origin + active * cell, origin, cell, cell * n);   // active column
                        g.FillRectangle(white, origin + active * cell, origin + active * cell, cell, cell);
                        for (int i = 0; i <= n; i++)
                        {
                            g.DrawLine(grid, origin + i * cell, origin, origin + i * cell, origin + n * cell);
                            g.DrawLine(grid, origin, origin + i * cell, origin + n * cell, origin + i * cell);
                        }
                        g.DrawRectangle(border, origin + active * cell, origin + active * cell, cell, cell);
                    }
                    if (lens)
                    {
                        // magnifying glass over the bottom-right corner: the "only while searching" variant
                        g.SmoothingMode = SmoothingMode.AntiAlias;
                        float r = size * 0.22f, cx = size * 0.62f, cy = size * 0.62f, w = size >= 32 ? 3f : 2f;
                        using (var glass = new SolidBrush(Color.FromArgb(220, 255, 255, 255)))
                        using (var rim = new Pen(Color.FromArgb(255, 60, 60, 60), w))
                        using (var handle = new Pen(Color.FromArgb(255, 60, 60, 60), w + 1) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                        {
                            g.FillEllipse(glass, cx - r, cy - r, 2 * r, 2 * r);
                            g.DrawEllipse(rim, cx - r, cy - r, 2 * r, 2 * r);
                            g.DrawLine(handle, cx + r * 0.7f, cy + r * 0.7f, size - 2, size - 2);
                        }
                    }
                }
                return PictureConverter.ToPicture((Image)bmp.Clone());
            }
        }

        // AxHost exposes the Image -> IPictureDisp conversion Office needs, but only to subclasses
        private sealed class PictureConverter : AxHost
        {
            private PictureConverter() : base("") { }
            public static object ToPicture(Image image) { return GetIPictureDispFromPicture(image); }
        }
    }
}
