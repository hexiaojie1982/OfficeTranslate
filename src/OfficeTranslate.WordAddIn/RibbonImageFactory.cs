using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace OfficeTranslate.WordAddIn
{
    internal static class RibbonImageFactory
    {
        private static readonly Dictionary<string, object> Pictures = new Dictionary<string, object>();
        private static readonly List<Bitmap> Bitmaps = new List<Bitmap>();
        private static readonly object Sync = new object();
        private static readonly Color Blue = Color.FromArgb(44, 105, 220);
        private static readonly Color Dark = Color.FromArgb(47, 61, 82);
        private static readonly Color Red = Color.FromArgb(221, 68, 68);

        public static object Get(string controlId)
        {
            lock (Sync)
            {
                if (Pictures.TryGetValue(controlId, out var existing)) return existing;
                var bitmap = Draw(controlId);
                Bitmaps.Add(bitmap);
                var picture = PictureDispConverter.Convert(bitmap);
                Pictures[controlId] = picture;
                return picture;
            }
        }

        public static object GetLanguage(string language)
        {
            var key = "Language:" + language;
            lock (Sync)
            {
                if (Pictures.TryGetValue(key, out var existing)) return existing;
                var bitmap = DrawLanguageFlag(language);
                Bitmaps.Add(bitmap);
                var picture = PictureDispConverter.Convert(bitmap);
                Pictures[key] = picture;
                return picture;
            }
        }

        private static Bitmap DrawLanguageFlag(string language)
        {
            if (language == "自动检测")
            {
                var globe = new Bitmap(32, 32, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                using (var graphics = Graphics.FromImage(globe))
                {
                    graphics.SmoothingMode = SmoothingMode.AntiAlias;
                    DrawLanguage(graphics, false);
                }
                return globe;
            }

            var bitmap = new Bitmap(32, 32, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                var rect = new Rectangle(3, 7, 26, 18);
                if (language == "简体中文") DrawChina(g, rect);
                else if (language == "繁體中文") DrawHongKong(g, rect);
                else if (language == "英语") DrawUnitedStates(g, rect);
                else if (language == "日语") DrawJapan(g, rect);
                else if (language == "韩语") DrawKorea(g, rect);
                else if (language == "法语") DrawVerticalTricolor(g, rect, Color.FromArgb(0, 85, 164), Color.White, Color.FromArgb(239, 65, 53));
                else if (language == "德语") DrawHorizontalTricolor(g, rect, Color.Black, Color.FromArgb(221, 0, 0), Color.FromArgb(255, 206, 0));
                else if (language == "西班牙语") DrawSpain(g, rect);
                else if (language == "俄语") DrawHorizontalTricolor(g, rect, Color.White, Color.FromArgb(0, 57, 166), Color.FromArgb(213, 43, 30));
                else if (language == "葡萄牙语") DrawPortugal(g, rect);
                else if (language == "意大利语") DrawVerticalTricolor(g, rect, Color.FromArgb(0, 146, 70), Color.White, Color.FromArgb(206, 43, 55));
                else DrawSaudiArabia(g, rect);
                using (var border = new Pen(Color.FromArgb(105, 115, 128), 1F)) g.DrawRectangle(border, rect);
            }
            return bitmap;
        }

        private static void DrawHorizontalTricolor(Graphics g, Rectangle r, Color a, Color b, Color c)
        {
            using (var ba = new SolidBrush(a)) using (var bb = new SolidBrush(b)) using (var bc = new SolidBrush(c))
            { g.FillRectangle(ba, r.X, r.Y, r.Width, 6); g.FillRectangle(bb, r.X, r.Y + 6, r.Width, 6); g.FillRectangle(bc, r.X, r.Y + 12, r.Width, 6); }
        }

        private static void DrawVerticalTricolor(Graphics g, Rectangle r, Color a, Color b, Color c)
        {
            using (var ba = new SolidBrush(a)) using (var bb = new SolidBrush(b)) using (var bc = new SolidBrush(c))
            { g.FillRectangle(ba, r.X, r.Y, 9, r.Height); g.FillRectangle(bb, r.X + 9, r.Y, 8, r.Height); g.FillRectangle(bc, r.X + 17, r.Y, 9, r.Height); }
        }

        private static void DrawChina(Graphics g, Rectangle r)
        {
            using (var red = new SolidBrush(Color.FromArgb(222, 41, 16))) g.FillRectangle(red, r);
            using (var yellow = new SolidBrush(Color.FromArgb(255, 222, 0))) g.FillEllipse(yellow, r.X + 4, r.Y + 4, 5, 5);
        }

        private static void DrawHongKong(Graphics g, Rectangle r)
        {
            using (var red = new SolidBrush(Color.FromArgb(222, 41, 16))) g.FillRectangle(red, r);
            using (var white = new SolidBrush(Color.White))
            {
                var cx = r.X + r.Width / 2F;
                var cy = r.Y + r.Height / 2F;
                for (var i = 0; i < 5; i++)
                {
                    var angle = -Math.PI / 2 + i * Math.PI * 2 / 5;
                    var px = cx + (float)Math.Cos(angle) * 3.2F;
                    var py = cy + (float)Math.Sin(angle) * 3.2F;
                    g.FillEllipse(white, px - 2.2F, py - 1.4F, 4.4F, 2.8F);
                }
            }
            using (var red = new SolidBrush(Color.FromArgb(222, 41, 16))) g.FillEllipse(red, r.X + 11.5F, r.Y + 7.5F, 3F, 3F);
        }

        private static void DrawUnitedStates(Graphics g, Rectangle r)
        {
            using (var white = new SolidBrush(Color.White)) g.FillRectangle(white, r);
            using (var red = new SolidBrush(Color.FromArgb(178, 34, 52))) for (var y = 0; y < 18; y += 4) g.FillRectangle(red, r.X, r.Y + y, r.Width, 2);
            using (var blue = new SolidBrush(Color.FromArgb(60, 59, 110))) g.FillRectangle(blue, r.X, r.Y, 11, 9);
        }

        private static void DrawJapan(Graphics g, Rectangle r)
        {
            using (var white = new SolidBrush(Color.White)) g.FillRectangle(white, r);
            using (var red = new SolidBrush(Color.FromArgb(188, 0, 45))) g.FillEllipse(red, r.X + 9, r.Y + 5, 9, 9);
        }

        private static void DrawKorea(Graphics g, Rectangle r)
        {
            using (var white = new SolidBrush(Color.White)) g.FillRectangle(white, r);
            using (var red = new SolidBrush(Color.FromArgb(205, 46, 58))) g.FillPie(red, r.X + 9, r.Y + 5, 9, 9, 180, 180);
            using (var blue = new SolidBrush(Color.FromArgb(0, 71, 160))) g.FillPie(blue, r.X + 9, r.Y + 5, 9, 9, 0, 180);
        }

        private static void DrawSpain(Graphics g, Rectangle r)
        {
            using (var red = new SolidBrush(Color.FromArgb(170, 21, 27))) g.FillRectangle(red, r);
            using (var yellow = new SolidBrush(Color.FromArgb(241, 191, 0))) g.FillRectangle(yellow, r.X, r.Y + 5, r.Width, 8);
        }

        private static void DrawPortugal(Graphics g, Rectangle r)
        {
            using (var green = new SolidBrush(Color.FromArgb(4, 106, 56))) g.FillRectangle(green, r.X, r.Y, 11, r.Height);
            using (var red = new SolidBrush(Color.FromArgb(218, 41, 28))) g.FillRectangle(red, r.X + 11, r.Y, 15, r.Height);
            using (var yellow = new SolidBrush(Color.FromArgb(255, 204, 0))) g.FillEllipse(yellow, r.X + 8, r.Y + 6, 6, 6);
        }

        private static void DrawSaudiArabia(Graphics g, Rectangle r)
        {
            using (var green = new SolidBrush(Color.FromArgb(0, 108, 53))) g.FillRectangle(green, r);
            using (var white = new Pen(Color.White, 2F)) g.DrawLine(white, r.X + 7, r.Y + 12, r.X + 20, r.Y + 12);
        }

        private static Bitmap Draw(string id)
        {
            var bitmap = new Bitmap(32, 32, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                if (id.EndsWith("SourceLanguageMenu", StringComparison.Ordinal)) DrawLanguage(g, false);
                else if (id.EndsWith("TargetLanguageMenu", StringComparison.Ordinal)) DrawLanguage(g, true);
                else if (id.EndsWith("Selection", StringComparison.Ordinal)) DrawSelection(g);
                else if (id.EndsWith("Document", StringComparison.Ordinal)) DrawDocument(g);
                else if (id.EndsWith("Bilingual", StringComparison.Ordinal)) DrawBilingual(g);
                else if (id.EndsWith("ImageOcr", StringComparison.Ordinal)) DrawImageOcr(g);
                else if (id.EndsWith("Cancel", StringComparison.Ordinal)) DrawCancel(g);
                else if (id.EndsWith("Settings", StringComparison.Ordinal)) DrawSettings(g);
                else DrawSwap(g);
            }
            return bitmap;
        }

        private static void DrawSelection(Graphics g)
        {
            using (var pen = new Pen(Dark, 1.8F))
            using (var bluePen = new Pen(Blue, 2.4F) { StartCap = LineCap.Round, EndCap = LineCap.Round })
            using (var fill = new SolidBrush(Color.FromArgb(40, Blue)))
            {
                g.DrawRectangle(pen, 5, 4, 18, 23);
                g.FillRectangle(fill, 8, 11, 12, 8);
                g.DrawLine(bluePen, 9, 13, 19, 13); g.DrawLine(bluePen, 9, 17, 17, 17);
                g.DrawLine(bluePen, 21, 23, 28, 23); g.DrawLine(bluePen, 25, 19, 28, 23); g.DrawLine(bluePen, 25, 27, 28, 23);
            }
        }

        private static void DrawLanguage(Graphics g, bool target)
        {
            using (var pen = new Pen(Dark, 1.7F) { StartCap = LineCap.Round, EndCap = LineCap.Round })
            using (var accent = new Pen(Blue, 2F) { StartCap = LineCap.Round, EndCap = LineCap.Round })
            {
                g.DrawEllipse(pen, 5, 5, 22, 22);
                g.DrawEllipse(pen, 11, 5, 10, 22);
                g.DrawLine(pen, 6, 12, 26, 12);
                g.DrawLine(pen, 6, 20, 26, 20);
                if (target)
                {
                    g.DrawLine(accent, 18, 26, 29, 26);
                    g.DrawLine(accent, 25, 22, 29, 26);
                }
                else
                {
                    g.DrawLine(accent, 3, 26, 14, 26);
                    g.DrawLine(accent, 3, 26, 7, 22);
                }
            }
        }

        private static void DrawDocument(Graphics g)
        {
            using (var pen = new Pen(Dark, 1.8F))
            using (var bluePen = new Pen(Blue, 2F) { StartCap = LineCap.Round, EndCap = LineCap.Round })
            {
                g.DrawRectangle(pen, 5, 3, 21, 26);
                g.DrawLine(bluePen, 9, 9, 22, 9); g.DrawLine(bluePen, 9, 14, 22, 14); g.DrawLine(bluePen, 9, 19, 19, 19);
                g.DrawLine(bluePen, 16, 25, 27, 25); g.DrawLine(bluePen, 23, 21, 27, 25); g.DrawLine(bluePen, 23, 29, 27, 25);
            }
        }

        private static void DrawCancel(Graphics g)
        {
            using (var pen = new Pen(Red, 2.4F) { StartCap = LineCap.Round, EndCap = LineCap.Round })
            {
                g.DrawEllipse(pen, 4, 4, 24, 24); g.DrawLine(pen, 10, 10, 22, 22); g.DrawLine(pen, 22, 10, 10, 22);
            }
        }

        private static void DrawBilingual(Graphics g)
        {
            using (var pen = new Pen(Dark, 1.7F) { StartCap = LineCap.Round, EndCap = LineCap.Round })
            using (var bluePen = new Pen(Blue, 2F) { StartCap = LineCap.Round, EndCap = LineCap.Round })
            {
                g.DrawRectangle(pen, 4, 4, 24, 24);
                g.DrawLine(pen, 8, 9, 20, 9);
                g.DrawLine(pen, 8, 13, 17, 13);
                g.DrawLine(bluePen, 8, 20, 24, 20);
                g.DrawLine(bluePen, 8, 24, 21, 24);
            }
        }

        private static void DrawImageOcr(Graphics g)
        {
            using (var pen = new Pen(Dark, 1.7F))
            using (var accent = new Pen(Blue, 2F) { StartCap = LineCap.Round, EndCap = LineCap.Round })
            {
                g.DrawRectangle(pen, 4, 6, 24, 20);
                g.DrawEllipse(pen, 8, 10, 5, 5);
                g.DrawLine(pen, 5, 23, 12, 17); g.DrawLine(pen, 12, 17, 17, 21); g.DrawLine(pen, 17, 21, 22, 15); g.DrawLine(pen, 22, 15, 28, 22);
                g.DrawLine(accent, 7, 29, 25, 29);
            }
        }

        private static void DrawSettings(Graphics g)
        {
            using (var pen = new Pen(Dark, 2F) { StartCap = LineCap.Round, EndCap = LineCap.Round })
            using (var accent = new Pen(Blue, 2.2F))
            {
                g.DrawEllipse(accent, 10, 10, 12, 12); g.DrawEllipse(pen, 14, 14, 4, 4);
                for (var i = 0; i < 8; i++)
                {
                    var angle = Math.PI * i / 4;
                    var x1 = 16 + (float)Math.Cos(angle) * 8; var y1 = 16 + (float)Math.Sin(angle) * 8;
                    var x2 = 16 + (float)Math.Cos(angle) * 12; var y2 = 16 + (float)Math.Sin(angle) * 12;
                    g.DrawLine(pen, x1, y1, x2, y2);
                }
            }
        }

        private static void DrawSwap(Graphics g)
        {
            using (var pen = new Pen(Blue, 2.2F) { StartCap = LineCap.Round, EndCap = LineCap.Round })
            {
                g.DrawLine(pen, 5, 11, 25, 11); g.DrawLine(pen, 21, 7, 25, 11); g.DrawLine(pen, 21, 15, 25, 11);
                g.DrawLine(pen, 27, 21, 7, 21); g.DrawLine(pen, 11, 17, 7, 21); g.DrawLine(pen, 11, 25, 7, 21);
            }
        }

        private sealed class PictureDispConverter : AxHost
        {
            private PictureDispConverter() : base(string.Empty) { }
            public static object Convert(Image image) => GetIPictureDispFromPicture(image);
        }
    }
}
