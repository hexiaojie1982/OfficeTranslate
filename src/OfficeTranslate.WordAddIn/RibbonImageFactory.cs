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
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                var rect = new Rectangle(2, 7, 28, 18);
                using (var shadow = new SolidBrush(Color.FromArgb(38, 0, 0, 0)))
                    g.FillRectangle(shadow, rect.X + 1, rect.Y + 1, rect.Width, rect.Height);
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
                using (var border = new Pen(Color.FromArgb(112, 122, 136), 1F))
                    g.DrawRectangle(border, rect.X - 0.5F, rect.Y - 0.5F, rect.Width + 1F, rect.Height + 1F);
            }
            return bitmap;
        }

        private static void DrawHorizontalTricolor(Graphics g, Rectangle r, Color a, Color b, Color c)
        {
            using (var ba = new SolidBrush(a)) using (var bb = new SolidBrush(b)) using (var bc = new SolidBrush(c))
            {
                var third = r.Height / 3F;
                g.FillRectangle(ba, r.X, r.Y, r.Width, third);
                g.FillRectangle(bb, r.X, r.Y + third, r.Width, third);
                g.FillRectangle(bc, r.X, r.Y + third * 2F, r.Width, r.Height - third * 2F);
            }
        }

        private static void DrawVerticalTricolor(Graphics g, Rectangle r, Color a, Color b, Color c)
        {
            using (var ba = new SolidBrush(a)) using (var bb = new SolidBrush(b)) using (var bc = new SolidBrush(c))
            {
                var third = r.Width / 3F;
                g.FillRectangle(ba, r.X, r.Y, third, r.Height);
                g.FillRectangle(bb, r.X + third, r.Y, third, r.Height);
                g.FillRectangle(bc, r.X + third * 2F, r.Y, r.Width - third * 2F, r.Height);
            }
        }

        private static void DrawChina(Graphics g, Rectangle r)
        {
            using (var red = new SolidBrush(Color.FromArgb(238, 28, 37))) g.FillRectangle(red, r);
            using (var yellow = new SolidBrush(Color.FromArgb(255, 222, 0)))
            {
                FillStar(g, yellow, r.X + 6F, r.Y + 5.2F, 3.4F, -90F);
                FillStar(g, yellow, r.X + 12.2F, r.Y + 2.5F, 1.25F, -72F);
                FillStar(g, yellow, r.X + 14.3F, r.Y + 5.2F, 1.25F, -90F);
                FillStar(g, yellow, r.X + 13.7F, r.Y + 8.4F, 1.25F, -106F);
                FillStar(g, yellow, r.X + 10.9F, r.Y + 10.6F, 1.25F, -124F);
            }
        }

        private static void DrawHongKong(Graphics g, Rectangle r)
        {
            using (var red = new SolidBrush(Color.FromArgb(222, 41, 16))) g.FillRectangle(red, r);
            var cx = r.X + r.Width / 2F;
            var cy = r.Y + r.Height / 2F;
            using (var white = new SolidBrush(Color.White))
            {
                for (var i = 0; i < 5; i++)
                {
                    var state = g.Save();
                    g.TranslateTransform(cx, cy);
                    g.RotateTransform(i * 72F);
                    using (var petal = new GraphicsPath())
                    {
                        petal.StartFigure();
                        petal.AddBezier(0F, 0F, -2.1F, -1.8F, -1.7F, -5.2F, 0F, -6F);
                        petal.AddBezier(0F, -6F, 2.6F, -5.3F, 3.1F, -2F, 0F, 0F);
                        g.FillPath(white, petal);
                    }
                    using (var dot = new SolidBrush(Color.FromArgb(222, 41, 16))) g.FillEllipse(dot, -0.45F, -4.5F, 0.9F, 0.9F);
                    g.Restore(state);
                }
            }
        }

        private static void DrawUnitedStates(Graphics g, Rectangle r)
        {
            using (var white = new SolidBrush(Color.White)) g.FillRectangle(white, r);
            var stripe = r.Height / 13F;
            using (var red = new SolidBrush(Color.FromArgb(178, 34, 52)))
                for (var i = 0; i < 13; i += 2) g.FillRectangle(red, r.X, r.Y + stripe * i, r.Width, stripe);
            var cantonWidth = r.Width * 0.42F;
            var cantonHeight = stripe * 7F;
            using (var blue = new SolidBrush(Color.FromArgb(60, 59, 110))) g.FillRectangle(blue, r.X, r.Y, cantonWidth, cantonHeight);
            using (var stars = new SolidBrush(Color.White))
                for (var row = 0; row < 4; row++)
                    for (var column = 0; column < 5; column++)
                        g.FillEllipse(stars, r.X + 1.2F + column * 2.2F + (row % 2) * 0.55F, r.Y + 0.7F + row * 2F, 0.75F, 0.75F);
        }

        private static void DrawJapan(Graphics g, Rectangle r)
        {
            using (var white = new SolidBrush(Color.White)) g.FillRectangle(white, r);
            using (var red = new SolidBrush(Color.FromArgb(188, 0, 45)))
                g.FillEllipse(red, r.X + r.Width / 2F - 4.5F, r.Y + r.Height / 2F - 4.5F, 9F, 9F);
        }

        private static void DrawKorea(Graphics g, Rectangle r)
        {
            using (var white = new SolidBrush(Color.White)) g.FillRectangle(white, r);
            var cx = r.X + r.Width / 2F;
            var cy = r.Y + r.Height / 2F;
            var taegeuk = new RectangleF(cx - 4.5F, cy - 4.5F, 9F, 9F);
            using (var red = new SolidBrush(Color.FromArgb(205, 46, 58))) g.FillEllipse(red, taegeuk);
            using (var blue = new SolidBrush(Color.FromArgb(0, 71, 160)))
            {
                g.FillPie(blue, taegeuk.X, taegeuk.Y, taegeuk.Width, taegeuk.Height, 0, 180);
                g.FillEllipse(blue, cx, cy - 2.25F, 4.5F, 4.5F);
            }
            using (var red = new SolidBrush(Color.FromArgb(205, 46, 58))) g.FillEllipse(red, cx - 4.5F, cy - 2.25F, 4.5F, 4.5F);
            using (var black = new Pen(Color.Black, 1.15F))
            {
                DrawTrigram(g, black, r.X + 5.5F, r.Y + 4F, -34F, 0);
                DrawTrigram(g, black, r.Right - 5.5F, r.Y + 4F, 34F, 1);
                DrawTrigram(g, black, r.X + 5.5F, r.Bottom - 4F, 34F, 2);
                DrawTrigram(g, black, r.Right - 5.5F, r.Bottom - 4F, -34F, 3);
            }
        }

        private static void DrawSpain(Graphics g, Rectangle r)
        {
            using (var red = new SolidBrush(Color.FromArgb(170, 21, 27))) g.FillRectangle(red, r);
            using (var yellow = new SolidBrush(Color.FromArgb(241, 191, 0))) g.FillRectangle(yellow, r.X, r.Y + r.Height * 0.25F, r.Width, r.Height * 0.5F);
            using (var crestRed = new SolidBrush(Color.FromArgb(170, 21, 27))) g.FillRectangle(crestRed, r.X + 7.2F, r.Y + 7F, 2.6F, 4F);
            using (var crestBlue = new SolidBrush(Color.FromArgb(0, 60, 140))) g.FillRectangle(crestBlue, r.X + 8F, r.Y + 8F, 1F, 1.8F);
        }

        private static void DrawPortugal(Graphics g, Rectangle r)
        {
            var split = r.Width * 0.4F;
            using (var green = new SolidBrush(Color.FromArgb(4, 106, 56))) g.FillRectangle(green, r.X, r.Y, split, r.Height);
            using (var red = new SolidBrush(Color.FromArgb(218, 41, 28))) g.FillRectangle(red, r.X + split, r.Y, r.Width - split, r.Height);
            var cx = r.X + split;
            var cy = r.Y + r.Height / 2F;
            using (var yellow = new Pen(Color.FromArgb(255, 204, 0), 1.2F))
            {
                g.DrawEllipse(yellow, cx - 3.7F, cy - 3.7F, 7.4F, 7.4F);
                g.DrawLine(yellow, cx - 3.2F, cy, cx + 3.2F, cy);
                g.DrawLine(yellow, cx, cy - 3.2F, cx, cy + 3.2F);
            }
            using (var white = new SolidBrush(Color.White)) g.FillRectangle(white, cx - 2.2F, cy - 2.8F, 4.4F, 5.2F);
            using (var blue = new SolidBrush(Color.FromArgb(0, 77, 152))) g.FillEllipse(blue, cx - 1.1F, cy - 1F, 2.2F, 2.2F);
        }

        private static void DrawSaudiArabia(Graphics g, Rectangle r)
        {
            using (var green = new SolidBrush(Color.FromArgb(0, 108, 53))) g.FillRectangle(green, r);
            using (var white = new Pen(Color.White, 1.1F) { StartCap = LineCap.Round, EndCap = LineCap.Round })
            {
                // A compact calligraphic suggestion remains recognizable after
                // Office scales the 32 px source down in the menu.
                g.DrawLine(white, r.X + 7, r.Y + 5, r.X + 21, r.Y + 5);
                g.DrawLine(white, r.X + 8, r.Y + 7, r.X + 20, r.Y + 7);
                g.DrawLine(white, r.X + 9, r.Y + 9, r.X + 19, r.Y + 9);
                g.DrawLine(white, r.X + 7, r.Y + 13, r.X + 21, r.Y + 13);
                g.DrawLine(white, r.X + 20, r.Y + 12, r.X + 23, r.Y + 11.2F);
            }
        }

        private static void FillStar(Graphics g, Brush brush, float centerX, float centerY, float outerRadius, float rotationDegrees)
        {
            var points = new PointF[10];
            var innerRadius = outerRadius * 0.382F;
            for (var i = 0; i < points.Length; i++)
            {
                var radius = i % 2 == 0 ? outerRadius : innerRadius;
                var angle = (rotationDegrees + i * 36F) * Math.PI / 180D;
                points[i] = new PointF(centerX + (float)Math.Cos(angle) * radius, centerY + (float)Math.Sin(angle) * radius);
            }
            g.FillPolygon(brush, points);
        }

        private static void DrawTrigram(Graphics g, Pen pen, float centerX, float centerY, float angle, int brokenRow)
        {
            var state = g.Save();
            g.TranslateTransform(centerX, centerY);
            g.RotateTransform(angle);
            for (var row = 0; row < 3; row++)
            {
                var y = (row - 1) * 1.7F;
                if (row == brokenRow % 3)
                {
                    g.DrawLine(pen, -3F, y, -0.6F, y);
                    g.DrawLine(pen, 0.6F, y, 3F, y);
                }
                else g.DrawLine(pen, -3F, y, 3F, y);
            }
            g.Restore(state);
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
