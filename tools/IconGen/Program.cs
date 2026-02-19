using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;

namespace IconGen
{
    class Program
    {
        static void Main(string[] args)
        {
            string outDir = Path.GetFullPath(Path.Combine(Environment.CurrentDirectory, "../../src/ReleasePack.Installer")); // Root of installer proj
            Directory.CreateDirectory(outDir);

            using (var bmp = new Bitmap(600, 100))
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.HighQuality;
                
                // Gradient Background
                using (var brush = new LinearGradientBrush(new Rectangle(0, 0, 600, 100), 
                    Color.FromArgb(30, 30, 30), Color.FromArgb(50, 50, 60), LinearGradientMode.Horizontal))
                {
                    g.FillRectangle(brush, 0, 0, 600, 100);
                }
                
                // Text
                using (var font = new Font("Segoe UI", 24, FontStyle.Bold))
                {
                    g.DrawString("SolidWorks Release Pack", font, Brushes.White, new PointF(20, 20));
                }
                
                using (var font = new Font("Segoe UI", 10))
                {
                    g.DrawString("Professional Automation Suite", font, Brushes.LightGray, new PointF(24, 60));
                }

                // Abstract graphic
                using (var pen = new Pen(Color.FromArgb(40, 120, 200), 2))
                {
                    g.DrawLine(pen, 500, 20, 550, 80);
                    g.DrawLine(pen, 550, 80, 580, 50);
                    g.DrawEllipse(pen, 540, 40, 20, 20);
                }

                string path = Path.Combine(outDir, "Banner.png");
                bmp.Save(path, ImageFormat.Png);
                Console.WriteLine($"Generated: {path}");
            }
        }
    }
}
