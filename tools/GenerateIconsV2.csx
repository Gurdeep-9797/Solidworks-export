using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

var sizes = new[] { 16, 20, 32, 40, 64, 96, 128 };
string outDir = Path.Combine(Environment.CurrentDirectory, "src/ReleasePack.AddIn/UI/Icons");
Directory.CreateDirectory(outDir);

foreach (var size in sizes)
{
    using (var bmp = new Bitmap(size, size))
    using (var g = Graphics.FromImage(bmp))
    {
        g.Clear(Color.Transparent);
        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;

        // Draw a 'Package' icon style background
        var rect = new Rectangle(1, 1, size - 2, size - 2);
        using (var brush = new SolidBrush(Color.FromArgb(40, 120, 200))) // SolidWorks Blue
        {
            g.FillRectangle(brush, rect);
        }
        
        using (var pen = new Pen(Color.White, Math.Max(1, size / 16f)))
        {
            g.DrawRectangle(pen, rect);
            g.DrawLine(pen, 1, 1, size-2, size-2);
            g.DrawLine(pen, 1, size-2, size-2, 1);
        }

        string path = Path.Combine(outDir, $"icon_{size}.png");
        bmp.Save(path, ImageFormat.Png);
        Console.WriteLine($"Generated: {path}");
    }
}
