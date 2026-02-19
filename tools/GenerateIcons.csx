// Generate Release Pack add-in icons at all 6 SolidWorks-required sizes.
// Run with: dotnet-script GenerateIcons.csx
// or:       dotnet run --project . -- generate-icons
//
// Output: Icons/ folder with main_20.png through main_128.png
//         and toolbar_20.png through toolbar_128.png

using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;

string outputDir = Path.Combine(AppContext.BaseDirectory, "Icons");
if (args.Length > 0 && Directory.Exists(args[0]))
    outputDir = args[0];

Directory.CreateDirectory(outputDir);

int[] sizes = { 20, 32, 40, 64, 96, 128 };

foreach (int size in sizes)
{
    // Main icon (gear + export arrow)
    using (var bmp = new Bitmap(size, size))
    using (var g = Graphics.FromImage(bmp))
    {
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.Clear(Color.Transparent);

        float pad = size * 0.08f;
        float s = size - pad * 2;

        // Gear circle (dark blue)
        using (var brush = new SolidBrush(Color.FromArgb(27, 58, 92)))
        {
            g.FillEllipse(brush, pad, pad, s, s);
        }

        // Inner circle cutout (white)
        float inner = s * 0.45f;
        float innerOff = pad + (s - inner) / 2;
        using (var brush = new SolidBrush(Color.White))
        {
            g.FillEllipse(brush, innerOff, innerOff, inner, inner);
        }

        // Export arrow (orange)
        float cx = size / 2f;
        float cy = size / 2f;
        float arrowSize = s * 0.25f;
        using (var pen = new Pen(Color.FromArgb(232, 125, 47), Math.Max(1.5f, size / 16f)))
        {
            pen.EndCap = LineCap.ArrowAnchor;
            g.DrawLine(pen, cx, cy + arrowSize * 0.3f, cx + arrowSize, cy - arrowSize * 0.7f);
        }

        bmp.Save(Path.Combine(outputDir, $"main_{size}.png"), System.Drawing.Imaging.ImageFormat.Png);
    }

    // Toolbar icon (same design — SolidWorks uses IconList for command buttons)
    using (var bmp = new Bitmap(size, size))
    using (var g = Graphics.FromImage(bmp))
    {
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.Clear(Color.Transparent);

        float pad = size * 0.1f;
        float s = size - pad * 2;

        // Document shape (rounded rect)
        using (var brush = new SolidBrush(Color.FromArgb(27, 58, 92)))
        {
            var rect = new RectangleF(pad, pad, s * 0.7f, s);
            g.FillRectangle(brush, rect);
        }

        // Arrow (orange, export)
        float arrowX = pad + s * 0.5f;
        float arrowY = size * 0.5f;
        float arrowLen = s * 0.35f;
        using (var pen = new Pen(Color.FromArgb(232, 125, 47), Math.Max(1.5f, size / 14f)))
        {
            pen.EndCap = LineCap.ArrowAnchor;
            g.DrawLine(pen, arrowX, arrowY, arrowX + arrowLen, arrowY);
        }

        bmp.Save(Path.Combine(outputDir, $"toolbar_{size}.png"), System.Drawing.Imaging.ImageFormat.Png);
    }

    Console.WriteLine($"Generated {size}x{size} icons");
}

Console.WriteLine($"Icons saved to: {outputDir}");
