namespace Dazzle.Pdf;

using PdfSharp.Drawing;
using PdfSharp.Pdf;

// Writes the pages PdfLayout sets. The layout runs twice: the first pass finds the page every
// element begins on, the second sets the pages again with the page numbers of
// current-node-page-number filled in, and those are written.
public class PdfRenderer
{
    /// <summary>Where every element begins, by the name of its location mark, after <see cref="Render"/>.</summary>
    public IReadOnlyDictionary<string, PdfElementLocation> Locations { get; private set; } = new Dictionary<string, PdfElementLocation>();

    public void Render(List<PdfPageSequence> pageSequences, Stream stream)
    {
        PdfFontCache.EnsureFontResolver();
        var fonts = new PdfFontCache();
        var sizes = new Dictionary<string, (double, double)?>();
        (double, double)? ImageSize(string path)
        {
            if (sizes.TryGetValue(path, out var known)) return known;
            (double, double)? size = null;
            if (File.Exists(path))
            {
                try
                {
                    using var image = XImage.FromFile(path);
                    size = (image.PointWidth, image.PointHeight);
                }
                catch (Exception)
                {
                    // not an image PDFsharp can read: the graphic is not set
                }
            }
            sizes[path] = size;
            return size;
        }

        var first = new PdfLayout(new Dictionary<string, int>(), fonts, ImageSize);
        first.Run(pageSequences);
        var layout = new PdfLayout(first.ElementPages, fonts, ImageSize);
        layout.Run(pageSequences);
        Locations = layout.ElementLocations;
        Write(layout.Pages, stream);
    }

    private static void Write(List<PdfLayout.Page> pages, Stream stream)
    {
        var doc = new PdfDocument();
        var images = new Dictionary<string, XImage>();
        try
        {
            foreach (var p in pages)
            {
                var page = doc.AddPage();
                page.Width = XUnit.FromPoint(p.Width);
                page.Height = XUnit.FromPoint(p.Height);
                using var gfx = XGraphics.FromPdfPage(page);
                foreach (var op in p.Ops)
                {
                    switch (op.Kind)
                    {
                        case PdfLayout.OpKind.Text:
                            gfx.DrawString(op.Text!, op.Font!, new XSolidBrush(op.Color), op.X, op.Y, XStringFormats.BaseLineLeft);
                            break;
                        case PdfLayout.OpKind.Line:
                            if (op.Thickness > 0)
                                gfx.DrawLine(new XPen(op.Color, op.Thickness), op.X, op.Y, op.X2, op.Y2);
                            break;
                        case PdfLayout.OpKind.Rect:
                            gfx.DrawRectangle(new XSolidBrush(op.Color), op.X, op.Y, op.W, op.H);
                            break;
                        case PdfLayout.OpKind.Image:
                            if (!images.TryGetValue(op.Path!, out var image))
                            {
                                image = XImage.FromFile(op.Path!);
                                images[op.Path!] = image;
                            }
                            gfx.DrawImage(image, op.X, op.Y, op.W, op.H);
                            break;
                    }
                }
            }
            PageLabels(doc, pages);
            doc.Save(stream, false);
        }
        finally
        {
            foreach (var image in images.Values)
                image.Dispose();
        }
    }

    // A page label where the format of the page numbers changes, so that a viewer shows the
    // numbers as the pages do
    private static void PageLabels(PdfDocument doc, List<PdfLayout.Page> pages)
    {
        if (pages.Count == 0 || pages.All(p => p.NumberFormat == "1")) return;
        var nums = new PdfArray();
        string? previous = null;
        for (int i = 0; i < pages.Count; i++)
        {
            var format = pages[i].NumberFormat;
            if (format == previous) continue;
            var label = new PdfDictionary();
            label.Elements.Add("/S", new PdfName(format switch
            {
                "i" => "/r",
                "I" => "/R",
                "a" => "/a",
                "A" => "/A",
                _ => "/D"
            }));
            label.Elements.Add("/St", new PdfInteger(pages[i].Number));
            nums.Elements.Add(new PdfInteger(i));
            nums.Elements.Add(label);
            previous = format;
        }
        var labels = new PdfDictionary();
        labels.Elements.Add("/Nums", nums);
        doc.Internals.Catalog.Elements.Add("/PageLabels", labels);
    }
}
