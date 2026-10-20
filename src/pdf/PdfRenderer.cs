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
        var balance = Balance(pageSequences, first, fonts, ImageSize);
        var layout = new PdfLayout(first.ElementPages, fonts, ImageSize, balance);
        layout.Run(pageSequences);
        Locations = layout.ElementLocations;
        Write(layout.Pages, stream);
    }

    // Balanced columns: on the last page of a column set with balance-columns?, the columns but the
    // last end at a height, the smallest at which the last column is not longer, so that the columns
    // are about equally long, the first ones rather the longer. That height is found by setting the
    // document again, one set after the other in the order they begin: from the content divided by
    // the number of columns upwards, by half of what the last column is too long, at least 0.5 pt.
    // A set with a column break on its last page is not balanced, and neither is one the height
    // would move to another page.
    private static Dictionary<int, (int Page, double Height)> Balance(List<PdfPageSequence> pageSequences,
        PdfLayout first, PdfFontCache fonts, Func<string, (double, double)?> imageSize)
    {
        var balance = new Dictionary<int, (int Page, double Height)>();
        var known = first.ElementPages;
        var ends = first.ColumnSetEnds;
        for (int set = 0; set < ends.Count; set++)
        {
            var end = ends[set];
            int n = end.Ends.Length;
            if (!end.Balance || end.HardBreak || end.Page < 0 || n < 2) continue;
            var content = end.Ends.Sum(e => e - end.Top);
            if (content <= 0) continue;
            var height = content / n;
            for (int round = 0; ; round++)
            {
                balance[set] = (end.Page, height);
                var trial = new PdfLayout(known, fonts, imageSize, balance);
                trial.Run(pageSequences);
                var result = trial.ColumnSetEnds.Count > set ? trial.ColumnSetEnds[set] : null;
                if (result == null || result.Page != end.Page || result.HardBreak || round == 40)
                {
                    // the sets after this one stay where the layout without this height has them
                    balance.Remove(set);
                    break;
                }
                var last = result.Ends[n - 1] - result.Top;
                if (last <= height + 0.001)
                {
                    ends = trial.ColumnSetEnds;
                    break;
                }
                height += Math.Max(0.5, (last - height) / 2);
            }
        }
        return balance;
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
                            if (op.Spacing == 0)
                                gfx.DrawString(op.Text!, op.Font!, new XSolidBrush(op.Color), op.X, op.Y, XStringFormats.BaseLineLeft);
                            else
                            {
                                // letter-spacing: each character on its own, the spacing after it
                                var brush = new XSolidBrush(op.Color);
                                var cx = op.X;
                                foreach (var rune in op.Text!.EnumerateRunes())
                                {
                                    var s = rune.ToString();
                                    gfx.DrawString(s, op.Font!, brush, cx, op.Y, XStringFormats.BaseLineLeft);
                                    cx += gfx.MeasureString(s, op.Font!).Width + op.Spacing;
                                }
                            }
                            break;
                        case PdfLayout.OpKind.Line:
                            if (op.Thickness > 0)
                                gfx.DrawLine(Pen(op), op.X, op.Y, op.X2, op.Y2);
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
    // The pen of a line. line-dash goes to PDFsharp relative to the width of the pen; a pattern of odd length is taken twice, so
    // that it repeats with dashes and gaps alternating as in PostScript, where PDFsharp would add a gap of its own.
    private static XPen Pen(PdfLayout.Op op)
    {
        var pen = new XPen(op.Color, op.Thickness);
        if (op.Dash is { Length: > 0 } dash && dash.Any(d => d > 0))
        {
            var pattern = dash.Length % 2 == 1 ? [.. dash, .. dash] : dash;
            pen.DashStyle = XDashStyle.Custom;
            pen.DashPattern = Array.ConvertAll(pattern, d => d / op.Thickness);
            pen.DashOffset = op.DashOffset / op.Thickness;
        }
        return pen;
    }

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
