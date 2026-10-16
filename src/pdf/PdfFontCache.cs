namespace Dazzle.Pdf;

using PdfSharp.Drawing;
using PdfSharp.Fonts;
using PdfSharp.Pdf;

// The fonts of a document: one XFont per family, size and style, the widths of the texts set in
// them, and the measures of a line. The typefaces come from the font resolver of PDFsharp: the
// caller's (GlobalFontSettings.FontResolver), else the system's.
internal sealed class PdfFontCache
{
    private readonly Dictionary<(string, double, bool, bool), XFont> fonts_ = new();
    private readonly Dictionary<(XFont, string), double> widths_ = new();
    private readonly XGraphics measure_ = XGraphics.CreateMeasureContext(new XSize(595, 842), XGraphicsUnit.Point, XPageDirection.Downwards);

    // Families tried, in this order, when the resolver has none for the family asked for
    private static readonly string[] Fallbacks = { "Arial", "Helvetica", "Liberation Sans", "DejaVu Sans", "Noto Sans", "Times New Roman" };

    private static bool resolverChecked_;

    public static void EnsureFontResolver()
    {
        if (resolverChecked_) return;
        resolverChecked_ = true;
        if (GlobalFontSettings.FontResolver != null) return;
        if (OperatingSystem.IsWindows())
            GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        else
            GlobalFontSettings.FontResolver = new MacFontResolver();
    }

    public XFont Font(PdfCharacteristics c)
    {
        var key = (c.FontFamily, c.FontSizePt, c.IsBold, c.IsItalic);
        if (fonts_.TryGetValue(key, out var font)) return font;
        var style = c.IsBold && c.IsItalic ? XFontStyleEx.BoldItalic
            : c.IsBold ? XFontStyleEx.Bold : c.IsItalic ? XFontStyleEx.Italic : XFontStyleEx.Regular;
        var options = new XPdfFontOptions(PdfFontEncoding.Unicode);
        Exception? first = null;
        foreach (var family in new[] { c.FontFamily }.Concat(Fallbacks))
        {
            try
            {
                font = new XFont(family, Math.Max(0.1, c.FontSizePt), style, options);
                break;
            }
            catch (Exception e)
            {
                first ??= e;
            }
        }
        if (font == null)
            throw new DssslProcessingException($"No font for the family \"{c.FontFamily}\": {first?.Message}");
        fonts_[key] = font;
        return font;
    }

    public double Width(string text, XFont font)
    {
        if (text.Length == 0) return 0;
        if (widths_.TryGetValue((font, text), out var w)) return w;
        w = measure_.MeasureString(text, font).Width;
        widths_[(font, text)] = w;
        return w;
    }

    private double PerEm(XFont font) => font.Size / font.FontFamily.GetEmHeight(font.Style);

    // The height of a line of the font set solid: the line spacing of its family
    public double LineHeight(XFont font) => font.FontFamily.GetLineSpacing(font.Style) * PerEm(font);

    public double Descent(XFont font) => font.FontFamily.GetCellDescent(font.Style) * PerEm(font);

    public double Ascent(XFont font) => font.FontFamily.GetCellAscent(font.Style) * PerEm(font);

    // Where the scores go: the underline below the placement path, its thickness, the line
    // through and the line over the text above the placement path (over the capitals)
    public (double Underline, double Thickness, double Strike, double Over) ScoreMetrics(XFont font)
    {
        var m = font.Metrics;
        var unit = font.Size / m.UnitsPerEm;
        var thickness = m.UnderlineThickness > 0 ? m.UnderlineThickness * unit : font.Size / 20;
        var underline = m.UnderlinePosition != 0 ? -m.UnderlinePosition * unit : font.Size / 10;
        var strike = m.StrikethroughPosition > 0 ? m.StrikethroughPosition * unit : font.Size * 0.3;
        var cap = m.CapHeight > 0 ? m.CapHeight * unit : font.Size * 0.7;
        return (underline, thickness, strike, cap + 2 * thickness);
    }
}
