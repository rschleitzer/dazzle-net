namespace Dazzle.Pdf;

using OpenJade.Style;
using Symbol = OpenJade.Style.FOTBuilder.Symbol;

public class PdfCharacteristics
{
    // Font
    public string FontFamily { get; set; } = "Helvetica";
    public long FontSize { get; set; } = 10000; // DSSSL units: 1pt = 1000
    public Symbol FontWeight { get; set; } = Symbol.symbolMedium;
    public Symbol FontPosture { get; set; } = Symbol.symbolUpright;

    // Color
    public byte ColorR { get; set; } = 0;
    public byte ColorG { get; set; } = 0;
    public byte ColorB { get; set; } = 0;
    public bool HasBackgroundColor { get; set; } = false;
    public byte BackgroundR { get; set; } = 255;
    public byte BackgroundG { get; set; } = 255;
    public byte BackgroundB { get; set; } = 255;

    // Indentation (DSSSL length units)
    public long StartIndent { get; set; } = 0;
    public long EndIndent { get; set; } = 0;
    public long FirstLineStartIndent { get; set; } = 0;

    // Spacing
    public long LineSpacing { get; set; } = 0; // 0 = auto
    public long SpaceBefore { get; set; } = 0;
    public long SpaceAfter { get; set; } = 0;

    // Alignment
    public Symbol Quadding { get; set; } = Symbol.symbolStart;

    // Page flow
    public bool KeepWithNext { get; set; } = false;
    public bool KeepWithPrevious { get; set; } = false;
    public Symbol BreakBefore { get; set; } = Symbol.symbolFalse;
    public Symbol BreakAfter { get; set; } = Symbol.symbolFalse;

    // Page dimensions (set on page sequence level)
    public long PageWidth { get; set; } = 595000;  // A4 width in DSSSL units (~210mm)
    public long PageHeight { get; set; } = 842000; // A4 height in DSSSL units (~297mm)
    public long LeftMargin { get; set; } = 72000;  // ~1 inch
    public long RightMargin { get; set; } = 72000;
    public long TopMargin { get; set; } = 72000;
    public long BottomMargin { get; set; } = 72000;

    // Position point shift (positive = superscript, negative = subscript)
    public long PositionPointShift { get; set; } = 0;

    // Verbatim mode: symbolAsis preserves whitespace/linebreaks (programlisting)
    public Symbol Lines { get; set; } = Symbol.symbolWrap;
    public Symbol InputWhitespaceTreatment { get; set; } = Symbol.symbolCollapse;

    // Page number format: "1" = arabic, "i" = lowercase roman, "I" = uppercase roman
    public string PageNumberFormat { get; set; } = "1";

    public PdfCharacteristics Clone()
    {
        return (PdfCharacteristics)MemberwiseClone();
    }

    // Convert DSSSL length units to points (1pt = 1000 DSSSL units)
    public static float ToPoints(long dsssl)
    {
        return dsssl / 1000f;
    }

    public float FontSizePt => ToPoints(FontSize);
    public float StartIndentPt => ToPoints(StartIndent);
    public float EndIndentPt => ToPoints(EndIndent);
    public float FirstLineStartIndentPt => ToPoints(FirstLineStartIndent);
    public float LineSpacingPt => ToPoints(LineSpacing);
    public float SpaceBeforePt => ToPoints(SpaceBefore);
    public float SpaceAfterPt => ToPoints(SpaceAfter);
    public float PageWidthPt => ToPoints(PageWidth);
    public float PageHeightPt => ToPoints(PageHeight);
    public float LeftMarginPt => ToPoints(LeftMargin);
    public float RightMarginPt => ToPoints(RightMargin);
    public float TopMarginPt => ToPoints(TopMargin);
    public float BottomMarginPt => ToPoints(BottomMargin);

    public bool IsBold => FontWeight == Symbol.symbolBold
                       || FontWeight == Symbol.symbolUltraBold
                       || FontWeight == Symbol.symbolSemiBold;

    public bool IsItalic => FontPosture == Symbol.symbolItalic
                         || FontPosture == Symbol.symbolOblique;
}
