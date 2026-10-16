namespace Dazzle.Pdf;

using OpenJade.Style;
using Symbol = OpenJade.Style.FOTBuilder.Symbol;

// The inherited characteristics a flow object is made under. Lengths are in DSSSL units
// (1pt = 1000). The initial values are those of ISO/IEC 10179 as the RTF backend of Jade
// has them, except for input-whitespace-treatment (see below).
public class PdfCharacteristics
{
    // Font
    public string FontFamily { get; set; } = "Helvetica";
    public long FontSize { get; set; } = 10000;
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

    // Indentation
    public long StartIndent { get; set; } = 0;
    public long EndIndent { get; set; } = 0;
    public long FirstLineStartIndent { get; set; } = 0;

    // Lines: the distance of the placement paths, exactly when min-leading is #f (the initial
    // value), at least that and the height of the line's content plus min-leading otherwise
    public long LineSpacing { get; set; } = 12000;
    public long? MinLeading { get; set; } = null;

    // Alignment
    public Symbol Quadding { get; set; } = Symbol.symbolStart;
    public Symbol DisplayAlignment { get; set; } = Symbol.symbolStart;

    // Page dimensions (set on page sequence level)
    public long PageWidth { get; set; } = 595276;
    public long PageHeight { get; set; } = 841890;
    public long LeftMargin { get; set; } = 72000;
    public long RightMargin { get; set; } = 72000;
    public long TopMargin { get; set; } = 72000;
    public long BottomMargin { get; set; } = 72000;
    public long HeaderMargin { get; set; } = 0;
    public long FooterMargin { get; set; } = 0;

    // Position point shift (positive = raised)
    public long PositionPointShift { get; set; } = 0;

    // Whitespace. The standard's initial input-whitespace-treatment is 'preserve; the backend
    // collapses unless told otherwise, as Dazzle.Pdf always has and as the Scaly engine does.
    public Symbol Lines { get; set; } = Symbol.symbolWrap;
    public Symbol InputWhitespaceTreatment { get; set; } = Symbol.symbolCollapse;

    // Line fields
    public long FieldWidth { get; set; } = 0;
    public Symbol FieldAlign { get; set; } = Symbol.symbolStart;

    // Rules and borders
    public long LineThickness { get; set; } = 1000;

    // Cell margins (table cells)
    public long CellBeforeRowMargin { get; set; } = 0;
    public long CellAfterRowMargin { get; set; } = 0;
    public long CellBeforeColumnMargin { get; set; } = 0;
    public long CellAfterColumnMargin { get; set; } = 0;
    public bool CellBackground { get; set; } = false;
    public Symbol CellRowAlignment { get; set; } = Symbol.symbolStart;

    // Page number format: "1" = arabic, "i" = lowercase roman, "I" = uppercase roman, "a", "A"
    public string PageNumberFormat { get; set; } = "1";

    // Score: the open score flow object ('before, 'through, 'after, or a length below the
    // placement path), and whether spaces are scored
    public Symbol ScoreType { get; set; } = Symbol.symbolFalse;
    public long? ScoreLength { get; set; } = null;
    public bool ScoreSpaces { get; set; } = true;

    // Dazzle extension characteristics (see PdfFOTBuilder.GetExtensions)
    public long MinRowHeight { get; set; } = 0;
    public long RowHeight { get; set; } = 0;

    public PdfCharacteristics Clone()
    {
        return (PdfCharacteristics)MemberwiseClone();
    }

    // Convert DSSSL length units to points (1pt = 1000 DSSSL units)
    public static double ToPoints(long dsssl)
    {
        return dsssl / 1000.0;
    }

    public double FontSizePt => ToPoints(FontSize);

    public bool IsBold => FontWeight == Symbol.symbolBold
                       || FontWeight == Symbol.symbolUltraBold
                       || FontWeight == Symbol.symbolSemiBold
                       || FontWeight == Symbol.symbolExtraBold;

    public bool IsItalic => FontPosture == Symbol.symbolItalic
                         || FontPosture == Symbol.symbolOblique;

    public bool Scored => ScoreType != Symbol.symbolFalse || ScoreLength.HasValue;

    // Verbatim: record ends are line breaks and spaces are kept
    public bool Verbatim => Lines == Symbol.symbolAsis || Lines == Symbol.symbolAsisWrap
                         || Lines == Symbol.symbolAsisTruncate
                         || InputWhitespaceTreatment == Symbol.symbolPreserve;

    // Whether lines are broken where they do not fit
    public bool Wraps => Lines == Symbol.symbolWrap || Lines == Symbol.symbolAsisWrap;
}
