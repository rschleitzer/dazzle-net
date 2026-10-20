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
    // min-post-line-spacing: the space below the placement path of a line is at least this, the
    // rest of the line's height is above (#f: the descent of its text, the rest above)
    public long? MinPostLineSpacing { get; set; } = null;
    // min-pre-line-spacing: the space above the placement path of a line is at least this, the rest of the line's height is below
    // (#f: as min-post-line-spacing and the descent leave it)
    public long? MinPreLineSpacing { get; set; } = null;

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

    // The page model (12.6.4, 12.6.5): the models of a page-sequence and of a column-set-sequence
    public PageModel[] InitialPageModels { get; set; } = System.Array.Empty<PageModel>();
    public PageModel[] RepeatPageModels { get; set; } = System.Array.Empty<PageModel>();
    public ColumnSetModel? ColumnSetModel { get; set; } = null;
    public long Span { get; set; } = 1;

    // Dazzle extension characteristics (see PdfFOTBuilder.GetExtensions)
    public long MinRowHeight { get; set; } = 0;
    public long RowHeight { get; set; } = 0;
    // grows-line?: whether an inline graphic makes its line taller where it does not fit (#f: it reaches beyond, as a word
    // processor's form field box does)
    public bool GrowsLine { get; set; } = true;
    // line-dash: the lengths of dashes and gaps, alternately, and the offset into them where a line starts; null is solid
    public long[]? LineDash { get; set; } = null;
    public long LineDashOffset { get; set; } = 0;
    // letter-spacing: space added after every character of the text, also after the last and after a space (negative narrows),
    // as a word processor's character spacing
    public long LetterSpacing { get; set; } = 0;
    // text-line-height and text-descent: the height of a line and the part of it below the placement path that this text needs,
    // instead of what its font's metrics say (0: the font's), as word processors give some fonts; line-height-factor: in
    // thousandths, not 0: every line of the paragraph is as high as its own text needs times this, the rest below the placement path
    public long TextLineHeight { get; set; } = 0;
    public long TextDescent { get; set; } = 0;
    public long LineHeightFactor { get; set; } = 0;
    // column-set and balance-columns?: what a display-group sets up its columns with; ColumnSetHere
    // says that the flow object being started specified column-set itself, not inherited it
    public string? ColumnSet { get; set; } = null;
    public bool BalanceColumns { get; set; } = false;
    public bool ColumnSetHere { get; set; } = false;
    // first-page-top-margin and first-page-bottom-margin of a simple-page-sequence: the margins of
    // its first page where they differ from the others (0: as the others)
    public long FirstPageTopMargin { get; set; } = 0;
    public long FirstPageBottomMargin { get; set; } = 0;

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
