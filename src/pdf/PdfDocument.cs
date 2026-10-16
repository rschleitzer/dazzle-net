namespace Dazzle.Pdf;

using OpenJade.Style;
using Symbol = OpenJade.Style.FOTBuilder.Symbol;

// The flow object tree the FOT builder collects and PdfLayout sets. Every node keeps a snapshot
// of the inherited characteristics it was made under.
public abstract class PdfNode
{
    public PdfCharacteristics Characteristics { get; set; }

    protected PdfNode(PdfCharacteristics characteristics)
    {
        Characteristics = characteristics.Clone();
    }
}

// The display characteristics of a block (DisplayNIC)
public class PdfDisplay
{
    public long SpaceBefore;
    public bool SpaceBeforeConditional = true;
    public long SpaceAfter;
    public bool SpaceAfterConditional = true;
    public bool KeepWithNext;
    public bool KeepWithPrevious;
    public bool BreakBefore;
    public bool BreakAfter;
}

// Container node that can hold child nodes
public abstract class PdfContainerNode : PdfNode
{
    public List<PdfNode> Children { get; } = new();
    public PdfDisplay Display { get; set; } = new();

    protected PdfContainerNode(PdfCharacteristics characteristics) : base(characteristics) { }

    public void Add(PdfNode child) => Children.Add(child);
}

// Root: a simple page sequence with header/footer slots and content
public class PdfPageSequence : PdfContainerNode
{
    // 24 header/footer slots (FOTBuilder.HF flags)
    public List<PdfNode>[] HeaderFooter { get; }

    public PdfPageSequence(PdfCharacteristics characteristics, int headerFooterSlots)
        : base(characteristics)
    {
        HeaderFooter = new List<PdfNode>[headerFooterSlots];
        for (int i = 0; i < headerFooterSlots; i++)
            HeaderFooter[i] = new();
    }
}

// Block-level paragraph containing inline runs
public class PdfParagraph : PdfContainerNode
{
    public PdfParagraph(PdfCharacteristics characteristics) : base(characteristics) { }
}

// Display group or display box: vertical container for block-level content
public class PdfDisplayGroup : PdfContainerNode
{
    public PdfDisplayGroup(PdfCharacteristics characteristics) : base(characteristics) { }
}

// Scroll: vertical flow container
public class PdfScroll : PdfContainerNode
{
    public PdfScroll(PdfCharacteristics characteristics) : base(characteristics) { }
}

// Sequence, link, score, inline box: inline container
public class PdfSequence : PdfContainerNode
{
    public PdfSequence(PdfCharacteristics characteristics) : base(characteristics) { }
}

// Line field: inline container at least field-width wide
public class PdfLineField : PdfContainerNode
{
    public PdfLineField(PdfCharacteristics characteristics) : base(characteristics) { }
}

// Leader: inline fill with a pattern
public class PdfLeader : PdfContainerNode
{
    public PdfLeader(PdfCharacteristics characteristics) : base(characteristics) { }
}

// Leaf: a run of text with a style snapshot
public class PdfTextRun : PdfNode
{
    public string Text { get; }

    public PdfTextRun(string text, PdfCharacteristics characteristics) : base(characteristics)
    {
        Text = text;
    }
}

// Leaf: page number of the page the text is set on
public class PdfPageNumber : PdfNode
{
    public PdfPageNumber(PdfCharacteristics characteristics) : base(characteristics) { }
}

// Leaf: page number of the page an element begins on
public class PdfNodePageNumber : PdfNode
{
    public string LocationName { get; }

    public PdfNodePageNumber(PdfCharacteristics characteristics, string locationName)
        : base(characteristics)
    {
        LocationName = locationName;
    }
}

// Leaf: the place where an element begins
public class PdfLocationMark : PdfNode
{
    public string LocationName { get; }

    public PdfLocationMark(string locationName)
        : base(new PdfCharacteristics())
    {
        LocationName = locationName;
    }
}

// Leaf: paragraph-break, the paragraph goes on with a new first line
public class PdfParagraphBreak : PdfNode
{
    public PdfParagraphBreak(PdfCharacteristics characteristics) : base(characteristics) { }
}

// Leaf: horizontal or vertical rule
public class PdfRule : PdfNode
{
    public Symbol Orientation { get; }
    public bool HasLength { get; }
    public long Length { get; }
    public PdfDisplay Display { get; set; } = new();

    public PdfRule(PdfCharacteristics characteristics, Symbol orientation, bool hasLength, long length)
        : base(characteristics)
    {
        Orientation = orientation;
        HasLength = hasLength;
        Length = length;
    }
}

// Leaf: external image
public class PdfExternalGraphic : PdfNode
{
    public string SystemId { get; }
    public bool IsDisplay { get; }
    public bool HasMaxWidth { get; }
    public long MaxWidth { get; }
    public bool HasMaxHeight { get; }
    public long MaxHeight { get; }
    public PdfDisplay Display { get; set; } = new();

    public PdfExternalGraphic(PdfCharacteristics characteristics, string systemId,
        bool isDisplay, bool hasMaxWidth, long maxWidth, bool hasMaxHeight, long maxHeight)
        : base(characteristics)
    {
        SystemId = systemId;
        IsDisplay = isDisplay;
        HasMaxWidth = hasMaxWidth;
        MaxWidth = maxWidth;
        HasMaxHeight = hasMaxHeight;
        MaxHeight = maxHeight;
    }
}

// Border specification for an edge, captured from a table-border flow object
public class PdfBorder
{
    public long LineThickness { get; set; } = 1000;
    public byte ColorR { get; set; }
    public byte ColorG { get; set; }
    public byte ColorB { get; set; }

    public double LineThicknessPt => PdfCharacteristics.ToPoints(LineThickness);
}

// Table: columns, header rows, body rows, footer rows, and the borders of the table
public class PdfTable : PdfNode
{
    public List<PdfTableColumn> Columns { get; } = new();
    public List<PdfTableRow> HeaderRows { get; } = new();
    public List<PdfTableRow> BodyRows { get; } = new();
    public List<PdfTableRow> FooterRows { get; } = new();
    public PdfDisplay Display { get; set; } = new();
    public bool HasWidth { get; set; }
    public long Width { get; set; }

    public PdfBorder? BeforeRowBorder { get; set; }
    public PdfBorder? AfterRowBorder { get; set; }
    public PdfBorder? BeforeColumnBorder { get; set; }
    public PdfBorder? AfterColumnBorder { get; set; }

    public PdfTable(PdfCharacteristics characteristics) : base(characteristics) { }
}

// Column definition within a table
public class PdfTableColumn
{
    public uint ColumnIndex { get; set; }
    public uint NColumnsSpanned { get; set; } = 1;
    public bool HasWidth { get; set; }
    public long Width { get; set; }
    public double TableUnitFactor { get; set; }
}

// Row in a table (contains cells)
public class PdfTableRow : PdfContainerNode
{
    public PdfTableRow(PdfCharacteristics characteristics) : base(characteristics) { }
}

// Cell in a table row (contains flow content)
public class PdfTableCell : PdfContainerNode
{
    public uint ColumnIndex { get; }
    public uint NColumnsSpanned { get; }
    public uint NRowsSpanned { get; }

    public PdfBorder? BeforeRowBorder { get; set; }
    public PdfBorder? AfterRowBorder { get; set; }
    public PdfBorder? BeforeColumnBorder { get; set; }
    public PdfBorder? AfterColumnBorder { get; set; }

    public PdfTableCell(PdfCharacteristics characteristics, uint columnIndex,
        uint nColumnsSpanned, uint nRowsSpanned)
        : base(characteristics)
    {
        ColumnIndex = columnIndex;
        NColumnsSpanned = nColumnsSpanned;
        NRowsSpanned = nRowsSpanned;
    }
}
