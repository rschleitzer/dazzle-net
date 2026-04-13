namespace Dazzle.Pdf;

using OpenJade.Style;
using Symbol = OpenJade.Style.FOTBuilder.Symbol;

// Base class for all intermediate document nodes
public abstract class PdfNode
{
    public PdfCharacteristics Characteristics { get; set; }

    protected PdfNode(PdfCharacteristics characteristics)
    {
        Characteristics = characteristics.Clone();
    }
}

// Container node that can hold child nodes
public abstract class PdfContainerNode : PdfNode
{
    public List<PdfNode> Children { get; } = new();

    protected PdfContainerNode(PdfCharacteristics characteristics) : base(characteristics) { }

    public void Add(PdfNode child) => Children.Add(child);
}

// Root: a simple page sequence with header/footer slots and content
public class PdfPageSequence : PdfContainerNode
{
    // DSSSL defines 24 header/footer slots (4 page types × 6 parts)
    // For Phase 1, we store them but only render the basic ones
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

// Display group: vertical container for block-level content
public class PdfDisplayGroup : PdfContainerNode
{
    public PdfDisplayGroup(PdfCharacteristics characteristics) : base(characteristics) { }
}

// Scroll: vertical flow container (similar to display group for PDF purposes)
public class PdfScroll : PdfContainerNode
{
    public PdfScroll(PdfCharacteristics characteristics) : base(characteristics) { }
}

// Sequence: inline container
public class PdfSequence : PdfContainerNode
{
    public PdfSequence(PdfCharacteristics characteristics) : base(characteristics) { }
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

// Leaf: page number placeholder
public class PdfPageNumber : PdfNode
{
    public PdfPageNumber(PdfCharacteristics characteristics) : base(characteristics) { }
}

// Leaf: horizontal or vertical rule
public class PdfRule : PdfNode
{
    public Symbol Orientation { get; }
    public bool HasLength { get; }
    public long Length { get; }

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
