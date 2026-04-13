namespace Dazzle.Pdf;

using OpenSP;
using OpenJade.Style;
using OpenJade.Grove;
using Char = System.UInt32;
using Symbol = OpenJade.Style.FOTBuilder.Symbol;
using LengthSpec = OpenJade.Style.FOTBuilder.LengthSpec;
using DeviceRGBColor = OpenJade.Style.FOTBuilder.DeviceRGBColor;
using DisplaySpace = OpenJade.Style.FOTBuilder.DisplaySpace;

public class PdfFOTBuilder : FOTBuilder
{
    private CmdLineApp app_;
    private string outputFilename_;

    // Characteristics stack: tracks current style state
    private PdfCharacteristics current_ = new();
    private Stack<PdfCharacteristics> characteristicsStack_ = new();

    // Document tree being built
    private System.Collections.Generic.List<PdfPageSequence> pageSequences_ = new();
    private PdfPageSequence? pageSequence_;
    private Stack<PdfContainerNode> containerStack_ = new();

    // Header/footer slot tracking: when set, content goes to HF slot instead of main content
    private uint? hfSlot_;


    public PdfFOTBuilder(CmdLineApp app, string outputFilename)
    {
        app_ = app;
        outputFilename_ = outputFilename;
    }

    // Current container we're adding content to (null if outside page sequence)
    private PdfContainerNode? CurrentContainer =>
        containerStack_.Count > 0 ? containerStack_.Peek()
        : pageSequence_;

    // Add a node: if inside a container (e.g. paragraph), add there.
    // Otherwise, if in HF mode, add to the HF slot.
    // Otherwise, add to the page sequence.
    private void AddNode(PdfNode node)
    {
        if (containerStack_.Count > 0)
            containerStack_.Peek().Add(node);
        else if (hfSlot_.HasValue && pageSequence_ != null)
            pageSequence_.HeaderFooter[hfSlot_.Value].Add(node);
        else
            pageSequence_?.Add(node);
    }

    // ==================== Page sequence ====================

    public override void startSimplePageSequence(FOTBuilder?[] headerFooter)
    {
        pageSequence_ = new PdfPageSequence(current_, nHF);

        // All header/footer content goes to this builder (serial mode)
        for (int i = 0; i < nHF; i++)
            headerFooter[i] = this;
    }

    public override void startSimplePageSequenceHeaderFooter(uint flags)
    {
        hfSlot_ = flags;
        characteristicsStack_.Push(current_.Clone());
    }

    public override void endSimplePageSequenceHeaderFooter(uint flags)
    {
        hfSlot_ = null;
        if (characteristicsStack_.Count > 0)
            current_ = characteristicsStack_.Pop();
    }

    public override void endAllSimplePageSequenceHeaderFooter()
    {
        // HF data is already stored in pageSequence_.HeaderFooter
    }

    public override void endSimplePageSequence()
    {
        if (pageSequence_ != null)
        {
            pageSequences_.Add(pageSequence_);
            pageSequence_ = null;
        }
        containerStack_.Clear();
    }

    public void Finish()
    {
        if (pageSequences_.Count == 0)
            return;
        var renderer = new PdfRenderer();
        renderer.Render(pageSequences_, outputFilename_);
    }

    // ==================== Block-level flow objects ====================

    // start()/end() are called by the engine around every flow object.
    // set* calls happen between start() and the specific startXxx() call.
    // We push characteristics on start() so set* modifications are scoped.
    // The DSSSL engine calls set* before startXxx, then children, then endXxx.
    // start()/end() are used for header/footer processing, not for wrapping every FO.
    // Each compound FO manages its own characteristics push/pop.

    public override void start()
    {
        characteristicsStack_.Push(current_.Clone());
    }

    public override void end()
    {
        if (characteristicsStack_.Count > 0)
            current_ = characteristicsStack_.Pop();
    }

    public override void startParagraph(ParagraphNIC nic)
    {
        ApplyDisplayNIC(nic);
        var para = new PdfParagraph(current_);
        AddNode(para);
        characteristicsStack_.Push(current_.Clone());
        containerStack_.Push(para);
    }

    public override void endParagraph()
    {
        if (containerStack_.Count > 0) containerStack_.Pop();
        if (characteristicsStack_.Count > 0) current_ = characteristicsStack_.Pop();
    }

    public override void startDisplayGroup(DisplayGroupNIC nic)
    {
        ApplyDisplayNIC(nic);
        var group = new PdfDisplayGroup(current_);
        AddNode(group);
        characteristicsStack_.Push(current_.Clone());
        containerStack_.Push(group);
    }

    public override void endDisplayGroup()
    {
        if (containerStack_.Count > 0) containerStack_.Pop();
        if (characteristicsStack_.Count > 0) current_ = characteristicsStack_.Pop();
    }

    public override void startScroll()
    {
        var scroll = new PdfScroll(current_);
        AddNode(scroll);
        characteristicsStack_.Push(current_.Clone());
        containerStack_.Push(scroll);
    }

    public override void endScroll()
    {
        if (containerStack_.Count > 0) containerStack_.Pop();
        if (characteristicsStack_.Count > 0) current_ = characteristicsStack_.Pop();
    }

    public override void startSequence()
    {
        var seq = new PdfSequence(current_);
        AddNode(seq);
        characteristicsStack_.Push(current_.Clone());
        containerStack_.Push(seq);
    }

    public override void endSequence()
    {
        if (containerStack_.Count > 0) containerStack_.Pop();
        if (characteristicsStack_.Count > 0) current_ = characteristicsStack_.Pop();
    }

    // ==================== Inline content ====================

    public override void characters(Char[] data, nuint size)
    {
        var sb = new System.Text.StringBuilder((int)size);
        for (nuint i = 0; i < size; i++)
            sb.Append((char)data[i]);
        AddNode(new PdfTextRun(sb.ToString(), current_));
    }

    public override void pageNumber()
    {
        AddNode(new PdfPageNumber(current_));
    }

    // ==================== Atomic flow objects ====================

    public override void rule(RuleNIC nic)
    {
        ApplyDisplayNIC(nic);
        AddNode(new PdfRule(current_, nic.orientation,
            nic.hasLength, nic.hasLength ? nic.length.length : 0));
    }

    public override void externalGraphic(ExternalGraphicNIC nic)
    {
        ApplyDisplayNIC(nic);
        AddNode(new PdfExternalGraphic(current_,
            nic.entitySystemId.ToString(),
            nic.isDisplay,
            nic.hasMaxWidth, nic.hasMaxWidth ? nic.maxWidth.length : 0,
            nic.hasMaxHeight, nic.hasMaxHeight ? nic.maxHeight.length : 0));
    }

    // ==================== Extension flow objects (Phase 2) ====================

    public override void extension(ExtensionFlowObj fo, NodePtr currentNode)
    {
        // Phase 2: dispatch to PDF-specific extension flow objects
    }

    public override void startExtension(ExtensionFlowObj fo, NodePtr currentNode,
        System.Collections.Generic.List<FOTBuilder> fotbs)
    {
        // Phase 2
    }

    public override void endExtension(ExtensionFlowObj fo)
    {
        // Phase 2
    }

    public static FOTBuilder.ExtensionTableEntry[] GetExtensions()
    {
        // Phase 2: barcode, html-content, page-overlay, attachment
        return System.Array.Empty<FOTBuilder.ExtensionTableEntry>();
    }

    // ==================== Characteristics (set* methods) ====================

    // Font
    public override void setFontSize(long size) => current_.FontSize = size;
    public override void setFontFamilyName(StringC name) => current_.FontFamily = name.ToString();
    public override void setFontWeight(Symbol weight) => current_.FontWeight = weight;
    public override void setFontPosture(Symbol posture) => current_.FontPosture = posture;

    // Color
    public override void setColor(DeviceRGBColor color)
    {
        current_.ColorR = color.red;
        current_.ColorG = color.green;
        current_.ColorB = color.blue;
    }

    public override void setBackgroundColor(DeviceRGBColor color)
    {
        current_.HasBackgroundColor = true;
        current_.BackgroundR = color.red;
        current_.BackgroundG = color.green;
        current_.BackgroundB = color.blue;
    }

    public override void setBackgroundColor()
    {
        current_.HasBackgroundColor = false;
    }

    // Indentation
    public override void setStartIndent(LengthSpec indent) => current_.StartIndent = indent.length;
    public override void setEndIndent(LengthSpec indent) => current_.EndIndent = indent.length;
    public override void setFirstLineStartIndent(LengthSpec indent) => current_.FirstLineStartIndent = indent.length;

    // Spacing
    public override void setLineSpacing(LengthSpec spacing) => current_.LineSpacing = spacing.length;

    // Alignment
    public override void setQuadding(Symbol quadding) => current_.Quadding = quadding;

    // Page dimensions
    public override void setPageWidth(long width) => current_.PageWidth = width;
    public override void setPageHeight(long height) => current_.PageHeight = height;
    public override void setLeftMargin(long margin) => current_.LeftMargin = margin;
    public override void setRightMargin(long margin) => current_.RightMargin = margin;
    public override void setTopMargin(long margin) => current_.TopMargin = margin;
    public override void setBottomMargin(long margin) => current_.BottomMargin = margin;

    // ==================== Helpers ====================

    private void ApplyDisplayNIC(FOTBuilder.DisplayNIC nic)
    {
        current_.SpaceBefore = nic.spaceBefore.nominal.length;
        current_.SpaceAfter = nic.spaceAfter.nominal.length;
        current_.KeepWithNext = nic.keepWithNext;
        current_.KeepWithPrevious = nic.keepWithPrevious;
        current_.BreakBefore = nic.breakBefore;
        current_.BreakAfter = nic.breakAfter;
    }
}
