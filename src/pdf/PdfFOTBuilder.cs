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

    public PdfFOTBuilder(Messenger messenger, string outputFilename)
    {
        app_ = null!;
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

    // True when the current container is a table row (sequences should not be pushed)
    private bool InTableRow => containerStack_.Count > 0 && containerStack_.Peek() is PdfTableRow;

    // ==================== Page sequence ====================

    public override void startSimplePageSequence(FOTBuilder?[] headerFooter)
    {
        characteristicsStack_.Push(current_.Clone());
        pageSequence_ = new PdfPageSequence(current_, nHF);

        // Adopt any containers started before this page sequence (e.g. scroll wrapping body)
        if (containerStack_.Count > 0)
        {
            var stack = containerStack_.ToArray(); // top-first order
            var root = stack[stack.Length - 1];    // bottom = outermost container
            pageSequence_.Add(root);
        }

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
        end();
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
        end();
    }

    public void Finish()
    {
        if (pageSequences_.Count == 0)
            return;
        var renderer = new PdfRenderer();
        renderer.Render(pageSequences_, outputFilename_);
    }

    public void Finish(Stream stream)
    {
        if (pageSequences_.Count == 0)
            return;
        var renderer = new PdfRenderer();
        renderer.Render(pageSequences_, stream);
    }

    // ==================== Block-level flow objects ====================

    // The engine calls pushStyle (which invokes set* on the FOTBuilder),
    // then processInner (which calls startXxx, children, endXxx), then popStyle.
    // The base class startXxx/endXxx default to calling start()/end().
    // end() uses RTF-style "remove-top, copy-from-new-top" semantics so that
    // set* modifications are scoped to the current FO and don't leak into siblings.

    public override void start()
    {
        characteristicsStack_.Push(current_.Clone());
    }

    public override void end()
    {
        if (characteristicsStack_.Count > 0)
        {
            // RTF semantics: remove top, then restore a copy from the new top.
            // This ensures set* modifications are scoped to the current FO
            // and don't leak into subsequent siblings.
            characteristicsStack_.Pop();
            if (characteristicsStack_.Count > 0)
                current_ = characteristicsStack_.Peek().Clone();
        }
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
        end();
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
        end();
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
        end();
    }

    public override void startSequence()
    {
        if (InTableRow)
        {
            // Inside a table row, the DSSSL engine wraps each cell in an implicit
            // sequence. Skip it — only table-cell nodes belong as row children.
            characteristicsStack_.Push(current_.Clone());
            return;
        }

        // Sequences are inline pass-through containers in DSSSL — they don't have
        // their own display spacing. Clear inherited SpaceBefore/SpaceAfter so they
        // don't inflate inter-element gaps at the rendering level.
        var seqChars = current_.Clone();
        seqChars.SpaceBefore = 0;
        seqChars.SpaceAfter = 0;
        var seq = new PdfSequence(seqChars);
        AddNode(seq);
        characteristicsStack_.Push(current_.Clone());
        containerStack_.Push(seq);
    }

    public override void endSequence()
    {
        if (containerStack_.Count > 0 && containerStack_.Peek() is not PdfSequence)
        {
            // Matching the skipped startSequence inside a table row
            end();
            return;
        }
        if (containerStack_.Count > 0) containerStack_.Pop();
        end();
    }

    // ==================== Inline content ====================

    public override void characters(Char[] data, nuint size)
    {
        bool verbatim = current_.Lines == Symbol.symbolAsis
            || current_.InputWhitespaceTreatment == Symbol.symbolPreserve;

        var sb = new System.Text.StringBuilder((int)size);
        for (nuint i = 0; i < size; i++)
        {
            char c = (char)data[i];
            if (verbatim)
            {
                // Preserve whitespace: keep CR/LF as newlines
                if (c == '\r')
                {
                    if (i + 1 < size && data[i + 1] == '\n')
                        i++; // skip CR in CR+LF
                    sb.Append('\n');
                }
                else
                    sb.Append(c);
            }
            else
            {
                // Normal mode: collapse record-ends and runs of whitespace
                if (c == '\r')
                    continue;
                if (c == '\n' || c == ' ' || c == '\t')
                {
                    // Collapse any run of whitespace (spaces, tabs, newlines) into one space
                    sb.Append(' ');
                    while (i + 1 < size)
                    {
                        char next = (char)data[i + 1];
                        if (next == ' ' || next == '\t' || next == '\n' || next == '\r')
                            i++;
                        else
                            break;
                    }
                }
                else
                    sb.Append(c);
            }
        }
        var text = sb.ToString();
        if (text.Length > 0)
            AddNode(new PdfTextRun(text, current_));
    }

    public override void pageNumber()
    {
        AddNode(new PdfPageNumber(current_));
    }

    // Position point shift (superscript/subscript via DSSSL characteristic)
    public override void setPositionPointShift(LengthSpec shift) =>
        current_.PositionPointShift = shift.length;

    public override void startLeader(LeaderNIC nic)
    {
        var leader = new PdfLeader(current_);
        AddNode(leader);
        characteristicsStack_.Push(current_.Clone());
        containerStack_.Push(leader);
    }

    public override void endLeader()
    {
        if (containerStack_.Count > 0) containerStack_.Pop();
        end();
    }

    // ==================== Node tracking & cross-references ====================

    public override void startNode(NodePtr node, StringC processingMode)
    {
        if (processingMode.size() == 0 && node != null)
        {
            var name = GetLocationName(node);
            if (name != null)
                AddNode(new PdfLocationMark(name));
        }
    }

    public override void currentNodePageNumber(NodePtr node)
    {
        start();
        var name = GetLocationName(node);
        if (name != null)
            AddNode(new PdfNodePageNumber(current_, name));
        end();
    }

    private static string? GetLocationName(NodePtr node)
    {
        GroveString id = new GroveString();
        if (node.getId(ref id) == OpenJade.Grove.AccessResult.accessOK && id.size() > 0)
        {
            var sb = new System.Text.StringBuilder("n");
            sb.Append(node.groveIndex());
            sb.Append('_');
            for (nuint i = 0; i < id.size(); i++)
                sb.Append((char)id[i]);
            return sb.ToString();
        }
        ulong idx = 0;
        if (node.elementIndex(ref idx) == OpenJade.Grove.AccessResult.accessOK)
            return $"n{node.groveIndex()}_{idx}";
        return null;
    }

    // ==================== Atomic flow objects ====================

    // Cell border ports: callback fires with characteristics (color, line-thickness) already
    // set on current_ via set* calls from the border sosofo (when it's a table-border FO).
    private enum CellBorderPort { BeforeRow, AfterRow, BeforeColumn, AfterColumn }
    private long lineThickness_ = 500; // DSSSL units; 0.5pt default

    // True when setColor or setLineThickness has been called since the last border capture.
    // Used to distinguish "undefined" border ports (no DSSSL characteristic set) from
    // defined ones (where the engine pushes characteristics before the port callback).
    private bool borderCharacteristicsFresh_ = false;

    // Saved state at startTableCell, so we can undo border-port mutations to current_.
    // (openjade-net's StyleStack.pop does not invoke set* to restore FOT builder state.)
    private PdfCharacteristics? cellBaseChars_;
    private long cellBaseLineThickness_;

    public override void setLineThickness(long thickness)
    {
        lineThickness_ = thickness;
        borderCharacteristicsFresh_ = true;
    }

    public override void rule(RuleNIC nic)
    {
        start();
        ApplyDisplayNIC(nic);
        AddNode(new PdfRule(current_, nic.orientation,
            nic.hasLength, nic.hasLength ? nic.length.length : 0));
        end();
    }

    private PdfTableCell? FindEnclosingCell()
    {
        foreach (var node in containerStack_)
            if (node is PdfTableCell c) return c;
        return null;
    }

    public override void externalGraphic(ExternalGraphicNIC nic)
    {
        start();
        ApplyDisplayNIC(nic);
        AddNode(new PdfExternalGraphic(current_,
            nic.entitySystemId.ToString(),
            nic.isDisplay,
            nic.hasMaxWidth, nic.hasMaxWidth ? nic.maxWidth.length : 0,
            nic.hasMaxHeight, nic.hasMaxHeight ? nic.maxHeight.length : 0));
        end();
    }

    // ==================== Table flow objects ====================

    private PdfTable? currentTable_;
    private enum TablePartSection { Body, Header, Footer }
    private TablePartSection tablePartSection_ = TablePartSection.Body;
    private PdfTableRow? currentTableRow_;

    public override void startTable(TableNIC nic)
    {
        ApplyDisplayNIC(nic);
        var table = new PdfTable(current_);
        AddNode(table);
        characteristicsStack_.Push(current_.Clone());
        currentTable_ = table;
    }

    public override void endTable()
    {
        currentTable_ = null;
        end();
    }

    public override void tableColumn(TableColumnNIC nic)
    {
        currentTable_?.Columns.Add(new PdfTableColumn
        {
            ColumnIndex = nic.columnIndex,
            NColumnsSpanned = nic.nColumnsSpanned,
            HasWidth = nic.hasWidth,
            Width = nic.hasWidth ? nic.width.length : 0,
            TableUnitFactor = nic.width.tableUnitFactor
        });
    }

    public override void startTablePartSerial(TablePartNIC nic)
    {
        characteristicsStack_.Push(current_.Clone());
        tablePartSection_ = TablePartSection.Body;
    }

    public override void endTablePartSerial() { end(); }

    public override void startTablePartHeader()
    {
        characteristicsStack_.Push(current_.Clone());
        tablePartSection_ = TablePartSection.Header;
    }

    public override void endTablePartHeader()
    {
        tablePartSection_ = TablePartSection.Body;
        end();
    }

    public override void startTablePartFooter()
    {
        characteristicsStack_.Push(current_.Clone());
        tablePartSection_ = TablePartSection.Footer;
    }

    public override void endTablePartFooter()
    {
        tablePartSection_ = TablePartSection.Body;
        end();
    }

    public override void startTableRow()
    {
        characteristicsStack_.Push(current_.Clone());
        var row = new PdfTableRow(current_);
        currentTableRow_ = row;
        containerStack_.Push(row);

        if (currentTable_ != null)
        {
            switch (tablePartSection_)
            {
                case TablePartSection.Header: currentTable_.HeaderRows.Add(row); break;
                case TablePartSection.Footer: currentTable_.FooterRows.Add(row); break;
                default: currentTable_.BodyRows.Add(row); break;
            }
        }
    }

    public override void endTableRow()
    {
        if (containerStack_.Count > 0) containerStack_.Pop();
        currentTableRow_ = null;
        end();
    }

    public override void startTableCell(TableCellNIC nic)
    {
        characteristicsStack_.Push(current_.Clone());
        // Snapshot state before any border-port mutations (the engine pushes border
        // characteristics but doesn't unwind them after the port callback).
        cellBaseChars_ = current_.Clone();
        cellBaseLineThickness_ = lineThickness_;
        borderCharacteristicsFresh_ = false;
        if (nic.missing)
            return; // Skip auto-generated missing cells

        var cell = new PdfTableCell(current_, nic.columnIndex,
            nic.nColumnsSpanned, nic.nRowsSpanned);
        currentTableRow_?.Add(cell);
        containerStack_.Push(cell);
    }

    public override void endTableCell()
    {
        if (containerStack_.Count > 0 && containerStack_.Peek() is PdfTableCell)
            containerStack_.Pop();
        end();
    }

    // Table border callbacks: tracked but not deeply modeled in Phase 1
    public override void tableBeforeRowBorder() { }
    public override void tableAfterRowBorder() { }
    public override void tableBeforeColumnBorder() { }
    public override void tableAfterColumnBorder() { }
    public override void tableCellBeforeRowBorder()    => CaptureCellBorder(CellBorderPort.BeforeRow);
    public override void tableCellAfterRowBorder()     => CaptureCellBorder(CellBorderPort.AfterRow);
    public override void tableCellBeforeColumnBorder() => CaptureCellBorder(CellBorderPort.BeforeColumn);
    public override void tableCellAfterColumnBorder()  => CaptureCellBorder(CellBorderPort.AfterColumn);

    private void CaptureCellBorder(CellBorderPort port)
    {
        // Only capture if the DSSSL actually defined this border (setColor/setLineThickness
        // were called by the engine between last capture and this callback).
        if (!borderCharacteristicsFresh_) return;
        borderCharacteristicsFresh_ = false;

        var cell = FindEnclosingCell();
        if (cell == null) return;
        var border = new PdfCellBorder
        {
            LineThickness = lineThickness_,
            ColorR = current_.ColorR,
            ColorG = current_.ColorG,
            ColorB = current_.ColorB
        };
        switch (port)
        {
            case CellBorderPort.BeforeRow:    cell.BeforeRowBorder    = border; break;
            case CellBorderPort.AfterRow:     cell.AfterRowBorder     = border; break;
            case CellBorderPort.BeforeColumn: cell.BeforeColumnBorder = border; break;
            case CellBorderPort.AfterColumn:  cell.AfterColumnBorder  = border; break;
        }

        // Restore pre-border state — the engine's style pop doesn't revert FOT builder state,
        // so the border color would otherwise leak into the cell's content rendering.
        if (cellBaseChars_ != null)
        {
            current_ = cellBaseChars_.Clone();
            lineThickness_ = cellBaseLineThickness_;
        }
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
        borderCharacteristicsFresh_ = true;
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

    // Verbatim/whitespace
    public override void setLines(Symbol lines) => current_.Lines = lines;
    public override void setInputWhitespaceTreatment(Symbol treatment) => current_.InputWhitespaceTreatment = treatment;

    // Cell margins and alignment
    public override void setCellBeforeRowMargin(LengthSpec margin) => current_.CellBeforeRowMargin = margin.length;
    public override void setCellAfterRowMargin(LengthSpec margin) => current_.CellAfterRowMargin = margin.length;
    public override void setCellBeforeColumnMargin(LengthSpec margin) => current_.CellBeforeColumnMargin = margin.length;
    public override void setCellAfterColumnMargin(LengthSpec margin) => current_.CellAfterColumnMargin = margin.length;
    public override void setCellBeforeRowMargin(long margin) => current_.CellBeforeRowMargin = margin;
    public override void setCellAfterRowMargin(long margin) => current_.CellAfterRowMargin = margin;
    public override void setCellBeforeColumnMargin(long margin) => current_.CellBeforeColumnMargin = margin;
    public override void setCellAfterColumnMargin(long margin) => current_.CellAfterColumnMargin = margin;
    public override void setCellBackground(bool background) => current_.CellBackground = background;
    public override void setCellRowAlignment(Symbol alignment) => current_.CellRowAlignment = alignment;

    // Page number format
    public override void setPageNumberFormat(StringC format) => current_.PageNumberFormat = format.ToString();

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
