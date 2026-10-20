namespace Dazzle.Pdf;

using OpenSP;
using OpenJade.Style;
using OpenJade.Grove;
using Char = System.UInt32;
using Symbol = OpenJade.Style.FOTBuilder.Symbol;
using LengthSpec = OpenJade.Style.FOTBuilder.LengthSpec;
using DeviceRGBColor = OpenJade.Style.FOTBuilder.DeviceRGBColor;
using DisplaySpace = OpenJade.Style.FOTBuilder.DisplaySpace;

// Collects the flow object tree; PdfLayout sets it into pages and PdfRenderer writes them.
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

    // Header/footer slot tracking: when set, content goes to HF slot instead of main content,
    // unless it goes into a container opened in the header or footer itself
    private uint? hfSlot_;
    private int hfDepth_;

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

    // Add a node: if inside a container (e.g. paragraph), add there.
    // Otherwise, if in HF mode, add to the HF slot.
    // Otherwise, add to the page sequence.
    private void AddNode(PdfNode node)
    {
        // span applies to what is flowed into a column set; there each area is one column wide
        if (columnSetSequences_ > 0 && current_.Span != 1 && !(node is PdfDisplayGroup g && g.ColumnSetSequence))
            errors_.Add($"span {current_.Span} in a column-set-sequence: an area spans one column only");
        if (hfSlot_.HasValue && pageSequence_ != null && containerStack_.Count <= hfDepth_)
            pageSequence_.HeaderFooter[hfSlot_.Value].Add(node);
        else if (containerStack_.Count > 0)
            containerStack_.Peek().Add(node);
        else
            pageSequence_?.Add(node);
    }

    // A container node becomes the place of what follows until its end
    private void Open(PdfContainerNode node)
    {
        AddNode(node);
        current_.ColumnSetHere = false;
        characteristicsStack_.Push(current_.Clone());
        containerStack_.Push(node);
    }

    private void Close()
    {
        if (containerStack_.Count > 0) containerStack_.Pop();
        end();
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
        hfDepth_ = containerStack_.Count;
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

    // A page-sequence: like a simple-page-sequence without headers and footers, its pages from its models
    public override void startPageSequence()
    {
        characteristicsStack_.Push(current_.Clone());
        pageSequence_ = new PdfPageSequence(current_, nHF)
        {
            InitialPageModels = current_.InitialPageModels,
            RepeatPageModels = current_.RepeatPageModels
        };
        if (containerStack_.Count > 0)
        {
            var stack = containerStack_.ToArray();
            pageSequence_.Add(stack[stack.Length - 1]);
        }
    }

    public override void endPageSequence() => endSimplePageSequence();

    public override void setInitialPageModels(PageModel[] models) => current_.InitialPageModels = models;
    public override void setRepeatPageModels(PageModel[] models) => current_.RepeatPageModels = models;
    public override void setColumnSetModel(ColumnSetModel? model) => current_.ColumnSetModel = model;
    public override void setSpan(long n) => current_.Span = n;

    // What the backend cannot set although the language allows it; DssslProcessor fails the run with these
    private readonly System.Collections.Generic.List<string> errors_ = new();
    public System.Collections.Generic.IReadOnlyList<string> Errors => errors_;
    private int columnSetSequences_;

    // A column-set-sequence: its content set into the columns of its model
    public override void startColumnSetSequence(DisplayNIC nic)
    {
        var group = new PdfDisplayGroup(current_) { Display = Display(nic), ColumnSetSequence = true };
        var model = current_.ColumnSetModel;
        if (model == null)
            errors_.Add("a column-set-sequence without a column-set-model");
        else
        {
            group.ColumnX = model.columnXOrigin.Select(x => x / 1000.0).ToArray();
            group.ColumnWidth = model.columnWidth.Select(w => w / 1000.0).ToArray();
            group.BalanceColumns = model.balance;
        }
        Open(group);
        columnSetSequences_++;
    }

    public override void endColumnSetSequence()
    {
        columnSetSequences_--;
        Close();
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
        using var stream = File.Create(outputFilename_);
        Render(stream);
    }

    public void Finish(Stream stream)
    {
        if (pageSequences_.Count == 0)
            return;
        Render(stream);
    }

    private void Render(Stream stream)
    {
        var renderer = new PdfRenderer();
        renderer.Render(pageSequences_, stream);
        var locations = new System.Collections.Generic.Dictionary<string, PdfElementLocation>();
        foreach (var (name, id) in idNames_)
            if (renderer.Locations.TryGetValue(name, out var location))
                locations[id] = location;
        ElementLocations = locations;
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
            characteristicsStack_.Pop();
            if (characteristicsStack_.Count > 0)
                current_ = characteristicsStack_.Peek().Clone();
        }
    }

    public override void startParagraph(ParagraphNIC nic)
    {
        var para = new PdfParagraph(current_) { Display = Display(nic) };
        Open(para);
    }

    public override void endParagraph() => Close();

    public override void paragraphBreak(ParagraphNIC nic)
    {
        AddNode(new PdfParagraphBreak(current_));
        atomic();
    }

    public override void startDisplayGroup(DisplayGroupNIC nic)
    {
        var group = new PdfDisplayGroup(current_) { Display = Display(nic) };
        if (current_.ColumnSetHere && ParseColumnSet(current_.ColumnSet) is { } columns)
        {
            group.ColumnX = columns.X;
            group.ColumnWidth = columns.Width;
            group.BalanceColumns = current_.BalanceColumns;
        }
        // what the group contains does not set up the columns again (Open forgets that it was here)
        Open(group);
    }

    public override void endDisplayGroup() => Close();

    public override void startBox(BoxNIC nic)
    {
        if (nic.isDisplay)
            Open(new PdfDisplayGroup(current_) { Display = Display(nic) });
        else
            Open(new PdfSequence(current_));
    }

    public override void endBox() => Close();

    public override void startScroll()
    {
        Open(new PdfScroll(current_));
    }

    public override void endScroll() => Close();

    public override void startSequence()
    {
        if (InTableRow)
        {
            // Inside a table row, the DSSSL engine wraps each cell in an implicit
            // sequence. Skip it — only table-cell nodes belong as row children.
            characteristicsStack_.Push(current_.Clone());
            return;
        }
        Open(new PdfSequence(current_));
    }

    public override void endSequence()
    {
        if (containerStack_.Count > 0 && containerStack_.Peek() is not PdfSequence)
        {
            // Matching the skipped startSequence inside a table row
            end();
            return;
        }
        Close();
    }

    public override void startLineField(LineFieldNIC nic)
    {
        Open(new PdfLineField(current_));
    }

    public override void endLineField() => Close();

    // ==================== Inline content ====================

    public override void characters(Char[] data, nuint size)
    {
        bool verbatim = current_.Verbatim;

        var sb = new System.Text.StringBuilder((int)size);
        for (nuint i = 0; i < size; i++)
        {
            uint u = data[i];
            if (verbatim)
            {
                // Preserve whitespace: keep CR/LF as newlines
                if (u == '\r')
                {
                    if (i + 1 < size && data[i + 1] == '\n')
                        i++; // skip CR in CR+LF
                    sb.Append('\n');
                }
                else
                    AppendChar(sb, u);
            }
            else
            {
                // Normal mode: collapse record-ends and runs of whitespace
                if (u == '\r')
                    continue;
                if (u == '\n' || u == ' ' || u == '\t')
                {
                    // Collapse any run of whitespace (spaces, tabs, newlines) into one space
                    sb.Append(' ');
                    while (i + 1 < size)
                    {
                        uint next = data[i + 1];
                        if (next == ' ' || next == '\t' || next == '\n' || next == '\r')
                            i++;
                        else
                            break;
                    }
                }
                else
                    AppendChar(sb, u);
            }
        }
        var text = sb.ToString();
        if (text.Length > 0)
            AddNode(new PdfTextRun(text, current_));
    }

    private static void AppendChar(System.Text.StringBuilder sb, uint u)
    {
        if (u > 0xFFFF)
            sb.Append(char.ConvertFromUtf32((int)u));
        else
            sb.Append((char)u);
    }

    public override void pageNumber()
    {
        AddNode(new PdfPageNumber(current_));
        atomic();
    }

    // Position point shift (superscript/subscript via DSSSL characteristic)
    public override void setPositionPointShift(LengthSpec shift) =>
        current_.PositionPointShift = shift.length;

    public override void startLeader(LeaderNIC nic)
    {
        Open(new PdfLeader(current_));
    }

    public override void endLeader() => Close();

    // Score: the line goes with the text under the characteristics of the score. They are set
    // before start(), so that the score's entry on the stack has them: after a flow object
    // inside the score ends, the characteristics come back from that entry.
    public override void startScore(Symbol type)
    {
        current_.ScoreType = type;
        current_.ScoreLength = null;
        start();
    }

    public override void startScore(LengthSpec length)
    {
        current_.ScoreType = Symbol.symbolFalse;
        current_.ScoreLength = length.length;
        start();
    }

    public override void startScore(Char c)
    {
        // A character score is set as an underline
        current_.ScoreType = Symbol.symbolAfter;
        current_.ScoreLength = null;
        start();
    }

    public override void setScoreSpaces(bool score) => current_.ScoreSpaces = score;

    // ==================== Node tracking & cross-references ====================

    public override void startNode(NodePtr node, StringC processingMode)
    {
        if (processingMode.size() == 0 && node != null)
        {
            var name = GetLocationName(node);
            if (name == null)
                return;
            if (TryGetId(node, out var id))
                idNames_[name] = id;
            // Between the cells of a table (a row element, say) the mark waits for the next cell,
            // so that it lands on the page of its row and not after the whole table.
            if (currentTable_ != null && FindEnclosingCell() == null)
                pendingTableMarks_.Add(name);
            else
                AddNode(new PdfLocationMark(name));
        }
    }

    private readonly System.Collections.Generic.List<string> pendingTableMarks_ = new();
    private readonly System.Collections.Generic.Dictionary<string, string> idNames_ = new();

    /// <summary>
    /// Where the elements with an ID begin, by their ID, after <see cref="Finish(Stream)"/>.
    /// </summary>
    public System.Collections.Generic.IReadOnlyDictionary<string, PdfElementLocation> ElementLocations { get; private set; }
        = new System.Collections.Generic.Dictionary<string, PdfElementLocation>();

    private static bool TryGetId(NodePtr node, out string id)
    {
        GroveString value = new GroveString();
        if (node.getId(ref value) == OpenJade.Grove.AccessResult.accessOK && value.size() > 0)
        {
            var sb = new System.Text.StringBuilder();
            for (nuint i = 0; i < value.size(); i++)
                sb.Append((char)value[i]);
            id = sb.ToString();
            return true;
        }
        id = "";
        return false;
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

    // Border ports: the callback fires with the characteristics (color, line-thickness) of the
    // table-border flow object already set on current_.
    private enum BorderPort { BeforeRow, AfterRow, BeforeColumn, AfterColumn }
    private long lineThickness_ = 1000; // DSSSL units; 1pt default

    // True when setColor or setLineThickness has been called since the last border capture.
    // Used to distinguish "undefined" border ports (no DSSSL characteristic set) from
    // defined ones (where the engine pushes characteristics before the port callback).
    private bool borderCharacteristicsFresh_ = false;
    private bool borderPresent_ = true;

    // Saved state at startTableCell, so we can undo border-port mutations to current_.
    // (openjade-net's StyleStack.pop does not invoke set* to restore FOT builder state.)
    private PdfCharacteristics? cellBaseChars_;
    private long cellBaseLineThickness_;

    public override void setLineThickness(long thickness)
    {
        lineThickness_ = thickness;
        current_.LineThickness = thickness;
        borderCharacteristicsFresh_ = true;
    }

    // line-dash: the first length of the list is the offset, the others alternately dash and gap; without them a line is solid
    public override void setLineDash(System.Collections.Generic.List<long> dashes, long offset)
    {
        current_.LineDash = dashes.Count == 0 ? null : dashes.ToArray();
        current_.LineDashOffset = offset;
        borderCharacteristicsFresh_ = true;
    }

    public override void setBorderPresent(bool present)
    {
        borderPresent_ = present;
        borderCharacteristicsFresh_ = true;
    }

    public override void rule(RuleNIC nic)
    {
        start();
        AddNode(new PdfRule(current_, nic.orientation,
            nic.hasLength, nic.hasLength ? nic.length.length : 0) { Display = Display(nic) });
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
        AddNode(new PdfExternalGraphic(current_,
            nic.entitySystemId.ToString(),
            nic.isDisplay,
            nic.hasMaxWidth, nic.hasMaxWidth ? nic.maxWidth.length : 0,
            nic.hasMaxHeight, nic.hasMaxHeight ? nic.maxHeight.length : 0)
        { Display = Display(nic), ScaleType = nic.scaleType, ScaleX = nic.scale[0], ScaleY = nic.scale[1] });
        end();
    }

    // ==================== Table flow objects ====================

    private PdfTable? currentTable_;
    private enum TablePartSection { Body, Header, Footer }
    private TablePartSection tablePartSection_ = TablePartSection.Body;
    private PdfTableRow? currentTableRow_;

    // Enclosing tables while a table in a cell is built
    private readonly Stack<(PdfTable? table, PdfTableRow? row, TablePartSection section)> tableStack_ = new();

    public override void startTable(TableNIC nic)
    {
        var table = new PdfTable(current_)
        {
            Display = Display(nic),
            HasWidth = nic.widthType == TableNIC.WidthType.widthExplicit,
            Width = nic.width.length
        };
        AddNode(table);
        characteristicsStack_.Push(current_.Clone());
        tableStack_.Push((currentTable_, currentTableRow_, tablePartSection_));
        currentTable_ = table;
        tableBaseChars_ = current_.Clone();
        tableBaseLineThickness_ = lineThickness_;
        borderCharacteristicsFresh_ = false;
    }

    public override void endTable()
    {
        currentTable_ = null;
        if (tableStack_.Count > 0)
            (currentTable_, currentTableRow_, tablePartSection_) = tableStack_.Pop();
        foreach (var name in pendingTableMarks_)
            AddNode(new PdfLocationMark(name));
        pendingTableMarks_.Clear();
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

    public override void startTablePart(TablePartNIC nic,
        out FOTBuilder? header, out FOTBuilder? footer)
    {
        characteristicsStack_.Push(current_.Clone());
        tablePartSection_ = TablePartSection.Body;
        header = new TablePartPortProxy(this, TablePartSection.Header);
        footer = new TablePartPortProxy(this, TablePartSection.Footer);
    }

    public override void endTablePart() { end(); }

    // Forwards all calls to the parent PdfFOTBuilder, but sets tablePartSection_
    // on entry so rows created during port processing go to HeaderRows/FooterRows.
    private class TablePartPortProxy : FOTBuilder
    {
        private readonly PdfFOTBuilder parent_;
        private readonly TablePartSection section_;

        public TablePartPortProxy(PdfFOTBuilder parent, TablePartSection section)
        {
            parent_ = parent;
            section_ = section;
        }

        public PdfFOTBuilder Parent => parent_;

        public override void startTableRow()
        {
            var saved = parent_.tablePartSection_;
            parent_.tablePartSection_ = section_;
            parent_.startTableRow();
            parent_.tablePartSection_ = saved;
        }

        public override void endTableRow() => parent_.endTableRow();
        public override void startTableCell(TableCellNIC nic) => parent_.startTableCell(nic);
        public override void endTableCell() => parent_.endTableCell();
        public override void tableCellBeforeRowBorder()    => parent_.tableCellBeforeRowBorder();
        public override void tableCellAfterRowBorder()     => parent_.tableCellAfterRowBorder();
        public override void tableCellBeforeColumnBorder() => parent_.tableCellBeforeColumnBorder();
        public override void tableCellAfterColumnBorder()  => parent_.tableCellAfterColumnBorder();
        public override void startParagraph(ParagraphNIC nic) => parent_.startParagraph(nic);
        public override void endParagraph() => parent_.endParagraph();
        public override void paragraphBreak(ParagraphNIC nic) => parent_.paragraphBreak(nic);
        public override void startDisplayGroup(DisplayGroupNIC nic) => parent_.startDisplayGroup(nic);
        public override void endDisplayGroup() => parent_.endDisplayGroup();
        public override void startBox(BoxNIC nic) => parent_.startBox(nic);
        public override void endBox() => parent_.endBox();
        public override void startSequence() => parent_.startSequence();
        public override void endSequence() => parent_.endSequence();
        public override void startLineField(LineFieldNIC nic) => parent_.startLineField(nic);
        public override void endLineField() => parent_.endLineField();
        public override void startLeader(LeaderNIC nic) => parent_.startLeader(nic);
        public override void endLeader() => parent_.endLeader();
        public override void startScore(Symbol type) => parent_.startScore(type);
        public override void startScore(LengthSpec length) => parent_.startScore(length);
        public override void startScore(uint c) => parent_.startScore(c);
        public override void endScore() => parent_.endScore();
        public override void characters(uint[] data, nuint size) => parent_.characters(data, size);
        public override void rule(RuleNIC nic) => parent_.rule(nic);
        public override void externalGraphic(ExternalGraphicNIC nic) => parent_.externalGraphic(nic);
        public override void pageNumber() => parent_.pageNumber();
        public override void currentNodePageNumber(NodePtr node) => parent_.currentNodePageNumber(node);
        // the elements of a header or footer row have their place as those of the body
        public override void startNode(NodePtr node, StringC processingMode) => parent_.startNode(node, processingMode);

        // Forward characteristic setters
        public override void setFontSize(long size) => parent_.setFontSize(size);
        public override void setFontFamilyName(StringC name) => parent_.setFontFamilyName(name);
        public override void setFontWeight(Symbol weight) => parent_.setFontWeight(weight);
        public override void setFontPosture(Symbol posture) => parent_.setFontPosture(posture);
        public override void setColor(DeviceRGBColor color) => parent_.setColor(color);
        public override void setBackgroundColor(DeviceRGBColor color) => parent_.setBackgroundColor(color);
        public override void setBackgroundColor() => parent_.setBackgroundColor();
        public override void setStartIndent(LengthSpec i) => parent_.setStartIndent(i);
        public override void setEndIndent(LengthSpec i) => parent_.setEndIndent(i);
        public override void setFirstLineStartIndent(LengthSpec i) => parent_.setFirstLineStartIndent(i);
        public override void setLineSpacing(LengthSpec s) => parent_.setLineSpacing(s);
        public override void setMinLeading(OptLengthSpec l) => parent_.setMinLeading(l);
        public override void setMinPostLineSpacing(OptLengthSpec s) => parent_.setMinPostLineSpacing(s);
        public override void setMinPreLineSpacing(OptLengthSpec s) => parent_.setMinPreLineSpacing(s);
        public override void setQuadding(Symbol q) => parent_.setQuadding(q);
        public override void setDisplayAlignment(Symbol a) => parent_.setDisplayAlignment(a);
        public override void setFieldWidth(LengthSpec w) => parent_.setFieldWidth(w);
        public override void setFieldAlign(Symbol a) => parent_.setFieldAlign(a);
        public override void setPositionPointShift(LengthSpec s) => parent_.setPositionPointShift(s);
        public override void setLineThickness(long t) => parent_.setLineThickness(t);
        public override void setLineDash(System.Collections.Generic.List<long> dashes, long offset) => parent_.setLineDash(dashes, offset);
        public override void setBorderPresent(bool b) => parent_.setBorderPresent(b);
        public override void setCellBeforeRowMargin(LengthSpec m) => parent_.setCellBeforeRowMargin(m);
        public override void setCellAfterRowMargin(LengthSpec m) => parent_.setCellAfterRowMargin(m);
        public override void setCellBeforeColumnMargin(LengthSpec m) => parent_.setCellBeforeColumnMargin(m);
        public override void setCellAfterColumnMargin(LengthSpec m) => parent_.setCellAfterColumnMargin(m);
        public override void setCellBeforeRowMargin(long m) => parent_.setCellBeforeRowMargin(m);
        public override void setCellAfterRowMargin(long m) => parent_.setCellAfterRowMargin(m);
        public override void setCellBeforeColumnMargin(long m) => parent_.setCellBeforeColumnMargin(m);
        public override void setCellAfterColumnMargin(long m) => parent_.setCellAfterColumnMargin(m);
        public override void setCellBackground(bool b) => parent_.setCellBackground(b);
        public override void setCellRowAlignment(Symbol a) => parent_.setCellRowAlignment(a);
        public override void setLines(Symbol l) => parent_.setLines(l);
        public override void setInputWhitespaceTreatment(Symbol t) => parent_.setInputWhitespaceTreatment(t);
        public override void setScoreSpaces(bool b) => parent_.setScoreSpaces(b);

        // Extension characteristics: the setters resolve the proxy to its parent
        public override void extensionSet(Action<FOTBuilder, bool> func, bool arg) => func(parent_, arg);
        public override void extensionSet(Action<FOTBuilder, StringC> func, StringC arg) => func(parent_, arg);
        public override void extensionSet(Action<FOTBuilder, long> func, long arg) => func(parent_, arg);

        public override void start() => parent_.start();
        public override void end() => parent_.end();
        public override void atomic() => parent_.atomic();
    }

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
        foreach (var name in pendingTableMarks_)
            cell.Add(new PdfLocationMark(name));
        pendingTableMarks_.Clear();
    }

    public override void endTableCell()
    {
        if (containerStack_.Count > 0 && containerStack_.Peek() is PdfTableCell)
            containerStack_.Pop();
        end();
    }

    // The borders of the table and of its cells
    private PdfCharacteristics? tableBaseChars_;
    private long tableBaseLineThickness_;

    public override void tableBeforeRowBorder()    => CaptureTableBorder(BorderPort.BeforeRow);
    public override void tableAfterRowBorder()     => CaptureTableBorder(BorderPort.AfterRow);
    public override void tableBeforeColumnBorder() => CaptureTableBorder(BorderPort.BeforeColumn);
    public override void tableAfterColumnBorder()  => CaptureTableBorder(BorderPort.AfterColumn);
    public override void tableCellBeforeRowBorder()    => CaptureCellBorder(BorderPort.BeforeRow);
    public override void tableCellAfterRowBorder()     => CaptureCellBorder(BorderPort.AfterRow);
    public override void tableCellBeforeColumnBorder() => CaptureCellBorder(BorderPort.BeforeColumn);
    public override void tableCellAfterColumnBorder()  => CaptureCellBorder(BorderPort.AfterColumn);

    // The border the port defines, or null when the stylesheet set nothing for it or said
    // border-present?: #f
    private PdfBorder? TakeBorder()
    {
        if (!borderCharacteristicsFresh_) return null;
        borderCharacteristicsFresh_ = false;
        var present = borderPresent_;
        borderPresent_ = true;
        if (!present) return null;
        return new PdfBorder
        {
            LineThickness = lineThickness_,
            LineDash = current_.LineDash,
            LineDashOffset = current_.LineDashOffset,
            ColorR = current_.ColorR,
            ColorG = current_.ColorG,
            ColorB = current_.ColorB
        };
    }

    private void CaptureTableBorder(BorderPort port)
    {
        var border = TakeBorder();
        if (currentTable_ != null && border != null)
        {
            switch (port)
            {
                case BorderPort.BeforeRow:    currentTable_.BeforeRowBorder    = border; break;
                case BorderPort.AfterRow:     currentTable_.AfterRowBorder     = border; break;
                case BorderPort.BeforeColumn: currentTable_.BeforeColumnBorder = border; break;
                case BorderPort.AfterColumn:  currentTable_.AfterColumnBorder  = border; break;
            }
        }
        if (tableBaseChars_ != null)
        {
            current_ = tableBaseChars_.Clone();
            lineThickness_ = tableBaseLineThickness_;
        }
    }

    private void CaptureCellBorder(BorderPort port)
    {
        var border = TakeBorder();
        var cell = FindEnclosingCell();
        if (cell != null && border != null)
        {
            switch (port)
            {
                case BorderPort.BeforeRow:    cell.BeforeRowBorder    = border; break;
                case BorderPort.AfterRow:     cell.AfterRowBorder     = border; break;
                case BorderPort.BeforeColumn: cell.BeforeColumnBorder = border; break;
                case BorderPort.AfterColumn:  cell.AfterColumnBorder  = border; break;
            }
        }

        // Restore pre-border state — the engine's style pop doesn't revert FOT builder state,
        // so the border color would otherwise leak into the cell's content rendering.
        if (cellBaseChars_ != null)
        {
            current_ = cellBaseChars_.Clone();
            lineThickness_ = cellBaseLineThickness_;
        }
    }

    // ==================== Extension flow objects ====================

    public override void extension(ExtensionFlowObj fo, NodePtr currentNode)
    {
    }

    public override void startExtension(ExtensionFlowObj fo, NodePtr currentNode,
        System.Collections.Generic.List<FOTBuilder> fotbs)
    {
    }

    public override void endExtension(ExtensionFlowObj fo)
    {
    }

    // Extension characteristics, for what DSSSL has no characteristic for; declared in a style
    // sheet with
    //   (declare-characteristic min-row-height "UNREGISTERED::Dazzle//Characteristic::min-row-height" 0pt)
    // and then specified like any inherited characteristic.
    public const string CharacteristicPrefix = "UNREGISTERED::Dazzle//Characteristic::";

    // The setter receives the builder the engine is talking to, which may be a port proxy.
    private static PdfFOTBuilder Of(FOTBuilder fotb) =>
        fotb is TablePartPortProxy proxy ? proxy.Parent : (PdfFOTBuilder)fotb;

    public static FOTBuilder.ExtensionTableEntry[] GetExtensions()
    {
        return new[]
        {
            // the height of a table-row: at least, or exactly (0 = from its content)
            new FOTBuilder.ExtensionTableEntry
            {
                pubid = CharacteristicPrefix + "min-row-height",
                lengthSetter = (fotb, n) => Of(fotb).current_.MinRowHeight = n
            },
            new FOTBuilder.ExtensionTableEntry
            {
                pubid = CharacteristicPrefix + "row-height",
                lengthSetter = (fotb, n) => Of(fotb).current_.RowHeight = n
            },
            // the columns a display-group sets its content in: the start of the first column from the
            // start of the area and its width, then for every further column the gap before it and its
            // width, in points ("0 240 14 241"); the content flows down a column, then into the next,
            // and from the last into the first of the next page
            new FOTBuilder.ExtensionTableEntry
            {
                pubid = CharacteristicPrefix + "column-set",
                stringSetter = (fotb, s) =>
                {
                    var b = Of(fotb);
                    b.current_.ColumnSet = s.ToString();
                    b.current_.ColumnSetHere = true;
                }
            },
            // whether the columns of the group's last page are made equally long
            new FOTBuilder.ExtensionTableEntry
            {
                pubid = CharacteristicPrefix + "balance-columns?",
                boolSetter = (fotb, v) => Of(fotb).current_.BalanceColumns = v
            },
            // the margins of the first page of a simple-page-sequence where they differ from the
            // others, as a word processor's taller header on a title page pushes the text down
            new FOTBuilder.ExtensionTableEntry
            {
                pubid = CharacteristicPrefix + "first-page-top-margin",
                lengthSetter = (fotb, n) => Of(fotb).current_.FirstPageTopMargin = n
            },
            new FOTBuilder.ExtensionTableEntry
            {
                pubid = CharacteristicPrefix + "first-page-bottom-margin",
                lengthSetter = (fotb, n) => Of(fotb).current_.FirstPageBottomMargin = n
            },
            // whether an inline external-graphic makes its line taller where it does not fit (#t, the initial value); with #f it
            // leaves the line as its text and line spacing make it and reaches beyond, above or below (with position-point-shift)
            new FOTBuilder.ExtensionTableEntry
            {
                pubid = CharacteristicPrefix + "grows-line?",
                boolSetter = (fotb, v) => Of(fotb).current_.GrowsLine = v
            },
            // space added after every character of the text, also after the last and after a space; negative narrows
            new FOTBuilder.ExtensionTableEntry
            {
                pubid = CharacteristicPrefix + "letter-spacing",
                lengthSetter = (fotb, n) => Of(fotb).current_.LetterSpacing = n
            },
            // the line height and the descent this text needs instead of its font's (0pt: the font's)
            new FOTBuilder.ExtensionTableEntry
            {
                pubid = CharacteristicPrefix + "text-line-height",
                lengthSetter = (fotb, n) => Of(fotb).current_.TextLineHeight = n
            },
            new FOTBuilder.ExtensionTableEntry
            {
                pubid = CharacteristicPrefix + "text-descent",
                lengthSetter = (fotb, n) => Of(fotb).current_.TextDescent = n
            },
            // in thousandths: every line as high as its own text needs times this, the rest below the placement path (0: off)
            new FOTBuilder.ExtensionTableEntry
            {
                pubid = CharacteristicPrefix + "line-height-factor",
                integerSetter = (fotb, n) => Of(fotb).current_.LineHeightFactor = n
            }
        };
    }

    // The columns of a column-set, or null where it does not say at least one column of a width
    private static (double[] X, double[] Width)? ParseColumnSet(string? set)
    {
        if (set == null) return null;
        var parts = set.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length < 2 || parts.Length % 2 != 0) return null;
        var numbers = new double[parts.Length];
        for (int i = 0; i < parts.Length; i++)
            if (!double.TryParse(parts[i], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out numbers[i]))
                return null;
        int n = parts.Length / 2;
        var x = new double[n];
        var w = new double[n];
        double at = numbers[0];
        for (int k = 0; k < n; k++)
        {
            if (k > 0) at += numbers[2 * k];
            x[k] = at;
            w[k] = numbers[2 * k + 1];
            if (w[k] <= 0) return null;
            at += w[k];
        }
        return (x, w);
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

    // Lines
    public override void setLineSpacing(LengthSpec spacing) => current_.LineSpacing = spacing.length;
    public override void setMinLeading(OptLengthSpec leading) =>
        current_.MinLeading = leading.hasLength ? leading.length.length : null;
    public override void setMinPreLineSpacing(OptLengthSpec spacing) =>
        current_.MinPreLineSpacing = spacing.hasLength ? spacing.length.length : null;
    public override void setMinPostLineSpacing(OptLengthSpec spacing) =>
        current_.MinPostLineSpacing = spacing.hasLength ? spacing.length.length : null;

    // Alignment
    public override void setQuadding(Symbol quadding) => current_.Quadding = quadding;
    public override void setDisplayAlignment(Symbol alignment) => current_.DisplayAlignment = alignment;

    // Line fields
    public override void setFieldWidth(LengthSpec width) => current_.FieldWidth = width.length;
    public override void setFieldAlign(Symbol align) => current_.FieldAlign = align;

    // Page dimensions
    public override void setPageWidth(long width) => current_.PageWidth = width;
    public override void setPageHeight(long height) => current_.PageHeight = height;
    public override void setLeftMargin(long margin) => current_.LeftMargin = margin;
    public override void setRightMargin(long margin) => current_.RightMargin = margin;
    public override void setTopMargin(long margin) => current_.TopMargin = margin;
    public override void setBottomMargin(long margin) => current_.BottomMargin = margin;
    public override void setHeaderMargin(long margin) => current_.HeaderMargin = margin;
    public override void setFooterMargin(long margin) => current_.FooterMargin = margin;

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

    private static PdfDisplay Display(FOTBuilder.DisplayNIC nic) => new()
    {
        SpaceBefore = nic.spaceBefore.nominal.length,
        SpaceBeforeConditional = nic.spaceBefore.conditional,
        SpaceAfter = nic.spaceAfter.nominal.length,
        SpaceAfterConditional = nic.spaceAfter.conditional,
        KeepWithNext = nic.keepWithNext,
        KeepWithPrevious = nic.keepWithPrevious,
        BreakBefore = nic.breakBefore == Symbol.symbolPage,
        BreakAfter = nic.breakAfter == Symbol.symbolPage,
        BreakBeforeColumn = nic.breakBefore == Symbol.symbolColumn,
        BreakAfterColumn = nic.breakAfter == Symbol.symbolColumn
    };
}
