namespace Dazzle.Pdf;

using PdfSharp.Drawing;
using Symbol = OpenJade.Style.FOTBuilder.Symbol;
using HF = OpenJade.Style.FOTBuilder.HF;
using PageModel = OpenJade.Style.PageModel;

// The layout: sets the flow object tree into pages. It follows the PDF backend of the Scaly
// port of OpenJade (packages/dazzle/…/PdfFOTBuilder.scaly): blocks flow down the page, a
// paragraph is broken greedily into lines, everything placed is a list of drawing operations
// of the page being set, and a run of them can still be moved (a table row that does not fit,
// a block that must stay with the next one). Inside a table cell, a header and a footer no page
// is broken.
//
// The semantics are those of ISO/IEC 10179 as the RTF backend of Jade has them: line-spacing is
// the distance of the lines, exactly when min-leading is #f and at least when it is a length;
// display space collapses to the larger of the space after and the space before, and
// conditional space is dropped at the start of an area (a page, a cell); cell margins are the
// distance between the edges of a cell and its content.
//
// Lengths are points (double), y downwards from the top of the page.
internal sealed class PdfLayout
{
    // ==================== Results ====================

    public sealed class Page
    {
        public double Width, Height;
        public int Number;              // physical, from 1
        public string NumberFormat = "1";
        public bool FirstOfSequence;
        public List<Op> Ops = new();
    }

    public enum OpKind { Text, Line, Rect, Image, Mark }

    public sealed class Op
    {
        public OpKind Kind;
        public double X, Y, X2, Y2, W, H, Thickness;
        public XColor Color;
        public string? Text;
        public XFont? Font;
        public string? Path;
        public double Spacing;              // letter-spacing of a text: after every character
        public double[]? Dash;              // line-dash of a line: dashes and gaps in points, null solid
        public double DashOffset;
        public bool Background;             // the background of a table cell

        public void Shift(double dx, double dy)
        {
            X += dx; X2 += dx;
            Y += dy; Y2 += dy;
        }
    }

    public List<Page> Pages { get; } = new();
    public Dictionary<string, int> ElementPages { get; } = new();
    // Where every element begins: the page in the document (from 1) and the top left of its first
    // line or cell in points from the top left of the page
    public Dictionary<string, PdfElementLocation> ElementLocations { get; } = new();

    // ==================== State ====================

    private readonly Dictionary<string, int> knownPages_;
    private readonly PdfFontCache fonts_;
    private readonly Func<string, (double Width, double Height)?> imageSize_;

    // The area blocks are set into: the body of the page, a table cell, a header or footer
    private sealed class Area
    {
        public List<Op> Ops = new();
        public double X, Width;
        public double Y, Top, Bottom;
        public bool NoBreak;
        public double PendingConditional, PendingFixed;
        public bool AtTop = true;
        public int KeepAt = -1;
        public double KeepY;
        public Symbol? Align;                 // the alignment of a header or footer part
        public List<string> Marks = new();    // elements waiting for the next line
        // where the cell backgrounds of tables go on this page: under the contents and borders of every table on it, so
        // that the background of a row does not cover the border of the row or the table before
        public List<Op>? BackgroundOps;
        public int BackgroundAt;
    }

    private Area area_ = new();
    private bool pageOpen_;
    private int pageNo_;
    private int fixedOps_;                    // ops in front of which nothing is moved

    // A column set being set (a display-group with column-set): the body of the page is one of its
    // columns at a time. Where a column is full, the next one begins at the top of the set on the
    // same page; after the last, the first on the next page.
    private sealed class ColumnSet
    {
        public int Index;                     // in the order the sets begin in the document
        public double[] X = null!, Width = null!;
        public double OuterX, OuterWidth;     // the area outside the set
        public bool Balance;
        public int Column;
        public double Top;                    // of the set on the page being set
        public double[] Ends = null!;         // where each column ends on that page
        public bool HardBreak;                // a column break on that page
    }
    private ColumnSet? columns_;
    private int columnSets_;

    /// <summary>
    /// Where a column set ended: the page (from 1), the top of the set on it and where each column
    /// ends there; whether it is to be balanced and had a column break on that page.
    /// </summary>
    public sealed record ColumnSetEnd(int Page, double Top, double[] Ends, bool Balance, bool HardBreak);
    public List<ColumnSetEnd> ColumnSetEnds { get; } = new();

    // How long the columns but the last may be on the last page of a set that is balanced, by the
    // index of the set: found by the renderer over several layouts
    private readonly IReadOnlyDictionary<int, (int Page, double Height)> balance_;

    // The page sequence being set, and how many of its pages are written
    private PdfPageSequence? seq_;
    private bool seqFirst_;
    private int seqPage_;

    // The page number a page-number flow object shows (headers and footers ask for their page)
    private int? hfPage_;

    public PdfLayout(Dictionary<string, int> knownPages, PdfFontCache fonts,
        Func<string, (double Width, double Height)?> imageSize,
        IReadOnlyDictionary<int, (int Page, double Height)>? balance = null)
    {
        knownPages_ = knownPages;
        fonts_ = fonts;
        imageSize_ = imageSize;
        balance_ = balance ?? new Dictionary<int, (int, double)>();
    }

    public void Run(List<PdfPageSequence> sequences)
    {
        foreach (var seq in sequences)
        {
            if (seq.Children.Count == 0) continue;
            if (pageOpen_) EmitPage();
            seq_ = seq;
            seqFirst_ = true;
            seqPage_ = 0;
            columns_ = null;
            area_ = new Area
            {
                X = PageX,
                Width = PageWidthOfBody,
                Top = PageTop,
                Bottom = PageBottom
            };
            area_.Y = area_.Top;
            FlowChildren(seq.Children, null, false);
            if (pageOpen_) EmitPage();
        }
    }

    private static double Pt(long dsssl) => dsssl / 1000.0;

    // ==================== Pages ====================

    private void OpenPage()
    {
        pageNo_++;
        pageOpen_ = true;
        area_.Ops = new List<Op>();
        area_.AtTop = true;
        area_.PendingConditional = 0;
        area_.Top = PageTop;
        area_.Bottom = PageBottom;
        area_.Y = area_.Top;
        area_.KeepAt = -1;
        fixedOps_ = 0;
        // a column set stands in the body of the page it is on
        if (columns_ is { } open)
        {
            open.OuterX = PageX;
            open.OuterWidth = PageWidthOfBody;
        }
        else
        {
            area_.X = PageX;
            area_.Width = PageWidthOfBody;
        }
        if (columns_ is { } set)
        {
            set.Top = area_.Top;
            set.Ends = Enumerable.Repeat(set.Top, set.X.Length).ToArray();
            set.HardBreak = false;
            EnterColumn(0);
        }
    }

    // The model of the page being set in a page-sequence: the k-th page takes the k-th of
    // initial-page-models, the pages after them go round repeat-page-models; null in a simple-page-sequence
    private PageModel? Model
    {
        get
        {
            var seq = seq_!;
            if (seq.InitialPageModels == null) return null;
            if (seqPage_ < seq.InitialPageModels.Length) return seq.InitialPageModels[seqPage_];
            var repeat = seq.RepeatPageModels!;
            if (repeat.Length == 0)
                throw new DssslProcessingException(
                    $"page {seqPage_ + 1} of a page-sequence has no page model: initial-page-models has "
                    + $"{seq.InitialPageModels.Length} and repeat-page-models none");
            return repeat[(seqPage_ - seq.InitialPageModels.Length) % repeat.Length];
        }
    }

    // The body of the page: the region of its page model (whose origin is at the lower left, y going up),
    // or the margins of the simple-page-sequence, on its first page first-page-top-margin and
    // first-page-bottom-margin where they are given
    private double PageTop
    {
        get
        {
            if (Model is { } m) return Pt(m.height - m.regionYOrigin - m.regionHeight);
            var c = seq_!.Characteristics;
            return Pt(seqFirst_ && c.FirstPageTopMargin > 0 ? c.FirstPageTopMargin : c.TopMargin);
        }
    }

    private double PageBottom
    {
        get
        {
            if (Model is { } m) return Pt(m.height - m.regionYOrigin);
            var c = seq_!.Characteristics;
            return Pt(c.PageHeight - (seqFirst_ && c.FirstPageBottomMargin > 0 ? c.FirstPageBottomMargin : c.BottomMargin));
        }
    }

    private double PageX
    {
        get
        {
            if (Model is { } m) return Pt(m.regionXOrigin);
            return Pt(seq_!.Characteristics.LeftMargin);
        }
    }

    private double PageWidthOfBody
    {
        get
        {
            if (Model is { } m) return Pt(m.regionWidth);
            var c = seq_!.Characteristics;
            return Pt(c.PageWidth - c.LeftMargin - c.RightMargin);
        }
    }

    // The body of the page becomes column k of the set from its top on this page; on the last page of
    // a set that is balanced, the columns but the last end where the balance says
    private void EnterColumn(int k)
    {
        var set = columns_!;
        var a = area_;
        set.Column = k;
        a.X = set.OuterX + set.X[k];
        a.Width = set.Width[k];
        a.Y = set.Top;
        a.Bottom = PageBottom;
        if (k < set.X.Length - 1 && balance_.TryGetValue(set.Index, out var limit) && limit.Page == Pages.Count + 1)
            a.Bottom = Math.Min(a.Bottom, set.Top + limit.Height);
    }

    // On to the next column of the set on this page; false after the last
    private bool NextColumn()
    {
        var set = columns_;
        if (set == null || set.Column + 1 >= set.X.Length) return false;
        set.Ends[set.Column] = area_.Y;
        EnterColumn(set.Column + 1);
        area_.AtTop = true;
        area_.PendingConditional = 0;
        area_.KeepAt = -1;
        fixedOps_ = area_.Ops.Count;
        return true;
    }

    // A column break: the next column, after the last the next page; outside a column set the next page
    private void ColumnBreak()
    {
        if (area_.NoBreak) return;
        if (columns_ == null)
        {
            if (pageOpen_ && !area_.AtTop) EmitPage();
            return;
        }
        if (!pageOpen_) OpenPage();
        if (NextColumn())
            columns_.HardBreak = true;
        else
        {
            EmitPage();
            OpenPage();
        }
    }

    // A display-group with column-set: its content is set into the columns, the set begins where the
    // group begins and ends below its longest column on the page it ends on
    private void FlowColumnSet(PdfDisplayGroup group)
    {
        var a = area_;
        var set = new ColumnSet
        {
            Index = columnSets_++, X = group.ColumnX!, Width = group.ColumnWidth!, Balance = group.BalanceColumns,
            OuterX = a.X, OuterWidth = a.Width
        };
        columns_ = set;
        // the space before the set (after the block before it) is above all of its columns
        if (pageOpen_)
        {
            a.Y += PendingSpace(a);
            a.PendingConditional = a.PendingFixed = 0;
        }
        set.Top = pageOpen_ ? a.Y : a.Top;
        set.Ends = Enumerable.Repeat(set.Top, set.X.Length).ToArray();
        bool atTop = a.AtTop;
        EnterColumn(0);
        a.AtTop = atTop;
        a.KeepAt = -1;
        fixedOps_ = a.Ops.Count;
        FlowChildren(group.Children, null, group.Display.KeepWithNext);
        set.Ends[set.Column] = a.Y;
        var end = set.Ends.Max();
        ColumnSetEnds.Add(new ColumnSetEnd(pageOpen_ ? Pages.Count + 1 : -1, set.Top, set.Ends.ToArray(), set.Balance, set.HardBreak));
        columns_ = null;
        a.X = set.OuterX;
        a.Width = set.OuterWidth;
        a.Bottom = PageBottom;
        a.Y = end;
        if (end > set.Top + 0.001) a.AtTop = false;
        a.KeepAt = -1;
        fixedOps_ = a.Ops.Count;
    }

    private void EmitPage()
    {
        var c = seq_!.Characteristics;
        var model = Model;
        var page = new Page
        {
            Width = Pt(model?.width ?? c.PageWidth),
            Height = Pt(model?.height ?? c.PageHeight),
            Number = pageNo_,
            NumberFormat = c.PageNumberFormat,
            FirstOfSequence = seqFirst_,
            Ops = area_.Ops
        };
        page.Ops.AddRange(HeadersAndFooters(page));
        foreach (var op in page.Ops)
            if (op.Kind == OpKind.Mark && op.Text != null && !ElementPages.ContainsKey(op.Text))
            {
                ElementPages[op.Text] = pageNo_;
                ElementLocations[op.Text] = new PdfElementLocation(Pages.Count + 1, op.X, op.Y);
            }
        Pages.Add(page);
        seqFirst_ = false;
        seqPage_++;
        pageOpen_ = false;
        area_.Ops = new List<Op>();
        area_.KeepAt = -1;
    }

    // The page (or the column) is full: write it and go on on a new one (in the next column). What
    // must stay with the block being set goes along, unless that is half a page.
    private void NewPage()
    {
        var a = area_;
        List<Op>? moved = null;
        double fromY = a.KeepY, oldY = a.Y, oldX = a.X;
        double top = columns_?.Top ?? a.Top;
        if (a.KeepAt >= fixedOps_ && a.KeepAt < a.Ops.Count && fromY > top
            && (oldY - fromY) * 2 < a.Bottom - top)
        {
            moved = a.Ops.GetRange(a.KeepAt, a.Ops.Count - a.KeepAt);
            a.Ops.RemoveRange(a.KeepAt, a.Ops.Count - a.KeepAt);
        }
        if (!NextColumn())
        {
            EmitPage();
            OpenPage();
        }
        if (moved != null)
        {
            var newTop = a.Y;
            a.KeepAt = a.Ops.Count;
            a.KeepY = newTop;
            foreach (var op in moved) op.Shift(a.X - oldX, newTop - fromY);
            a.Ops.AddRange(moved);
            a.Y = newTop + (oldY - fromY);
            a.AtTop = false;
        }
    }

    private double PendingSpace(Area a) =>
        a.AtTop ? a.PendingFixed : Math.Max(a.PendingConditional, a.PendingFixed);

    // Make room for something `height` high: a page is opened if none is, and broken if it does
    // not fit. Then the pending space is put in, and the elements waiting for a line are marked.
    private void MakeRoom(double height, bool keep)
    {
        var a = area_;
        if (!a.NoBreak && !pageOpen_)
            OpenPage();
        if (!a.NoBreak && !a.AtTop && a.Y + PendingSpace(a) + height > a.Bottom + 0.001)
            NewPage();
        a.Y += PendingSpace(a);
        a.PendingConditional = a.PendingFixed = 0;
        if (keep)
        {
            if (a.KeepAt < 0)
            {
                a.KeepAt = a.Ops.Count;
                a.KeepY = a.Y;
            }
        }
        else
            a.KeepAt = -1;
        foreach (var name in a.Marks)
            a.Ops.Add(new Op { Kind = OpKind.Mark, X = a.X, Y = a.Y, Text = name });
        a.Marks.Clear();
        a.AtTop = false;
    }

    private static void AddPending(Area a, long space, bool conditional)
    {
        var s = Pt(space);
        if (conditional) a.PendingConditional = Math.Max(a.PendingConditional, s);
        else a.PendingFixed = Math.Max(a.PendingFixed, s);
    }

    private void BlockStart(PdfDisplay d)
    {
        if (d.BreakBefore && pageOpen_ && !area_.AtTop && !area_.NoBreak)
            EmitPage();
        else if (d.BreakBeforeColumn)
            ColumnBreak();
        AddPending(area_, d.SpaceBefore, d.SpaceBeforeConditional);
    }

    private void BlockEnd(PdfDisplay d)
    {
        AddPending(area_, d.SpaceAfter, d.SpaceAfterConditional);
        if (d.BreakAfter && pageOpen_ && !area_.NoBreak)
            EmitPage();
        else if (d.BreakAfterColumn)
            ColumnBreak();
    }

    // ==================== The flow ====================

    private enum AtomKind { Text, Space, Newline, Tab, SoftHyphen, Leader, Mark, Image, FieldStart, FieldEnd }

    private sealed class Atom
    {
        public AtomKind Kind;
        public PdfCharacteristics Chars = null!;
        public XFont? Font;
        public string? Text;
        public double Width;
        public double Height;
        public bool Breakable;
        public bool Stretches;            // a space that justification widens
        public bool Holds;                // a zero width space: the line is there
    }

    private readonly List<Atom> atoms_ = new();
    private int paraBase_;
    private PdfCharacteristics? paraChars_;
    private bool paraFirst_ = true;
    private bool paraKeep_;

    // Set a list of nodes: what is inline is collected into lines under the paragraph's
    // characteristics (those of the first text where there is no paragraph), a block ends the
    // lines before it.
    private void FlowChildren(List<PdfNode> children, PdfCharacteristics? p, bool keep)
    {
        var (oldBase, oldChars, oldFirst, oldKeep) = (paraBase_, paraChars_, paraFirst_, paraKeep_);
        paraBase_ = atoms_.Count;
        paraChars_ = p;
        paraFirst_ = true;
        paraKeep_ = keep;
        foreach (var n in children)
            FlowNode(n);
        FlushLines();
        (paraBase_, paraChars_, paraFirst_, paraKeep_) = (oldBase, oldChars, oldFirst, oldKeep);
    }

    private void FlowInside(List<PdfNode> children)
    {
        foreach (var n in children)
            FlowNode(n);
    }

    private void FlowNode(PdfNode n)
    {
        switch (n)
        {
            case PdfTextRun run:
                TextAtoms(run.Text, run.Characteristics);
                break;
            case PdfLineField field:
                atoms_.Add(new Atom { Kind = AtomKind.FieldStart, Chars = field.Characteristics, Width = Pt(field.Characteristics.FieldWidth) });
                FlowInside(field.Children);
                atoms_.Add(new Atom { Kind = AtomKind.FieldEnd, Chars = field.Characteristics });
                break;
            case PdfLeader leader:
                LeaderAtom(leader);
                break;
            case PdfSequence seq:
                FlowInside(seq.Children);
                break;
            case PdfLocationMark mark:
                atoms_.Add(new Atom { Kind = AtomKind.Mark, Chars = mark.Characteristics, Text = mark.LocationName });
                break;
            case PdfPageNumber pn:
            {
                var page = hfPage_ ?? (pageOpen_ ? pageNo_ : pageNo_ + 1);
                TextAtoms(FormatPageNumber(page, seq_?.Characteristics.PageNumberFormat ?? "1"), pn.Characteristics);
                break;
            }
            case PdfNodePageNumber npn:
                TextAtoms(knownPages_.TryGetValue(npn.LocationName, out var p)
                    ? FormatPageNumber(p, FormatOfPage(p)) : "00", npn.Characteristics);
                break;
            case PdfExternalGraphic g when !g.IsDisplay:
                ImageAtom(g);
                break;
            case PdfParagraphBreak:
                FlushLines();
                paraFirst_ = true;
                break;
            case PdfRule r when r.Orientation != Symbol.symbolHorizontal:
                break;
            default:
                FlushLines();
                FlowBlock(n);
                break;
        }
    }

    private void FlowBlock(PdfNode n)
    {
        switch (n)
        {
            case PdfParagraph para:
                BlockStart(para.Display);
                FlowChildren(para.Children, para.Characteristics, para.Display.KeepWithNext);
                BlockEnd(para.Display);
                break;
            case PdfDisplayGroup group:
                BlockStart(group.Display);
                // a column set in a column set, a cell, a header or a footer is an ordinary group; for a
                // column-set-sequence that would set otherwise than it asks
                if (group.ColumnSetSequence && (columns_ != null || area_.NoBreak))
                    throw new DssslProcessingException(
                        "a column-set-sequence in a column set, a table cell, a header or a footer is not supported");
                if (group.ColumnX != null && columns_ == null && !area_.NoBreak)
                    FlowColumnSet(group);
                else
                    FlowChildren(group.Children, null, group.Display.KeepWithNext);
                BlockEnd(group.Display);
                break;
            case PdfScroll scroll:
                FlowChildren(scroll.Children, null, false);
                break;
            case PdfRule rule:
                BlockStart(rule.Display);
                FlowRule(rule);
                BlockEnd(rule.Display);
                break;
            case PdfExternalGraphic graphic:
                BlockStart(graphic.Display);
                FlowGraphic(graphic);
                BlockEnd(graphic.Display);
                break;
            case PdfTable table:
                BlockStart(table.Display);
                FlowTable(table);
                BlockEnd(table.Display);
                break;
            case PdfContainerNode other:
                FlowChildren(other.Children, null, false);
                break;
        }
    }

    // A horizontal rule across the area, or of its length
    private void FlowRule(PdfRule n)
    {
        var c = n.Characteristics;
        var thickness = Pt(c.LineThickness);
        MakeRoom(thickness, n.Display.KeepWithNext);
        var available = area_.Width - Pt(c.StartIndent) - Pt(c.EndIndent);
        var length = n.HasLength ? Pt(n.Length) : available;
        var x = area_.X + Pt(c.StartIndent) + Offset(c.DisplayAlignment, available - length);
        var at = area_.Y + thickness / 2;
        area_.Ops.Add(new Op { Kind = OpKind.Line, X = x, Y = at, X2 = x + length, Y2 = at, Thickness = thickness, Color = ColorOf(c),
            Dash = Dash(c.LineDash), DashOffset = Pt(c.LineDashOffset) });
        area_.Y += thickness;
    }

    private static double Offset(Symbol alignment, double room) =>
        alignment == Symbol.symbolCenter ? room / 2 : alignment == Symbol.symbolEnd ? room : 0;

    private void FlowGraphic(PdfExternalGraphic n)
    {
        var c = n.Characteristics;
        var available = area_.Width - Pt(c.StartIndent) - Pt(c.EndIndent);
        var size = GraphicSize(n, available);
        if (size == null) return;
        var (width, height) = size.Value;
        MakeRoom(height, n.Display.KeepWithNext);
        var x = area_.X + Pt(c.StartIndent) + Offset(c.DisplayAlignment, available - width);
        area_.Ops.Add(new Op { Kind = OpKind.Image, X = x, Y = area_.Y, W = width, H = height, Path = ResolveImagePath(n.SystemId) });
        area_.Y += height;
    }

    // The size a graphic is set at (scale 'max-uniform): uniformly the largest that fits the
    // maximum width and height; without either its own size, at most the width available.
    private (double, double)? GraphicSize(PdfExternalGraphic n, double available)
    {
        var own = imageSize_(ResolveImagePath(n.SystemId));
        if (own == null || own.Value.Width <= 0 || own.Value.Height <= 0) return null;
        var (w, h) = own.Value;
        // scale: factors (OpenJade's symbolFalse) scale the graphic as they are; 'max scales each direction on its own to
        // max-width and max-height
        if (n.ScaleType == Symbol.symbolFalse)
            return (w * n.ScaleX, h * n.ScaleY);
        if (n.ScaleType == Symbol.symbolMax && (n.HasMaxWidth || n.HasMaxHeight))
        {
            var sw = n.HasMaxWidth ? Pt(n.MaxWidth) / w : n.HasMaxHeight ? Pt(n.MaxHeight) / h : 1;
            var sh = n.HasMaxHeight ? Pt(n.MaxHeight) / h : sw;
            return (w * sw, h * sh);
        }
        if (!n.HasMaxWidth && !n.HasMaxHeight)
        {
            if (w > available && available > 0) { h = h * available / w; w = available; }
            return (w, h);
        }
        // scale 'max-uniform (the initial value): the largest uniform scale within max-width and max-height; a max-width the
        // stylesheet gives is the limit, even where the graphic then overflows its line
        var maxW = n.HasMaxWidth ? Pt(n.MaxWidth) : double.MaxValue;
        var maxH = n.HasMaxHeight ? Pt(n.MaxHeight) : double.MaxValue;
        var scale = Math.Min(maxW / w, maxH / h);
        return (w * scale, h * scale);
    }

    public static string ResolveImagePath(string systemId)
    {
        string path = systemId;
        if (path.StartsWith("file://"))
            path = new Uri(path).LocalPath;
        if (!System.IO.Path.IsPathRooted(path))
            path = System.IO.Path.GetFullPath(path);
        return path;
    }

    // ==================== Atoms ====================

    // The atoms of a run of text: words and the spaces between them; in verbatim text lines,
    // the breaks between them, every space and the tabs
    private void TextAtoms(string text, PdfCharacteristics c)
    {
        var font = fonts_.Font(c);
        bool verbatim = c.Verbatim;
        // letter-spacing comes after every character, also after a space
        var ls = Pt(c.LetterSpacing);
        var word = new System.Text.StringBuilder();
        // the text before this run in the paragraph, past location marks: a word may continue across runs
        Atom? before = null;
        for (int k = atoms_.Count - 1; k >= paraBase_; k--)
            if (atoms_[k].Kind != AtomKind.Mark) { before = atoms_[k]; break; }
        bool LetterBefore() => word.Length > 0 ? char.IsLetter(word[^1])
            : before is { Kind: AtomKind.Text, Text: { Length: > 0 } b } && char.IsLetter(b[^1]);
        // a hyphen between letters at the end of the previous run lets the line break before this one
        bool hyphen = before is { Kind: AtomKind.Text, Text: { Length: > 1 } p } && p[^1] == '-' && char.IsLetter(p[^2]);
        void Word()
        {
            if (word.Length == 0) return;
            var s = word.ToString();
            atoms_.Add(new Atom { Kind = AtomKind.Text, Chars = c, Font = font, Text = s, Width = fonts_.Width(s, font) + s.EnumerateRunes().Count() * ls });
            word.Clear();
        }
        foreach (var ch in text)
        {
            switch (ch)
            {
                case '\n' when verbatim:
                    Word();
                    atoms_.Add(new Atom { Kind = AtomKind.Newline, Chars = c });
                    break;
                case '\t' when verbatim:
                    Word();
                    atoms_.Add(new Atom { Kind = AtomKind.Tab, Chars = c, Font = font });
                    break;
                case ' ':
                case '\n':
                case '\t':
                    Word();
                    // outside verbatim text a run of white space is one space, also across runs
                    if (!verbatim && atoms_.Count > paraBase_ && atoms_[^1].Kind == AtomKind.Space && !atoms_[^1].Chars.Verbatim)
                        break;
                    atoms_.Add(new Atom { Kind = AtomKind.Space, Chars = c, Font = font, Width = fonts_.Width(" ", font) + ls, Breakable = !verbatim || c.Wraps, Stretches = true });
                    break;
                case '\u200B':
                    Word();
                    atoms_.Add(new Atom { Kind = AtomKind.Space, Chars = c, Font = font, Width = 0, Breakable = true, Holds = true });
                    break;
                case '\u00AD':
                    Word();
                    atoms_.Add(new Atom { Kind = AtomKind.SoftHyphen, Chars = c, Font = font, Width = fonts_.Width("-", font) + ls });
                    break;
                case '\u200C':
                case '\u200D':
                case '\uFEFF':
                    break;
                case '-' when LetterBefore():
                    // after a hyphen between letters the line may break, as in a compound word
                    word.Append(ch);
                    hyphen = true;
                    break;
                default:
                    if (hyphen && char.IsLetter(ch))
                    {
                        Word();
                        atoms_.Add(new Atom { Kind = AtomKind.Space, Chars = c, Font = font, Width = 0, Breakable = c.Wraps });
                    }
                    hyphen = false;
                    word.Append(ch);
                    break;
            }
            if (ch != '-') hyphen = false;
        }
        Word();
    }

    // A leader: its pattern is the text in it, a dot where there is none
    private void LeaderAtom(PdfLeader n)
    {
        var c = n.Characteristics;
        string pattern = ".";
        foreach (var child in n.Children)
            if (child is PdfTextRun t && t.Text.Length > 0)
            {
                c = t.Characteristics;
                pattern = t.Text;
                break;
            }
        var font = fonts_.Font(c);
        atoms_.Add(new Atom { Kind = AtomKind.Leader, Chars = c, Font = font, Text = pattern, Width = fonts_.Width(pattern, font) });
    }

    private void ImageAtom(PdfExternalGraphic g)
    {
        var size = GraphicSize(g, area_.Width);
        if (size == null) return;
        atoms_.Add(new Atom { Kind = AtomKind.Image, Chars = g.Characteristics, Text = ResolveImagePath(g.SystemId), Width = size.Value.Item1, Height = size.Value.Item2 });
    }

    // ==================== Lines ====================

    // The default tab stops of a tab in verbatim text, as in RTF and Word
    private const double DefaultTab = 36;

    // The width of the atoms from..to set side by side from `start`; a leader counts as one
    // pattern, a line field as at least its width, a soft hyphen only at the end of the line.
    private double LineWidth(int from, int to, double start = 0)
    {
        double x = start, fieldX = 0, fieldWidth = 0;
        for (int i = from; i < to; i++)
        {
            var a = atoms_[i];
            switch (a.Kind)
            {
                case AtomKind.FieldStart: fieldX = x; fieldWidth = a.Width; break;
                case AtomKind.FieldEnd: if (x < fieldX + fieldWidth) x = fieldX + fieldWidth; break;
                case AtomKind.Tab: x = (Math.Floor(x / DefaultTab + 1e-6) + 1) * DefaultTab; break;
                case AtomKind.SoftHyphen: if (i == to - 1) x += a.Width; break;
                case AtomKind.Mark:
                case AtomKind.Newline: break;
                default: x += a.Width; break;
            }
        }
        return x - start;
    }

    // The atoms collected since the paragraph began become lines
    private void FlushLines()
    {
        int b = paraBase_, end = atoms_.Count;
        if (end == b) return;
        var pc = paraChars_;
        bool visible = false;
        for (int i = b; i < end; i++)
        {
            var k = atoms_[i].Kind;
            if (k is AtomKind.Text or AtomKind.Image or AtomKind.Leader or AtomKind.Newline or AtomKind.Tab
                || k == AtomKind.Space && (atoms_[i].Chars.Verbatim || atoms_[i].Holds))
            {
                pc ??= atoms_[i].Chars;
                visible = true;
                break;
            }
        }
        if (!visible || pc == null)
        {
            for (int i = b; i < end; i++)
                if (atoms_[i].Kind == AtomKind.Mark)
                    area_.Marks.Add(atoms_[i].Text!);
            atoms_.RemoveRange(b, end - b);
            return;
        }
        var p = pc;
        bool wraps = p.Wraps;
        int from = b;
        bool wrapped = false;
        int hyphenLines = 0;          // lines in a row ending at a soft hyphen, for hyphenation-ladder-count
        while (from < end)
        {
            var indent = Pt(p.StartIndent) + (paraFirst_ ? Pt(p.FirstLineStartIndent) : 0);
            var available = area_.Width - indent - Pt(p.EndIndent);
            // a line does not begin with a space, outside verbatim text; after a line the text was broken at for its width, the
            // spaces at the break stay at the end of that line also in verbatim text (as Word has them, and as XSL's
            // suppress-at-line-break)
            while (from < end && atoms_[from].Kind == AtomKind.Space && !atoms_[from].Holds
                   && (!atoms_[from].Chars.Verbatim || wrapped && atoms_[from].Breakable))
                from++;
            if (from >= end) break;
            int to = from, next = -1;
            int breakAt = -1, breakNext = -1, spaceAt = -1;
            bool ended = false;
            // a line breaks only after it holds something: a location mark at its start does not count, so a first word
            // wider than the line overflows it instead of leaving a line with only the mark
            bool held = false;
            while (to < end)
            {
                var a = atoms_[to];
                if (a.Kind == AtomKind.Newline)
                {
                    ended = true;
                    next = to + 1;
                    break;
                }
                // a space that runs past the edge does not break the line: it stays at the end of the line as Word leaves it
                // hanging there, and the line breaks there when the next text does not fit
                if (wraps && held && a.Kind != AtomKind.Mark && !(a.Kind == AtomKind.Space && a.Breakable)
                    && LineWidth(from, to + 1, indent) > available + Squeeze(p, from, to, spaceAt) + 0.001)
                {
                    if (breakAt > from)
                    {
                        to = breakAt;
                        next = breakNext;
                    }
                    else if (a.Kind == AtomKind.Space && a.Breakable)
                        next = to + 1;
                    else
                        next = to;
                    break;
                }
                if (a.Kind == AtomKind.Space && a.Breakable)
                {
                    breakAt = spaceAt = to;
                    breakNext = to + 1;
                }
                else if (a.Kind == AtomKind.SoftHyphen && LineWidth(from, to + 1, indent) <= available + HyphenSqueeze(p, from, to, breakAt) + 0.001
                         && HyphenBreaks(p, from, to, spaceAt, indent, available, hyphenLines))
                {
                    breakAt = to + 1;     // the hyphen stays on the line
                    breakNext = to + 1;
                }
                if (a.Kind != AtomKind.Mark) held = true;
                to++;
            }
            if (next < 0) next = to;
            bool last = to >= end || ended;
            wrapped = !ended && to < end;
            hyphenLines = !last && to > from && atoms_[to - 1].Kind == AtomKind.SoftHyphen ? hyphenLines + 1 : 0;
            SetLine(p, from, to, indent, available, last);
            paraFirst_ = false;
            from = next;
            if (ended && from >= end)
            {
                // a record end at the very end of the text ends the line, it makes no new one
                break;
            }
        }
        atoms_.RemoveRange(b, end - b);
    }

    // Set one line of a paragraph at the place of the next block
    private void SetLine(PdfCharacteristics p, int from, int to, double indent, double available, bool last)
    {
        double natural = 0, descent = 0, image = 0, imageBelow = 0;
        int spaces = 0, leaders = 0;
        bool text = false;
        for (int i = from; i < to; i++)
        {
            var a = atoms_[i];
            if (a.Kind is AtomKind.Text or AtomKind.Leader or AtomKind.SoftHyphen or AtomKind.Space)
            {
                if (a.Font == null) continue;
                text = true;
                natural = Math.Max(natural, TextLineHeight(a.Chars, a.Font));
                descent = Math.Max(descent, TextDescent(a.Chars, a.Font) - Pt(a.Chars.PositionPointShift));
            }
            if (a.Kind == AtomKind.Image && a.Chars.GrowsLine)
            {
                // position-point-shift moves an inline graphic like text: what it lowers below the placement path is descent
                var shift = Pt(a.Chars.PositionPointShift);
                image = Math.Max(image, a.Height + shift);
                imageBelow = Math.Max(imageBelow, -shift);
            }
            if (a.Kind == AtomKind.Space && a.Stretches && i > from) spaces++;
            if (a.Kind == AtomKind.Leader) leaders++;
        }
        if (!text)
        {
            var font = fonts_.Font(p);
            natural = TextLineHeight(p, font);
            descent = TextDescent(p, font);
        }
        descent = Math.Max(descent, imageBelow);
        // the line spacing: exactly, or at least with min-leading
        var height = Pt(p.LineSpacing);
        if (p.MinLeading.HasValue)
            height = Math.Max(height, natural + Pt(p.MinLeading.Value));
        if (image > 0)
            height = Math.Max(height, image + descent);
        // Below the placement path: the descent, or min-post-line-spacing where that is more; with min-pre-line-spacing alone,
        // the rest of the line below the space it asks above. A line area is min-pre-line-spacing before and
        // min-post-line-spacing after its placement path (12.6.6.1), the line spacing beyond that is space between lines; at the
        // end of an area the line needs only its own extent down to its descent, and the rest of min-post-line-spacing may lie
        // past the end (as word processors do with the extra of a multiple line spacing)
        var below = p.MinPostLineSpacing.HasValue ? Math.Max(descent, Pt(p.MinPostLineSpacing.Value))
            : p.MinPreLineSpacing.HasValue ? Math.Max(descent, height - Pt(p.MinPreLineSpacing.Value)) : descent;
        var past = p.MinPostLineSpacing.HasValue ? below - descent : 0;
        // line-height-factor: the line as high as its own text needs times the factor, the rest below the placement path, which
        // may lie past the end of the area like the rest of min-post-line-spacing
        if (p.LineHeightFactor > 0)
        {
            var factor = p.LineHeightFactor / 1000.0;
            height = Math.Max(natural * factor, image + descent);
            below = descent + Math.Max(0, height - natural);
            past = below - descent;
        }
        MakeRoom(height - past, paraKeep_);
        var top = area_.Y;
        var baseline = top + height - below;
        var width = LineWidth(from, to, indent);
        double x = area_.X + indent, extra = 0, fill = 0;
        var quadding = p.Quadding;
        if (area_.Align.HasValue && quadding == Symbol.symbolStart)
            quadding = area_.Align.Value;
        if (leaders > 0)
        {
            if (available > width) fill = (available - width) / leaders;
        }
        else if (quadding == Symbol.symbolEnd)
            x += available - width;
        else if (quadding == Symbol.symbolCenter)
            x += (available - width) / 2;
        else if (quadding == Symbol.symbolJustify && !last && spaces > 0 && available > width)
            extra = (available - width) / spaces;
        else if (quadding == Symbol.symbolJustify && spaces > 0 && width > available + 0.001 && p.JustifySqueeze > 0)
            extra = (available - width) / spaces;      // justify-squeeze: the spaces narrowed so that the line fits
        if (p.HasBackgroundColor)
            area_.Ops.Add(new Op { Kind = OpKind.Rect, X = area_.X + Pt(p.StartIndent), Y = top,
                W = area_.Width - Pt(p.StartIndent) - Pt(p.EndIndent), H = height,
                Color = XColor.FromArgb(p.BackgroundR, p.BackgroundG, p.BackgroundB) });
        SetAtoms(from, to, x, baseline, top, extra, fill, area_.X);
        area_.Y = top + height;
    }

    // justify-squeeze: how far a justified line from..to may run past the edge, its spaces narrowed: the given thousandths of
    // the width of each space between its words, and with justify-squeeze-word at most so many thousandths of the word that
    // does not fit, from the last space the line could break at (with it) to the atom at to, also where a soft hyphen breaks it
    private double Squeeze(PdfCharacteristics p, int from, int to, int spaceAt)
    {
        if (p.JustifySqueeze <= 0 || p.Quadding != Symbol.symbolJustify || spaceAt <= from) return 0;
        double sum = 0;
        for (int i = from + 1; i < to; i++)
            if (atoms_[i].Kind == AtomKind.Space && atoms_[i].Stretches)
                sum += atoms_[i].Width;
        var squeeze = sum * p.JustifySqueeze / 1000.0;
        if (p.JustifySqueezeWord > 0)
            squeeze = Math.Min(squeeze, LineWidth(spaceAt, to + 1) * p.JustifySqueezeWord / 1000.0);
        return squeeze;
    }

    // justify-squeeze at the soft hyphen at to: the share of the spaces, with justify-squeeze-hyphen at most so many thousandths
    // of the syllable before the hyphen, from the last place the line could break (a space with it), the hyphen not counted
    private double HyphenSqueeze(PdfCharacteristics p, int from, int to, int breakAt)
    {
        if (p.JustifySqueeze <= 0 || p.Quadding != Symbol.symbolJustify) return 0;
        double sum = 0;
        for (int i = from + 1; i < to; i++)
            if (atoms_[i].Kind == AtomKind.Space && atoms_[i].Stretches)
                sum += atoms_[i].Width;
        var squeeze = sum * p.JustifySqueeze / 1000.0;
        // the bound from the syllable before the hyphen holds from the second syllable on; the first is only bound by the spaces
        if (p.JustifySqueezeHyphen > 0 && breakAt > from && atoms_[breakAt - 1].Kind == AtomKind.SoftHyphen)
            squeeze = Math.Min(squeeze, LineWidth(breakAt, to) * p.JustifySqueezeHyphen / 1000.0);
        return squeeze;
    }

    // Whether the line may break at the soft hyphen at to, the word beginning after the space at spaceAt (or at the line's start)
    private bool HyphenBreaks(PdfCharacteristics p, int from, int to, int spaceAt, double indent, double available, int hyphenLines)
    {
        // hyphenation-ladder-count: no more lines in a row at a hyphen
        if (p.HyphenationLadderCount > 0 && hyphenLines >= p.HyphenationLadderCount) return false;
        // hyphenation-next-char?: the character after the hyphen fits on the line as well
        if (p.HyphenationNextChar && to + 1 < atoms_.Count && atoms_[to + 1] is { Kind: AtomKind.Text, Text.Length: > 0 } after
            && LineWidth(from, to, indent) + fonts_.Width(after.Text![..1], after.Font!) + Pt(after.Chars.LetterSpacing) > available + 0.001)
            return false;
        if (spaceAt <= from) return true;
        if (p.Quadding == Symbol.symbolJustify && p.HyphenationStretch > 0)
        {
            // hyphenation-stretch: the line without the word, justified, widens each space by more than the share of a space
            int spaces = 0;
            for (int i = from + 1; i < spaceAt; i++)
                if (atoms_[i].Kind == AtomKind.Space && atoms_[i].Stretches) spaces++;
            var room = available - LineWidth(from, spaceAt, indent);
            return spaces == 0 || room / spaces > atoms_[spaceAt].Width * p.HyphenationStretch / 1000.0;
        }
        // hyphenation-zone: more room than that left before the word, from its start or the end of the word before, with
        // hyphenation-zone-indent? and the start indent of the line
        if (p.HyphenationZone > 0)
            return available - LineWidth(from, p.HyphenationZoneSpace ? spaceAt : spaceAt + 1, indent)
                   + (p.HyphenationZoneIndent ? indent : 0) > Pt(p.HyphenationZone) + 0.001;
        return true;
    }

    // The line height and the descent of a text: text-line-height and text-descent where they are set, else its font's
    private double TextLineHeight(PdfCharacteristics c, XFont font) => c.TextLineHeight > 0 ? Pt(c.TextLineHeight) : fonts_.LineHeight(font);

    private double TextDescent(PdfCharacteristics c, XFont font) => c.TextDescent > 0 ? Pt(c.TextDescent) : fonts_.Descent(font);

    // Put the atoms from..to side by side with their baseline at `baseline`. Every space gets
    // `extra` more, every leader `fill` more than one pattern.
    private void SetAtoms(int from, int to, double start, double baseline, double top, double extra, double fill, double lineStart)
    {
        double x = start, fieldX = 0, fieldWidth = 0;
        // the score being drawn: where it begins, where its last scored text ends
        PdfCharacteristics? score = null;
        XFont? scoreFont = null;
        double scoreFrom = 0, scoreTo = 0;
        void EndScore()
        {
            if (score != null && scoreTo > scoreFrom)
                DrawScore(score, scoreFont!, scoreFrom, scoreTo, baseline);
            score = null;
        }
        for (int i = from; i < to; i++)
        {
            var a = atoms_[i];
            switch (a.Kind)
            {
                case AtomKind.Text:
                {
                    var c = a.Chars;
                    if (c.Scored)
                    {
                        if (score != null && !SameScore(score, c)) EndScore();
                        if (score == null) { score = c; scoreFont = a.Font; scoreFrom = x; }
                        scoreTo = x + a.Width;
                    }
                    else
                        EndScore();
                    area_.Ops.Add(new Op { Kind = OpKind.Text, X = x, Y = baseline - Pt(c.PositionPointShift), Text = a.Text, Font = a.Font, Color = ColorOf(c),
                        Spacing = Pt(c.LetterSpacing) });
                    x += a.Width;
                    break;
                }
                case AtomKind.Space:
                    if (score != null && !(a.Chars.Scored && a.Chars.ScoreSpaces)) EndScore();
                    x += a.Width + (a.Stretches && i > from ? extra : 0);
                    break;
                case AtomKind.Tab:
                    EndScore();
                    x = lineStart + (Math.Floor((x - lineStart) / DefaultTab + 1e-6) + 1) * DefaultTab;
                    break;
                case AtomKind.SoftHyphen:
                    if (i == to - 1)
                    {
                        area_.Ops.Add(new Op { Kind = OpKind.Text, X = x, Y = baseline - Pt(a.Chars.PositionPointShift), Text = "-", Font = a.Font, Color = ColorOf(a.Chars) });
                        if (score != null) scoreTo = x + a.Width;
                        x += a.Width;
                    }
                    break;
                case AtomKind.Leader:
                {
                    // as many patterns as fit, flush with the end of the room
                    EndScore();
                    var room = a.Width + fill;
                    var count = a.Width > 0 ? (int)Math.Floor(room / a.Width + 1e-9) : 0;
                    if (count > 0)
                        area_.Ops.Add(new Op { Kind = OpKind.Text, X = x + room - count * a.Width, Y = baseline,
                            Text = string.Concat(Enumerable.Repeat(a.Text, count)), Font = a.Font, Color = ColorOf(a.Chars) });
                    x += room;
                    break;
                }
                case AtomKind.Image:
                    EndScore();
                    area_.Ops.Add(new Op { Kind = OpKind.Image, X = x, Y = baseline - a.Height - Pt(a.Chars.PositionPointShift), W = a.Width, H = a.Height, Path = a.Text });
                    x += a.Width;
                    break;
                case AtomKind.Mark:
                    area_.Ops.Add(new Op { Kind = OpKind.Mark, X = x, Y = top, Text = a.Text });
                    break;
                case AtomKind.FieldStart:
                {
                    fieldX = x;
                    fieldWidth = a.Width;
                    // a line field in a score is an area of the score's content: it is scored over its whole width
                    if (a.Chars.Scored)
                    {
                        if (score != null && !SameScore(score, a.Chars)) EndScore();
                        if (score == null) { score = a.Chars; scoreFont = fonts_.Font(a.Chars); scoreFrom = x; scoreTo = x; }
                        scoreTo = Math.Max(scoreTo, x + fieldWidth);
                    }
                    // the content of the field: at its start, end or middle
                    int depth = 0, j = i + 1;
                    for (; j < to; j++)
                    {
                        if (atoms_[j].Kind == AtomKind.FieldStart) depth++;
                        else if (atoms_[j].Kind == AtomKind.FieldEnd && depth-- == 0) break;
                    }
                    var content = LineWidth(i + 1, j);
                    if (content < fieldWidth)
                        x += Offset(a.Chars.FieldAlign, fieldWidth - content);
                    break;
                }
                case AtomKind.FieldEnd:
                    if (x < fieldX + fieldWidth) x = fieldX + fieldWidth;
                    break;
            }
        }
        EndScore();
    }

    private static bool SameScore(PdfCharacteristics a, PdfCharacteristics b) =>
        a.ScoreType == b.ScoreType && a.ScoreLength == b.ScoreLength
        && a.ColorR == b.ColorR && a.ColorG == b.ColorG && a.ColorB == b.ColorB
        && (a.LineDash ?? []).SequenceEqual(b.LineDash ?? []) && a.LineDashOffset == b.LineDashOffset;

    // line-dash in points; null for a solid line
    private static double[]? Dash(long[]? dash) => dash == null ? null : Array.ConvertAll(dash, d => Pt(d));

    // A score: 'after under the text, 'through through it, 'before over it, a length below the
    // placement path by that much
    private void DrawScore(PdfCharacteristics c, XFont font, double x1, double x2, double baseline)
    {
        var (underline, thickness, strike, over) = fonts_.ScoreMetrics(font);
        double y = c.ScoreLength.HasValue ? baseline + Pt(c.ScoreLength.Value)
            : c.ScoreType == Symbol.symbolThrough ? baseline - strike
            : c.ScoreType == Symbol.symbolBefore ? baseline - over
            : baseline + underline;
        area_.Ops.Add(new Op { Kind = OpKind.Line, X = x1, Y = y, X2 = x2, Y2 = y, Thickness = thickness, Color = ColorOf(c),
            Dash = Dash(c.LineDash), DashOffset = Pt(c.LineDashOffset) });
    }

    private static XColor ColorOf(PdfCharacteristics c) => XColor.FromArgb(c.ColorR, c.ColorG, c.ColorB);

    // ==================== Headers and footers ====================


    private IEnumerable<Op> HeadersAndFooters(Page page)
    {
        var seq = seq_!;
        var c = seq.Characteristics;
        int flags = (page.FirstOfSequence ? (int)HF.firstHF : (int)HF.otherHF)
                  | (page.Number % 2 == 1 ? (int)HF.frontHF : (int)HF.backHF);
        var result = new List<Op>();
        foreach (var header in new[] { true, false })
            foreach (var (part, align) in new[] { ((int)HF.leftHF, Symbol.symbolStart), ((int)HF.centerHF, Symbol.symbolCenter), ((int)HF.rightHF, Symbol.symbolEnd) })
            {
                int slot = flags | (header ? (int)HF.headerHF : (int)HF.footerHF) | part;
                if (slot >= seq.HeaderFooter.Length || seq.HeaderFooter[slot].Count == 0) continue;
                var (ops, height) = SetPart(seq.HeaderFooter[slot], Pt(c.LeftMargin),
                    Pt(c.PageWidth - c.LeftMargin - c.RightMargin), align, page.Number);
                // header-margin and footer-margin: the distance of the header from the top of the
                // page and of the footer from its bottom (as \headery and \footery in RTF);
                // without them the part stands in the middle of the margin, as in the Scaly engine
                double dy = header
                    ? (c.HeaderMargin > 0 ? Pt(c.HeaderMargin) : Math.Max(0, (Pt(c.TopMargin) - height) / 2))
                    : Pt(c.PageHeight) - (c.FooterMargin > 0 ? Pt(c.FooterMargin) + height : Math.Max(height, (Pt(c.BottomMargin) + height) / 2));
                foreach (var op in ops)
                {
                    if (op.Kind == OpKind.Mark) continue;
                    op.Shift(0, dy);
                    result.Add(op);
                }
            }
        return result;
    }

    private (List<Op> Ops, double Height) SetPart(List<PdfNode> nodes, double x, double width, Symbol align, int page)
    {
        var saved = area_;
        var savedHf = hfPage_;
        area_ = new Area { X = x, Width = width, NoBreak = true, Bottom = double.MaxValue, Align = align };
        hfPage_ = page;
        FlowChildren(nodes, null, false);
        var result = (area_.Ops, area_.Y);
        area_ = saved;
        hfPage_ = savedHf;
        return result;
    }

    // ==================== Tables ====================

    private sealed class CellBox
    {
        public PdfTableCell Cell = null!;
        public int Row;            // index in the group
        public int RowSpan;
        public double X, Width;    // relative to the table
        public List<Op> Ops = null!;
        public double Height;      // margins and content
    }

    // A table: the widths of its columns, then its rows one below the other in groups the row
    // spans hold together, the header rows first and again at the top of every further page.
    private void FlowTable(PdfTable table)
    {
        var c = table.Characteristics;
        var available = area_.Width - Pt(c.StartIndent) - Pt(c.EndIndent);
        int columns = table.Columns.Count == 0 ? 0 : table.Columns.Max(col => (int)(col.ColumnIndex + col.NColumnsSpanned));
        foreach (var row in table.HeaderRows.Concat(table.BodyRows).Concat(table.FooterRows))
        {
            int inRow = 0;
            foreach (var cell in row.Children.OfType<PdfTableCell>())
                inRow = Math.Max(inRow, (int)(cell.ColumnIndex + cell.NColumnsSpanned));
            columns = Math.Max(columns, inRow);
        }
        if (columns == 0) return;
        var tableWidth = table.HasWidth ? Pt(table.Width) : available;

        // The widths: a length is taken as it is, what is left over is shared by the table
        // units; a column that says nothing has one unit
        var fixedW = new double[columns];
        var units = new double[columns];
        for (int k = 0; k < columns; k++) units[k] = 1;
        foreach (var col in table.Columns)
        {
            if (col.NColumnsSpanned != 1 || col.ColumnIndex >= columns) continue;
            if (col.HasWidth)
            {
                fixedW[col.ColumnIndex] = Pt(col.Width);
                units[col.ColumnIndex] = col.TableUnitFactor;
            }
        }
        var left = Math.Max(0, tableWidth - fixedW.Sum());
        var totalUnits = units.Sum();
        var widths = new double[columns];
        for (int k = 0; k < columns; k++)
            widths[k] = fixedW[k] + (totalUnits > 0 ? left * units[k] / totalUnits : 0);
        var total = widths.Sum();
        // from the start of the area: a row that goes on in the next column of a column set moves with it
        var x = Pt(c.StartIndent) + (total < available ? Offset(c.DisplayAlignment, available - total) : 0);
        var starts = new double[columns + 1];
        for (int k = 0; k < columns; k++) starts[k + 1] = starts[k] + widths[k];

        bool first = true;
        var headerGroups = Groups(table.HeaderRows);
        foreach (var (rows, body) in headerGroups.Select(g => (g, false))
            .Concat(Groups(table.BodyRows).Select(g => (g, true)))
            .Concat(Groups(table.FooterRows).Select(g => (g, false))))
        {
            PlaceGroup(table, rows, x, starts, total, body, headerGroups, first);
            first = false;
        }
        if (!first && table.AfterRowBorder != null)
            area_.Ops.Add(BorderLine(area_.X + x, area_.Y, area_.X + x + total, area_.Y, table.AfterRowBorder));
    }

    // Rows held together by cells that span them
    private static List<List<PdfTableRow>> Groups(List<PdfTableRow> rows)
    {
        var groups = new List<List<PdfTableRow>>();
        int reach = -1;
        for (int r = 0; r < rows.Count; r++)
        {
            if (r > reach) groups.Add(new List<PdfTableRow>());
            groups[^1].Add(rows[r]);
            foreach (var cell in rows[r].Children.OfType<PdfTableCell>())
                reach = Math.Max(reach, r + (int)Math.Max(1, cell.NRowsSpanned) - 1);
            reach = Math.Max(reach, r);
        }
        return groups;
    }

    private void PlaceGroup(PdfTable table, List<PdfTableRow> rows, double x, double[] starts, double total,
        bool body, List<List<PdfTableRow>> headerGroups, bool first)
    {
        var (ops, height, rowTops) = SetGroup(rows, starts, total);
        var a = area_;
        if (!a.NoBreak && pageOpen_ && !a.AtTop && a.Y + PendingSpace(a) + height > a.Bottom + 0.001)
        {
            NewPage();
            if (body)
                foreach (var head in headerGroups)
                    PlaceGroup(table, head, x, starts, total, false, headerGroups, false);
        }
        MakeRoom(height, false);
        var top = a.Y;
        if (first && table.BeforeRowBorder != null)
            ops.Add(BorderLine(0, 0, total, 0, table.BeforeRowBorder));
        if (table.BeforeColumnBorder != null)
            ops.Add(BorderLine(0, 0, 0, height, table.BeforeColumnBorder));
        if (table.AfterColumnBorder != null)
            ops.Add(BorderLine(total, 0, total, height, table.AfterColumnBorder));
        if (a.BackgroundOps != a.Ops)
        {
            a.BackgroundOps = a.Ops;
            a.BackgroundAt = a.Ops.Count;
        }
        foreach (var op in ops)
        {
            op.Shift(a.X + x, top);
            if (op.Background)
                a.Ops.Insert(a.BackgroundAt++, op);
            else
                a.Ops.Add(op);
        }
        a.Y = top + height;
    }

    // The rows of a group: every cell set into its columns as an area of its own, the rows as
    // high as their highest cell (a spanning cell makes its last row higher where it needs to),
    // then backgrounds, contents at their row alignment, and borders. Relative to the group.
    private (List<Op> Ops, double Height, double[] RowTops) SetGroup(List<PdfTableRow> rows, double[] starts, double total)
    {
        int columns = starts.Length - 1;
        var boxes = new List<CellBox>();
        for (int r = 0; r < rows.Count; r++)
        {
            int column = 0;
            foreach (var cell in rows[r].Children.OfType<PdfTableCell>())
            {
                if (cell.ColumnIndex < columns) column = (int)cell.ColumnIndex;
                if (column >= columns) continue;
                int span = (int)Math.Max(1, cell.NColumnsSpanned);
                int end = Math.Min(columns, column + span);
                var cc = cell.Characteristics;
                var box = new CellBox
                {
                    Cell = cell, Row = r, RowSpan = Math.Min((int)Math.Max(1, cell.NRowsSpanned), rows.Count - r),
                    X = starts[column], Width = starts[end] - starts[column]
                };
                var inner = box.Width - Pt(cc.CellBeforeColumnMargin) - Pt(cc.CellAfterColumnMargin);
                var (ops, contentHeight) = SetCell(cell, inner);
                box.Ops = ops;
                box.Height = Pt(cc.CellBeforeRowMargin) + contentHeight + Pt(cc.CellAfterRowMargin);
                boxes.Add(box);
                column = end;
            }
        }

        var heights = new double[rows.Count];
        var exact = new bool[rows.Count];
        for (int r = 0; r < rows.Count; r++)
        {
            // the characteristics of a table-row reach the FOT builder with its cells (the engine
            // pushes the row's style for every cell), so the row is as high as its cells ask; the
            // row node itself carries the characteristics around it (in a table in a cell those of
            // the outer row), which count only for a row without cells
            var asked = rows[r].Children.OfType<PdfTableCell>().Select(cell => cell.Characteristics).ToList();
            if (asked.Count == 0) asked.Add(rows[r].Characteristics);
            var h = boxes.Where(b => b.Row == r && b.RowSpan == 1).Select(b => b.Height).DefaultIfEmpty(0).Max();
            var rowHeight = asked.Max(rc => rc.RowHeight);
            if (rowHeight > 0) { h = Pt(rowHeight); exact[r] = true; }
            else h = Math.Max(h, Pt(asked.Max(rc => rc.MinRowHeight)));
            heights[r] = h;
        }
        foreach (var b in boxes.Where(b => b.RowSpan > 1))
        {
            int last = b.Row + b.RowSpan - 1;
            var sum = heights.Skip(b.Row).Take(b.RowSpan).Sum();
            if (sum < b.Height && !exact[last]) heights[last] += b.Height - sum;
        }
        var tops = new double[rows.Count + 1];
        for (int r = 0; r < rows.Count; r++) tops[r + 1] = tops[r] + heights[r];

        var result = new List<Op>();
        foreach (var b in boxes)
        {
            var cc = b.Cell.Characteristics;
            if (cc.CellBackground && cc.HasBackgroundColor)
                result.Add(new Op { Kind = OpKind.Rect, Background = true, X = b.X, Y = tops[b.Row], W = b.Width, H = tops[b.Row + b.RowSpan] - tops[b.Row],
                    Color = XColor.FromArgb(cc.BackgroundR, cc.BackgroundG, cc.BackgroundB) });
        }
        foreach (var b in boxes)
        {
            var cc = b.Cell.Characteristics;
            var room = tops[b.Row + b.RowSpan] - tops[b.Row] - b.Height;
            var down = room > 0 ? Offset(cc.CellRowAlignment, room) : 0;
            foreach (var op in b.Ops)
            {
                op.Shift(b.X + Pt(cc.CellBeforeColumnMargin), tops[b.Row] + Pt(cc.CellBeforeRowMargin) + down);
                // the backgrounds of a table in the cell stay over the background of the cell
                op.Background = false;
                result.Add(op);
            }
        }
        foreach (var b in boxes)
        {
            double y1 = tops[b.Row], y2 = tops[b.Row + b.RowSpan], x1 = b.X, x2 = b.X + b.Width;
            if (b.Cell.BeforeRowBorder != null) result.Add(BorderLine(x1, y1, x2, y1, b.Cell.BeforeRowBorder));
            if (b.Cell.AfterRowBorder != null) result.Add(BorderLine(x1, y2, x2, y2, b.Cell.AfterRowBorder));
            if (b.Cell.BeforeColumnBorder != null) result.Add(BorderLine(x1, y1, x1, y2, b.Cell.BeforeColumnBorder));
            if (b.Cell.AfterColumnBorder != null) result.Add(BorderLine(x2, y1, x2, y2, b.Cell.AfterColumnBorder));
        }
        return (result, tops[rows.Count], tops);
    }

    // The content of a cell, set into an area of its own width from the top left; conditional
    // space at its start and end is dropped
    private (List<Op> Ops, double Height) SetCell(PdfTableCell cell, double width)
    {
        var saved = area_;
        area_ = new Area { X = 0, Width = Math.Max(0, width), NoBreak = true, Bottom = double.MaxValue };
        FlowChildren(cell.Children, null, false);
        var height = area_.Y + area_.PendingFixed;
        var marks = area_.Marks;
        foreach (var name in marks)
            area_.Ops.Add(new Op { Kind = OpKind.Mark, X = 0, Y = 0, Text = name });
        var ops = area_.Ops;
        area_ = saved;
        return (ops, height);
    }

    private static Op BorderLine(double x1, double y1, double x2, double y2, PdfBorder border) => new()
    {
        Kind = OpKind.Line, X = x1, Y = y1, X2 = x2, Y2 = y2, Thickness = border.LineThicknessPt,
        Color = XColor.FromArgb(border.ColorR, border.ColorG, border.ColorB),
        Dash = Dash(border.LineDash), DashOffset = Pt(border.LineDashOffset)
    };

    // ==================== Page numbers ====================

    private string FormatOfPage(int page) =>
        page >= 1 && page <= Pages.Count ? Pages[page - 1].NumberFormat : seq_?.Characteristics.PageNumberFormat ?? "1";

    public static string FormatPageNumber(int number, string format) => format switch
    {
        "i" => ToRoman(number).ToLowerInvariant(),
        "I" => ToRoman(number),
        "a" => ToAlpha(number).ToLowerInvariant(),
        "A" => ToAlpha(number),
        _ => number.ToString(System.Globalization.CultureInfo.InvariantCulture)
    };

    private static string ToRoman(int number)
    {
        if (number <= 0) return number.ToString(System.Globalization.CultureInfo.InvariantCulture);
        var sb = new System.Text.StringBuilder();
        var values = new[] { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
        var symbols = new[] { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
        for (int i = 0; i < values.Length; i++)
            while (number >= values[i])
            {
                sb.Append(symbols[i]);
                number -= values[i];
            }
        return sb.ToString();
    }

    private static string ToAlpha(int number)
    {
        if (number <= 0) return number.ToString(System.Globalization.CultureInfo.InvariantCulture);
        var sb = new System.Text.StringBuilder();
        while (number > 0)
        {
            number--;
            sb.Insert(0, (char)('A' + number % 26));
            number /= 26;
        }
        return sb.ToString();
    }
}
