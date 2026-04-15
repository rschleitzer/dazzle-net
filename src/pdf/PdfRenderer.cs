namespace Dazzle.Pdf;

using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using Symbol = OpenJade.Style.FOTBuilder.Symbol;
using HF = OpenJade.Style.FOTBuilder.HF;
using MdColor = MigraDoc.DocumentObjectModel.Color;
using MdUnit = MigraDoc.DocumentObjectModel.Unit;

public class PdfRenderer
{
    public void Render(List<PdfPageSequence> pageSequences, string outputPath)
    {
        var doc = BuildDocument(pageSequences);
        var renderer = new PdfDocumentRenderer();
        renderer.Document = doc;
        renderer.RenderDocument();
        renderer.PdfDocument.Save(outputPath);
    }

    public void Render(List<PdfPageSequence> pageSequences, Stream stream)
    {
        var doc = BuildDocument(pageSequences);
        var renderer = new PdfDocumentRenderer();
        renderer.Document = doc;
        renderer.RenderDocument();
        renderer.PdfDocument.Save(stream, false);
    }

    private Document BuildDocument(List<PdfPageSequence> pageSequences)
    {
        var doc = new Document();
        var nonEmpty = pageSequences.Where(ps => ps.Children.Count > 0).ToList();
        foreach (var ps in nonEmpty)
            RenderPageSequence(doc, ps);
        return doc;
    }

    // ==================== Page Sequence → Section ====================

    private void RenderPageSequence(Document doc, PdfPageSequence pageSequence)
    {
        var chars = pageSequence.Characteristics;
        var section = doc.AddSection();

        section.PageSetup.PageWidth = MdUnit.FromPoint(chars.PageWidthPt);
        section.PageSetup.PageHeight = MdUnit.FromPoint(chars.PageHeightPt);
        section.PageSetup.LeftMargin = MdUnit.FromPoint(chars.LeftMarginPt);
        section.PageSetup.RightMargin = MdUnit.FromPoint(chars.RightMarginPt);
        section.PageSetup.TopMargin = MdUnit.FromPoint(chars.TopMarginPt);
        section.PageSetup.BottomMargin = MdUnit.FromPoint(chars.BottomMarginPt);
        section.PageSetup.DifferentFirstPageHeaderFooter = true;

        var format = chars.PageNumberFormat;

        // Headers
        RenderHeaderFooter(section.Headers.FirstPage, pageSequence,
            (int)(HF.firstHF | HF.frontHF | HF.headerHF), format);
        RenderHeaderFooter(section.Headers.Primary, pageSequence,
            (int)(HF.otherHF | HF.frontHF | HF.headerHF), format);

        // Footers
        RenderHeaderFooter(section.Footers.FirstPage, pageSequence,
            (int)(HF.firstHF | HF.frontHF | HF.footerHF), format);
        RenderHeaderFooter(section.Footers.Primary, pageSequence,
            (int)(HF.otherHF | HF.frontHF | HF.footerHF), format);

        // Content
        RenderChildren(section, pageSequence.Children);
    }

    // ==================== Header/Footer ====================

    private static void RenderHeaderFooter(HeaderFooter hf, PdfPageSequence pageSequence,
        int baseFlags, string format)
    {
        var left = pageSequence.HeaderFooter[baseFlags | (int)HF.leftHF];
        var center = pageSequence.HeaderFooter[baseFlags | (int)HF.centerHF];
        var right = pageSequence.HeaderFooter[baseFlags | (int)HF.rightHF];

        bool hasContent = left.Count > 0 || center.Count > 0 || right.Count > 0;
        if (!hasContent) return;

        // Use a 3-column table for left/center/right layout
        var table = hf.AddTable();
        table.Borders.Visible = false;
        var pageWidth = pageSequence.Characteristics.PageWidthPt
            - pageSequence.Characteristics.LeftMarginPt
            - pageSequence.Characteristics.RightMarginPt;
        var colWidth = pageWidth / 3;
        table.AddColumn(MdUnit.FromPoint(colWidth));
        table.AddColumn(MdUnit.FromPoint(colWidth));
        table.AddColumn(MdUnit.FromPoint(colWidth));

        var row = table.AddRow();

        RenderHFCell(row.Cells[0], left, ParagraphAlignment.Left, format);
        RenderHFCell(row.Cells[1], center, ParagraphAlignment.Center, format);
        RenderHFCell(row.Cells[2], right, ParagraphAlignment.Right, format);
    }

    private static void RenderHFCell(Cell cell, List<PdfNode> nodes,
        ParagraphAlignment alignment, string format)
    {
        if (nodes.Count == 0) return;
        var para = cell.AddParagraph();
        para.Format.Alignment = alignment;

        foreach (var node in nodes)
        {
            switch (node)
            {
                case PdfParagraph p:
                    RenderHFInline(para, p.Children, format);
                    break;
                case PdfContainerNode c:
                    RenderHFInline(para, c.Children, format);
                    break;
                default:
                    RenderHFInline(para, new List<PdfNode> { node }, format);
                    break;
            }
        }
    }

    private static void RenderHFInline(Paragraph para, List<PdfNode> children, string format)
    {
        foreach (var child in children)
        {
            switch (child)
            {
                case PdfTextRun run:
                    var ft = para.AddFormattedText(run.Text);
                    ApplyFont(ft.Font, run.Characteristics);
                    break;
                case PdfPageNumber pn:
                    para.AddPageField();
                    break;
                case PdfSequence seq:
                    RenderHFInline(para, seq.Children, format);
                    break;
            }
        }
    }

    // ==================== Content rendering ====================

    private void RenderChildren(Section section, List<PdfNode> children)
    {
        bool pendingBreak = false;
        foreach (var child in children)
        {
            bool breakBefore = (child is PdfParagraph or PdfDisplayGroup)
                && child.Characteristics.BreakBefore == Symbol.symbolPage;

            if (pendingBreak || breakBefore)
                section.AddPageBreak();
            pendingBreak = false;

            RenderNode(section, child);

            if ((child is PdfParagraph or PdfDisplayGroup)
                && child.Characteristics.BreakAfter == Symbol.symbolPage)
                pendingBreak = true;
        }
    }

    private void RenderChildrenToCell(Cell cell, List<PdfNode> children)
    {
        foreach (var child in children)
            RenderNodeToContainer(cell, child);
    }

    private void RenderNode(Section section, PdfNode node)
    {
        switch (node)
        {
            case PdfTable table:
                RenderTable(section, table);
                break;
            case PdfParagraph para:
                RenderParagraph(section, para);
                break;
            case PdfDisplayGroup group:
                RenderChildren(section, group.Children);
                break;
            case PdfScroll scroll:
                RenderChildren(section, scroll.Children);
                break;
            case PdfSequence seq:
                RenderChildren(section, seq.Children);
                break;
            case PdfTextRun run:
                var p = section.AddParagraph();
                var ft = p.AddFormattedText(run.Text);
                ApplyFont(ft.Font, run.Characteristics);
                break;
            case PdfRule rule:
                RenderRule(section, rule);
                break;
            case PdfExternalGraphic graphic:
                RenderExternalGraphic(section, graphic);
                break;
            case PdfLocationMark loc:
                var bm = section.AddParagraph();
                bm.Format.SpaceBefore = 0;
                bm.Format.SpaceAfter = 0;
                bm.Format.Font.Size = 1;
                bm.AddBookmark(loc.LocationName);
                break;
        }
    }

    private void RenderNodeToContainer(Cell cell, PdfNode node)
    {
        switch (node)
        {
            case PdfParagraph para:
                RenderParagraphToCell(cell, para);
                break;
            case PdfDisplayGroup group:
                RenderChildrenToCell(cell, group.Children);
                break;
            case PdfScroll scroll:
                RenderChildrenToCell(cell, scroll.Children);
                break;
            case PdfSequence seq:
                RenderChildrenToCell(cell, seq.Children);
                break;
            case PdfTextRun run:
                var p = cell.AddParagraph();
                var ft = p.AddFormattedText(run.Text);
                ApplyFont(ft.Font, run.Characteristics);
                break;
            case PdfExternalGraphic graphic:
                RenderExternalGraphicToCell(cell, graphic);
                break;
        }
    }

    // ==================== Paragraph ====================

    private void RenderParagraph(Section section, PdfParagraph para)
    {
        var chars = para.Characteristics;
        var segments = FlattenInline(para.Children);
        int leaderIdx = segments.FindIndex(s => s.Kind == SegmentKind.Leader);

        if (leaderIdx >= 0)
        {
            RenderLeaderParagraph(section, chars, segments, leaderIdx);
            return;
        }

        var mdPara = section.AddParagraph();
        ApplyParagraphFormat(mdPara, chars);
        RenderSegments(mdPara, segments);
    }

    private void RenderParagraphToCell(Cell cell, PdfParagraph para)
    {
        var chars = para.Characteristics;
        var segments = FlattenInline(para.Children);
        var mdPara = cell.AddParagraph();
        ApplyParagraphFormat(mdPara, chars);
        RenderSegments(mdPara, segments);
    }

    private void RenderLeaderParagraph(Section section, PdfCharacteristics chars,
        List<InlineSegment> segments, int leaderIdx)
    {
        var mdPara = section.AddParagraph();
        ApplyParagraphFormat(mdPara, chars);

        // Right-aligned tab stop with dot leader
        var pageWidth = section.PageSetup.PageWidth - section.PageSetup.LeftMargin
            - section.PageSetup.RightMargin;
        char dot = '.';
        if (segments[leaderIdx].Text != null && segments[leaderIdx].Text!.Length > 0)
            dot = segments[leaderIdx].Text![0];
        var leader = dot == '.' ? TabLeader.Dots : TabLeader.Lines;
        mdPara.Format.TabStops.AddTabStop(pageWidth, TabAlignment.Right, leader);

        // Render text before leader
        for (int i = 0; i < leaderIdx; i++)
            RenderSegment(mdPara, segments[i]);

        mdPara.AddTab();

        // Render text after leader (page number etc.)
        for (int i = leaderIdx + 1; i < segments.Count; i++)
            RenderSegment(mdPara, segments[i]);
    }

    // ==================== Inline segments ====================

    private enum SegmentKind { Text, PageNumber, NodePageNumber, Leader, ExternalGraphic }
    private record InlineSegment(SegmentKind Kind, string? Text, string? LocationName,
        PdfCharacteristics Chars, PdfExternalGraphic? Graphic = null);

    private static List<InlineSegment> FlattenInline(List<PdfNode> children)
    {
        var result = new List<InlineSegment>();
        FlattenInlineRecursive(children, result);
        return result;
    }

    private static void FlattenInlineRecursive(List<PdfNode> children, List<InlineSegment> result)
    {
        foreach (var child in children)
        {
            switch (child)
            {
                case PdfTextRun run:
                    result.Add(new(SegmentKind.Text, run.Text, null, run.Characteristics));
                    break;
                case PdfPageNumber pn:
                    result.Add(new(SegmentKind.PageNumber, null, null, pn.Characteristics));
                    break;
                case PdfNodePageNumber npn:
                    result.Add(new(SegmentKind.NodePageNumber, null, npn.LocationName,
                        npn.Characteristics));
                    break;
                case PdfLeader leader:
                    string dot = ".";
                    if (leader.Children.Count > 0 && leader.Children[0] is PdfTextRun lr)
                        dot = lr.Text;
                    result.Add(new(SegmentKind.Leader, dot, null, leader.Characteristics));
                    break;
                case PdfSequence seq:
                    FlattenInlineRecursive(seq.Children, result);
                    break;
                case PdfParagraph nestedPara:
                    FlattenInlineRecursive(nestedPara.Children, result);
                    break;
                case PdfExternalGraphic graphic:
                    result.Add(new(SegmentKind.ExternalGraphic, null, null,
                        graphic.Characteristics, graphic));
                    break;
            }
        }
    }

    private static void RenderSegments(Paragraph para, List<InlineSegment> segments)
    {
        foreach (var seg in segments)
            RenderSegment(para, seg);
    }

    private static void RenderSegment(Paragraph para, InlineSegment seg)
    {
        switch (seg.Kind)
        {
            case SegmentKind.Text:
                if (seg.Chars.PositionPointShift > 0)
                {
                    var sup = para.AddFormattedText(seg.Text!);
                    ApplyFont(sup.Font, seg.Chars);
                    sup.Superscript = true;
                }
                else if (seg.Chars.PositionPointShift < 0)
                {
                    var sub = para.AddFormattedText(seg.Text!);
                    ApplyFont(sub.Font, seg.Chars);
                    sub.Subscript = true;
                }
                else
                {
                    var ft = para.AddFormattedText(seg.Text!);
                    ApplyFont(ft.Font, seg.Chars);
                }
                break;
            case SegmentKind.PageNumber:
                para.AddPageField();
                break;
            case SegmentKind.NodePageNumber:
                // MigraDoc doesn't have cross-reference page numbers.
                // Add a bookmark reference placeholder.
                para.AddText("?");
                break;
            case SegmentKind.ExternalGraphic when seg.Graphic != null:
                var path = ResolveImagePath(seg.Graphic.SystemId);
                if (File.Exists(path))
                    para.AddImage(path);
                break;
        }
    }

    // ==================== Table ====================

    private void RenderTable(Section section, PdfTable pdfTable)
    {
        var table = section.AddTable();
        table.Borders.Width = 0.5;
        table.Borders.Color = Colors.Black;

        // Column definitions
        foreach (var col in pdfTable.Columns)
        {
            if (col.HasWidth && col.Width > 0)
                table.AddColumn(MdUnit.FromPoint(col.WidthPt));
            else
                table.AddColumn();
        }

        // Header rows
        foreach (var row in pdfTable.HeaderRows)
        {
            var mdRow = table.AddRow();
            mdRow.HeadingFormat = true;
            RenderTableRowCells(mdRow, row);
        }

        // Body rows
        foreach (var row in pdfTable.BodyRows)
        {
            var mdRow = table.AddRow();
            RenderTableRowCells(mdRow, row);
        }

        // Footer rows
        foreach (var row in pdfTable.FooterRows)
        {
            var mdRow = table.AddRow();
            RenderTableRowCells(mdRow, row);
        }
    }

    private void RenderTableRowCells(Row mdRow, PdfTableRow pdfRow)
    {
        int colIdx = 0;
        foreach (var child in pdfRow.Children)
        {
            if (child is not PdfTableCell pdfCell) continue;

            var cell = mdRow.Cells[colIdx];
            var chars = pdfCell.Characteristics;

            // Cell spanning
            if (pdfCell.NColumnsSpanned > 1)
                cell.MergeRight = (int)(pdfCell.NColumnsSpanned - 1);
            if (pdfCell.NRowsSpanned > 1)
                cell.MergeDown = (int)(pdfCell.NRowsSpanned - 1);

            // Background color
            if (chars.HasBackgroundColor)
                cell.Shading.Color = ToMdColor(chars.BackgroundR, chars.BackgroundG, chars.BackgroundB);

            // Vertical alignment
            if (chars.CellRowAlignment == Symbol.symbolCenter)
                cell.VerticalAlignment = VerticalAlignment.Center;
            else if (chars.CellRowAlignment == Symbol.symbolEnd)
                cell.VerticalAlignment = VerticalAlignment.Bottom;

            // Cell content
            RenderChildrenToCell(cell, pdfCell.Children);

            // Apply cell margins to the paragraphs inside
            foreach (var obj in cell.Elements)
            {
                if (obj is Paragraph p)
                {
                    if (chars.CellBeforeColumnMargin > 0)
                        p.Format.LeftIndent = MdUnit.FromPoint(chars.CellBeforeColumnMarginPt);
                    if (chars.CellAfterColumnMargin > 0)
                        p.Format.RightIndent = MdUnit.FromPoint(chars.CellAfterColumnMarginPt);
                    if (chars.CellBeforeRowMargin > 0)
                        p.Format.SpaceBefore = MdUnit.FromPoint(chars.CellBeforeRowMarginPt);
                    if (chars.CellAfterRowMargin > 0)
                        p.Format.SpaceAfter = MdUnit.FromPoint(chars.CellAfterRowMarginPt);
                }
            }

            colIdx += (int)pdfCell.NColumnsSpanned;
        }
    }

    // ==================== Rule ====================

    private static void RenderRule(Section section, PdfRule rule)
    {
        var chars = rule.Characteristics;
        if (rule.Orientation == Symbol.symbolHorizontal)
        {
            var para = section.AddParagraph();
            para.Format.SpaceBefore = MdUnit.FromPoint(chars.SpaceBeforePt);
            para.Format.SpaceAfter = MdUnit.FromPoint(chars.SpaceAfterPt);
            para.Format.Borders.Bottom.Width = 1;
            para.Format.Borders.Bottom.Color = ToMdColor(chars.ColorR, chars.ColorG, chars.ColorB);
        }
    }

    // ==================== External Graphic ====================

    private static string ResolveImagePath(string systemId)
    {
        string path = systemId;
        if (path.StartsWith("file://"))
            path = new Uri(path).LocalPath;
        if (!Path.IsPathRooted(path))
            path = Path.GetFullPath(path);
        return path;
    }

    private static void RenderExternalGraphic(Section section, PdfExternalGraphic graphic)
    {
        string path = ResolveImagePath(graphic.SystemId);
        if (!File.Exists(path)) return;

        var image = section.AddImage(path);
        if (graphic.HasMaxWidth)
            image.Width = MdUnit.FromPoint(PdfCharacteristics.ToPoints(graphic.MaxWidth));
        if (graphic.HasMaxHeight)
            image.Height = MdUnit.FromPoint(PdfCharacteristics.ToPoints(graphic.MaxHeight));
    }

    private static void RenderExternalGraphicToCell(Cell cell, PdfExternalGraphic graphic)
    {
        string path = ResolveImagePath(graphic.SystemId);
        if (!File.Exists(path)) return;

        var para = cell.AddParagraph();
        var image = para.AddImage(path);
        if (graphic.HasMaxWidth)
            image.Width = MdUnit.FromPoint(PdfCharacteristics.ToPoints(graphic.MaxWidth));
        if (graphic.HasMaxHeight)
            image.Height = MdUnit.FromPoint(PdfCharacteristics.ToPoints(graphic.MaxHeight));
    }

    // ==================== Formatting helpers ====================

    private static void ApplyParagraphFormat(Paragraph para, PdfCharacteristics chars)
    {
        ApplyFont(para.Format.Font, chars);

        if (chars.SpaceBefore > 0)
            para.Format.SpaceBefore = MdUnit.FromPoint(chars.SpaceBeforePt);
        if (chars.SpaceAfter > 0)
            para.Format.SpaceAfter = MdUnit.FromPoint(chars.SpaceAfterPt);

        if (chars.StartIndent > 0)
            para.Format.LeftIndent = MdUnit.FromPoint(chars.StartIndentPt);
        if (chars.EndIndent > 0)
            para.Format.RightIndent = MdUnit.FromPoint(chars.EndIndentPt);
        if (chars.FirstLineStartIndent != 0)
            para.Format.FirstLineIndent = MdUnit.FromPoint(chars.FirstLineStartIndentPt);

        if (chars.LineSpacing > 0 && chars.FontSize > 0)
            para.Format.LineSpacing = MdUnit.FromPoint(chars.LineSpacingPt);

        if (chars.Quadding == Symbol.symbolCenter)
            para.Format.Alignment = ParagraphAlignment.Center;
        else if (chars.Quadding == Symbol.symbolEnd)
            para.Format.Alignment = ParagraphAlignment.Right;
        else if (chars.Quadding == Symbol.symbolJustify)
            para.Format.Alignment = ParagraphAlignment.Justify;

        if (chars.HasBackgroundColor)
            para.Format.Shading.Color = ToMdColor(chars.BackgroundR, chars.BackgroundG, chars.BackgroundB);

        if (chars.KeepWithNext)
            para.Format.KeepWithNext = true;
    }

    private static void ApplyFont(Font font, PdfCharacteristics chars)
    {
        font.Name = chars.FontFamily;
        font.Size = chars.FontSizePt;
        font.Color = ToMdColor(chars.ColorR, chars.ColorG, chars.ColorB);

        if (chars.IsBold)
            font.Bold = true;
        if (chars.IsItalic)
            font.Italic = true;
    }

    private static MdColor ToMdColor(byte r, byte g, byte b)
    {
        return new MdColor(r, g, b);
    }
}
