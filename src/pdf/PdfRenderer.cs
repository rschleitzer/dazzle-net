namespace Dazzle.Pdf;

using QuestPDF.Fluent;
using QuestPDF.Infrastructure;
using Symbol = OpenJade.Style.FOTBuilder.Symbol;
using HF = OpenJade.Style.FOTBuilder.HF;

public class PdfRenderer
{
    public void Render(List<PdfPageSequence> pageSequences, string outputPath)
    {
        QuestPDF.Settings.License = LicenseType.Community;

        var nonEmpty = pageSequences.Where(ps => ps.Children.Count > 0).ToList();
        Document.Create(container =>
        {
            foreach (var pageSequence in nonEmpty)
                RenderPageSequence(container, pageSequence);
        }).GeneratePdf(outputPath);
    }

    public void Render(List<PdfPageSequence> pageSequences, Stream stream)
    {
        QuestPDF.Settings.License = LicenseType.Community;

        var nonEmpty = pageSequences.Where(ps => ps.Children.Count > 0).ToList();
        Document.Create(container =>
        {
            foreach (var pageSequence in nonEmpty)
                RenderPageSequence(container, pageSequence);
        }).GeneratePdf(stream);
    }

    private int pageSequenceIndex_;

    private void RenderPageSequence(IDocumentContainer container, PdfPageSequence pageSequence)
    {
        var chars = pageSequence.Characteristics;
        var sectionName = $"__ps_{pageSequenceIndex_++}";

        container.Page(page =>
        {
            page.Size(chars.PageWidthPt, chars.PageHeightPt, Unit.Point);
            page.MarginLeft(chars.LeftMarginPt, Unit.Point);
            page.MarginRight(chars.RightMarginPt, Unit.Point);
            page.MarginTop(chars.TopMarginPt, Unit.Point);
            page.MarginBottom(0);

            // Header: first-page with ShowOnce, other-pages with SkipOnce
            page.Header().Column(col =>
            {
                col.Item().ShowOnce().Element(c =>
                    RenderHeaderFooter(c, pageSequence,
                        (int)(HF.firstHF | HF.frontHF | HF.headerHF)));
                col.Item().SkipOnce().Element(c =>
                    RenderHeaderFooter(c, pageSequence,
                        (int)(HF.otherHF | HF.frontHF | HF.headerHF)));
            });

            // Footer: full margin height, positioned like RTF (footerMargin=0.5in from page edge)
            const float footerMarginPt = 36f;
            page.Footer()
                .Height(chars.BottomMarginPt, Unit.Point)
                .AlignBottom()
                .PaddingBottom(footerMarginPt, Unit.Point)
                .Column(col =>
                {
                    col.Item().ShowOnce().Element(c =>
                        RenderHeaderFooter(c, pageSequence,
                            (int)(HF.firstHF | HF.frontHF | HF.footerHF)));
                    col.Item().SkipOnce().Element(c =>
                        RenderHeaderFooter(c, pageSequence,
                            (int)(HF.otherHF | HF.frontHF | HF.footerHF)));
                });

            page.Content().Section(sectionName).Column(col =>
            {
                RenderChildren(col, pageSequence.Children);
            });
        });
    }

    private static void RenderHeaderFooter(IContainer hfContainer, PdfPageSequence pageSequence,
        int baseFlags)
    {
        var left = pageSequence.HeaderFooter[baseFlags | (int)HF.leftHF];
        var center = pageSequence.HeaderFooter[baseFlags | (int)HF.centerHF];
        var right = pageSequence.HeaderFooter[baseFlags | (int)HF.rightHF];

        bool hasContent = left.Count > 0 || center.Count > 0 || right.Count > 0;
        if (!hasContent) return;

        var format = pageSequence.Characteristics.PageNumberFormat;

        hfContainer.Row(row =>
        {
            // Left-aligned
            row.RelativeItem().Text(text =>
            {
                RenderHFSlotInline(text, left, format);
            });

            // Center-aligned
            row.RelativeItem().Text(text =>
            {
                text.AlignCenter();
                RenderHFSlotInline(text, center, format);
            });

            // Right-aligned
            row.RelativeItem().Text(text =>
            {
                text.AlignRight();
                RenderHFSlotInline(text, right, format);
            });
        });
    }

    private static void RenderHFSlotInline(TextDescriptor text, List<PdfNode> nodes, string format)
    {
        foreach (var node in nodes)
        {
            switch (node)
            {
                case PdfParagraph para:
                    RenderHFInlineContent(text, para.Children, format);
                    break;
                case PdfContainerNode container:
                    RenderHFInlineContent(text, container.Children, format);
                    break;
                default:
                    RenderHFInlineContent(text, new List<PdfNode> { node }, format);
                    break;
            }
        }
    }

    private static void RenderHFInlineContent(TextDescriptor text, List<PdfNode> children,
        string format)
    {
        foreach (var child in children)
        {
            switch (child)
            {
                case PdfTextRun run:
                    text.Span(run.Text).Style(BuildTextStyle(run.Characteristics));
                    break;
                case PdfPageNumber pn:
                    text.CurrentPageNumber()
                        .Format(n => FormatPageNumber(n, format))
                        .Style(BuildTextStyle(pn.Characteristics));
                    break;
                case PdfSequence seq:
                    RenderHFInlineContent(text, seq.Children, format);
                    break;
            }
        }
    }

    private static string FormatPageNumber(int? pageNumber, string format)
    {
        if (pageNumber == null) return "0";
        int n = pageNumber.Value;
        return format switch
        {
            "i" => ToRomanLower(n),
            "I" => ToRomanLower(n).ToUpperInvariant(),
            _ => n.ToString()
        };
    }

    private static string ToRomanLower(int number)
    {
        if (number <= 0) return number.ToString();
        ReadOnlySpan<(int value, string numeral)> table =
        [
            (1000, "m"), (900, "cm"), (500, "d"), (400, "cd"),
            (100, "c"), (90, "xc"), (50, "l"), (40, "xl"),
            (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i")
        ];
        var sb = new System.Text.StringBuilder();
        foreach (var (value, numeral) in table)
            while (number >= value) { sb.Append(numeral); number -= value; }
        return sb.ToString();
    }

    private static void RenderInlineContent(TextDescriptor text, List<PdfNode> children)
    {
        foreach (var child in children)
        {
            switch (child)
            {
                case PdfTextRun run:
                    text.Span(run.Text).Style(BuildTextStyle(run.Characteristics));
                    break;
                case PdfPageNumber pn:
                    text.CurrentPageNumber().Style(BuildTextStyle(pn.Characteristics));
                    break;
                case PdfSequence seq:
                    RenderInlineContent(text, seq.Children);
                    break;
                case PdfLeader leader:
                    // Emit leader character (typically ".") repeated as fill
                    string dot = ".";
                    if (leader.Children.Count > 0 && leader.Children[0] is PdfTextRun lr)
                        dot = lr.Text;
                    text.Span(" " + new string(dot[0], 40) + " ")
                        .Style(BuildTextStyle(leader.Characteristics));
                    break;
                case PdfNodePageNumber npn:
                    text.BeginPageNumberOfSection(npn.LocationName);
                    break;
            }
        }
    }

    private static bool HasDisplayBreak(PdfNode node) =>
        node is PdfParagraph or PdfDisplayGroup;

    private void RenderChildren(ColumnDescriptor col, List<PdfNode> children,
        float accumSpace = 0, float prevBottomHalfLeading = 0)
    {
        bool pendingBreak = false;
        foreach (var child in children)
        {
            bool breakBefore = HasDisplayBreak(child)
                && child.Characteristics.BreakBefore == Symbol.symbolPage;

            // Collapse break-after + break-before into a single page break
            if (pendingBreak || breakBefore)
            {
                col.Item().PageBreak();
                accumSpace = 0;
            }
            pendingBreak = false;

            // DSSSL space model: collapse adjacent SpaceAfter/SpaceBefore (take max).
            // Container nodes (display-group, scroll, sequence) don't add visible
            // spacing themselves — their SpaceBefore/After propagates through to
            // children, like RTF's accumSpace_ carries across nesting boundaries.
            // Skip invisible nodes (location marks) — they don't affect spacing
            if (child is PdfLocationMark)
            {
                col.Item().Element(container => RenderNode(container, child));
                continue;
            }

            bool isContainer = child is PdfDisplayGroup or PdfScroll or PdfSequence;
            float spaceBefore = child.Characteristics.SpaceBeforePt;
            float collapsedSpace = Math.Max(accumSpace, spaceBefore);

            // QuestPDF's LineHeight distributes half-leading above and below each
            // line. This creates visual space at paragraph boundaries that stacks
            // with our explicit spacing. Compensate by subtracting the half-leading
            // from the previous element's bottom and current element's top.
            float topHalfLeading = HalfLeading(child.Characteristics);
            // Compensate for ~half of QuestPDF's edge leading (empirically tuned)
            float leadingCompensation = (prevBottomHalfLeading + topHalfLeading) * 0.5f;
            float adjustedSpace = Math.Max(0, collapsedSpace - leadingCompensation);


            if (isContainer)
            {
                var cnode = (PdfContainerNode)child;
                // Capture values for lambda (avoid C# closure over modified variable)
                float capturedCollapsed = collapsedSpace;
                float capturedPrevBHL = prevBottomHalfLeading;
                col.Item().Element(cont =>
                {
                    var styled = ApplyBlockCharacteristics(cont, cnode.Characteristics);
                    styled.Column(innerCol =>
                    {
                        RenderChildren(innerCol, cnode.Children,
                            capturedCollapsed, capturedPrevBHL);
                    });
                });
                prevBottomHalfLeading = LastLeafHalfLeading(cnode);
            }
            else
            {
                float nodeTopHL = topHalfLeading;
                float leafCompensation = (prevBottomHalfLeading + nodeTopHL) * 0.5f;
                float adjustedSpaceFinal = Math.Max(0, collapsedSpace - leafCompensation);
                if (adjustedSpaceFinal > 0)
                    col.Item().PaddingTop(adjustedSpaceFinal, Unit.Point)
                        .Element(container => RenderNode(container, child));
                else
                    col.Item().Element(container => RenderNode(container, child));
                prevBottomHalfLeading = HalfLeading(child.Characteristics);
            }

            accumSpace = child.Characteristics.SpaceAfterPt;

            if (HasDisplayBreak(child)
                && child.Characteristics.BreakAfter == Symbol.symbolPage)
                pendingBreak = true;
        }
    }

    // QuestPDF LineHeight distributes extra space (leading) as half above and half
    // below each line. This half-leading must be subtracted from inter-element
    // spacing to match RTF/Word behavior where line-spacing is internal to paragraphs.
    private static float HalfLeading(PdfCharacteristics chars)
    {
        if (chars.LineSpacing > 0 && chars.FontSize > 0)
            return Math.Max(0, (chars.LineSpacingPt - chars.FontSizePt) / 2);
        return 0;
    }

    private static float FirstLeafHalfLeading(PdfContainerNode container)
    {
        foreach (var child in container.Children)
        {
            // Paragraphs are block-level leaves for spacing purposes.
            // Only recurse into pass-through containers (display-group, scroll, sequence).
            if (child is PdfParagraph para)
                return HalfLeading(para.Characteristics);
            if (child is PdfContainerNode nested)
                return FirstLeafHalfLeading(nested);
            return HalfLeading(child.Characteristics);
        }
        return 0;
    }

    private static float LastLeafHalfLeading(PdfContainerNode container)
    {
        for (int i = container.Children.Count - 1; i >= 0; i--)
        {
            if (container.Children[i] is PdfParagraph para)
                return HalfLeading(para.Characteristics);
            if (container.Children[i] is PdfContainerNode nested)
                return LastLeafHalfLeading(nested);
            return HalfLeading(container.Children[i].Characteristics);
        }
        return 0;
    }

    private void RenderNode(IContainer container, PdfNode node)
    {
        switch (node)
        {
            case PdfParagraph para:
                RenderParagraph(container, para);
                break;
            case PdfDisplayGroup group:
                RenderContainerNode(container, group);
                break;
            case PdfScroll scroll:
                RenderContainerNode(container, scroll);
                break;
            case PdfSequence seq:
                RenderSequence(container, seq);
                break;
            case PdfTextRun run:
                RenderTextRun(container, run);
                break;
            case PdfPageNumber pn:
                RenderPageNumber(container, pn);
                break;
            case PdfRule rule:
                RenderRule(container, rule);
                break;
            case PdfExternalGraphic graphic:
                RenderExternalGraphic(container, graphic);
                break;
            case PdfLocationMark loc:
                container.Section(loc.LocationName);
                break;
        }
    }

    private static void RenderParagraph(IContainer container, PdfParagraph para)
    {
        var chars = para.Characteristics;

        // DSSSL start-indent is absolute from the reference area edge.
        // first-line-start-indent is relative to start-indent.
        // Negative first-line-start-indent = hanging indent (first line outdented).
        // QuestPDF only supports non-negative ParagraphFirstLineIndentation,
        // so for hanging indents we reduce PaddingLeft to the first-line position.
        var styled = ApplyBlockCharacteristics(container, chars, applyIndent: false);
        if (chars.EndIndent > 0)
            styled = styled.PaddingRight(chars.EndIndentPt, Unit.Point);

        var segments = FlattenInline(para.Children);
        int leaderIdx = segments.FindIndex(s => s.Kind == SegmentKind.Leader);

        if (leaderIdx >= 0)
        {
            // TOC-style: leader check takes priority over hanging indent
            float firstLinePt = Math.Max(0, chars.StartIndentPt + chars.FirstLineStartIndentPt);
            styled = styled.PaddingLeft(firstLinePt, Unit.Point);
        }
        else if (chars.FirstLineStartIndent < 0)
        {
            // Hanging indent: label at firstLine position, body at start-indent.
            // Render as Row: left column = label (hangWidth), right column = body text.
            float firstLinePt = Math.Max(0, chars.StartIndentPt + chars.FirstLineStartIndentPt);
            float hangWidth = Math.Abs(chars.FirstLineStartIndentPt);
            styled = styled.PaddingLeft(firstLinePt, Unit.Point);

            styled.Row(row =>
            {
                row.ConstantItem(hangWidth, Unit.Point).Text(text =>
                {
                    text.DefaultTextStyle(BuildTextStyle(chars));
                    if (segments.Count > 0)
                        RenderSegments(text, segments.GetRange(0, 1));
                });
                row.RelativeItem().Text(text =>
                {
                    text.DefaultTextStyle(BuildTextStyle(chars));
                    if (segments.Count > 1)
                        RenderSegments(text, segments.GetRange(1, segments.Count - 1));
                });
            });
            return;
        }
        else
        {
            if (chars.StartIndent > 0)
                styled = styled.PaddingLeft(chars.StartIndentPt, Unit.Point);
        }

        // Leader-based or normal rendering

        if (leaderIdx >= 0)
        {
            var before = segments.GetRange(0, leaderIdx);
            var leaderSeg = segments[leaderIdx];
            var after = segments.GetRange(leaderIdx + 1, segments.Count - leaderIdx - 1);

            styled.Row(row =>
            {
                // Entry title (auto-width, left-aligned)
                row.AutoItem().AlignBottom().Text(text =>
                {
                    text.DefaultTextStyle(BuildTextStyle(chars));
                    RenderSegments(text, before);
                });

                // Dot leader (fills remaining space)
                char dot = leaderSeg.Text != null && leaderSeg.Text.Length > 0
                    ? leaderSeg.Text[0] : '.';
                row.RelativeItem()
                    .AlignBottom()
                    .PaddingHorizontal(2, Unit.Point)
                    .Text(text =>
                    {
                        text.DefaultTextStyle(BuildTextStyle(leaderSeg.Chars));
                        text.ClampLines(1, "");
                        text.Span(new string(dot, 200));
                    });

                // Page number (auto-width, right-aligned)
                row.AutoItem().AlignBottom().Text(text =>
                {
                    text.DefaultTextStyle(BuildTextStyle(chars));
                    RenderSegments(text, after);
                });
            });
        }
        else if (segments.Any(s => s.Kind == SegmentKind.ExternalGraphic))
        {
            // Paragraph with embedded graphics: render as column of text + images
            RenderParagraphWithGraphics(styled, chars, segments);
        }
        else
        {
            styled.Text(text =>
            {
                text.DefaultTextStyle(BuildTextStyle(chars));

                if (chars.FirstLineStartIndent > 0)
                    text.ParagraphFirstLineIndentation(chars.FirstLineStartIndentPt, Unit.Point);

                if (chars.Quadding == Symbol.symbolCenter)
                    text.AlignCenter();
                else if (chars.Quadding == Symbol.symbolEnd)
                    text.AlignRight();
                else if (chars.Quadding == Symbol.symbolJustify)
                    text.Justify();

                RenderSegments(text, segments);
            });
        }
    }

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
                    // Recurse into nested paragraphs (e.g. figure wrappers
                    // containing Paragraph/Paragraph/Sequence with graphics)
                    FlattenInlineRecursive(nestedPara.Children, result);
                    break;
                case PdfExternalGraphic graphic:
                    result.Add(new(SegmentKind.ExternalGraphic, null, null,
                        graphic.Characteristics, graphic));
                    break;
            }
        }
    }

    private static void RenderSegments(TextDescriptor text, List<InlineSegment> segments)
    {
        foreach (var seg in segments)
        {
            switch (seg.Kind)
            {
                case SegmentKind.Text:
                    text.Span(seg.Text!).Style(BuildTextStyle(seg.Chars));
                    break;
                case SegmentKind.PageNumber:
                    text.CurrentPageNumber();
                    break;
                case SegmentKind.NodePageNumber:
                    text.BeginPageNumberOfSection(seg.LocationName!);
                    break;
                // ExternalGraphic is handled by RenderParagraphWithGraphics, not inline
            }
        }
    }

    private static string ResolveImagePath(string systemId)
    {
        string path = systemId;
        if (path.StartsWith("file://"))
            path = new Uri(path).LocalPath;
        if (!Path.IsPathRooted(path))
            path = Path.GetFullPath(path);
        return path;
    }

    // Render a paragraph that contains external graphics as a Column of text+images
    private static void RenderParagraphWithGraphics(IContainer container,
        PdfCharacteristics chars, List<InlineSegment> segments)
    {
        container.Column(col =>
        {
            var textSegs = new List<InlineSegment>();
            foreach (var seg in segments)
            {
                if (seg.Kind == SegmentKind.ExternalGraphic && seg.Graphic != null)
                {
                    // Flush accumulated text segments
                    if (textSegs.Count > 0)
                    {
                        var captured = new List<InlineSegment>(textSegs);
                        col.Item().Text(text =>
                        {
                            text.DefaultTextStyle(BuildTextStyle(chars));
                            RenderSegments(text, captured);
                        });
                        textSegs.Clear();
                    }
                    // Render graphic as block element
                    var graphic = seg.Graphic;
                    col.Item().Element(c => RenderExternalGraphic(c, graphic));
                }
                else
                {
                    textSegs.Add(seg);
                }
            }
            // Flush remaining text
            if (textSegs.Count > 0)
            {
                col.Item().Text(text =>
                {
                    text.DefaultTextStyle(BuildTextStyle(chars));
                    RenderSegments(text, textSegs);
                });
            }
        });
    }

    private void RenderContainerNode(IContainer container, PdfContainerNode group)
    {
        var chars = group.Characteristics;
        var styled = ApplyBlockCharacteristics(container, chars);

        styled.Column(col =>
        {
            RenderChildren(col, group.Children);
        });
    }

    private void RenderSequence(IContainer container, PdfSequence seq)
    {
        // Sequence is a pass-through container — render children in a column
        container.Column(col =>
        {
            RenderChildren(col, seq.Children);
        });
    }

    private static void RenderTextRun(IContainer container, PdfTextRun run)
    {
        container.Text(text =>
        {
            text.Span(run.Text).Style(BuildTextStyle(run.Characteristics));
        });
    }

    private static void RenderPageNumber(IContainer container, PdfPageNumber pn)
    {
        container.Text(text =>
        {
            text.DefaultTextStyle(BuildTextStyle(pn.Characteristics));
            text.CurrentPageNumber();
        });
    }

    private static void RenderRule(IContainer container, PdfRule rule)
    {
        var chars = rule.Characteristics;
        var color = Color.FromRGB(chars.ColorR, chars.ColorG, chars.ColorB);

        if (rule.Orientation == Symbol.symbolHorizontal)
        {
            if (rule.HasLength)
                container.Width(PdfCharacteristics.ToPoints(rule.Length), Unit.Point)
                         .LineHorizontal(1, Unit.Point)
                         .LineColor(color);
            else
                container.LineHorizontal(1, Unit.Point)
                         .LineColor(color);
        }
        else
        {
            container.LineVertical(1, Unit.Point)
                     .LineColor(color);
        }
    }

    private static void RenderExternalGraphic(IContainer container, PdfExternalGraphic graphic)
    {
        string path = ResolveImagePath(graphic.SystemId);
        if (!File.Exists(path))
        {
            container.Text(text => text.Span($"[Image: {graphic.SystemId}]"));
            return;
        }

        var styled = container;
        if (graphic.HasMaxWidth)
            styled = styled.MaxWidth(PdfCharacteristics.ToPoints(graphic.MaxWidth), Unit.Point);
        if (graphic.HasMaxHeight)
            styled = styled.MaxHeight(PdfCharacteristics.ToPoints(graphic.MaxHeight), Unit.Point);
        styled.Image(path).FitWidth();
    }

    // Apply block-level characteristics (spacing, background).
    // Indentation is only applied for paragraphs (applyIndent=true) since
    // DSSSL start-indent is absolute from the reference area edge, not relative
    // to the parent container.
    private static IContainer ApplyBlockCharacteristics(IContainer container, PdfCharacteristics chars,
        bool applyIndent = false)
    {
        IContainer result = container;

        // SpaceBefore/SpaceAfter are handled by RenderChildren with DSSSL collapsing
        // (max of adjacent spaces, not sum). Not applied here.

        if (applyIndent)
        {
            if (chars.StartIndent > 0)
                result = result.PaddingLeft(chars.StartIndentPt, Unit.Point);
            if (chars.EndIndent > 0)
                result = result.PaddingRight(chars.EndIndentPt, Unit.Point);
        }

        if (chars.HasBackgroundColor)
            result = result.Background(Color.FromRGB(chars.BackgroundR, chars.BackgroundG, chars.BackgroundB));

        return result;
    }

    // Build a QuestPDF TextStyle from characteristics
    private static TextStyle BuildTextStyle(PdfCharacteristics chars)
    {
        var style = TextStyle.Default
            .FontFamily(chars.FontFamily)
            .FontSize(chars.FontSizePt)
            .FontColor(Color.FromRGB(chars.ColorR, chars.ColorG, chars.ColorB));

        if (chars.IsBold)
            style = style.Bold();
        if (chars.IsItalic)
            style = style.Italic();

        if (chars.LineSpacing > 0 && chars.FontSizePt > 0)
        {
            float lineHeight = chars.LineSpacingPt / chars.FontSizePt;
            if (lineHeight >= 1.0f)
                style = style.LineHeight(lineHeight);
        }

        return style;
    }
}
