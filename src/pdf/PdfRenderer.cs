namespace Dazzle.Pdf;

using QuestPDF.Fluent;
using QuestPDF.Infrastructure;
using Symbol = OpenJade.Style.FOTBuilder.Symbol;

public class PdfRenderer
{
    public void Render(List<PdfPageSequence> pageSequences, string outputPath)
    {
        QuestPDF.Settings.License = LicenseType.Community;

        Document.Create(container =>
        {
            foreach (var pageSequence in pageSequences)
                RenderPageSequence(container, pageSequence);
        }).GeneratePdf(outputPath);
    }

    public void Render(List<PdfPageSequence> pageSequences, Stream stream)
    {
        QuestPDF.Settings.License = LicenseType.Community;

        Document.Create(container =>
        {
            foreach (var pageSequence in pageSequences)
                RenderPageSequence(container, pageSequence);
        }).GeneratePdf(stream);
    }

    private void RenderPageSequence(IDocumentContainer container, PdfPageSequence pageSequence)
    {
        var chars = pageSequence.Characteristics;

        container.Page(page =>
        {
            page.Size(chars.PageWidthPt, chars.PageHeightPt, Unit.Point);
            page.MarginLeft(chars.LeftMarginPt, Unit.Point);
            page.MarginRight(chars.RightMarginPt, Unit.Point);
            page.MarginTop(chars.TopMarginPt, Unit.Point);
            page.MarginBottom(chars.BottomMarginPt, Unit.Point);

            page.Content().Column(col =>
            {
                RenderChildren(col, pageSequence.Children);
            });
        });
    }

    private void RenderChildren(ColumnDescriptor col, List<PdfNode> children)
    {
        foreach (var child in children)
        {
            col.Item().Element(container => RenderNode(container, child));
        }
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
        }
    }

    private static void RenderParagraph(IContainer container, PdfParagraph para)
    {
        var chars = para.Characteristics;
        var styled = ApplyBlockCharacteristics(container, chars);

        styled.Text(text =>
        {
            text.DefaultTextStyle(BuildTextStyle(chars));

            if (chars.Quadding == Symbol.symbolCenter)
                text.AlignCenter();
            else if (chars.Quadding == Symbol.symbolEnd)
                text.AlignRight();
            else if (chars.Quadding == Symbol.symbolJustify)
                text.Justify();

            foreach (var child in para.Children)
            {
                switch (child)
                {
                    case PdfTextRun run:
                        text.Span(run.Text).Style(BuildTextStyle(run.Characteristics));
                        break;
                    case PdfPageNumber:
                        text.CurrentPageNumber();
                        break;
                }
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
        try
        {
            var styled = container;
            if (graphic.HasMaxWidth)
                styled = styled.MaxWidth(PdfCharacteristics.ToPoints(graphic.MaxWidth), Unit.Point);
            if (graphic.HasMaxHeight)
                styled = styled.MaxHeight(PdfCharacteristics.ToPoints(graphic.MaxHeight), Unit.Point);
            styled.Image(graphic.SystemId);
        }
        catch
        {
            container.Text(text =>
            {
                text.Span($"[Image: {graphic.SystemId}]");
            });
        }
    }

    // Apply block-level characteristics (spacing, indentation)
    private static IContainer ApplyBlockCharacteristics(IContainer container, PdfCharacteristics chars)
    {
        IContainer result = container;

        if (chars.SpaceBefore > 0)
            result = result.PaddingTop(chars.SpaceBeforePt, Unit.Point);
        if (chars.SpaceAfter > 0)
            result = result.PaddingBottom(chars.SpaceAfterPt, Unit.Point);
        if (chars.StartIndent > 0)
            result = result.PaddingLeft(chars.StartIndentPt, Unit.Point);
        if (chars.EndIndent > 0)
            result = result.PaddingRight(chars.EndIndentPt, Unit.Point);

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
            style = style.LineHeight(chars.LineSpacingPt / chars.FontSizePt);

        return style;
    }
}
