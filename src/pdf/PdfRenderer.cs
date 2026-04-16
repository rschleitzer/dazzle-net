namespace Dazzle.Pdf;

using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using Symbol = OpenJade.Style.FOTBuilder.Symbol;
using HF = OpenJade.Style.FOTBuilder.HF;
using PdfSharp.Drawing;
using PdfSharp.Fonts;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Advanced;
using MdColor = MigraDoc.DocumentObjectModel.Color;
using MdUnit = MigraDoc.DocumentObjectModel.Unit;

public class PdfRenderer
{
    // Placeholder for non-Arabic page numbers in headers/footers.
    // Replaced after rendering via XGraphics overlay.
    private const string PageNumberPlaceholder = "~pn~";

    // Page number format per section (index → DSSSL format string)
    private List<string> sectionFormats_ = new();

    // Font info for non-Arabic page numbers (captured during HF rendering)
    // Available content width for the current section (for image sizing)
    private static MdUnit availableWidth_;

    private static string hfPageFontFamily_ = "Times New Roman";
    private static double hfPageFontSize_ = 10;
    private static bool hfPageFontBold_;
    private static bool hfPageFontItalic_;

    public void Render(List<PdfPageSequence> pageSequences, string outputPath)
    {
        EnsureFontResolver();

        // Pass 1: render to compute page layout and bookmark positions
        var doc1 = BuildDocument(pageSequences);
        var r1 = new PdfDocumentRenderer();
        r1.Document = doc1;
        r1.RenderDocument();
        var resolvedPages = ResolvePageNumbers(r1.PdfDocument);

        // Pass 2: replace PdfNodePageNumber with formatted text, render final PDF
        RewritePageNumbers(pageSequences, resolvedPages);
        var doc2 = BuildDocument(pageSequences);
        var r2 = new PdfDocumentRenderer();
        r2.Document = doc2;
        r2.RenderDocument();
        var sectionPages = ApplyPageLabels(r2.PdfDocument);
        ReplaceHeaderFooterPageNumbers(r2.PdfDocument, sectionPages);
        r2.PdfDocument.Save(outputPath);
    }

    public void Render(List<PdfPageSequence> pageSequences, Stream stream)
    {
        EnsureFontResolver();

        var doc1 = BuildDocument(pageSequences);
        var r1 = new PdfDocumentRenderer();
        r1.Document = doc1;
        r1.RenderDocument();
        var resolvedPages = ResolvePageNumbers(r1.PdfDocument);

        RewritePageNumbers(pageSequences, resolvedPages);
        var doc2 = BuildDocument(pageSequences);
        var r2 = new PdfDocumentRenderer();
        r2.Document = doc2;
        r2.RenderDocument();
        var sectionPages = ApplyPageLabels(r2.PdfDocument);
        ReplaceHeaderFooterPageNumbers(r2.PdfDocument, sectionPages);
        r2.PdfDocument.Save(stream, false);
    }

    // Returns ordered list of (sectionIdx, startPage) pairs.
    private List<(int secIdx, int startPage)> ApplyPageLabels(PdfDocument pdfDoc)
    {
        if (sectionFormats_.Count == 0)
            return new();

        // Section marker bookmarks "__sec_0", "__sec_1", ... were inserted during
        // document building. Find their page numbers in the PDF named destinations.
        var pageForSection = new Dictionary<int, int>(); // sectionIdx → 0-based page

        // PdfSharp stores named destinations in /Names → /Dests → /Names array.
        // Objects may be indirect references, so dereference where needed.
        var namesObj = pdfDoc.Internals.Catalog.Elements["/Names"];
        var namesDict = Deref(namesObj) as PdfDictionary;
        var destsObj = namesDict?.Elements["/Dests"];
        var destsDict = Deref(destsObj) as PdfDictionary;
        var namesArrayObj = destsDict?.Elements["/Names"];
        var namesArray = Deref(namesArrayObj) as PdfArray;

        if (namesArray != null)
        {
            for (int i = 0; i + 1 < namesArray.Elements.Count; i += 2)
            {
                var nameItem = namesArray.Elements[i] as PdfString;
                if (nameItem == null) continue;
                string name = nameItem.Value;
                if (!name.StartsWith("__sec_")) continue;
                int secIdx = int.Parse(name.Substring(6));

                // Destination is stored as a PdfLiteral: "[N 0 R /XYZ x y null]"
                // Extract the object number N to find the page.
                var dest = namesArray.Elements[i + 1];
                string destStr = dest?.ToString() ?? "";
                var match = System.Text.RegularExpressions.Regex.Match(destStr, @"(\d+)\s+\d+\s+R\b");
                if (match.Success)
                {
                    int objNum = int.Parse(match.Groups[1].Value);
                    for (int p = 0; p < pdfDoc.PageCount; p++)
                    {
                        if (pdfDoc.Pages[p].Reference?.ObjectNumber == objNum)
                        {
                            pageForSection[secIdx] = p;
                            break;
                        }
                    }
                }
            }
        }

        // Fallback: if no bookmarks found, just label all pages with first format
        if (pageForSection.Count == 0)
            pageForSection[0] = 0;

        // Only emit a new page label entry when the format changes.
        // Consecutive sections with the same format continue numbering.
        var nums = new PdfArray();
        string prevFormat = null;
        foreach (var (secIdx, pageIdx) in pageForSection.OrderBy(kv => kv.Value))
        {
            var format = secIdx < sectionFormats_.Count
                ? sectionFormats_[secIdx] : "1";

            if (format == prevFormat)
                continue; // same format, numbering continues

            var labelDict = new PdfDictionary();
            labelDict.Elements.Add("/S", new PdfName(DssslFormatToPdfStyle(format)));
            labelDict.Elements.Add("/St", new PdfInteger(pageIdx + 1)); // 1-based physical page

            nums.Elements.Add(new PdfInteger(pageIdx));
            nums.Elements.Add(labelDict);
            prevFormat = format;
        }

        if (nums.Elements.Count > 0)
        {
            var pageLabels = new PdfDictionary();
            pageLabels.Elements.Add("/Nums", nums);
            pdfDoc.Internals.Catalog.Elements.Add("/PageLabels", pageLabels);
        }

        return pageForSection.OrderBy(kv => kv.Value)
            .Select(kv => (kv.Key, kv.Value)).ToList();
    }

    /// Replace the placeholder string in PDF content streams with formatted page numbers.
    private void ReplaceHeaderFooterPageNumbers(PdfDocument pdfDoc,
        List<(int secIdx, int startPage)> sectionPages)
    {
        if (sectionPages.Count == 0) return;

        // Encode the placeholder as it appears in the PDF content stream.
        // MigraDoc writes Unicode text using a font-specific encoding.
        // The placeholder chars ◇◇ (U+25C7) are encoded as glyph IDs in the content.
        // We need to find and replace them in the raw stream bytes.
        // PdfSharp uses WinAnsiEncoding by default — but ◇ is not in WinAnsi.
        // So MigraDoc will use a CIDFont with Unicode glyph IDs.
        // The hex-encoded form of U+25C7 U+25C7 in a CIDFont is <25C725C7>.

        for (int p = 0; p < pdfDoc.PageCount; p++)
        {
            // Determine which section this page belongs to and compute formatted number
            int secIdx = 0;
            int secStart = 0;
            for (int s = sectionPages.Count - 1; s >= 0; s--)
            {
                if (p >= sectionPages[s].startPage)
                {
                    secIdx = sectionPages[s].secIdx;
                    secStart = sectionPages[s].startPage;
                    break;
                }
            }
            string format = secIdx < sectionFormats_.Count ? sectionFormats_[secIdx] : "1";
            if (format == "1") continue; // Arabic pages don't need replacement

            // Physical page number (1-based) — DSSSL counter runs continuously
            int physicalPage = p + 1;
            string formatted = FormatPageNumber(physicalPage, format);

            OverlayPageNumber(pdfDoc.Pages[p], formatted);
        }
    }

    /// Find the placeholder in the content stream, get its coordinates,
    /// then draw the formatted page number on top using XGraphics.
    private static void OverlayPageNumber(PdfPage page, string formatted)
    {
        byte[] needle = System.Text.Encoding.ASCII.GetBytes(PageNumberPlaceholder);

        for (int ci = 0; ci < page.Contents.Elements.Count; ci++)
        {
            var item = page.Contents.Elements[ci];
            if (item is PdfReference r) item = r.Value;
            if (item is not PdfDictionary dict) continue;
            var stream = dict.Stream;
            if (stream == null) continue;

            stream.TryUncompress();
            byte[] data = stream.Value;
            if (data == null) continue;

            int idx = FindBytes(data, needle);
            if (idx < 0) continue;

            // Parse the Td coordinates before the placeholder.
            // Content stream looks like: ... X Y Td\n(~pn~) Tj ...
            // We need to trace absolute position by walking the text state.
            string before = System.Text.Encoding.ASCII.GetString(data, 0, idx);

            // Find the font size from the last Tf operator
            double fontSize = 10;
            foreach (System.Text.RegularExpressions.Match tfm in
                System.Text.RegularExpressions.Regex.Matches(before, @"/F\d+\s+([\d.]+)\s+Tf"))
            {
                fontSize = double.Parse(tfm.Groups[1].Value,
                    System.Globalization.CultureInfo.InvariantCulture);
            }

            // Find the last BT...Td chain to get absolute position.
            // After BT, Td coordinates are cumulative. We need to sum all Td's since last BT.
            double absX = 0, absY = 0;
            int lastBt = before.LastIndexOf("BT");
            if (lastBt >= 0)
            {
                string sincebt = before.Substring(lastBt);
                foreach (System.Text.RegularExpressions.Match tdm in
                    System.Text.RegularExpressions.Regex.Matches(sincebt,
                        @"([\d.]+)\s+([\d.]+)\s+Td"))
                {
                    absX += double.Parse(tdm.Groups[1].Value,
                        System.Globalization.CultureInfo.InvariantCulture);
                    absY += double.Parse(tdm.Groups[2].Value,
                        System.Globalization.CultureInfo.InvariantCulture);
                }
            }

            // PDF origin = bottom-left. XGraphics origin = top-left.
            double pageHeight = page.Height.Point;
            double gfxY = pageHeight - absY;

            using var gfx = XGraphics.FromPdfPage(page, XGraphicsPdfPageOptions.Append);

            // Build the correct font (captured during HF rendering)
            var xStyle = XFontStyleEx.Regular;
            if (hfPageFontBold_ && hfPageFontItalic_) xStyle = XFontStyleEx.BoldItalic;
            else if (hfPageFontBold_) xStyle = XFontStyleEx.Bold;
            else if (hfPageFontItalic_) xStyle = XFontStyleEx.Italic;
            var font = new XFont(hfPageFontFamily_, fontSize, xStyle);

            // Measure placeholder and replacement widths for right-alignment
            double placeholderWidth = gfx.MeasureString(PageNumberPlaceholder, font).Width;
            double replacementWidth = gfx.MeasureString(formatted, font).Width;

            // The right edge of the placeholder text = absX + placeholderWidth.
            // For right-alignment, draw the replacement so its right edge matches.
            double rightEdge = absX + placeholderWidth;
            double drawX = rightEdge - replacementWidth;

            // Cover old placeholder with white
            gfx.DrawRectangle(XBrushes.White, absX - 1, gfxY - fontSize, placeholderWidth + 2, fontSize + 4);

            // Draw the formatted page number
            gfx.DrawString(formatted, font, XBrushes.Black,
                drawX, gfxY,
                new XStringFormat
                {
                    Alignment = XStringAlignment.Near,
                    LineAlignment = XLineAlignment.BaseLine
                });
            return;
        }
    }

    private static int FindBytes(byte[] haystack, byte[] needle)
    {
        for (int i = 0; i <= haystack.Length - needle.Length; i++)
        {
            bool match = true;
            for (int j = 0; j < needle.Length; j++)
            {
                if (haystack[i + j] != needle[j]) { match = false; break; }
            }
            if (match) return i;
        }
        return -1;
    }

    private static PdfItem Deref(PdfItem item)
    {
        while (item is PdfReference r)
            item = r.Value;
        return item;
    }

    private static string DssslFormatToPdfStyle(string format) => format switch
    {
        "i" => "/r",  // lowercase roman
        "I" => "/R",  // uppercase roman
        "a" => "/a",  // lowercase alpha
        "A" => "/A",  // uppercase alpha
        _ => "/D"     // decimal (Arabic)
    };

    // ==================== Two-pass page number resolution ====================

    /// Extract bookmark→page and section→page mappings from the rendered PDF.
    /// Returns a map: locationName → formatted page number string.
    private Dictionary<string, string> ResolvePageNumbers(PdfDocument pdfDoc)
    {
        var result = new Dictionary<string, string>();
        var bookmarkPage = new Dictionary<string, int>();   // bookmark → 0-based page
        var sectionStart = new Dictionary<int, int>();       // sectionIdx → 0-based page

        // Parse Named Destinations
        var namesObj = pdfDoc.Internals.Catalog.Elements["/Names"];
        var namesDict = Deref(namesObj) as PdfDictionary;
        var destsObj = namesDict?.Elements["/Dests"];
        var destsDict = Deref(destsObj) as PdfDictionary;
        var namesArray = Deref(destsDict?.Elements["/Names"]) as PdfArray;

        if (namesArray == null)
            return result;

        for (int i = 0; i + 1 < namesArray.Elements.Count; i += 2)
        {
            var nameItem = namesArray.Elements[i] as PdfString;
            if (nameItem == null) continue;
            string name = nameItem.Value;

            var destStr = namesArray.Elements[i + 1]?.ToString() ?? "";
            var match = System.Text.RegularExpressions.Regex.Match(destStr, @"(\d+)\s+\d+\s+R\b");
            if (!match.Success) continue;

            int objNum = int.Parse(match.Groups[1].Value);
            int pageIdx = -1;
            for (int p = 0; p < pdfDoc.PageCount; p++)
            {
                if (pdfDoc.Pages[p].Reference?.ObjectNumber == objNum)
                {
                    pageIdx = p;
                    break;
                }
            }
            if (pageIdx < 0) continue;

            if (name.StartsWith("__sec_"))
                sectionStart[int.Parse(name.Substring(6))] = pageIdx;
            else
                bookmarkPage[name] = pageIdx;
        }

        // Build ordered section boundaries: [(sectionIdx, startPage, format), ...]
        var sections = sectionStart.OrderBy(kv => kv.Value).ToList();

        // For each bookmark, find its section and compute formatted page number
        foreach (var (bm, pageIdx) in bookmarkPage)
        {
            int secIdx = 0;
            int secStart = 0;
            for (int s = sections.Count - 1; s >= 0; s--)
            {
                if (pageIdx >= sections[s].Value)
                {
                    secIdx = sections[s].Key;
                    secStart = sections[s].Value;
                    break;
                }
            }
            // Physical page number (1-based) — DSSSL counter runs continuously
            int physicalPage = pageIdx + 1;
            string format = secIdx < sectionFormats_.Count ? sectionFormats_[secIdx] : "1";
            result[bm] = FormatPageNumber(physicalPage, format);
        }

        return result;
    }

    /// Walk the PdfNode tree and replace PdfNodePageNumber with PdfTextRun.
    private static void RewritePageNumbers(List<PdfPageSequence> pageSequences,
        Dictionary<string, string> resolvedPages)
    {
        if (resolvedPages.Count == 0) return;
        foreach (var ps in pageSequences)
            RewriteInChildren(ps.Children, resolvedPages);
    }

    private static void RewriteInChildren(List<PdfNode> children,
        Dictionary<string, string> resolvedPages)
    {
        for (int i = 0; i < children.Count; i++)
        {
            if (children[i] is PdfNodePageNumber npn
                && resolvedPages.TryGetValue(npn.LocationName, out var formatted))
            {
                children[i] = new PdfTextRun(formatted, npn.Characteristics);
            }
            else if (children[i] is PdfContainerNode container)
            {
                RewriteInChildren(container.Children, resolvedPages);
            }
        }
    }

    private static string FormatPageNumber(int number, string format) => format switch
    {
        "i" => ToRoman(number).ToLower(),
        "I" => ToRoman(number),
        "a" => ToAlpha(number).ToLower(),
        "A" => ToAlpha(number),
        _ => number.ToString()
    };

    private static string ToRoman(int number)
    {
        if (number <= 0) return number.ToString();
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
        if (number <= 0) return number.ToString();
        var sb = new System.Text.StringBuilder();
        while (number > 0)
        {
            number--;
            sb.Insert(0, (char)('A' + number % 26));
            number /= 26;
        }
        return sb.ToString();
    }

    private static bool fontResolverSet_;

    private static void EnsureFontResolver()
    {
        if (fontResolverSet_)
            return;
        if (GlobalFontSettings.FontResolver == null)
            GlobalFontSettings.FontResolver = new MacFontResolver();
        fontResolverSet_ = true;
    }

    private Document BuildDocument(List<PdfPageSequence> pageSequences)
    {
        sectionFormats_.Clear();
        var doc = new Document();
        var nonEmpty = pageSequences.Where(ps => ps.Children.Count > 0).ToList();
        int sectionIdx = 0;
        foreach (var ps in nonEmpty)
        {
            sectionFormats_.Add(ps.Characteristics.PageNumberFormat);
            RenderPageSequence(doc, ps, sectionIdx);
            sectionIdx++;
        }
        return doc;
    }

    // ==================== Page Sequence → Section ====================

    private void RenderPageSequence(Document doc, PdfPageSequence pageSequence, int sectionIdx)
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
        availableWidth_ = section.PageSetup.PageWidth - section.PageSetup.LeftMargin
            - section.PageSetup.RightMargin;

        // Don't set StartingNumber — DSSSL page counter continues across sections.
        // MigraDoc continues numbering from the previous section by default.

        // Section marker bookmark for page label assignment
        var marker = section.AddParagraph();
        marker.Format.SpaceBefore = 0;
        marker.Format.SpaceAfter = 0;
        marker.Format.Font.Size = 1;
        marker.Format.LineSpacing = MdUnit.FromPoint(0);
        marker.AddBookmark($"__sec_{sectionIdx}");

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
        prevSpaceAfterPt_ = 0;
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

        // DSSSL simple-page-sequence headers/footers are single-line.
        // Use a single paragraph with tab stops for left/center/right alignment.
        var pageWidth = pageSequence.Characteristics.PageWidthPt
            - pageSequence.Characteristics.LeftMarginPt
            - pageSequence.Characteristics.RightMarginPt;

        var para = hf.AddParagraph();
        para.Format.TabStops.AddTabStop(MdUnit.FromPoint(pageWidth / 2), TabAlignment.Center);
        para.Format.TabStops.AddTabStop(MdUnit.FromPoint(pageWidth), TabAlignment.Right);

        // Left content
        RenderHFInline(para, left, format);
        // Center content (always emit tab to maintain alignment)
        para.AddTab();
        RenderHFInline(para, center, format);
        // Right content (always emit tab to maintain alignment)
        para.AddTab();
        RenderHFInline(para, right, format);
    }

    private static void RenderHFInline(Paragraph para, List<PdfNode> nodes, string format)
    {
        foreach (var node in nodes)
        {
            switch (node)
            {
                case PdfTextRun run:
                    var ft = para.AddFormattedText(run.Text);
                    ApplyFont(ft.Font, run.Characteristics);
                    break;
                case PdfPageNumber pn:
                    if (format != "1")
                    {
                        var pft = para.AddFormattedText(PageNumberPlaceholder);
                        ApplyFont(pft.Font, pn.Characteristics);
                        hfPageFontFamily_ = pn.Characteristics.FontFamily;
                        hfPageFontSize_ = pn.Characteristics.FontSizePt;
                        hfPageFontBold_ = pn.Characteristics.IsBold;
                        hfPageFontItalic_ = pn.Characteristics.IsItalic;
                    }
                    else
                    {
                        var pft = para.AddFormattedText();
                        ApplyFont(pft.Font, pn.Characteristics);
                        pft.AddPageField();
                    }
                    break;
                case PdfParagraph p:
                    RenderHFInline(para, p.Children, format);
                    break;
                case PdfContainerNode c:
                    RenderHFInline(para, c.Children, format);
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
        float saved = prevSpaceAfterPt_;
        prevSpaceAfterPt_ = 0;
        foreach (var child in children)
            RenderNodeToContainer(cell, child);
        prevSpaceAfterPt_ = saved;
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
                CollapseDisplaySpace(group.Characteristics);
                RenderChildren(section, group.Children);
                // Collapse group's SpaceAfter with last child's SpaceAfter
                prevSpaceAfterPt_ = Math.Max(group.Characteristics.SpaceAfterPt, prevSpaceAfterPt_);
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
                CollapseDisplaySpace(group.Characteristics);
                RenderChildrenToCell(cell, group.Children);
                prevSpaceAfterPt_ = Math.Max(group.Characteristics.SpaceAfterPt, prevSpaceAfterPt_);
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

        // Empty paragraphs (no visible text): don't emit a MigraDoc paragraph,
        // just let their spacing participate in collapsing.
        if (segments.All(s => s.Kind == SegmentKind.Text && string.IsNullOrWhiteSpace(s.Text)))
        {
            float spaceBefore = chars.SpaceBeforePt;
            prevSpaceAfterPt_ = Math.Max(
                Math.Max(spaceBefore, prevSpaceAfterPt_),
                chars.SpaceAfterPt);
            return;
        }

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
                    // DSSSL already reduced font-size and set position-point-shift.
                    // MigraDoc's Superscript further shrinks the font (~58%), causing
                    // double reduction. Compensate by scaling the font size up so that
                    // after MigraDoc's shrink the result matches the DSSSL size.
                    var sup = para.AddFormattedText(seg.Text!);
                    var compensated = seg.Chars.Clone();
                    compensated.FontSize = (long)(seg.Chars.FontSize / 0.58);
                    ApplyFont(sup.Font, compensated);
                    sup.Superscript = true;
                }
                else if (seg.Chars.PositionPointShift < 0)
                {
                    var sub = para.AddFormattedText(seg.Text!);
                    var compensated = seg.Chars.Clone();
                    compensated.FontSize = (long)(seg.Chars.FontSize / 0.58);
                    ApplyFont(sub.Font, compensated);
                    sub.Subscript = true;
                }
                else
                {
                    AddTextWithLineBreaks(para, seg.Text!, seg.Chars);
                }
                break;
            case SegmentKind.PageNumber:
                para.AddPageField();
                break;
            case SegmentKind.NodePageNumber:
                para.AddPageRefField(seg.LocationName!);
                break;
            case SegmentKind.ExternalGraphic when seg.Graphic != null:
                var imgPath = ResolveImagePath(seg.Graphic.SystemId);
                if (File.Exists(imgPath))
                {
                    var img = para.AddImage(imgPath);
                    img.LockAspectRatio = true;
                    if (seg.Graphic.HasMaxWidth)
                    {
                        var maxW = MdUnit.FromPoint(PdfCharacteristics.ToPoints(seg.Graphic.MaxWidth));
                        img.Width = maxW < availableWidth_ ? maxW : availableWidth_;
                    }
                    else
                    {
                        img.Width = availableWidth_;
                    }
                    if (seg.Graphic.HasMaxHeight)
                        img.Height = MdUnit.FromPoint(PdfCharacteristics.ToPoints(seg.Graphic.MaxHeight));
                }
                break;
        }
    }

    // ==================== Table ====================

    private void RenderTable(Section section, PdfTable pdfTable)
    {
        // Apply display spacing via a spacer paragraph (MigraDoc tables have no SpaceBefore)
        var tChars = pdfTable.Characteristics;
        float tSpaceBefore = tChars.SpaceBeforePt;
        float tCollapsed = Math.Max(tSpaceBefore, prevSpaceAfterPt_);
        if (tCollapsed > 0)
        {
            var spacer = section.AddParagraph();
            spacer.Format.SpaceBefore = 0;
            spacer.Format.SpaceAfter = MdUnit.FromPoint(tCollapsed);
            spacer.Format.Font.Size = 1;
            spacer.Format.LineSpacing = MdUnit.FromPoint(0);
        }
        prevSpaceAfterPt_ = tChars.SpaceAfterPt;

        var table = section.AddTable();
        table.Borders.Visible = false;

        // Table indentation
        if (tChars.StartIndent > 0)
            table.Rows.LeftIndent = MdUnit.FromPoint(tChars.StartIndentPt);

        // Column definitions
        if (pdfTable.Columns.Count > 0)
        {
            foreach (var col in pdfTable.Columns)
            {
                if (col.HasWidth && col.Width > 0)
                    table.AddColumn(MdUnit.FromPoint(col.WidthPt));
                else
                    table.AddColumn();
            }
        }
        else
        {
            // No explicit table-column FOs (e.g. simplelist with %simplelist-column-width% #f).
            // Infer column count from cells and distribute page width equally.
            int colCount = InferColumnCount(pdfTable);
            if (colCount == 0)
                return;
            var pageWidth = section.PageSetup.PageWidth - section.PageSetup.LeftMargin
                - section.PageSetup.RightMargin;
            var colWidth = pageWidth / colCount;
            for (int i = 0; i < colCount; i++)
                table.AddColumn(colWidth);
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

    private static int InferColumnCount(PdfTable pdfTable)
    {
        int max = 0;
        foreach (var row in pdfTable.HeaderRows.Concat(pdfTable.BodyRows).Concat(pdfTable.FooterRows))
        {
            int cols = 0;
            foreach (var child in row.Children)
                if (child is PdfTableCell cell)
                    cols += (int)cell.NColumnsSpanned;
            if (cols > max) max = cols;
        }
        return max;
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
        // Constrain to the available content width so images don't overflow
        var available = section.PageSetup.PageWidth - section.PageSetup.LeftMargin
            - section.PageSetup.RightMargin;

        if (graphic.HasMaxWidth)
        {
            var maxW = MdUnit.FromPoint(PdfCharacteristics.ToPoints(graphic.MaxWidth));
            image.Width = maxW < available ? maxW : available;
        }
        else
        {
            image.LockAspectRatio = true;
            image.Width = available;
        }
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

    /// Collapse a display-group's SpaceBefore with the previous element's SpaceAfter.
    /// Updates prevSpaceAfterPt_ and resets it for the group's children.
    private void CollapseDisplaySpace(PdfCharacteristics chars)
    {
        float spaceBefore = chars.SpaceBeforePt;
        float collapsed = Math.Max(spaceBefore, prevSpaceAfterPt_);
        // The collapsed space is now "consumed" — children start fresh
        prevSpaceAfterPt_ = collapsed;
    }

    // DSSSL display-space collapsing: the gap between two adjacent block elements
    // is max(space-after of first, space-before of second), not the sum.
    // MigraDoc adds them, so we only emit SpaceBefore (collapsed) and skip SpaceAfter.
    private float prevSpaceAfterPt_;

    private void ApplyParagraphFormat(Paragraph para, PdfCharacteristics chars)
    {
        ApplyFont(para.Format.Font, chars);

        float spaceBefore = chars.SpaceBeforePt;
        float collapsed = Math.Max(spaceBefore, prevSpaceAfterPt_);
        if (collapsed > 0)
            para.Format.SpaceBefore = MdUnit.FromPoint(collapsed);
        // Don't set SpaceAfter — it's consumed by the next element's collapsing.
        prevSpaceAfterPt_ = chars.SpaceAfterPt;

        if (chars.StartIndent > 0)
            para.Format.LeftIndent = MdUnit.FromPoint(chars.StartIndentPt);
        if (chars.EndIndent > 0)
            para.Format.RightIndent = MdUnit.FromPoint(chars.EndIndentPt);
        if (chars.FirstLineStartIndent != 0)
            para.Format.FirstLineIndent = MdUnit.FromPoint(chars.FirstLineStartIndentPt);

        if (chars.LineSpacing > 0 && chars.FontSize > 0)
        {
            para.Format.LineSpacingRule = LineSpacingRule.AtLeast;
            para.Format.LineSpacing = MdUnit.FromPoint(chars.LineSpacingPt);
        }

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

    /// Add text to a paragraph, converting \n to MigraDoc line breaks.
    /// Add text to a paragraph. For verbatim text (containing \n),
    /// convert newlines to MigraDoc line breaks and preserve spaces
    /// as non-breaking spaces so MigraDoc doesn't collapse them.
    private static void AddTextWithLineBreaks(Paragraph para, string text, PdfCharacteristics chars)
    {
        bool verbatim = text.Contains('\n');
        if (!verbatim)
        {
            var ft = para.AddFormattedText(text);
            ApplyFont(ft.Font, chars);
            return;
        }

        // Replace regular spaces with non-breaking spaces to preserve indentation
        text = text.Replace(' ', '\u00A0');

        var lines = text.Split('\n');
        for (int i = 0; i < lines.Length; i++)
        {
            if (lines[i].Length > 0)
            {
                var ft = para.AddFormattedText(lines[i]);
                ApplyFont(ft.Font, chars);
            }
            if (i < lines.Length - 1)
                para.AddLineBreak();
        }
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
