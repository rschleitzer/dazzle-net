# dazzle-net

A DSSSL processor for .NET. It runs DSSSL stylesheets (ISO/IEC 10179) over
SGML and XML documents with the [OpenJade](https://www.nuget.org/packages/OpenJade)
style engine — the C# port of James Clark's original — and adds two things
the original does not have: a flow object class for writing whole directory
trees, and a PDF backend.

What the packages add to DSSSL, where they read the standard where it leaves room, and which of its
optional parts they have, is described in [`docs/dazzle.xml`](docs/dazzle.xml) (DocBook).

## Packages

| Package | What it is |
|---|---|
| [Dazzle.Cli](https://www.nuget.org/packages/Dazzle.Cli) | the command-line tool `dazzle` |
| [Dazzle.Pdf](https://www.nuget.org/packages/Dazzle.Pdf) | the PDF backend as a library, with a small API |
| [Dazzle](https://www.nuget.org/packages/Dazzle) | the transformation backend with the `directory` flow object class |

All target .NET 10.

## The tool

```
dotnet tool install -g Dazzle.Cli

dazzle -t pdf -o book.pdf -d stylesheet.dsl document.xml
```

```
Usage: dazzle [options] DSSSL-spec document
  -t type    Output type (sgml, xml, pdf)
  -o file    Output file
  -d spec    DSSSL specification
  -V var     Define variable
  -c catalog Use SGML catalog
```

The DTDs of the DSSSL architecture and a catalog for them come with the tool.

## Generating files: the `directory` flow object class

OpenJade's transformation backend (`-t sgml`, `-t xml`) writes text files
through the `entity` flow object class. dazzle adds `directory`, which
creates a directory and makes the entities inside it relative to it — so a
stylesheet can lay out a whole tree of generated files, which is what makes
DSSSL usable as a code generator.

```scheme
(declare-flow-object-class entity
  "UNREGISTERED::James Clark//Flow Object Class::entity")
(declare-flow-object-class directory
  "UNREGISTERED::Dazzle//Flow Object Class::directory")

(make directory
  path: "generated"
  (make entity
    system-id: "Parser.cs"
    (literal "// ...")))
```

## PDF

The PDF backend sets the pages itself and writes them with
[PDFsharp](https://www.nuget.org/packages/PDFsharp), the way the PDF backend
of the Scaly port of OpenJade does: blocks flow down the page, paragraphs are
broken into lines, tables are set row by row, and the flow object tree is laid
out twice so that `current-node-page-number-sosofo` knows its page. It sets
simple page sequences with headers and footers, page sequences and column sets
of the full page model (below), paragraphs and
`paragraph-break`, display groups, boxes, line fields, leaders, `score`, rules,
external graphics and tables (spanned columns and rows, borders of the table
and its cells, header rows again on every page, tables in cells) — enough to set
a book with the DocBook print stylesheets.

An error in the stylesheet fails the run (`DssslProcessingException`, with the
place in the stylesheet), and so does what the backend cannot set although the
language allows it.

The characteristics mean what ISO/IEC 10179 says, as the RTF backend of Jade has
them:

- `line-spacing` is the distance of the lines: exactly that with `min-leading: #f`
  (the initial value), at least that and the height of the line's text plus
  `min-leading` otherwise; below the placement path of a line is the descent of
  its text, or `min-post-line-spacing` where that is more, and the rest of the
  line spacing is above; with `min-pre-line-spacing` alone, that much is above
  and the rest below;
- a line area is `min-pre-line-spacing` before and `min-post-line-spacing` after
  its placement path (12.6.6.1): at the end of an area the last line needs only
  its extent down to the descent of its text, and the rest of
  `min-post-line-spacing` may lie past the end, as word processors do with the
  extra of a multiple line spacing (since 0.7.0);
- display space collapses to the larger of the space after one block and the
  space before the next, and `conditional?` space is dropped at the start of an
  area — the top of a page, a table cell;
- the cell margins are the distance between the edges of a cell and its content;
- `score` draws its line after (`'after`), through or before the text, or at a
  length below it, over the spaces only with `score-spaces?: #t`;
- `input-whitespace-treatment: 'preserve` keeps every space, and where `lines:`
  wraps a line still breaks at them; the spaces where a line is broken for its
  width stay at the end of that line, as XSL's `suppress-at-line-break` has it,
  and do not indent the next; a tab goes to the next multiple of 36pt;
- a line breaks at a space, after a soft hyphen (which then shows) and after a
  hyphen between letters;
- a cell is placed in the column after the previous cell of its row unless its
  `column-number` says otherwise, also under a cell spanning rows;
- an external graphic takes its `scale`: `'max-uniform` (the initial value) the
  largest uniform scale within `max-width` and `max-height`, `'max` each direction
  on its own, numbers as they are; a graphic may overflow its line;
- a line field in a `score` is scored over its whole width;
- the backgrounds of table cells lie under the contents and borders of every
  table on the page;
- a line breaks only once it holds something: a first word wider than the line
  overflows it, and a hyphen between letters is a break also where the letters
  are in different runs;
- `break-before: 'column` and `break-after: 'column` go on in the next column of
  a column set (below), and outside one on a new page;
- a table row is as high as its cells ask: the style of a `table-row` reaches the
  backend with its cells, and a table in a cell does not take on the height of
  the row around it.

One deviation: the initial `input-whitespace-treatment` is `'collapse`, not
`'preserve`.

Header and footer parts are set as blocks across the width of the page body,
`center-` and `right-` parts aligned to the middle and the end; they stand at
`header-margin` from the top and `footer-margin` from the bottom of the page,
and without these in the middle of the margin.

The fonts come from PDFsharp's font resolver: the one the program sets in
`GlobalFontSettings.FontResolver`, else the system's fonts.

From a program:

```csharp
using Dazzle.Pdf;

DssslProcessor.GeneratePdf("document.xml", "stylesheet.dsl", "document.pdf");

byte[] pdf = DssslProcessor.GeneratePdfBytes("document.xml", "stylesheet.dsl");
```

A third overload writes to a `Stream`. All take an optional path to an SGML
catalog; without one the bundled catalog is used.

`GeneratePdfWithLocations` also says where the elements that have an ID (an
attribute declared `ID` in the document's DTD) begin: the page, counted from 1,
and the top left of the element's first line, or of the content of its table
cell, in points from the top left of the page. A program that draws on the PDF
afterwards can place things relative to an element, the way word processors
anchor a frame to a paragraph:

```csharp
PdfResult result = DssslProcessor.GeneratePdfWithLocations("document.xml", "stylesheet.dsl");
PdfElementLocation where = result.Locations["intro"];   // where.Page, where.X, where.Y
```

The locations come from the marks the backend sets for every element processed
in the initial processing mode, the same marks `current-node-page-number-sosofo`
reads; an element processed only in a named mode has none. The top of a first
line is below the element's `space-before`.

### The full page model

Since 0.6.0, on OpenJade 1.4.0, the backend has a part of the full page model of
ISO/IEC 10179 (12.6.4, 12.6.5): the features `page` and `multi-column`,
restricted, and not `nested-column-set`. A page model has the size of the page
and one region filled top to bottom (its origin at the lower left, y going up);
a column-set model has one column subset with its columns and `balance?`.
`page-sequence` takes `initial-page-models` and `repeat-page-models`: the k-th
page takes the k-th initial model, the pages after them go round the repeat
models. `column-set-sequence` takes `column-set-model` and the display
characteristics; its content flows down a column into the next, from the last
into the first column of the next page, and the set ends below its longest
column. Everything else of the page model is an error where it is written.

```scheme
(define-page-model first
  (width 595pt) (height 842pt)
  (region (x-origin 70pt) (y-origin 56pt) (width 469pt) (height 600pt)))
(define-page-model other
  (width 595pt) (height 842pt)
  (region (x-origin 70pt) (y-origin 56pt) (width 469pt) (height 730pt)))
(define-column-set-model two
  (column-subset (column (x-origin 0pt) (width 220pt))
                 (column (x-origin 249pt) (width 220pt))
                 (balance? #t)))

(root (make page-sequence initial-page-models: (list first)
                          repeat-page-models: (list other)
        (process-children)))
(element section (make column-set-sequence column-set-model: two
                   (process-children)))
```

A column-set-sequence in a table cell, a header or a footer, and `span` other
than 1 in a column set, are errors.

### Extension characteristics

DSSSL has no height for a table row. The backend takes it as an extension
characteristic, which a stylesheet declares and then specifies on a
`table-row` (or anything around it):

```scheme
(declare-characteristic min-row-height
  "UNREGISTERED::Dazzle//Characteristic::min-row-height" 0pt)
(declare-characteristic row-height
  "UNREGISTERED::Dazzle//Characteristic::row-height" 0pt)

(make table-row min-row-height: 20pt ...)
```

| Characteristic | Value | Meaning |
|---|---|---|
| `min-row-height` | length | a `table-row` is at least this high |
| `row-height` | length | a `table-row` is exactly this high |
| `column-set` | string | the columns a `display-group` sets its content in (deprecated: `column-set-sequence`) |
| `balance-columns?` | boolean | the columns of the group's last page are made equally long (deprecated: `balance?`) |
| `first-page-top-margin` | length | the top margin of the first page of a `simple-page-sequence` (0pt: as the others; deprecated: `page-sequence`) |
| `first-page-bottom-margin` | length | its bottom margin (0pt: as the others; deprecated: `page-sequence`) |

The last four are deprecated since 0.6.0 and stay for one version: the full page
model above does the same in the terms of the standard.

DSSSL sets columns in the full page model, which the backend did not have
before 0.6.0. A `display-group` that specifies `column-set` sets its content in
columns instead: the string gives, in points, the start of the first column from the
start of the area and its width, then for every further column the gap before
it and its width (`"0 240 14 241"`; one column with an offset narrows the area).
The set begins where the group begins and the space before it, content flows down
a column into the next, from the last into the first column of the next page,
and the set ends below its longest column. A paragraph's indents are measured
from its column. At the top of a column `conditional?` space is dropped as at the
top of a page. With `balance-columns?: #t` the columns of the set's last page end
at the smallest height at which the last column is not longer than the others
(the first ones are rather the longer); the document is set again for that, once
per set and height tried. A set with a column break on its last page is not
balanced. A column set in a column set, a table cell, a header or a footer is an
ordinary group.

```scheme
(declare-characteristic column-set
  "UNREGISTERED::Dazzle//Characteristic::column-set" "")
(declare-characteristic balance-columns?
  "UNREGISTERED::Dazzle//Characteristic::balance-columns?" #f)

(make display-group column-set: "0 240 14 241" balance-columns?: #t ...)
```

## License

MIT, see `LICENSE`. OpenJade and OpenSP carry the license of the original.
