# dazzle-net

A DSSSL processor for .NET. It runs DSSSL stylesheets (ISO/IEC 10179) over
SGML and XML documents with the [OpenJade](https://www.nuget.org/packages/OpenJade)
style engine — the C# port of James Clark's original — and adds two things
the original does not have: a flow object class for writing whole directory
trees, and a PDF backend.

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
simple page sequences with headers and footers, paragraphs and
`paragraph-break`, display groups, boxes, line fields, leaders, `score`, rules,
external graphics and tables (spanned columns and rows, borders of the table
and its cells, header rows again on every page, tables in cells) — enough to set
a book with the DocBook print stylesheets.

The characteristics mean what ISO/IEC 10179 says, as the RTF backend of Jade has
them:

- `line-spacing` is the distance of the lines: exactly that with `min-leading: #f`
  (the initial value), at least that and the height of the line's text plus
  `min-leading` otherwise;
- display space collapses to the larger of the space after one block and the
  space before the next, and `conditional?` space is dropped at the start of an
  area — the top of a page, a table cell;
- the cell margins are the distance between the edges of a cell and its content;
- `score` draws its line after (`'after`), through or before the text, or at a
  length below it, over the spaces only with `score-spaces?: #t`;
- `input-whitespace-treatment: 'preserve` keeps every space, and where `lines:`
  wraps a line still breaks at them; a tab goes to the next multiple of 36pt;
- a line breaks at a space, after a soft hyphen (which then shows) and after a
  hyphen between letters;
- a cell is placed in the column after the previous cell of its row unless its
  `column-number` says otherwise, also under a cell spanning rows.

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

## License

MIT, see `LICENSE`. OpenJade and OpenSP carry the license of the original.
