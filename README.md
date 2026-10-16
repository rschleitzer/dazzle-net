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

The PDF backend renders the flow object tree with
[PDFsharp/MigraDoc](https://www.nuget.org/packages/PDFsharp-MigraDoc). It is
rudimentary: page sequences with headers and footers, paragraphs, display
groups, leaders and tables — enough to set a book with the DocBook print
stylesheets, and not more.

From a program:

```csharp
using Dazzle.Pdf;

DssslProcessor.GeneratePdf("document.xml", "stylesheet.dsl", "document.pdf");

byte[] pdf = DssslProcessor.GeneratePdfBytes("document.xml", "stylesheet.dsl");
```

A third overload writes to a `Stream`. All take an optional path to an SGML
catalog; without one the bundled catalog is used.

Beyond the basics the backend sets `score` with `type: 'after` as an underline
(`score-spaces?: #f` underlines the words only; MigraDoc has no overline or
strike-through), negative indents, a tab character in text with
`input-whitespace-treatment: 'preserve`, preserved runs of spaces that still
break where `lines:` wraps, leaders in table cells, and cell margins the way
the RTF backend of Jade adds them to the paragraphs of the cell. A table in a
table cell, which MigraDoc cannot nest, is set as its cells one after another.

### Extension characteristics

What DSSSL has no characteristic for, the backend takes as extension
characteristics. A stylesheet declares the ones it uses and then specifies
them like any inherited characteristic; one that does not declare them gets
the behaviour of earlier versions.

```scheme
(declare-characteristic tab-stops
  "UNREGISTERED::Dazzle//Characteristic::tab-stops" "")
(declare-characteristic line-spacing-exact?
  "UNREGISTERED::Dazzle//Characteristic::line-spacing-exact?" #f)
(declare-characteristic min-row-height
  "UNREGISTERED::Dazzle//Characteristic::min-row-height" 0pt)
(declare-characteristic row-height
  "UNREGISTERED::Dazzle//Characteristic::row-height" 0pt)
(declare-characteristic exact-cell-margins?
  "UNREGISTERED::Dazzle//Characteristic::exact-cell-margins?" #f)

(make paragraph
  tab-stops: "100pt 280pt:center:dots 15cm:end:lines"
  ...)
```

| Characteristic | Value | Meaning |
|---|---|---|
| `tab-stops` | string | tab stops of a paragraph, separated by white space, each `position[:alignment[:leader]]`; the position with the unit `pt`, `mm`, `cm` or `in`, the alignment `start`, `end`, `center` or `decimal`, the leader `none`, `dots`, `dashes`, `lines`, `heavy` or `middle-dots` |
| `line-spacing-exact?` | boolean | `line-spacing` is the exact distance of the lines, also below the size of the font; otherwise it is a minimum |
| `min-row-height` | length | a `table-row` is at least this high |
| `row-height` | length | a `table-row` is exactly this high |
| `exact-cell-margins?` | boolean | on a `table`: the cell margins are the whole distance between the edge of a cell and its content; otherwise MigraDoc pads every cell by 1.2 mm left and right in addition |

## License

MIT, see `LICENSE`. OpenJade and OpenSP carry the license of the original.
