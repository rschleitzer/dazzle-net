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

## License

MIT, see `LICENSE`. OpenJade and OpenSP carry the license of the original.
