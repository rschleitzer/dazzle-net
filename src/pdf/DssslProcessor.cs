namespace Dazzle.Pdf;

using OpenSP;
using OpenJade.Grove;
using OpenJade.SPGrove;
using OpenJade.Style;
using Char = System.UInt32;

/// <summary>
/// Drives the OpenJade DSSSL engine to produce PDF output.
/// No CLI app hierarchy, no environment variables — just .NET objects.
/// </summary>
public static class DssslProcessor
{
    /// <summary>
    /// Process an XML document with a DSSSL stylesheet and write PDF to a stream.
    /// </summary>
    /// <param name="xmlPath">Path to the XML document</param>
    /// <param name="dssslPath">Path to the DSSSL stylesheet (.dsl)</param>
    /// <param name="output">Stream to write the PDF to</param>
    /// <param name="catalogPath">Optional path to an SGML catalog file (defaults to bundled catalog)</param>
    public static void GeneratePdf(string xmlPath, string dssslPath, Stream output,
        string? catalogPath = null)
    {
        catalogPath ??= BundledCatalogPath();

        // Coding system: XMLCodingSystem auto-detects UTF-8/UTF-16 from XML declaration.
        // No SP_CHARSET_FIXED or SP_ENCODING env vars needed.
        var codingSystemKit = CodingSystemKit.make(null)!;
        var inputCodingSystemKit = new ConstPtr<InputCodingSystemKit>(codingSystemKit);
        var systemCharset = codingSystemKit.systemCharset();
        var xmlCodingSystem = new XMLCodingSystem(inputCodingSystemKit.pointer()!);

        // Entity manager with file system storage, using XML coding system for UTF-8 support
        var storageMgr = new PosixStorageManager("OSFILE", systemCharset, xmlCodingSystem, 5);
        var entityManager = ExtendEntityManager.make(storageMgr, xmlCodingSystem, inputCodingSystemKit, true);

        // SGML catalog (resolves PUBLIC identifiers for DSSSL DTD, document DTD, etc.)
        if (File.Exists(catalogPath))
        {
            var catalogs = new Vector<StringC>();
            catalogs.push_back(new StringC(catalogPath));
            var catalogManager = SOCatalogManager.make(catalogs, 0, systemCharset, systemCharset, true);
            entityManager.setCatalogManager(catalogManager);
        }

        // Messenger: captures errors with file:line:column location
        var messageStream = new StrOutputCharStream();
        var messenger = new MessageReporter(messageStream);

        // Merge xml.dcl + document path into one sysid (xml.dcl tells OpenSP this is XML)
        var xmlDclPath = Path.Combine(
            Path.GetDirectoryName(catalogPath!) ?? "", "xml.dcl");
        var docSysid = new StringC();
        var sysids = new Vector<StringC>(2);
        sysids[0] = new StringC(xmlDclPath);
        sysids[1] = new StringC(xmlPath);
        entityManager.mergeSystemIds(sysids, false, systemCharset, messenger, docSysid);

        // Parse the XML document into a grove
        var docParams = new SgmlParser.Params();
        docParams.sysid = docSysid;
        docParams.entityManager = new Ptr<EntityManager>(entityManager);
        docParams.options = new ParserOptions();

        var docParser = new SgmlParser(docParams);
        docParser.allLinkTypesActivated();

        var rootNode = new NodePtr();
        var groveBuilder = GroveBuilder.make(0, messenger, null, false, ref rootNode);
        docParser.parseAll(groveBuilder, 0);
        if (groveBuilder is GroveBuilderMessageEventHandler gbeh)
            gbeh.markComplete();

        // Check for XML parsing errors
        ThrowOnErrors(messageStream, "XML document parsing failed");

        // Parse the DSSSL specification
        var specParams = new SgmlParser.Params();
        specParams.sysid = new StringC(dssslPath);
        specParams.entityManager = new Ptr<EntityManager>(entityManager);
        specParams.options = new ParserOptions();

        var specParser = new SgmlParser(specParams);
        specParser.allLinkTypesActivated();

        // Create FOT builder and style engine
        var extensions = PdfFOTBuilder.GetExtensions();
        var groveManager = new SimpleGroveManager(entityManager, messenger);

        using var styleEngine = new StyleEngine(
            messenger, groveManager, 72000, false, false, false, extensions);

        styleEngine.parseSpec(specParser, systemCharset, new StringC(), messenger);
        ThrowOnErrors(messageStream, "DSSSL stylesheet parsing failed");

        // Process document → FOT builder → PDF
        var fotBuilder = new PdfFOTBuilder(messenger, "");
        styleEngine.process(rootNode, fotBuilder);
        ThrowOnErrors(messageStream, "DSSSL processing failed");
        fotBuilder.Finish(output);
    }

    /// <summary>
    /// Process an XML document with a DSSSL stylesheet and write PDF to a file.
    /// </summary>
    public static void GeneratePdf(string xmlPath, string dssslPath, string outputPath,
        string? catalogPath = null)
    {
        using var stream = File.Create(outputPath);
        GeneratePdf(xmlPath, dssslPath, stream, catalogPath);
    }

    /// <summary>
    /// Process an XML document with a DSSSL stylesheet and return PDF as byte array.
    /// </summary>
    public static byte[] GeneratePdfBytes(string xmlPath, string dssslPath,
        string? catalogPath = null)
    {
        using var ms = new MemoryStream();
        GeneratePdf(xmlPath, dssslPath, ms, catalogPath);
        return ms.ToArray();
    }

    /// <summary>
    /// Returns the path to the SGML catalog bundled with Dazzle.Pdf.
    /// </summary>
    public static string BundledCatalogPath()
    {
        var assemblyDir = Path.GetDirectoryName(
            System.Reflection.Assembly.GetExecutingAssembly().Location) ?? "";
        return Path.Combine(assemblyDir, "sgml", "catalog");
    }

    private static void ThrowOnErrors(StrOutputCharStream messageStream, string context)
    {
        var str = new StringC();
        messageStream.extractString(str);
        var messages = str.ToString().Trim();
        if (!string.IsNullOrEmpty(messages))
            throw new DssslProcessingException($"{context}:\n{messages}");
    }

    // Minimal GroveManager for loading additional groves (cross-references)
    private class SimpleGroveManager : GroveManager
    {
        private readonly ExtendEntityManager entityManager_;
        private readonly Messenger messenger_;
        private readonly Dictionary<string, NodePtr> groveTable_ = new();

        public SimpleGroveManager(ExtendEntityManager entityManager, Messenger messenger)
        {
            entityManager_ = entityManager;
            messenger_ = messenger;
        }

        public override bool load(StringC sysid, System.Collections.Generic.List<StringC> active,
            NodePtr parent, ref NodePtr rootNode, System.Collections.Generic.List<StringC> architecture)
        {
            var key = sysid.ToString();
            if (groveTable_.TryGetValue(key, out var existing))
            {
                rootNode = existing;
                return true;
            }

            var parms = new SgmlParser.Params();
            parms.sysid = sysid;
            parms.entityManager = new Ptr<EntityManager>(entityManager_);
            parms.options = new ParserOptions();

            var parser = new SgmlParser(parms);
            foreach (var linkType in active)
                parser.activateLinkType(linkType);
            parser.allLinkTypesActivated();

            var newRoot = new NodePtr();
            var builder = GroveBuilder.make(0, messenger_, null, false, ref newRoot);
            parser.parseAll(builder, 0);
            if (builder is GroveBuilderMessageEventHandler gbeh)
                gbeh.markComplete();

            groveTable_[key] = newRoot;
            rootNode = newRoot;
            return true;
        }

        public override bool readEntity(StringC name, out StringC content)
        {
            content = new StringC();
            return false;
        }

        public override void mapSysid(ref StringC sysid) { }
    }
}

/// <summary>
/// Exception thrown when DSSSL processing encounters errors (with file:line:column details).
/// </summary>
public class DssslProcessingException : Exception
{
    public DssslProcessingException(string message) : base(message) { }
}
