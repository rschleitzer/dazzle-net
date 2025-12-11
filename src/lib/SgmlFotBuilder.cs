// Based on TransformFOTBuilder from OpenJade
// Copyright (c) 1997 James Clark

namespace Dazzle;

using OpenSP;
using OpenJade.Style;
using OpenJade.Grove;
using System.Text;
using Char = System.UInt32;
using Boolean = System.Boolean;

// SGML/XML FOT Builder - produces transformed SGML/XML output
public class SgmlFotBuilder : FOTBuilder
{
    private CmdLineApp? app_;
    private OutputCharStream? os_;
    private bool xml_;
    private System.Collections.Generic.List<StringC>? options_;
    private System.Collections.Generic.List<StringC> openElements_;
    private ReState state_;
    private bool preserveSdata_;
    private System.Collections.Generic.List<bool> preserveSdataStack_;
    private System.Collections.Generic.Stack<OpenFile> openFileStack_ = new();

    // Tracks an open output file for entity flow objects
    private class OpenFile
    {
        public StringC systemId = new StringC();
        public OutputCharStream? saveOs;
        public FileOutputByteStream? fileByteStream;
        public OutputCharStream? os;
    }

    public enum ReState
    {
        stateMiddle,
        stateStartOfElement,
        statePendingRe
    }

    // Document type NIC for doctype declaration
    public class DocumentTypeNIC
    {
        public StringC name = new StringC();
        public StringC publicId = new StringC();
        public StringC systemId = new StringC();
    }

    // Element NIC for element construction
    public class ElementNIC
    {
        public StringC gi = new StringC();
        public System.Collections.Generic.List<StringC> attributes = new System.Collections.Generic.List<StringC>();
    }

    // Constructor
    public SgmlFotBuilder(CmdLineApp app, bool xml, System.Collections.Generic.List<StringC> options)
    {
        app_ = app;
        xml_ = xml;
        options_ = options;
        openElements_ = new System.Collections.Generic.List<StringC>();
        state_ = ReState.stateMiddle;
        preserveSdata_ = false;
        preserveSdataStack_ = new System.Collections.Generic.List<bool>();

        // Create output stream
        os_ = app_.makeStdOut();

        // Output XML declaration if in XML mode
        if (xml_)
            output("<?xml version=\"1.0\"?>\n");
    }

    private void output(string s)
    {
        if (os_ == null) return;
        foreach (char c in s)
            os_.put((Char)c);
    }

    private void output(StringC s)
    {
        if (os_ == null) return;
        for (nuint i = 0; i < s.size(); i++)
            os_.put(s[i]);
    }

    private void flushPendingRe()
    {
        if (state_ == ReState.statePendingRe)
        {
            output("\r");
            state_ = ReState.stateMiddle;
        }
    }

    private void flushPendingReCharRef()
    {
        if (state_ == ReState.statePendingRe)
        {
            output("&#13;");
            state_ = ReState.stateMiddle;
        }
    }

    // Output attributes
    private void outputAttributes(System.Collections.Generic.List<StringC> atts)
    {
        for (int i = 0; i + 1 < atts.Count; i += 2)
        {
            output(" ");
            output(atts[i]);
            output("=\"");
            outputAttributeValue(atts[i + 1]);
            output("\"");
        }
    }

    private void outputAttributeValue(StringC value)
    {
        for (nuint i = 0; i < value.size(); i++)
        {
            Char c = value[i];
            switch (c)
            {
                case '<': output("&lt;"); break;
                case '>': output("&gt;"); break;
                case '&': output("&amp;"); break;
                case '"': output("&quot;"); break;
                default:
                    if (os_ != null)
                        os_.put(c);
                    break;
            }
        }
    }

    public override void start()
    {
        preserveSdataStack_.Add(preserveSdata_);
    }

    public override void end()
    {
        if (preserveSdataStack_.Count > 0)
        {
            preserveSdata_ = preserveSdataStack_[preserveSdataStack_.Count - 1];
            preserveSdataStack_.RemoveAt(preserveSdataStack_.Count - 1);
        }
    }

    public override void characters(Char[] data, nuint size)
    {
        if (state_ == ReState.stateStartOfElement)
            state_ = ReState.stateMiddle;
        else
            flushPendingRe();

        for (nuint i = 0; i < size; i++)
        {
            Char c = data[i];
            switch (c)
            {
                case '<': output("&lt;"); break;
                case '>': output("&gt;"); break;
                case '&': output("&amp;"); break;
                case '\r':
                    if (state_ == ReState.stateMiddle)
                        state_ = ReState.statePendingRe;
                    else
                        flushPendingReCharRef();
                    break;
                default:
                    if (os_ != null)
                        os_.put(c);
                    break;
            }
        }
    }

    public override void charactersFromNode(NodePtr nd, Char[] data, nuint size)
    {
        characters(data, size);
    }

    // Start an element
    public void startElement(ElementNIC nic)
    {
        flushPendingRe();
        output("<");
        output(nic.gi);
        outputAttributes(nic.attributes);
        output(">");
        openElements_.Add(new StringC(nic.gi));
        state_ = ReState.stateStartOfElement;
    }

    // End an element
    public void endElement()
    {
        if (openElements_.Count > 0)
        {
            StringC gi = openElements_[openElements_.Count - 1];
            openElements_.RemoveAt(openElements_.Count - 1);

            if (state_ == ReState.stateStartOfElement)
                state_ = ReState.stateMiddle;
            else
                flushPendingRe();

            output("</");
            output(gi);
            output(">");
            state_ = ReState.statePendingRe;
        }
    }

    // Empty element
    public void emptyElement(ElementNIC nic)
    {
        flushPendingRe();
        output("<");
        output(nic.gi);
        outputAttributes(nic.attributes);
        if (xml_)
            output("/>");
        else
            output(">");
        state_ = ReState.statePendingRe;
    }

    // Processing instruction
    public void processingInstruction(StringC data)
    {
        flushPendingRe();
        output("<?");
        output(data);
        if (xml_)
            output("?>");
        else
            output(">");
        state_ = ReState.statePendingRe;
    }

    // Document type declaration
    public void documentType(DocumentTypeNIC nic)
    {
        flushPendingRe();
        output("<!DOCTYPE ");
        output(nic.name);
        if (nic.publicId.size() > 0)
        {
            output(" PUBLIC \"");
            output(nic.publicId);
            output("\"");
            if (nic.systemId.size() > 0)
            {
                output(" \"");
                output(nic.systemId);
                output("\"");
            }
        }
        else if (nic.systemId.size() > 0)
        {
            output(" SYSTEM \"");
            output(nic.systemId);
            output("\"");
        }
        output(">\n");
    }

    // Entity reference
    public void entityRef(StringC name)
    {
        flushPendingRe();
        output("&");
        output(name);
        output(";");
    }

    // Start external entity - opens a new output file
    public void startEntity(StringC systemId)
    {
        flushPendingRe();
        var ofp = new OpenFile();
        ofp.systemId = new StringC(systemId);
        ofp.saveOs = os_;

        // Convert StringC to filename string
        string filename = systemId.ToString();
        if (!string.IsNullOrEmpty(filename))
        {
            ofp.fileByteStream = new FileOutputByteStream();
            if (ofp.fileByteStream.open(filename))
            {
                // Use RecordOutputCharStream to handle line endings properly
                ofp.os = new RecordOutputCharStream(
                    new EncodeOutputCharStream(ofp.fileByteStream, app_!.outputCodingSystem()!));
                os_ = ofp.os;
            }
            else
            {
                // Could not open file
                Console.Error.WriteLine($"ERROR: could not open file '{filename}'");
                ofp.fileByteStream = null;
                ofp.os = null;
            }
        }
        openFileStack_.Push(ofp);
    }

    // End external entity - closes the current file and restores previous output
    public void endEntity()
    {
        flushPendingRe();
        if (openFileStack_.Count > 0)
        {
            var of = openFileStack_.Pop();
            if (of.os != null)
            {
                of.os.flush();
            }
            if (of.fileByteStream != null)
            {
                of.fileByteStream.close();
            }
            os_ = of.saveOs;
        }
    }

    public override void formattingInstruction(StringC s)
    {
        flushPendingRe();
        output(s);
    }

    // Extension flow object handling
    public override void extension(ExtensionFlowObj fo, NodePtr currentNode)
    {
        if (fo is SgmlExtensionFlowObj tfo)
            tfo.atomic(this, currentNode);
    }

    public override void startExtension(ExtensionFlowObj fo, NodePtr currentNode, System.Collections.Generic.List<FOTBuilder?> fotbs)
    {
        if (fo is SgmlCompoundExtensionFlowObj tfo)
            tfo.start(this, currentNode);
    }

    public override void endExtension(ExtensionFlowObj fo)
    {
        if (fo is SgmlCompoundExtensionFlowObj tfo)
            tfo.end(this);
    }

    // Setters
    public void setPreserveSdata(bool flag)
    {
        preserveSdata_ = flag;
    }

    // Additional flow object support
    public override void startParagraph(ParagraphNIC nic)
    {
        start();
    }

    public override void endParagraph()
    {
        output("\n");
        end();
    }

    public override void startDisplayGroup(DisplayGroupNIC nic)
    {
        start();
    }

    public override void endDisplayGroup()
    {
        end();
    }

    public override void startScroll()
    {
        start();
    }

    public override void endScroll()
    {
        end();
    }

    public override void startSequence()
    {
        start();
    }

    public override void endSequence()
    {
        end();
    }

    // Static extensions array for registration
    public static FOTBuilder.ExtensionTableEntry[] GetExtensions()
    {
        return new FOTBuilder.ExtensionTableEntry[]
        {
            new FOTBuilder.ExtensionTableEntry
            {
                pubid = "UNREGISTERED::James Clark//Flow Object Class::entity",
                flowObj = new EntityFlowObj()
            },
            new FOTBuilder.ExtensionTableEntry
            {
                pubid = "UNREGISTERED::James Clark//Flow Object Class::entity-ref",
                flowObj = new EntityRefFlowObj()
            },
            new FOTBuilder.ExtensionTableEntry
            {
                pubid = "UNREGISTERED::James Clark//Flow Object Class::element",
                flowObj = new ElementFlowObj()
            },
            new FOTBuilder.ExtensionTableEntry
            {
                pubid = "UNREGISTERED::James Clark//Flow Object Class::empty-element",
                flowObj = new EmptyElementFlowObj()
            },
            new FOTBuilder.ExtensionTableEntry
            {
                pubid = "UNREGISTERED::James Clark//Flow Object Class::document-type",
                flowObj = new DocumentTypeFlowObj()
            },
            new FOTBuilder.ExtensionTableEntry
            {
                pubid = "UNREGISTERED::James Clark//Flow Object Class::processing-instruction",
                flowObj = new ProcessingInstructionFlowObj()
            },
            new FOTBuilder.ExtensionTableEntry
            {
                pubid = "UNREGISTERED::James Clark//Flow Object Class::formatting-instruction",
                flowObj = new FormattingInstructionFlowObj()
            }
        };
    }
}

// Base class for atomic SGML extension flow objects
public class SgmlExtensionFlowObj : FOTBuilder.ExtensionFlowObj
{
    public virtual void atomic(SgmlFotBuilder fotb, NodePtr nd) { }
}

// Base class for compound SGML extension flow objects
public class SgmlCompoundExtensionFlowObj : FOTBuilder.CompoundExtensionFlowObj
{
    public virtual void start(SgmlFotBuilder fotb, NodePtr nd) { }
    public virtual void end(SgmlFotBuilder fotb) { }
}

// Entity flow object - creates a new output file
public class EntityFlowObj : SgmlCompoundExtensionFlowObj
{
    private StringC systemId_ = new StringC();

    public override void start(SgmlFotBuilder fotb, NodePtr nd)
    {
        fotb.startEntity(systemId_);
    }

    public override void end(SgmlFotBuilder fotb)
    {
        fotb.endEntity();
    }

    public override bool hasNIC(StringC name)
    {
        return name.ToString() == "system-id";
    }

    public override void setNIC(StringC name, IExtensionFlowObjValue value)
    {
        if (name.ToString() == "system-id")
            value.convertString(out systemId_);
    }

    public override FOTBuilder.ExtensionFlowObj copy()
    {
        var c = new EntityFlowObj();
        c.systemId_ = new StringC(systemId_);
        return c;
    }
}

// Entity reference flow object
public class EntityRefFlowObj : SgmlExtensionFlowObj
{
    private StringC name_ = new StringC();

    public override void atomic(SgmlFotBuilder fotb, NodePtr nd)
    {
        fotb.entityRef(name_);
    }

    public override bool hasNIC(StringC name)
    {
        return name.ToString() == "name";
    }

    public override void setNIC(StringC name, IExtensionFlowObjValue value)
    {
        if (name.ToString() == "name")
            value.convertString(out name_);
    }

    public override FOTBuilder.ExtensionFlowObj copy()
    {
        var c = new EntityRefFlowObj();
        c.name_ = new StringC(name_);
        return c;
    }
}

// Element flow object
public class ElementFlowObj : SgmlCompoundExtensionFlowObj
{
    private SgmlFotBuilder.ElementNIC nic_ = new SgmlFotBuilder.ElementNIC();

    public override void start(SgmlFotBuilder fotb, NodePtr nd)
    {
        fotb.startElement(nic_);
    }

    public override void end(SgmlFotBuilder fotb)
    {
        fotb.endElement();
    }

    public override bool hasNIC(StringC name)
    {
        string n = name.ToString();
        return n == "gi" || n == "attributes";
    }

    public override void setNIC(StringC name, IExtensionFlowObjValue value)
    {
        string n = name.ToString();
        if (n == "gi")
            value.convertString(out nic_.gi);
        // attributes would need special handling
    }

    public override FOTBuilder.ExtensionFlowObj copy()
    {
        var c = new ElementFlowObj();
        c.nic_.gi = new StringC(nic_.gi);
        return c;
    }
}

// Empty element flow object
public class EmptyElementFlowObj : SgmlExtensionFlowObj
{
    private SgmlFotBuilder.ElementNIC nic_ = new SgmlFotBuilder.ElementNIC();

    public override void atomic(SgmlFotBuilder fotb, NodePtr nd)
    {
        fotb.emptyElement(nic_);
    }

    public override bool hasNIC(StringC name)
    {
        string n = name.ToString();
        return n == "gi" || n == "attributes";
    }

    public override void setNIC(StringC name, IExtensionFlowObjValue value)
    {
        string n = name.ToString();
        if (n == "gi")
            value.convertString(out nic_.gi);
    }

    public override FOTBuilder.ExtensionFlowObj copy()
    {
        var c = new EmptyElementFlowObj();
        c.nic_.gi = new StringC(nic_.gi);
        return c;
    }
}

// Document type flow object
public class DocumentTypeFlowObj : SgmlExtensionFlowObj
{
    private SgmlFotBuilder.DocumentTypeNIC nic_ = new SgmlFotBuilder.DocumentTypeNIC();

    public override void atomic(SgmlFotBuilder fotb, NodePtr nd)
    {
        fotb.documentType(nic_);
    }

    public override bool hasNIC(StringC name)
    {
        string n = name.ToString();
        return n == "name" || n == "system-id" || n == "public-id";
    }

    public override void setNIC(StringC name, IExtensionFlowObjValue value)
    {
        string n = name.ToString();
        if (n == "name")
            value.convertString(out nic_.name);
        else if (n == "system-id")
            value.convertString(out nic_.systemId);
        else if (n == "public-id")
            value.convertString(out nic_.publicId);
    }

    public override FOTBuilder.ExtensionFlowObj copy()
    {
        var c = new DocumentTypeFlowObj();
        c.nic_.name = new StringC(nic_.name);
        c.nic_.systemId = new StringC(nic_.systemId);
        c.nic_.publicId = new StringC(nic_.publicId);
        return c;
    }
}

// Processing instruction flow object
public class ProcessingInstructionFlowObj : SgmlExtensionFlowObj
{
    private StringC data_ = new StringC();

    public override void atomic(SgmlFotBuilder fotb, NodePtr nd)
    {
        fotb.processingInstruction(data_);
    }

    public override bool hasNIC(StringC name)
    {
        return name.ToString() == "data";
    }

    public override void setNIC(StringC name, IExtensionFlowObjValue value)
    {
        if (name.ToString() == "data")
            value.convertString(out data_);
    }

    public override FOTBuilder.ExtensionFlowObj copy()
    {
        var c = new ProcessingInstructionFlowObj();
        c.data_ = new StringC(data_);
        return c;
    }
}

// Formatting instruction flow object
public class FormattingInstructionFlowObj : SgmlExtensionFlowObj
{
    private StringC data_ = new StringC();

    public override void atomic(SgmlFotBuilder fotb, NodePtr nd)
    {
        fotb.formattingInstruction(data_);
    }

    public override bool hasNIC(StringC name)
    {
        return name.ToString() == "data";
    }

    public override void setNIC(StringC name, IExtensionFlowObjValue value)
    {
        if (name.ToString() == "data")
            value.convertString(out data_);
    }

    public override FOTBuilder.ExtensionFlowObj copy()
    {
        var c = new FormattingInstructionFlowObj();
        c.data_ = new StringC(data_);
        return c;
    }
}
