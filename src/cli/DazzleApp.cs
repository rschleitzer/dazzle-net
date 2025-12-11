// Dazzle Application - DSSSL processor with custom backends

namespace Dazzle.Cli;

using OpenSP;
using OpenJade.Style;
using OpenJade.Jade;
using Dazzle;
using Char = System.UInt32;

public class DazzleApp : DssslApp
{
    public enum OutputType
    {
        sgmlType,
        xmlType
    }

    private static readonly string[] outputTypeNames = {
        "sgml",
        "xml"
    };

    private OutputType outputType_ = OutputType.sgmlType;
    private string outputFilename_ = "";
    private System.Collections.Generic.List<StringC> outputOptions_;

    public DazzleApp() : base(72000)
    {
        outputOptions_ = new System.Collections.Generic.List<StringC>();
        registerOption('t', "(sgml|xml)");
        registerOption('o', "output_file");
    }

    public override void processOption(char opt, string? arg)
    {
        if (arg == null) arg = "";
        switch (opt)
        {
            case 't':
                {
                    int dashPos = arg.IndexOf('-');
                    string typeName = dashPos >= 0 ? arg.Substring(0, dashPos) : arg;

                    bool found = false;
                    for (int i = 0; i < outputTypeNames.Length; i++)
                    {
                        if (typeName == outputTypeNames[i])
                        {
                            outputType_ = (OutputType)i;
                            found = true;
                            break;
                        }
                    }

                    if (!found)
                    {
                        Console.Error.WriteLine("Error: Unknown output type: " + arg);
                    }

                    // Parse sub-options
                    if (dashPos >= 0)
                    {
                        string subOpts = arg.Substring(dashPos);
                        System.Text.StringBuilder currentOpt = new System.Text.StringBuilder();
                        foreach (char c in subOpts)
                        {
                            if (c == '-')
                            {
                                if (currentOpt.Length > 0)
                                {
                                    StringC sc = new StringC();
                                    sc.assign(currentOpt.ToString());
                                    outputOptions_.Add(sc);
                                    currentOpt.Clear();
                                }
                            }
                            else
                            {
                                currentOpt.Append(c);
                            }
                        }
                        if (currentOpt.Length > 0)
                        {
                            StringC sc = new StringC();
                            sc.assign(currentOpt.ToString());
                            outputOptions_.Add(sc);
                        }
                    }
                }
                break;
            case 'o':
                if (arg.Length == 0)
                    Console.Error.WriteLine("Error: Empty output filename");
                else
                    outputFilename_ = arg;
                break;
            default:
                base.processOption(opt, arg);
                break;
        }
    }

    public override FOTBuilder? makeFOTBuilder(out FOTBuilder.ExtensionTableEntry[]? ext)
    {
        ext = SgmlFotBuilder.GetExtensions();

        return new SgmlFotBuilder(
            this,
            outputType_ == OutputType.xmlType,
            outputOptions_);
    }

    public string programName()
    {
        return "dazzle";
    }

    public void printUsage()
    {
        Console.WriteLine("Usage: dazzle [options] DSSSL-spec document");
        Console.WriteLine("Options:");
        Console.WriteLine("  -t type    Output type (sgml, xml)");
        Console.WriteLine("  -o file    Output file");
        Console.WriteLine("  -d spec    DSSSL specification");
        Console.WriteLine("  -V var     Define variable");
        Console.WriteLine("  -c catalog Use SGML catalog");
    }
}
