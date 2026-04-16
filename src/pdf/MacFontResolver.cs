namespace Dazzle.Pdf;

using PdfSharp.Fonts;

internal class MacFontResolver : IFontResolver
{
    private static readonly string[] SearchPaths =
    {
        "/System/Library/Fonts/Supplemental/",
        "/System/Library/Fonts/",
        "/Library/Fonts/",
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Library/Fonts/")
    };

    // Map DSSSL/PostScript font names to macOS TTF file names
    private static readonly Dictionary<string, string> Aliases = new(StringComparer.OrdinalIgnoreCase)
    {
        { "Helvetica", "Arial" },
        { "Courier", "Courier New" },
        { "Times Roman", "Times New Roman" },
        { "Times", "Times New Roman" },
    };

    public FontResolverInfo ResolveTypeface(string familyName, bool bold, bool italic)
    {
        string mapped = MapFamily(familyName);
        string faceName = BuildFaceName(mapped, bold, italic);

        if (FindFontFile(faceName) != null)
            return new FontResolverInfo(faceName);

        // Try base family name (some fonts only have a single file)
        if (FindFontFile(mapped) != null)
            return new FontResolverInfo(mapped, bold, italic);

        // Last resort: fall back to Arial which is always available
        string fallback = BuildFaceName("Arial", bold, italic);
        if (FindFontFile(fallback) != null)
            return new FontResolverInfo(fallback);

        return new FontResolverInfo("Arial", bold, italic);
    }

    public byte[] GetFont(string faceName)
    {
        string path = FindFontFile(faceName);
        if (path != null)
            return File.ReadAllBytes(path);

        // Fallback
        path = FindFontFile("Arial");
        if (path != null)
            return File.ReadAllBytes(path);

        return null;
    }

    private static string MapFamily(string familyName)
    {
        return Aliases.TryGetValue(familyName, out var mapped) ? mapped : familyName;
    }

    private static string BuildFaceName(string familyName, bool bold, bool italic)
    {
        if (bold && italic)
            return familyName + " Bold Italic";
        if (bold)
            return familyName + " Bold";
        if (italic)
            return familyName + " Italic";
        return familyName;
    }

    private static string FindFontFile(string faceName)
    {
        foreach (var dir in SearchPaths)
        {
            var ttf = Path.Combine(dir, faceName + ".ttf");
            if (File.Exists(ttf))
                return ttf;
        }
        return null;
    }
}
