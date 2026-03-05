namespace PdfTableExtractor.App.Utils;

public static class FilenameHelper
{
    private static readonly Regex IdRegex = new(@"([A-Za-z]{2,6}\d{3,6})", RegexOptions.Compiled);

    public static string ExtractWorksheetIdentifier(string pdfPath)
    {
        var name = Path.GetFileNameWithoutExtension(pdfPath);
        var tokens = name.Split(new[] { '_', '-', ' ' }, StringSplitOptions.RemoveEmptyEntries);
        for (int i = tokens.Length - 1; i >= 0; i--)
        {
            var m = IdRegex.Match(tokens[i]);
            if (m.Success) return SanitizeSheetName(m.Groups[1].Value.ToUpperInvariant());
        }
        var fallback = IdRegex.Match(name);
        if (fallback.Success) return SanitizeSheetName(fallback.Groups[1].Value.ToUpperInvariant());
        return SanitizeSheetName(name);
    }

    public static string SanitizeSheetName(string name)
    {
        if (string.IsNullOrWhiteSpace(name)) name = "Sheet";
        var invalid = new HashSet<char>(@":\/?*[]".ToCharArray());
        var cleaned = new string(name.Select(c => invalid.Contains(c) ? '_' : c).ToArray()).Trim();
        if (cleaned.Length > 31) cleaned = cleaned[..31];
        return string.IsNullOrWhiteSpace(cleaned) ? "Sheet" : cleaned;
    }
}
