using UglyToad.PdfPig.Content;

namespace PdfTableExtractor.App.Utils;

public sealed class PdfLine
{
    public double Y { get; init; }
    public List<Word> Words { get; init; } = new();
    public string Text => string.Join(" ", Words.OrderBy(w => w.BoundingBox.Left).Select(w => w.Text));
}

public static class PdfLineGrouper
{
    public static List<PdfLine> GroupWordsIntoLines(IEnumerable<Word> words, double yTolerance = 2.0)
    {
        var ordered = words.OrderByDescending(w => w.BoundingBox.Top).ThenBy(w => w.BoundingBox.Left).ToList();
        var lines = new List<PdfLine>();
        foreach (var word in ordered)
        {
            var y = word.BoundingBox.Top;
            var line = lines.FirstOrDefault(l => Math.Abs(l.Y - y) <= yTolerance);
            if (line is null) lines.Add(new PdfLine { Y = y, Words = new List<Word> { word } });
            else line.Words.Add(word);
        }
        foreach (var l in lines) l.Words.Sort((a, b) => a.BoundingBox.Left.CompareTo(b.BoundingBox.Left));
        return lines.OrderByDescending(l => l.Y).ToList();
    }
}
