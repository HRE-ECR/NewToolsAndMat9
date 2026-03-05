using PdfTableExtractor.App.Models;
using PdfTableExtractor.App.Utils;
using Serilog;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace PdfTableExtractor.App.Services;

public sealed class PdfExtractionService
{
    private static readonly string Section5Title = "materials and special tools";
    private static readonly string Section6Title = "references";

    public async Task<List<TableBlock>> ExtractAsync(string pdfPath, ExtractionMode mode, CancellationToken ct)
        => await Task.Run(() => ExtractInternal(pdfPath, mode, ct), ct);

    private List<TableBlock> ExtractInternal(string pdfPath, ExtractionMode mode, CancellationToken ct)
    {
        var blocks = new List<TableBlock>();
        using var document = PdfDocument.Open(pdfPath);

        bool inSection5 = false;
        ColumnLayout? currentLayout = null;
        TableBlock? currentBlock = null;

        for (int pageNumber = 1; pageNumber <= document.NumberOfPages; pageNumber++)
        {
            ct.ThrowIfCancellationRequested();

            var page = document.GetPage(pageNumber);
            var words = page.GetWords().ToList();
            var lines = PdfLineGrouper.GroupWordsIntoLines(words);

            foreach (var line in lines)
            {
                var t = TextNormalize.NormalizeLoose(line.Text);
                if (!inSection5 && t.Contains(Section5Title)) { inSection5 = true; break; }
                if (inSection5 && (t.StartsWith("6 ") || t.Contains(Section6Title))) { inSection5 = false; currentLayout = null; currentBlock = null; break; }
            }

            if (!inSection5) continue;

            foreach (var line in lines)
            {
                ct.ThrowIfCancellationRequested();

                var norm = TextNormalize.NormalizeLoose(line.Text);
                if (norm.StartsWith("6 ") || norm.Contains(Section6Title)) { inSection5 = false; currentLayout = null; currentBlock = null; break; }

                if (IsHeaderLine(line.Words))
                {
                    var subsection = DetectSubsection(norm);
                    if (subsection is null || !IsSubsectionRelevant(subsection, mode)) { currentLayout = null; currentBlock = null; continue; }
                    currentBlock = new TableBlock { Title = subsection };
                    blocks.Add(currentBlock);
                    currentLayout = ColumnLayout.FromHeaderWords(line.Words);
                    continue;
                }

                if (currentLayout is null || currentBlock is null) continue;

                var row = ExtractRowFromLine(line.Words, currentLayout);
                if (row is null) continue;

                if (string.IsNullOrWhiteSpace(row.Item))
                {
                    var prev = currentBlock.Rows.LastOrDefault();
                    if (prev != null)
                    {
                        if (!string.IsNullOrWhiteSpace(row.Description)) prev.Description = Merge(prev.Description, row.Description);
                        if (!string.IsNullOrWhiteSpace(row.PartNumber)) prev.PartNumber = Merge(prev.PartNumber, row.PartNumber);
                        if (!string.IsNullOrWhiteSpace(row.Qty)) prev.Qty = Merge(prev.Qty, row.Qty);
                    }
                    continue;
                }

                currentBlock.Rows.Add(row);
            }
        }

        return blocks.Where(b => b.Rows.Count > 0).ToList();
    }

    private static string Merge(string a, string b)
    {
        a = a?.Trim() ?? "";
        b = b?.Trim() ?? "";
        if (a.Length == 0) return b;
        if (b.Length == 0) return a;
        return a + " " + b;
    }

    private static bool IsHeaderLine(List<Word> words)
    {
        var tokens = words.Select(w => TextNormalize.NormalizeToken(w.Text)).ToList();
        bool hasItem = tokens.Contains("item");
        bool hasQty = tokens.Any(t => t.Contains("qty") || t.Contains("quantity"));
        var joined = string.Join("", tokens);
        bool hasPart = tokens.Contains("part") || joined.Contains("partnumber") || joined.Contains("partno");
        return hasItem && hasPart && hasQty;
    }

    private static string? DetectSubsection(string normalizedLine)
    {
        if (normalizedLine.Contains("materials")) return "Materials";
        if (normalizedLine.Contains("consumables")) return "Consumables";
        if (normalizedLine.Contains("special tools") || normalizedLine.Contains("special tooling")) return "Special Tools";
        if (normalizedLine.Contains("test equipment")) return "Test Equipment";
        if (normalizedLine.Contains("tools")) return "Tools";
        return null;
    }

    private static bool IsSubsectionRelevant(string subsection, ExtractionMode mode)
        => mode switch
        {
            ExtractionMode.Material => subsection is "Materials" or "Consumables",
            ExtractionMode.Tooling => subsection is "Tools" or "Special Tools" or "Test Equipment",
            _ => false
        };

    private static TableRow? ExtractRowFromLine(List<Word> words, ColumnLayout layout)
    {
        var itemWords = new List<Word>();
        var descWords = new List<Word>();
        var partWords = new List<Word>();
        var qtyWords = new List<Word>();

        foreach (var w in words)
        {
            var cx = (w.BoundingBox.Left + w.BoundingBox.Right) / 2.0;
            if (layout.CoshhIgnoreLeft.HasValue && layout.CoshhIgnoreRight.HasValue)
                if (cx >= layout.CoshhIgnoreLeft.Value && cx <= layout.CoshhIgnoreRight.Value) continue;

            if (cx >= layout.ItemLeft && cx < layout.ItemRight) itemWords.Add(w);
            else if (cx >= layout.DescLeft && cx < layout.DescRight) descWords.Add(w);
            else if (cx >= layout.PartLeft && cx < layout.PartRight) partWords.Add(w);
            else if (cx >= layout.QtyLeft && cx <= layout.QtyRight) qtyWords.Add(w);
        }

        static string Join(List<Word> ws) => string.Join(" ", ws.OrderBy(x => x.BoundingBox.Left).Select(x => x.Text)).Trim();

        var row = new TableRow { Item = Join(itemWords), Description = Join(descWords), PartNumber = Join(partWords), Qty = Join(qtyWords) };
        if (row.Item.Length == 0 && row.Description.Length == 0 && row.PartNumber.Length == 0 && row.Qty.Length == 0) return null;
        return row;
    }

    private sealed class ColumnLayout
    {
        public double ItemLeft { get; init; }
        public double ItemRight { get; init; }
        public double DescLeft { get; init; }
        public double DescRight { get; init; }
        public double PartLeft { get; init; }
        public double PartRight { get; init; }
        public double QtyLeft { get; init; }
        public double QtyRight { get; init; }
        public double? CoshhIgnoreLeft { get; init; }
        public double? CoshhIgnoreRight { get; init; }

        public static ColumnLayout FromHeaderWords(List<Word> headerWords)
        {
            Word? item = FindWord(headerWords, "item");
            Word? qty = FindWord(headerWords, "qty", "quantity");
            Word? coshh = FindWord(headerWords, "coshh");
            Word? part = FindWord(headerWords, "part");

            if (item is null || qty is null || part is null)
            {
                var left = headerWords.Min(w => w.BoundingBox.Left);
                var right = headerWords.Max(w => w.BoundingBox.Right);
                var width = right - left;
                return new ColumnLayout
                {
                    ItemLeft = left,
                    ItemRight = left + width * 0.15,
                    DescLeft = left + width * 0.15,
                    DescRight = left + width * 0.65,
                    PartLeft = left + width * 0.65,
                    PartRight = left + width * 0.85,
                    QtyLeft = left + width * 0.85,
                    QtyRight = right
                };
            }

            var itemLeft = item.BoundingBox.Left;
            var partLeft = part.BoundingBox.Left;
            var qtyLeft = qty.BoundingBox.Left;
            var headerRight = headerWords.Max(w => w.BoundingBox.Right);

            var itemRight = headerWords
                .Where(w => w.BoundingBox.Left > item.BoundingBox.Left + 1 && w.BoundingBox.Left < partLeft)
                .Select(w => w.BoundingBox.Left)
                .DefaultIfEmpty(item.BoundingBox.Right + 20)
                .Min();

            double? coshhLeft = coshh?.BoundingBox.Left;
            double? coshhRight = coshh is null ? null : partLeft;

            var descLeft = itemRight;
            var descRight = coshhLeft ?? partLeft;

            return new ColumnLayout
            {
                ItemLeft = itemLeft,
                ItemRight = itemRight,
                DescLeft = descLeft,
                DescRight = descRight,
                PartLeft = partLeft,
                PartRight = qtyLeft,
                QtyLeft = qtyLeft,
                QtyRight = headerRight + 2,
                CoshhIgnoreLeft = coshhLeft,
                CoshhIgnoreRight = coshhRight
            };
        }

        private static Word? FindWord(List<Word> words, params string[] targets)
        {
            foreach (var w in words)
            {
                var n = TextNormalize.NormalizeToken(w.Text);
                if (targets.Any(t => n == t)) return w;
            }
            foreach (var w in words)
            {
                var n = TextNormalize.NormalizeToken(w.Text);
                if (targets.Any(t => n.Contains(t))) return w;
            }
            return null;
        }
    }
}
