namespace PdfTableExtractor.App.Models;

public sealed class TableBlock
{
    public string Title { get; set; } = "";
    public List<TableRow> Rows { get; set; } = new();
}
