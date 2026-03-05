namespace PdfTableExtractor.App.Models;

public sealed class ExtractionResult
{
    public string SourcePdfPath { get; set; } = "";
    public string WorksheetName { get; set; } = "";
    public bool Success { get; set; }
    public string Message { get; set; } = "";
    public List<TableBlock> Blocks { get; set; } = new();
}
