using PdfTableExtractor.App.Models;
using PdfTableExtractor.App.Utils;
using Serilog;

namespace PdfTableExtractor.App.Services;

public sealed class RunOrchestrator
{
    private readonly PdfExtractionService _pdf;
    private readonly ExcelExportService _excel;

    public RunOrchestrator(PdfExtractionService pdf, ExcelExportService excel)
    {
        _pdf = pdf;
        _excel = excel;
    }

    public async Task<(string? workbookPath, List<ExtractionResult> results)> RunAsync(
        List<string> pdfFiles,
        string outputFolder,
        ExtractionMode mode,
        IProgress<(int current, int total, string message)> progress,
        CancellationToken ct)
    {
        Directory.CreateDirectory(outputFolder);
        var results = new List<ExtractionResult>();

        for (int i = 0; i < pdfFiles.Count; i++)
        {
            ct.ThrowIfCancellationRequested();
            var file = pdfFiles[i];
            var id = FilenameHelper.ExtractWorksheetIdentifier(file);
            progress.Report((i + 1, pdfFiles.Count, "Processing " + (i + 1) + "/" + pdfFiles.Count + ": " + Path.GetFileName(file)));

            try
            {
                var blocks = await _pdf.ExtractAsync(file, mode, ct);
                results.Add(new ExtractionResult
                {
                    SourcePdfPath = file,
                    WorksheetName = id,
                    Success = blocks.Count > 0,
                    Message = blocks.Count > 0 ? "OK" : "No relevant tables found",
                    Blocks = blocks
                });
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Failed processing {File}", file);
                results.Add(new ExtractionResult
                {
                    SourcePdfPath = file,
                    WorksheetName = id,
                    Success = false,
                    Message = ex.Message
                });
            }
        }

        string? workbookPath = null;
        if (results.Any(r => r.Success))
            workbookPath = _excel.SaveWorkbook(results, outputFolder, mode);

        return (workbookPath, results);
    }
}
