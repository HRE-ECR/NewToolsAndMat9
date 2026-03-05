using ClosedXML.Excel;
using PdfTableExtractor.App.Models;
using Serilog;

namespace PdfTableExtractor.App.Services;

public sealed class ExcelExportService
{
    public string SaveWorkbook(List<ExtractionResult> results, string outputFolder, ExtractionMode mode)
    {
        var now = DateTime.Now;
        var dateText = now.ToString("dd MMM yyyy");
        var workbookName = mode + " - " + dateText + ".xlsx";
        var fullPath = Path.Combine(outputFolder, workbookName);

        using var wb = new XLWorkbook();

        foreach (var r in results.Where(x => x.Success))
        {
            var sheetName = r.WorksheetName;
            if (wb.Worksheets.Any(ws => ws.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase)))
                sheetName = sheetName.Length > 28 ? sheetName[..28] + "_2" : sheetName + "_2";

            var ws = wb.Worksheets.Add(sheetName);

            ws.Cell(1, 1).Value = "Item";
            ws.Cell(1, 2).Value = mode == ExtractionMode.Material ? "Material / Consumables" : "Tools / Special Tools / Test Equipment";
            ws.Cell(1, 3).Value = "Part Number";
            ws.Cell(1, 4).Value = "QTY";
            ws.Range(1, 1, 1, 4).Style.Font.Bold = true;
            ws.SheetView.FreezeRows(1);

            int row = 2;
            foreach (var block in r.Blocks)
            {
                ws.Range(row, 1, row, 4).Merge();
                ws.Cell(row, 1).Value = block.Title;
                ws.Cell(row, 1).Style.Font.Bold = true;
                row++;

                foreach (var tr in block.Rows)
                {
                    ws.Cell(row, 1).SetValue(tr.Item);
                    ws.Cell(row, 2).SetValue(tr.Description);
                    ws.Cell(row, 3).SetValue(tr.PartNumber);
                    ws.Cell(row, 4).SetValue(tr.Qty);
                    row++;
                }
                row++;
            }

            ws.Columns().AdjustToContents();
        }

        wb.SaveAs(fullPath);
        Log.Information("Excel workbook saved: {Path}", fullPath);
        return fullPath;
    }
}
