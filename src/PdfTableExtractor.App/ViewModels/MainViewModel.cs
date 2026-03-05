using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using PdfTableExtractor.App.Models;
using PdfTableExtractor.App.Services;
using Serilog;
using Serilog.Events;
using System.Windows;

namespace PdfTableExtractor.App.ViewModels;

using WinForms = System.Windows.Forms;

public partial class MainViewModel : ObservableObject
{
    private readonly PdfExtractionService _pdfService = new();
    private readonly ExcelExportService _excelService = new();
    private readonly RunOrchestrator _orchestrator;

    private readonly StringBuilder _logBuilder = new();
    private const int MaxLogChars = 250_000;

    public MainViewModel()
    {
        _orchestrator = new RunOrchestrator(_pdfService, _excelService);
        IsMaterialMode = true;
        StatusText = "Ready.";
        UpdateCanStart();
        SelectedPdfFiles.CollectionChanged += (_, __) => UpdateCanStart();
    }

    public ObservableCollection<string> SelectedPdfFiles { get; } = new();

    [ObservableProperty] private string outputFolder = "";
    partial void OnOutputFolderChanged(string value) => UpdateCanStart();

    [ObservableProperty] private string statusText = "";
    [ObservableProperty] private int progressPercent;

    [ObservableProperty] private bool isBusy;
    partial void OnIsBusyChanged(bool value) => UpdateCanStart();

    [ObservableProperty] private string liveLogText = "";

    [ObservableProperty] private bool isMaterialMode;
    partial void OnIsMaterialModeChanged(bool value) { if (value) IsToolingMode = false; UpdateCanStart(); }

    [ObservableProperty] private bool isToolingMode;
    partial void OnIsToolingModeChanged(bool value) { if (value) IsMaterialMode = false; UpdateCanStart(); }

    [ObservableProperty] private bool canStart;

    [RelayCommand]
    private void ClearLog()
    {
        _logBuilder.Clear();
        LiveLogText = "";
    }

    [RelayCommand]
    private void SelectPdfFiles()
    {
        var dlg = new Microsoft.Win32.OpenFileDialog { Filter = "PDF files (*.pdf)|*.pdf", Multiselect = true };
        if (dlg.ShowDialog() == true)
        {
            SelectedPdfFiles.Clear();
            foreach (var f in dlg.FileNames.OrderBy(x => x)) SelectedPdfFiles.Add(f);
            AppendLogLocal("Selected " + SelectedPdfFiles.Count + " PDF(s).");
        }
    }

    [RelayCommand]
    private void SelectFolder()
    {
        using var dlg = new WinForms.FolderBrowserDialog { Description = "Select folder containing PDF files", UseDescriptionForTitle = true, ShowNewFolderButton = false };
        if (dlg.ShowDialog() == WinForms.DialogResult.OK && Directory.Exists(dlg.SelectedPath))
        {
            var files = Directory.GetFiles(dlg.SelectedPath, "*.pdf", SearchOption.TopDirectoryOnly).OrderBy(x => x).ToList();
            SelectedPdfFiles.Clear();
            foreach (var f in files) SelectedPdfFiles.Add(f);
            AppendLogLocal("Found " + SelectedPdfFiles.Count + " PDF(s) in folder.");
        }
    }

    [RelayCommand]
    private void SelectOutputFolder()
    {
        using var dlg = new WinForms.FolderBrowserDialog { Description = "Select output folder", UseDescriptionForTitle = true, ShowNewFolderButton = true };
        if (dlg.ShowDialog() == WinForms.DialogResult.OK && Directory.Exists(dlg.SelectedPath))
        {
            OutputFolder = dlg.SelectedPath;
            AppendLogLocal("Output folder: " + OutputFolder);
        }
    }

    [RelayCommand]
    private async Task StartExtraction()
    {
        if (!CanStart) return;
        IsBusy = true;
        ProgressPercent = 0;
        StatusText = "Starting...";

        try
        {
            ConfigureSerilogForRun(OutputFolder);
            var mode = IsMaterialMode ? ExtractionMode.Material : ExtractionMode.Tooling;
            var pdfs = SelectedPdfFiles.ToList();

            var progress = new Progress<(int current, int total, string message)>(p =>
            {
                StatusText = p.message;
                ProgressPercent = (int)Math.Round((double)p.current / p.total * 100);
            });

            var (workbookPath, results) = await _orchestrator.RunAsync(pdfs, OutputFolder, mode, progress, CancellationToken.None);
            var ok = results.Count(r => r.Success);
            var fail = results.Count - ok;
            StatusText = workbookPath is not null ? ("Done. OK:" + ok + " Fail:" + fail) : ("Done. OK:" + ok + " Fail:" + fail + " (no workbook)");
        }
        catch (Exception ex)
        {
            StatusText = "Error occurred. See log.";
            Log.Error(ex, "Run failed");
        }
        finally
        {
            IsBusy = false;
        }
    }

    private void UpdateCanStart()
    {
        CanStart = !IsBusy && SelectedPdfFiles.Count > 0 && Directory.Exists(OutputFolder) && (IsMaterialMode || IsToolingMode);
    }

    private void ConfigureSerilogForRun(string outputFolder)
    {
        Directory.CreateDirectory(outputFolder);
        var runLogPath = Path.Combine(outputFolder, "run.log");
        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Information()
            .MinimumLevel.Override("Microsoft", LogEventLevel.Warning)
            .WriteTo.File(runLogPath, rollingInterval: RollingInterval.Day, retainedFileCountLimit: 14, shared: true)
            .WriteTo.Sink(new PdfTableExtractor.App.Services.SerilogUiSink(System.Windows.Application.Current.Dispatcher, AppendLogLocal))
            .CreateLogger();
        AppendLogLocal("Logging to: " + runLogPath);
    }

    private void AppendLogLocal(string line)
    {
        if (_logBuilder.Length > MaxLogChars) _logBuilder.Remove(0, _logBuilder.Length / 2);
        _logBuilder.AppendLine(line);
        LiveLogText = _logBuilder.ToString();
    }
}
