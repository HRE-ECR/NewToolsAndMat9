using System.Windows;
using System.Windows.Threading;
using Serilog;

namespace PdfTableExtractor.App;

public partial class App : System.Windows.Application
{
    private static string CrashLogPath => Path.Combine(AppContext.BaseDirectory, "crash.log");
    private static string AppLogPath => Path.Combine(AppContext.BaseDirectory, "app.log");

    public App()
    {
        AppDomain.CurrentDomain.UnhandledException += (_, e) =>
        {
            if (e.ExceptionObject is Exception ex)
                WriteCrash("AppDomain.UnhandledException", ex);
            else
                File.AppendAllText(CrashLogPath,
                    "[" + DateTime.UtcNow.ToString("u") + "] AppDomain.UnhandledException: " +
                    (e.ExceptionObject?.ToString() ?? "(null)") + Environment.NewLine);
        };

        TaskScheduler.UnobservedTaskException += (_, e) =>
        {
            WriteCrash("TaskScheduler.UnobservedTaskException", e.Exception);
            e.SetObserved();
        };
    }

    private void App_Startup(object sender, StartupEventArgs e)
    {
        try
        {
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Information()
                .WriteTo.File(
                    path: AppLogPath,
                    rollingInterval: RollingInterval.Day,
                    retainedFileCountLimit: 14,
                    shared: true)
                .CreateLogger();

            Log.Information("App starting. Version={Version} BaseDir={BaseDir}",
                typeof(App).Assembly.GetName().Version?.ToString() ?? "?", AppContext.BaseDirectory);

            var window = new MainWindow();
            MainWindow = window;
            window.Show();
        }
        catch (Exception ex)
        {
            WriteCrash("Startup", ex);
            var msg = "The application failed to start." + Environment.NewLine + Environment.NewLine
                      + "Error: " + ex.Message + Environment.NewLine + Environment.NewLine
                      + "A crash log was written to:" + Environment.NewLine + CrashLogPath;

            System.Windows.MessageBox.Show(msg, "PdfTableExtractor - Startup Error",
                MessageBoxButton.OK, MessageBoxImage.Error);

            Shutdown(-1);
        }
    }

    private void App_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        WriteCrash("DispatcherUnhandledException", e.Exception);
        var msg = "A fatal error occurred and the app must close." + Environment.NewLine + Environment.NewLine
                  + "Error: " + e.Exception.Message + Environment.NewLine + Environment.NewLine
                  + "Crash log:" + Environment.NewLine + CrashLogPath;

        System.Windows.MessageBox.Show(msg, "PdfTableExtractor - Fatal Error",
            MessageBoxButton.OK, MessageBoxImage.Error);

        e.Handled = true;
        Shutdown(-1);
    }

    private static void WriteCrash(string source, Exception ex)
    {
        try
        {
            var sb = new StringBuilder();
            sb.Append('[').Append(DateTime.UtcNow.ToString("u")).Append("] ").AppendLine(source);
            sb.AppendLine("Message: " + ex.Message);
            sb.AppendLine("Type: " + (ex.GetType().FullName ?? "(unknown)"));
            sb.AppendLine("Stack:");
            sb.AppendLine(ex.ToString());
            sb.AppendLine(new string('-', 80));
            File.AppendAllText(CrashLogPath, sb.ToString());
        }
        catch { }
    }
}
