using Serilog.Core;
using Serilog.Events;
using System.Windows.Threading;

namespace PdfTableExtractor.App.Services;

public sealed class SerilogUiSink : ILogEventSink
{
    private readonly Dispatcher _dispatcher;
    private readonly Action<string> _appendLine;

    public SerilogUiSink(Dispatcher dispatcher, Action<string> appendLine)
    {
        _dispatcher = dispatcher;
        _appendLine = appendLine;
    }

    public void Emit(LogEvent logEvent)
    {
        var line = "[" + logEvent.Timestamp.ToString("HH:mm:ss") + "] " + logEvent.Level + ": " + logEvent.RenderMessage();
        if (logEvent.Exception != null) line += " | EX: " + logEvent.Exception.Message;
        _dispatcher.BeginInvoke(() => _appendLine(line));
    }
}
