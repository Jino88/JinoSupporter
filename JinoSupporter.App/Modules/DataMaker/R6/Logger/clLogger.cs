using System;
using System.Collections.Concurrent;
using System.IO;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Threading;

namespace DataMaker.Logger
{
    public static class clLogger
    {
        public static RichTextBox WpfTextBox { get; set; }
        public static bool ShowVerboseUiLogs { get; set; } = false;

        private static string _logFilePath = null;
        private static readonly object _lockObj = new object();
        private static readonly ConcurrentQueue<(string Message, System.Windows.Media.Brush? Brush)> _pendingUiLogs = new();
        private static DispatcherTimer? _uiFlushTimer;

        public static void InitializeFileLogging(string logDirectory = null)
        {
            try
            {
                if (string.IsNullOrEmpty(logDirectory))
                {
                    logDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs");
                }

                if (!Directory.Exists(logDirectory))
                {
                    Directory.CreateDirectory(logDirectory);
                }

                string fileName = $"log_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
                _logFilePath = Path.Combine(logDirectory, fileName);

                File.WriteAllText(_logFilePath, $"=== Log Started at {DateTime.Now:yyyy-MM-dd HH:mm:ss} ==={Environment.NewLine}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to initialize file logging: {ex.Message}");
                _logFilePath = null;
            }
        }

        public static void Log(string message)
        {
            WriteLog(message, forceUi: false, uiBrush: null);
        }

        public static void LogImportant(string message)
        {
            WriteLog(message, forceUi: true, uiBrush: null);
        }

        public static void LogWarning(string message)
        {
            WriteLog($"WARNING: {message}", forceUi: true, uiBrush: System.Windows.Media.Brushes.DarkOrange);
        }

        public static void LogError(string message)
        {
            WriteLog($"ERROR: {message}", forceUi: true, uiBrush: System.Windows.Media.Brushes.Red);
        }

        public static void LogException(Exception ex, string context)
        {
            WriteLog($"ERROR in {context}: {ex.Message}", forceUi: true, uiBrush: System.Windows.Media.Brushes.Red);
            WriteLog($"Stack Trace: {ex.StackTrace}", forceUi: false, uiBrush: null);
        }

        public static string GetLogFilePath()
        {
            return _logFilePath;
        }

        private static void WriteLog(string message, bool forceUi, System.Windows.Media.Brush uiBrush)
        {
            string logMessage = $"[{DateTime.Now:HH:mm:ss}] {message}";

            if (WpfTextBox != null && (forceUi || ShouldShowInUi(message)))
            {
                EnqueueUiLog(logMessage, uiBrush);
            }

            Console.WriteLine(logMessage);

            if (!string.IsNullOrEmpty(_logFilePath))
            {
                try
                {
                    lock (_lockObj)
                    {
                        File.AppendAllText(_logFilePath, logMessage + Environment.NewLine);
                    }
                }
                catch
                {
                }
            }
        }

        private static void EnqueueUiLog(string logMessage, System.Windows.Media.Brush? uiBrush)
        {
            _pendingUiLogs.Enqueue((logMessage, uiBrush));

            if (WpfTextBox == null)
            {
                return;
            }

            if (WpfTextBox.Dispatcher.CheckAccess())
            {
                EnsureUiFlushTimer();
                return;
            }

            _ = WpfTextBox.Dispatcher.BeginInvoke(new Action(EnsureUiFlushTimer));
        }

        private static void EnsureUiFlushTimer()
        {
            if (WpfTextBox == null)
            {
                return;
            }

            if (_uiFlushTimer == null)
            {
                _uiFlushTimer = new DispatcherTimer(DispatcherPriority.Background, WpfTextBox.Dispatcher)
                {
                    Interval = TimeSpan.FromMilliseconds(120)
                };
                _uiFlushTimer.Tick += (_, _) => FlushPendingUiLogs();
            }

            if (!_uiFlushTimer.IsEnabled)
            {
                _uiFlushTimer.Start();
            }
        }

        private static void FlushPendingUiLogs()
        {
            if (WpfTextBox == null)
            {
                _pendingUiLogs.Clear();
                _uiFlushTimer?.Stop();
                return;
            }

            int appended = 0;
            while (appended < 40 && _pendingUiLogs.TryDequeue(out var entry))
            {
                AppendParagraphToUi(entry.Message, entry.Brush);
                appended++;
            }

            if (_pendingUiLogs.IsEmpty)
            {
                _uiFlushTimer?.Stop();
            }

            WpfTextBox.ScrollToEnd();
        }

        private static void AppendParagraphToUi(string logMessage, System.Windows.Media.Brush? uiBrush)
        {
            Paragraph paragraph = new Paragraph(new Run(logMessage))
            {
                Margin = new System.Windows.Thickness(0)
            };

            if (uiBrush != null)
            {
                paragraph.Foreground = uiBrush;
            }

            WpfTextBox.Document.Blocks.Add(paragraph);
        }

        private static bool ShouldShowInUi(string message)
        {
            if (ShowVerboseUiLogs)
            {
                return true;
            }

            if (string.IsNullOrWhiteSpace(message))
            {
                return false;
            }

            string[] importantKeywords =
            {
                "ERROR",
                "WARNING",
                "failed",
                "completed",
                "success",
                "saved",
                "updated",
                "loaded",
                "loading json",
                "loading existing data",
                "starting report generation",
                "processing [",
                "creating combined allgroupsdetailreport",
                "checking for unmapped items",
                "unmapped items check completed",
                "auto-loading db",
                "log file:",
                "settings updated.",
                "model groups saved",
                "model groups loaded",
                "skip"
            };

            foreach (string keyword in importantKeywords)
            {
                if (message.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }
    }
}
