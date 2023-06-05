using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace WPFExceptionHandler
{
    public enum LogEntryType
    {
        Info,
        Warning,
        GenericError,
        CriticalError
    }

    public class ExceptionManagement
    {
        private static string _exceptionLogFilePathAlt = "";
        // TODO: hier wird ein Fehler geworfen und führt zu einem StackOverflow!
        private static string _exceptionLogFilePath
        {
            get
            {
                string path = Path.Combine(AppContext.BaseDirectory != null ? AppContext.BaseDirectory : AppDomain.CurrentDomain.BaseDirectory);
                path = Path.Combine(Path.GetTempPath(), DateTime.Now.Ticks.ToString() + "_temp.log");
                //try
                //{
                //    if (!Directory.Exists(path))
                //    {
                //        Directory.CreateDirectory(path);
                //        path = Path.Combine(path, DateTime.Now.ToString("yyyy-MM-dd") + "_debug.log");
                //    }
                //}
                //catch (Exception ex)
                //{
                //    Debug.Print(ex.Message);
                //    Debug.Print("Error creating log file with app path.");
                //    Debug.Print("Creating log path in temp dir.");
                //    path = Path.Combine(Path.GetTempPath(), DateTime.Now.Ticks.ToString() + "_temp.log");
                //}
                return path;
            }
        }
        public static string ExceptionLogFilePath { get => string.IsNullOrWhiteSpace(_exceptionLogFilePathAlt) ? _exceptionLogFilePath : _exceptionLogFilePathAlt; }
        public delegate void ExceptionCatchedEventHandler(object sender, ExceptionCatchedEventArgs args);
        public delegate void LogDebugAddEvenHandler(object sender, LogDebugAddEventArgs args);
        public static List<LogEntry> LogEntries { get; }

        public static event ExceptionCatchedEventHandler ExceptionCatched;
        public static void RaiseExceptionCatchedEvent(object sender, Exception exception)
        {
            ExceptionCatched?.Invoke(sender, new ExceptionCatchedEventArgs(exception));
        }
        public static event LogDebugAddEvenHandler LogDebugAdded;
        public static void RaiseLogDebugAdded(object sender, string message, LogEntryType entryType, DateTime timeStamp)
        {
            LogEntry entry = new LogEntry(message, entryType, timeStamp);
            LogEntries?.Add(entry);
            LogDebugAdded?.Invoke(sender, new LogDebugAddEventArgs(entry));
        }

        static ExceptionManagement()
        {
            LogEntries = new List<LogEntry>();

            Application.Current.Exit += new ExitEventHandler((s, e) =>
            {
                if (e.ApplicationExitCode != 0)
                    LogCriticalError("ExceptionManagement exited with fault (exit code = " + e.ApplicationExitCode + ").");
                else
                    LogDebug("ExceptionManagement exited without any fault.");
            });
        }

        public static void CreateExceptionManagement(Application app, AppDomain domain)
        {
            string appUri = new FileInfo(domain.BaseDirectory).Name;
            app.DispatcherUnhandledException += new DispatcherUnhandledExceptionEventHandler((s, e) =>
            {
                LogGenericError(e.Exception);
                e.Handled = true;
            });

            app.Dispatcher.UnhandledException += new DispatcherUnhandledExceptionEventHandler((s, e) =>
            {
                LogGenericError(e.Exception);
                e.Handled = true;
            });

            domain.UnhandledException += new UnhandledExceptionEventHandler((s, e) =>
            {
                LogCriticalError(e.);
            });

            domain.FirstChanceException += new EventHandler<FirstChanceExceptionEventArgs>((s, e) =>
            {
                LogCriticalError(e.Exception);
            });

            app.Exit += new ExitEventHandler((s, e) =>
            {
                if (e.ApplicationExitCode != 0)
                    LogCriticalError("Application exited with fault (exit code = " + e.ApplicationExitCode + ").");
                else
                    LogDebug("Application exited without any fault.");
            });
        }

        private static void OnUnobservedTaskException(object sender, UnobservedTaskExceptionEventArgs e)
        {
            LogGenericError(e.Exception);
            e.SetObserved();
        }

        public static bool SetAlternativeLogFilePath(string filePath)
        {
            bool filePathValid;

            try
            {
                string dirPath = Path.GetDirectoryName(filePath);
                string fileName = Path.GetFileNameWithoutExtension(filePath);
                string fileExtension = Path.GetExtension(filePath);
                char[] invalidFileNameChars = Path.GetInvalidFileNameChars();
                char[] invalidPathChars = Path.GetInvalidPathChars();

                bool invalidPath = invalidPathChars.Any(dirChar => invalidPathChars.Contains(dirChar));
                bool invalidFileName = invalidFileNameChars.Any(fileNameChar => invalidFileNameChars.Contains(fileNameChar));
                filePathValid = !(invalidPath || invalidFileName);

                if (filePathValid) _exceptionLogFilePathAlt = filePath;
            }
            catch (Exception exception)
            {
                filePathValid = false;
                LogGenericError(exception.Message);
                LogGenericError("Alternative file path could not be set ('" + filePath + ").");

            }

            return filePathValid;
        }

        private static void CreateLogFile()
        {
            if (File.Exists(ExceptionLogFilePath))
                return;
            
            File.Create(ExceptionLogFilePath);
            LogDebug("Log created.");
        }

        private static string CreateMessageString(string message, LogEntryType entryType, DateTime timeStamp)
        {
            string threadname = Thread.CurrentThread.Name;
            string formattedtimestamp = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss.fff");
            string logentrytype = Enum.GetName(typeof(LogEntryType), entryType);
            if (string.IsNullOrEmpty(threadname))
                return string.Format("{0}: ({2}) {3}", formattedtimestamp, logentrytype, message);
            else
                return string.Format("{0}: [{1}] ({2}) {3}", formattedtimestamp, threadname, logentrytype, message);
        }
        public static void LogDebug(string message) => WriteLogEntry(message, LogEntryType.Info);
        public static void LogDebugBool(string message, bool value, string trueString, string falseString) => WriteLogEntry(message.Replace("%1", value ? trueString : falseString), LogEntryType.Info);
        public static void LogWarning(string message) => WriteLogEntry(message, LogEntryType.Warning);
        public static void LogGenericError(string message) => WriteLogEntry(message, LogEntryType.GenericError);
        public static void LogCriticalError(string message) => WriteLogEntry(message, LogEntryType.CriticalError);
        public static void LogGenericError(Exception exception) => WriteLogEntry(exception.Message, LogEntryType.GenericError);
        public static void LogCriticalError(Exception exception) => WriteLogEntry(exception.Message, LogEntryType.CriticalError);
        private static void WriteLogEntry(string message, LogEntryType entryType)
        {
            CreateLogFile();
            DateTime timeStamp = DateTime.Now;
            string exceptionMessage = CreateMessageString(message, entryType, timeStamp);
            string dirPath = Path.GetDirectoryName(ExceptionLogFilePath);
            if (dirPath == null) return;
            if (!Directory.Exists(dirPath)) Directory.CreateDirectory(dirPath);
            File.AppendAllText(ExceptionLogFilePath, exceptionMessage);
            RaiseLogDebugAdded(new object(), message, entryType, timeStamp);
        }

        public static string[] GetAllLines() => File.ReadAllLines(ExceptionLogFilePath);
        public static void CatchException(object sender, FirstChanceExceptionEventArgs args)
        {
            LogCriticalError(message: $"Error in '{args.Exception.TargetSite?.Name}': {args.Exception.Message}");
            RaiseExceptionCatchedEvent(new object(), args.Exception);
        }
    }

    public class LogEntry
    {
        public DateTime TimeStamp { get; }
        public string TimeStampText => TimeStamp.ToShortDateString() + " " + TimeStamp.ToLongTimeString();
        public string Message { get; }
        public LogEntryType EntryType { get; }
        public string EntryTypeText => Enum.GetName(typeof(LogEntryType), EntryType);

        public LogEntry(string message, LogEntryType entryType, DateTime timeStamp)
        {
            Message = message;
            EntryType = entryType;
            TimeStamp = timeStamp;
        }

        public override bool Equals(object obj)
        {
            return obj is LogEntry entry &&
                   Message == entry.Message;
        }

        public override int GetHashCode()
        {
            return 460171812 + EqualityComparer<string>.Default.GetHashCode(Message);
        }
    }

    public class ExceptionCatchedEventArgs
    {
        public Exception CatchedException { get; }
        public ExceptionCatchedEventArgs(Exception catchedException) => CatchedException = catchedException;

        public override bool Equals(object obj)
        {
            return obj is ExceptionCatchedEventArgs args &&
                   EqualityComparer<Exception>.Default.Equals(CatchedException, args.CatchedException);
        }

        public override int GetHashCode()
        {
            return -1688175730 + EqualityComparer<Exception>.Default.GetHashCode(CatchedException);
        }
    }

    public class LogDebugAddEventArgs
    {
        public LogEntry Entry { get; }
        public LogDebugAddEventArgs(LogEntry entry) => Entry = entry;

        public override bool Equals(object obj)
        {
            return obj is LogDebugAddEventArgs args &&
                   EqualityComparer<LogEntry>.Default.Equals(Entry, args.Entry);
        }

        public override int GetHashCode()
        {
            return 1970138155 + EqualityComparer<LogEntry>.Default.GetHashCode(Entry);
        }
    }
}
