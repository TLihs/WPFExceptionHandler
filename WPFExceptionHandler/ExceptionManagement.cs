using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace WPFExceptionHandler
{
    public static class ExceptionManagement
    {
        private enum LogEntryType
        {
            Info,
            Warning,
            GenericError,
            CriticalError
        }

        private static string _exceptionLogPathAlt;
        private static string _exceptionLogPath;
        private static bool _initialized = false;
        private static FileStream _logFileStream;
        private static Queue<byte[]> _logEntries = new Queue<byte[]>();

        public static string ExceptionLogFilePath => string.IsNullOrWhiteSpace(_exceptionLogPathAlt) ? Path.Combine(_exceptionLogPath, DateTime.Now.ToString("yyyy-MM-dd") + ".log") : _exceptionLogPathAlt;
        public delegate void ExceptionCatchedEventHandler(object sender, ExceptionCatchedEventArgs args);

        public static event ExceptionCatchedEventHandler ExceptionCatched;
        public static void RaiseExceptionCatchedEvent(object sender, Exception exception)
        {
            ExceptionCatched?.Invoke(sender, new ExceptionCatchedEventArgs(exception));
        }

        static ExceptionManagement()
        {
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
            if (_initialized)
            {
                LogWarning("ExceptionManagement::CreateExceptionManagement(...) - ExceptionManagement already initialized.");
                return;
            }

            string appname = System.Reflection.Assembly.GetEntryAssembly().GetName().Name;
            Debug.Print("BaseDirectory: " + appname);

            _exceptionLogPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), appname, "Log");

            TaskScheduler.UnobservedTaskException += new EventHandler<UnobservedTaskExceptionEventArgs>((s, e) =>
            {
                try
                {
                    LogGenericError(e.Exception);
                    e.SetObserved();
                }
                catch (Exception ex)
                {
                    LogCriticalError("Error occured while logging exception: " + ex.Message);
                }
            });

            app.DispatcherUnhandledException += new DispatcherUnhandledExceptionEventHandler((s, e) =>
            {
                try
                {
                    LogGenericError(e.Exception);
                    e.Handled = true;
                }
                catch (Exception ex)
                {
                    LogCriticalError("Error occured while logging exception: " + ex.Message);
                }
            });

            app.Dispatcher.UnhandledException += new DispatcherUnhandledExceptionEventHandler((s, e) =>
            {
                try
                {
                    LogGenericError(e.Exception);
                    e.Handled = true;
                }
                catch (Exception ex)
                {
                    LogCriticalError("Error occured while logging exception: " + ex.Message);
                }
            });

            domain.UnhandledException += new UnhandledExceptionEventHandler((s, e) =>
            {
                bool exceptioncatched = false;
                
                if (e != null)
                    try
                    {
                        LogCriticalError((Exception)e.ExceptionObject);
                        exceptioncatched = true;
                    }
                    catch (Exception ex)
                    {
                        LogCriticalError("Error occured while logging exception: " + ex.Message);
                    }
                
                if (!exceptioncatched)
                    LogCriticalError("Unknown critical exception occured.");

                if (e.IsTerminating)
                    LogCriticalError("App is terminating.");
            });

            domain.FirstChanceException += new EventHandler<FirstChanceExceptionEventArgs>((s, e) =>
            {
                try
                {
                    LogCriticalError(e.Exception);
                }
                catch (Exception ex)
                {
                    LogCriticalError("Error occured while logging exception: " + ex.Message);
                }
            });

            app.Exit += new ExitEventHandler((s, e) =>
            {
                if (e.ApplicationExitCode != 0)
                    LogCriticalError("Application exited with fault (exit code = " + e.ApplicationExitCode + ").");
                else
                    LogDebug("Application exited without any fault.");
            });

            _initialized = true;
        }

        public static bool SetAlternativeLogFilePath(string filePath)
        {
            bool filePathValid = false;

            try
            {
                if (Uri.IsWellFormedUriString(filePath, UriKind.Absolute))
                {
                    string fileExtension = Path.GetExtension(filePath);

                    filePathValid = (fileExtension == "log");
                }

                if (filePathValid)
                    _exceptionLogPathAlt = filePath;
                else
                    LogGenericError(string.Format("Alternative file path invalid ('{0}').", filePath));
            }
            catch (Exception exception)
            {
                filePathValid = false;
                LogGenericError(exception.Message);
                LogGenericError(string.Format("Alternative file path could not be set ('{0}').", filePath));
            }

            return filePathValid;
        }

        private static void CreateLogFile()
        {
            string logpath = ExceptionLogFilePath;
            bool newlog = false;

            FileInfo logfileinfo = new FileInfo(logpath);
            if (!File.Exists(logpath))
            {
                Directory.CreateDirectory(logfileinfo.DirectoryName);
                newlog = true;
            }

            if (_logFileStream == null)
            {
                _logFileStream = logfileinfo.Open(FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.Read);

                if (newlog)
                    LogDebug("Log created.");
                else
                {
                    _logFileStream.Position = _logFileStream.Length;
                    LogDebug("Log reopened.");
                }
            }
        }

        private static string CreateMessageString(string message, LogEntryType entryType, DateTime timeStamp)
        {
            string threadname = Thread.CurrentThread.Name;
            string formattedtimestamp = timeStamp.ToString("yyyy-MM-dd hh:mm:ss.fff");
            string logentrytype = Enum.GetName(typeof(LogEntryType), entryType);
            if (string.IsNullOrEmpty(threadname))
                return string.Format("{0}: ({1}) {2}\r\n", formattedtimestamp, logentrytype, message);
            else
                return string.Format("{0}: [{1}] ({2}) {3}\r\n", formattedtimestamp, threadname, logentrytype, message);
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
            DateTime timestamp = DateTime.Now;

            CreateLogFile();

            Task.Run(() =>
            {
                string logmessage = CreateMessageString(message, entryType, timestamp);
                byte[] logmessagebytes = Encoding.UTF8.GetBytes(logmessage);
                try
                {
                    while (_logEntries.Count > 0)
                        _logFileStream.Write(_logEntries.Dequeue(), 0, logmessagebytes.Length);
                    _logFileStream.Write(logmessagebytes, 0, logmessagebytes.Length);
                    _logFileStream.Flush();
                }
                catch (Exception ex)
                {
                    _logEntries.Enqueue(logmessagebytes);
                }
            });
        }

        public static string[] GetAllLines() => File.ReadAllLines(ExceptionLogFilePath);
        public static void CatchException(object sender, FirstChanceExceptionEventArgs args)
        {
            LogCriticalError(message: $"Error in '{args.Exception.TargetSite?.Name}': {args.Exception.Message}");
            RaiseExceptionCatchedEvent(new object(), args.Exception);
        }

        private class LogEntry
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

        private class LogDebugAddEventArgs
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
}
