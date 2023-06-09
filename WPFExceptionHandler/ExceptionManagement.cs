using Microsoft.VisualStudio;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace WPFExceptionHandler
{
    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    public static class ExceptionManagement
    {
        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// Enumeration of types of log entries that are defined for exception handling and/or debug logging.
        /// </summary>
        public enum LogEntryType
        {
            Info,
            Warning,
            GenericError,
            CriticalError
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        public delegate void LogDebugAddedEventHandler(object sender, LogDebugAddedEventArgs args);

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        public static event LogDebugAddedEventHandler LogDebugAdded;
        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        private static void RaiseLogDebugAdded(object sender, LogEntry entry)
        {
            LogDebugAdded?.Invoke(sender, new LogDebugAddedEventArgs(entry));
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        private static string _exceptionLogPathAlt;
        private static string _exceptionLogPath;
        private static bool _initialized = false;
        private static bool _shutdownStarted = false;
        private static bool _disposed = false;
        private static bool _includeDebugInformation = true;
        private static FileStream _logFileStream;
        private static Queue<byte[]> _logEntries = new Queue<byte[]>();
        private static Queue<Task> _pendingLogTasks = new Queue<Task>();
        private static Task _currentLogTask = null;

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        public static string ExceptionLogFilePath => string.IsNullOrWhiteSpace(_exceptionLogPathAlt) ? 
            Path.Combine(_exceptionLogPath, DateTime.Now.ToString("yyyy-MM-dd") + ".log") : _exceptionLogPathAlt;
        public static bool UseFileLogging { get; set; }

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        static ExceptionManagement()
        {
            Application.Current.Exit += new ExitEventHandler((s, e) =>
            {
                if (e.ApplicationExitCode != 0)
                    LogCriticalError("ExceptionManagement exited with fault (exit code = " + e.ApplicationExitCode + ").");
                else
                    LogDebug("ExceptionManagement exited without any fault.");
            });

            UseFileLogging = false;
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        public static void CreateExceptionManagement(Application app, AppDomain domain, bool includeDebugInformation = true)
        {
            if (_initialized)
            {
                LogWarning("ExceptionManagement::CreateExceptionManagement(...) - ExceptionManagement already initialized.");
                return;
            }

            _includeDebugInformation = includeDebugInformation;
            string appname = System.Reflection.Assembly.GetEntryAssembly().GetName().Name;

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
                        LogGenericError((Exception)e.ExceptionObject);
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

            //domain.FirstChanceException += new EventHandler<FirstChanceExceptionEventArgs>((s, e) =>
            //{
            //    try
            //    {
            //        LogGenericError(e.Exception);
            //        Console.WriteLine(e.Exception.Source.ToString());
            //    }
            //    catch (Exception ex)
            //    {
            //        LogCriticalError("Error occured while logging exception: " + ex.Message);
            //    }
            //});

            app.Exit += new ExitEventHandler((s, e) =>
            {
                if (e.ApplicationExitCode != 0)
                    LogCriticalError("Application exited with fault (exit code = " + e.ApplicationExitCode + ").");
                else
                    LogDebug("Application exited without any fault.");

                Shutdown();
            });

            _initialized = true;
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
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

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        private static void Shutdown()
        {
            _shutdownStarted = true;
            _logFileStream.Close();
            _disposed = true;
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
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
                {
                    LogDebug("Log created.");
                }
                else
                {
                    _logFileStream.Position = _logFileStream.Length;
                    LogDebug("Log reopened.");
                }
            }
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
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

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        private static void WriteLogEntry(string message, LogEntryType entryType)
        {
            if (_disposed)
                return;
            
            DateTime timestamp = DateTime.Now;

            if (UseFileLogging)
                CreateLogFile();

            if (_currentLogTask != null)
            {
                _pendingLogTasks.Enqueue(WriteNextLogEntryAsync(message, entryType, timestamp));
                _currentLogTask.GetAwaiter().OnCompleted(() =>
                {
                    if (_pendingLogTasks.Count > 0)
                        _currentLogTask = _pendingLogTasks.Dequeue();
                    else
                        _currentLogTask = null;
                });
            }
            else
                _currentLogTask = WriteNextLogEntryAsync(message, entryType, timestamp);

            RaiseLogDebugAdded(null, new LogEntry(message, entryType, timestamp));
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        private static async Task WriteNextLogEntryAsync(string message, LogEntryType entryType, DateTime timestamp)
        {
            string logmessage = CreateMessageString(message, entryType, timestamp);
            Console.Write(logmessage);
            if (UseFileLogging)
            {
                byte[] logmessagebytes = Encoding.UTF8.GetBytes(logmessage); ;
                try
                {
                    _logEntries.Enqueue(logmessagebytes);
                    while (_logEntries.Count > 0)
                    {
                        logmessagebytes = _logEntries.Dequeue();
                        await _logFileStream.WriteAsync(logmessagebytes, 0, logmessagebytes.Length);
                    }
                    await _logFileStream.FlushAsync();

                    if (_shutdownStarted)
                    {
                        _logFileStream.Close();
                        _disposed = true;
                    }
                }
                catch (Exception ex)
                {
                    if (!(_disposed || _shutdownStarted))
                    {
                        Console.WriteLine(ex.ToString());
                        _logEntries.Enqueue(logmessagebytes);
                    }
                }
            }

            _currentLogTask = null;
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        public static string[] GetAllLines() => File.ReadAllLines(ExceptionLogFilePath);

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        public static void LogDebug(string message)
        {
            if (_includeDebugInformation)
                WriteLogEntry(message, LogEntryType.Info);
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        public static void LogDebugBool(string message, bool value, string trueString, string falseString)
        {
            if (_includeDebugInformation)
                WriteLogEntry(message.Replace("%1", value ? trueString : falseString), LogEntryType.Info);
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        public static void LogWarning(string message) => WriteLogEntry(message, LogEntryType.Warning);

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        public static void LogGenericError(string message) => WriteLogEntry(message, LogEntryType.GenericError);

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        public static void LogCriticalError(string message) => WriteLogEntry(message, LogEntryType.CriticalError);

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        public static void LogGenericError(Exception exception) => WriteLogEntry(exception.Message, LogEntryType.GenericError);

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        public static void LogCriticalError(Exception exception) => WriteLogEntry(exception.Message, LogEntryType.CriticalError);

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        public static int RunSafe(LogMessage message, Action action)
        {
            try
            {
                action.Invoke();
                return VSConstants.S_OK;
            }
            catch (Exception ex)
            {
                if (message.IsEmptyMessage)
                    message.Message = "<Untraced>";
                WriteLogEntry(string.Format("{0}: {1}", message.Message, ex.Message), message.EntryType);
                return ex.HResult;
            }
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        public static int RunSafe(LogMessage message, Func<bool> action)
        {
            try
            {
                bool result = action.Invoke();
                return result ? VSConstants.S_OK : VSConstants.S_FALSE;
            }
            catch (Exception ex)
            {
                if (message.IsEmptyMessage)
                    message.Message = "<Untraced>";
                WriteLogEntry(string.Format("{0}: {1}", message.Message, ex.Message), message.EntryType);
                return ex.HResult;
            }
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        public static string GetHRToMessage(int hr)
        {
            return Marshal.GetExceptionForHR(hr).Message;
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        public struct LogMessage
        {
            public string Message { get; set; }
            public LogEntryType EntryType { get; }
            public bool IsEmptyMessage => string.IsNullOrEmpty(Message);

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            public LogMessage(string message = "", LogEntryType entryType = LogEntryType.GenericError)
            {
                Message = message;
                EntryType = entryType;
            }
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        public class LogEntry
        {
            public DateTime TimeStamp { get; }
            public string TimeStampText => TimeStamp.ToString("yyyy-MM-dd hh:mm:ss.fff");
            public string Message { get; protected set; }
            public LogEntryType EntryType { get; protected set; }
            public string EntryTypeText => Enum.GetName(typeof(LogEntryType), EntryType);

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            public LogEntry(string message, LogEntryType entryType, DateTime timeStamp)
            {
                Message = message;
                EntryType = entryType;
                TimeStamp = timeStamp;
            }

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            public override bool Equals(object obj)
            {
                return obj is LogEntry entry &&
                       Message == entry.Message;
            }

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            public override int GetHashCode()
            {
                return 460171812 + EqualityComparer<string>.Default.GetHashCode(Message);
            }

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            public override string ToString()
            {
                return CreateMessageString(Message, EntryType, TimeStamp);
            }
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        public class LogDebugAddedEventArgs
        {
            public LogEntry Entry { get; }

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            public LogDebugAddedEventArgs(LogEntry entry)
            {
                Entry = entry;
            }

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            public override bool Equals(object obj)
            {
                if (!(obj is LogDebugAddedEventArgs))
                    return false;

                LogDebugAddedEventArgs args = (LogDebugAddedEventArgs) obj;
                return Entry == args.Entry;
            }

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            public override int GetHashCode()
            {
                return 1970138155 + EqualityComparer<LogEntry>.Default.GetHashCode(Entry);
            }
        }
    }
}
