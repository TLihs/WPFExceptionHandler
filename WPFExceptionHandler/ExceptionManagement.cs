// WPF Exception Handler Library
// Copyright (c) 2023 Toni Lihs
// Licensed under MIT License

using Microsoft.VisualStudio;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace WPFExceptionHandler
{
    public static class ExceptionManagement
    {
        /// <summary>
        /// Enumeration of types of log entries that are defined for exception handling and/or debug logging.
        /// </summary>
        public enum LogEntryType
        {
            LE_INFO,
            LE_WARNING,
            LE_ERROR_GENERIC,
            LE_ERROR_CRITICAL
        }

        /// <summary>
        /// 
        /// </summary>
        public delegate void LogDebugAddedEventHandler(object sender, LogDebugAddedEventArgs args);

        /// <summary>
        /// 
        /// </summary>
        public static event LogDebugAddedEventHandler LogDebugAdded;

        /// <summary>
        /// 
        /// </summary>
        private static void RaiseLogDebugAdded(object sender, LogEntry entry)
        {
            LogDebugAdded?.Invoke(sender, new LogDebugAddedEventArgs(entry));
        }

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

        /// <summary>
        /// 
        /// </summary>
        public static string EHExceptionLogFilePath => string.IsNullOrWhiteSpace(_exceptionLogPathAlt) ? 
            Path.Combine(_exceptionLogPath, DateTime.Now.ToString("yyyy-MM-dd") + ".log") : _exceptionLogPathAlt;
        /// <summary>
        /// 
        /// </summary>
        public static bool EHUseFileLogging { get; set; }

        /// <summary>
        /// 
        /// </summary>
        static ExceptionManagement()
        {
            Application.Current.Exit += new ExitEventHandler((s, e) =>
            {
                if (e.ApplicationExitCode != 0)
                    EHLogCriticalError("ExceptionManagement exited with fault (exit code = " + e.ApplicationExitCode + ").");
                else
                    EHLogDebug("ExceptionManagement exited without any fault.");
            });
        }

        /// <summary>
        /// 
        /// </summary>
        public static void CreateExceptionManagement(Application app, AppDomain domain, bool includeDebugInformation = true, bool useFileLogging = false)
        {
            if (_initialized)
            {
                EHLogWarning("ExceptionManagement::CreateExceptionManagement(...) - ExceptionManagement already initialized.");
                return;
            }

            EHUseFileLogging = useFileLogging;
            _includeDebugInformation = includeDebugInformation;
            string appname = System.Reflection.Assembly.GetEntryAssembly().GetName().Name;

            _exceptionLogPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), appname, "Log");

            TaskScheduler.UnobservedTaskException += new EventHandler<UnobservedTaskExceptionEventArgs>((s, e) =>
            {
                try
                {
                    EHLogGenericError(e.Exception);
                    e.SetObserved();
                }
                catch (Exception ex)
                {
                    EHLogCriticalError("Error occured while logging exception: " + ex.Message);
                }
            });

            app.DispatcherUnhandledException += new DispatcherUnhandledExceptionEventHandler((s, e) =>
            {
                try
                {
                    EHLogGenericError(e.Exception);
                    e.Handled = true;
                }
                catch (Exception ex)
                {
                    EHLogCriticalError("Error occured while logging exception: " + ex.Message);
                }
            });

            app.Dispatcher.UnhandledException += new DispatcherUnhandledExceptionEventHandler((s, e) =>
            {
                try
                {
                    EHLogGenericError(e.Exception);
                    e.Handled = true;
                }
                catch (Exception ex)
                {
                    EHLogCriticalError("Error occured while logging exception: " + ex.Message);
                }
            });

            domain.UnhandledException += new UnhandledExceptionEventHandler((s, e) =>
            {
                bool exceptioncatched = false;
                
                if (e != null)
                    try
                    {
                        EHLogGenericError((Exception)e.ExceptionObject);
                        exceptioncatched = true;
                    }
                    catch (Exception ex)
                    {
                        EHLogCriticalError("Error occured while logging exception: " + ex.Message);
                    }
                
                if (!exceptioncatched)
                    EHLogCriticalError("Unknown critical exception occured.");

                if (e.IsTerminating)
                    EHLogCriticalError("App is terminating.");
            });

            domain.FirstChanceException += new EventHandler<FirstChanceExceptionEventArgs>((s, e) =>
            {
                try
                {
                    EHLogGenericError(e.Exception);
                    Console.WriteLine(e.Exception.Source.ToString());
                }
                catch (Exception ex)
                {
                    EHLogCriticalError("Error occured while logging exception: " + ex.Message);
                }
            });

            app.Exit += new ExitEventHandler((s, e) =>
            {
                if (e.ApplicationExitCode != 0)
                    EHLogCriticalError("Application exited with fault (exit code = " + e.ApplicationExitCode + ").");
                else
                    EHLogDebug("Application exited without any fault.");

                Shutdown();
            });

            _initialized = true;
        }

        /// <summary>
        /// 
        /// </summary>
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
                    EHLogGenericError(string.Format("Alternative file path invalid ('{0}').", filePath));
            }
            catch (Exception exception)
            {
                filePathValid = false;
                EHLogGenericError(exception.Message);
                EHLogGenericError(string.Format("Alternative file path could not be set ('{0}').", filePath));
            }

            return filePathValid;
        }

        /// <summary>
        /// 
        /// </summary>
        private static void Shutdown()
        {
            _shutdownStarted = true;
            if (EHUseFileLogging)
                _logFileStream.Close();
            _disposed = true;
        }

        /// <summary>
        /// 
        /// </summary>
        private static void CreateLogFile()
        {
            string logpath = EHExceptionLogFilePath;
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
                    EHLogDebug("Log created.");
                }
                else
                {
                    _logFileStream.Position = _logFileStream.Length;
                    EHLogDebug("Log reopened.");
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
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

        /// <summary>
        /// 
        /// </summary>
        private static void WriteLogEntry(string message, LogEntryType entryType)
        {
            if (_disposed)
                return;
            
            DateTime timestamp = DateTime.Now;

            if (EHUseFileLogging)
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

        /// <summary>
        /// 
        /// </summary>
        private static async Task WriteNextLogEntryAsync(string message, LogEntryType entryType, DateTime timestamp)
        {
            string logmessage = CreateMessageString(message, entryType, timestamp);
            Console.Write(logmessage);
            if (EHUseFileLogging)
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
        }

        /// <summary>
        /// 
        /// </summary>
        private static string FormatException(Exception ex)
        {
            string formattedexception;

            try
            {
                formattedexception = string.Format("{0}: {1} ({2})", ex.Message, ex.StackTrace, ex.Source);
            }
            catch (Exception e)
            {
                EHLogGenericError(e);
                formattedexception = ex.Message;
            }
            
            return formattedexception;
        }

        /// <summary>
        /// 
        /// </summary>
        public static string[] GetAllLines() => File.ReadAllLines(EHExceptionLogFilePath);

        /// <summary>
        /// 
        /// </summary>
        public static void EHLogDebug(string message, params string[] formatParameters)
        {
            if (_includeDebugInformation)
            {
                if (formatParameters != null & formatParameters.Length > 0)
                    WriteLogEntry(string.Format(message, formatParameters), LogEntryType.LE_INFO);
                else
                    WriteLogEntry(message, LogEntryType.LE_INFO);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public static void EHLogDebugBool(string message, bool value, string trueString, string falseString)
        {
            if (_includeDebugInformation)
                WriteLogEntry(message.Replace("%1", value ? trueString : falseString), LogEntryType.LE_INFO);
        }

        /// <summary>
        /// 
        /// </summary>
        public static void EHLogWarning(string message, params string[] formatParameters)
        {
            if (formatParameters != null & formatParameters.Length > 0)
                WriteLogEntry(string.Format(message, formatParameters), LogEntryType.LE_WARNING);
            else
                WriteLogEntry(message, LogEntryType.LE_WARNING);
        }

        /// <summary>
        /// 
        /// </summary>
        public static void EHLogGenericError(string message, params string[] formatParameters)
        {
            if (formatParameters != null & formatParameters.Length > 0)
                WriteLogEntry(string.Format(message, formatParameters), LogEntryType.LE_ERROR_GENERIC);
            else
                WriteLogEntry(message, LogEntryType.LE_ERROR_GENERIC);
        }

        /// <summary>
        /// 
        /// </summary>
        public static void EHLogCriticalError(string message, params string[] formatParameters)
        {
            if (formatParameters != null & formatParameters.Length > 0)
                WriteLogEntry(string.Format(message, formatParameters), LogEntryType.LE_ERROR_CRITICAL);
            else
                WriteLogEntry(message, LogEntryType.LE_ERROR_CRITICAL);
        }

        /// <summary>
        /// 
        /// </summary>
        public static void EHLogGenericError(Exception exception) => WriteLogEntry(FormatException(exception), LogEntryType.LE_ERROR_GENERIC);

        /// <summary>
        /// 
        /// </summary>
        public static void EHLogCriticalError(Exception exception) => WriteLogEntry(FormatException(exception), LogEntryType.LE_ERROR_CRITICAL);

        /// <summary>
        /// 
        /// </summary>
        public static int RunSafe(Action action, LogMessage exceptionMessage)
        {
            try
            {
                action.Invoke();
                return VSConstants.S_OK;
            }
            catch (Exception ex)
            {
                if (exceptionMessage.IsEmptyMessage)
                    exceptionMessage.Message = "<Untraced>";
                WriteLogEntry(string.Format("{0}: {1}", exceptionMessage.Message, ex.Message), exceptionMessage.EntryType);
                return ex.HResult;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public static int RunSafe(Func<bool> action, LogMessage exceptionMessage)
        {
            try
            {
                bool result = action.Invoke();
                return result ? VSConstants.S_OK : VSConstants.S_FALSE;
            }
            catch (Exception ex)
            {
                if (exceptionMessage.IsEmptyMessage)
                    exceptionMessage.Message = "<Untraced>";
                WriteLogEntry(string.Format("{0}: {1}", exceptionMessage.Message, ex.Message), exceptionMessage.EntryType);
                return ex.HResult;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public static string GetHRToMessage(int hr)
        {
            return Marshal.GetExceptionForHR(hr).Message;
        }

        /// <summary>
        /// 
        /// </summary>
        public struct LogMessage
        {
            /// <summary>
            /// 
            /// </summary>
            public string Message { get; set; }
            /// <summary>
            /// 
            /// </summary>
            public LogEntryType EntryType { get; }
            /// <summary>
            /// 
            /// </summary>
            public bool IsEmptyMessage => string.IsNullOrEmpty(Message);

            /// <summary>
            /// 
            /// </summary>
            public LogMessage(string message = "", LogEntryType entryType = LogEntryType.LE_ERROR_GENERIC)
            {
                Message = message;
                EntryType = entryType;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public class LogEntry
        {
            /// <summary>
            /// 
            /// </summary>
            public DateTime TimeStamp { get; }
            /// <summary>
            /// 
            /// </summary>
            public string TimeStampText => TimeStamp.ToString("yyyy-MM-dd hh:mm:ss.fff");
            /// <summary>
            /// 
            /// </summary>
            public string Message { get; protected set; }
            /// <summary>
            /// 
            /// </summary>
            public LogEntryType EntryType { get; protected set; }
            /// <summary>
            /// 
            /// </summary>
            public string EntryTypeText => Enum.GetName(typeof(LogEntryType), EntryType);

            /// <summary>
            /// 
            /// </summary>
            public LogEntry(string message, LogEntryType entryType, DateTime timeStamp)
            {
                Message = message;
                EntryType = entryType;
                TimeStamp = timeStamp;
            }

            /// <summary>
            /// 
            /// </summary>
            public override bool Equals(object obj)
            {
                return obj is LogEntry entry &&
                       Message == entry.Message;
            }

            /// <summary>
            /// 
            /// </summary>
            public override int GetHashCode()
            {
                return 460171812 + EqualityComparer<string>.Default.GetHashCode(Message);
            }

            /// <summary>
            /// 
            /// </summary>
            public override string ToString()
            {
                return CreateMessageString(Message, EntryType, TimeStamp);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public class LogDebugAddedEventArgs
        {
            /// <summary>
            /// 
            /// </summary>
            public LogEntry Entry { get; }

            /// <summary>
            /// 
            /// </summary>
            public LogDebugAddedEventArgs(LogEntry entry)
            {
                Entry = entry;
            }

            /// <summary>
            /// 
            /// </summary>
            public override bool Equals(object obj)
            {
                if (!(obj is LogDebugAddedEventArgs))
                    return false;

                LogDebugAddedEventArgs args = (LogDebugAddedEventArgs) obj;
                return Entry == args.Entry;
            }

            /// <summary>
            /// 
            /// </summary>
            public override int GetHashCode()
            {
                return 1970138155 + EqualityComparer<LogEntry>.Default.GetHashCode(Entry);
            }
        }
    }
}
