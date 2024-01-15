// WPF Exception Handler Library
// Copyright (c) 2023 Toni Lihs
// Licensed under MIT License

using Microsoft.VisualStudio;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using static WPFExceptionHandler.ExceptionManagement;

namespace WPFExceptionHandler
{
    public static class ExceptionManagement
    {
        /// <summary>
        /// Enumeration of types of log entries that are defined for exception handling and/or debug logging.
        /// </summary>
        public enum LogEntryTypes
        {
            LE_NOT_INITIALIZED,
            LE_INFO,
            LE_WARNING,
            LE_ERROR_GENERIC,
            LE_ERROR_CRITICAL
        }

        public enum ExceptionManagementStates
        {
            EMS_INITIALIZED = 0x0001,
            EMS_FILESTREAM_OPENED = 0x0100,
            EMS_ACCESSING_FILESTREAM = 0x0200,
            EMS_FILESTREAM_NOT_ACCESSIBLE = 0x0400,
            EMS_DISPOSING = 0x0010,
            EMS_DISPOSED = 0x0020
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        public delegate void LogDebugAddedEventHandler(object sender, LogDebugAddedEventArgs args);

        /// <summary>
        /// 
        /// </summary>
        public static event LogDebugAddedEventHandler LogDebugAdded;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="entry"></param>
        private static void RaiseLogDebugAdded(object sender, LogMessage entry)
        {
            LogDebugAdded?.Invoke(sender, new LogDebugAddedEventArgs(entry));
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        public delegate void ExceptionCaughtEventHandler(object sender, ExceptionCaughtEventArgs args);

        /// <summary>
        /// 
        /// </summary>
        public static event ExceptionCaughtEventHandler ExceptionCaught;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="exception"></param>
        /// <param name="isCritical"></param>
        private static void RaiseExceptionCaught(object sender, Exception exception, bool isHandled = false, bool isCritical = false)
        {
            ExceptionCaught?.Invoke(sender, new ExceptionCaughtEventArgs(exception, isHandled, isCritical));
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="message"></param>
        /// <param name="isCritical"></param>
        private static void RaiseExceptionCaught(object sender, string message, bool isHandled = false, bool isCritical = false)
        {
            ExceptionCaught?.Invoke(sender, new ExceptionCaughtEventArgs(message, isHandled, isCritical));
        }

        private static readonly char[] _NEW_LINE_CHARS = { '\r', '\n' };
        
        private static string _exceptionLogPathAlt;
        private static string _exceptionLogPath;
        private static bool _includeDebugInformation = true;
        private static FileStream _logFileStream;
        private static Queue<LogMessage> _logEntries = new Queue<LogMessage>();
        private static Exception _lastCaughtException = null;
        private static ExceptionManagementStates _state = 0;
        private static Thread _loggingThread = null;

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
            ThreadStart threadStart = new ThreadStart(WriteLoggingContinuously);
            _loggingThread = new Thread(threadStart)
            {
                IsBackground = true,
                Name = "Exception Management Logging",
                Priority = ThreadPriority.BelowNormal,
            };
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="app"></param>
        /// <param name="domain"></param>
        /// <param name="includeDebugInformation"></param>
        /// <param name="useFileLogging"></param>
        public static void CreateExceptionManagement(Application app, AppDomain domain, bool includeDebugInformation = true, bool useFileLogging = false)
        {
            if ((_state & ExceptionManagementStates.EMS_INITIALIZED) == ExceptionManagementStates.EMS_INITIALIZED)
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
                if (e.Observed)
                    return;

                try
                {
                    EHLogGenericError(e.Exception);
                    e.SetObserved();
                    _lastCaughtException = e.Exception;
                    RaiseExceptionCaught(s, e.Exception, true);
                }
                catch (Exception ex)
                {
                    EHLogCriticalError("Error occured while logging exception: " + ex.Message);
                }
            });

            app.DispatcherUnhandledException += new DispatcherUnhandledExceptionEventHandler((s, e) =>
            {
                if (e.Handled)
                    return;

                try
                {
                    EHLogGenericError(e.Exception);
                    // TODO: Maybe it's not a good idea to set every unhandled exception as handled...
                    e.Handled = true;
                    _lastCaughtException = e.Exception;
                    RaiseExceptionCaught(s, e.Exception, true);
                }
                catch (Exception ex)
                {
                    EHLogCriticalError("Error occured while logging exception: " + ex.Message);
                }
            });

            app.Dispatcher.UnhandledException += new DispatcherUnhandledExceptionEventHandler((s, e) =>
            {
                if (e.Handled)
                    return;

                try
                {
                    EHLogGenericError(e.Exception);
                    // TODO: Maybe it's not a good idea to set every unhandled exception as handled...
                    e.Handled = true;
                    _lastCaughtException = e.Exception;
                    RaiseExceptionCaught(s, e.Exception, true);
                }
                catch (Exception ex)
                {
                    EHLogCriticalError("Error occured while logging exception: " + ex.Message);
                }
            });

            domain.UnhandledException += new UnhandledExceptionEventHandler((s, e) =>
            {
                bool exceptioncaught = false;
                Debug.Print("Domain.UnhandledException");
                
                if (e != null && e.ExceptionObject != null)
                    if (e.ExceptionObject == _lastCaughtException)
                        return;

                    try
                    {
                        if (e.IsTerminating)
                            EHLogCriticalError((Exception)e.ExceptionObject);
                        else
                            EHLogGenericError((Exception)e.ExceptionObject);
                        exceptioncaught = true;
                        _lastCaughtException = (Exception)e.ExceptionObject;
                    }
                    catch
                    {
                        // TODO: Logging might fail and we don't want to create another exception
                        // by trying to log again, but without a try-catch-block. Maybe we should
                        // try to create a specific exception for when logging fails, so it can
                        // be separated from not logging related exceptions (whichever should arise
                        // from this try block ...
                        // EHLogCriticalError("Error occured while logging exception: " + ex.Message);
                    }
                
                if (!exceptioncaught)
                {
                    // See TODO above ...
                    // EHLogCriticalError("Unknown critical exception occured.");
                    RaiseExceptionCaught(s, "Unknown critical exception occured.");
                }
                else if (e.IsTerminating)
                {
                    EHLogCriticalError("App is terminating.");
                    RaiseExceptionCaught(s, (Exception)e.ExceptionObject, false, true);
                }
                else
                {
                    RaiseExceptionCaught(s, (Exception)e.ExceptionObject);
                }
            });

            domain.FirstChanceException += new EventHandler<FirstChanceExceptionEventArgs>((s, e) =>
            {
                if (e.Exception == _lastCaughtException)
                    return;
                Debug.Print("Domain.FirstChanceException");

                try
                {
                    EHLogGenericError(e.Exception);
                    _lastCaughtException = e.Exception;
                    RaiseExceptionCaught(s, e.Exception);
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

                _state |= ExceptionManagementStates.EMS_DISPOSING;
                _state ^= ExceptionManagementStates.EMS_INITIALIZED;

                while (_loggingThread.IsAlive)
                {

                }
            });

            if (app != Application.Current)
                Application.Current.Exit += new ExitEventHandler((s, e) =>
                {
                    if (e.ApplicationExitCode != 0)
                        EHLogCriticalError("ExceptionManagement exited with fault (exit code = " + e.ApplicationExitCode + ").");
                    else
                        EHLogDebug("ExceptionManagement exited without any fault.");
                });

            _state |= ExceptionManagementStates.EMS_INITIALIZED;
            _loggingThread.Start();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
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
                    EHLogGenericError($"Alternative file path invalid ('{filePath}').");
            }
            catch (Exception exception)
            {
                filePathValid = false;
                EHLogGenericError(exception.Message);
                EHLogGenericError($"Alternative file path could not be set ('{filePath}').");
            }

            return filePathValid;
        }

        /// <summary>
        /// 
        /// </summary>
        private static void CreateLogFile()
        {
            if ((_state & ExceptionManagementStates.EMS_FILESTREAM_OPENED) == ExceptionManagementStates.EMS_FILESTREAM_OPENED)
                return;

            if ((_state & ExceptionManagementStates.EMS_FILESTREAM_NOT_ACCESSIBLE) == ExceptionManagementStates.EMS_FILESTREAM_NOT_ACCESSIBLE)
                return;

            if ((_state & ExceptionManagementStates.EMS_DISPOSING) == ExceptionManagementStates.EMS_DISPOSING)
                return;

            if ((_state & ExceptionManagementStates.EMS_DISPOSED) == ExceptionManagementStates.EMS_DISPOSED)
                return;

            string logpath = EHExceptionLogFilePath;
            bool newlog = false;

            FileInfo logfileinfo;
            try
            {
                logfileinfo = new FileInfo(logpath);
                if (!File.Exists(logpath))
                {
                    Directory.CreateDirectory(logfileinfo.DirectoryName);
                    newlog = true;
                }
            }
            catch(Exception ex)
            {
                _state |= ExceptionManagementStates.EMS_FILESTREAM_NOT_ACCESSIBLE;
                EHUseFileLogging = false;
                Debug.Print(ex.Message);
                if (Console.IsOutputRedirected)
                    Console.WriteLine(ex.Message);
                return;
            }

            try
            {
                _logFileStream?.Flush();
                _logFileStream?.Close();
                _logFileStream = null;
                _logFileStream = logfileinfo.Open(FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.Read);

                if (newlog)
                    EHLogDebug("Log created.");
                else
                {
                    _logFileStream.Position = _logFileStream.Length;
                    EHLogDebug("Log reopened.");
                }
                _state |= ExceptionManagementStates.EMS_FILESTREAM_OPENED;
            }
            catch (Exception ex)
            {
                _state |= ExceptionManagementStates.EMS_FILESTREAM_NOT_ACCESSIBLE;
                EHUseFileLogging = false;
                Console.WriteLine(ex.Message);
                return;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private static void CloseFileStream()
        {
            if ((_state & ExceptionManagementStates.EMS_FILESTREAM_OPENED) == ExceptionManagementStates.EMS_FILESTREAM_OPENED)
            {
                _logFileStream.Close();
                _logFileStream = null;
                _state |= ExceptionManagementStates.EMS_DISPOSED;
                _state ^= ExceptionManagementStates.EMS_DISPOSING;
                _state ^= ExceptionManagementStates.EMS_FILESTREAM_OPENED;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        /// <param name="entryType"></param>
        /// <param name="timeStamp"></param>
        /// <returns></returns>
        private static string CreateMessageString(string message, LogEntryTypes entryType, DateTime timeStamp)
        {
            string threadname = Thread.CurrentThread.Name;
            string logentrytype = Enum.GetName(typeof(LogEntryTypes), entryType);
            if (string.IsNullOrEmpty(threadname))
                return $"{timeStamp:yyyy-MM-dd hh:mm:ss.fff}: ({logentrytype}) {message}\r\n";
            else
                return $"{timeStamp:yyyy-MM-dd hh:mm:ss.fff}: [{threadname}] ({logentrytype}) {message}\r\n";
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        /// <param name="entryType"></param>
        private static void WriteLogEntry(string message, LogEntryTypes entryType)
        {
            if ((_state & ExceptionManagementStates.EMS_DISPOSED) == ExceptionManagementStates.EMS_DISPOSED)
            {
                Debug.Print("ExceptionManagement already disposed");
                if (Console.IsOutputRedirected)
                    Console.WriteLine("ExceptionManagement already disposed");
                return;
            }

            LogMessage logmessage = new LogMessage(DateTime.Now, message, entryType);

            lock (_logEntries)
                _logEntries.Enqueue(logmessage);

            RaiseLogDebugAdded(null, logmessage);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        /// <param name="entryType"></param>
        /// <param name="timestamp"></param>
        /// <returns></returns>
        private static void WriteNextLogEntry(LogMessage message)
        {
            string logmessage = message.ToString();
            string logmessagetrimmed = logmessage.TrimEnd(_NEW_LINE_CHARS);
            Debug.Print(logmessagetrimmed);
            if (Console.IsOutputRedirected)
                Console.WriteLine(logmessagetrimmed);

            if ((_state & ExceptionManagementStates.EMS_FILESTREAM_OPENED) == ExceptionManagementStates.EMS_FILESTREAM_OPENED)
            {
                _state |= ExceptionManagementStates.EMS_ACCESSING_FILESTREAM;

                try
                {
                    byte[] buffer = Encoding.UTF8.GetBytes(logmessage);
                    if (_logFileStream == null)
                    {
                        _state ^= ExceptionManagementStates.EMS_FILESTREAM_OPENED;
                        throw new ArgumentNullException("File stream is null");
                    }
                    _logFileStream.Write(buffer, 0, buffer.Length);
                    _logFileStream.Flush();
                }
                catch (Exception ex)
                {
                    Debug.Print(ex.ToString());
                    if (Console.IsOutputRedirected)
                        Console.WriteLine(ex.ToString());
                }
                finally
                {
                    _state ^= ExceptionManagementStates.EMS_ACCESSING_FILESTREAM;
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private static void WriteLoggingContinuously()
        {
            try
            {
                LogMessage logmessage;

                while ((_state & ExceptionManagementStates.EMS_DISPOSED) != ExceptionManagementStates.EMS_DISPOSED)
                {
                    if (EHUseFileLogging && (_state & ExceptionManagementStates.EMS_FILESTREAM_OPENED) != ExceptionManagementStates.EMS_FILESTREAM_OPENED)
                        CreateLogFile();

                    if (!((_state & ExceptionManagementStates.EMS_ACCESSING_FILESTREAM) == ExceptionManagementStates.EMS_ACCESSING_FILESTREAM))
                    {
                        try
                        {
                            lock (_logEntries)
                            {
                                if (_logEntries.Count > 0)
                                    logmessage = _logEntries.Dequeue();
                                else
                                {
                                    Thread.Sleep(10);
                                    continue;
                                }
                            }
                            WriteNextLogEntry(logmessage);
                        }
                        catch (Exception ex)
                        {
                            Debug.Print(ex.ToString());
                            if (Console.IsOutputRedirected)
                                Console.WriteLine(ex.ToString());
                        }
                    }

                    if ((_state & ExceptionManagementStates.EMS_DISPOSING) == ExceptionManagementStates.EMS_DISPOSING)
                        lock (_logEntries)
                            if (_logEntries.Count == 0)
                                CloseFileStream();
                }
            }
            catch (Exception ex)
            {
                Debug.Print(ex.Message);
                if (Console.IsOutputRedirected)
                    Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ex"></param>
        /// <returns></returns>
        private static string FormatException(Exception ex)
        {
            try
            {
                return $"{ex.Message}: {ex.StackTrace} ({ex.Source})";
            }
            catch
            {
                return $"{ex.Message}: {ex.StackTrace} (<unavailable>)";
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private static void CorrectNullOrEmpty(ref object[] formatParameters)
        {
            if (formatParameters == null)
                return;

            for (int index = 0; index < formatParameters.Length; index++)
                if (formatParameters[index] == null)
                    formatParameters[index] = "[null]";
                else if (formatParameters[index] is string parmeters)
                    if (string.IsNullOrEmpty(parmeters))
                        formatParameters[index] = "[empty]";
                    else if (string.IsNullOrWhiteSpace(parmeters))
                        formatParameters[index] = "[whitespace]";
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static string[] GetAllLines() => File.ReadAllLines(EHExceptionLogFilePath);

        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        /// <param name="formatParameters"></param>
        public static void EHLogDebug(string message, params object[] formatParameters)
        {
            if (_includeDebugInformation)
            {
                if (formatParameters != null && formatParameters.Length > 0)
                {
                    CorrectNullOrEmpty(ref formatParameters);
                    WriteLogEntry(string.Format(message, formatParameters), LogEntryTypes.LE_INFO);
                }
                else
                    WriteLogEntry(message, LogEntryTypes.LE_INFO);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        /// <param name="formatParameters"></param>
        public static void EHLogWarning(string message, params object[] formatParameters)
        {
            if (formatParameters != null && formatParameters.Length > 0)
            {
                CorrectNullOrEmpty(ref formatParameters);
                WriteLogEntry(string.Format(message, formatParameters), LogEntryTypes.LE_WARNING);
            }
            else
                WriteLogEntry(message, LogEntryTypes.LE_WARNING);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="exception"></param>
        public static void EHLogGenericError(Exception exception) => WriteLogEntry(FormatException(exception), LogEntryTypes.LE_ERROR_GENERIC);

        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        /// <param name="formatParameters"></param>
        public static void EHLogGenericError(string message, params object[] formatParameters)
        {
            if (formatParameters != null && formatParameters.Length > 0)
            {
                CorrectNullOrEmpty(ref formatParameters);
                WriteLogEntry(string.Format(message, formatParameters), LogEntryTypes.LE_ERROR_GENERIC);
            }
            else
                WriteLogEntry(message, LogEntryTypes.LE_ERROR_GENERIC);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="exception"></param>
        public static void EHLogCriticalError(Exception exception) => WriteLogEntry(FormatException(exception), LogEntryTypes.LE_ERROR_CRITICAL);

        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        /// <param name="formatParameters"></param>
        public static void EHLogCriticalError(string message, params object[] formatParameters)
        {
            if (formatParameters != null && formatParameters.Length > 0)
            {
                CorrectNullOrEmpty(ref formatParameters);
                WriteLogEntry(string.Format(message, formatParameters), LogEntryTypes.LE_ERROR_CRITICAL);
            }
            else
                WriteLogEntry(message, LogEntryTypes.LE_ERROR_CRITICAL);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="type"></param>
        /// <param name="message"></param>
        /// <param name="owner"></param>
        /// <param name="formatParameters"></param>
        public static void EHMsgBox(LogEntryTypes type, string message, Window owner = null, params object[] formatParameters)
        {
            string formattedmessage;
            
            if (formatParameters != null && formatParameters.Length > 0)
            {
                CorrectNullOrEmpty(ref formatParameters);
                formattedmessage = string.Format(message, formatParameters);
            }
            else
                formattedmessage = message;

            WriteLogEntry(formattedmessage, type);
            string title = "Info";
            MessageBoxImage severityimage = MessageBoxImage.Information;
            switch (type)
            {
                case LogEntryTypes.LE_WARNING:
                    title = "Warning";
                    severityimage = MessageBoxImage.Warning;
                    break;
                case LogEntryTypes.LE_ERROR_GENERIC | LogEntryTypes.LE_ERROR_CRITICAL:
                    title = "Error";
                    severityimage = MessageBoxImage.Error;
                    break;
            }

            if (owner != null)
                MessageBox.Show(owner, formattedmessage, title, MessageBoxButton.OK, severityimage);
            else
                MessageBox.Show(formattedmessage, title, MessageBoxButton.OK, severityimage);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        /// <param name="owner"></param>
        /// <param name="formatParameters"></param>
        /// <returns></returns>
        public static MessageBoxResult EHMsgBoxYesNo(string message, Window owner = null, params object[] formatParameters)
        {
            string formattedmessage;

            if (formatParameters != null && formatParameters.Length > 0)
            {
                CorrectNullOrEmpty(ref formatParameters);
                formattedmessage = string.Format(message, formatParameters);
            }
            else
                formattedmessage = message;

            WriteLogEntry(formattedmessage, LogEntryTypes.LE_INFO);
            if (owner != null)
                return MessageBox.Show(owner, formattedmessage, "Question", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes, MessageBoxOptions.DefaultDesktopOnly);
            else
                return MessageBox.Show(formattedmessage, "Question", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes, MessageBoxOptions.DefaultDesktopOnly);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="action"></param>
        /// <param name="exceptionMessage"></param>
        /// <returns></returns>
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
                WriteLogEntry($"{exceptionMessage.Message}: {ex.Message}", exceptionMessage.EntryType);
                return ex.HResult;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="action"></param>
        /// <param name="exceptionMessage"></param>
        /// <returns></returns>
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
                WriteLogEntry($"{exceptionMessage.Message}: {ex.Message}", exceptionMessage.EntryType);
                return ex.HResult;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="hr"></param>
        /// <returns></returns>
        public static void SafeLogException(Exception ex, string functionName, string propertyName, string message, bool criticalError)
        {
            if (criticalError)
            {
                EHLogGenericError($"{functionName}() - ({propertyName}) {message}");
                EHLogGenericError(ex);
            }
            else
                EHLogWarning($"{functionName}() - ({propertyName}) {message}");
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="hr"></param>
        /// <returns></returns>
        public static string GetHRToMessage(int hr)
        {
            try
            {
                return Marshal.GetExceptionForHR(hr)?.Message;
            }
            catch (Exception ex)
            {
                EHLogGenericError(ex);
                return string.Empty;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public struct LogMessage
        {
            /// <summary>
            /// 
            /// </summary>
            public DateTime Timestamp;
            /// <summary>
            /// 
            /// </summary>
            public string Message;
            /// <summary>
            /// 
            /// </summary>
            public LogEntryTypes EntryType;
            /// <summary>
            /// 
            /// </summary>
            public bool IsEmptyMessage => string.IsNullOrWhiteSpace(Message);

            /// <summary>
            /// 
            /// </summary>
            /// <param name="message"></param>
            /// <param name="entryType"></param>
            public LogMessage(DateTime timestamp, string message, LogEntryTypes entryType = LogEntryTypes.LE_NOT_INITIALIZED)
            {
                Timestamp = timestamp;
                Message = message;
                EntryType = entryType;
            }

            /// <summary>
            /// 
            /// </summary>
            public override string ToString()
            {
                return CreateMessageString(Message, EntryType, Timestamp);
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
            public LogMessage Entry { get; }

            /// <summary>
            /// 
            /// </summary>
            /// <param name="entry"></param>
            public LogDebugAddedEventArgs(LogMessage entry)
            {
                Entry = entry;
            }

            /// <summary>
            /// 
            /// </summary>
            /// <param name="obj"></param>
            /// <returns></returns>
            public override bool Equals(object obj)
            {
                if (!(obj is LogDebugAddedEventArgs))
                    return false;

                LogDebugAddedEventArgs args = (LogDebugAddedEventArgs) obj;
                return Entry.Equals(args.Entry);
            }

            /// <summary>
            /// 
            /// </summary>
            /// <returns></returns>
            public override int GetHashCode()
            {
                return 37 + EqualityComparer<LogMessage>.Default.GetHashCode(Entry);
            }
        }

        public class ExceptionCaughtEventArgs
        {
            private string _message;
            
            /// <summary>
            /// 
            /// </summary>
            public Exception Exception { get; }
            /// <summary>
            /// 
            /// </summary>
            public string Message => Exception == null ? _message : Exception.Message;
            /// <summary>
            /// 
            /// </summary>
            public bool IsHandled { get; }
            /// <summary>
            /// 
            /// </summary>
            public bool IsCritical { get; }

            /// <summary>
            /// 
            /// </summary>
            /// <param name="exception"></param>
            /// <param name="isHandled"></param>
            /// <param name="isCritical"></param>
            public ExceptionCaughtEventArgs(Exception exception, bool isHandled, bool isCritical = false)
            {
                Exception = exception;
                IsHandled = isHandled;
                IsCritical = isCritical;
            }

            /// <summary>
            /// 
            /// </summary>
            /// <param name="message"></param>
            /// <param name="isHandled"></param>
            /// <param name="isCritical"></param>
            public ExceptionCaughtEventArgs(string message, bool isHandled, bool isCritical = false)
            {
                _message = message;
                IsHandled = isHandled;
                IsCritical = isCritical;
            }

            /// <summary>
            /// 
            /// </summary>
            /// <param name="obj"></param>
            /// <returns></returns>
            public override bool Equals(object obj)
            {
                if (!(obj is ExceptionCaughtEventArgs))
                    return false;

                ExceptionCaughtEventArgs args = (ExceptionCaughtEventArgs)obj;
                return this == args;
            }

            /// <summary>
            /// 
            /// </summary>
            /// <returns></returns>
            public override int GetHashCode()
            {
                return 41 + (Exception != null ? EqualityComparer<Exception>.Default.GetHashCode(Exception) : EqualityComparer<string>.Default.GetHashCode(Message));
            }
        }
    }
}
