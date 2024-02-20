// .NET 8 Exception Handler Library
// Copyright (c) 2024 Toni Lihs
// Licensed under MIT License

using Microsoft.Win32.SafeHandles;
using System.Diagnostics;
using System.IO;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Threading;

namespace NET8ExceptionHandler
{
    public static class ExceptionManagement
    {
        /// <summary>
        /// Enumeration of types of log entries that are defined
        /// for exception handling and/or debug logging.
        /// </summary>
        public enum LogEntryTypes
        {
            NOT_INITIALIZED,
            DEBUG,
            INFO,
            WARNING,
            GENERIC_ERROR,
            CRITICAL_ERROR
        }

        /// <summary>
        /// 
        /// </summary>
        public enum ExceptionManagementStates
        {
            EMS_INITIALIZED = 0x0001,
            EMS_DISPOSING = 0x0010,
            EMS_DISPOSED = 0x0020,
            EMS_FILEHANDLE_OPENED = 0x0100,
            EMS_ACCESSING_FILEHANDLE = 0x0200,
            EMS_FILEHANDLE_NOT_ACCESSIBLE = 0x0400,
            EMS_MESSAGE_BUFFER_FULL = 0x0800
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        public delegate void LogEntryAddedEventHandler(
            object? sender, LogEntryAddedEventArgs args);

        /// <summary>
        /// 
        /// </summary>
        public static event LogEntryAddedEventHandler? LogEntryAdded = null;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="entry"></param>
        private static void RaiseLogEntryAdded(object? sender, LogMessage entry)
        {
            try
            {
                LogEntryAdded?.Invoke(sender, new LogEntryAddedEventArgs(entry));
            }
            catch (Exception e)
            {
                EHUseRaiseLogEntryAddedEvent = false;
                EHLogGenericError(e);
                EHLogWarning("Deactivated raising of log entry added events.");
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        public delegate void ExceptionCaughtEventHandler(object sender,
            ExceptionCaughtEventArgs args);

        /// <summary>
        /// 
        /// </summary>
        public static event ExceptionCaughtEventHandler? ExceptionCaught = null;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="exception"></param>
        /// <param name="isCritical"></param>
        private static void RaiseExceptionCaught(object? sender, Exception exception,
            bool isHandled = false, bool isCritical = false)
        {
            try
            {
                ExceptionCaught?.Invoke(sender ?? new(), new ExceptionCaughtEventArgs(
                    exception, isHandled, isCritical));
            }
            catch (Exception e)
            {
                EHUseRaiseExceptionCaughtEvent = false;
                EHLogGenericError(e);
                EHLogWarning("Deactivated raising of exception caught events.");
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="message"></param>
        /// <param name="isCritical"></param>
        private static void RaiseExceptionCaught(object sender, string message,
            bool isHandled = false, bool isCritical = false)
        {
            try
            {
                ExceptionCaught?.Invoke(sender, new ExceptionCaughtEventArgs(
                    message, isHandled, isCritical));
            }
            catch (Exception e)
            {
                EHUseRaiseExceptionCaughtEvent = false;
                EHLogGenericError(e);
                EHLogWarning("Deactivated raising of exception caught events.");
            }
        }

        private static readonly char[] _NEW_LINE_CHARS = ['\r', '\n'];

        private static string _exceptionLogPathAlt = string.Empty;
        private static string _exceptionLogPath = string.Empty;
        private static bool _includeDebugInformation = true;
        private static SafeFileHandle? _fileHandle = null;
        private static Exception? _lastCaughtException = null;
        private static ExceptionManagementStates _state = 0;
        private static readonly Queue<LogMessage> _logEntries = new();
        private static readonly Thread _loggingThread;
        private static readonly MemoryStream _buffer = new();
        private static readonly Thread _bufferThread;

        /// <summary>
        /// 
        /// </summary>
        public static string EHExceptionLogFilePath =>
            string.IsNullOrWhiteSpace(_exceptionLogPathAlt) ?
            Path.Combine(_exceptionLogPath, DateTime.Now.ToString("yyyy-MM-dd") + ".log") : 
            _exceptionLogPathAlt;
        /// <summary>
        /// 
        /// </summary>
        public static bool EHUseFileLogging { get; set; }
        public static bool EHUseRaiseLogEntryAddedEvent { get; set; } = true;
        public static bool EHUseRaiseExceptionCaughtEvent { get; set; } = true;

        /// <summary>
        /// 
        /// </summary>
        static ExceptionManagement()
        {
            ThreadStart threadStart = new(RefreshBufferContinuously);
            _bufferThread = new(threadStart)
            {
                IsBackground = true,
                Name = "Exception Management Buffer Refresh",
                Priority = ThreadPriority.BelowNormal
            };

            threadStart = new(FlushBufferContinuously);
            _loggingThread = new(threadStart)
            {
                IsBackground = true,
                Name = "Exception Management File Flush",
                Priority = ThreadPriority.BelowNormal
            };
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="app"></param>
        /// <param name="domain"></param>
        /// <param name="includeDebugInformation"></param>
        /// <param name="useFileLogging"></param>
        public static bool CreateExceptionManagement(Application app, AppDomain domain,
            bool includeDebugInformation = true, bool useFileLogging = false)
        {
            if ((_state & ExceptionManagementStates.EMS_INITIALIZED) ==
                ExceptionManagementStates.EMS_INITIALIZED)
            {
                EHLogWarning(
                    "ExceptionManagement::CreateExceptionManagement(..) - " +
                    "ExceptionManagement already initialized.");
                return false;
            }

            EHUseFileLogging = useFileLogging;
            _includeDebugInformation = includeDebugInformation;
            string? appname = System.Reflection.Assembly.GetEntryAssembly()?.GetName()?.Name;
            if (appname == null)
            {
                EHLogGenericError(
                    "ExceptionManagement::CreateExceptionManagement(..) - " +
                    "App name could not be determined. Logging disabled.");
                return false;
            }

            _exceptionLogPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                appname, "Log");

            TaskScheduler.UnobservedTaskException += new EventHandler<UnobservedTaskExceptionEventArgs>((s, e) =>
            {
                if (e.Observed)
                    return;

                try
                {
                    EHLogGenericError(e.Exception);
                    e.SetObserved();
                    _lastCaughtException = e.Exception;
                    if (EHUseRaiseExceptionCaughtEvent)
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
                    if (EHUseRaiseExceptionCaughtEvent)
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
                    if (EHUseRaiseExceptionCaughtEvent)
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

                if (e == null)
                    return;

                try
                {
                    if (e.ExceptionObject != null)
                    {
                        if (e.IsTerminating)
                            EHLogCriticalError((Exception)e.ExceptionObject);
                        else
                            EHLogGenericError((Exception)e.ExceptionObject);
                        _lastCaughtException = (Exception)e.ExceptionObject;
                    }
                    exceptioncaught = true;
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

                if (!exceptioncaught && EHUseRaiseExceptionCaughtEvent)
                {
                    // See TODO above ...
                    // EHLogCriticalError("Unknown critical exception occured.");
                    RaiseExceptionCaught(s, "Unknown critical exception occured.");
                }
                else if (e.IsTerminating)
                {
                    EHLogCriticalError("App is terminating.");
                    if (e.ExceptionObject != null && EHUseRaiseExceptionCaughtEvent)
                        RaiseExceptionCaught(s, (Exception)e.ExceptionObject, false, true);
                }
                else
                {
                    if (e.ExceptionObject != null && EHUseRaiseExceptionCaughtEvent)
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
                    if (EHUseRaiseExceptionCaughtEvent)
                        RaiseExceptionCaught(s, e.Exception);
                }
                catch (Exception ex)
                {
                    EHLogCriticalError("Error occured while logging exception: " +
                        ex.Message);
                }
            });

            app.Exit += new ExitEventHandler((s, e) =>
            {
                if (e.ApplicationExitCode != 0)
                    EHLogCriticalError("Application exited with fault " +
                        "(exit code = " + e.ApplicationExitCode + ").");
                else
                    EHLogDebug("Application exited without any fault.");

                _state |= ExceptionManagementStates.EMS_DISPOSING;
                _state ^= ExceptionManagementStates.EMS_INITIALIZED;

                _bufferThread.Join();
                _loggingThread.Join();

                CloseFileHandle();
            });

            if (app != Application.Current)
                Application.Current.Exit += new ExitEventHandler((s, e) =>
                {
                    if (e.ApplicationExitCode != 0)
                        EHLogCriticalError(
                            "ExceptionManagement exited with fault " +
                            "(exit code = " + e.ApplicationExitCode + ").");
                    else
                        EHLogDebug("ExceptionManagement exited without any fault.");
                });

            _state |= ExceptionManagementStates.EMS_INITIALIZED;

            _bufferThread.Start();
            _loggingThread.Start();

            return true;
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
            if ((_state & ExceptionManagementStates.EMS_FILEHANDLE_OPENED) ==
                ExceptionManagementStates.EMS_FILEHANDLE_OPENED)
                return;

            if ((_state & ExceptionManagementStates.EMS_FILEHANDLE_NOT_ACCESSIBLE) ==
                ExceptionManagementStates.EMS_FILEHANDLE_NOT_ACCESSIBLE)
                return;

            if ((_state & ExceptionManagementStates.EMS_DISPOSING) ==
                ExceptionManagementStates.EMS_DISPOSING)
                return;

            if ((_state & ExceptionManagementStates.EMS_DISPOSED) ==
                ExceptionManagementStates.EMS_DISPOSED)
                return;

            string logpath = EHExceptionLogFilePath;
            bool newlog = false;

            FileInfo logfileinfo;
            try
            {
                logfileinfo = new FileInfo(logpath);
                if (!File.Exists(logpath))
                {
                    if (logfileinfo.DirectoryName == null)
                        throw new DirectoryNotFoundException(
                            "ExceptionManagement::[static]CreateLogFile() - " +
                            $"Failed to get directory name from path '{logpath}'");
                    Directory.CreateDirectory(logfileinfo.DirectoryName);
                    newlog = true;
                }
            }
            catch (Exception ex)
            {
                _state |= ExceptionManagementStates.EMS_FILEHANDLE_NOT_ACCESSIBLE;
                EHUseFileLogging = false;
                Debug.Print(ex.Message);
                if (Console.IsOutputRedirected)
                    Console.WriteLine(ex.Message);
                return;
            }

            try
            {
                _fileHandle?.Close();
                _fileHandle?.Dispose();

                _fileHandle = File.OpenHandle(logfileinfo.FullName, FileMode.OpenOrCreate,
                    FileAccess.ReadWrite, FileShare.Read, FileOptions.WriteThrough);

                if (newlog)
                    EHLogDebug("--------------------Log created---------------------");
                else
                    EHLogDebug("--------------------Log reopened--------------------");
                _state |= ExceptionManagementStates.EMS_FILEHANDLE_OPENED;
            }
            catch (Exception ex)
            {
                _state |= ExceptionManagementStates.EMS_FILEHANDLE_NOT_ACCESSIBLE;
                EHUseFileLogging = false;
                Console.WriteLine(ex.Message);
                return;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private static void CloseFileHandle()
        {
            if ((_state & ExceptionManagementStates.EMS_FILEHANDLE_OPENED) ==
                ExceptionManagementStates.EMS_FILEHANDLE_OPENED)
            {
                _fileHandle?.Close();
                _fileHandle?.Dispose();
                _state |= ExceptionManagementStates.EMS_DISPOSED;
                _state ^= ExceptionManagementStates.EMS_DISPOSING;
                _state ^= ExceptionManagementStates.EMS_FILEHANDLE_OPENED;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        /// <param name="entryType"></param>
        /// <param name="timeStamp"></param>
        /// <returns></returns>
        private static string CreateMessageString(string? threadName,
            string message, LogEntryTypes entryType, DateTime timeStamp)
        {
            string? logentrytype = Enum.GetName(typeof(LogEntryTypes), entryType);
            if (string.IsNullOrEmpty(threadName))
                return $"{timeStamp:yyyy-MM-dd hh:mm:ss.fff}: " +
                    $"({logentrytype}) {message}\r\n";
            else
                return $"{timeStamp:yyyy-MM-dd hh:mm:ss.fff}: " +
                    $"[{threadName}] ({logentrytype}) {message}\r\n";
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ex"></param>
        private static void ProvideExceptionInfoToListeners(Exception ex)
        {
            Debug.Print(FormatException(ex));
            if (Console.IsOutputRedirected)
                Console.WriteLine(FormatException(ex));
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        /// <param name="entryType"></param>
        private static void WriteLogEntry(string message, LogEntryTypes entryType)
        {
            if ((_state & ExceptionManagementStates.EMS_DISPOSED) ==
                ExceptionManagementStates.EMS_DISPOSED)
            {
                Debug.Print("ExceptionManagement already disposed");
                if (Console.IsOutputRedirected)
                    Console.WriteLine("ExceptionManagement already disposed");
                return;
            }

            LogMessage logmessage = new(DateTime.Now,
                Thread.CurrentThread.Name, message, entryType);

            lock (_logEntries)
                _logEntries.Enqueue(logmessage);

            if (EHUseRaiseLogEntryAddedEvent)
                RaiseLogEntryAdded(null, logmessage);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        /// <param name="entryType"></param>
        /// <param name="timestamp"></param>
        /// <returns></returns>
        private static void WriteNextLogEntry()
        {
            if ((_state & ExceptionManagementStates.EMS_FILEHANDLE_OPENED) ==
                ExceptionManagementStates.EMS_FILEHANDLE_OPENED)
            {
                _state |= ExceptionManagementStates.EMS_ACCESSING_FILEHANDLE;

                try
                {
                    byte[] buffer;
                    int length = 0;
                    lock (_buffer)
                    {
                        length = (int)_buffer.Length;
                        buffer = new byte[length];
                        _buffer.Position = 0;
                        _buffer.Read(buffer, 0, length);
                        _buffer.SetLength(0);
                        _buffer.Capacity = 0;
                    }

                    if (length == 0)
                        return;

                    string message = Encoding.UTF8.GetString(buffer).TrimEnd(_NEW_LINE_CHARS);
                    Debug.Print(message);
                    if (Console.IsOutputRedirected)
                        Console.WriteLine(message);

                    if (_fileHandle == null || _fileHandle.IsInvalid || _fileHandle.IsClosed)
                        throw new ArgumentNullException("File stream is null");
                    using FileStream stream = new(_fileHandle, FileAccess.Write);
                    stream.Write(buffer, 0, length);
                    stream.Flush();
                    _state ^= ExceptionManagementStates.EMS_ACCESSING_FILEHANDLE;
                }
                catch (Exception ex)
                {
                    ProvideExceptionInfoToListeners(ex);
                    _state ^= ExceptionManagementStates.EMS_ACCESSING_FILEHANDLE;
                    _state |= ExceptionManagementStates.EMS_DISPOSING;

                    throw;
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private static void FlushBufferContinuously()
        {
            try
            {
                while ((_state & ExceptionManagementStates.EMS_DISPOSED) !=
                    ExceptionManagementStates.EMS_DISPOSED)
                {
                    if (EHUseFileLogging && (_state &
                        ExceptionManagementStates.EMS_FILEHANDLE_OPENED) !=
                        ExceptionManagementStates.EMS_FILEHANDLE_OPENED)
                        CreateLogFile();

                    if ((_state & ExceptionManagementStates.EMS_ACCESSING_FILEHANDLE) !=
                        ExceptionManagementStates.EMS_ACCESSING_FILEHANDLE &&
                        (_state & ExceptionManagementStates.EMS_FILEHANDLE_OPENED) ==
                        ExceptionManagementStates.EMS_FILEHANDLE_OPENED)
                    {
                        WriteNextLogEntry();
                    }
                    else
                    {
                        Thread.Sleep(5);
                    }

                    if ((_state & ExceptionManagementStates.EMS_DISPOSING) ==
                        ExceptionManagementStates.EMS_DISPOSING)
                    {
                        long length = 0;

                        if ((_state & ExceptionManagementStates.EMS_ACCESSING_FILEHANDLE) !=
                            ExceptionManagementStates.EMS_ACCESSING_FILEHANDLE &&
                            (_state & ExceptionManagementStates.EMS_FILEHANDLE_OPENED) ==
                            ExceptionManagementStates.EMS_FILEHANDLE_OPENED &&
                            (_state & ExceptionManagementStates.EMS_MESSAGE_BUFFER_FULL) !=
                            ExceptionManagementStates.EMS_MESSAGE_BUFFER_FULL)
                        {
                            lock (_buffer)
                                length = _buffer.Length;
                        }

                        if (length == 0)
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                ProvideExceptionInfoToListeners(ex);
                throw;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private static void RefreshBufferContinuously()
        {
            try
            {
                while ((_state & ExceptionManagementStates.EMS_DISPOSED) !=
                    ExceptionManagementStates.EMS_DISPOSED &&
                    (_state & ExceptionManagementStates.EMS_DISPOSING) !=
                    ExceptionManagementStates.EMS_DISPOSING)
                {
                    if ((_state & ExceptionManagementStates.EMS_FILEHANDLE_OPENED) ==
                        ExceptionManagementStates.EMS_FILEHANDLE_OPENED &&
                        (_state & ExceptionManagementStates.EMS_MESSAGE_BUFFER_FULL) !=
                        ExceptionManagementStates.EMS_MESSAGE_BUFFER_FULL)
                    {
                        LogMessage? logmessage = null;
                        lock (_logEntries)
                            if (_logEntries.Count > 0)
                                logmessage = _logEntries.Dequeue();

                        if (logmessage != null)
                        {
                            string? logtext = logmessage.ToString();
                            if (logtext != null)
                            {
                                byte[] logmessagebytes = Encoding.UTF8.GetBytes(logtext);
                                lock (_buffer)
                                    _buffer.Write(logmessagebytes, 0, logmessagebytes.Length);
                            }
                        }
                        else
                        {
                            Thread.Sleep(5);
                        }
                    }
                    else
                    {
                        Thread.Sleep(5);
                    }
                }
            }
            catch (Exception ex)
            {
                ProvideExceptionInfoToListeners(ex);
                throw;
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
        public static void EHLogDebug(string message)
        {
            if (_includeDebugInformation)
                    WriteLogEntry(message, LogEntryTypes.DEBUG);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        /// <param name="formatParameters"></param>
        public static void EHLogInfo(string message)
        {
            WriteLogEntry(message, LogEntryTypes.INFO);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        /// <param name="formatParameters"></param>
        public static void EHLogWarning(string message)
        {
            WriteLogEntry(message, LogEntryTypes.WARNING);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="exception"></param>
        public static void EHLogGenericError(Exception exception)
            => WriteLogEntry(FormatException(exception), LogEntryTypes.GENERIC_ERROR);

        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        /// <param name="formatParameters"></param>
        public static void EHLogGenericError(string message)
        {
            WriteLogEntry(message, LogEntryTypes.GENERIC_ERROR);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="exception"></param>
        public static void EHLogCriticalError(Exception exception)
            => WriteLogEntry(FormatException(exception), LogEntryTypes.CRITICAL_ERROR);

        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        /// <param name="formatParameters"></param>
        public static void EHLogCriticalError(string message)
        {
            WriteLogEntry(message, LogEntryTypes.CRITICAL_ERROR);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="type"></param>
        /// <param name="message"></param>
        /// <param name="owner"></param>
        /// <param name="formatParameters"></param>
        public static void EHMsgBox(LogEntryTypes type, string message,
            Window? owner = null)
        {
            WriteLogEntry(message, type);
            string title = "Info";
            MessageBoxImage severityimage = MessageBoxImage.Information;
            switch (type)
            {
                case LogEntryTypes.WARNING:
                    title = "Warning";
                    severityimage = MessageBoxImage.Warning;
                    break;
                case LogEntryTypes.GENERIC_ERROR | LogEntryTypes.CRITICAL_ERROR:
                    title = "Error";
                    severityimage = MessageBoxImage.Error;
                    break;
            }

            if (owner != null)
                MessageBox.Show(owner, message, title,
                    MessageBoxButton.OK, severityimage);
            else
                MessageBox.Show(message, title,
                    MessageBoxButton.OK, severityimage);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        /// <param name="owner"></param>
        /// <param name="formatParameters"></param>
        /// <returns></returns>
        public static MessageBoxResult EHMsgBoxYesNo(string message,
            Window? owner = null)
        {
            WriteLogEntry(message, LogEntryTypes.INFO);
            if (owner != null)
                return MessageBox.Show(owner, message, "Question",
                    MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes,
                    MessageBoxOptions.DefaultDesktopOnly);
            else
                return MessageBox.Show(message, "Question", MessageBoxButton.YesNo,
                    MessageBoxImage.Question, MessageBoxResult.Yes,
                    MessageBoxOptions.DefaultDesktopOnly);
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
                return 0;
            }
            catch (Exception ex)
            {
                if (exceptionMessage.IsEmptyMessage)
                    exceptionMessage.Message = "<Untraced>";
                WriteLogEntry($"{exceptionMessage.Message}: {ex.Message}",
                    exceptionMessage.EntryType);
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
                return result ? 0 : 1;
            }
            catch (Exception ex)
            {
                if (exceptionMessage.IsEmptyMessage)
                    exceptionMessage.Message = "<Untraced>";
                WriteLogEntry($"{exceptionMessage.Message}: {ex.Message}",
                    exceptionMessage.EntryType);
                return ex.HResult;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ex"></param>
        /// <param name="functionName"></param>
        /// <param name="propertyName"></param>
        /// <param name="message"></param>
        /// <param name="criticalError"></param>
        public static void SafeLogException(Exception ex, string functionName,
            string propertyName, string message, bool criticalError)
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
        public static string? GetHRToMessage(int hr)
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
        /// <param name="timestamp"></param>
        /// <param name="threadName"></param>
        /// <param name="message"></param>
        /// <param name="entryType"></param>
        public struct LogMessage(DateTime timestamp, string? threadName,
            string message, LogEntryTypes entryType = LogEntryTypes.NOT_INITIALIZED)
        {
            /// <summary>
            /// 
            /// </summary>
            public DateTime Timestamp = timestamp;
            /// <summary>
            /// 
            /// </summary>
            public string? ThreadName = threadName;
            /// <summary>
            /// 
            /// </summary>
            public string Message = message;
            /// <summary>
            /// 
            /// </summary>
            public LogEntryTypes EntryType = entryType;
            /// <summary>
            /// 
            /// </summary>
            public readonly bool IsEmptyMessage => string.IsNullOrWhiteSpace(Message);

            /// <summary>
            /// 
            /// </summary>
            public override readonly string ToString()
            {
                return CreateMessageString(ThreadName, Message, EntryType, Timestamp);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="entry"></param>
        public class LogEntryAddedEventArgs(LogMessage entry)
        {
            /// <summary>
            /// 
            /// </summary>
            public LogMessage Entry { get; } = entry;

            /// <summary>
            /// 
            /// </summary>
            /// <param name="obj"></param>
            /// <returns></returns>
            public override bool Equals(object? obj)
            {
                if (obj is not LogEntryAddedEventArgs)
                    return false;

                LogEntryAddedEventArgs args = (LogEntryAddedEventArgs)obj;
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
            private readonly string _message;

            /// <summary>
            /// 
            /// </summary>
            public Exception? Exception { get; }
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
            public ExceptionCaughtEventArgs(Exception exception, bool isHandled,
                bool isCritical = false)
            {
                Exception = exception;
                _message = string.Empty;
                IsHandled = isHandled;
                IsCritical = isCritical;
            }

            /// <summary>
            /// 
            /// </summary>
            /// <param name="message"></param>
            /// <param name="isHandled"></param>
            /// <param name="isCritical"></param>
            public ExceptionCaughtEventArgs(string message, bool isHandled,
                bool isCritical = false)
            {
                Exception = null;
                _message = message;
                IsHandled = isHandled;
                IsCritical = isCritical;
            }

            /// <summary>
            /// 
            /// </summary>
            /// <param name="obj"></param>
            /// <returns></returns>
            public override bool Equals(object? obj)
            {
                if (obj is not ExceptionCaughtEventArgs)
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
                return 41 + (Exception != null ?
                    EqualityComparer<Exception>.Default.GetHashCode(Exception) :
                    EqualityComparer<string>.Default.GetHashCode(Message));
            }
        }
    }
}
