using Microsoft.VisualStudio;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

using static WPFExceptionHandler.ExceptionManagement;

namespace ExceptionTest
{
    /// <summary>
    /// Interaktionslogik für "App.xaml"
    /// </summary>
    public partial class App : Application
    {
        private LogDebugAddedEventHandler _logDebugAddedEventHandler = new LogDebugAddedEventHandler((s, e) =>
        {
            RunSafe(new LogMessage("LogDebugAdded failed", LogEntryType.Warning), () =>
            {
                MainWindow mainwindow = (MainWindow)Current.MainWindow;
                if (mainwindow != null && mainwindow.IsVisible)
                    mainwindow.TestMessages.Add(e.Entry.ToString());
            });
        });

        App()
        {
            CreateExceptionManagement(this, AppDomain.CurrentDomain, true);
            UseFileLogging = true;
            LogDebugAdded += _logDebugAddedEventHandler;
        }

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            LogDebug("Debug Test");
            LogWarning("Warning Test");

            LogMessage testExceptionMessage = new LogMessage("LogMessageTest", LogEntryType.Warning);
            Action testAction = () => throw new Exception("TestException thrown");
            int result =  RunSafe(testExceptionMessage, testAction);

            //if (result == VSConstants.S_OK)
            //    App.Current.Shutdown();
        }

        protected override void OnExit(ExitEventArgs e)
        {
            LogDebugAdded -= _logDebugAddedEventHandler;

            base.OnExit(e);
        }
    }
}
