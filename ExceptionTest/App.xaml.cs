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
        App()
        {
            CreateExceptionManagement(this, AppDomain.CurrentDomain, false);
            LogDebugAdded += new LogDebugAddedEventHandler((s, e) =>
            {
                Console.WriteLine("Event received!");
            });
        }

        protected override void OnStartup(StartupEventArgs e)
        {
            LogDebug("Debug Test");
            LogWarning("Warning Test");
        }
    }
}
