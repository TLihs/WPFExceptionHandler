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
            CreateExceptionManagement(this, AppDomain.CurrentDomain);
        }

        protected override void OnStartup(StartupEventArgs e)
        {
            LogDebug("Test");
            //throw new Exception("TestException");
        }
    }
}
