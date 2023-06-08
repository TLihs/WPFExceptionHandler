using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using static WPFExceptionHandler.ExceptionManagement;

namespace ExceptionTest
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private static int _testCount = 0;
        
        public ObservableCollection<string> TestMessages { get; }
        public string StringifiedTestMessages => string.Join("", TestMessages);
        
        public MainWindow()
        {
            TestMessages = new ObservableCollection<string>();
            //TestMessages.CollectionChanged += (s, e) => OnPropertyChanged("StringifiedTestMessages");

            InitializeComponent();
        }

        private void Button_ThrowException_Click(object sender, RoutedEventArgs e)
        {
            LogMessage message = new LogMessage("TestMessage");
            for (int i = 0; i < 5000; i++)
            {
                LogDebug("test ist test test ist test ist test ist test test ist test ist test ist test test ist test ist test ist test test ist test ist test ist test test ist test ist test ist test test ist test ist test");
            }

            _testCount++;
        }
    }
}
