using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows;

namespace TfsMiner
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnConfig_Click(object sender, RoutedEventArgs e)
        {
            Process.Start("settings.json");
        }

        private Task<string> Mine()
        {
            return Task.Run(() =>
            {
                var miner = new Miner();
                miner.Mine(count => Dispatcher.Invoke(() => progressLabel.Content = count));
                string path = Path.GetFullPath("changes.xlsx");
                miner.ExportToExcel(path);
                return path;
            });
        }

        private async void mine_Click(object sender, RoutedEventArgs e)
        {
            statusLabel.Content = "working...";
            string path = await Mine();
            statusLabel.Content = $"report saved to {path}";
        }
    }
}
