using Npgsql;
using System.Configuration;
using UserMonitoringApp.Services;
using System.Windows;

namespace UserMonitoringApp
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void LoadReport_Click(object sender, RoutedEventArgs e)
        {
            var service = new MonitoringService();

            var data = service.GetAnomalyReport(
                dateFrom.SelectedDate.Value,
                dateTo.SelectedDate.Value,
                int.Parse(thresholdBox.Text)
            );

            dataGrid.ItemsSource = data;
        }
    }
}