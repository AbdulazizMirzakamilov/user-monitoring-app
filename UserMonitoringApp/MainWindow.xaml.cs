using Npgsql;
using System.Configuration;
using System.Windows;
using System.Windows.Controls;
using UserMonitoringApp.Models;
using UserMonitoringApp.Services;

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
            if (dateFrom.SelectedDate == null || dateTo.SelectedDate == null)
            {
                MessageBox.Show("Выберите даты");
                return;
            }

            if (dateFrom.SelectedDate > dateTo.SelectedDate)
            {
                MessageBox.Show("Дата 'с' не может быть больше даты 'по'");
                return;
            }

            if (!int.TryParse(thresholdBox.Text, out int threshold))
            {
                MessageBox.Show("Порог должен быть числом");
                return;
            }

            var service = new MonitoringService();

            var data = service.GetAnomalyReport(
                dateFrom.SelectedDate.Value,
                dateTo.SelectedDate.Value,
                threshold
            );

            dataGrid.ItemsSource = null;
            _anomalyData = data;
            dataGrid.ItemsSource = _anomalyData;
        }

        private void LoadIpReport_Click(object sender, RoutedEventArgs e)
        {
            if (ipDateFrom.SelectedDate == null || ipDateTo.SelectedDate == null)
            {
                MessageBox.Show("Выберите даты");
                return;
            }

            if (ipDateFrom.SelectedDate > ipDateTo.SelectedDate)
            {
                MessageBox.Show("Дата 'с' не может быть больше даты 'по'");
                return;
            }

            var service = new MonitoringService();

            var data = service.GetIpReport(
                ipDateFrom.SelectedDate.Value,
                ipDateTo.SelectedDate.Value
            );

            ipGrid.ItemsSource = null;
            ipGrid.ItemsSource = data;
        }

        private void LoadContinuousReport_Click(object sender, RoutedEventArgs e)
        {
            if (contDateFrom.SelectedDate == null || contDateTo.SelectedDate == null)
            {
                MessageBox.Show("Выберите даты");
                return;
            }

            if (contDateFrom.SelectedDate > contDateTo.SelectedDate)
            {
                MessageBox.Show("Дата 'с' не может быть больше даты 'по'");
                return;
            }

            if (!int.TryParse(contThresholdBox.Text, out int threshold))
            {
                MessageBox.Show("Порог должен быть числом");
                return;
            }

            var service = new MonitoringService();

            var data = service.GetContinuousWorkReport(
                contDateFrom.SelectedDate.Value,
                contDateTo.SelectedDate.Value,
                threshold
            );

            contGrid.ItemsSource = null;
            contGrid.ItemsSource = data;
        }

        private List<AnomalyReportItem> _anomalyData;

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_anomalyData == null) return;

            var text = searchBox.Text;

            var filtered = _anomalyData
                .Where(x =>
                    x.Username.Contains(text, StringComparison.OrdinalIgnoreCase) ||
                    x.FullName.Contains(text, StringComparison.OrdinalIgnoreCase))
                .ToList();

            dataGrid.ItemsSource = filtered;
        }
    }
}