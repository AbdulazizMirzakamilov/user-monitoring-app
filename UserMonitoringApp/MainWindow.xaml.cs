using Npgsql;
using System.Configuration;
using System.Windows;
using System.Windows.Controls;
using UserMonitoringApp.Models;
using UserMonitoringApp.Services;
using ClosedXML.Excel;
using Microsoft.Win32;

namespace UserMonitoringApp
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            dateFrom.SelectedDate = DateTime.Now.AddDays(-7);
            dateTo.SelectedDate = DateTime.Now;

            ipDateFrom.SelectedDate = DateTime.Now.AddDays(-7);
            ipDateTo.SelectedDate = DateTime.Now;

            contDateFrom.SelectedDate = DateTime.Now.AddDays(-7);
            contDateTo.SelectedDate = DateTime.Now;
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

            if (data.Count == 0)
            {
                MessageBox.Show("Данные не найдены");
            }

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

            if (data.Count == 0)
            {
                MessageBox.Show("Данные не найдены");
            }

            ipGrid.ItemsSource = null;
            _ipData = data;
            ipGrid.ItemsSource = _ipData;
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

            if (data.Count == 0)
            {
                MessageBox.Show("Данные не найдены");
            }

            contGrid.ItemsSource = null;
            _contData = data;
            contGrid.ItemsSource = _contData;
        }

        private List<AnomalyReportItem> _anomalyData;
        private List<IpReportItem> _ipData;
        private List<ContinuousWorkItem> _contData;

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

        private void IpSearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_ipData == null) return;

            var text = ipSearchBox.Text;

            var filtered = _ipData
                .Where(x =>
                    x.Username.Contains(text, StringComparison.OrdinalIgnoreCase) ||
                    x.FullName.Contains(text, StringComparison.OrdinalIgnoreCase))
                .ToList();

            ipGrid.ItemsSource = filtered;
        }

        private void ExportToExcel<T>(List<T> data, string fileName)
        {
            var dialog = new SaveFileDialog
            {
                Filter = "Excel files (*.xlsx)|*.xlsx",
                FileName = fileName
            };

            if (dialog.ShowDialog() != true) return;

            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Report");

            var properties = typeof(T).GetProperties();

            // Заголовки
            for (int i = 0; i < properties.Length; i++)
            {
                worksheet.Cell(1, i + 1).Value = properties[i].Name;
            }

            // Данные
            for (int row = 0; row < data.Count; row++)
            {
                for (int col = 0; col < properties.Length; col++)
                {
                    var value = properties[col].GetValue(data[row]);
                    worksheet.Cell(row + 2, col + 1).Value = value?.ToString();
                }
            }

            workbook.SaveAs(dialog.FileName);
        }

        private void ExportAnomaly_Click(object sender, RoutedEventArgs e)
        {
            if (_anomalyData == null || _anomalyData.Count == 0)
            {
                MessageBox.Show("Нет данных для экспорта");
                return;
            }

            ExportToExcel(_anomalyData, "AnomalyReport.xlsx");
        }
        private void ExportIp_Click(object sender, RoutedEventArgs e)
        {
            if (_ipData == null || _ipData.Count == 0)
            {
                MessageBox.Show("Нет данных для экспорта");
                return;
            }

            ExportToExcel(_ipData, "IpReport.xlsx");
        }
        private void ExportContinuous_Click(object sender, RoutedEventArgs e)
        {
            if (_contData == null || _contData.Count == 0)
            {
                MessageBox.Show("Нет данных для экспорта");
                return;
            }

            ExportToExcel(_contData, "ContinuousReport.xlsx");
        }
    }
}