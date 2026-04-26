using ClosedXML.Excel;
using Microsoft.Win32;
using System.Windows;
using System.Windows.Controls;
using UserMonitoringApp.Models;
using UserMonitoringApp.Services;

namespace UserMonitoringApp
{
    public partial class MainWindow : Window
    {
        private List<AnomalyReportItem> _anomalyData;
        private List<IpReportItem> _ipData;
        private List<ContinuousWorkItem> _contData;

        public MainWindow()
        {
            InitializeComponent();

            // Инициализация дат по умолчанию
            var startDate = DateTime.Now.AddDays(-30);
            var endDate = DateTime.Now;

            dateFrom.SelectedDate = startDate;
            dateTo.SelectedDate = endDate;

            ipDateFrom.SelectedDate = startDate;
            ipDateTo.SelectedDate = endDate;

            contDateFrom.SelectedDate = startDate;
            contDateTo.SelectedDate = endDate;

            opDateFrom.SelectedDate = startDate;
            opDateTo.SelectedDate = endDate;
        }

        #region Обработчики загрузки данных

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

        private void LoadOperationStats_Click(object sender, RoutedEventArgs e)
        {
            var service = new MonitoringService();
            var data = service.GetOperationStats(
                opDateFrom.SelectedDate.Value,
                opDateTo.SelectedDate.Value
            );

            operationGrid.ItemsSource = data;
        }

        private void LoadDailyActivity_Click(object sender, RoutedEventArgs e)
        {
            if (date.SelectedDate == null) return;

            var service = new MonitoringService();
            var data = service.GetDailyActivity(date.SelectedDate.Value);

            dailyGrid.ItemsSource = data;
        }

        #endregion

        #region Поиск и фильтрация

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_anomalyData == null) return;

            var text = searchBox.Text;
            var filtered = _anomalyData
                .Where(x => x.Username.Contains(text, StringComparison.OrdinalIgnoreCase) ||
                            x.FullName.Contains(text, StringComparison.OrdinalIgnoreCase))
                .ToList();

            dataGrid.ItemsSource = filtered;
        }

        private void IpSearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_ipData == null) return;

            var text = ipSearchBox.Text;
            var filtered = _ipData
                .Where(x => x.Username.Contains(text, StringComparison.OrdinalIgnoreCase) ||
                            x.FullName.Contains(text, StringComparison.OrdinalIgnoreCase))
                .ToList();

            ipGrid.ItemsSource = filtered;
        }

        #endregion

        #region Экспорт в Excel

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

            // Словарь для маппинга заголовков
            var headers = new Dictionary<string, string>
            {
                { "Username", "Пользователь" },
                { "FullName", "ФИО" },
                { "RequestsCount", "Количество запросов" },
                { "IpCount", "Количество IP" },
                { "Date", "Дата" },
                { "DaysCount", "Дней подряд" },
                { "Name", "Операция" },
                { "TotalCount", "Количество" },
                { "TotalRequests", "Всего запросов" }
            };

            // Формирование шапки таблицы
            for (int i = 0; i < properties.Length; i++)
            {
                var propName = properties[i].Name;
                var headerText = headers.ContainsKey(propName) ? headers[propName] : propName;

                var cell = worksheet.Cell(1, i + 1);
                cell.Value = headerText;
                cell.Style.Font.Bold = true;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#EEF1F7");
            }

            // Заполнение данными
            for (int row = 0; row < data.Count; row++)
            {
                for (int col = 0; col < properties.Length; col++)
                {
                    var value = properties[col].GetValue(data[row]);
                    var cell = worksheet.Cell(row + 2, col + 1);

                    if (value is DateTime dt)
                    {
                        cell.Value = dt;
                        cell.Style.DateFormat.Format = "dd.MM.yyyy";
                    }
                    else
                    {
                        cell.Value = value?.ToString() ?? "";
                    }

                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                }
            }

            // Оформление: границы и автоширина
            var range = worksheet.Range(1, 1, data.Count + 1, properties.Length);
            range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            worksheet.Columns().AdjustToContents();

            // Подвал отчета
            int infoRow = data.Count + 3;
            worksheet.Cell(infoRow, 1).Value = "Дата формирования:";
            worksheet.Cell(infoRow, 2).Value = DateTime.Now.ToString("dd.MM.yyyy HH:mm");

            worksheet.Cell(infoRow + 1, 1).Value = "Сформировал:";
            worksheet.Cell(infoRow + 1, 2).Value = Environment.UserName;

            worksheet.Range(infoRow, 1, infoRow + 1, 1).Style.Font.Bold = true;
            worksheet.Columns().AdjustToContents();

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

        #endregion
    }
}