using ClosedXML.Excel;
using Microsoft.Win32;
using System.Windows;
using System.Windows.Controls;
using UserMonitoringApp.Models;
using UserMonitoringApp.Services;
using System.ComponentModel;

namespace UserMonitoringApp
{
    public partial class MainWindow : Window
    {
        private readonly MonitoringService _monitoringService = new MonitoringService();

        // Списки для хранения загруженных данных
        private List<AnomalyReportItem> _anomalyData;
        private List<IpReportItem> _ipData;
        private List<ContinuousWorkItem> _contData;

        public MainWindow()
        {
            InitializeComponent();

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

        #region Обработчики загрузки данных (с обработкой ошибок - ПУНКТ 3)

        private void LoadReport_Click(object sender, RoutedEventArgs e)
        {
            if (dateFrom.SelectedDate == null || dateTo.SelectedDate == null)
            {
                MessageBox.Show("Выберите даты", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!int.TryParse(thresholdBox.Text, out int threshold))
            {
                MessageBox.Show("Порог должен быть числом", "Ошибка ввода", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                var data = _monitoringService.GetAnomalyReport(
                    dateFrom.SelectedDate.Value,
                    dateTo.SelectedDate.Value,
                    threshold
                );

                if (data.Count == 0) MessageBox.Show("Данные по аномалиям не найдены");

                dataGrid.ItemsSource = null;
                _anomalyData = data;
                dataGrid.ItemsSource = _anomalyData;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка базы данных: {ex.Message}", "Критическая ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadIpReport_Click(object sender, RoutedEventArgs e)
        {
            if (ipDateFrom.SelectedDate == null || ipDateTo.SelectedDate == null) return;

            try
            {
                var data = _monitoringService.GetIpReport(
                    ipDateFrom.SelectedDate.Value,
                    ipDateTo.SelectedDate.Value
                );

                if (data.Count == 0) MessageBox.Show("Данные по IP не найдены");

                ipGrid.ItemsSource = null;
                _ipData = data;
                ipGrid.ItemsSource = _ipData;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при получении IP-отчета: {ex.Message}");
            }
        }

        private void LoadContinuousReport_Click(object sender, RoutedEventArgs e)
        {
            if (contDateFrom.SelectedDate == null || contDateTo.SelectedDate == null) return;
            if (!int.TryParse(contThresholdBox.Text, out int threshold)) return;

            try
            {
                var data = _monitoringService.GetContinuousWorkReport(
                    contDateFrom.SelectedDate.Value,
                    contDateTo.SelectedDate.Value,
                    threshold
                );

                contGrid.ItemsSource = null;
                _contData = data;
                contGrid.ItemsSource = _contData;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при расчете непрерывной работы: {ex.Message}");
            }
        }

        private void LoadOperationStats_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var data = _monitoringService.GetOperationStats(
                    opDateFrom.SelectedDate.Value,
                    opDateTo.SelectedDate.Value
                );
                operationGrid.ItemsSource = data;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка статистики: {ex.Message}");
            }
        }

        private void LoadDailyActivity_Click(object sender, RoutedEventArgs e)
        {
            if (date.SelectedDate == null) return;
            try
            {
                var data = _monitoringService.GetDailyActivity(date.SelectedDate.Value);
                dailyGrid.ItemsSource = data;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка за день: {ex.Message}");
            }
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

        private void ExportToExcel<T>(List<T> data, string fileName, string reportName)
        {
            var dialog = new SaveFileDialog
            {
                Filter = "Excel files (*.xlsx)|*.xlsx",
                FileName = fileName
            };

            if (dialog.ShowDialog() != true) return;

            try
            {
                using var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("Отчет");
                var properties = typeof(T).GetProperties();

                // Название отчета
                var titleCell = worksheet.Cell(1, 1);
                titleCell.Value = reportName;
                titleCell.Style.Font.Bold = true;
                titleCell.Style.Font.FontSize = 14;
                titleCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Range(1, 1, 1, properties.Length).Merge();

                // Формирование шапки таблицы
                for (int i = 0; i < properties.Length; i++)
                {
                    var attribute = properties[i].GetCustomAttributes(typeof(DisplayNameAttribute), true)
                        .FirstOrDefault() as DisplayNameAttribute;

                    var headerText = attribute != null ? attribute.DisplayName : properties[i].Name;

                    var cell = worksheet.Cell(2, i + 1);
                    cell.Value = headerText;
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#EEF1F7");
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                }

                // Заполнение данными
                for (int row = 0; row < data.Count; row++)
                {
                    for (int col = 0; col < properties.Length; col++)
                    {
                        var value = properties[col].GetValue(data[row]);
                        var cell = worksheet.Cell(row + 3, col + 1);

                        if (value is DateTime dt)
                        {
                            cell.Value = dt;
                            cell.Style.DateFormat.Format = "dd.MM.yyyy";
                        }
                        else if (value is int || value is long || value is double || value is decimal)
                        {
                            cell.Value = Convert.ToDouble(value);
                        }
                        else
                        {
                            cell.Value = value?.ToString() ?? "";
                        }

                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    }
                }

                // Информационный блок
                int footerRow = data.Count + 5;

                worksheet.Cell(footerRow, 1).Value = "Дата формирования:";
                worksheet.Cell(footerRow, 2).Value = DateTime.Now.ToString("dd.MM.yyyy HH:mm");

                worksheet.Cell(footerRow + 1, 1).Value = "Сформировал:";
                worksheet.Cell(footerRow + 1, 2).Value = Environment.UserName;

                // Автоподбор ширины колонок
                worksheet.Columns().AdjustToContents();

                // Стили для подписей в подвале
                worksheet.Range(footerRow, 1, footerRow + 1, 1).Style.Font.Bold = true;

                workbook.SaveAs(dialog.FileName);
                MessageBox.Show("Данные успешно экспортированы!", "Экспорт", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при создании Excel: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void ExportAnomaly_Click(object sender, RoutedEventArgs e)
        {
            if (_anomalyData == null || _anomalyData.Count == 0) return;
            ExportToExcel(_anomalyData, "AnomalyReport.xlsx", "Отчет по аномальной активности");
        }

        private void ExportIp_Click(object sender, RoutedEventArgs e)
        {
            if (_ipData == null || _ipData.Count == 0) return;
            ExportToExcel(_ipData, "IpReport.xlsx", "Отчет по подозрительным IP-адресам");
        }

        private void ExportContinuous_Click(object sender, RoutedEventArgs e)
        {
            if (_contData == null || _contData.Count == 0) return;
            ExportToExcel(_contData, "ContinuousReport.xlsx", "Отчет по непрерывной активности");
        }

        #endregion
    }
}