using System.ComponentModel;

namespace UserMonitoringApp.Models;

public class AnomalyReportItem
{
    [DisplayName("Логин")]
    public string Username { get; set; } = string.Empty;

    [DisplayName("ФИО пользователя")]
    public string FullName { get; set; } = string.Empty;

    [DisplayName("Кол-во запросов")]
    public int RequestsCount { get; set; }
}