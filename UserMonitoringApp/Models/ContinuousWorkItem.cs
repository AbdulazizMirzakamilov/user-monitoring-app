using System.ComponentModel;

namespace UserMonitoringApp.Models;

public class ContinuousWorkItem
{
    [DisplayName("ФИО пользователя")]
    public string FullName { get; set; } = string.Empty;

    [DisplayName("Дней подряд")]
    public int DaysCount { get; set; }
}