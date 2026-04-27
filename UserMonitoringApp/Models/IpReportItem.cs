using System;
using System.ComponentModel;

namespace UserMonitoringApp.Models;

public class IpReportItem
{
    [DisplayName("Идентификатор пользователя")]
    public string Username { get; set; } = string.Empty;

    [DisplayName("ФИО пользователя")]
    public string FullName { get; set; } = string.Empty;

    [DisplayName("Кол-во IP")]
    public int IpCount { get; set; }

    [DisplayName("Дата входа")]
    public DateTime Date { get; set; }
}