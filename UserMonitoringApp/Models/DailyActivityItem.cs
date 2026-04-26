namespace UserMonitoringApp.Models
{
    public class DailyActivityItem
    {
        public string Username { get; set; }
        public string FullName { get; set; }
        public DateTime Day { get; set; }
        public int TotalRequests { get; set; }
    }
}
