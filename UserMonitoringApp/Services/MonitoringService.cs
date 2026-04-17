using Npgsql;
using System.Configuration;
using UserMonitoringApp.Models;

namespace UserMonitoringApp.Services
{
    public class MonitoringService
    {
        private readonly string _connectionString;

        public MonitoringService()
        {
            _connectionString = ConfigurationManager
                .ConnectionStrings["DefaultConnection"]
                .ConnectionString;
        }

        public List<AnomalyReportItem> GetAnomalyReport(DateTime from, DateTime to, int threshold)
        {
            var result = new List<AnomalyReportItem>();

            using var conn = new NpgsqlConnection(_connectionString);
            conn.Open();

            string sql = @"
                SELECT 
                    u.username,
                    u.last_name || ' ' || u.first_name AS full_name,
                    SUM(CASE WHEN ao.operation_id = 6 THEN ao.count ELSE 0 END) AS requests_count
                FROM activity_log al
                JOIN users u ON u.user_id = al.user_id
                JOIN activity_operations ao ON ao.activity_id = al.activity_id
                WHERE al.recorded_at BETWEEN @from AND @to
                GROUP BY u.username, full_name
                HAVING SUM(CASE WHEN ao.operation_id = 6 THEN ao.count ELSE 0 END) > @threshold
                ORDER BY requests_count DESC;
            ";

            using var cmd = new NpgsqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("from", from);
            cmd.Parameters.AddWithValue("to", to);
            cmd.Parameters.AddWithValue("threshold", threshold);

            using var reader = cmd.ExecuteReader();

            while (reader.Read())
            {
                result.Add(new AnomalyReportItem
                {
                    Username = reader.GetString(0),
                    FullName = reader.GetString(1),
                    RequestsCount = reader.GetInt32(2)
                });
            }

            return result;
        }
    }
}
