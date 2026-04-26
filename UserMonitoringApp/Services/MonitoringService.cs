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
                    u.last_name || ' ' || u.first_name || ' ' || COALESCE(u.patronymic, '') AS full_name,
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

        public List<IpReportItem> GetIpReport(DateTime from, DateTime to)
        {
            var result = new List<IpReportItem>();

            using var conn = new NpgsqlConnection(_connectionString);
            conn.Open();

            string sql = @"
                SELECT 
                    u.username,
                    u.last_name || ' ' || u.first_name || ' ' || COALESCE(u.patronymic, '') AS full_name,
                    COUNT(DISTINCT ll.ip_address) AS ip_count,
                    DATE(ll.logged_at) AS log_date
                FROM login_log ll
                JOIN users u ON u.user_id = ll.user_id
                WHERE ll.logged_at BETWEEN @from AND @to
                  AND ll.is_success = TRUE
                GROUP BY u.username, full_name, log_date
                HAVING COUNT(DISTINCT ll.ip_address) > 1
                ORDER BY log_date DESC;
            ";

            using var cmd = new NpgsqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("from", from);
            cmd.Parameters.AddWithValue("to", to);

            using var reader = cmd.ExecuteReader();

            while (reader.Read())
            {
                result.Add(new IpReportItem
                {
                    Username = reader.GetString(0),
                    FullName = reader.GetString(1),
                    IpCount = reader.GetInt32(2),
                    Date = reader.GetDateTime(3)
                });
            }

            return result;
        }

        public List<ContinuousWorkItem> GetContinuousWorkReport(DateTime from, DateTime to, int threshold)
        {
            var result = new List<ContinuousWorkItem>();

            using var conn = new NpgsqlConnection(_connectionString);
            conn.Open();

            string sql = @"
        WITH activity_days AS (
            SELECT 
                user_id,
                DATE(recorded_at) AS day
            FROM activity_log
            WHERE recorded_at BETWEEN @from AND @to
            GROUP BY user_id, day
        ),
        grouped AS (
            SELECT 
                user_id,
                day,
                day - (ROW_NUMBER() OVER (PARTITION BY user_id ORDER BY day)) * INTERVAL '1 day' AS grp
            FROM activity_days
        ),
        series AS (
            SELECT 
                user_id,
                COUNT(*) AS days_count
            FROM grouped
            GROUP BY user_id, grp
        )
        SELECT 
            u.last_name || ' ' || u.first_name || ' ' || COALESCE(u.patronymic, '') AS full_name,
            s.days_count
        FROM series s
        JOIN users u ON u.user_id = s.user_id
        WHERE s.days_count >= @threshold
        ORDER BY s.days_count DESC;
    ";

            using var cmd = new NpgsqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("from", from);
            cmd.Parameters.AddWithValue("to", to);
            cmd.Parameters.AddWithValue("threshold", threshold);

            using var reader = cmd.ExecuteReader();

            while (reader.Read())
            {
                result.Add(new ContinuousWorkItem
                {
                    FullName = reader.GetString(0),
                    DaysCount = reader.GetInt32(1)
                });
            }

            return result;
        }
    }
}
