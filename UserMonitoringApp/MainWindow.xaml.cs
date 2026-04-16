using Npgsql;
using System.Configuration;
using System.Windows;

namespace UserMonitoringApp
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            TestConnection();
        }

        private void TestConnection()
        {
            try
            {
                var connSettings = ConfigurationManager.ConnectionStrings["DefaultConnection"];

                if (connSettings == null)
                {
                    MessageBox.Show("Строка подключения не найдена");
                    return;
                }

                string connString = connSettings.ConnectionString;

                using var conn = new NpgsqlConnection(connString);
                conn.Open();

                MessageBox.Show("Подключение к БД успешно");
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
        }
    }
}