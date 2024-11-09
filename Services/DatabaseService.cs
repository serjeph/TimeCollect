using Npgsql;
using System;
using System.Collections.Generic;
using System.Windows;
using TimeCollect.Models;

namespace TimeCollect.Services
{
    public class DatabaseService
    {
        private readonly string _connectionString;

        public DatabaseService(DatabaseSettings settings)
        {
            _connectionString = $"" +
                $"Host={settings.Host};" +
                $"Database={settings.Database};" +
                $"Username={settings.Username};" +
                $"Password={settings.Password};" +
                $"Port={settings.Port};";
        }

        public void InsertData(IList<IList<object>> data, string sheetName)
        {
            try
            {
                string tableName = $"timesheet_{sheetName}";

                // Create table if it does not exist
                CreateTableIfNotExists(tableName);

                using (var conn = new NpgsqlConnection(_connectionString))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand())
                    {
                        cmd.Connection = conn;
                        cmd.CommandText = $@"INSERT INTO {tableName} ( uuid, " +
                            "employee_id, " +
                            "row_id, " +
                            "year, " +
                            "month, " +
                            "day, " +
                            "week_type, " +
                            "employee_name, " +
                            "project_code, " +
                            "task_type, " +
                            "work_type, " +
                            "is_actual, " +
                            "worked_hours ) " +
                            "VALUES " +
                            "( @uuid, " +
                            "@employee_id, " +
                            "@row_id, " +
                            "@year, " +
                            "@month, " +
                            "@day, " +
                            "@week_type, " +
                            "@employee_name, " +
                            "@project_code, " +
                            "@task_type, " +
                            "@work_type, " +
                            "@is_actual, " +
                            "@worked_hours ) " +
                            "ON CONFLICT (uuid) DO UPDATE " +
                            "SET " +
                            "   employee_id = EXCLUDED.employee_id," +
                            "   row_id = EXCLUDED.row_id," +
                            "   year = EXCLUDED.year," +
                            "   month = EXCLUDED.month," +
                            "   day = EXCLUDED.day," +
                            "   week_type = EXCLUDED.week_type," +
                            "   employee_name = EXCLUDED.employee_name," +
                            "   project_code = EXCLUDED.project_code," +
                            "   task_type = EXCLUDED.task_type," +
                            "   work_type = EXCLUDED.work_type," +
                            "   is_actual = EXCLUDED.is_actual," +
                            "   worked_hours = EXCLUDED.worked_hours;";

                        cmd.Parameters.Clear(); // Clear previous parameters
                        // Add parameters
                        cmd.Parameters.AddWithValue("uuid", NpgsqlTypes.NpgsqlDbType.Integer);
                        cmd.Parameters.AddWithValue("employee_id", NpgsqlTypes.NpgsqlDbType.Integer);
                        cmd.Parameters.AddWithValue("row_id", NpgsqlTypes.NpgsqlDbType.Integer);
                        cmd.Parameters.AddWithValue("year", NpgsqlTypes.NpgsqlDbType.Integer);
                        cmd.Parameters.AddWithValue("month", NpgsqlTypes.NpgsqlDbType.Integer);
                        cmd.Parameters.AddWithValue("day", NpgsqlTypes.NpgsqlDbType.Integer);
                        cmd.Parameters.AddWithValue("week_type", NpgsqlTypes.NpgsqlDbType.Varchar);
                        cmd.Parameters.AddWithValue("employee_name", NpgsqlTypes.NpgsqlDbType.Varchar);
                        cmd.Parameters.AddWithValue("project_code", NpgsqlTypes.NpgsqlDbType.Varchar);
                        cmd.Parameters.AddWithValue("task_type", NpgsqlTypes.NpgsqlDbType.Varchar);
                        cmd.Parameters.AddWithValue("work_type", NpgsqlTypes.NpgsqlDbType.Varchar);
                        cmd.Parameters.AddWithValue("is_actual", NpgsqlTypes.NpgsqlDbType.Varchar);
                        cmd.Parameters.AddWithValue("worked_hours", NpgsqlTypes.NpgsqlDbType.Numeric);

                        foreach (var rowData in data)
                        {
                            cmd.Parameters["uuid"].Value = Convert.ToInt32(rowData[0]);
                            cmd.Parameters["employee_id"].Value = Convert.ToInt32(rowData[1]);
                            cmd.Parameters["row_id"].Value = Convert.ToInt32(rowData[2]);
                            cmd.Parameters["year"].Value = Convert.ToInt32(rowData[3]);
                            cmd.Parameters["month"].Value = Convert.ToInt32(rowData[4]);
                            cmd.Parameters["day"].Value = Convert.ToInt32(rowData[5]);
                            cmd.Parameters["week_type"].Value = rowData[6].ToString();
                            cmd.Parameters["employee_name"].Value = rowData[7].ToString();
                            cmd.Parameters["project_code"].Value = rowData[8].ToString();
                            cmd.Parameters["task_type"].Value = rowData[9].ToString();
                            cmd.Parameters["work_type"].Value = rowData[10].ToString();
                            cmd.Parameters["is_actual"].Value = rowData[11].ToString();
                            cmd.Parameters["worked_hours"].Value = Convert.ToDecimal(rowData[12]);

                            cmd.ExecuteNonQuery();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred while inserting into the database: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public List<string> GetColumnHeaders(string tableName)
        {
            var columnHeaders = new List<string>();

            try
            {
                using (var conn = new NpgsqlConnection(_connectionString))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand("SELECT column_name " +
                        "FROM information_schema.columns " +
                        "WHERE table_schema = 'public' " +
                        $"AND table_name = '{tableName}';", conn))
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            columnHeaders.Add(reader.GetString(0));
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error getting column headers: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            return columnHeaders;
        }

        private void CreateTableIfNotExists(string tableName)
        {
            try
            {
                using (var conn = new NpgsqlConnection(_connectionString))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand())
                    {
                        cmd.Connection = conn;
                        cmd.CommandText = $@"
                            CREATE TABLE IF NOT EXISTS {tableName} (
                                uuid int PRIMARY KEY,
                                employee_id int,
                                row_id int,
                                year int,
                                month int,
                                day int,
                                week_type VARCHAR(50),
                                employee_name VARCHAR(100),
                                project_code VARCHAR(100),
                                task_type VARCHAR(100),
                                work_type VARCHAR(100),
                                is_actual VARCHAR(10),
                                worked_hours NUMERIC(5,2)
                            )";
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating table: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

    }
}
