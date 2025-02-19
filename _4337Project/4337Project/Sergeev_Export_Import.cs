using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Globalization;
using OfficeOpenXml;
using System.Linq;

namespace _4337Project
{
    public class Sergeev_Export_Import
    {
        public static void ImportData(string filePath, string connectionString, string tableName)
        {
            FileInfo fileInfo = new FileInfo(filePath);
            if (!fileInfo.Exists) throw new FileNotFoundException("Файл не найден.", filePath);

            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                if (package.Workbook.Worksheets.Count == 0)
                    throw new InvalidOperationException("Файл не содержит ни одного листа.");

                ExcelWorksheet worksheet = package.Workbook.Worksheets["Лист1"];
                worksheet.Calculate();
                int rowCount = worksheet.Dimension?.Rows ?? 0;

                if (rowCount == 0) throw new InvalidOperationException("Лист пустой.");

                CreateTable(connectionString, tableName);
                SaveDataToTable(connectionString, tableName, worksheet, rowCount);
            }
        }

        private static void CreateTable(string connectionString, string tableName)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                string createTableQuery = $@"
                IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = '{tableName}')
                BEGIN
                    CREATE TABLE {tableName} (
                        Id INT IDENTITY(1,1) PRIMARY KEY,
                        [Код заказа] NVARCHAR(50),
                        [Дата создания] DATE,
                        [Время заказа] TIME,
                        [Код клиента] NVARCHAR(50),
                        [Услуги] NVARCHAR(MAX),
                        [Статус] NVARCHAR(50),
                        [Дата закрытия] DATE,
                        [Время проката] NVARCHAR(50)
                    )
                END";
                using (SqlCommand command = new SqlCommand(createTableQuery, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }

        private static void SaveDataToTable(string connectionString, string tableName, ExcelWorksheet worksheet, int rowCount)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                for (int row = 2; row <= rowCount; row++)
                {
                    string orderCode = worksheet.Cells[row, 2].Text.Trim(); // Код заказа (столбец 2)
                    string orderDateText = worksheet.Cells[row, 3].Text.Trim(); // Дата создания (столбец 3)
                    string orderTimeText = worksheet.Cells[row, 4].Text.Trim(); // Время заказа (столбец 4)
                    string clientCode = worksheet.Cells[row, 5].Text.Trim(); // Код клиента (столбец 5)
                    string services = worksheet.Cells[row, 6].Text.Trim(); // Услуги (столбец 6)
                    string status = worksheet.Cells[row, 7].Text.Trim(); // Статус (столбец 7)
                    string closeDateText = worksheet.Cells[row, 8].Text.Trim(); // Дата закрытия (столбец 8)
                    string rentalTime = worksheet.Cells[row, 9].Text.Trim(); // Время проката (столбец 9)

                    // Парсинг даты и времени
                    DateTime? orderDate = ParseDate(orderDateText);
                    TimeSpan? orderTime = ParseTime(orderTimeText);
                    DateTime? closeDate = ParseDate(closeDateText);

                    // SQL-запрос
                    string insertQuery = $@"
                    INSERT INTO {tableName} (
                        [Код заказа], [Дата создания], [Время заказа], [Код клиента], 
                        [Услуги], [Статус], [Дата закрытия], [Время проката]
                    ) VALUES (@p1, @p2, @p3, @p4, @p5, @p6, @p7, @p8)";

                    using (SqlCommand command = new SqlCommand(insertQuery, connection))
                    {
                        command.Parameters.AddWithValue("@p1", (object)orderCode ?? DBNull.Value);
                        command.Parameters.AddWithValue("@p2", (object)orderDate ?? DBNull.Value);
                        command.Parameters.AddWithValue("@p3", (object)orderTime ?? DBNull.Value);
                        command.Parameters.AddWithValue("@p4", (object)clientCode ?? DBNull.Value);
                        command.Parameters.AddWithValue("@p5", (object)services ?? DBNull.Value);
                        command.Parameters.AddWithValue("@p6", (object)status ?? DBNull.Value);
                        command.Parameters.AddWithValue("@p7", (object)closeDate ?? DBNull.Value);
                        command.Parameters.AddWithValue("@p8", (object)rentalTime ?? DBNull.Value);
                        command.ExecuteNonQuery();
                    }
                }
            }
        }

        public static void ExportData(string connectionString, string tableName, string outputFilePath)
        {
            List<Dictionary<string, object>> data = GetDataFromTable(connectionString, tableName);
            var groupedData = data.GroupBy(row => row["Дата создания"]?.ToString() ?? "Нет даты");
            CreateExcel(groupedData, outputFilePath);
        }

        private static List<Dictionary<string, object>> GetDataFromTable(string connectionString, string tableName)
        {
            List<Dictionary<string, object>> data = new List<Dictionary<string, object>>();
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                string query = $"SELECT * FROM {tableName}";
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Dictionary<string, object> row = new Dictionary<string, object>();
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                string columnName = reader.GetName(i);
                                row[columnName] = reader.IsDBNull(i) ? null : reader.GetValue(i);
                            }
                            data.Add(row);
                        }
                    }
                }
            }
            return data;
        }

        private static void CreateExcel(IEnumerable<IGrouping<string, Dictionary<string, object>>> groupedData, string outputFilePath)
        {
            FileInfo newFile = new FileInfo(outputFilePath);
            if (newFile.Exists) newFile.Delete();

            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                foreach (var group in groupedData)
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(group.Key);
                    string[] columnNames = { "Id", "Код заказа", "Дата создания", "Время заказа", "Код клиента", "Услуги", "Статус", "Дата закрытия", "Время проката" };

                    // Заголовки столбцов
                    for (int i = 0; i < columnNames.Length; i++)
                    {
                        worksheet.Cells[1, i + 1].Value = columnNames[i];
                    }

                    // Данные
                    int row = 2;
                    foreach (var record in group)
                    {
                        for (int i = 0; i < columnNames.Length; i++)
                        {
                            var cell = worksheet.Cells[row, i + 1];
                            var value = record[columnNames[i]] ?? "Нет данных";

                            // Убедитесь, что "Код заказа" записывается как текст
                            if (columnNames[i] == "Код заказа")
                            {
                                cell.Value = "'" + value.ToString(); // Добавляем апостроф для принудительного текстового формата
                                cell.Style.Numberformat.Format = "@"; // Устанавливаем текстовый формат
                            }
                            else
                            {
                                cell.Value = value;
                            }
                        }
                        row++;
                    }
                }
                package.Save();
            }
        }

        private static DateTime? ParseDate(string dateText)
        {
            if (string.IsNullOrEmpty(dateText)) return null;
            if (double.TryParse(dateText, out double dateNum))
            {
                return DateTime.FromOADate(dateNum);
            }
            else if (DateTime.TryParseExact(dateText, "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate))
            {
                return parsedDate;
            }
            return null;
        }

        private static TimeSpan? ParseTime(string timeText)
        {
            if (string.IsNullOrEmpty(timeText)) return null;
            if (TimeSpan.TryParseExact(timeText, "hh\\:mm", CultureInfo.InvariantCulture, out TimeSpan parsedTime))
            {
                return parsedTime;
            }
            return null;
        }
    }
}
    