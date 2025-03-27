using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using Newtonsoft.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace _4337Project
{
    public class Sergeev_Export_Import
    {
        
        public static void ImportData(string filePath, string connectionString, string tableName)
        {
            try
            {
                if (!File.Exists(filePath))
                    throw new FileNotFoundException("Файл не найден.", filePath);

                string jsonData = File.ReadAllText(filePath);
                List<Order> orders = ImportJsonData(jsonData);

                CreateTable(connectionString, tableName);
                SaveDataToTable(connectionString, tableName, orders);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при импорте: {ex.Message}");
                throw;
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

        private static void SaveDataToTable(string connectionString, string tableName, List<Order> orders)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                foreach (var order in orders)
                {
                    string insertQuery = $@"
                    INSERT INTO {tableName} (
                        [Код заказа], [Дата создания], [Время заказа], [Код клиента], 
                        [Услуги], [Статус], [Дата закрытия], [Время проката]
                    ) VALUES (@p1, @p2, @p3, @p4, @p5, @p6, @p7, @p8)";

                    using (SqlCommand command = new SqlCommand(insertQuery, connection))
                    {
                        command.Parameters.AddWithValue("@p1", order.CodeOrder ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@p2", order.ParsedCreateDate.HasValue ? order.ParsedCreateDate.Value.ToString("yyyy-MM-dd") : (object)DBNull.Value);
                        command.Parameters.AddWithValue("@p3", order.ParsedCreateTime.HasValue ? order.ParsedCreateTime.Value.ToString(@"hh\:mm") : (object)DBNull.Value);
                        command.Parameters.AddWithValue("@p4", order.CodeClient ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@p5", order.Services ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@p6", order.Status ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@p7", order.ParsedClosedDate.HasValue ? order.ParsedClosedDate.Value.ToString("yyyy-MM-dd") : (object)DBNull.Value);
                        command.Parameters.AddWithValue("@p8", order.ProkatTime ?? (object)DBNull.Value);

                        command.ExecuteNonQuery();
                    }
                }
            }
        }

        public static void ExportData(string connectionString, string tableName, string outputFilePath)
        {
            try
            {
                List<Order> orders = GetDataFromTable(connectionString, tableName);
                CreateWordDocument(orders, outputFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при экспорте: {ex.Message}");
                throw;
            }
        }

        private static List<Order> GetDataFromTable(string connectionString, string tableName)
        {
            List<Order> orders = new List<Order>();
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
                            orders.Add(new Order
                            {
                                Id = (int)reader["Id"],
                                CodeOrder = reader["Код заказа"].ToString(),
                                ParsedCreateDate = reader["Дата создания"] as DateTime?,
                                ParsedCreateTime = reader["Время заказа"] as TimeSpan?,
                                CodeClient = reader["Код клиента"].ToString(),
                                Services = reader["Услуги"].ToString(),
                                Status = reader["Статус"].ToString(),
                                ParsedClosedDate = reader["Дата закрытия"] as DateTime?,
                                ProkatTime = reader["Время проката"].ToString()
                            });
                        }
                    }
                }
            }
            return orders;
        }

        private static void CreateWordDocument(List<Order> orders, string outputFilePath)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(outputFilePath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = new Body();

                Paragraph title = new Paragraph(new Run(new Text("Отчёт о заказах")));
                title.ParagraphProperties = new ParagraphProperties(new Justification() { Val = JustificationValues.Center });
                body.Append(title);

                Table table = new Table();

                // Добавляем заголовки
                TableRow headerRow = new TableRow();
                string[] headers = { "Id", "Код заказа", "Код клиента", "Услуги" };
                foreach (var header in headers)
                {
                    headerRow.Append(new TableCell(new Paragraph(new Run(new Text(header)))));
                }
                table.Append(headerRow);

                // Добавляем данные
                foreach (var order in orders)
                {
                    TableRow row = new TableRow();
                    row.Append(new TableCell(new Paragraph(new Run(new Text(order.Id.ToString())))));
                    row.Append(new TableCell(new Paragraph(new Run(new Text(order.CodeOrder)))));
                    row.Append(new TableCell(new Paragraph(new Run(new Text(order.CodeClient)))));
                    row.Append(new TableCell(new Paragraph(new Run(new Text(order.Services)))));

                    table.Append(row);
                }

                body.Append(table);
                mainPart.Document.Append(body);
                mainPart.Document.Save();
            }
        }
        private static List<Order> ImportJsonData(string jsonData)
        {
            var settings = new JsonSerializerSettings
            {
                DateFormatString = "dd.MM.yyyy", 
                NullValueHandling = NullValueHandling.Ignore 
            };

            var orders = JsonConvert.DeserializeObject<List<Order>>(jsonData, settings);

            foreach (var order in orders)
            {
                if (!string.IsNullOrEmpty(order.CreateDate))
                {
                    if (DateTime.TryParseExact(order.CreateDate, "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var createDate))
                    {
                        order.ParsedCreateDate = createDate;
                    }
                    else
                    {
                        order.ParsedCreateDate = null; 
                    }
                }

                if (!string.IsNullOrEmpty(order.CreateTime))
                {
                    if (TimeSpan.TryParse(order.CreateTime, out var createTime))
                    {
                        order.ParsedCreateTime = createTime;
                    }
                    else
                    {
                        order.ParsedCreateTime = null; 
                    }
                }

                if (!string.IsNullOrEmpty(order.ClosedDate))
                {
                    if (DateTime.TryParseExact(order.ClosedDate, "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var closedDate))
                    {
                        order.ParsedClosedDate = closedDate;
                    }
                    else
                    {
                        order.ParsedClosedDate = null; 
                    }
                }
            }

            return orders;
        }
    }

    public class Order
    {
        public int Id { get; set; }
        public string CodeOrder { get; set; }

        public string CreateDate { get; set; } 
        public string CreateTime { get; set; } 
        public string CodeClient { get; set; }
        public string Services { get; set; }
        public string Status { get; set; }
        public string ClosedDate { get; set; } 
        public string ProkatTime { get; set; }

        [JsonIgnore] 
        public DateTime? ParsedCreateDate { get; set; }

        [JsonIgnore] 
        public TimeSpan? ParsedCreateTime { get; set; }

        [JsonIgnore] 
        public DateTime? ParsedClosedDate { get; set; }
    }
}