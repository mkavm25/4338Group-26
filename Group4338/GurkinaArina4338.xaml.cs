using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Windows;
using OfficeOpenXml;

namespace Group4338
{
    public partial class GurkinaArina4338 : Window
    {
        private readonly string dbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "clients.db");
        private readonly string connectionString;

        public  GurkinaArina4338()
        {
            InitializeComponent();
            ExcelPackage.License.SetNonCommercialPersonal("Гуркина Арина");
            connectionString = $"Data Source={dbPath};Version=3;";
            CreateDatabaseAndTable();
        }

        private void CreateDatabaseAndTable()
        {
            try
            {
                if (!File.Exists(dbPath))
                {
                    SQLiteConnection.CreateFile(dbPath);
                    Log("Создан новый файл базы данных.");
                }

                string createTableQuery = @"
                    CREATE TABLE IF NOT EXISTS Clients (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        FullName TEXT NOT NULL,
                        Email TEXT,
                        Street TEXT NOT NULL
                    )";

                using (var connection = new SQLiteConnection(connectionString))
                {
                    connection.Open();
                    using (var command = new SQLiteCommand(createTableQuery, connection))
                    {
                        command.ExecuteNonQuery();
                    }
                    Log("Таблица Clients создана или уже существует.");
                }
            }
            catch (Exception ex)
            {
                Log($"Ошибка при создании БД: {ex.Message}");
            }
        }

        private void ImportButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
                openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                openFileDialog.Title = "Выберите файл 3.xlsx для импорта";

                if (openFileDialog.ShowDialog() == true)
                {
                    string filePath = openFileDialog.FileName;
                    Log($"Выбран файл: {filePath}");

                    var clients = LoadFromExcel(filePath);

                    if (clients.Count > 0)
                    {
                        SaveToDatabase(clients);
                        Log($"Успешно импортировано {clients.Count} записей.");
                    }
                    else
                    {
                        Log("Нет данных для импорта.");
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"Ошибка при импорте: {ex.Message}");
            }
        }

        private List<Client> LoadFromExcel(string filePath)
        {
            var clients = new List<Client>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                if (worksheet == null)
                {
                    Log("Лист не найден в файле Excel.");
                    return clients;
                }

                int rowCount = worksheet.Dimension?.Rows ?? 0;
                Log($"Найдено строк в Excel: {rowCount}");

                for (int row = 2; row <= rowCount; row++)
                {
                    try
                    {
                        string fullName = worksheet.Cells[row, 1].Text?.Trim();
                        string street = worksheet.Cells[row, 6].Text?.Trim();
                        string email = worksheet.Cells[row, 9].Text?.Trim();

                        if (!string.IsNullOrEmpty(fullName) && !string.IsNullOrEmpty(street))
                        {
                            clients.Add(new Client
                            {
                                FullName = fullName,
                                Street = street,
                                Email = email ?? ""
                            });
                        }
                    }
                    catch (Exception ex)
                    {
                        Log($"Ошибка в строке {row}: {ex.Message}");
                    }
                }
            }

            Log($"Загружено {clients.Count} клиентов из Excel.");
            return clients;
        }

        private void SaveToDatabase(List<Client> clients)
        {
            string deleteQuery = "DELETE FROM Clients";

            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();

                using (var transaction = connection.BeginTransaction())
                {
                    try
                    {
                        using (var deleteCmd = new SQLiteCommand(deleteQuery, connection))
                        {
                            deleteCmd.ExecuteNonQuery();
                        }

                        string insertQuery = "INSERT INTO Clients (FullName, Email, Street) VALUES (@FullName, @Email, @Street)";

                        foreach (var client in clients)
                        {
                            using (var insertCmd = new SQLiteCommand(insertQuery, connection))
                            {
                                insertCmd.Parameters.AddWithValue("@FullName", client.FullName);
                                insertCmd.Parameters.AddWithValue("@Email", client.Email ?? "");
                                insertCmd.Parameters.AddWithValue("@Street", client.Street);
                                insertCmd.ExecuteNonQuery();
                            }
                        }

                        transaction.Commit();
                        Log($"Сохранено {clients.Count} записей в БД.");
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        Log($"Ошибка при сохранении в БД: {ex.Message}");
                        throw;
                    }
                }
            }
        }

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var clients = GetClientsFromDatabase();

                if (clients.Count == 0)
                {
                    Log("Нет данных для экспорта. Сначала выполните импорт.");
                    return;
                }

                var groupedByStreet = clients.GroupBy(c => c.Street).OrderBy(g => g.Key);

                Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog();
                saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                saveFileDialog.FileName = "Экспорт_клиенты.xlsx";
                saveFileDialog.Title = "Сохранить файл Excel";

                if (saveFileDialog.ShowDialog() == true)
                {
                    string filePath = saveFileDialog.FileName;
                    ExportToExcel(groupedByStreet, filePath);
                    Log($"Данные экспортированы в файл: {filePath}");
                }
            }
            catch (Exception ex)
            {
                Log($"Ошибка при экспорте: {ex.Message}");
            }
        }

        private List<Client> GetClientsFromDatabase()
        {
            var clients = new List<Client>();
            string selectQuery = "SELECT Id, FullName, Email, Street FROM Clients ORDER BY Street, FullName";

            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                using (var command = new SQLiteCommand(selectQuery, connection))
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        clients.Add(new Client
                        {
                            Id = Convert.ToInt32(reader["Id"]),
                            FullName = reader["FullName"].ToString(),
                            Email = reader["Email"].ToString(),
                            Street = reader["Street"].ToString()
                        });
                    }
                }
            }

            return clients;
        }

        private void ExportToExcel(IEnumerable<IGrouping<string, Client>> groupedData, string filePath)
        {
            using (var package = new ExcelPackage())
            {
                int sheetCounter = 1;

                foreach (var group in groupedData)
                {
                    string sheetName = group.Key.Length > 30 ? group.Key.Substring(0, 30) : group.Key;
                    if (package.Workbook.Worksheets.Any(w => w.Name == sheetName))
                    {
                        sheetName = $"{sheetName}_{sheetCounter++}";
                    }

                    var worksheet = package.Workbook.Worksheets.Add(sheetName);

                    worksheet.Cells[1, 1].Value = "Код клиента";
                    worksheet.Cells[1, 2].Value = "ФИО";
                    worksheet.Cells[1, 3].Value = "E-mail";

                    using (var range = worksheet.Cells[1, 1, 1, 3])
                    {
                        range.Style.Font.Bold = true;
                    }

                    int row = 2;
                    foreach (var client in group)
                    {
                        worksheet.Cells[row, 1].Value = client.Id;
                        worksheet.Cells[row, 2].Value = client.FullName;
                        worksheet.Cells[row, 3].Value = client.Email;
                        row++;
                    }

                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                }

                FileInfo excelFile = new FileInfo(filePath);
                package.SaveAs(excelFile);
            }
        }

        private void Log(string message)
        {
            Dispatcher.Invoke(() =>
            {
                LogTextBox.AppendText($"{DateTime.Now:HH:mm:ss} - {message}\n");
                LogTextBox.ScrollToEnd();
            });
        }
    }

    public class Client
    {
        public int Id { get; set; }
        public string FullName { get; set; }
        public string Email { get; set; }
        public string Street { get; set; }
    }
}