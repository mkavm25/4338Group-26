using System;
using System.IO;
using System.Windows;
using Group4338.Services;

namespace Group4338
{
    public partial class GurkinaArina4338 : Window
    {
        private JsonImportService importService;
        private WordExportService exportService;

        public GurkinaArina4338()
        {
            InitializeComponent();

            importService = new JsonImportService();
            exportService = new WordExportService();

            Log("Программа запущена. Готов к работе.");
            CheckDatabase();
        }

        private void CheckDatabase()
        {
            string dbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "clients.db");
            if (File.Exists(dbPath))
            {
                Log($"Файл БД найден: {dbPath}");
            }
            else
            {
                Log("БД будет создана при первом импорте.");
            }
        }

        private void ImportButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Log("Начало импорта JSON...");
                importService.ImportFromJsonFile();
                Log("Импорт завершен.");
            }
            catch (Exception ex)
            {
                Log($"Ошибка: {ex.Message}");
            }
        }

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Log("Начало экспорта в Word...");
                exportService.ExportToWord();
                Log("Экспорт завершен.");
            }
            catch (Exception ex)
            {
                Log($"Ошибка: {ex.Message}");
            }
        }

        private void ResetDatabaseButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string dbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "clients.db");
                if (File.Exists(dbPath))
                {
                    File.Delete(dbPath);
                    Log("Файл БД удален. При следующем импорте создастся новая БД.");
                }
                else
                {
                    Log("Файл БД не найден.");
                }
            }
            catch (Exception ex)
            {
                Log($"Ошибка при удалении БД: {ex.Message}");
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
}