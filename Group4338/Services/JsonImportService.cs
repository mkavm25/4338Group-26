using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using Microsoft.Win32;
using System.Windows;
using Group4338.Models;
using Group4338.Database;

namespace Group4338.Services
{
    public class JsonImportService
    {
        private DatabaseHelper dbHelper = new DatabaseHelper();

        public void ImportFromJsonFile()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "JSON files (*.json)|*.json";
            openFileDialog.Title = "Выберите файл 3.json для импорта";

            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    string jsonContent = File.ReadAllText(openFileDialog.FileName);
                    var clients = JsonConvert.DeserializeObject<List<Client>>(jsonContent);

                    if (clients != null && clients.Count > 0)
                    {
                        dbHelper.SaveClients(clients);
                        MessageBox.Show(
                            $"Успешно импортировано {clients.Count} клиентов!",
                            "Импорт завершен",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information);
                    }
                    else
                    {
                        MessageBox.Show("Файл не содержит данных", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при импорте: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }
}