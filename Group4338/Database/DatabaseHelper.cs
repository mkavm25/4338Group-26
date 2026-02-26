using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using Group4338.Models;

namespace Group4338.Database
{
    public class DatabaseHelper
    {
        private readonly string connectionString;
        private readonly string dbPath;

        public DatabaseHelper()
        {
            dbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "clients.db");
            connectionString = $"Data Source={dbPath};Version=3;";

            // Удаляем старую БД, если она существует с неправильной структурой
            DeleteOldDatabase();

            // Создаем новую БД с правильной структурой
            CreateDatabaseAndTable();
        }

        private void DeleteOldDatabase()
        {
            try
            {
                if (File.Exists(dbPath))
                {
                    bool hasCorrectStructure = CheckDatabaseStructure();

                    if (!hasCorrectStructure)
                    {
                        File.Delete(dbPath);
                        Console.WriteLine("Старая БД удалена. Будет создана новая с правильной структурой.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при проверке БД: {ex.Message}");
            }
        }

        private bool CheckDatabaseStructure()
        {
            try
            {
                using (var connection = new SQLiteConnection(connectionString))
                {
                    connection.Open();

                    string checkTableQuery = "SELECT name FROM sqlite_master WHERE type='table' AND name='Clients'";
                    using (var cmd = new SQLiteCommand(checkTableQuery, connection))
                    {
                        var result = cmd.ExecuteScalar();
                        if (result == null)
                            return false; 
                    }

                    string checkColumnQuery = "PRAGMA table_info(Clients)";
                    using (var cmd = new SQLiteCommand(checkColumnQuery, connection))
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string columnName = reader["name"].ToString();
                            if (columnName == "CodeClient")
                                return true;
                        }
                    }
                }
                return false; 
            }
            catch
            {
                return false;
            }
        }

        private void CreateDatabaseAndTable()
        {
            if (!File.Exists(dbPath))
            {
                SQLiteConnection.CreateFile(dbPath);
            }

            string createTableQuery = @"
                CREATE TABLE IF NOT EXISTS Clients (
                    Id INTEGER PRIMARY KEY,
                    FullName TEXT,
                    CodeClient TEXT,
                    BirthDate TEXT,
                    [Index] TEXT,
                    City TEXT,
                    Street TEXT,
                    Home INTEGER,
                    Kvartira INTEGER,
                    E_mail TEXT
                )";

            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                using (var command = new SQLiteCommand(createTableQuery, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }

        public void SaveClients(List<Client> clients)
        {
            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();

                using (var clearCmd = new SQLiteCommand("DELETE FROM Clients", connection))
                {
                    clearCmd.ExecuteNonQuery();
                }

                foreach (var client in clients)
                {
                    string sql = @"
                        INSERT INTO Clients 
                        (Id, FullName, CodeClient, BirthDate, [Index], City, Street, Home, Kvartira, E_mail)
                        VALUES 
                        (@id, @full, @code, @birth, @index, @city, @street, @home, @kv, @email)";

                    using (var cmd = new SQLiteCommand(sql, connection))
                    {
                        cmd.Parameters.AddWithValue("@id", client.Id);
                        cmd.Parameters.AddWithValue("@full", client.FullName);
                        cmd.Parameters.AddWithValue("@code", client.CodeClient);
                        cmd.Parameters.AddWithValue("@birth", client.BirthDate);
                        cmd.Parameters.AddWithValue("@index", client.Index);
                        cmd.Parameters.AddWithValue("@city", client.City);
                        cmd.Parameters.AddWithValue("@street", client.Street);
                        cmd.Parameters.AddWithValue("@home", client.Home);
                        cmd.Parameters.AddWithValue("@kv", client.Kvartira);
                        cmd.Parameters.AddWithValue("@email", client.E_mail);
                        cmd.ExecuteNonQuery();
                    }
                }
            }
        }

        public List<Client> GetAllClients()
        {
            var clients = new List<Client>();

            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                string sql = "SELECT Id, FullName, CodeClient, BirthDate, [Index], City, Street, Home, Kvartira, E_mail FROM Clients";
                using (var cmd = new SQLiteCommand(sql, connection))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        clients.Add(new Client
                        {
                            Id = Convert.ToInt32(reader["Id"]),
                            FullName = reader["FullName"].ToString(),
                            CodeClient = reader["CodeClient"].ToString(),
                            BirthDate = reader["BirthDate"].ToString(),
                            Index = reader["Index"].ToString(),
                            City = reader["City"].ToString(),
                            Street = reader["Street"].ToString(),
                            Home = Convert.ToInt32(reader["Home"]),
                            Kvartira = Convert.ToInt32(reader["Kvartira"]),
                            E_mail = reader["E_mail"].ToString()
                        });
                    }
                }
            }

            return clients;
        }
    }
}