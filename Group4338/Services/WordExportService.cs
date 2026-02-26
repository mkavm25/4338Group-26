using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Win32;
using System.Windows;
using Group4338.Models;
using Group4338.Database;

namespace Group4338.Services
{
    public class WordExportService
    {
        private DatabaseHelper dbHelper = new DatabaseHelper();

        public void ExportToWord()
        {
            var clients = dbHelper.GetAllClients();

            if (clients.Count == 0)
            {
                MessageBox.Show("Сначала импортируйте данные!", "Нет данных", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Word files (*.docx)|*.docx";
            saveFileDialog.FileName = $"Клиенты_по_улицам_{DateTime.Now:yyyyMMdd}.docx";

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    CreateWordDocument(clients, saveFileDialog.FileName);
                    MessageBox.Show("Документ успешно создан!", "Экспорт завершен", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при экспорте: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void CreateWordDocument(List<Client> clients, string filePath)
        {
            var groups = clients
                .GroupBy(c => c.Street)
                .OrderBy(g => g.Key);

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                bool isFirstGroup = true;

                foreach (var group in groups)
                {
                    if (!isFirstGroup)
                    {
                        body.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
                    }

                    // Заголовок с названием улицы
                    Paragraph titleParagraph = new Paragraph();
                    Run titleRun = new Run();
                    titleRun.AppendChild(new Text($"Улица: {group.Key}"));
                    titleParagraph.AppendChild(titleRun);

                    titleParagraph.ParagraphProperties = new ParagraphProperties();
                    titleParagraph.ParagraphProperties.Justification = new Justification() { Val = JustificationValues.Center };
                    titleParagraph.ParagraphProperties.SpacingBetweenLines = new SpacingBetweenLines() { After = "300" };

                    titleRun.RunProperties = new RunProperties();
                    titleRun.RunProperties.Append(new Bold());

                    body.AppendChild(titleParagraph);

                    // Создание таблицы
                    Table table = new Table();

                    // Заголовок таблицы
                    TableRow headerRow = new TableRow();
                    headerRow.Append(CreateTableCell("Код клиента", true));
                    headerRow.Append(CreateTableCell("ФИО", true));
                    headerRow.Append(CreateTableCell("E-mail", true));
                    table.Append(headerRow);

                    // Строки с данными (сортировка по ФИО)
                    foreach (var client in group.OrderBy(c => c.FullName))
                    {
                        TableRow dataRow = new TableRow();
                        dataRow.Append(CreateTableCell(client.CodeClient));
                        dataRow.Append(CreateTableCell(client.FullName));
                        dataRow.Append(CreateTableCell(client.E_mail));
                        table.Append(dataRow);
                    }

                    body.AppendChild(table);
                    body.AppendChild(new Paragraph());

                    isFirstGroup = false;
                }

                mainPart.Document.Save();
            }
        }

        private TableCell CreateTableCell(string text, bool isHeader = false)
        {
            TableCell cell = new TableCell();

            Paragraph paragraph = new Paragraph();
            Run run = new Run();
            run.AppendChild(new Text(text ?? ""));

            if (isHeader)
            {
                run.RunProperties = new RunProperties();
                run.RunProperties.Append(new Bold());
            }

            paragraph.AppendChild(run);
            cell.AppendChild(paragraph);

            return cell;
        }
    }
}