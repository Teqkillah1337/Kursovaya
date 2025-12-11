using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;

namespace ComputerPassport
{
    public partial class ComparisonPage : Page
    {
        private readonly ComparisonResult _result;
        private readonly Exporter.ComputerInfo _savedInfo;
        private readonly Exporter.ComputerInfo _currentInfo;

        public ComparisonPage(ComparisonResult result, Exporter.ComputerInfo currentInfo, Exporter.ComputerInfo savedInfo)
        {
            InitializeComponent();
            _result = result;
            _currentInfo = currentInfo;
            _savedInfo = savedInfo;
            DisplayComparisonResults();
        }

        private void DisplayComparisonResults()
        {
            // Добавляем инвентарные номера в заголовок
            string currentInventory = string.IsNullOrEmpty(_currentInfo.InventoryNumber) ? "нет" : _currentInfo.InventoryNumber;
            string savedInventory = string.IsNullOrEmpty(_savedInfo.InventoryNumber) ? "нет" : _savedInfo.InventoryNumber;

            txtTitle.Text = $"Сравнение паспортов: {_currentInfo.UniqueId} ↔ {_savedInfo.UniqueId}";
            txtDates.Text = $"Текущий: {_currentInfo.CollectedAt:dd.MM.yyyy HH:mm} | Инв.№: {currentInventory} | Сохранённый: {_savedInfo.CollectedAt:dd.MM.yyyy HH:mm} | Инв.№: {savedInventory}";

            // Сводка
            DisplaySummary();

            // Детали изменений
            DisplayChangesDetails();

            // По разделам
            DisplayChangesBySections();

            // Статус
            UpdateStatus();
        }

        private void DisplaySummary()
        {
            var summary = new StringBuilder();
            summary.AppendLine("ОБЩАЯ СВОДКА ИЗМЕНЕНИЙ");
            summary.AppendLine("======================");
            summary.AppendLine();

            int totalChanges = _result.AddedItems.Count + _result.RemovedItems.Count + _result.ChangedItems.Count;

            // Проверяем изменение инвентарного номера
            bool inventoryChanged = _currentInfo.InventoryNumber != _savedInfo.InventoryNumber;
            if (inventoryChanged)
            {
                totalChanges++; // Учитываем изменение инвентарного номера
            }

            if (totalChanges == 0)
            {
                summary.AppendLine("🎉 Конфигурации идентичны!");
                summary.AppendLine("   Не обнаружено никаких изменений в системе.");
            }
            else
            {
                summary.AppendLine($"📈 Обнаружено изменений: {totalChanges}");
                summary.AppendLine();

                if (inventoryChanged)
                {
                    string oldInventory = string.IsNullOrEmpty(_savedInfo.InventoryNumber) ? "не указан" : _savedInfo.InventoryNumber;
                    string newInventory = string.IsNullOrEmpty(_currentInfo.InventoryNumber) ? "не указан" : _currentInfo.InventoryNumber;
                    summary.AppendLine($"📝 Изменен инвентарный номер:");
                    summary.AppendLine($"    Было: {oldInventory}");
                    summary.AppendLine($"    Стало: {newInventory}");
                    summary.AppendLine();
                }

                summary.AppendLine($"➕ Добавлено: {_result.AddedItems.Count}");
                summary.AppendLine($"➖ Удалено: {_result.RemovedItems.Count}");
                summary.AppendLine($"✏️ Изменено: {_result.ChangedItems.Count}");
                summary.AppendLine();
                summary.AppendLine("💡 Используйте вкладки 'Детали изменений' и 'По разделам'");
                summary.AppendLine("   для просмотра подробной информации.");
            }

            txtSummary.Text = summary.ToString();
        }

        private void DisplayChangesDetails()
        {
            spChanges.Children.Clear();

            // Добавленные элементы
            if (_result.AddedItems.Count > 0)
            {
                AddSectionHeader(spChanges, $"➕ ДОБАВЛЕНО ({_result.AddedItems.Count})", "#4CAF50");
                foreach (var item in _result.AddedItems)
                {
                    AddChangeItem(spChanges, item, "#1E1E2E");
                }
            }

            // Удаленные элементы
            if (_result.RemovedItems.Count > 0)
            {
                AddSectionHeader(spChanges, $"➖ УДАЛЕНО ({_result.RemovedItems.Count})", "#F44336");
                foreach (var item in _result.RemovedItems)
                {
                    AddChangeItem(spChanges, item, "#1E1E2E");
                }
            }

            // Измененные элементы (БЕЗ инвентарного номера - он будет только в разделе "По разделам")
            if (_result.ChangedItems.Count > 0)
            {
                AddSectionHeader(spChanges, $"✏️ ИЗМЕНЕНО ({_result.ChangedItems.Count})", "#FF9800");
                foreach (var item in _result.ChangedItems)
                {
                    AddChangeItem(spChanges, item, "#1E1E2E");
                }
            }

            if (_result.AddedItems.Count == 0 && _result.RemovedItems.Count == 0 && _result.ChangedItems.Count == 0)
            {
                var noChangesText = new TextBlock
                {
                    Text = "🎉 Изменений не обнаружено",
                    Foreground = System.Windows.Media.Brushes.LightGreen,
                    FontSize = 14,
                    FontWeight = FontWeights.Bold,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Margin = new Thickness(0, 20, 0, 0)
                };
                spChanges.Children.Add(noChangesText);
            }
        }

        private void DisplayChangesBySections()
        {
            spBySections.Children.Clear();

            var sections = new Dictionary<string, List<string>>
            {
                ["Реквизиты"] = new List<string>(),
                ["Аппаратная конфигурация"] = new List<string>(),
                ["Программное обеспечение"] = new List<string>(),
                ["Драйверы"] = new List<string>(),
                ["Периферийные устройства"] = new List<string>()
            };

            // Добавляем изменение инвентарного номера в раздел "Реквизиты" (ТОЛЬКО ЗДЕСЬ)
            if (_currentInfo.InventoryNumber != _savedInfo.InventoryNumber)
            {
                string oldInventory = string.IsNullOrEmpty(_savedInfo.InventoryNumber) ? "не указан" : _savedInfo.InventoryNumber;
                string newInventory = string.IsNullOrEmpty(_currentInfo.InventoryNumber) ? "не указан" : _currentInfo.InventoryNumber;
                sections["Реквизиты"].Add($"✏️ ИЗМЕНЕН ИНВЕНТАРНЫЙ НОМЕР: Было: {oldInventory} → Стало: {newInventory}");
            }

            // Группируем изменения по разделам
            GroupChangesBySection(_result.AddedItems, "➕ ", sections);
            GroupChangesBySection(_result.RemovedItems, "➖ ", sections);
            GroupChangesBySection(_result.ChangedItems, "✏️ ", sections);

            foreach (var section in sections)
            {
                if (section.Value.Count > 0)
                {
                    AddSectionHeader(spBySections, $"{section.Key} ({section.Value.Count} изменений)", "#2196F3");
                    foreach (var item in section.Value)
                    {
                        AddChangeItem(spBySections, item, "#1E1E2E");
                    }
                }
            }

            if (spBySections.Children.Count == 0)
            {
                var noChangesText = new TextBlock
                {
                    Text = "🎉 Изменений не обнаружено",
                    Foreground = System.Windows.Media.Brushes.LightGreen,
                    FontSize = 14,
                    FontWeight = FontWeights.Bold,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Margin = new Thickness(0, 20, 0, 0)
                };
                spBySections.Children.Add(noChangesText);
            }
        }

        private void GroupChangesBySection(List<string> changes, string prefix, Dictionary<string, List<string>> sections)
        {
            foreach (var change in changes)
            {
                var section = GetSectionFromChange(change);
                if (sections.ContainsKey(section))
                {
                    sections[section].Add(prefix + change.Substring(change.IndexOf(']') + 2));
                }
            }
        }

        private string GetSectionFromChange(string change)
        {
            if (change.Contains("[Аппаратная конфигурация]")) return "Аппаратная конфигурация";
            if (change.Contains("[Программное обеспечение]")) return "Программное обеспечение";
            if (change.Contains("[Драйверы]")) return "Драйверы";
            if (change.Contains("[Периферийные устройства]")) return "Периферийные устройства";
            if (change.Contains("[Реквизиты]")) return "Реквизиты";
            return "Другие";
        }

        private void AddSectionHeader(Panel parent, string text, string color)
        {
            var header = new TextBlock
            {
                Text = text,
                Foreground = System.Windows.Media.Brushes.White,
                Background = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString(color),
                FontWeight = FontWeights.Bold,
                FontSize = 12,
                Padding = new Thickness(10, 5, 10, 5),
                Margin = new Thickness(0, 10, 0, 5)
            };
            parent.Children.Add(header);
        }

        private void AddChangeItem(Panel parent, string text, string color)
        {
            var item = new TextBlock
            {
                Text = text,
                Foreground = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString(color),
                FontFamily = new System.Windows.Media.FontFamily("Consolas"),
                FontSize = 10,
                Margin = new Thickness(15, 2, 0, 2),
                TextWrapping = TextWrapping.Wrap
            };
            parent.Children.Add(item);
        }

        private void UpdateStatus()
        {
            int totalChanges = _result.AddedItems.Count + _result.RemovedItems.Count + _result.ChangedItems.Count;

            // Учитываем изменение инвентарного номера
            if (_currentInfo.InventoryNumber != _savedInfo.InventoryNumber)
            {
                totalChanges++;
            }

            if (totalChanges == 0)
            {
                txtStatus.Text = "✅ Конфигурации идентичны - изменений не обнаружено";
                txtStatus.Foreground = System.Windows.Media.Brushes.LightGreen;
            }
            else
            {
                int changedCount = _result.ChangedItems.Count;
                if (_currentInfo.InventoryNumber != _savedInfo.InventoryNumber)
                {
                    changedCount++;
                }

                txtStatus.Text = $"⚠ Обнаружено изменений: {totalChanges} (➕{_result.AddedItems.Count} ➖{_result.RemovedItems.Count} ✏️{changedCount})";
                txtStatus.Foreground = System.Windows.Media.Brushes.Orange;
            }
        }

        private void BtnExportSummary_Click(object sender, RoutedEventArgs e)
        {
            ExportComparisonReport(false);
        }

        private void BtnExportFull_Click(object sender, RoutedEventArgs e)
        {
            ExportComparisonReport(true);
        }

        private void ExportComparisonReport(bool fullReport)
        {
            var saveDialog = new SaveFileDialog
            {
                Filter = "Текстовые файлы (*.txt)|*.txt|HTML файлы (*.html)|*.html",
                FileName = $"сравнение_паспортов_{DateTime.Now:yyyyMMddHHmmss}",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            };

            if (saveDialog.ShowDialog() == true)
            {
                try
                {
                    string content = fullReport ?
                        GenerateFullComparisonReport() :
                        GenerateSummaryReport();

                    if (saveDialog.FilterIndex == 2) // HTML
                    {
                        content = ConvertToHtmlReport(content);
                    }

                    File.WriteAllText(saveDialog.FileName, content, Encoding.UTF8);
                    MessageBox.Show($"Отчёт сохранён в файл:\n{saveDialog.FileName}", "Экспорт завершён",
                                  MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при сохранении: {ex.Message}", "Ошибка",
                                  MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private string GenerateSummaryReport()
        {
            var report = new StringBuilder();
            report.AppendLine("ОТЧЁТ СРАВНЕНИЯ ПАСПОРТОВ КОМПЬЮТЕРА - СВОДКА");
            report.AppendLine("==============================================");
            report.AppendLine();

            string currentInventory = string.IsNullOrEmpty(_currentInfo.InventoryNumber) ? "не указан" : _currentInfo.InventoryNumber;
            string savedInventory = string.IsNullOrEmpty(_savedInfo.InventoryNumber) ? "не указан" : _savedInfo.InventoryNumber;

            report.AppendLine($"Текущий паспорт: {_currentInfo.UniqueId} (Инв.№: {currentInventory})");
            report.AppendLine($"Сохранённый паспорт: {_savedInfo.UniqueId} (Инв.№: {savedInventory})");
            report.AppendLine($"Дата сравнения: {DateTime.Now:dd.MM.yyyy HH:mm:ss}");
            report.AppendLine();

            int totalChanges = _result.AddedItems.Count + _result.RemovedItems.Count + _result.ChangedItems.Count;

            // Проверяем изменение инвентарного номера
            bool inventoryChanged = _currentInfo.InventoryNumber != _savedInfo.InventoryNumber;
            if (inventoryChanged)
            {
                totalChanges++;
            }

            if (totalChanges == 0)
            {
                report.AppendLine("РЕЗУЛЬТАТ: Конфигурации идентичны - изменений не обнаружено");
            }
            else
            {
                report.AppendLine($"РЕЗУЛЬТАТ: Обнаружено изменений: {totalChanges}");

                if (inventoryChanged)
                {
                    report.AppendLine("  - Изменен инвентарный номер");
                }

                report.AppendLine($"  - Добавлено: {_result.AddedItems.Count}");
                report.AppendLine($"  - Удалено: {_result.RemovedItems.Count}");
                report.AppendLine($"  - Изменено: {_result.ChangedItems.Count + (inventoryChanged ? 1 : 0)}");
            }

            return report.ToString();
        }

        private string GenerateFullComparisonReport()
        {
            var report = new StringBuilder();
            report.AppendLine("ПОЛНЫЙ ОТЧЁТ СРАВНЕНИЯ ПАСПОРТОВ КОМПЬЮТЕРА");
            report.AppendLine("===========================================");
            report.AppendLine();

            string currentInventory = string.IsNullOrEmpty(_currentInfo.InventoryNumber) ? "не указан" : _currentInfo.InventoryNumber;
            string savedInventory = string.IsNullOrEmpty(_savedInfo.InventoryNumber) ? "не указан" : _savedInfo.InventoryNumber;

            report.AppendLine($"Текущий паспорт: {_currentInfo.UniqueId} (Инв.№: {currentInventory}, сформирован: {_currentInfo.CollectedAt:dd.MM.yyyy HH:mm:ss})");
            report.AppendLine($"Сохранённый паспорт: {_savedInfo.UniqueId} (Инв.№: {savedInventory}, сформирован: {_savedInfo.CollectedAt:dd.MM.yyyy HH:mm:ss})");
            report.AppendLine($"Дата сравнения: {DateTime.Now:dd.MM.yyyy HH:mm:ss}");
            report.AppendLine();

            int totalChanges = _result.AddedItems.Count + _result.RemovedItems.Count + _result.ChangedItems.Count;

            // Проверяем изменение инвентарного номера
            bool inventoryChanged = _currentInfo.InventoryNumber != _savedInfo.InventoryNumber;
            if (inventoryChanged)
            {
                totalChanges++;
            }

            if (totalChanges == 0)
            {
                report.AppendLine("РЕЗУЛЬТАТ: Конфигурации идентичны - изменений не обнаружено");
            }
            else
            {
                report.AppendLine($"ОБНАРУЖЕНО ИЗМЕНЕНИЙ: {totalChanges}");
                report.AppendLine();

                // Показываем изменение инвентарного номера (БЕЗ ДУБЛИРОВАНИЯ)
                if (inventoryChanged)
                {
                    report.AppendLine("ИЗМЕНЕН ИНВЕНТАРНЫЙ НОМЕР:");
                    report.AppendLine($"  Было: {savedInventory}");
                    report.AppendLine($"  Стало: {currentInventory}");
                    report.AppendLine();
                }

                if (_result.AddedItems.Count > 0)
                {
                    report.AppendLine("ДОБАВЛЕНО:");
                    foreach (var item in _result.AddedItems)
                        report.AppendLine($"  [+] {item}");
                    report.AppendLine();
                }

                if (_result.RemovedItems.Count > 0)
                {
                    report.AppendLine("УДАЛЕНО:");
                    foreach (var item in _result.RemovedItems)
                        report.AppendLine($"  [-] {item}");
                    report.AppendLine();
                }

                if (_result.ChangedItems.Count > 0)
                {
                    report.AppendLine("ИЗМЕНЕНО:");
                    foreach (var item in _result.ChangedItems)
                        report.AppendLine($"  [*] {item}");
                }
            }

            return report.ToString();
        }

        private string ConvertToHtmlReport(string textContent)
        {
            var html = new StringBuilder();
            html.AppendLine("<!DOCTYPE html>");
            html.AppendLine("<html><head><meta charset='utf-8'>");
            html.AppendLine("<title>Отчёт сравнения паспортов</title>");
            html.AppendLine("<style>");
            html.AppendLine("body { font-family: Arial, sans-serif; margin: 20px; }");
            html.AppendLine(".header { background: #2C2C3C; color: white; padding: 10px; text-align: center; }");
            html.AppendLine(".added { color: #4CAF50; }");
            html.AppendLine(".removed { color: #F44336; }");
            html.AppendLine(".changed { color: #FF9800; }");
            html.AppendLine("</style>");
            html.AppendLine("</head><body>");
            html.AppendLine("<div class='header'><h1>Отчёт сравнения паспортов компьютера</h1></div>");
            html.AppendLine("<pre>" + textContent.Replace("\n", "<br>").Replace(" ", "&nbsp;") + "</pre>");
            html.AppendLine("</body></html>");
            return html.ToString();
        }
    }
}