using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace ComputerPassport
{
    public partial class MainWindow : Window
    {
        private Exporter.ComputerInfo currentInfo;
        private BackgroundWorker collectionWorker;
        private bool isCollectionInProgress;
        private bool isComparisonViewActive = false;

        public MainWindow()
        {
            InitializeComponent();
            InitializeBackgroundWorker();

            // Автоматически определяем папку для сохранения
            InitializeSaveFolder();

            // Автоматически выбрать первый элемент в комбобоксе
            if (cmbExportFormat.Items.Count > 0)
                cmbExportFormat.SelectedIndex = 0;

            UpdateControlsState();
        }

        private void InitializeSaveFolder()
        {
            try
            {
                // Получаем путь к папке "Документы"
                string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                // Создаем подпапку "ПаспортыПК" в Документах
                string saveFolder = Path.Combine(documentsPath, "ПаспортыПК");

                // Устанавливаем путь в текстовое поле
                txtSaveFolder.Text = saveFolder;

                // Создаем папку если не существует
                if (!Directory.Exists(saveFolder))
                {
                    Directory.CreateDirectory(saveFolder);
                }

                CheckSaveFolder();
            }
            catch (Exception)
            {
                // Если не удалось, используем папку приложения
                string appPath = AppDomain.CurrentDomain.BaseDirectory;
                txtSaveFolder.Text = Path.Combine(appPath, "ПаспортыПК");
                CheckSaveFolder();

                MessageBox.Show("Не удалось создать папку в Документах.\n" +
                               $"Файлы будут сохраняться в: {txtSaveFolder.Text}",
                               "Информация");
            }
        }

        private void InitializeBackgroundWorker()
        {
            collectionWorker = new BackgroundWorker();
            collectionWorker.DoWork += CollectionWorker_DoWork;
            collectionWorker.RunWorkerCompleted += CollectionWorker_RunWorkerCompleted;
            collectionWorker.WorkerReportsProgress = false;
        }

        private void BtnBack_Click(object sender, RoutedEventArgs e)
        {
            ShowMainView();
        }

        private void ShowMainView()
        {
            mainContent.Visibility = Visibility.Visible;
            comparisonContent.Visibility = Visibility.Collapsed;
            aboutContent.Visibility = Visibility.Collapsed;
            btnBack.Visibility = Visibility.Collapsed;
            txtNavigationTitle.Text = "Паспорт Компьютера";
            isComparisonViewActive = false;

            // Очищаем frames
            comparisonFrame.Content = null;
            aboutFrame.Content = null;
        }

        private void ShowComparisonView()
        {
            mainContent.Visibility = Visibility.Collapsed;
            comparisonContent.Visibility = Visibility.Visible;
            aboutContent.Visibility = Visibility.Collapsed;
            btnBack.Visibility = Visibility.Visible;
            txtNavigationTitle.Text = "Сравнение паспортов";
            isComparisonViewActive = true;
        }

        private void ShowAboutView()
        {
            mainContent.Visibility = Visibility.Collapsed;
            comparisonContent.Visibility = Visibility.Collapsed;
            aboutContent.Visibility = Visibility.Visible;
            btnBack.Visibility = Visibility.Visible;
            txtNavigationTitle.Text = "О программе";

            // Загружаем страницу "О программе"
            aboutFrame.Content = new AboutPage();
        }

        private void BtnAbout_Click(object sender, RoutedEventArgs e)
        {
            ShowAboutView();
        }

        private void BtnBrowseFolder_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                ValidateNames = false,
                CheckFileExists = false,
                CheckPathExists = true,
                FileName = "Выберите папку"
            };

            if (Directory.Exists(txtSaveFolder.Text))
            {
                dialog.InitialDirectory = txtSaveFolder.Text;
            }
            else
            {
                dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            }

            if (dialog.ShowDialog() == true)
            {
                string selectedPath = Path.GetDirectoryName(dialog.FileName);
                if (!string.IsNullOrEmpty(selectedPath))
                {
                    txtSaveFolder.Text = selectedPath;
                    CheckSaveFolder();
                }
            }
        }

        private void CheckSaveFolder()
        {
            try
            {
                if (Directory.Exists(txtSaveFolder.Text))
                {
                    txtFolderStatus.Text = "✅ Папка доступна для сохранения";
                    txtFolderStatus.Foreground = new SolidColorBrush(Color.FromArgb(0xFF, 0x33, 0x33, 0x33));

                    // Показываем информацию о свободном месте
                    DriveInfo drive = new DriveInfo(Path.GetPathRoot(txtSaveFolder.Text));
                    if (drive.IsReady)
                    {
                        long freeSpaceGB = drive.AvailableFreeSpace / (1024 * 1024 * 1024);
                        txtFolderStatus.Text += $" (Свободно: {freeSpaceGB} GB)";
                    }
                }
                else
                {
                    txtFolderStatus.Text = "⚠ Папка не существует - будет создана автоматически";
                    txtFolderStatus.Foreground = System.Windows.Media.Brushes.Orange;
                }
            }
            catch
            {
                txtFolderStatus.Text = "❌ Ошибка проверки папки";
                txtFolderStatus.Foreground = System.Windows.Media.Brushes.LightCoral;
            }
        }

        private void BtnCollect_Click(object sender, RoutedEventArgs e)
        {
            if (isCollectionInProgress)
            {
                MessageBox.Show("Сбор информации уже выполняется...");
                return;
            }

            // ЗАПРАШИВАЕМ ИНВЕНТАРНЫЙ НОМЕР ПЕРЕД СБОРОМ
            var inventoryDialog = new InventoryNumberDialog();
            inventoryDialog.Owner = this;
            inventoryDialog.WindowStartupLocation = WindowStartupLocation.CenterOwner;

            if (inventoryDialog.ShowDialog() == true)
            {
                string inventoryNumber = inventoryDialog.InventoryNumber;

                txtOutput.Text = "Идет сбор системной информации...\nПожалуйста, подождите...";
                isCollectionInProgress = true;
                UpdateControlsState();

                // Запускаем сбор в фоновом потоке с инвентарным номером
                var hardware = cbHardware.IsChecked == true;
                var software = cbSoftware.IsChecked == true;
                var drivers = cbDrivers.IsChecked == true;
                var peripherals = cbPeripherals.IsChecked == true;

                collectionWorker.RunWorkerAsync(new object[] { hardware, software, drivers, peripherals, inventoryNumber });
            }
        }

        private void CollectionWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                if (e.Argument is object[] parameters && parameters.Length == 5)
                {
                    bool hardware = (bool)parameters[0];
                    bool software = (bool)parameters[1];
                    bool drivers = (bool)parameters[2];
                    bool peripherals = (bool)parameters[3];
                    string inventoryNumber = (string)parameters[4];

                    // Выполняем сбор и устанавливаем инвентарный номер
                    var result = InfoCollector.Collect(hardware, software, drivers, peripherals);
                    result.InventoryNumber = inventoryNumber; // ДОБАВЛЯЕМ ИНВЕНТАРНЫЙ НОМЕР
                    e.Result = result;
                }
                else
                {
                    // По умолчанию собираем все без инвентарного номера
                    e.Result = InfoCollector.Collect(true, true, true, true);
                }
            }
            catch (Exception ex)
            {
                e.Result = ex;
            }
        }

        private void CollectionWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            isCollectionInProgress = false;

            if (e.Error != null)
            {
                txtOutput.Text = $"Произошла ошибка:\n{e.Error.Message}\n\n{e.Error.StackTrace}";
                currentInfo = null;
            }
            else if (e.Result is Exception exception)
            {
                txtOutput.Text = $"Произошла ошибка:\n{exception.Message}\n\n{exception.StackTrace}";
                currentInfo = null;
            }
            else
            {
                currentInfo = e.Result as Exporter.ComputerInfo;

                if (currentInfo != null)
                {
                    DisplayCollectedInfo();
                }
                else
                {
                    txtOutput.Text = "Не удалось собрать информацию о системе.";
                }
            }

            UpdateControlsState();
        }

        private void UpdateControlsState()
        {
            // Всегда доступны
            btnCollect.IsEnabled = !isCollectionInProgress;
            btnBrowseFolder.IsEnabled = !isCollectionInProgress;

            // Зависят от наличия данных
            bool hasData = currentInfo != null && !isCollectionInProgress;
            btnExport.IsEnabled = hasData && cmbExportFormat.SelectedItem != null;
            cmbExportFormat.IsEnabled = hasData;
            btnCompare.IsEnabled = hasData && !isCollectionInProgress;

            // Чекбоксы
            cbHardware.IsEnabled = !isCollectionInProgress;
            cbSoftware.IsEnabled = !isCollectionInProgress;
            cbDrivers.IsEnabled = !isCollectionInProgress;
            cbPeripherals.IsEnabled = !isCollectionInProgress;
        }

        private void DisplayCollectedInfo()
        {
            if (currentInfo == null) return;

            var displayText = new StringBuilder();
            displayText.AppendLine($"ПАСПОРТ КОМПЬЮТЕРА № {currentInfo.UniqueId}");

            // Показываем инвентарный номер
            if (!string.IsNullOrEmpty(currentInfo.InventoryNumber))
            {
                displayText.AppendLine($"Инвентарный номер: {currentInfo.InventoryNumber}");
            }

            displayText.AppendLine($"Дата формирования: {currentInfo.CollectedAt:dd.MM.yyyy HH:mm:ss}");
            displayText.AppendLine();

            // Показываем аппаратную конфигурацию
            if (cbHardware.IsChecked == true && currentInfo.Hardware.Count > 0)
            {
                displayText.AppendLine("=== АППАРАТНАЯ КОНФИГУРАЦИЯ ===");
                foreach (var kv in currentInfo.Hardware)
                {
                    displayText.AppendLine($"{kv.Key}: {kv.Value}");
                }
                displayText.AppendLine();
            }

            // Показываем программное обеспечение
            if (cbSoftware.IsChecked == true && currentInfo.Software.Count > 0)
            {
                displayText.AppendLine("=== ПРОГРАММНОЕ ОБЕСПЕЧЕНИЕ ===");
                foreach (var kv in currentInfo.Software)
                {
                    displayText.AppendLine($"{kv.Key}: {kv.Value}");
                }
                displayText.AppendLine();
            }

            // Показываем драйверы
            if (cbDrivers.IsChecked == true && currentInfo.Drivers.Count > 0)
            {
                displayText.AppendLine("=== ДРАЙВЕРЫ УСТРОЙСТВ ===");
                foreach (var kv in currentInfo.Drivers)
                {
                    displayText.AppendLine($"{kv.Key}: {kv.Value}");
                }
                displayText.AppendLine();
            }

            // Показываем периферийные устройства
            if (cbPeripherals.IsChecked == true && currentInfo.Peripherals.Count > 0)
            {
                displayText.AppendLine("=== ПЕРИФЕРИЙНЫЕ УСТРОЙСТВА ===");
                foreach (var kv in currentInfo.Peripherals)
                {
                    displayText.AppendLine($"{kv.Key}: {kv.Value}");
                }
                displayText.AppendLine();
            }

            txtOutput.Text = displayText.ToString();
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            if (!ValidateBeforeExport()) return;

            if (cmbExportFormat.SelectedItem == null)
            {
                MessageBox.Show("Выберите формат экспорта.", "Формат не выбран");
                return;
            }

            // Инвентарный номер уже есть в currentInfo - просто экспортируем
            ExportWithInventoryNumber(currentInfo.InventoryNumber);
        }
        private void ExportWithInventoryNumber(string inventoryNumber)
        {
            var selectedItem = cmbExportFormat.SelectedItem as ComboBoxItem;
            string format = selectedItem?.Tag as string;

            try
            {
                string resultPath = null;

                switch (format)
                {
                    case "HTML":
                        resultPath = Exporter.ExportToHtml(currentInfo,
                            cbHardware.IsChecked == true,
                            cbSoftware.IsChecked == true,
                            cbDrivers.IsChecked == true,
                            cbPeripherals.IsChecked == true,
                            txtSaveFolder.Text);
                        break;

                    case "PDF":
                        resultPath = Exporter.ExportToPdf(currentInfo,
                            cbHardware.IsChecked == true,
                            cbSoftware.IsChecked == true,
                            cbDrivers.IsChecked == true,
                            cbPeripherals.IsChecked == true,
                            txtSaveFolder.Text);
                        break;

                    case "JSON":
                        var fileName = GenerateFileName("паспорт", inventoryNumber, "json");
                        var fullPath = Path.Combine(txtSaveFolder.Text, fileName);
                        var options = new JsonSerializerOptions { WriteIndented = true };
                        string json = JsonSerializer.Serialize(currentInfo, options);
                        File.WriteAllText(fullPath, json);
                        resultPath = fullPath;
                        break;
                }

                if (resultPath != null)
                {
                    if (MessageBox.Show($"Файл успешно экспортирован:\n{resultPath}\n\nХотите открыть папку с файлом?",
                        "Экспорт завершен", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    {
                        System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{resultPath}\"");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте: {ex.Message}", "Ошибка экспорта");
            }
        }

        private string GenerateFileName(string baseName, string inventoryNumber, string extension)
        {
            var fileName = new StringBuilder(baseName);

            if (!string.IsNullOrEmpty(inventoryNumber))
            {
                fileName.Append($"_{inventoryNumber}");
            }

            fileName.Append($"_{currentInfo.UniqueId}_{DateTime.Now:yyyyMMddHHmmss}.{extension}");
            return fileName.ToString();
        }

        private void CmbExportFormat_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateControlsState();
        }

        private void BtnCompare_Click(object sender, RoutedEventArgs e)
        {
            if (currentInfo == null)
            {
                MessageBox.Show("Сначала соберите информацию о системе.");
                return;
            }

            var openFileDialog = new OpenFileDialog
            {
                Filter = "JSON файлы (*.json)|*.json",
                InitialDirectory = txtSaveFolder.Text
            };

            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    string json = File.ReadAllText(openFileDialog.FileName);
                    var savedInfo = JsonSerializer.Deserialize<Exporter.ComputerInfo>(json);

                    // Выполняем сравнение - currentInfo уже содержит инвентарный номер
                    var comparisonResult = ComparePassports(currentInfo, savedInfo);

                    var comparisonPage = new ComparisonPage(comparisonResult, currentInfo, savedInfo);
                    comparisonFrame.Content = comparisonPage;
                    ShowComparisonView();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при сравнении: {ex.Message}", "Ошибка");
                }
            }
        }

        private bool ValidateBeforeExport()
        {
            if (currentInfo == null)
            {
                MessageBox.Show("Сначала соберите информацию о системе.", "Нет данных");
                return false;
            }

            if (string.IsNullOrWhiteSpace(txtSaveFolder.Text))
            {
                MessageBox.Show("Укажите папку для сохранения.", "Папка не указана");
                return false;
            }

            // Создаем папку если не существует
            if (!Directory.Exists(txtSaveFolder.Text))
            {
                try
                {
                    Directory.CreateDirectory(txtSaveFolder.Text);
                    CheckSaveFolder();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Не удалось создать папку: {ex.Message}", "Ошибка");
                    return false;
                }
            }

            return true;
        }

        // Обновляем состояние при изменении чекбоксов
        private void CheckBox_CheckedChanged(object sender, RoutedEventArgs e)
        {
            if (currentInfo != null && !isCollectionInProgress)
            {
                DisplayCollectedInfo();
            }
        }

        #region Классы и методы для сравнения

        private ComparisonResult ComparePassports(Exporter.ComputerInfo current, Exporter.ComputerInfo saved)
        {
            var result = new ComparisonResult
            {
                Title = $"Сравнение паспортов:\nТекущий: {current.UniqueId}\nСохранённый: {saved.UniqueId}"
            };

            // Проверяем изменение инвентарного номера
            if (current.InventoryNumber != saved.InventoryNumber)
            {
                string oldInventory = string.IsNullOrEmpty(saved.InventoryNumber) ? "не указан" : saved.InventoryNumber;
                string newInventory = string.IsNullOrEmpty(current.InventoryNumber) ? "не указан" : current.InventoryNumber;
                result.ChangedItems.Add($"[Реквизиты] ИЗМЕНЕН ИНВЕНТАРНЫЙ НОМЕР\n  Было: {oldInventory}\n  Стало: {newInventory}");
            }

            // Сравниваем аппаратную конфигурацию
            CompareDictionaries(current.Hardware, saved.Hardware, "Аппаратная конфигурация", result);

            // Сравниваем программное обеспечение
            CompareDictionaries(current.Software, saved.Software, "Программное обеспечение", result);

            // Сравниваем драйверы
            CompareDictionaries(current.Drivers, saved.Drivers, "Драйверы", result);

            // Сравниваем периферийные устройства
            CompareDictionaries(current.Peripherals, saved.Peripherals, "Периферийные устройства", result);

            return result;
        }

        private void CompareDictionaries(Dictionary<string, string> currentDict, Dictionary<string, string> savedDict, string sectionName, ComparisonResult result)
        {
            // Ищем добавленные элементы
            foreach (var item in currentDict)
            {
                if (!savedDict.ContainsKey(item.Key))
                {
                    result.AddedItems.Add($"[{sectionName}] ДОБАВЛЕНО: {item.Key} = {item.Value}");
                }
                else if (savedDict[item.Key] != item.Value)
                {
                    result.ChangedItems.Add($"[{sectionName}] ИЗМЕНЕНО: {item.Key}\n  Было: {savedDict[item.Key]}\n  Стало: {item.Value}");
                }
            }

            // Ищем удаленные элементы
            foreach (var item in savedDict)
            {
                if (!currentDict.ContainsKey(item.Key))
                {
                    result.RemovedItems.Add($"[{sectionName}] УДАЛЕНО: {item.Key} = {item.Value}");
                }
            }
        }

        #endregion
    }
}