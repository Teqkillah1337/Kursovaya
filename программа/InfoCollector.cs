using System;
using System.Collections.Generic;
using System.Management;
using System.Net.NetworkInformation;
using Microsoft.Win32;

namespace ComputerPassport
{
    public static class InfoCollector
    {
        public static Exporter.ComputerInfo Collect(bool hardware, bool software, bool drivers, bool peripherals)
        {
            var info = new Exporter.ComputerInfo
            {
                UniqueId = GenerateUniqueId(),
                CollectedAt = DateTime.Now,
                Hardware = new Dictionary<string, string>(),
                Software = new Dictionary<string, string>(),
                Drivers = new Dictionary<string, string>(),
                Peripherals = new Dictionary<string, string>()
            };

            // Собираем только выбранные разделы
            if (hardware) GetHardwareInfo(info.Hardware);
            if (software) GetSoftwareInfo(info.Software);
            if (drivers) GetDriverInfo(info.Drivers);
            if (peripherals) GetPeripheralInfo(info.Peripherals);

            return info;
        }

        private static string GenerateUniqueId()
        {
            try
            {
                var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Processor");
                foreach (var mo in searcher.Get())
                {
                    return $"CPU-{mo["ProcessorId"]}";
                }
            }
            catch { }
            return $"PC-{DateTime.Now:yyyyMMddHHmmss}";
        }

        private static void GetHardwareInfo(Dictionary<string, string> hardware)
        {
            try
            {
                // Процессор
                var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Processor");
                foreach (var mo in searcher.Get())
                {
                    hardware["Процессор"] = $"{mo["Name"]}";
                    hardware["Ядра процессора"] = $"{mo["NumberOfCores"]}";
                    hardware["Потоки процессора"] = $"{mo["NumberOfLogicalProcessors"]}";
                    hardware["Частота процессора"] = $"{mo["MaxClockSpeed"]} MHz";
                    hardware["Сокет процессора"] = $"{mo["SocketDesignation"]}";
                    break;
                }

                // Оперативная память
                searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PhysicalMemory");
                long totalMemory = 0;
                int memoryModules = 0;
                foreach (var mo in searcher.Get())
                {
                    totalMemory += Convert.ToInt64(mo["Capacity"]);
                    memoryModules++;
                    hardware[$"Модуль памяти {memoryModules}"] =
                        $"{Convert.ToInt64(mo["Capacity"]) / (1024 * 1024 * 1024)} GB, {mo["Speed"]} MHz, {mo["Manufacturer"]}";
                }
                hardware["Общий объем памяти"] = $"{totalMemory / (1024 * 1024 * 1024)} GB";
                hardware["Количество модулей памяти"] = $"{memoryModules}";

                // Материнская плата
                searcher = new ManagementObjectSearcher("SELECT * FROM Win32_BaseBoard");
                foreach (var mo in searcher.Get())
                {
                    hardware["Материнская плата"] = $"{mo["Manufacturer"]} {mo["Product"]}";
                    hardware["Версия платы"] = $"{mo["Version"]}";
                    break;
                }

                // BIOS
                searcher = new ManagementObjectSearcher("SELECT * FROM Win32_BIOS");
                foreach (var mo in searcher.Get())
                {
                    hardware["BIOS"] = $"{mo["Manufacturer"]} {mo["SMBIOSBIOSVersion"]}";
                    hardware["Версия BIOS"] = $"{mo["Version"]}";
                    break;
                }

                // Жесткие диски
                searcher = new ManagementObjectSearcher("SELECT * FROM Win32_DiskDrive");
                int diskCount = 0;
                foreach (var mo in searcher.Get())
                {
                    diskCount++;
                    var size = Convert.ToInt64(mo["Size"]) / (1024 * 1024 * 1024);
                    hardware[$"Диск {diskCount}"] =
                        $"{mo["Model"]}, Серийный номер: {mo["SerialNumber"]?.ToString()?.Trim()}, Размер: {size} GB";
                }

                // Видеокарты
                searcher = new ManagementObjectSearcher("SELECT * FROM Win32_VideoController");
                int gpuCount = 0;
                foreach (var mo in searcher.Get())
                {
                    if (mo["Name"] != null && !mo["Name"].ToString().Contains("Microsoft"))
                    {
                        gpuCount++;
                        var memory = Convert.ToInt64(mo["AdapterRAM"]) / (1024 * 1024);
                        hardware[$"Видеокарта {gpuCount}"] =
                            $"{mo["Name"]}, Память: {memory} MB, Драйвер: {mo["DriverVersion"]}";
                    }
                }

                // Сетевые интерфейсы
                searcher = new ManagementObjectSearcher("SELECT * FROM Win32_NetworkAdapter WHERE NetEnabled=True");
                int netCount = 0;
                foreach (var mo in searcher.Get())
                {
                    netCount++;
                    hardware[$"Сетевой адаптер {netCount}"] =
                        $"{mo["Name"]}, MAC: {mo["MACAddress"]}";
                }

                // Звуковые устройства
                searcher = new ManagementObjectSearcher("SELECT * FROM Win32_SoundDevice");
                int soundCount = 0;
                foreach (var mo in searcher.Get())
                {
                    soundCount++;
                    hardware[$"Звуковое устройство {soundCount}"] = $"{mo["Name"]}";
                }
            }
            catch (Exception ex)
            {
                hardware["Ошибка сбора"] = ex.Message;
            }
        }

        private static void GetSoftwareInfo(Dictionary<string, string> software)
        {
            try
            {
                // Операционная система
                var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem");
                foreach (var mo in searcher.Get())
                {
                    software["Операционная система"] = $"{mo["Caption"]}";
                    software["Версия ОС"] = $"{mo["Version"]}";
                    software["Архитектура ОС"] = $"{mo["OSArchitecture"]}";
                    software["Сборка ОС"] = $"{mo["BuildNumber"]}";
                    software["Производитель ОС"] = $"{mo["Manufacturer"]}";
                    break;
                }

                // Установленные программы через реестр (ВСЕ программы)
                GetInstalledPrograms(software);

                // Ключевые программы
                FindKeySoftware(software);
            }
            catch (Exception ex)
            {
                software["Ошибка сбора"] = ex.Message;
            }
        }

        private static void GetInstalledPrograms(Dictionary<string, string> software)
        {
            try
            {
                string[] registryPaths = {
                    @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                    @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
                };

                int programCount = 0;
                var allPrograms = new List<string>();

                foreach (string registryPath in registryPaths)
                {
                    using (RegistryKey key = Registry.LocalMachine.OpenSubKey(registryPath))
                    {
                        if (key != null)
                        {
                            foreach (string subKeyName in key.GetSubKeyNames())
                            {
                                using (RegistryKey subKey = key.OpenSubKey(subKeyName))
                                {
                                    string displayName = subKey?.GetValue("DisplayName") as string;
                                    string publisher = subKey?.GetValue("Publisher") as string;
                                    string displayVersion = subKey?.GetValue("DisplayVersion") as string;

                                    if (!string.IsNullOrEmpty(displayName))
                                    {
                                        programCount++;
                                        string programInfo = displayName;
                                        if (!string.IsNullOrEmpty(displayVersion))
                                            programInfo += $", Версия: {displayVersion}";
                                        if (!string.IsNullOrEmpty(publisher))
                                            programInfo += $", Производитель: {publisher}";

                                        allPrograms.Add(programInfo);
                                    }
                                }
                            }
                        }
                    }
                }

                // Добавляем ВСЕ программы в словарь
                for (int i = 0; i < allPrograms.Count; i++)
                {
                    software[$"Установленная программа {i + 1}"] = allPrograms[i];
                }
            }
            catch (Exception ex)
            {
                software["Ошибка сбора программ"] = ex.Message;
            }
        }

        private static void FindKeySoftware(Dictionary<string, string> software)
        {
            try
            {
                // Поиск браузеров
                string[] browsers = { "chrome", "firefox", "opera", "edge", "safari" };
                foreach (var browser in browsers)
                {
                    var path = Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" + browser + ".exe", "", null);
                    if (path != null)
                    {
                        software["Браузер " + browser.ToUpper()] = "Установлен";
                    }
                }

                // Поиск офисных приложений
                var officePath = Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office", "InstallRoot", null);
                if (officePath != null)
                {
                    software["Microsoft Office"] = "Установлен";
                }

                // Антивирусы
                CheckAntivirus(software);
            }
            catch { }
        }

        private static void CheckAntivirus(Dictionary<string, string> software)
        {
            try
            {
                var searcher = new ManagementObjectSearcher(@"root\SecurityCenter2", "SELECT * FROM AntiVirusProduct");
                foreach (var mo in searcher.Get())
                {
                    software["Антивирус"] = $"{mo["displayName"]}";
                    break;
                }
            }
            catch { }
        }

        private static void GetDriverInfo(Dictionary<string, string> drivers)
        {
            try
            {
                // ВСЕ драйверы устройств
                var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PnPSignedDriver WHERE DeviceClass IS NOT NULL");
                int driverCount = 0;
                foreach (var mo in searcher.Get())
                {
                    driverCount++;
                    var deviceName = mo["DeviceName"]?.ToString();
                    var driverVersion = mo["DriverVersion"]?.ToString();

                    if (!string.IsNullOrEmpty(deviceName))
                    {
                        drivers[$"Драйвер {driverCount}"] =
                            $"{deviceName}, Версия: {driverVersion}";
                    }
                }
            }
            catch (Exception ex)
            {
                drivers["Ошибка сбора"] = ex.Message;
            }
        }

        private static void GetPeripheralInfo(Dictionary<string, string> peripherals)
        {
            try
            {
                // USB устройства
                var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_USBHub");
                int usbCount = 0;
                foreach (var mo in searcher.Get())
                {
                    usbCount++;
                    peripherals[$"USB устройство {usbCount}"] = $"{mo["Name"]}";
                }

                // Принтеры
                searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Printer");
                int printerCount = 0;
                foreach (var mo in searcher.Get())
                {
                    printerCount++;
                    peripherals[$"Принтер {printerCount}"] = $"{mo["Name"]}";
                }

                // Мыши и клавиатуры
                searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PointingDevice");
                int mouseCount = 0;
                foreach (var mo in searcher.Get())
                {
                    mouseCount++;
                    peripherals[$"Устройство ввода {mouseCount}"] = $"{mo["Name"]}";
                }

                // Мониторы
                searcher = new ManagementObjectSearcher("SELECT * FROM Win32_DesktopMonitor");
                int monitorCount = 0;
                foreach (var mo in searcher.Get())
                {
                    monitorCount++;
                    peripherals[$"Монитор {monitorCount}"] = $"{mo["Name"]}";
                }
            }
            catch (Exception ex)
            {
                peripherals["Ошибка сбора"] = ex.Message;
            }
        }
    }
}