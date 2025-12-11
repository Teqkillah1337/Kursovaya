using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace ComputerPassport
{
    public static class Exporter
    {
        public class ComputerInfo
        {
            public string UniqueId { get; set; }
            public DateTime CollectedAt { get; set; }
            public string InventoryNumber { get; set; } // Новое свойство для инвентарного номера
            public Dictionary<string, string> Hardware { get; set; }
            public Dictionary<string, string> Software { get; set; }
            public Dictionary<string, string> Drivers { get; set; }
            public Dictionary<string, string> Peripherals { get; set; }
        }

        public static string ExportToHtml(ComputerInfo info, bool h = true, bool s = true, bool d = true, bool p = true, string folderPath = null)
        {
            var sb = new StringBuilder();

            sb.AppendLine("<!DOCTYPE html>");
            sb.AppendLine("<html><head><meta charset='utf-8'>");
            sb.AppendLine("<title>Паспорт компьютера</title>");
            sb.AppendLine("<style>");
            sb.AppendLine("body { font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }");
            sb.AppendLine(".passport { background: white; border: 2px solid #333; padding: 20px; max-width: 1000px; margin: 0 auto; }");
            sb.AppendLine(".header { text-align: center; border-bottom: 2px solid #333; padding-bottom: 10px; margin-bottom: 20px; }");
            sb.AppendLine(".section { margin-bottom: 20px; }");
            sb.AppendLine(".section h2 { background: #333; color: white; padding: 5px 10px; margin: 0; }");
            sb.AppendLine(".section table { width: 100%; border-collapse: collapse; }");
            sb.AppendLine(".section td { padding: 8px; border: 1px solid #ddd; vertical-align: top; }");
            sb.AppendLine(".section td:first-child { font-weight: bold; width: 30%; background: #f9f9f9; }");
            sb.AppendLine(".footer { text-align: center; margin-top: 30px; font-style: italic; }");
            sb.AppendLine(".inventory { background: #e8f5e8; padding: 5px 10px; border-radius: 4px; display: inline-block; margin-top: 5px; }");
            sb.AppendLine("</style>");
            sb.AppendLine("</head><body>");

            sb.AppendLine("<div class='passport'>");

            sb.AppendLine("<div class='header'>");
            sb.AppendLine("<h1>ПАСПОРТ КОМПЬЮТЕРА</h1>");

            // Добавляем инвентарный номер в заголовок
            if (!string.IsNullOrEmpty(info.InventoryNumber))
            {
                sb.AppendLine($"<div class='inventory'>Инвентарный номер: {info.InventoryNumber}</div>");
            }

            sb.AppendLine($"<h2>№ {info.UniqueId} от {info.CollectedAt:dd.MM.yyyy}</h2>");
            sb.AppendLine("</div>");

            if (h && info.Hardware.Count > 0)
            {
                sb.AppendLine("<div class='section'>");
                sb.AppendLine("<h2>АППАРАТНАЯ КОНФИГУРАЦИЯ</h2>");
                sb.AppendLine("<table>");
                foreach (var kv in info.Hardware)
                {
                    sb.AppendLine($"<tr><td>{kv.Key}</td><td>{kv.Value}</td></tr>");
                }
                sb.AppendLine("</table>");
                sb.AppendLine("</div>");
            }

            if (s && info.Software.Count > 0)
            {
                sb.AppendLine("<div class='section'>");
                sb.AppendLine("<h2>ПРОГРАММНОЕ ОБЕСПЕЧЕНИЕ</h2>");
                sb.AppendLine("<table>");
                foreach (var kv in info.Software)
                {
                    sb.AppendLine($"<tr><td>{kv.Key}</td><td>{kv.Value}</td></tr>");
                }
                sb.AppendLine("</table>");
                sb.AppendLine("</div>");
            }

            if (d && info.Drivers.Count > 0)
            {
                sb.AppendLine("<div class='section'>");
                sb.AppendLine("<h2>ДРАЙВЕРЫ УСТРОЙСТВ</h2>");
                sb.AppendLine("<table>");
                foreach (var kv in info.Drivers)
                {
                    sb.AppendLine($"<tr><td>{kv.Key}</td><td>{kv.Value}</td></tr>");
                }
                sb.AppendLine("</table>");
                sb.AppendLine("</div>");
            }

            if (p && info.Peripherals.Count > 0)
            {
                sb.AppendLine("<div class='section'>");
                sb.AppendLine("<h2>ПЕРИФЕРИЙНЫЕ УСТРОЙСТВА</h2>");
                sb.AppendLine("<table>");
                foreach (var kv in info.Peripherals)
                {
                    sb.AppendLine($"<tr><td>{kv.Key}</td><td>{kv.Value}</td></tr>");
                }
                sb.AppendLine("</table>");
                sb.AppendLine("</div>");
            }

            sb.AppendLine("<div class='footer'>");
            sb.AppendLine($"<p>Паспорт сформирован: {info.CollectedAt:dd.MM.yyyy HH:mm:ss}</p>");
            sb.AppendLine("</div>");

            sb.AppendLine("</div>");
            sb.AppendLine("</body></html>");

            // Генерация имени файла с инвентарным номером
            string fileName;
            if (!string.IsNullOrEmpty(info.InventoryNumber))
            {
                fileName = $"паспорт_{info.InventoryNumber}_{info.UniqueId}_{DateTime.Now:yyyyMMddHHmmss}.html";
            }
            else
            {
                fileName = $"паспорт_{info.UniqueId}_{DateTime.Now:yyyyMMddHHmmss}.html";
            }

            if (string.IsNullOrEmpty(folderPath))
            {
                folderPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            }

            if (!Directory.Exists(folderPath))
            {
                try
                {
                    Directory.CreateDirectory(folderPath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Не удалось создать папку: {ex.Message}");
                    return null;
                }
            }

            var fullPath = Path.Combine(folderPath, fileName);

            try
            {
                File.WriteAllText(fullPath, sb.ToString(), Encoding.UTF8);
                return fullPath;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении HTML: {ex.Message}");
                return null;
            }
        }

        public static string ExportToPdf(ComputerInfo info, bool h = true, bool s = true, bool d = true, bool p = true, string folderPath = null)
        {
            // Генерация имени файла с инвентарным номером
            string fileName;
            if (!string.IsNullOrEmpty(info.InventoryNumber))
            {
                fileName = $"паспорт_{info.InventoryNumber}_{info.UniqueId}_{DateTime.Now:yyyyMMddHHmmss}.pdf";
            }
            else
            {
                fileName = $"паспорт_{info.UniqueId}_{DateTime.Now:yyyyMMddHHmmss}.pdf";
            }

            if (string.IsNullOrEmpty(folderPath))
            {
                folderPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            }

            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            var fullPath = Path.Combine(folderPath, fileName);

            try
            {
                // Регистрируем шрифты для поддержки русского языка
                RegisterRussianFonts();

                using (var fs = new FileStream(fullPath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    var document = new Document(PageSize.A4, 40, 40, 40, 40);
                    var writer = PdfWriter.GetInstance(document, fs);

                    document.Open();

                    // Создаем стилизованный PDF похожий на HTML
                    AddPdfHeader(document, info);

                    if (h && info.Hardware.Count > 0)
                    {
                        AddPdfSection(document, "АППАРАТНАЯ КОНФИГУРАЦИЯ", info.Hardware);
                    }

                    if (s && info.Software.Count > 0)
                    {
                        AddPdfSection(document, "ПРОГРАММНОЕ ОБЕСПЕЧЕНИЕ", info.Software);
                    }

                    if (d && info.Drivers.Count > 0)
                    {
                        AddPdfSection(document, "ДРАЙВЕРЫ УСТРОЙСТВ", info.Drivers);
                    }

                    if (p && info.Peripherals.Count > 0)
                    {
                        AddPdfSection(document, "ПЕРИФЕРИЙНЫЕ УСТРОЙСТВА", info.Peripherals);
                    }

                    AddPdfFooter(document, info);

                    document.Close();
                }

                return fullPath;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при создании PDF: {ex.Message}\n\nПопробуйте использовать HTML экспорт.", "Ошибка");
                return null;
            }
        }

        private static void RegisterRussianFonts()
        {
            // Регистрируем системные шрифты для поддержки русского
            if (!FontFactory.IsRegistered("ArialUnicode"))
            {
                // Пытаемся найти Arial Unicode MS или используем стандартный шрифт
                string[] fontPaths = {
                    @"C:\Windows\Fonts\arial.ttf",
                    @"C:\Windows\Fonts\arialuni.ttf",
                    @"C:\Windows\Fonts\times.ttf"
                };

                foreach (string fontPath in fontPaths)
                {
                    if (File.Exists(fontPath))
                    {
                        FontFactory.Register(fontPath, "ArialUnicode");
                        break;
                    }
                }
            }
        }

        private static void AddPdfHeader(Document document, ComputerInfo info)
        {
            // Основной заголовок
            var titleFont = FontFactory.GetFont("ArialUnicode", BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 20, Font.BOLD, BaseColor.BLACK);
            var title = new Paragraph("ПАСПОРТ КОМПЬЮТЕРА", titleFont)
            {
                Alignment = Element.ALIGN_CENTER,
                SpacingAfter = 10f
            };
            document.Add(title);

            // Инвентарный номер если указан
            if (!string.IsNullOrEmpty(info.InventoryNumber))
            {
                var inventoryFont = FontFactory.GetFont("ArialUnicode", BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 12, Font.BOLD, BaseColor.DARK_GRAY);
                var inventory = new Paragraph($"Инвентарный номер: {info.InventoryNumber}", inventoryFont)
                {
                    Alignment = Element.ALIGN_CENTER,
                    SpacingAfter = 5f
                };
                document.Add(inventory);
            }

            // Подзаголовок с номером и датой
            var subtitleFont = FontFactory.GetFont("ArialUnicode", BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 14, Font.BOLD, BaseColor.DARK_GRAY);
            var subtitle = new Paragraph($"№ {info.UniqueId} от {info.CollectedAt:dd.MM.yyyy}", subtitleFont)
            {
                Alignment = Element.ALIGN_CENTER,
                SpacingAfter = 25f
            };
            document.Add(subtitle);
        }

        private static void AddPdfSection(Document document, string sectionTitle, Dictionary<string, string> data)
        {
            // Заголовок секции с серым фоном как в HTML
            var sectionFont = FontFactory.GetFont("ArialUnicode", BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 12, Font.BOLD, BaseColor.WHITE);
            var sectionHeader = new Paragraph(sectionTitle, sectionFont)
            {
                SpacingBefore = 20f,
                SpacingAfter = 10f
            };

            // Создаем ячейку с серым фоном для заголовка
            var headerCell = new PdfPCell(sectionHeader)
            {
                BackgroundColor = new BaseColor(51, 51, 51), // Темно-серый как в HTML
                BorderWidth = 0,
                Padding = 8f,
                HorizontalAlignment = Element.ALIGN_LEFT
            };

            var headerTable = new PdfPTable(1)
            {
                WidthPercentage = 100
            };
            headerTable.AddCell(headerCell);
            document.Add(headerTable);

            // Таблица с данными
            var dataTable = new PdfPTable(2)
            {
                WidthPercentage = 100
            };
            dataTable.SetWidths(new float[] { 35, 65 });

            var boldFont = FontFactory.GetFont("ArialUnicode", BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 10, Font.BOLD, BaseColor.BLACK);
            var normalFont = FontFactory.GetFont("ArialUnicode", BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 10, Font.NORMAL, BaseColor.BLACK);

            foreach (var kv in data)
            {
                // Ячейка с названием параметра (светло-серый фон)
                var nameCell = new PdfPCell(new Phrase(kv.Key, boldFont))
                {
                    BorderWidth = 0.5f,
                    BorderColor = new BaseColor(221, 221, 221),
                    Padding = 8f,
                    BackgroundColor = new BaseColor(249, 249, 249), // Светло-серый как в HTML
                    HorizontalAlignment = Element.ALIGN_LEFT
                };
                dataTable.AddCell(nameCell);

                // Ячейка со значением
                var valueCell = new PdfPCell(new Phrase(kv.Value ?? "", normalFont))
                {
                    BorderWidth = 0.5f,
                    BorderColor = new BaseColor(221, 221, 221),
                    Padding = 8f,
                    HorizontalAlignment = Element.ALIGN_LEFT
                };
                dataTable.AddCell(valueCell);
            }

            document.Add(dataTable);
        }

        private static void AddPdfFooter(Document document, ComputerInfo info)
        {
            // Добавляем отступ
            document.Add(new Paragraph(" ") { SpacingBefore = 20f });

            var footerFont = FontFactory.GetFont("ArialUnicode", BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 10, Font.ITALIC, BaseColor.DARK_GRAY);
            var footer = new Paragraph($"Паспорт сформирован: {info.CollectedAt:dd.MM.yyyy HH:mm:ss}", footerFont)
            {
                Alignment = Element.ALIGN_CENTER,
                SpacingBefore = 10f
            };
            document.Add(footer);
        }

        private static string Escape(object o)
        {
            if (o == null) return "";
            return o.ToString()
                .Replace("\\", "\\\\")
                .Replace("{", "\\{")
                .Replace("}", "\\}")
                .Replace("\r", "")
                .Replace("\n", "\\par ");
        }
    }
}