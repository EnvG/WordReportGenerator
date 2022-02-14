using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Linq;

namespace WordReportGenerator
{
    public class Report
    {
        /// <summary>
        /// Тело документа
        /// </summary>
        private Body body;

        public Report()
        {
            // Инициализация тела документа
            this.body = new Body();
        }

        #region Методы манипуляции с документом
        /// <summary>
        /// Добавить параграф в документ
        /// </summary>
        /// <param name="content">Содержание параграфа</param>
        public void AddParagraph (string content)
        {
            // Получение параграфа с заданным содержанием
            Paragraph paragraph = this.GetParagraph(content);

            // Добавление параграфа в тело документа
            this.body.Append(paragraph);
        }
        /// <summary>
        /// Добавить таблицу в документ
        /// </summary>
        /// <param name="content">Таблица, добавляемая в документ</param>
        /// <param name="width">Ширина столбца</param>
        public void AddTable<T>(List<T[]> content, string width = "500")
        {
            // Количество строк таблицы
            int rowsCount = content.Count;
            // Количество столбцов таблицы
            int columnsCount = content.First().Length;
            Table table = new Table();

            // Границы таблицы
            TableProperties borders = this.GetTableBorder();
            table.Append(borders);

            for (int i = 0; i < rowsCount; i++)
            {
                TableRow row = new TableRow();
                for (int j = 0; j < columnsCount; j++)
                {
                    TableCell cell = new TableCell();
                    cell.AppendChild(new TableCellWidth() { Width = $"{width}", Type = TableWidthUnitValues.Pct });
                    // Значение в ячейке
                    Paragraph value = this.GetParagraph(content[i][j].ToString());

                    // Добавление значения в ячейку
                    cell.Append(value);
                    // Добавление ячейки в строку
                    row.Append(cell);
                }

                // Добавление строки в таблицу
                table.Append(row);
            }

            this.body.Append(table);
        }
        /// <summary>
        /// Добавить параграф в документ
        /// </summary>
        /// <param name="runs">Текст, из которого состоит параграф</param>
        public void AddParagraph(params Run[] runs)
        {
            Paragraph paragraph = new Paragraph();
            foreach (Run run in runs)
            {
                paragraph.Append(run);
            }

            this.body.Append(paragraph);
        }
        #endregion
        /// <summary>
        /// Получить текст для параграфа
        /// </summary>
        /// <param name="content">Содержание</param>
        /// <param name="verticalPosition">Вертикальное выравнивание</param>
        /// <returns></returns>
        public static Run GetRun(string content, VerticalPositionValues verticalPosition = VerticalPositionValues.Baseline)
        {
            Run run = new Run(new RunProperties(new VerticalTextAlignment { Val = verticalPosition }), new Text(content) { Space = SpaceProcessingModeValues.Preserve });

            return run;
        }
        /// <summary>
        /// Получение параграфа с данным содержанием
        /// </summary>
        /// <param name="content">Содержание параграфа</param>
        /// <returns></returns>
        private Paragraph GetParagraph(string content, VerticalPositionValues verticalPosition = VerticalPositionValues.Baseline)
        {
            Paragraph paragraph = new Paragraph();
            Run run = new Run(new RunProperties(new VerticalTextAlignment { Val = verticalPosition }));
            Text text = new Text(content)
            {
                Space = SpaceProcessingModeValues.Preserve
            };

            // Добавление текста в набор
            run.Append(text);
            // Добавление набора текста в параграф
            paragraph.Append(run);

            return paragraph;
        }
        /// <summary>
        /// Получение свойств с границами для таблицы
        /// </summary>
        /// <param name="color">Цвет границ (по умолчанию чёрный)</param>
        /// <returns></returns>
        private TableProperties GetTableBorder(string color = "#000000")
        {
            // Свойства
            TableProperties properties = new TableProperties();

            // Границы
            TableBorders borders = new TableBorders();
            #region Задание границ
            // Добавление левой границы
            borders.Append(new LeftBorder()
            {
                Val = new EnumValue<BorderValues>(BorderValues.Thick),
                Color = "#000000",
            });
            // Добавление верхней границы
            borders.Append(new TopBorder()
            {
                Val = new EnumValue<BorderValues>(BorderValues.Thick),
                Color = "#000000",
            });
            // Добавление правой границы
            borders.Append(new RightBorder()
            {
                Val = new EnumValue<BorderValues>(BorderValues.Thick),
                Color = "#000000",
            });
            // Добавление нижней границы
            borders.Append(new BottomBorder()
            {
                Val = new EnumValue<BorderValues>(BorderValues.Thick),
                Color = "#000000",
            });
            // Добавление внутренних горихонтальных границ
            borders.Append(new InsideHorizontalBorder()
            {
                Val = new EnumValue<BorderValues>(BorderValues.Thick),
                Color = "#000000",
            });
            // Добавление внутренних вертикальных границ
            borders.Append(new InsideVerticalBorder()
            {
                Val = new EnumValue<BorderValues>(BorderValues.Thick),
                Color = "#000000",
            });
            #endregion

            properties.Append(borders);


            return properties;
        }

        /// <summary>
        /// Сохранить документ по указанном пути
        /// </summary>
        /// <param name="path">Путь, по которому нужно сохранить документ</param>
        public void SaveDocument(string path)
        {
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document, true))
            {
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document();
                mainPart.Document.Body = this.body;
                mainPart.Document.Save();
                wordDocument.Save();
            }
        }
    }
}
