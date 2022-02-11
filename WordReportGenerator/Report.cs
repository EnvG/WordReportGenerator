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
        public void AddTable (List<string[]> content)
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
                    // Значение в ячейке
                    Paragraph value = this.GetParagraph(content[i][j]);

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
        #endregion

        /// <summary>
        /// Получение параграфа с данным содержанием
        /// </summary>
        /// <param name="content">Содержание параграфа</param>
        /// <returns></returns>
        private Paragraph GetParagraph (string content)
        {
            Paragraph paragraph = new Paragraph();
            Run run = new Run();
            Text text = new Text(content);

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
