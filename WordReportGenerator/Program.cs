using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordReportGenerator
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string path = "Example 1.docx";
            List<string[]> table = GetRandomTable();

            // Создание нового отчёта
            Report report = new Report();
            // Добавление параграфов
            report.AddParagraph("Это первый параграф документа");
            report.AddParagraph("Это второй параграф документа");
            // Добавление таблицы
            report.AddTable(table);
            report.SaveDocument(path);
        }

        /// <summary>
        /// Получить таблицу со случайнычи числами
        /// </summary>
        /// <param name="rowsCount">Количество строк таблицы</param>
        /// <param name="columnsCount">Количество столбцов таблицы</param>
        /// <param name="maxValue">Максимальное значение</param>
        /// <returns></returns>
        private static List<string[]> GetRandomTable(byte rowsCount = 5, byte columnsCount = 7, double maxValue = 25)
        {
            List<string[]> table = new List<string[]>();
            Random random = new Random();

            for (int i = 0; i < rowsCount; i++)
            {
                List<string> row = new List<string>();
                for (int j = 0; j < columnsCount; j++)
                {
                    // Так как NextDouble() возвращает число от 0 до 1, то умножив его на 10 и максимальное значение, можно получить число от 0 до максимального значения
                    row.Add($"{Math.Round(random.NextDouble() * 10 * maxValue, 2)}");
                }
                table.Add(row.ToArray());
            }

            return table;
        }
    }
}
