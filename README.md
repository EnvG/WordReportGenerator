# WordReportGenerator
[C#] Генератор отчётов Word (основан на OpenXML)


——————— ДОСТУПНЫЕ ФЦНКЦИИ ———————


  — Добавление текста в документ
  
  — Добавление таблиц в документ.


——————— ПОДКЛЮЧЕНИЕ ———————


Для подключения в проект скачайте файл WordReportGenerator ———> Report.cs и добавьте его в свой проект. Для работы требуется также добавить Nuget DocumentFormat.OpenXml.


——————— ПРИМЕРЫ ———————


Примеры кода (взяты из WordReportGenerator ———> Program.cs):

            // Создание нового отчёта
            Report report = new Report();
						
            // Добавление параграфов
            report.AddParagraph("Это первый параграф документа");
            report.AddParagraph("Это второй параграф документа");
						
            // Добавление таблицы
            report.AddTable(table);
						
	    // Сохранение документа по нужному пути
            report.SaveDocument(path);
