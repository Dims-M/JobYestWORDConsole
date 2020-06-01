
using GemBox.Document;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Paragraph = GemBox.Document.Paragraph;
using Run = GemBox.Document.Run;

namespace JobYestWORDConsole
{
    class Program
    {
        static void Main(string[] args)
        {

            // Init(); // тестовое создание документа и запись в него
            //TestDoc(); //Запись в уже созданный документ

            // TestOpenXML(); //библиотека OpenXML

            // GemBoxDocument();
            //  GemBoxDocument2();
            // GemBoxDocument0();
            //GemBoxDocumentTest();
            // GetBoxDoc("Writing.docx");
            // GetBoxDocPDF("Writing.docx");
            // GetBoxCreateWord();
            // GetBoxCreateWord2(); // Метод работает!!!!!

            //Печать на принтере нарямую из кода
            // JobPrintDoc.TestPrint();

            // JobPrintDoc.ConverdToBase64String();

            string pathFolder = @"C:\\1\\";
            string videoUrl = "https://www.youtube.com/watch?v=lzm5llVmR2E";
            string nameFile = "test";

            RabotaYoutube.SaveMP3(pathFolder, videoUrl, nameFile);


            Console.WriteLine("Для выхода нажмите любую клавишу!");
            Console.ReadKey();

        }

        /// <summary>
        /// Тестовой метод 2
        /// </summary>
         static void GetBoxCreateWord2()
        {
            string[] temZnach = new string[] { "Gthdjt", "Массив представляет", "мы можем" };
            string temp = "";
            // создание ссылок дирикторий(папок для картинки). Корневая папка
            var originalDirectory = new DirectoryInfo(string.Format(@"~Archive_Documents\\Uploads"));

            //Бесплатное Лицензия
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            // Создание нового дока.
            var document = new DocumentModel();
            try
            {
                for (int i = 0; i < temZnach.Length; i++)
                {
                    temp += (temZnach[i] + Environment.NewLine);
                    document.Content.LoadText(temp);

                };

                document.Save("Archive_Documents/TestSaveDoc.docx");
               // document.Save("`Archive_Documents/TestSaveDoc.pdf");

            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка при работе" + ex);
            }
        }


        /// <summary>
        /// Тестовой вариант
        /// </summary>
        static void GetBoxCreateWord()
        {
            string[] temZnach = new string[] { "Gthdjt", "Массив представляет", "мы можем" };
            string temp = "";
            //Лицензия
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            // Создание нового дока.
            var document = new DocumentModel();
            try
            {
                for (int i = 0; i < temZnach.Length; i++)
                {
                   // temp += (temZnach[i] + Environment.NewLine);
                    temp += (temZnach[i] + "\t\n");
                   // temp += (temZnach[i]);
                    document.Content.LoadText(temp);

                };

                //document.Sections.Add(
                //    new Section(document,
                //      new Paragraph(document,
                //          new Run(document, temp),
                //          new SpecialCharacter(document, SpecialCharacterType.LineBreak)
                //      // new Run(document, ""/*"***"*//*"\xFC" + "\xF0" + "\x32"*/) { CharacterFormat = { FontName = "Wingdings", Size = 48 } }
                //      )));
               
           

                //// Добавление нового раздела с двумя абзацами, содержащими некоторый текст и символы
                //document.Sections.Add(
                //new Section(document,
                //    new Paragraph(document,
                //        new Run(document, "Это наш первый абзац с символами, добавленными в новую строку."),
                //        new SpecialCharacter(document, SpecialCharacterType.LineBreak),
                //        new Run(document, ""/*"***"*//*"\xFC" + "\xF0" + "\x32"*/) { CharacterFormat = { FontName = "Wingdings", Size = 48 } }),
                //    new Paragraph(document, "Это наш второй абзац."))
                //);

            // Save Word document to file's path.
            document.Save("TestSaveDoc.docx");
            document.Save("TestSaveDoc.pdf");

            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка при работе" + ex);
            }
        }


        static void GetBoxDocPDF(string ebillpath)
        {
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            try
            {
                DocumentModel doc = new DocumentModel();

            CharacterFormat charFormat = doc.DefaultCharacterFormat;
            charFormat.Size = 10;
            charFormat.FontName = "Courier New";

            ParagraphFormat parFormat = doc.DefaultParagraphFormat;
            parFormat.SpaceAfter = 0;
            parFormat.LineSpacing = 1;

            // It seems you want to skip first line with 'clearedtop'.
            // So maybe you could just use this instead.
            string text = string.Concat(File.ReadLines(ebillpath).Skip(1));
            doc.Content.LoadText(text);

            Section section = doc.Sections[0];
            PageMargins margins = section.PageSetup.PageMargins;
            margins.Bottom = 36;
            margins.Top = 36;
            margins.Right = 36;
            margins.Left = 36;

            doc.Save(@"test.pdf");

            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка при работе" + ex);
            }
        }
        static void GetBoxDoc(string ebillpath)
        {
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");
            try
            {

            using (var sr = new StreamReader (ebillpath))
            {
                var doc = new DocumentModel();
                doc.DefaultCharacterFormat.Size = 10;
                doc.DefaultCharacterFormat.FontName = "Courier New";

                var section = new Section(doc);
                doc.Sections.Add(section);
                string line;

                var clearedtop = false;

                while ((line = sr.ReadLine()) != null)
                {
                    if (string.IsNullOrEmpty(line) && !clearedtop)
                    {
                        continue;
                    }

                    clearedtop = true;
                    Paragraph paragraph2 = new Paragraph(doc, new Run(doc, line));
                    section.Blocks.Add(paragraph2);
                }

                PageSetup pageSetup = new PageSetup(); // section.PageSetup;
                var pm = new PageMargins();
                pm.Bottom = 36;
                pm.Top = 36;
                pm.Right = 36;
                pm.Left = 36;
                pageSetup.PageMargins = pm;

                doc.Save(@"test.pdf");
            }

            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка при работе"+ ex);
            }
        }

        static void GemBoxDocumentTest()
        {
            string[] temZnach = new string[] { "Gthdjt", "Массив представляет", "мы можем" };

            //Лицензии
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            // обьект для работы с вордом
            DocumentModel document = new DocumentModel();

            //настройки шрифта
            document.DefaultCharacterFormat.Size = 25;

            //производные элементы, имеющие определенный набор свойств,
            //используемых для определения страниц
            Section section = new Section(document);

            //Добавляем текущий обьект в документ
            document.Sections.Add(section);

            //Представляет собой абзац содержания в документе.
            Paragraph paragraph = new Paragraph(document);

            //Документ.Секции блоков.
            section.Blocks.Add(paragraph);

            for (int i = 0; i < temZnach.Length; i++)
            {
                Run run = new Run(document, temZnach[i]);
                // document.Content.LoadText(temZnach[i], new CharacterFormat() { FontName = "Arial" });
                paragraph.Inlines.Add(run);
            }

            // Представляет область текста с общим набором свойств.
           // Run run = new Run(document, "Programming language: C++, C# and Java");

            //Абзац в строчках.
          //  paragraph.Inlines.Add(run);

            document.Save("Doc1.docx");
           
        }

        static void GemBoxDocument0()
        {
            string[] temZnach = new string[] { "Gthdjt", "Массив представляет", "мы можем" };

            try
            {

           
            // Бесплатная версия.
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            // Создаем новый докумен
            var document = new DocumentModel();

            for (int i = 0; i< temZnach.Length; i++ )
            {
                document.Content.LoadText(temZnach[i], new CharacterFormat() { FontName = "Arial" });
            }

            // Добавьте в документ обычный текст
            //document.Content.LoadText("Внутри все очень просто, но удобно....", new CharacterFormat() { FontName = "Arial" });

            // Вставьте текст в формате RTF в начале документа.
            var position = document.Content.Start.LoadText(@"{\rtf1\ansi\deff0{\fonttbl{\f0 Arial Black;}}{\colortbl ;\red255\green128\blue64;}\f0\cf1 Это богатый форматированный текст 1.}",
                LoadOptions.RtfDefault);

            // Вставьте текст в формате HTML после предыдущего текста.
            position.LoadText("<p style='font-family:Arial Narrow;color:royalblue;'>Это еще один богатый форматированный текст 2 .</p>",
                LoadOptions.HtmlDefault);

            // Сохраните документ Word в путь к файлу.
            document.Save("Writing.docx");

            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка при работе \t\n" + ex);
            }
        }
        static void GemBoxDocument1()
        {

            string[] temZnach = new string[] { "Gthdjt", "Массив представляет", "мы можем" };
            // If using Professional version, put your serial key below.
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            //Создаем  новый пустой документ 
            DocumentModel document = new DocumentModel();

            //секция документа.
            Section section = new Section(document);

            // Запысываем пустую секцию в документ
            document.Sections.Add(section);

            //Представляет собой абзац содержания в документе.
            Paragraph paragraph = new Paragraph(document);
            Paragraph paragraph2 = new Paragraph(document);

            // добавляем блок в секцию 
            section.Blocks.Add(paragraph);
            section.Blocks.Add(paragraph2);


            // Инициализирует новый экземпляр класса Run с указанным текстом.
            Run run = new Run(document, "Hello World!");

            Run run2 = new Run(document, temZnach.ToString());

            paragraph.Inlines.Add(run);
            paragraph.Inlines.Add(run2);

            document.Save("TestGemBox.docx");
        }
        static void GemBoxDocument3()
        {
            DocumentModel doc = null;
           
            try
            {
                string CreateDoc = @"c:\1\TestGemBox.docx";

                doc = new DocumentModel();
                doc.Save(CreateDoc);

                // Путь до файла
                string destFileName2 = @"c:\1\TestGemBox2.docx";

                //// Create a new empty Word file.
                //var doc = new DocumentModel();

                // Обязательная строка, указываем, что мы используем лимитированную версию
                ComponentInfo.SetLicense("FREE-LIMITED-KEY");

                // Загружаем в память наш документ
                doc = DocumentModel.Load(CreateDoc);

                string[] data = new string[] { "Alex", "Gulynin", "27" };

               //Коллекция закладок
                BookmarkCollection wBookmarks = doc.Bookmarks;

               // ContentRange - это область содержимого в документе
                ContentRange wRange;

                int i = 0;

               // Пробегаем по всем закладкам в документе
                foreach (Bookmark mark in doc.Bookmarks)
                {
                 //   Получаем содержимое закладки
                   wRange = mark.GetContent(false);
                 //   Загружаем туда нужный текст
                    wRange.LoadText(data[i].ToString());
                    i++;
                }

                // Сохраняем изменения в нашем документе
                doc.Save(destFileName2);
                doc = null;
            }
            catch (Exception ex)
            {
                doc = null;
                Console.WriteLine("Во время выполнения программы произошла ошибка! Текст ошибки: {0}", ex.Message);
                Console.ReadLine();
            }
        }

        static void GemBoxDocument2()
        {
            // Set license key to use GemBox.Document in Free mode.
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            // Create a new empty Word file.
            var doc = new DocumentModel();

            // Add a new document content.
            doc.Sections.Add(new Section(doc, new Paragraph(doc, "Hello world!")));

            // Save to DOCX and PDF files.
            doc.Save("Document.docx");
            //doc.Save("Document.pdf");
        }

        #region Тестовые методы
        //static void Init()
        //{/// <summary>
        // /// Основной объект Application, который является предком всех остальных объектов
        // /// </summary>
        //    Microsoft.Office.Interop.Word.Application winword = new Microsoft.Office.Interop.Word.Application();

        //    object missing = System.Reflection.Missing.Value;

        //    /// <summary>
        //    /// Чтобы открыть существующий документ или создать новый, необходимо создать новый объект Document.
        //    /// </summary>
        //    Microsoft.Office.Interop.Word.Document document =
        //    winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);

        //    try
        //    {

        //        /// <summary>
        //        /// Добавление текста в документ
        //        /// </summary>

        //        //Добавление текста в документ
        //        document.Content.SetRange(0, 0);
        //        document.Content.Text = "Реквизиты клиента." + Environment.NewLine;

        //        /// <summary>
        //        /// Добавление колонтитулов
        //        /// </summary>

        //        //Добавление верхнего колонтитула
        //        foreach (Microsoft.Office.Interop.Word.Section section in document.Sections)
        //        {
        //            Microsoft.Office.Interop.Word.Range headerRange =
        //            section.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
        //            headerRange.Fields.Add(headerRange, Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage);
        //            headerRange.ParagraphFormat.Alignment =
        //            Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
        //            headerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdBlue;
        //            headerRange.Font.Size = 10;
        //            headerRange.Text = "Сеть центров по выдаче денежных займов размещает условия предоставления денежных средств частным лицам." +
        //                "Имеется калькулятор расчёта кредитования с расчетом процентов и конечной общей стоимости"
        //                + Environment.NewLine + "https://kassaone.ru";
        //        }

        //        //Добавление нижнего колонтитула
        //        foreach (Microsoft.Office.Interop.Word.Section wordSection in document.Sections)
        //        {
        //            Microsoft.Office.Interop.Word.Range footerRange =
        //            wordSection.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;

        //            footerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdDarkRed;
        //            footerRange.Font.Size = 10;
        //            footerRange.ParagraphFormat.Alignment =
        //            Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
        //            footerRange.Text = "Ваш Банк" + Environment.NewLine + "https://kassaone.ru";
        //        }


        //        /// <summary>
        //        /// применить к тексту определенный стиль.
        //        /// </summary>

        //        //Добавление текста со стилем Заголовок 1
        //        Microsoft.Office.Interop.Word.Paragraph para1 = document.Content.Paragraphs.Add(ref missing);
        //        object styleHeading1 = "123";
        //        para1.Range.set_Style(styleHeading1);
        //        para1.Range.Text = "Ваша заявка будет расмотрена. ";
        //        para1.Range.InsertParagraphAfter();


        //        //Сохранение документа
        //        //var temn = Assembly.GetExecutingAssembly().Location;
        //        var path = System.IO.Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName);
        //        string s = Environment.CurrentDirectory;
        //        // object filename = @"C:temp1.docx";

        //        object filename = path;
        //        document.SaveAs(ref filename);

        //        // document.SaveAs2(true);
        //        //Закрытие текущего документа
        //        document.Close(ref missing, ref missing, ref missing);
        //        document = null;

        //        //Закрытие приложения Word
        //        winword.Quit(ref missing, ref missing, ref missing);
        //        winword = null;
        //        winword.Quit();

        //        Console.WriteLine($"Вроде документ сформировался!!!{ Environment.NewLine}");
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine($"Произошла ошибка!!!{ Environment.NewLine}. {ex}");
        //    }

        //    finally
        //    {
        //        //winword.Quit(SaveChanges: false, OriginalFormat: false, RouteDocument: false);
        //        //System.Runtime.InteropServices.Marshal.ReleaseComObject(winword);


        //        //Закрытие текущего документа
        //        document.Close(ref missing, ref missing, ref missing);
        //        document = null;
        //        // Закрытие приложения Word
        //        winword.Quit(ref missing, ref missing, ref missing);
        //        winword = null;

        //    }
        //}

        ///// <summary>
        ///// Запись(обновление) в уже созданный документ Word с помощью Microsoft.Office.Interop.Word
        ///// </summary>
        //static void TestDoc()
        //{
        //    Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
        //    try
        //    {

        //        Microsoft.Office.Interop.Word.Document doc = ap.Documents.Open(@"c:MyWord.docx", ReadOnly: false, Visible: false);
        //        doc.Activate();

        //        Microsoft.Office.Interop.Word.Selection sel = ap.Selection;

        //        if (sel != null)
        //        {
        //            switch (sel.Type)
        //            {
        //                case Microsoft.Office.Interop.Word.WdSelectionType.wdSelectionIP:
        //                    sel.TypeText(DateTime.Now.ToString());
        //                    sel.TypeParagraph();
        //                    sel.TypeText("Microsoft Word");
        //                    sel.TypeParagraph();
        //                    break;

        //                default:
        //                    Console.WriteLine("Тип выбора не обрабатывается; запись не выполняется");
        //                    break;

        //            }

        //            // Remove all meta data.
        //            doc.RemoveDocumentInformation(Microsoft.Office.Interop.Word.WdRemoveDocInfoType.wdRDIAll);

        //            ap.Documents.Save(NoPrompt: true, OriginalFormat: true);
        //        }
        //        else
        //        {
        //            Console.WriteLine("Можете приобрести выбор...не пишу, чтобы сделать документ..");
        //        }

        //        ap.Documents.Close(SaveChanges: false, OriginalFormat: false, RouteDocument: false);
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine("Исключение: " + ex.Message); // Could be that the document is already open (/) or Word is in Memory(?)
        //    }
        //    finally
        //    {
        //        ((Microsoft.Office.Interop.Word._Application)ap).Quit(SaveChanges: false, OriginalFormat: false, RouteDocument: false);
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(ap);
        //    }
        //}


        //static void TestOpenXML()
        //{
        //    try
        //    {
        //        // Получаем массив байтов из нашего файла
        //        byte[] textByteArray = File.ReadAllBytes(@"C:\MyWord.docx");
        //        // Массив данных
        //        string[] data = new string[3] { "27", "Гулынин", "Алексей" };
        //        // Начинаем работу с потоком
        //        using (MemoryStream stream = new MemoryStream())
        //        {
        //            // Записываем в поток наш word-файл
        //            stream.Write(textByteArray, 0, textByteArray.Length);

        //            // Открываем документ из потока с возможностью редактирования
        //            using (WordprocessingDocument doc = WordprocessingDocument.Open(stream, true))
        //            {
        //                // Ищем все закладки в документе
        //                var bookMarks = FindBookmarks(doc.MainDocumentPart.Document);

        //                int i = 0;
        //                foreach (var end in bookMarks)
        //                {
        //                    // В документе встречаются какие-то служебные закладки
        //                    // Таким способом отфильтровываем всё ненужное
        //                    // end.Key содержит имена наших закладок
        //                    if (end.Key != "Name" && end.Key != "Age" && end.Key != "Surname") continue;

        //                    // Создаём текстовый элемент
        //                    var textElement = new Text(data[i].ToString());

        //                    // Далее данный текст добавляем в закладку
        //                    var runElement = new Run(textElement);

        //                    end.Value.InsertAfterSelf(runElement);
        //                    i++;
        //                }
        //            }
        //            // Записываем всё в наш файл
        //            File.WriteAllBytes(@"D:\Test.docx", stream.ToArray());
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine("Во время выполнения программы произошла ошибка! Текст ошибки: {0}", ex.Message);
        //        Console.ReadLine();
        //    }
        //}

        // Получаем все закладки в документе
        // bStartWithNoEnds - словарь, который будет содержать только начало закладок,
        // чтобы потом по ним находить соответствующие им концы закладок
        // documentPart - элемент Word-документа
        // outs - конечный результат
        //private static Dictionary<string, BookmarkEnd> FindBookmarks(OpenXmlElement documentPart, Dictionary<string, BookmarkEnd> outs = null, Dictionary<string, string> bStartWithNoEnds = null)
        //{
        //    if (outs == null) { outs = new Dictionary<string, BookmarkEnd>(); }
        //    if (bStartWithNoEnds == null) { bStartWithNoEnds = new Dictionary<string, string>(); }

        //    // Проходимся по всем элементам на странице Word-документа
        //    foreach (var docElement in documentPart.Elements())
        //    {
        //        // BookmarkStart определяет начало закладки в рамках документа
        //        // маркер начала связан с маркером конца закладки
        //        if (docElement is BookmarkStart)
        //        {
        //            var bookmarkStart = docElement as BookmarkStart;
        //            // Записываем id и имя закладки
        //            bStartWithNoEnds.Add(bookmarkStart.Id, bookmarkStart.Name);
        //        }

        //        // BookmarkEnd определяет конец закладки в рамках документа
        //        if (docElement is BookmarkEnd)
        //        {
        //            var bookmarkEnd = docElement as BookmarkEnd;
        //            foreach (var startName in bStartWithNoEnds)
        //            {
        //                // startName.Key как раз и содержит id закладки
        //                // здесь проверяем, что есть связь между началом и концом закладки
        //                if (bookmarkEnd.Id == startName.Key)
        //                    // В конечный массив добавляем то, что нам и нужно получить
        //                    outs.Add(startName.Value, bookmarkEnd);
        //            }
        //        }
        //        // Рекурсивно вызываем данный метод, чтобы пройтись по всем элементам
        //        // word-документа
        //        FindBookmarks(docElement, outs, bStartWithNoEnds);
        //    }

        //    return outs;
        //}



        //static void Init33()
        //{
        //    string filePath = @"C:\MyWord1234.docx";
        //    DocX doc = DocX.Create(filePath);
        //    Paragraph p1 = template.InsertParagraph();
        //    p1.AppendLine("This line contains a ").Append("bold").Bold().Append(" word.");
        //    p1.AppendLine("Here is example with question mark?");
        //    p1.AppendLine();
        //    p1.AppendLine("Can you help me figure it out?");
        //    p1.AppendLine();

        #endregion

    }
}
