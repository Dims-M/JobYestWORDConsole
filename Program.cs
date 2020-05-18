using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace JobYestWORDConsole
{
    class Program
    {
        static void Main(string[] args)
        {

           // Init(); // тестовое создание документа и запись в него
           // TestDoc(); //Запись в уже созданный документ

            Console.WriteLine("Для выхода нажмите любую клавишу!");
            Console.ReadKey();

        }

        static void Init()
        {/// <summary>
         /// Основной объект Application, который является предком всех остальных объектов
         /// </summary>
            Microsoft.Office.Interop.Word.Application winword = new Microsoft.Office.Interop.Word.Application();

            object missing = System.Reflection.Missing.Value;

            /// <summary>
            /// Чтобы открыть существующий документ или создать новый, необходимо создать новый объект Document.
            /// </summary>
            Microsoft.Office.Interop.Word.Document document =
            winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);

            try
            {

            /// <summary>
            /// Добавление текста в документ
            /// </summary>

            //Добавление текста в документ
            document.Content.SetRange(0, 0);
            document.Content.Text = "Реквизиты клиента." + Environment.NewLine;

            /// <summary>
            /// Добавление колонтитулов
            /// </summary>

                //Добавление верхнего колонтитула
                foreach (Microsoft.Office.Interop.Word.Section section in document.Sections)
            {
                Microsoft.Office.Interop.Word.Range headerRange =
                section.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Fields.Add(headerRange, Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage);
                headerRange.ParagraphFormat.Alignment =
                Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                headerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdBlue;
                headerRange.Font.Size = 10;
                headerRange.Text = "Сеть центров по выдаче денежных займов размещает условия предоставления денежных средств частным лицам." +
                    "Имеется калькулятор расчёта кредитования с расчетом процентов и конечной общей стоимости"
                    + Environment.NewLine + "https://kassaone.ru";
            }

            //Добавление нижнего колонтитула
            foreach (Microsoft.Office.Interop.Word.Section wordSection in document.Sections)
            {
                Microsoft.Office.Interop.Word.Range footerRange =
                wordSection.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;

                footerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdDarkRed;
                footerRange.Font.Size = 10;
                footerRange.ParagraphFormat.Alignment =
                Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                footerRange.Text = "Ваш Банк" + Environment.NewLine + "https://kassaone.ru";
            }


            /// <summary>
            /// применить к тексту определенный стиль.
            /// </summary>

            //Добавление текста со стилем Заголовок 1
            Microsoft.Office.Interop.Word.Paragraph para1 = document.Content.Paragraphs.Add(ref missing);
            object styleHeading1 = "123";
            para1.Range.set_Style(styleHeading1);
            para1.Range.Text = "Ваша заявка будет расмотрена. ";
            para1.Range.InsertParagraphAfter();


            //Сохранение документа
            //var temn = Assembly.GetExecutingAssembly().Location;
            var path = System.IO.Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName);
            string s = Environment.CurrentDirectory;
            // object filename = @"C:temp1.docx";
            
            object filename = path;
                document.SaveAs(ref filename);
               // document.SaveAs2(true);
                //Закрытие текущего документа
                document.Close(ref missing, ref missing, ref missing);
                document = null;

                //Закрытие приложения Word
                winword.Quit(ref missing, ref missing, ref missing);
                winword = null;

                Console.WriteLine($"Вроде документ сформировался!!!{ Environment.NewLine}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Произошла ошибка!!!{ Environment.NewLine}. {ex}"); 
            }

            finally
            {
                //winword.Quit(SaveChanges: false, OriginalFormat: false, RouteDocument: false);
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(winword);


                //Закрытие текущего документа
                document.Close(ref missing, ref missing, ref missing);
                document = null;
               // Закрытие приложения Word
                winword.Quit(ref missing, ref missing, ref missing);
                winword = null;

            }
        }

        /// <summary>
        /// Запись(обновление) в уже созданный документ Word
        /// </summary>
      static void TestDoc()
        {
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            try
            {

                Microsoft.Office.Interop.Word.Document doc = ap.Documents.Open(@"c:MyWord.docx", ReadOnly: false, Visible: false);
                doc.Activate();

                Microsoft.Office.Interop.Word.Selection sel = ap.Selection;

                if (sel != null)
                {
                    switch (sel.Type)
                    {
                        case Microsoft.Office.Interop.Word.WdSelectionType.wdSelectionIP:
                            sel.TypeText(DateTime.Now.ToString());
                            sel.TypeParagraph();
                            sel.TypeText("Microsoft Word");
                            sel.TypeParagraph();
                            break;

                        default:
                            Console.WriteLine("Тип выбора не обрабатывается; запись не выполняется");
                            break;

                    }

                    // Remove all meta data.
                    doc.RemoveDocumentInformation(Microsoft.Office.Interop.Word.WdRemoveDocInfoType.wdRDIAll);

                    ap.Documents.Save(NoPrompt: true, OriginalFormat: true);
                }
                else
                {
                    Console.WriteLine("Можете приобрести выбор...не пишу, чтобы сделать документ..");
                }

                ap.Documents.Close(SaveChanges: false, OriginalFormat: false, RouteDocument: false);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Исключение: " + ex.Message); // Could be that the document is already open (/) or Word is in Memory(?)
            }
            finally
            {
                ((Microsoft.Office.Interop.Word._Application)ap).Quit(SaveChanges: false, OriginalFormat: false, RouteDocument: false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ap);
            }
        }


    }
}
