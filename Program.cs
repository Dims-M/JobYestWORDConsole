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

            Init();
            Console.ReadKey(true);
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

            /// <summary>
            /// Добавление текста в документ
            /// </summary>

            //Добавление текста в документ
            document.Content.SetRange(0, 0);
            document.Content.Text = "www.CSharpCoderR.com" + Environment.NewLine;

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
                headerRange.Text = "Верхний колонтитул" + Environment.NewLine + "www.CSharpCoderR.com";
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
                footerRange.Text = "Нижний колонтитул" + Environment.NewLine + "www.CSharpCoderR.com";
            }


            /// <summary>
            /// применить к тексту определенный стиль.
            /// </summary>

            //Добавление текста со стилем Заголовок 1
            Microsoft.Office.Interop.Word.Paragraph para1 = document.Content.Paragraphs.Add(ref missing);
            object styleHeading1 = "Заголовок 1";
            para1.Range.set_Style(styleHeading1);
            para1.Range.Text = "Исходники по языку программирования CSharp";
            para1.Range.InsertParagraphAfter();


            //Сохранение документа
            //var temn = Assembly.GetExecutingAssembly().Location;
            var path = System.IO.Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName);
            string s = Environment.CurrentDirectory;
            // object filename = @"C:temp1.docx";
            
            object filename = path;
                document.SaveAs(ref filename);
                //Закрытие текущего документа
                document.Close(ref missing, ref missing, ref missing);
                document = null;
                //Закрытие приложения Word
                winword.Quit(ref missing, ref missing, ref missing);
                winword = null;
            
        }

    }
}
