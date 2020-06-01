
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JobYestWORDConsole
{

    /// <summary>
    /// Класс для работы с документами Microsoft Wordс 
    /// </summary>
    public class JobMicrosoftWord
   
   {
        #region Тестовые методы

        //    /// <summary>
        //    /// Основной объект Application, который является предком всех остальных объектов
        //    /// </summary>
        //    static Microsoft.Office.Interop.Word.Application winword = new Microsoft.Office.Interop.Word.Application();


        //    static object missing = System.Reflection.Missing.Value;
        //    /// <summary>
        //    /// Чтобы открыть существующий документ или создать новый, необходимо создать новый объект Document.
        //    /// </summary>
        //    Microsoft.Office.Interop.Word.Document document =
        //    winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);

        //    /// <summary>
        //    /// Добавление текста в документ
        //    /// </summary>
        //    public void AddTextToWord()
        //    {
        //        //Добавление текста в документ
        //        document.Content.SetRange(0, 0);
        //        document.Content.Text = "www.CSharpCoderR.com" + Environment.NewLine;

        //    }

        //        /// <summary>
        //        /// Добавление колонтитулов
        //        /// </summary>
        //        public void AddColontitum()
        //    {
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
        //            headerRange.Text = "Верхний колонтитул" + Environment.NewLine + "www.CSharpCoderR.com";
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
        //            footerRange.Text = "Нижний колонтитул" + Environment.NewLine + "www.CSharpCoderR.com";
        //        }
        //    }

        //    /// <summary>
        //    /// применить к тексту определенный стиль.
        //    /// </summary>
        //    public void MuStil()
        //    {
        //        //Добавление текста со стилем Заголовок 1
        //        Microsoft.Office.Interop.Word.Paragraph para1 = document.Content.Paragraphs.Add(ref missing);
        //        object styleHeading1 = "Заголовок 1";
        //        para1.Range.set_Style(styleHeading1);
        //        para1.Range.Text = "Исходники по языку программирования CSharp";
        //        para1.Range.InsertParagraphAfter();
        //    }


        //    public void SaveTextWord()
        //    {
        //        //Сохранение документа
        //        object filename = @"d:temp1.docx";
        //        document.SaveAs(ref filename);
        //        //Закрытие текущего документа
        //        document.Close(ref missing, ref missing, ref missing);
        //        document = null;
        //        //Закрытие приложения Word
        //        winword.Quit(ref missing, ref missing, ref missing);
        //        winword = null;
        //    }

        #endregion
    }
}

