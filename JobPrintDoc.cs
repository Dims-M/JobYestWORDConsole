using GemBox.Document;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JobYestWORDConsole
{

   public static class JobPrintDoc
    {

        private static DocumentModel document;

        /// <summary>
        /// Печатает напримую в принтер. НЕ СЕРВЕРНЫЙ ВАРИАНТ
        /// </summary>
        static public void TestPrint()
        {
            // лицензия
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            try
            {

                // обьек для работы с документом
                DocumentModel document = DocumentModel.Load("Doc1.docx");

                // Установите параметры страницы документа Word.
                foreach (Section section in document.Sections)
                {
                    PageSetup pageSetup = section.PageSetup;
                    pageSetup.Orientation = Orientation.Landscape;
                    pageSetup.LineNumberRestartSetting = LineNumberRestartSetting.NewPage;
                    pageSetup.LineNumberDistanceFromText = 50;

                    PageMargins pageMargins = pageSetup.PageMargins;
                    pageMargins.Top = 20;
                    pageMargins.Left = 100;
                }

                // Печать документа Word на принтере по умолчанию (например, "Microsoft Print to Pdf").
                string printerName = null;
                document.Print(printerName);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка при работе" + ex);
            }


        }
   
        //не работает.
    static public void ConverdPdf(String AsBase64String )
        {
            string pdflocation = "TestAsBase64StringPDF";
            string fileName = "softcore.pdf";

            // put your base64 string converted from online platform here instead of V
            var base64String = AsBase64String;

            int mod4 = base64String.Length % 4;

            // as of my research this mod4 will be greater than 0 if the base 64 string is corrupted
            if (mod4 > 0)
            {
                base64String += new string('=', 4 - mod4);
            }
            pdflocation = pdflocation + fileName;

            byte[] data = Convert.FromBase64String(base64String);

            using (FileStream stream = System.IO.File.Create(pdflocation))
            {
                stream.Write(data, 0, data.Length);
            }

        }

        /// <summary>
        /// Конвертириуем пдв в строку Base64String и сохраняем в файл
        /// </summary>
        static public void ConverdToBase64String()
        {
            string textFromFile;

            var bytes = File.ReadAllBytes("rrrrtest.pdf");
            var base64 = Convert.ToBase64String(bytes);

            using (FileStream fstream = File.OpenRead($"222.pdf"))
            {
                // преобразуем строку в байты
                byte[] array = new byte[fstream.Length];
                // считываем данные
                fstream.Read(array, 0, array.Length);
                // декодируем байты в строку
                textFromFile = System.Text.Encoding.Default.GetString(array);
                // Console.WriteLine($"Текст из файла: {textFromFile}");
            }

           

            //using (StreamWriter writer = File.AppendText(path))
            //{
            //    writer.WriteLine(EncryptText(contents));
            //}

            string sorstext = textFromFile;
            string b64 = Convert.ToBase64String(Encoding.Default.GetBytes(sorstext));
            Console.WriteLine(b64);



            using (FileStream fstream = new FileStream(@"note.txt", FileMode.OpenOrCreate))
            {
                // преобразуем строку в байты
                byte[] array = System.Text.Encoding.Default.GetBytes(b64);
                // запись массива байтов в файл
                fstream.Write(array, 0, array.Length);
                //Console.WriteLine("Текст записан в файл");
            }
        }

    }

}
