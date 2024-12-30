using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Pdf;
using Sunny.UI;


namespace Tools.Utils
{
    internal class pdf2word
    {

        public static Task ConvertWordToImagesAsync(string pdfpath) {
            if (string.IsNullOrEmpty(pdfpath))
            {
                throw new ArgumentNullException(nameof(pdfpath));
            }
            return Task.Run(() =>
           {
               Document pdfDocument = new Document(pdfpath);
               DocSaveOptions saveOptions;
               FileInfo file = new FileInfo(pdfpath);
               string wordPath = Path.GetDirectoryName(file.FullName);
               string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(file.Name);

               saveOptions = new DocSaveOptions
               {
                   Format = DocSaveOptions.DocFormat.DocX,
                   Mode = DocSaveOptions.RecognitionMode.Textbox
               };
               wordPath = Path.Combine(wordPath, $"{fileNameWithoutExtension}.docx");
               pdfDocument.Save(wordPath, saveOptions); // 保存为Word文档
           });
        }
    }
}
