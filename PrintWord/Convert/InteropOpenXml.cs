using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using Microsoft.Office.Interop.Word;

using PrintWord.Interfaces;

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Xml.Linq;

using Document = Microsoft.Office.Interop.Word.Document;

namespace PrintWord.Convert
{
    internal class InteropOpenXml : IConvert
    {
        public void PasteHtml(string pathSave)
        {
            using (var wordprocessing = WordprocessingDocument.Create(pathSave + ".docx", WordprocessingDocumentType.Document))
            {
                var mainDocumentPart = wordprocessing.MainDocumentPart;

                if (mainDocumentPart == null)
                {
                    mainDocumentPart = wordprocessing.AddMainDocumentPart();
                    mainDocumentPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();
                    mainDocumentPart.Document.Append(new Body());
                }

                using (var fileStream = new FileStream(pathSave, FileMode.Open))
                {
                    var formatImportPart = mainDocumentPart
                        .AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Html, "generateId1");

                    formatImportPart.FeedData(fileStream);
                    mainDocumentPart.Document.Body.InsertAt(new AltChunk
                    {
                        Id = "generateId1"
                    }, 0);
                    mainDocumentPart.Document.Save();
                }

                wordprocessing.MainDocumentPart.Document.Save();
            }
        }

        public void PasteImages(string fileDocument, IEnumerable<string> images)
        {
            var application = new Application();
            var document = application.Documents.Open(FileName: fileDocument + ".docx", ReadOnly: false);

            try
            {
                foreach (var image in images)
                {
                    foreach (Range _rangeObject in document.StoryRanges)
                    {
                        if (_rangeObject.Find.Execute("[" + image + "]", Forward: true, Wrap: WdFindWrap.wdFindContinue))
                        {
                            _rangeObject.Select();
                            _rangeObject.Delete();
                            application.Selection.InlineShapes.AddPicture(Path.GetFullPath(images.ElementAt(0)), false, true, _rangeObject);
                        }
                    }
                }
            }
            finally
            {
                document?.Close();
                application?.Quit();

                Marshal.FinalReleaseComObject(document);
                Marshal.FinalReleaseComObject(application);
            }
        }
    }
}