using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using Microsoft.Office.Interop.Word;

using PrintWord.Interfaces;

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

using Document = Microsoft.Office.Interop.Word.Document;

namespace PrintWord.Convert
{
    internal class Html2Word : IConvert
    {
        private readonly string _pathFile;

        private Document _document;
        private readonly Application _application;

        public Html2Word(string pathFile)
        {
            _pathFile = pathFile;
            _application = new Application();
        }

        public void Convert()
        {
            using (var wordprocessing = WordprocessingDocument.Create(Path.GetFileNameWithoutExtension(_pathFile) + ".docx", WordprocessingDocumentType.Document))
            {
                var mainDocumentPart = GetDocumentPart(wordprocessing);

                using (var memoryStream = new FileStream(_pathFile, FileMode.Open))
                {
                    var formatImportPart = mainDocumentPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Html, "generateId1");

                    formatImportPart.FeedData(memoryStream);
                    mainDocumentPart.Document.Body.Append(new AltChunk
                    {
                        Id = "generateId1"
                    });
                }

                wordprocessing.MainDocumentPart.Document.Save();
            }
        }

        public void PasteImages(IEnumerable<string> images)
        {
            try
            {
                var application = new Application();
                var document = application.Documents.Open(FileName: Path.GetFullPath(_pathFile + ".docx"), ReadOnly: false);

                foreach (var image in images)
                {
                    foreach (Range rangeObject in document.StoryRanges)
                    {
                        if (rangeObject.Find.Execute("[" + image + "]", Forward: true, Wrap: WdFindWrap.wdFindContinue))
                        {
                            rangeObject.Select();
                            rangeObject.Delete();
                            application.Selection.InlineShapes.AddPicture(Path.GetFullPath(images.ElementAt(0)), false, true, rangeObject);
                        }
                    }
                }
            }
            finally
            {
                _document?.Close();
                _application?.Quit();
            }
        }

        private MainDocumentPart GetDocumentPart(WordprocessingDocument wordprocessing)
        {
            var mainDocumentPart = wordprocessing.MainDocumentPart;

            if (mainDocumentPart == null)
            {
                mainDocumentPart = wordprocessing.AddMainDocumentPart();
                mainDocumentPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();
                mainDocumentPart.Document.Append(new Body());
            }

            return mainDocumentPart;
        }
    }
}