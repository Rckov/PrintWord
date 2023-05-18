using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using PrintWord.Interfaces;

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace PrintWord.Convert
{
    internal class Html2Word : IConvert
    {
        public Html2Word()
        {

        }

        public void Dispose() { }

        public void Convert(string pathFile)
        {
            var generateId = "generateId1";

            using (var wordprocessing = WordprocessingDocument.Create(Path.GetFileNameWithoutExtension(pathFile) + ".docx", WordprocessingDocumentType.Document))
            {
                var mainDocumentPart = wordprocessing.MainDocumentPart;

                if (mainDocumentPart == null)
                {
                    mainDocumentPart = wordprocessing.AddMainDocumentPart();
                    mainDocumentPart.Document = new Document();
                    mainDocumentPart.Document.Append(new Body());
                }

                using (var memoryStream = new FileStream(pathFile, FileMode.Open))
                {
                    var formatImportPart = mainDocumentPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Html, generateId);

                    formatImportPart.FeedData(memoryStream);
                    mainDocumentPart.Document.Body.Append(new AltChunk
                    {
                        Id = generateId
                    });
                }

                wordprocessing.MainDocumentPart.Document.Save();
            }
        }

        public void PasteImages(IEnumerable<string> images)
        {
            throw new NotImplementedException();
        }

        public void SaveDocument(string pathFile)
        {
            throw new NotImplementedException();
        }
    }
}