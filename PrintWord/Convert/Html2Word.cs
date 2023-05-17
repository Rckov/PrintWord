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
            var generateId = "AltChunkId1";

            using (var memoryStream = new MemoryStream())
            using (var wordprocessing = WordprocessingDocument.Create(memoryStream, WordprocessingDocumentType.Document))
            {
                var mainPart = wordprocessing.MainDocumentPart;

                if (mainPart == null)
                {
                    mainPart = wordprocessing.AddMainDocumentPart();
                    Document document = new Document(new Body());
                    document.Save(mainPart);
                }

                var formatImports = mainPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Xhtml, generateId);

                using (var streamFormat = formatImports.GetStream(FileMode.Create, FileAccess.Write))
                {
                    using (var streamWriter = new StreamWriter(streamFormat))
                    {
                        streamWriter.Write(File.ReadAllText(pathFile));
                    }
                }

                mainPart.Document.Body.InsertAt(new AltChunk()
                {
                    Id = generateId
                }, 0);
                mainPart.Document.Save();

                File.WriteAllBytes(Path.GetFileNameWithoutExtension(pathFile) + ".docx", memoryStream.ToArray());
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