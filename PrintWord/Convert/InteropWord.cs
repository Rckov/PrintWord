using Microsoft.Office.Interop.Word;

using PrintWord.Interfaces;

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;

namespace PrintWord.Convert
{
    internal class InteropWord : IConvert
    {
        private Document _document;
        private readonly Application _application;

        public InteropWord()
        {
            _application = new Application();
        }

        public void Dispose()
        {
            _document?.Close();
            _application?.Quit();
        }

        public void Convert(string pathFile)
        {
            var pathTemp = Path.GetTempPath();
            var pathWord = Path.GetFileNameWithoutExtension(pathFile) + ".rtf";
            var pathTempWord = Path.Combine(pathTemp, pathWord);

            _document = _application.Documents.Open(FileName: pathFile, ReadOnly: false);
            _document.SaveAs(FileName: pathFile + ".rtf", FileFormat: WdSaveFormat.wdFormatRTF);

            if (!File.Exists(pathTempWord))
            {
                throw new Exception("Failed to convert html document to .rtf document");
            }
        }

        public void PasteImages(IEnumerable<string> images)
        {
            if (images.Count() == 0) return;
        }

        public void SaveDocument(string pathFile)
        {
            _document.SaveAs(FileName: pathFile + ".rtf", FileFormat: WdSaveFormat.wdFormatRTF);
        }

        private Document OpenDocument(string pathFile)
        {
            return _application.Documents.Open(FileName: pathFile, ReadOnly: false);
        }

        private void ReplaceWordParameters(string pathImage)
        {
            _application.Selection.InlineShapes.AddPicture(pathImage);
        }
    }
}