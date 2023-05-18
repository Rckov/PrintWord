using Microsoft.Office.Interop.Word;

using PrintWord.Interfaces;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

namespace PrintWord.Convert
{
    internal class InteropWord : IConvert
    {
        private Document _document;
        private Application _application;

        private readonly string _pathFile;

        public InteropWord(string pathFile)
        {
            _pathFile = pathFile;
            _application = new Application();
        }

        public void Dispose()
        {
            _document?.Close();
            Marshal.FinalReleaseComObject(_document);

            _application?.Quit();
            Marshal.FinalReleaseComObject(_application);
        }

        public void Convert()
        {
            var pathTemp = Path.GetTempPath();
            var pathWord = Path.GetFileNameWithoutExtension(_pathFile) + ".rtf";
            var pathTempWord = Path.Combine(pathTemp, pathWord);

            _document = _application.Documents.Open(FileName: _pathFile, ReadOnly: false);
            _document.SaveAs(FileName: _pathFile + ".rtf", FileFormat: WdSaveFormat.wdFormatRTF);

            if (!File.Exists(pathTempWord))
            {
                throw new Exception("Failed to convert html document to .rtf document");
            }

            Dispose();
        }

        public void PasteImages(IEnumerable<string> images)
        {
            _application = new Application();
            _document = _application.Documents.Open(FileName: _pathFile + ".rtf", ReadOnly: false);

            foreach (var image in images)
            {
                foreach (Range _rangeObject in _document.StoryRanges)
                {
                    if (_rangeObject.Find.Execute("["+ image + "]", Forward: true, Wrap: WdFindWrap.wdFindContinue))
                    {
                        _rangeObject.Select();
                        _rangeObject.Delete();
                        _application.Selection.InlineShapes.AddPicture(Path.GetFullPath(images.ElementAt(0)), false, true, _rangeObject);
                    }
                }
            }
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