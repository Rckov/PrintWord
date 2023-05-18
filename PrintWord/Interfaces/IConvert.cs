using System;
using System.Collections.Generic;

namespace PrintWord.Interfaces
{
    internal interface IConvert : IDisposable
    {
        void Convert();

        void PasteImages(IEnumerable<string> images);

        void SaveDocument(string pathFile);
    }
}