using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace PrintWord.Interfaces
{
    internal interface IConvert : IDisposable
    {
        void Convert(string pathFile);

        void PasteImages(IEnumerable<string> images);

        void SaveDocument(string pathFile);
    }
}