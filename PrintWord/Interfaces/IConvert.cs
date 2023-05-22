using System;
using System.Collections.Generic;

namespace PrintWord.Interfaces
{
    internal interface IConvert
    {
        void PasteHtml(string pathSave);

        void PasteImages(string fileDocument, IEnumerable<string> images);
    }
}