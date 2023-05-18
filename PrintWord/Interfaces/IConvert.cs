using System;
using System.Collections.Generic;

namespace PrintWord.Interfaces
{
    internal interface IConvert
    {
        void Convert();

        void PasteImages(IEnumerable<string> images);
    }
}