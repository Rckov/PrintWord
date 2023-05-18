using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using PrintWord.Interfaces;

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace PrintWord.Convert
{
    internal class Html2Word : IConvert
    {
        private readonly string _pathFile;
        private readonly WordprocessingDocument _wordprocessing;

        public Html2Word(string pathFile)
        {
            _pathFile = pathFile;
            _wordprocessing = WordprocessingDocument
                .Create(Path.GetFileNameWithoutExtension(_pathFile) + ".docx", WordprocessingDocumentType.Document);
        }

        public void Dispose()
        {
            _wordprocessing?.Dispose();
        }

        public void Convert()
        {
            var generateId = "generateId1";
            var mainDocumentPart = GetDocumentPart();

            using (var memoryStream = new FileStream(_pathFile, FileMode.Open))
            {
                var formatImportPart = mainDocumentPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Html, generateId);

                formatImportPart.FeedData(memoryStream);
                mainDocumentPart.Document.Body.InsertAfter(new AltChunk
                {
                    Id = generateId
                }, mainDocumentPart.Document.Body.Elements<Paragraph>().LastOrDefault());
                mainDocumentPart.Document.Save();
            }
        }

        public void PasteImages(IEnumerable<string> images)
        {
            foreach (var item in images)
            {
                var mainDocumentPart = GetDocumentPart();
                var imagePart = mainDocumentPart.AddImagePart(ImagePartType.Jpeg);
            
                using (FileStream stream = new FileStream(item, FileMode.Open))
                {
                    imagePart.FeedData(stream);
                }
            
                var drawElement = GetImageToBody(mainDocumentPart.GetIdOfPart(imagePart));
            
                //PasteImage(item);
            }
        }

        public void SaveDocument(string pathFile)
        {
            _wordprocessing.MainDocumentPart.Document.Save();
        }

        private MainDocumentPart GetDocumentPart(WordprocessingDocument wordprocessing = null)
        {
            var mainDocumentPart = (wordprocessing ?? _wordprocessing).MainDocumentPart;

            if (mainDocumentPart == null)
            {
                mainDocumentPart = _wordprocessing.AddMainDocumentPart();
                mainDocumentPart.Document = new Document();
                mainDocumentPart.Document.Append(new Body());
            }

            return mainDocumentPart;
        }

        private Drawing GetImageToBody(string relationshipId)
        {
            var element =
             new Drawing(
                 new DW.Inline(
                     new DW.Extent() { Cx = 990000L, Cy = 792000L },
                     new DW.EffectExtent()
                     {
                         LeftEdge = 0L,
                         TopEdge = 0L,
                         RightEdge = 0L,
                         BottomEdge = 0L
                     },
                     new DW.DocProperties()
                     {
                         Id = (UInt32Value)1U,
                         Name = "Picture 1"
                     },
                     new DW.NonVisualGraphicFrameDrawingProperties(
                         new A.GraphicFrameLocks() { NoChangeAspect = true }),
                     new A.Graphic(
                         new A.GraphicData(
                             new PIC.Picture(
                                 new PIC.NonVisualPictureProperties(
                                     new PIC.NonVisualDrawingProperties()
                                     {
                                         Id = (UInt32Value)0U,
                                         Name = "New Bitmap Image.jpg"
                                     },
                                     new PIC.NonVisualPictureDrawingProperties()),
                                 new PIC.BlipFill(
                                     new A.Blip(
                                         new A.BlipExtensionList(
                                             new A.BlipExtension()
                                             {
                                                 Uri =
                                                    "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                             })
                                     )
                                     {
                                         Embed = relationshipId,
                                         CompressionState =
                                         A.BlipCompressionValues.Print
                                     },
                                     new A.Stretch(
                                         new A.FillRectangle())),
                                 new PIC.ShapeProperties(
                                     new A.Transform2D(
                                         new A.Offset() { X = 0L, Y = 0L },
                                         new A.Extents() { Cx = 990000L, Cy = 792000L }),
                                     new A.PresetGeometry(
                                         new A.AdjustValueList()
                                     )
                                     { Preset = A.ShapeTypeValues.Rectangle }))
                         )
                         { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                 )
                 {
                     DistanceFromTop = (UInt32Value)0U,
                     DistanceFromBottom = (UInt32Value)0U,
                     DistanceFromLeft = (UInt32Value)0U,
                     DistanceFromRight = (UInt32Value)0U,
                     EditId = "50D07946"
                 });

            return element;
        }
    }
}