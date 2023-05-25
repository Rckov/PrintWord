using DocumentFormat.OpenXml.Wordprocessing;

using Microsoft.Office.Interop.Word;

using PrintWord.Convert;
using PrintWord.Convert.Enums;
using PrintWord.Interfaces;

using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

using Application = Microsoft.Office.Interop.Word.Application;
using Document = Microsoft.Office.Interop.Word.Document;

namespace PrintWord
{
    public partial class MainForm : Form
    {
        private IConvert _convert;

        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            combPrintType.DataSource = Enum.GetValues(typeof(ConvertType));
        }

        private void BtBrowse_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                //openFileDialog.Filter = "html document |*.html";
                txtPath.Text = openFileDialog.ShowDialog() == DialogResult.OK ? openFileDialog.FileName : txtPath.Text;
            }
        }

        private void BtConvert_Click(object sender, EventArgs e)
        {
            if (false)
            {
                var unused = MessageBox.Show("Путь к HTML файле не может быть пустым");
            }
            else
            {
                Application _application = new Application();
                Document _document = _application.Documents.Open(FileName: txtPath.Text, ReadOnly: false);

                var path = "C:\\Users\\Kuraz\\Desktop\\60_10_0010205_25_2020-11-06_evz07.xml.pdf";

                object missing = System.Reflection.Missing.Value;

                object oFalse = false;

                foreach (Range _rangeObject in _document.StoryRanges)
                {
                    if (_rangeObject.Find.Execute("[" + "PDF" + "]", Forward: true, Wrap: WdFindWrap.wdFindContinue))
                    {
                        _rangeObject.Select();
                        _rangeObject.Delete();
                        _application.Selection.InlineShapes.AddOLEObject(ref missing, path, ref missing, ref missing, ref missing, ref missing, ref missing, _rangeObject);
                    }
                }

                _document.Close();
                _application.Quit();
                //var listImages = new List<string>() { "СхемаП4.jpg" };
                //var type = Enum.Parse(typeof(ConvertType), combPrintType.SelectedValue.ToString());
                //
                //_convert = GetConverter((ConvertType)type);
                //_convert.PasteHtml(txtPath.Text);
                //_convert.PasteImages(txtPath.Text, listImages);
            }
        }

        private IConvert GetConverter(ConvertType type)
        {
            switch (type)
            {
                case ConvertType.Html2Word: return new InteropOpenXml();
                default: return default;
            }
        }
    }
}