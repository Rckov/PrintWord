using PrintWord.Convert;
using PrintWord.Convert.Enums;
using PrintWord.Interfaces;

using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

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
                openFileDialog.Filter = "html document |*.html";
                txtPath.Text = openFileDialog.ShowDialog() == DialogResult.OK ? openFileDialog.FileName : txtPath.Text;
            }
        }

        private void BtConvert_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtPath.Text))
            {
                MessageBox.Show("Путь к HTML файле не может быть пустым");
            }
            else
            {
                var listImages = new List<string>() { "СхемаП4.jpg" };
                var type = Enum.Parse(typeof(ConvertType), combPrintType.SelectedValue.ToString());

                _convert = GetConverter((ConvertType)type);
                _convert.PasteHtml(txtPath.Text);
                _convert.PasteImages(txtPath.Text, listImages);
            }
        }

        private IConvert GetConverter(ConvertType type)
        {
            switch (type)
            {
                case ConvertType.Html2Word: return new InteropOpenXml();
                case ConvertType.InteropWord: return new InteropOfficeWord(txtPath.Text);
                default: return default;
            }
        }
    }
}