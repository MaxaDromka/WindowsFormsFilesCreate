using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WindowsFormsFilesCreate
{
    public partial class WordForm : Form
    {
        public WordForm()
        {
            InitializeComponent();
        }

        private void WordForm_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string text = textBox1.Text;
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.Filter = "Word Document|*.docx";
            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(saveDialog.FileName, WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                    mainPart.Document = new Document();
                    Body body = mainPart.Document.AppendChild(new Body());
                    Paragraph para = body.AppendChild(new Paragraph());
                    Run run = para.AppendChild(new Run());
                    run.AppendChild(new Text(text));
                }
                MessageBox.Show("Файл успешно создан.");
            }
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
