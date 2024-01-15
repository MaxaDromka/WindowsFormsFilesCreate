using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using iTextSharp.text;
using PdfSharp.Drawing;

namespace WindowsFormsFilesCreate
{
    public partial class PdfForm : Form
    {
        public PdfForm()
        {
            InitializeComponent();
        }

        private void PdfForm_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string text = textBox1.Text;
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.Filter = "PDF Document|*.pdf";
            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                using (FileStream fs = new FileStream(saveDialog.FileName, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    Document document = new Document();
                    PdfWriter writer = PdfWriter.GetInstance(document, fs);
                    document.Open();
                    document.Add(new Paragraph(text));
                    document.Close();
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
