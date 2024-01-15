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
            saveDialog.Filter = "PDF Documents|*.pdf";
            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                PdfDocument pdf = new PdfDocument();
                PdfPage page = pdf.AddPage();
                XGraphics gfx = XGraphics.FromPdfPage(page);
                XFont font = new XFont("Arial", 12);

                gfx.DrawString(text, font, XBrushes.Black,
                    new XRect(0, 0, page.Width.Point, page.Height.Point),
                    XStringFormats.TopLeft);

                pdf.Save(saveDialog.FileName);
                MessageBox.Show("File saved successfully.");
            }
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
