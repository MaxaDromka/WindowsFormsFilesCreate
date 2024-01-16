using System;
using DocumentFormat.OpenXml;
using System.Windows.Forms;
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

                    // Добавляем дополнительные элементы
                    Table table = body.AppendChild(new Table());
                    table.AppendChild(new TableRow());
                    TableCell cell = table.AppendChild(new TableCell());
                    cell.AppendChild(new Paragraph());
                    run = cell.AppendChild(new Run());
                    run.AppendChild(new Text("Заказчик:"));
                    run = cell.AppendChild(new Run());
                    run.AppendChild(new Text("Адрес проведения работ:"));
                    run = cell.AppendChild(new Run());
                    run.AppendChild(new Text("Место установки:"));
                    run = cell.AppendChild(new Run());
                    run.AppendChild(new Text("Заменяемый прибор:"));
                    run = cell.AppendChild(new Run());
                    run.AppendChild(new Text("- тип – "));
                    run = cell.AppendChild(new Run());
                    run.AppendChild(new Text("- заводской номер – "));
                    run = cell.AppendChild(new Run());
                    run.AppendChild(new Text("- дата выпуска – "));
                    run = cell.AppendChild(new Run());
                    run.AppendChild(new Text("- показания прибора – "));
                    run = cell.AppendChild(new Run());
                    run.AppendChild(new Text("Установленный прибор:"));
                    run = cell.AppendChild(new Run());
                    run.AppendChild(new Text("- тип – "));
                    run = cell.AppendChild(new Run());
                    run.AppendChild(new Text("- заводской номер – "));
                    run = cell.AppendChild(new Run());
                    run.AppendChild(new Text("- дата выпуска – "));
                    run = cell.AppendChild(new Run());
                    run.AppendChild(new Text("- показания прибора – "));
                    run = cell.AppendChild(new Run());
                    run.AppendChild(new Text("- дата монтажа – "));
                    run = cell.AppendChild(new Run());
                    run.AppendChild(new Text("Работы проведены в соответствии с условиями договора."));
                    run = cell.AppendChild(new Run());
                    run.AppendChild(new Text("Замену произвёл:"));
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
