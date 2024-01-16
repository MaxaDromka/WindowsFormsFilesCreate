using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;
using System.Text;
using System.Windows.Forms;

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
            string templateFilePath = "D:\\Загрузки\\Текстовый документ.txt"; // Укажите путь к вашему шаблону

            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Text Document (*.txt)|*.txt";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string saveFilePath = saveFileDialog.FileName; // Получаем выбранный пользователем путь для сохранения заполненного документа
                        string[] textBoxValues = new string[]
                        {
                        textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text, textBox5.Text,
                        textBox6.Text, textBox7.Text, textBox8.Text, textBox9.Text, textBox10.Text,
                        textBox11.Text, textBox12.Text, textBox13.Text, textBox14.Text, textBox15.Text,
                        textBox16.Text,textBox17.Text
                        };
                        string documentText = File.ReadAllText(templateFilePath);

                        documentText = documentText.Replace("____________________________", textBoxValues[0]);
                        documentText = documentText.Replace("____________________________", textBoxValues[1]);                    
                        documentText = documentText.Replace("Адрес проведения работ ________________________________________________", "Адрес проведения работ: " + textBoxValues[2] + ".");
                        documentText = documentText.Replace("Место установки ____________________________________________.", "Место установки: " + textBoxValues[3] + ".");
                        documentText = documentText.Replace("Заменяемый прибор: ", "Заменяемый прибор: " + textBoxValues[4] + ".");
                        documentText = documentText.Replace("- тип – ____________________________________________;", "- тип – " + textBoxValues[5] + ".");
                        documentText = documentText.Replace("- заводской номер – ____________________________________________;", "- заводской номер – " + textBoxValues[6] + ".");
                        documentText = documentText.Replace("- дата выпуска – ____________________________________________;", "- дата выпуска – " + textBoxValues[7] + ".");
                        documentText = documentText.Replace("- показания прибора – ____________________________________________.", "- показания прибора – " + textBoxValues[8] + ".");
                        documentText = documentText.Replace("Установленный прибор: ", "Установленный прибор: " + textBoxValues[9] + ".");
                        documentText = documentText.Replace("- тип – ____________________________________________;", "- тип – " + textBoxValues[10] + ".");
                        documentText = documentText.Replace("- заводской номер – ____________________________________________;", "- заводской номер – " + textBoxValues[11] + ".");
                        documentText = documentText.Replace("- дата выпуска – ____________________________________________;", "- дата выпуска –" + textBoxValues[12] + ".");
                        documentText = documentText.Replace("- показания прибора – ____________________________________________;", "- показания прибора –" + textBoxValues[13] + ".");
                        documentText = documentText.Replace("- дата монтажа – ____________________________________________.", "- дата монтажа –" + textBoxValues[14] + ".");
                        documentText = documentText.Replace("Замену произвёл: __________________________________________________________.", "Замену произвёл: " + textBoxValues[15] + ".");
                        documentText = documentText.Replace("_", "");

                    File.WriteAllText(saveFilePath, documentText); 

                    MessageBox.Show("Файл успешно создан.");
                }
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
