using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using exportWord = Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        private readonly string TemplateFileName = @"C:\Users\usersql\Desktop\экзамен МДК0202 Карамова\WindowsFormsApp1\WindowsFormsApp1\bin\Debug\check.docx";
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int a = Convert.ToInt32(textBox1.Text);
            int b = Convert.ToInt32(textBox2.Text);
            if (radioButton1.Checked)
            {
                string al = "Алюминий";
                double c = a * b;
                c = c * 15.50 / 100;
                label4.Text = "Размера " + Convert.ToString(a) + "x" + Convert.ToString(b) + "см";
                label5.Text = "Материал " + al;
                label3.Text = "Стоимость " + Convert.ToString(c);
            }
            if (radioButton2.Checked)
            {
                string pl = "Пластик";
                double c = a * b;
                c = c * 9.90 / 100;
                label4.Text = "Размера " + Convert.ToString(a) + "x" + Convert.ToString(b) + "см";
                label5.Text = "Материал " + pl;
                label3.Text = "Стоимость " + Convert.ToString(c);
            }
            else
            {
                a = 0;
                b = 0;
                MessageBox.Show("vvedite dann");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
  
            string al = radioButton1.Text;
            string pl = radioButton2.Text;
            string c = label3.Text;
            string number = textBox3.Text;

            var wordApp = new Word.Application();
            wordApp.Visible = false;

            try
            {
                var wordDocument = wordApp.Documents.Open(TemplateFileName);
                ReplaceWordStub("{al}", al, wordDocument);
                ReplaceWordStub("{pl}", pl, wordDocument);
                ReplaceWordStub("{c}", c, wordDocument);
                ReplaceWordStub("{number}", number, wordDocument);

                wordDocument.SaveAs(@"C:\Users\usersql\Desktop\экзамен МДК0202 Карамова\WindowsFormsApp1\WindowsFormsApp1\bin\Debug\check3.docx");
            }
            catch
            {
                MessageBox.Show("Произошла ошибка");
            }
        }

        private void ReplaceWordStub(string v, string al, Word.Document wordDocument)
        {
            throw new NotImplementedException();
        }
    }
}
