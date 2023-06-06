using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TeorVer
{
    public partial class Form1 : Form
    {
        int amount = 0;
        bool changedPath = false;
        public Form1()
        {
            InitializeComponent();
        }

        private void Button_Start_Click(object sender, EventArgs e)
        {
            bool flag = true;
            try
            {
                amount = Convert.ToInt32(textBox1.Text);
                if (amount <= 0) {
                    throw new ArgumentException();
                }
                if (label2.Text == "Не указан")
                {
                    throw new NullReferenceException();
                }
            }
            catch (Exception ex) when (ex is ArgumentException || ex is FormatException)
            {
                MessageBox.Show("Укажите корректное число вариантов","Ошибка",MessageBoxButtons.OK);
                flag = false;
            }
            catch (Exception ex) when (ex is NullReferenceException)
            {
                MessageBox.Show("Укажите путь генерации", "Ошибка", MessageBoxButtons.OK);
                flag = false;
            }
            finally
            {
                if (flag)
                {
                    if (changedPath) Program.Generate(amount, folderBrowserDialog1.SelectedPath,openFileDialog1.FileName);
                    else Program.Generate(amount, folderBrowserDialog1.SelectedPath);
                }
            }
        }

        private void browser_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                label2.Text = folderBrowserDialog1.SelectedPath; 
            }
        }

        private void EditButton_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "json files|*.json";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                label5.Text = openFileDialog1.FileName;
                changedPath = true;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Воробьев А.Д. - Разработка архитектуры приложения, \nпрограммирование. \n\nКалайдина Г.В. - Постановка задачи. Разработка математических \nмоделей. Тестирование программы. \n\nШемякин Н.П. - Формирование вариантов задач, \nнаписание  исходного кода программы.", "О создателях", MessageBoxButtons.OK);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            changedPath = false;
            label5.Text = "Стандартный";
        }
    }
}
