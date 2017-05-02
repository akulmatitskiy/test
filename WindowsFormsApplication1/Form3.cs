using System;
using System.Collections.Generic;
using System.Xml.Serialization;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Xml.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Microsoft.Win32;

namespace WindowsFormsApplication1
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }

       
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                DataSet ds = new DataSet(); // создаем пока что пустой кэш данных
                DataTable dt = new DataTable(); // создаем пока что пустую таблицу данных
                dt.TableName = "Template"; // название таблицы
                dt.Columns.Add("Select"); // название колонок
                dt.Columns.Add("Name");
                dt.Columns.Add("Message");
                ds.Tables.Add(dt); //в ds создается таблица, с названием и колонками, созданными выше

                foreach (DataGridViewRow r in dataGridView1.Rows) // пока в dataGridView1 есть строки
                {
                    DataRow row = ds.Tables["Template"].NewRow(); // создаем новую строку в таблице, занесенной в ds
                    row["Select"] = r.Cells[0].Value;  //в столбец этой строки заносим данные из первого столбца dataGridView1
                    row["Name"] = r.Cells[1].Value; // то же самое со вторыми столбцами
                    row["Message"] = r.Cells[2].Value; //то же самое с третьими столбцами
                    ds.Tables["Template"].Rows.Add(row); //добавление всей этой строки в таблицу ds.
                }
                ds.WriteXml(@"template.xml");
                MessageBox.Show("XML файл успешно сохранен.", "Выполнено.");
            }
            catch
            {
                MessageBox.Show("Невозможно сохранить XML файл.", "Ошибка.");
            }

            // Close this window
            this.Close();
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            DataSet ds = new DataSet(); // создаем пока что пустой кэш данных
            DataTable dt = new DataTable(); // создаем пока что пустую таблицу данных
            dt.TableName = "Template"; // название таблицы
            dt.Columns.Add("Select"); // название колонок
            dt.Columns.Add("Name");
            dt.Columns.Add("Message");
            ds.Tables.Add(dt); //в ds создается таблица, с названием и колонками, созданными выше
            if (dataGridView1.Rows.Count > 0) //если в таблице больше нуля строк
            {
                MessageBox.Show("Очистите поле перед загрузкой нового файла.", "Ошибка.");
            }
            else
            {
                if (File.Exists(@"template.xml")) // если существует данный файл
                {
                    
                    ds.ReadXml(@"template.xml"); // записываем в него XML-данные из файла
            
                    foreach (DataRow item in ds.Tables["Template"].Rows)
                    {
                        int n = dataGridView1.Rows.Add(); // добавляем новую сроку в dataGridView1
                        dataGridView1.Rows[n].Cells[0].Value = item["Select"]; // заносим в первый столбец созданной строки данные из первого столбца таблицы ds.
                        dataGridView1.Rows[n].Cells[1].Value = item["Name"]; // то же самое со вторым столбцом
                        dataGridView1.Rows[n].Cells[2].Value = item["Message"]; // то же самое с третьим столбцом
                    }
                }
                else
                {
                    MessageBox.Show("XML файл не найден.", "Ошибка.");
                }
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("Заполните все поля.", "Ошибка.");
            }
            else
            {
                int n = dataGridView1.Rows.Add();
                dataGridView1.Rows[n].Cells[0].Value = comboBox1.Text;
                dataGridView1.Rows[n].Cells[1].Value = textBox1.Text;
                dataGridView1.Rows[n].Cells[2].Value = richTextBox1.Text;
            }

        }

       

        

        private void button4_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                int n = dataGridView1.SelectedRows[0].Index;
                dataGridView1.Rows[n].Cells[0].Value = comboBox1.Text;
                dataGridView1.Rows[n].Cells[1].Value = textBox1.Text; 
                dataGridView1.Rows[n].Cells[2].Value = richTextBox1.Text; 
            }
            else
            {
                MessageBox.Show("Выберите строку для редактирования.", "Ошибка.");
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                dataGridView1.Rows.RemoveAt(dataGridView1.SelectedRows[0].Index);
            }
            else
            {
                MessageBox.Show("Выберите строку для удаления.", "Ошибка.");
            }

        }
        
        private void dataGridView1_MouseClick_1(object sender, MouseEventArgs e)
        {
            comboBox1.Text = dataGridView1.SelectedRows[0].Cells[0].Value.ToString();
            textBox1.Text = dataGridView1.SelectedRows[0].Cells[1].Value.ToString();
            // int n = Convert.ToInt32(dataGridView1.SelectedRows[0].Cells[1].Value);
            richTextBox1.Text = dataGridView1.SelectedRows[0].Cells[2].Value.ToString();
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }
    }
}
