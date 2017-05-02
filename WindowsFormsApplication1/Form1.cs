using System;
using System.Collections.Generic;
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
using System.Reflection;
using ExcelObj = Microsoft.Office.Interop.Excel;




namespace WindowsFormsApplication1
{
     public partial class Form1 : Form
    {
        public static string cfgxmldir;
        public static string smsgate;
        public static string smsnum;
        public static string smsmessage;
        
        public Form1()
                    
        {
            InitializeComponent();
            cfgxmldir = (@"c:\xml");
            smsgate = ("http://192.168.12.80/goip/en/dosend.php?USERNAME=root&PASSWORD=root&smsprovider=4&smsnum="+smsnum+"&method=2&Memo="+smsmessage);


            RegistryKey op = Registry.CurrentUser.OpenSubKey("Rostretail");
            if (op != null)
            {
                cfgxmldir = (string)op.GetValue("xmldir");
                smsgate = (string)op.GetValue("smsgate");
            }
            else
            {
                RegistryKey key = Registry.CurrentUser.CreateSubKey("Rostretail");
                key.SetValue("xmldir", cfgxmldir);
                key.SetValue("smsgate", smsgate);
                key.Close();
            }
            
        }
     
        private DataTable CreateTable()
        {
            //создаём таблицу
            DataTable dt = new DataTable("person");
            //создаём три колонки
            DataColumn colCHECK = new DataColumn("Active", typeof(bool));
            DataColumn colPhone = new DataColumn("Phone", typeof(String));
            DataColumn colFName = new DataColumn("Firstname", typeof(String));
            DataColumn colLName = new DataColumn("Lastname", typeof(String));
            DataColumn colAddr = new DataColumn("Address", typeof(String));
            //добавляем колонки в таблицу
            dt.Columns.Add(colCHECK);
            dt.Columns.Add(colPhone);
            dt.Columns.Add(colFName);
            dt.Columns.Add(colLName);
            dt.Columns.Add(colAddr);
            return dt;
        }

        private static string GET(string pUSER, string pPASS, string pSMSNUM, string pSERVER, string pMESSAGE)
        {
            string sum = "";
            try
            {
                var Dataget = "http://" + pSERVER + "/goip/en/dosend.php?USERNAME=" + pUSER + "&PASSWORD=" + pPASS + "&smsprovider=4&smsnum=" + pSMSNUM + "&method=2&Memo=" + pMESSAGE;
                System.Net.WebRequest req = System.Net.WebRequest.Create(Dataget);
                System.Net.WebResponse resp = req.GetResponse();
                System.IO.Stream stream = resp.GetResponseStream();
                System.IO.StreamReader sr = new System.IO.StreamReader(stream);
                string Out = sr.ReadToEnd();

                sr.Close();
                var smspos1 = Out.IndexOf("messageid=");
                var smspos2 = Out.IndexOf("&USER");
                var numleng = smspos2 - smspos1;
                var smsid = Out.Substring((smspos1 + 10), (smspos2 - smspos1 - 10));

                var Dataansver = ("http://192.168.12.80/goip/en/resend.php?messageid=" + smsid + "&USERNAME=root&PASSWORD=root");

                System.Net.WebRequest ansv = System.Net.WebRequest.Create(Dataansver);
                System.Net.WebResponse respancv = ansv.GetResponse();
                System.IO.Stream streamav = respancv.GetResponseStream();
                System.IO.StreamReader sr1 = new System.IO.StreamReader(streamav);
                string Out1 = sr1.ReadToEnd();
                sr1.Close();
                return Out1;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return sum;
        }
        private DataTable ReadXml()
        {
            DataTable dt = null;
            try
            {
                string[] files = new DirectoryInfo(@cfgxmldir).GetFiles("*.xml", SearchOption.AllDirectories).Select(f => f.FullName).ToArray();
                
                //создаём таблицу
                dt = CreateTable();
                DataRow newRow = null;
                // вывод первого списка файлов
                for (int i = 0; i < files.Length; i++)
                {
                    // если существует данный файл
                    if (File.Exists(files[i]))
                    {
                        
                    //загружаем xml файл
                    XDocument xDoc = XDocument.Load(files[i]);

                    //получаем все узлы в xml файле
                    foreach (XElement elm in xDoc.Descendants("person"))
                    {
                        //создаём новую запись
                        newRow = dt.NewRow();
                        //проверяем наличие атрибутов (если требуется)
                        if (elm.HasAttributes)
                        {

                //проверяем наличие атрибута active
                if (elm.Element("is_active") != null)
                {
                                //получаем значение атрибута
                                newRow["Active"] = int.Parse(elm.Element("is_active").Value);

                        }
                       
                 }
                        
                        //проверяем наличие xml элемента mobile_phone
                        if (elm.Element("mobile_phone") != null)
                        {
                                //получаем значения элемента mobile_phone
                                newRow["Phone"] = elm.Element("mobile_phone").Value;
                        }
                        
                        //проверяем наличие xml элемента firstname
                        if (elm.Element("firstname") != null)
                        {
                                //получаем значения элемента firstname
                                newRow["Firstname"] = elm.Element("firstname").Value;
                        }
                        //проверяем наличие xml элемента lastname
                        if (elm.Element("lastname") != null)
                        {
                                //получаем значения элемента firstname
                                newRow["Lastname"] = elm.Element("lastname").Value;

                        }
                            //проверяем наличие xml элемента address
                            if (elm.Element("address") != null)
                            {
                                //получаем значения элемента firstname
                                newRow["Address"] = elm.Element("address").Value;

                            }

                            //добавляем новую запись в таблицу
                            dt.Rows.Add(newRow);
                       }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return dt;
        }
        private void UpdateLabelText()
        {
            // Set the labels to reflect the current state of the DataGridView.
            label9.Text = "Всего записей: " + dataGridView1.RowCount.ToString();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = ReadXml();
            UpdateLabelText();

                      
           var sw = new StreamWriter(@"config.csv", false, Encoding.UTF8);

            foreach (DataGridViewRow row in dataGridView1.Rows) //запись
                if (!row.IsNewRow)
                {
                    var first = true;
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        if (!first) sw.Write(";");
                        sw.Write(cell.Value.ToString());
                        first = false;
                    }
                    sw.WriteLine();
                }
            sw.Close();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        
        private void button2_Click(object sender, EventArgs e)
        {
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.DataSource = ReadXml();

            if (File.Exists(@"config.csv"))
            {
                dataGridView1.DataSource = ReadCSVFile(@"config.csv"); 
            }

            comboBox1.Items.Add("Все");
            comboBox1.Items.Add("Активные");
            comboBox1.Items.Add("Неактивные");
            comboBox1.SelectedIndex = 0;
            comboBox3.Items.Add("Телефон");
            comboBox3.Items.Add("Имя");
            comboBox3.Items.Add("Фамилия");
            comboBox3.SelectedIndex = 0;
            
        }

        private DataTable ReadCSVFile(string pathToCsvFile)
        {
            DataTable dt = null;
            //создаём таблицу
            dt = CreateTable();

            try
            {
            
                DataRow dr = null;
                string[] cfgValues = null;
                string[] cfg = File.ReadAllLines(pathToCsvFile);

                for (int i = 0; i < cfg.Length; i++)
                {
                    if (!String.IsNullOrEmpty(cfg[i]))
                    {
                        cfgValues = cfg[i].Split(';');
                        //создаём новую строку
                        dr = dt.NewRow();
                        dr["Active"] = Boolean.Parse(cfgValues[0]);
                        dr["Phone"] = cfgValues[1];
                        dr["Firstname"] = cfgValues[2];
                        dr["Lastname"] = cfgValues[3];
                        dr["Address"] = cfgValues[4];
                        //добавляем строку в таблицу
                        dt.Rows.Add(dr);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return dt;
        }
        private void cохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // диалоговое окно
            var save = new SaveFileDialog
            {
                AddExtension = true,
                DefaultExt = "bin",
                Filter = @"Текстовые файлы (*.txt)|*.txt|CSV-файл (*.csv)|*.csv|Bin-файл (*.bin)|*.bin",
                FilterIndex = 2,
                RestoreDirectory = true

            };

            if (save.ShowDialog() != DialogResult.OK) return;

            var sw = new StreamWriter(save.FileName, false, Encoding.UTF8);
         
                foreach (DataGridViewRow row in dataGridView1.Rows) //запись
                    if (!row.IsNewRow)
                    {
                        var first = true;
                        foreach (DataGridViewCell cell in row.Cells)
                        {
                            if (!first) sw.Write(";");
                            sw.Write(cell.Value.ToString());
                            first = false;
                        }
                        sw.WriteLine();
                    }
            sw.Close();

        }

        private void синхронизацияEStaffToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = ReadXml();
            UpdateLabelText();
            var sw = new StreamWriter(@"config.csv", false, Encoding.UTF8);

            foreach (DataGridViewRow row in dataGridView1.Rows) //запись
                if (!row.IsNewRow)
                {
                    var first = true;
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        if (!first) sw.Write(";");
                        sw.Write(cell.Value.ToString());
                        first = false;
                    }
                    sw.WriteLine();
                }
            sw.Close();

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }


        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            AboutBox1 frm = new AboutBox1();
            frm.ShowDialog();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void настройкиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form2 frm = new Form2();
            frm.ShowDialog();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
                {
                    checkBox1.Text = "Массовая рассылка";
                    maskedTextBox1.Enabled = false;
                }
        
                else
                {
                    checkBox1.Text = "Одиночная отправка";
                    maskedTextBox1.Enabled = true;                 }
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void редакторШаблоновToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 frm = new Form3();
            frm.ShowDialog();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            //  (string pUSER, string pPASS, string pSMSNUM, string pSERVER, string pMESSAGE)
            var spUSER = "root";
            var spPASS = "root";
            var spSMSNUM = "380689418354";
            var spSERVER = "192.168.12.80";
            var spMESSAGE = richTextBox1.Text;
            var Answer = GET(spUSER, spPASS, spSMSNUM, spSERVER, spMESSAGE);
            MessageBox.Show(Answer);
        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            //Задаем расширение имени файла по умолчанию.
            ofd.DefaultExt = "*.xls;*.xlsx";
            //Задаем строку фильтра имен файлов, которая определяет
            //варианты, доступные в поле "Файлы типа" диалогового
            //окна.
            ofd.Filter = "Excel Sheet(*.xlsx)|*.xlsx";
            //Задаем заголовок диалогового окна.
            ofd.Title = "Выберите документ для загрузки данных";
            ExcelObj.Application app = new ExcelObj.Application();
            ExcelObj.Workbook workbook;
            ExcelObj.Worksheet NwSheet;
            ExcelObj.Range ShtRange;
            DataTable dt = new DataTable();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = ofd.FileName;

                workbook = app.Workbooks.Open(ofd.FileName, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value);

                //Устанавливаем номер листа из котрого будут извлекаться данные
                //Листы нумеруются от 1
                NwSheet = (ExcelObj.Worksheet)workbook.Sheets.get_Item(1);
                ShtRange = NwSheet.UsedRange;
                for (int Cnum = 1; Cnum <= ShtRange.Columns.Count; Cnum++)
                {
                    dt.Columns.Add(
                       new DataColumn((ShtRange.Cells[1, Cnum] as ExcelObj.Range).Value2.ToString()));
                }
                dt.AcceptChanges();

                string[] columnNames = new String[dt.Columns.Count];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    columnNames[0] = dt.Columns[i].ColumnName;
                }

                for (int Rnum = 2; Rnum <= ShtRange.Rows.Count; Rnum++)
                {
                    DataRow dr = dt.NewRow();
                    for (int Cnum = 1; Cnum <= ShtRange.Columns.Count; Cnum++)
                    {
                        if ((ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2 != null)
                        {
                            dr[Cnum - 1] =
                (ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2.ToString();
                        }
                    }
                    dt.Rows.Add(dr);
                    dt.AcceptChanges();
                }

                dataGridView1.DataSource = dt;
                app.Quit();
            }
            else
                // Close this window
                this.Close();
        }
        }
}
