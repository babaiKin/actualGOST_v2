using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using HtmlAgilityPack;
using System.Net;
using System.Threading;
using System.Management;
using System.Data.OleDb;
using System.Text.RegularExpressions;
using Awesomium.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Diagnostics;
using MySql.Data.MySqlClient;


namespace ActualGost
{
    public partial class Form1 : Form
    {
        string upd;
        string eXt;
        string fileName;
        string saveFileName;
        int next;

        Excel.Application ObjWorkExcel = new Excel.Application(); //открыть эксель

        public Form1()
        {
            InitializeComponent();
            textBox1.Text = "91";
            textBox2.Text = "TYm897";
            textBox1.Visible = false;
            textBox2.Visible = false;
            label4.Visible = false;
            textBox3.ReadOnly = true;
            this.Text = "Актуализация ГОСТов";
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void btnFind_Click(object sender, EventArgs e)
        {
            checkBox1.Enabled = false;
            checkBox2.Enabled = false;
            btnFind.Enabled = false;

            //таймер
            //считает время выполнения программы
            Stopwatch testStopwatch = new Stopwatch();
            testStopwatch.Start();

            //eXt = Path.GetExtension(openFileDialog1.SafeFileName);
            //fileName = openFileDialog1.FileName;

            if (checkBox1.Checked & checkBox2.Checked)
            {
                try
                {
                    OpenFileDialog openFileDialog1 = new OpenFileDialog();
                    openFileDialog1.Filter = "Доступные форматы (*.txt ; *.xls ; *.xlsx)|*.txt; *.xls; *.xlsx";
                    openFileDialog1.FilterIndex = 2;
                    openFileDialog1.RestoreDirectory = true;
                    openFileDialog1.Title = "Select File";

                    if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        eXt = Path.GetExtension(openFileDialog1.SafeFileName);
                        fileName = openFileDialog1.FileName;

                        BellGiss();
                        Predstavitelstvo();

                        testStopwatch.Stop();
                        TimeSpan tSpan; tSpan = testStopwatch.Elapsed;
                        MessageBox.Show("Время выполнения операции - " + tSpan.ToString());

                        //MessageBox.Show("oooops.... sorry, not done yet");
                    }
                    //ObjWorkExcel.Visible = true;
                }
                catch (Exception err)
                { MessageBox.Show("FATAL ERROR: " + err); }
              }
            /*try
            {
                checkBox1.Enabled = false;
                checkBox2.Enabled = false;
                btnFind.Enabled = false;

                Predstavitelstvo();
                next = 2;
                BellGiss();

                checkBox1.Enabled = true;
                checkBox2.Enabled = true;
                btnFind.Enabled = true;
            }
            catch
            { MessageBox.Show("FATAL ERROR"); }*/

            else if (checkBox2.Checked)
                try
                {
                    OpenFileDialog openFileDialog1 = new OpenFileDialog();
                    openFileDialog1.Filter = "Доступные форматы (*.txt ; *.xls ; *.xlsx)|*.txt; *.xls; *.xlsx";
                    openFileDialog1.FilterIndex = 2;
                    openFileDialog1.RestoreDirectory = true;
                    openFileDialog1.Title = "Select File";

                    if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        eXt = Path.GetExtension(openFileDialog1.SafeFileName);
                        fileName = openFileDialog1.FileName;
                        saveFileName = openFileDialog1.SafeFileName;

                        Predstavitelstvo();

                        testStopwatch.Stop();
                        TimeSpan tSpan; tSpan = testStopwatch.Elapsed;
                        MessageBox.Show("Время выполнения операции - " + tSpan.ToString());

                        //ObjWorkExcel.Visible = true;
                    }

                }
                catch (Exception err)
                { MessageBox.Show("FATAL ERROR: " + err); }

            else if (checkBox1.Checked)
                try
                {
                    OpenFileDialog openFileDialog1 = new OpenFileDialog();
                    openFileDialog1.Filter = "Доступные форматы (*.txt ; *.xls ; *.xlsx)|*.txt; *.xls; *.xlsx";
                    openFileDialog1.FilterIndex = 2;
                    openFileDialog1.RestoreDirectory = true;
                    openFileDialog1.Title = "Select File";

                    if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        eXt = Path.GetExtension(openFileDialog1.SafeFileName);
                        fileName = openFileDialog1.FileName;
                        saveFileName = openFileDialog1.SafeFileName;

                        BellGiss();

                        testStopwatch.Stop();
                        TimeSpan tSpan; tSpan = testStopwatch.Elapsed;
                        MessageBox.Show("Время выполнения операции - " + tSpan.ToString());

                        //ObjWorkExcel.Visible = true;
                    }
                }
                catch (Exception err)
                { MessageBox.Show("FATAL ERROR: " + err); }

            checkBox1.Enabled = true;
            checkBox2.Enabled = true;
            btnFind.Enabled = true;

            Process[] ps1 = System.Diagnostics.Process.GetProcessesByName("EXCEL");
            foreach (Process p1 in ps1)
                p1.Kill();
            this.Close();
        }


        //БелГИСС
        public void BellGiss()
        {
            //int numCol = Convert.ToInt32(textBox2.Text);

            Object missingObj = System.Reflection.Missing.Value;
            Object trueObj = true;
            Object falseObj = false;

            progressBar1.Value = 0;
            progressBar1.Visible = true;
            label4.Visible = true;
            label4.Text = "Идет проверка по базе БелГИССа...";
            richTextBox1_output.Clear();
            btnFind.Enabled = false;

            logIn();            // функция авторизации
                                // остановка программы на x секунд 
            while (webControl1.IsLoading)
            {
                Application.DoEvents();
                //Thread.Sleep(1000);
            }

            //обрабатывание xls/xlsx файлов
            //if (Path.GetExtension(eXt) == ".xls" | Path.GetExtension(eXt) == ".xlsx")
            //{
            //добавление каждого ГОСТа в textBox
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(fileName, Type.Missing,
                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                   Type.Missing, Type.Missing, Type.Missing); //открыть файл
                        Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //получить 1 лист

                        var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//1 ячейку
                        string[] str = new string[lastCell.Row]; // массив значений с листа равен по размеру листу
                        string pattern = @"((\bСТБ\b)(\D)*([^А-я,(])*)";
                        Regex regex = new Regex(pattern);

                        for (int j = 3; j < lastCell.Row; j++) // по всем строкам
                        {
                            for (int currentcol = 5; currentcol <= 6; currentcol++)
                            {
                                str[j] = ObjWorkSheet.Cells[j + 3, currentcol].Text.ToString(); //считываем текст в строку

                                string input = str[j];

                                Match match = regex.Match(input);
                                if (match.Value != null)
                                    textBox3.Text = match.Value;

                                if (textBox3.Text != "") // проверка на пустые строки
                                {
                                    FindBell();              // функция выполнения поиска госта
                                    //WriteToRTB();          // функция записи в rtb

                                    ObjWorkSheet.Cells[j + 3, currentcol + 2].Value = upd;

                                    if (label2.Text == "1")
                                    {
                                        ObjWorkSheet.Cells[j + 3, currentcol + 2].Value = upd /*textBox1.Text + " действующий"*/;
                                        ObjWorkSheet.Cells[j + 3, currentcol + 2].Font.Color = Excel.XlRgbColor.rgbBlack;
                                    }

                                    else if (label2.Text == "2")
                                    {
                                        ObjWorkSheet.Cells[j + 3, currentcol + 2].Value = upd;
                                        ObjWorkSheet.Cells[j + 3, currentcol + 2].Font.Color = Excel.XlRgbColor.rgbRed;
                                    }

                                    else
                                    {
                                        //richTextBox1.AppendText("Информация по ГОСТу не найдена...\n");
                                        ObjWorkSheet.Cells[j + 3, currentcol + 2].Value = textBox3.Text + " Информация по ГОСТу не найдена...";
                                        ObjWorkSheet.Cells[j + 3, currentcol + 2].Font.Color = Excel.XlRgbColor.rgbOrangeRed;
                                    }
                                }
                            }

                            progressBar1.Maximum = lastCell.Row;
                            progressBar1.Value = j;
                        }

                        ObjWorkExcel.DisplayAlerts = false;
                        ObjWorkSheet.SaveAs(fileName);
                        ObjWorkExcel.Interactive = true;
                        ObjWorkExcel.ScreenUpdating = true;
                        ObjWorkExcel.UserControl = true;

                        btnFind.Enabled = true;
                    //}
                
                //webControl2.Dispose();
                progressBar1.Visible = false;
                label4.Visible = false;
                webControl2.Dispose();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                //}
        }

        public void logIn()
        {
            webControl2.Source = new Uri("https://ips3.belgiss.by/index.php");

            //ожидание полной загрузки страницы
            while (webControl2.IsLoading)
            {
                Application.DoEvents();
            }

            //поиск логина/пароля по айди блока
            //внесение данных в поле
            //нажатие кнопки
            try
            {
                webControl2.ExecuteJavascript("document.getElementById('form_auth_login').value=" + textBox1.Text);
                webControl2.ExecuteJavascript("document.getElementById('form_auth_password').value='" + textBox2.Text + "'");
                webControl2.ExecuteJavascript("$('*[value=Войти]').click()");
            }
            catch
            {  }
        }

        public void FindBell()
        {
            label2.Text = "";
            webControl2.Source = new Uri("https://ips3.belgiss.by/Search.php?fullseek=" + textBox3.Text);

            //ожидание полной загрузки страницы
            while (webControl2.IsLoading)
            {
                Application.DoEvents();
            }

            loadStatus();

            try
            {
                var html = webControl2.ExecuteJavascriptWithResult("document.documentElement.outerHTML").ToString();
                //var doc = new HtmlAgilityPack.HtmlDocument();
                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(html);

                int stri;     // начальная строка
                int col;      // колонки
                var count = Regex.Matches(html, @"<tr role[^<]*>.*?</tr>", RegexOptions.Singleline).Count; //количество строк <tr> на странице
                upd = "";

                for (stri = 2; stri <= count+1 ; stri++)  //откидываем лишние -11 строк, не относящиеяся к нашей таблице
                {
                    //вывод названия госта в rtb
                    HtmlNode gost = doc.DocumentNode.SelectSingleNode("/html/body/div[5]/div[2]/div[2]/div[1]/div[1]/div[1]/div[3]/div[4]/div[1]/table/tbody/tr[" + stri + "]/td[2]/a");
                    richTextBox1_output.AppendText(gost.InnerText.Trim() + ":\n");
                    upd = upd + gost.InnerText.Trim() + " ";

                    HtmlNode docType = doc.DocumentNode.SelectSingleNode("/html/body/div[5]/div[2]/div[2]/div[1]/div[1]/div[1]/div[3]/div[4]/div[1]/table/tbody/tr[" + stri + "]/td[5]");  //вид документа
                    HtmlNode change = doc.DocumentNode.SelectSingleNode("/html/body/div[5]/div[2]/div[2]/div[1]/div[1]/div[1]/div[3]/div[4]/div[1]/table/tbody/tr[" + stri + "]/td[8]");   //заменяющий
                    HtmlNode working = doc.DocumentNode.SelectSingleNode("/html/body/div[5]/div[2]/div[2]/div[1]/div[1]/div[1]/div[3]/div[4]/div[1]/table/tbody/tr[" + stri + "]/td[12]"); //действущий/недействующий

                    if (docType != null)
                        if (docType.InnerText.Trim().Replace("&nbsp;", "") != "")  //если поле пустое, то пропускает его.
                        {
                            upd = upd + " " + docType.InnerText.Trim().Replace("&nbsp;", "");
                        }

                    if (change != null)
                        if (change.InnerText.Trim().Replace("&nbsp;", "") != "")  //если поле пустое, то пропускает его.
                        {
                            upd = upd + " Заменяющий: " + change.InnerText.Trim().Replace("&nbsp;", "");
                        }

                    if (working != null)
                    {
                        if (working.InnerText.Trim().Replace("&nbsp;", "") == "Недействующий НД")
                        {
                            label2.Text = "2";
                            upd = upd + " " + working.InnerText.Trim().Replace("&nbsp;", "") + "\n";
                            //return;
                        }

                        else
                        {
                            label2.Text = "1";
                            upd = upd + " " + working.InnerText.Trim().Replace("&nbsp;", "") + "\n";
                            //return;
                        }
                    }
                }
               
                html = null;
                doc = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
            catch { }

        }

        public void loadStatus()
        {
            /*while (webControl2.IsLoading)
            {
                Application.DoEvents();
            }
            */
            Thread.Sleep(500);
            bool ex = true;
            while (ex = true)
            {
                //style загрузка
                var html = webControl2.ExecuteJavascriptWithResult("document.documentElement.outerHTML").ToString();

                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(html);

                HtmlNode loadStatus = doc.DocumentNode.SelectSingleNode("/html/body/div[5]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]");

                try
                {
                    if (loadStatus.Attributes["style"].Value == "display: block; ")
                    {
                        while (webControl2.IsLoading)
                        {
                            Application.DoEvents();
                        }
                    }

                    else
                    {
                        //return;
                        break;
                    }
                }
                catch
                { //return;
                    break; 
                }
            }
        }

        //Представительство
        string constr = "Server=servgost;" +
                                    "port=3306;" +
                                    "Database=gost;" +
                                    "Uid=admin;" +
                                    "Pwd=;" +
                                    "CharSet = cp1251;" +
                                    "Allow Zero Datetime=true; ";
        MySqlConnection mycon;
        MySqlCommand mycom;
        MySqlCommand mycom2;

        public void Predstavitelstvo()
        {
            mycon = new MySqlConnection(constr);

            textBox2.Text = "1";

            try
            {
                mycon = new MySqlConnection(constr);
                mycon.Open();

                mycom = new MySqlCommand(@"SELECT * FROM gost.s_service ", mycon);
                //        MessageBox.Show("CONNECTED !");
                mycom.CommandType = CommandType.Text;

                MySqlDataAdapter adapter = new MySqlDataAdapter();

                adapter.TableMappings.Add("Table", "s_service");

                adapter.SelectCommand = mycom;

                DataTable dataTable = new DataTable();
                DataSet dataSet = new DataSet("s_service");
                adapter.Fill(dataTable);
                dataGridView1.DataSource = dataTable;
                mycon.Close();
            }

            catch (InvalidCastException ee)
            {
                MessageBox.Show("Нет подключения к серверу" + ee.Message);
            }

            progressBar1.Value = 0;
            progressBar1.Visible = true;
            label4.Visible = true;
            label4.Text = "Идет проверка по базе Представительства...";

            if (textBox2.Text != "")
            {
                int numCol = Convert.ToInt32(textBox2.Text); //переменная для номера колонки

                //OpenFileDialog openFileDialog1 = new OpenFileDialog();
                //openFileDialog1.Filter = "Доступные форматы (*.xls ; *.xlsx)|*.xls; *.xlsx";
                //openFileDialog1.FilterIndex = 2;
                //openFileDialog1.RestoreDirectory = true;
                //openFileDialog1.Title = "Select an Excel File";

                //if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                //{
                    

                    btnFind.Enabled = false;
                    //button2.Enabled = false;
                    //textBox1.ReadOnly = true;
                    textBox2.ReadOnly = true;
                    richTextBox1_output.Clear();

                    // цикл по вытаскиванию строчки из xls
                    //string fileName = openFileDialog1.FileName;

                    //Excel.Application ObjWorkExcel = new Excel.Application(); //открыть эксель
                    Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(fileName, Type.Missing,
                               Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                               Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                               Type.Missing, Type.Missing, Type.Missing); //открыть файл
                    Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //получить 1 лист

                    var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//1 ячейку
                    string[] str = new string[lastCell.Row]; // массив значений с листа равен по размеру листу

                    for (int j = 0; j < lastCell.Row; j++) // по всем строкам
                    {
                        //for (int currentcol = numCol; currentcol <= numCol + 1; currentcol++)
                        for (int currentcol = numCol; currentcol <= 6; currentcol++)
                        {
                            str[j] = ObjWorkSheet.Cells[j + 1, currentcol].Text.ToString();//считываем текст в строку

                            //вытаскиваем регулярками ГОСТ
                            string input = str[j];
                            string pattern = @"((\bГОСТ\b)(\D)*([^А-я,(])*)";
                            Regex regex = new Regex(pattern);

                            // Получаем совпадения в экземпляре класса Match
                            Match match = regex.Match(input);
                            // отображаем все совпадения
                            //MessageBox.Show("" + match.Value);
                            //записываем в textBox
                            if (match.Value != null)
                                textBox1.Text = match.Value;

                            if (textBox1.Text != "") // проверка на пустые строки
                            {
                                FindPred();              // функция выполнения поиска госта
                                //WriteToRTB();        // функция записи в rtb

                                ObjWorkSheet.Cells[j + 1, currentcol + 2].Value = upd;

                                if (label2.Text == "1")
                                {
                                    ObjWorkSheet.Cells[j + 1, currentcol + 2].Value = upd /*textBox1.Text + " действующий"*/;
                                    ObjWorkSheet.Cells[j + 3, currentcol + 2].Font.Color = Excel.XlRgbColor.rgbBlack;
                                }

                                else if (label2.Text == "2")
                                {
                                    ObjWorkSheet.Cells[j + 1, currentcol + 2].Value = upd;
                                    ObjWorkSheet.Cells[j + 1, currentcol + 2].Font.Color = Excel.XlRgbColor.rgbRed;
                                }

                                else
                                {
                                    //richTextBox1.AppendText("Информация по ГОСТу не найдена...\n");
                                    ObjWorkSheet.Cells[j + 1, currentcol + 2].Value = textBox1.Text + "  Информация по ГОСТу не найдена...";
                                    ObjWorkSheet.Cells[j + 1, currentcol + 2].Font.Color = Excel.XlRgbColor.rgbOrangeRed;
                                }
                            }
                           
                            progressBar1.Maximum = lastCell.Row;
                            progressBar1.Value = j;
                        }
                    }
                    //ObjWorkExcel.Visible = true;
                    btnFind.Enabled = true;                   
                    //textBox1.ReadOnly = false;
                    textBox2.ReadOnly = false;
                    //object missing = Type.Missing;
                    label4.Visible = false;

                
                ObjWorkExcel.DisplayAlerts = false;
                ObjWorkSheet.SaveAs(fileName);
                ObjWorkExcel.Interactive = true;
                ObjWorkExcel.ScreenUpdating = true;
                ObjWorkExcel.UserControl = true;
            }
            else
                MessageBox.Show("Укажите номер колонки!");

            

            progressBar1.Visible = false;
            label4.Visible = false;

            FindWithChanges();

            dataGridView1.Dispose();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            //mycon.Close();
        }

        public void FindPred()
        {
            label2.Text = "";
            richTextBox1_output.AppendText(textBox1.Text + "\n");
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1[2, i].FormattedValue.ToString().Contains(textBox1.Text.Trim()))
                {
                    upd = Convert.ToString(dataGridView1[2, i].Value) + " ";
                    //if (Convert.ToString(dataGridView1[8, i].Value) == "")
                    //{
                    //label2.Text = "0";
                    //richTextBox1.AppendText("Нет информации\n");
                    //upd = upd + "Нет информации";
                    //return;
                    //}

                    if (Convert.ToString(dataGridView1[8, i].Value) == "Утратил силу в РФ" | Convert.ToString(dataGridView1[8, i].Value) == "Отменен" | Convert.ToString(dataGridView1[8, i].Value) == "Заменен")
                    {
                        label2.Text = "2";
                        //richTextBox1.AppendText(Convert.ToString(dataGridView1[8, i].Value));
                        upd = upd + Convert.ToString(dataGridView1[8, i].Value);
                        return;
                    }

                    else
                    {
                        label2.Text = "1";
                        //richTextBox1.AppendText("Действующий НД\n");
                        upd = upd + Convert.ToString(dataGridView1[8, i].Value);
                        //upd = "Действующий НД";
                        
                        return;
                    }
                }
            }
        }

        public void FindWithChanges()
        {
            try
            {
                mycon.Open();

                mycom2 = new MySqlCommand(@"SELECT * FROM gost.s_sub_in_part ", mycon);
                mycom2.CommandType = CommandType.Text;

                MySqlDataAdapter adapter2 = new MySqlDataAdapter();

                adapter2.TableMappings.Add("Table", "s_sub_in_part");

                adapter2.SelectCommand = mycom2;

                DataTable dataTable2 = new DataTable();
                DataSet dataSet2 = new DataSet("s_sub_in_part");
                adapter2.Fill(dataTable2);
                dataGridView2.DataSource = dataTable2;
                mycon.Close();

            }

            catch (InvalidCastException ee)
            {
                MessageBox.Show("Нет подключения к серверу" + ee.Message);
            }

            progressBar1.Value = 0;
            progressBar1.Visible = true;
            label4.Visible = true;
            label4.Text = "Идет поиск ГОСТов с изменениями в базе Представительства...";

            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(fileName, Type.Missing,
                               Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                               Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                               Type.Missing, Type.Missing, Type.Missing); //открыть файл
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //получить 1 лист

            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell); //1 ячейку
            string[] str = new string[lastCell.Row]; // массив значений с листа равен по размеру листу

            /*
            ObjWorkSheet.Cells[lastCell.Row , 1].Value = "Действует, но с ичастичными зменениями";
            ObjWorkSheet.Cells[lastCell.Row , 1].Font.Color = Excel.XlRgbColor.rgbGreen;

            ObjWorkSheet.Cells[lastCell.Row , 2].Value = "Незначительные ошибки, проблемы";
            ObjWorkSheet.Cells[lastCell.Row , 2].Font.Color = Excel.XlRgbColor.rgbOrange;

            ObjWorkSheet.Cells[lastCell.Row , 3].Value = "Отменен";
            ObjWorkSheet.Cells[lastCell.Row , 3].Font.Color = Excel.XlRgbColor.rgbRed;
            */

            for (int j = 0; j < lastCell.Row; j++) // по всем строкам
            {
                for (int currentcol = 7; currentcol <= 8; currentcol++)
                {
                    if (ObjWorkSheet.Cells[j + 1, currentcol].Text.ToString() != "")
                    {
                        str[j] = ObjWorkSheet.Cells[j + 1, currentcol].Text.ToString();//считываем текст в строку

                        //вытаскиваем регулярками Действующий ГОСТ
                        string input = str[j];
                        string pattern = @"(\bДействует\b)";
                        Regex regex = new Regex(pattern);

                        // Получаем совпадения в экземпляре класса Match
                        Match match = regex.Match(input);

                        if (match.Value == "Действует")
                        {
                            string[] str2 = new string[lastCell.Row];
                            str2[j] = ObjWorkSheet.Cells[j + 1, currentcol-2].Text.ToString();//считываем текст в строку

                            //вытаскиваем регулярками ГОСТ
                            string input2 = str2[j];
                            string pattern2 = @"((\bГОСТ\b)(\D)*([^А-я,(])*)";
                            Regex regex2 = new Regex(pattern2);

                            // Получаем совпадения в экземпляре класса Match
                            Match match2 = regex2.Match(input2);
                            // отображаем все совпадения
                            //MessageBox.Show("" + match.Value);
                            //записываем в textBox
                            
                            //if (match2.Value != null)
                            //    textBox1.Text = match2.Value;

                            //MessageBox.Show(ObjWorkSheet.Cells[j + 1, currentcol].Text.ToString() + " || " + match2.Value + " || yeap!");
                            
                            //поиск в другой таблице
                            for (int i = 0; i < dataGridView2.RowCount; i++)
                            {
                                //MessageBox.Show(dataGridView2[3, i].FormattedValue.ToString() + " || " + match2.Value.Trim());
                                //MessageBox.Show(match2.Value.Trim() + " заменен " + Convert.ToString(dataGridView2[3, i].Value));
                                if (dataGridView2[2, i].FormattedValue.ToString().Contains(match2.Value.Trim()))
                                {
                                    upd = Convert.ToString(dataGridView2[2, i].Value) + " заменен " + Convert.ToString(dataGridView2[3, i].Value);
                                    ObjWorkSheet.Cells[j + 1, currentcol].Value = upd;
                                    ObjWorkSheet.Cells[j + 1, currentcol].Font.Color = Excel.XlRgbColor.rgbGreen;
                                    //MessageBox.Show(Convert.ToString(dataGridView2[2, i].Value) + " заменен " + Convert.ToString(dataGridView2[3, i].Value));
                                    //return;
                                }
                            }
                        }
                    }
                }
                progressBar1.Maximum = lastCell.Row;
                progressBar1.Value = j;
            }

            progressBar1.Visible = false;
            label4.Visible = false;
            ObjWorkExcel.DisplayAlerts = false;
            ObjWorkSheet.SaveAs(fileName);
            ObjWorkExcel.Interactive = true;
            ObjWorkExcel.ScreenUpdating = true;
            ObjWorkExcel.UserControl = true;


            dataGridView2.Dispose();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
    }
}


    //progressBar1.Value = 0;
    //progressBar1.Visible = true;
    //progressBar1.Maximum = lastCell.Row;
    //progressBar1.Value = j;
    //progressBar1.Visible = false;