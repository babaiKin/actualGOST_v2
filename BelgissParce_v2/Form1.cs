using System;
using System.Windows.Forms;
using HtmlAgilityPack;
using System.Threading;
using System.Text.RegularExpressions;
using Awesomium.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Diagnostics;
using System.Net.Mail;
using System.Net;

using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management;
using System.Data.OleDb;


namespace ActualGost
{
    public partial class Form1 : Form
    {
        int lastColl = 8;
        //int last;
        string upd;
        string eXt;
        string fileName;
        string saveFileName;
        //int next;

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

        bool workFlag = true; //флаг для определения ошибки при проверки по базам. Если все ровно, то не меняется, иначе - false
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
                        AIS_RST();

                        testStopwatch.Stop();
                        TimeSpan tSpan; tSpan = testStopwatch.Elapsed;
                        MessageBox.Show("Время выполнения операции - " + tSpan.ToString());

                        //MessageBox.Show("oooops.... sorry, not done yet");
                    }
                    //ObjWorkExcel.Visible = true;
                }
                catch (Exception err)
                { MessageBox.Show("FATAL ERROR: " + err); workFlag = false; }
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

            else if (checkBox2.Checked) //проверка только по АИС РСТ
                try
                {
                    OpenFileDialog openFileDialog1 = new OpenFileDialog();
                    openFileDialog1.Filter = "Доступные форматы (*.txt ; *.xls ; *.xlsx)|*.txt; *.xls; *.xlsx";
                    openFileDialog1.FilterIndex = 2;
                    openFileDialog1.RestoreDirectory = true;
                    openFileDialog1.Title = "Select File";

                    if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        //label1.Text = Path.GetFileNameWithoutExtension(openFileDialog1.FileName);
                       
                        eXt = Path.GetExtension(openFileDialog1.SafeFileName);
                        fileName = openFileDialog1.FileName;
                        saveFileName = openFileDialog1.SafeFileName;
                        AIS_RST();

                        testStopwatch.Stop();
                        TimeSpan tSpan; tSpan = testStopwatch.Elapsed;
                        MessageBox.Show("Время выполнения операции - " + tSpan.ToString());

                        //ObjWorkExcel.Visible = true;
                    }
                    
                    
                }
                catch (Exception err)
                { MessageBox.Show("FATAL ERROR: " + err); workFlag = false; }

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
                { MessageBox.Show("FATAL ERROR: " + err); workFlag = false; }
            checkBox1.Enabled = true;
            checkBox2.Enabled = true;
            btnFind.Enabled = true;

            Process[] ps1 = System.Diagnostics.Process.GetProcessesByName("EXCEL");
            foreach (Process p1 in ps1)
                p1.Kill();

            //MessageBox.Show(fileName + " || " + saveFileName);


            if (workFlag == true)
            {
                try
                {
                    //CreateTestMessage1("exchange.nncsm.ru", 587);
                }
                catch
                { MessageBox.Show("ERROR"); }
                MessageBox.Show("Сообщение усепшно отправлено адресату");
            }
            this.Close();
        }

        public void CreateTestMessage1(string server, int port)
        {
            string from = "nikolaevn@nncsm.ru";
            string to = "chubanova@nncsm.ru";
            string subject = "Проверенная ОА";
            string body = "Проверенный файл ОА : " + saveFileName;
            string file = fileName;

            MailMessage message = new MailMessage(from, to, subject, body);

            Attachment data = new Attachment(file);
            message.Attachments.Add(data);

            SmtpClient client = new SmtpClient(server, port);
            
            // Credentials are necessary if the server requires the client 
            // to authenticate before it will send e-mail on the client's behalf.
            client.Credentials = CredentialCache.DefaultNetworkCredentials;
            //client.Credentials = new System.Net.NetworkCredential("login@mail.ru", "password"); // Указываем логин и пароль для авторизации

            try
            {
                client.Send(message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception caught in CreateTestMessage1(): {0}",
                            ex.ToString());
            }


            ///
            ///отправить это же сообщение лично мне для проверки
            ///

            MailMessage message2 = new MailMessage(from, from, subject, body);

            message2.Attachments.Add(data);
            client.Credentials = CredentialCache.DefaultNetworkCredentials;

            try
            {
                client.Send(message2);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception caught in CreateTestMessage1(): {0}",
                            ex.ToString());
            }
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

                        for (int j = 0; j < lastCell.Row; j++) // по всем строкам
                        {
                            for (int currentcol = 1; currentcol <= lastColl; currentcol++)
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

                                    ObjWorkSheet.Cells[j + 3, currentcol + lastColl].Value = upd;

                                    if (label2.Text == "1")
                                    {
                                        ObjWorkSheet.Cells[j + 3, currentcol + lastColl ].Value = upd /*textBox1.Text + " действующий"*/;
                                        ObjWorkSheet.Cells[j + 3, currentcol + lastColl ].Font.Color = Excel.XlRgbColor.rgbBlack;
                                    }

                                    else if (label2.Text == "2")
                                    {
                                        ObjWorkSheet.Cells[j + 3, currentcol + lastColl ].Value = upd;
                                        ObjWorkSheet.Cells[j + 3, currentcol + lastColl ].Font.Color = Excel.XlRgbColor.rgbRed;
                                    }

                                    else
                                    {
                                        //richTextBox1.AppendText("Информация по ГОСТу не найдена...\n");
                                        ObjWorkSheet.Cells[j + 3, currentcol + lastColl].Value = textBox3.Text + " Информация по ГОСТу не найдена...";
                                        ObjWorkSheet.Cells[j + 3, currentcol + lastColl].Font.Color = Excel.XlRgbColor.rgbOrangeRed;
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
                
                progressBar1.Visible = false;
                label4.Visible = false;
            if (checkBox1.Checked & checkBox2.Checked)
            { }
            else
                webControl2.Dispose();
            //webControl2.Dispose();
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
                //int col;      // колонки
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













        //АИС РСТ
        public void AIS_RST()
        {
            Object missingObj = System.Reflection.Missing.Value;
            Object trueObj = true;
            Object falseObj = false;

            progressBar1.Value = 0;
            progressBar1.Visible = true;
            label4.Visible = true;
            label4.Text = "Идет проверка по базе АИС РСТ...";
            richTextBox1_output.Clear();
            btnFind.Enabled = false;

            //Awesomium.Windows.Forms.WebControl webControl2 = new Awesomium.Windows.Forms.WebControl();
            logInAIS_RST();     // функция авторизации
                                // остановка программы на x секунд 
            //webControl2.Dispose();
            //GC.Collect();

            Thread.Sleep(1000);

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
            string pattern = @"((\bГОСТ\b)(\D)*([^А-я,(])*)";
            Regex regex = new Regex(pattern);

            int currentcol;
            for (int j = 0; j < lastCell.Row; j++) // по всем строкам
            {
                for (currentcol = 1; currentcol <= lastColl; currentcol++)
                {
                    str[j] = ObjWorkSheet.Cells[j + 3, currentcol].Text.ToString(); //считываем текст в строку

                    string input = str[j];

                    Match match = regex.Match(input);
                    if (match.Value != null)
                        textBox3.Text = match.Value;
                    //else
                        //currentcol++;

                    if (textBox3.Text != "") // проверка на пустые строки
                    {
                        FindAIS_RST();              // функция выполнения поиска госта
                                                    //WriteToRTB();          // функция записи в rtb
                        ObjWorkSheet.Cells[j + 3, currentcol + lastColl].Value = upd;

                        if (label2.Text == "1")
                        {
                            ObjWorkSheet.Cells[j + 3, currentcol + lastColl].Value = upd /*textBox1.Text + " действующий"*/;
                            ObjWorkSheet.Cells[j + 3, currentcol + lastColl].Font.Color = Excel.XlRgbColor.rgbBlack;
                        }

                        else if (label2.Text == "2")
                        {
                            ObjWorkSheet.Cells[j + 3, currentcol + lastColl].Value = upd;
                            ObjWorkSheet.Cells[j + 3, currentcol + lastColl].Font.Color = Excel.XlRgbColor.rgbRed;
                        }

                        else
                        {
                            //richTextBox1.AppendText("Информация по ГОСТу не найдена...\n");
                            ObjWorkSheet.Cells[j + 3, currentcol + lastColl].Value = upd;
                            ObjWorkSheet.Cells[j + 3, currentcol + lastColl].Font.Color = Excel.XlRgbColor.rgbOrangeRed;
                        }
                        //break;
                        GC.Collect();
                    }
                    
                }

                progressBar1.Maximum = lastCell.Row;
                progressBar1.Value = j;
                label1.Text = (j + " строка из " + lastCell.Row);
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

        public void logInAIS_RST()
        {
            webControl2.Source = new Uri("http://clnrasp.gostinfo.ru/");

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
                webControl2.ExecuteJavascript("document.getElementById('Email').value=" + "'gost@nncsm.ru'");
                webControl2.ExecuteJavascript("document.getElementById('Password').value=" + "'taburet18'");
                webControl2.ExecuteJavascript("$('*[value=Войти]').click()");
            }
            catch
            { }
        }



        //более не используется
        //перенес все в одну функцию
        //FindAIS_RST
        //был косяк с using(WebView)
        //хз как иначе это можно исправить
        public void GOSTout()
        {
            using (WebView webView =
                    WebCore.CreateWebView(1024, 768))
            {
                //---------------------------------------------------------------------------------------------//
                //--------------------------------- вывод информации по ГОСТу ---------------------------------//
                webView.Source = new Uri("http://clnrasp.gostinfo.ru/Material/details?id=" + dataID);

                try
                {
                    Thread.Sleep(1000);
                    //MessageBox.Show("" + dataID);
                    var html = webView.ExecuteJavascriptWithResult("document.documentElement.outerHTML").ToString();
                    HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                    doc.LoadHtml(html);

                    HtmlNode gostN = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[1]/div[1]/table/tbody/tr[1]/th");
                    HtmlNode gostV = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[1]/div[1]/table/tbody/tr[1]/td");
                    upd = upd + "\n" + gostN.InnerText.Trim() + " : " + gostV.InnerText.Trim();

                    HtmlNode statusN = doc.DocumentNode.SelectSingleNode(" /html/body/div/div/div/div[1]/div/div[2]/p");
                    status = statusN.InnerText.Trim(); //можно и без этого, сразу в условие подставлять иннерТекст

                    //richTextBox1_output.AppendText(status + "");
                    //MessageBox.Show("" + statusN.InnerText.Trim());


                    if (statusN.InnerText.Trim() == "Архивный")
                    {
                        ////////////////////////////////////////////////////////////////////////
                        //вывод идет по две колонки: название параметра (N) и сам параметр (V)//
                        ////////////////////////////////////////////////////////////////////////

                        label2.Text = "2"; //недействующий ГОСТ

                        int stri;     // начальная строка
                        var count = Regex.Matches(html, @"<tr class[^<]*>.*?</tr>", RegexOptions.Singleline).Count; //количество строк <tr> на странице
                        upd = "";

                        string[] text = { "Обозначение", "Статус", "Дата введения в действие", "Дата ограничения срока действия", "Обозначение заменяющего", "Примечания" };
                        int arrL = text.Length;

                        for (stri = 1; stri <= count; stri++)
                        {
                            HtmlNode textN = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[1]/div[1]/table/tbody/tr[" + stri + "]/th");
                            HtmlNode textV = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[1]/div[1]/table/tbody/tr[" + stri + "]/td");
                            int t;

                            for (t = 0; t < arrL; t++)
                            {
                                if (textN.InnerText.Trim() == text[t])
                                {
                                    if (textN.InnerText.Trim() == "Дата введения в действие")
                                    {
                                        upd = upd + "\n" + textN.InnerText.Trim() + " : " + textV.InnerText.Trim().Replace("0:00:00", "");
                                    }
                                    else if (textN.InnerText.Trim() == "Дата ограничения срока действия")
                                    {
                                        upd = upd + "\n" + textN.InnerText.Trim() + " : " + textV.InnerText.Trim().Replace("1:00:00", "");
                                    }
                                    else
                                        upd = upd + "\n" + textN.InnerText.Trim() + " : " + textV.InnerText.Trim();
                                }
                            }
                        }
                    }

                    else
                    {
                        label2.Text = "1";
                        upd = upd + "\n" + "Статус : Действует";
                        //upd = upd + " Действует";
                    }


                    richTextBox1_output.AppendText(upd + "\n");

                    html = null;
                    doc = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                }
                catch { }

                //--------------------------------- вывод информации по ГОСТу ---------------------------------//
                //---------------------------------------------------------------------------------------------//
            }
        }
        
        public void FindAIS_RST()
        {
            label2.Text = "";
            Thread.Sleep(1000);

            using (WebView webView =
                    WebCore.CreateWebView(1024, 768))
            {
                //Переменная, используемая для обозначения загрузки страницы
                bool finishedLoading = false;

                //Загружает страницу во view
                webView.Source = new Uri("http://clnrasp.gostinfo.ru/?order_by_column=sort&date_valid&order_by_direction=2&query=" + textBox3.Text);
                
                webView.LoadingFrameComplete += (sender, e) =>
                {
                    finishedLoading = true;
                };

                //ожидание полной загрузки страницы.
                while (!finishedLoading)
                {
                    Thread.Sleep(100);
                    WebCore.Update();
                }


                //webControl2.Source = new Uri("http://clnrasp.gostinfo.ru/?order_by_column=sort&date_valid&order_by_direction=2&query=" + textBox3.Text);


                //ожидание полной загрузки страницы
                //while (webControl2.IsLoading)
                //{
                //    Application.DoEvents();
                //}
                //Thread.Sleep(1000);

                //--------------------------------------------------------------------------------------------//
                //----------------------------------- блок поиска ID ГОСТа -----------------------------------//
                JSObject collection = webView.ExecuteJavascriptWithResult("document.getElementsByTagName('a')");

                if (collection == null)
                {
                    MessageBox.Show("ERROR :: collection = null"); //мб здесь ретурн или бреак надо впихать
                                                                   //return;
                }

                int length = (int)collection["length"];
                for (int i = 0; i < length; i++)
                {
                    JSObject element = collection.Invoke("item", i);
                    if (element["className"].ToString() == "rasp-details-modal")
                    {
                        hedefid = i;
                        break;
                    }
                }

                dynamic elements = (JSObject)webView.ExecuteJavascriptWithResult("document.getElementsByTagName('a')");

                if (elements == null)
                {
                    MessageBox.Show("ERROR :: elements = null"); //мб здесь ретурн или бреак  надо впихать
                }

                int lengths = (int)elements.length;

                if (lengths == 0)
                {
                    MessageBox.Show("ERROR :: lengths = 0"); //мб здесь ретурн или бреак  надо впихать
                }

                //MessageBox.Show("gaslughaslkghmawg");
                using (elements)
                {
                    upd = "";
                    //MessageBox.Show(hedefid + "");
                    if (!elements[hedefid].getAttribute("data-id"))
                    {
                        //MessageBox.Show("gut");
                        upd = upd + "\n" + "Обозначение : " + textBox3.Text;
                        upd = upd + "\n" + "Информация не найдена";
                        richTextBox1_output.AppendText(upd + "\n");
                    }

                    else
                    {
                        if (elements[hedefid].getAttribute("data-id") != string.Empty)
                        {
                            dataID = (elements[hedefid].getAttribute("data-id"));

                            
                                //---------------------------------------------------------------------------------------------//
                                //--------------------------------- вывод информации по ГОСТу ---------------------------------//
                                webView.Source = new Uri("http://clnrasp.gostinfo.ru/Material/details?id=" + dataID);

                                try
                                {
                                    Thread.Sleep(1000);
                                    //MessageBox.Show("" + dataID);
                                    var html = webView.ExecuteJavascriptWithResult("document.documentElement.outerHTML").ToString();
                                    HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                                    doc.LoadHtml(html);

                                    HtmlNode gostN = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[1]/div[1]/table/tbody/tr[1]/th");
                                    HtmlNode gostV = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[1]/div[1]/table/tbody/tr[1]/td");
                                    upd = upd + "\n" + gostN.InnerText.Trim() + " : " + gostV.InnerText.Trim();

                                    HtmlNode statusN = doc.DocumentNode.SelectSingleNode(" /html/body/div/div/div/div[1]/div/div[2]/p");
                                    status = statusN.InnerText.Trim(); //можно и без этого, сразу в условие подставлять иннерТекст

                                    //richTextBox1_output.AppendText(status + "");
                                    //MessageBox.Show("" + statusN.InnerText.Trim());


                                    if (statusN.InnerText.Trim() == "Архивный")
                                    {
                                        ////////////////////////////////////////////////////////////////////////
                                        //вывод идет по две колонки: название параметра (N) и сам параметр (V)//
                                        ////////////////////////////////////////////////////////////////////////

                                        label2.Text = "2"; //недействующий ГОСТ

                                        int stri;     // начальная строка
                                        var count = Regex.Matches(html, @"<tr class[^<]*>.*?</tr>", RegexOptions.Singleline).Count; //количество строк <tr> на странице
                                        upd = "";

                                        string[] text = { "Обозначение", "Статус", "Дата введения в действие", "Дата ограничения срока действия", "Обозначение заменяющего", "Примечания" };
                                        int arrL = text.Length;

                                        for (stri = 1; stri <= count; stri++)
                                        {
                                            HtmlNode textN = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[1]/div[1]/table/tbody/tr[" + stri + "]/th");
                                            HtmlNode textV = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[1]/div[1]/table/tbody/tr[" + stri + "]/td");
                                            int t;

                                            for (t = 0; t < arrL; t++)
                                            {
                                                if (textN.InnerText.Trim() == text[t])
                                                {
                                                    if (textN.InnerText.Trim() == "Дата введения в действие")
                                                    {
                                                        upd = upd + "\n" + textN.InnerText.Trim() + " : " + textV.InnerText.Trim().Replace("0:00:00", "");
                                                    }
                                                    else if (textN.InnerText.Trim() == "Дата ограничения срока действия")
                                                    {
                                                        upd = upd + "\n" + textN.InnerText.Trim() + " : " + textV.InnerText.Trim().Replace("1:00:00", "");
                                                    }
                                                    else
                                                        upd = upd + "\n" + textN.InnerText.Trim() + " : " + textV.InnerText.Trim();
                                                }
                                            }
                                        }

                                    }

                                    else
                                    {
                                        label2.Text = "1";
                                        upd = upd + "\n" + "Статус : Действует";
                                        //upd = upd + " Действует";
                                    }


                                    richTextBox1_output.AppendText(upd + "\n");

                                    html = null;
                                    doc = null;
                                    GC.Collect();
                                    GC.WaitForPendingFinalizers();
                                    GC.Collect();
                                }
                                catch { }

                                //--------------------------------- вывод информации по ГОСТу ---------------------------------//
                                //---------------------------------------------------------------------------------------------//
                                hedefid = 0;
                        }
                        else
                        {
                            MessageBox.Show("ERROR :: data-id = string.Empty"); //мб здесь ретурн или бреак  надо впихать 
                        }
                    }
                }
                //----------------------------------- блок поиска ID ГОСТа -----------------------------------//
                //--------------------------------------------------------------------------------------------//
                collection.Dispose();
                elements.Dispose();
                GC.Collect();
            }
            /*
            //---------------------------------------------------------------------------------------------//
            //--------------------------------- вывод информации по ГОСТу ---------------------------------//
            webControl2.Source = new Uri("http://clnrasp.gostinfo.ru/Material/details?id=" + dataID);

            try
            {
                Thread.Sleep(1000);
                //MessageBox.Show("" + dataID);
                var html = webControl2.ExecuteJavascriptWithResult("document.documentElement.outerHTML").ToString();
                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(html);
                upd = "";

                HtmlNode gostN = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[1]/div[1]/table/tbody/tr[1]/th");
                HtmlNode gostV = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[1]/div[1]/table/tbody/tr[1]/td");
                upd = upd + "\n" + gostN.InnerText.Trim() + " : " + gostV.InnerText.Trim();

                HtmlNode statusN = doc.DocumentNode.SelectSingleNode(" /html/body/div/div/div/div[1]/div/div[2]/p");
                status = statusN.InnerText.Trim(); //можно и без этого, сразу в условие подставлять иннерТекст

                //richTextBox1_output.AppendText(status + "");
                //MessageBox.Show("" + statusN.InnerText.Trim());


                if (statusN.InnerText.Trim() == "Архивный")
                {
                    ////////////////////////////////////////////////////////////////////////
                    //вывод идет по две колонки: название параметра (N) и сам параметр (V)//
                    ////////////////////////////////////////////////////////////////////////

                    label2.Text = "2"; //недействующий ГОСТ

                    int stri;     // начальная строка
                    var count = Regex.Matches(html, @"<tr class[^<]*>.*?</tr>", RegexOptions.Singleline).Count; //количество строк <tr> на странице
                    upd = "";

                    string[] text = {"Обозначение", "Статус", "Дата введения в действие", "Дата ограничения срока действия", "Обозначение заменяющего", "Примечания"};
                    int arrL = text.Length;

                    for (stri = 1; stri <= count; stri++)
                    {
                        HtmlNode textN = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[1]/div[1]/table/tbody/tr[" + stri + "]/th");
                        HtmlNode textV = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[1]/div[1]/table/tbody/tr[" + stri + "]/td");
                        int t;

                        for (t = 0; t < arrL; t++)
                        {
                            if (textN.InnerText.Trim() == text[t])
                            {
                                if (textN.InnerText.Trim() == "Дата введения в действие")
                                {
                                    upd = upd + "\n" + textN.InnerText.Trim() + " : " + textV.InnerText.Trim().Replace("0:00:00", "");
                                }
                                else if (textN.InnerText.Trim() == "Дата ограничения срока действия")
                                {
                                    upd = upd + "\n" + textN.InnerText.Trim() + " : " + textV.InnerText.Trim().Replace("1:00:00", "");
                                }
                                else
                                    upd = upd + "\n" + textN.InnerText.Trim() + " : " + textV.InnerText.Trim();
                            }
                        }
                    }


                        ////////не совсем правильное решение
                        ////////в разных ГОСТах этим строки расположены по разному
                        ////////
                        ////////label2.Text = "2"; //недействующий ГОСТ
                        //////////Обозначение
                        ////////HtmlNode gostN = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[1]/div[1]/table/tbody/tr[1]/th");
                        //////////HtmlNode gostV = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[1]/div[1]/table/tbody/tr[1]/td"); //прикрутил ранее

                        //////////Статус //если статус != действителен, то строки будут расположены по другому
                        ////////statusN = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[1]/div[1]/table/tbody/tr[3]/th");
                        ////////HtmlNode statusV = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[1]/div[1]/table/tbody/tr[3]/td");

                        //////////Дата введения
                        ////////HtmlNode dateInN = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[1]/div[1]/table/tbody/tr[8]/th");
                        ////////HtmlNode dateInV = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[1]/div[1]/table/tbody/tr[8]/td");

                        //////////Дата ограничения
                        ////////HtmlNode dateStopN = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[1]/div[1]/table/tbody/tr[9]/th");
                        ////////HtmlNode dateStopV = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[1]/div[1]/table/tbody/tr[9]/td");

                        //////////Обозначение заменяющего
                        ////////HtmlNode changedN = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[1]/div[1]/table/tbody/tr[15]/th");
                        ////////HtmlNode changedV = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[1]/div[1]/table/tbody/tr[15]/td");

                        //////////Примечание
                        ////////HtmlNode noteN = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[1]/div[1]/table/tbody/tr[16]/th");
                        ////////HtmlNode noteV = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[1]/div[1]/table/tbody/tr[16]/td");

                        //////////Thread.Sleep(1000);

                        //////////проверка на null
                        //////////хотя с чего бы ему тут быть...
                        ////////if (statusN != null && statusV != null)
                        ////////{
                        ////////    upd = upd + "\n" + statusN.InnerText.Trim() + " : " + statusV.InnerText.Trim();
                        ////////}

                        ////////if (gostN != null && gostV != null)
                        ////////{
                        ////////    upd = upd + "\n" + gostN.InnerText.Trim() + " : " + gostV.InnerText.Trim();
                        ////////}

                        ////////if (dateInN != null && dateInV != null)
                        ////////{
                        ////////    upd = upd + "\n" + dateInN.InnerText.Trim() + " : " + dateInV.InnerText.Trim().Replace("0:00:00", "");
                        ////////}

                        ////////if (dateStopN != null && dateStopV != null)
                        ////////{
                        ////////    upd = upd + "\n" + dateStopN.InnerText.Trim() + " : " + dateStopV.InnerText.Trim().Replace("1:00:00", "");
                        ////////}

                        ////////if (changedN != null && changedV != null)
                        ////////{
                        ////////    upd = upd + "\n" + changedN.InnerText.Trim() + " : " + changedV.InnerText.Trim();
                        ////////}

                        ////////if (noteN != null && noteV != null)
                        ////////{
                        ////////    upd = upd + "\n" + noteN.InnerText.Trim() + " : " + noteV.InnerText.Trim();
                        ////////}


                        /*
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
                        */
            //var webSession = this.webControl2.WebSession;
            //webControl2.Dispose();
            //webSession.ClearCache();
            //webSession.Dispose();
            WebCore.Update();
        }

                /*else
                {
                    label2.Text = "1";
                    upd = upd + "\n" + "Статус : Действует";
                    //upd = upd + " Действует";
                }
                    

                richTextBox1_output.AppendText(upd + "\n");

                html = null;
                doc = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
            catch { }

            //--------------------------------- вывод информации по ГОСТу ---------------------------------//
            //---------------------------------------------------------------------------------------------//
            
        }*/


        int hedefid = 0;
        string dataID;
        string status;



















        ///////////////////////////////////////////////////////////////////////////////////////
        ////////////////////////////// дальше идет тестовый шлак //////////////////////////////
        ////////////////////////////// дальше идет тестовый шлак //////////////////////////////
        ////////////////////////////// дальше идет тестовый шлак //////////////////////////////
        ///////////////////////////////////////////////////////////////////////////////////////

        ///////////////////////////////////////////////////////////////////////////////////////
        /////////////////        //////////////////////////////////////////////////////////////
        /////////////////   //   /////        /////        ///       ///       ////////////////
        /////////////////   //   /////   //   /////   //   ///   ///////   ////////////////////
        /////////////////   //   /////   //   /////   //   ///   ///////   ////////////////////
        /////////////////   //   /////        /////   //   ///       ///       ////////////////
        ///////////////            ///   //   /////   //   ///   ///////   ////////////////////
        ///////////////   //////   ///   //   ///     //   ///       ///       ////////////////
        ///////////////////////////////////////////////////////////////////////////////////////



        ///////////////////////////////////////////////////////////////////////////////////////
        ///////////////////   ///////   ///////////////////////////////////////////////////////
        ///////////////////   ///////   /////        ///        ///   ///   ///////////////////
        ///////////////////   //   //   /////   //   ///   //   ///   //   ////////////////////
        ///////////////////   //   //   /////   //   ///   //   ///   /   /////////////////////
        ///////////////////   //   //   /////   //   ///        ///      //////////////////////
        ///////////////////   //   //   /////   //   ///   //   ///   //   ////////////////////
        ///////////////////             ///     //   ///   //   ///   ///   ///////////////////
        ///////////////////////////////////////////////////////////////////////////////////////


        private void button1_Click(object sender, EventArgs e)
        {
            logInAIS_RST();

            while (webControl2.IsLoading)
            {
                Application.DoEvents();
            }
            Thread.Sleep(1000);

            label2.Text = "";
            webControl2.Source = new Uri("http://clnrasp.gostinfo.ru/?order_by_column=sort&date_valid&order_by_direction=2&query=" + "ГОСТ 1.8-2004");

            //ожидание полной загрузки страницы
            while (webControl2.IsLoading)
            {
                Application.DoEvents();
            }
            //Thread.Sleep(1000);


            //--------------------------------------------------------------------------------------------//
            //----------------------------------- блок поиска ID ГОСТа -----------------------------------//

            JSObject collection = webControl2.ExecuteJavascriptWithResult("document.getElementsByTagName('a')");
            if (collection == null)
                MessageBox.Show("ERROR"); //мб здесь ретурн или бреак надо впихать
                                          //return null;

            int length = (int)collection["length"];
            for (int i = 0; i < length; i++)
            {
                JSObject element = collection.Invoke("item", i);
                if (element["className"].ToString() == "rasp-details-modal")
                {
                    hedefid = i;
                    break;
                }
            }

            dynamic elements = (JSObject)webControl2.ExecuteJavascriptWithResult("document.getElementsByTagName('a')");

            if (elements == null)
                MessageBox.Show("ERROR :: elements = null"); //мб здесь ретурн или бреак  надо впихать
                                          //return null;

            int lengths = (int)elements.length;

            if (lengths == 0)
                MessageBox.Show("ERROR :: lengths = 0"); //мб здесь ретурн или бреак  надо впихать
                                          //return null;

            using (elements)
            {
                if (elements[hedefid].getAttribute("data-id") != string.Empty)
                    dataID = (elements[hedefid].getAttribute("data-id"));
                    //return elements[hedefid].getAttribute("data-user-id");
                else
                    MessageBox.Show("ERROR :: data-id = string.Empty"); //мб здесь ретурн или бреак  надо впихать
                                              //return null;
            }
            //----------------------------------- блок поиска ID ГОСТа -----------------------------------//
            //--------------------------------------------------------------------------------------------//
                        

            //--------------------------------------------------------------------------------------------//
            //--------------------------------- вывод информации по ГОСТу ---------------------------------//
            webControl2.Source = new Uri("http://clnrasp.gostinfo.ru/Material/details?id=" + dataID);


            try
            {
                Thread.Sleep(1000);
                //MessageBox.Show("" + dataID);
                var html = webControl2.ExecuteJavascriptWithResult("document.documentElement.outerHTML").ToString();
                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(html);
                upd = "";

                HtmlNode gostV = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[1]/div[1]/table/tbody/tr[1]/td");
                upd = upd + gostV.InnerText.Trim() + " ";

                HtmlNode statusN = doc.DocumentNode.SelectSingleNode(" /html/body/div/div/div/div[1]/div/div[2]/p");
                status = statusN.InnerText.Trim(); //можно и без этого, сразу в условие подставлять иннерТекст
                
                //richTextBox1_output.AppendText(status + "");
                //MessageBox.Show("" + statusN.InnerText.Trim());


                if (status == "Архивный")
                {
                    ////////////////////////////////////////////////////////////////////////
                    //вывод идет по две колонки: название параметра (N) и сам параметр (V)//
                    ////////////////////////////////////////////////////////////////////////

                    //Обозначение
                    HtmlNode gostN = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[1]/div[1]/table/tbody/tr[1]/th");
                    //HtmlNode gostV = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[1]/div[1]/table/tbody/tr[1]/td"); //прикрутил ранее

                    //Статус //если статус != действителен, то строки будут расположены по другому
                    statusN = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[1]/div[1]/table/tbody/tr[3]/th");
                    HtmlNode statusV = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[1]/div[1]/table/tbody/tr[3]/td");

                    //Дата введения
                    HtmlNode dateInN = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[1]/div[1]/table/tbody/tr[8]/th");
                    HtmlNode dateInV = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[1]/div[1]/table/tbody/tr[8]/td");

                    //Дата ограничения
                    HtmlNode dateStopN = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[1]/div[1]/table/tbody/tr[9]/th");
                    HtmlNode dateStopV = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[1]/div[1]/table/tbody/tr[9]/td");

                    //Обозначение заменяющего
                    HtmlNode changedN = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[1]/div[1]/table/tbody/tr[15]/th");
                    HtmlNode changedV = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[1]/div[1]/table/tbody/tr[15]/td");

                    //Примечание
                    HtmlNode noteN = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[1]/div[1]/table/tbody/tr[16]/th");
                    HtmlNode noteV = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div[1]/div[1]/table/tbody/tr[16]/td");

                    //проверка на null
                    //хотя с чего бы ему тут быть...
                    if (statusN != null && statusV != null)
                    {
                        upd = upd + "\n" + statusN.InnerText.Trim() + " : " + statusV.InnerText.Trim();
                    }

                    if (gostN != null && gostV != null)
                    {
                        upd = upd + "\n" + gostN.InnerText.Trim() + " : " + gostV.InnerText.Trim();
                    }

                    if (dateInN != null && dateInV != null)
                    {
                        upd = upd + "\n" + dateInN.InnerText.Trim() + " : " + dateInV.InnerText.Trim().Replace("0:00:00", "");
                    }

                    if (dateStopN != null && dateStopV != null)
                    {
                        upd = upd + "\n" + dateStopN.InnerText.Trim() + " : " + dateStopV.InnerText.Trim().Replace("1:00:00", "");
                    }

                    if (changedN != null && changedV != null)
                    {
                        upd = upd + "\n" + changedN.InnerText.Trim() + " : " + changedV.InnerText.Trim();
                    }

                    if (noteN != null && noteV != null)
                    {
                        upd = upd + "\n" + noteN.InnerText.Trim() + " : " + noteV.InnerText.Trim();
                    }
                    

                    /*
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
                    */
                }

                else
                    upd = upd + " Действует";

                richTextBox1_output.Text = upd;   
                 
                html = null;
                doc = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
            catch { }

            //--------------------------------- вывод информации по ГОСТу ---------------------------------//
            //---------------------------------------------------------------------------------------------//
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ///попробовать намутить создание бюллетени по гостам
            ///примерно как-то так...
            ///1-спарсить с сайта заказ и список гостов из него
            ///2-вписать в адр строку поисковый запрос с номером госта
            ///3-спарсить полученный результат, а именно краткое описание госта
        }

        // Представительство более не используется
        ////////public void Predstavitelstvo()
        ////////{
        ////////    mycon = new MySqlConnection(constr);

        ////////    textBox2.Text = "1";

        ////////    try
        ////////    {
        ////////        mycon = new MySqlConnection(constr);
        ////////        mycon.Open();

        ////////        mycom = new MySqlCommand(@"SELECT * FROM gost.s_service ", mycon);
        ////////        //        MessageBox.Show("CONNECTED !");
        ////////        mycom.CommandType = CommandType.Text;

        ////////        MySqlDataAdapter adapter = new MySqlDataAdapter();

        ////////        adapter.TableMappings.Add("Table", "s_service");

        ////////        adapter.SelectCommand = mycom;

        ////////        DataTable dataTable = new DataTable();
        ////////        DataSet dataSet = new DataSet("s_service");
        ////////        adapter.Fill(dataTable);
        ////////        dataGridView1.DataSource = dataTable;
        ////////        mycon.Close();
        ////////    }

        ////////    catch (InvalidCastException ee)
        ////////    {
        ////////        MessageBox.Show("Нет подключения к серверу" + ee.Message);
        ////////    }

        ////////    progressBar1.Value = 0;
        ////////    progressBar1.Visible = true;
        ////////    label4.Visible = true;
        ////////    label4.Text = "Идет проверка по базе Представительства...";

        ////////    if (textBox2.Text != "")
        ////////    {
        ////////        int numCol = Convert.ToInt32(textBox2.Text); //переменная для номера колонки

        ////////        //OpenFileDialog openFileDialog1 = new OpenFileDialog();
        ////////        //openFileDialog1.Filter = "Доступные форматы (*.xls ; *.xlsx)|*.xls; *.xlsx";
        ////////        //openFileDialog1.FilterIndex = 2;
        ////////        //openFileDialog1.RestoreDirectory = true;
        ////////        //openFileDialog1.Title = "Select an Excel File";

        ////////        //if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        ////////        //{


        ////////            btnFind.Enabled = false;
        ////////            //button2.Enabled = false;
        ////////            //textBox1.ReadOnly = true;
        ////////            textBox2.ReadOnly = true;
        ////////            richTextBox1_output.Clear();

        ////////            // цикл по вытаскиванию строчки из xls
        ////////            //string fileName = openFileDialog1.FileName;

        ////////            //Excel.Application ObjWorkExcel = new Excel.Application(); //открыть эксель
        ////////            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(fileName, Type.Missing,
        ////////                       Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        ////////                       Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        ////////                       Type.Missing, Type.Missing, Type.Missing); //открыть файл
        ////////            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //получить 1 лист

        ////////            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//1 ячейку
        ////////            string[] str = new string[lastCell.Row]; // массив значений с листа равен по размеру листу

        ////////            for (int j = 0; j < lastCell.Row; j++) // по всем строкам
        ////////            {
        ////////                //for (int currentcol = numCol; currentcol <= numCol + 1; currentcol++)
        ////////                for (int currentcol = numCol; currentcol <= lastColl; currentcol++)
        ////////                {
        ////////                    str[j] = ObjWorkSheet.Cells[j + 1, currentcol].Text.ToString();//считываем текст в строку

        ////////                    //вытаскиваем регулярками ГОСТ
        ////////                    string input = str[j];
        ////////                    string pattern = @"((\bГОСТ\b)(\D)*([^А-я,(])*)";
        ////////                    Regex regex = new Regex(pattern);

        ////////                    // Получаем совпадения в экземпляре класса Match
        ////////                    Match match = regex.Match(input);
        ////////                    // отображаем все совпадения
        ////////                    //MessageBox.Show("" + match.Value);
        ////////                    //записываем в textBox
        ////////                    if (match.Value != null)
        ////////                        textBox1.Text = match.Value;

        ////////                    if (textBox1.Text != "") // проверка на пустые строки
        ////////                    {
        ////////                        FindPred();              // функция выполнения поиска госта
        ////////                        //WriteToRTB();        // функция записи в rtb

        ////////                        ObjWorkSheet.Cells[j + 1, currentcol + lastColl].Value = upd;

        ////////                        if (label2.Text == "1")
        ////////                        {
        ////////                            ObjWorkSheet.Cells[j + 1, currentcol + lastColl].Value = upd /*textBox1.Text + " действующий"*/;
        ////////                            ObjWorkSheet.Cells[j + 3, currentcol + lastColl].Font.Color = Excel.XlRgbColor.rgbBlack;
        ////////                        }

        ////////                        else if (label2.Text == "2")
        ////////                        {
        ////////                            ObjWorkSheet.Cells[j + 1, currentcol + lastColl].Value = upd;
        ////////                            ObjWorkSheet.Cells[j + 1, currentcol + lastColl].Font.Color = Excel.XlRgbColor.rgbRed;
        ////////                        }

        ////////                        else
        ////////                        {
        ////////                            //richTextBox1.AppendText("Информация по ГОСТу не найдена...\n");
        ////////                            ObjWorkSheet.Cells[j + 1, currentcol + lastColl].Value = textBox1.Text + "  Информация по ГОСТу не найдена...";
        ////////                            ObjWorkSheet.Cells[j + 1, currentcol + lastColl].Font.Color = Excel.XlRgbColor.rgbOrangeRed;
        ////////                        }
        ////////                    }

        ////////                    progressBar1.Maximum = lastCell.Row;
        ////////                    progressBar1.Value = j;
        ////////                }
        ////////            }
        ////////            //ObjWorkExcel.Visible = true;
        ////////            btnFind.Enabled = true;                   
        ////////            //textBox1.ReadOnly = false;
        ////////            textBox2.ReadOnly = false;
        ////////            //object missing = Type.Missing;
        ////////            label4.Visible = false;


        ////////        ObjWorkExcel.DisplayAlerts = false;
        ////////        ObjWorkSheet.SaveAs(fileName);
        ////////        ObjWorkExcel.Interactive = true;
        ////////        ObjWorkExcel.ScreenUpdating = true;
        ////////        ObjWorkExcel.UserControl = true;
        ////////    }
        ////////    else
        ////////        MessageBox.Show("Укажите номер колонки!");



        ////////    progressBar1.Visible = false;
        ////////    label4.Visible = false;

        ////////    FindWithChanges();

        ////////    dataGridView1.Dispose();
        ////////    GC.Collect();
        ////////    GC.WaitForPendingFinalizers();
        ////////    GC.Collect();
        ////////    //mycon.Close();
        ////////}

        ////////public void FindPred()
        ////////{
        ////////    label2.Text = "";
        ////////    richTextBox1_output.AppendText(textBox1.Text + "\n");
        ////////    for (int i = 0; i < dataGridView1.RowCount; i++)
        ////////    {
        ////////        if (dataGridView1[2, i].FormattedValue.ToString().Contains(textBox1.Text.Trim()))
        ////////        {
        ////////            upd = Convert.ToString(dataGridView1[2, i].Value) + " ";
        ////////            //if (Convert.ToString(dataGridView1[8, i].Value) == "")
        ////////            //{
        ////////            //label2.Text = "0";
        ////////            //richTextBox1.AppendText("Нет информации\n");
        ////////            //upd = upd + "Нет информации";
        ////////            //return;
        ////////            //}

        ////////            if (Convert.ToString(dataGridView1[8, i].Value) == "Утратил силу в РФ" | Convert.ToString(dataGridView1[8, i].Value) == "Отменен" | Convert.ToString(dataGridView1[8, i].Value) == "Заменен")
        ////////            {
        ////////                label2.Text = "2";
        ////////                //richTextBox1.AppendText(Convert.ToString(dataGridView1[8, i].Value));
        ////////                upd = upd + Convert.ToString(dataGridView1[8, i].Value);
        ////////                return;
        ////////            }

        ////////            else
        ////////            {
        ////////                label2.Text = "1";
        ////////                //richTextBox1.AppendText("Действующий НД\n");
        ////////                upd = upd + Convert.ToString(dataGridView1[8, i].Value);
        ////////                //upd = "Действующий НД";

        ////////                return;
        ////////            }
        ////////        }
        ////////    }
        ////////}

        ////////public void FindWithChanges()
        ////////{
        ////////    try
        ////////    {
        ////////        mycon.Open();

        ////////        mycom2 = new MySqlCommand(@"SELECT * FROM gost.s_sub_in_part ", mycon);
        ////////        mycom2.CommandType = CommandType.Text;

        ////////        MySqlDataAdapter adapter2 = new MySqlDataAdapter();

        ////////        adapter2.TableMappings.Add("Table", "s_sub_in_part");

        ////////        adapter2.SelectCommand = mycom2;

        ////////        DataTable dataTable2 = new DataTable();
        ////////        DataSet dataSet2 = new DataSet("s_sub_in_part");
        ////////        adapter2.Fill(dataTable2);
        ////////        dataGridView2.DataSource = dataTable2;
        ////////        mycon.Close();

        ////////    }

        ////////    catch (InvalidCastException ee)
        ////////    {
        ////////        MessageBox.Show("Нет подключения к серверу" + ee.Message);
        ////////    }

        ////////    progressBar1.Value = 0;
        ////////    progressBar1.Visible = true;
        ////////    label4.Visible = true;
        ////////    label4.Text = "Идет поиск ГОСТов с изменениями в базе Представительства...";

        ////////    Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(fileName, Type.Missing,
        ////////                       Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        ////////                       Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        ////////                       Type.Missing, Type.Missing, Type.Missing); //открыть файл
        ////////    Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //получить 1 лист

        ////////    var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell); //1 ячейку
        ////////    string[] str = new string[lastCell.Row]; // массив значений с листа равен по размеру листу

        ////////    /*
        ////////    ObjWorkSheet.Cells[lastCell.Row , 1].Value = "Действует, но с ичастичными зменениями";
        ////////    ObjWorkSheet.Cells[lastCell.Row , 1].Font.Color = Excel.XlRgbColor.rgbGreen;

        ////////    ObjWorkSheet.Cells[lastCell.Row , 2].Value = "Незначительные ошибки, проблемы";
        ////////    ObjWorkSheet.Cells[lastCell.Row , 2].Font.Color = Excel.XlRgbColor.rgbOrange;

        ////////    ObjWorkSheet.Cells[lastCell.Row , 3].Value = "Отменен";
        ////////    ObjWorkSheet.Cells[lastCell.Row , 3].Font.Color = Excel.XlRgbColor.rgbRed;
        ////////    */

        ////////    for (int j = 0; j < lastCell.Row; j++) // по всем строкам
        ////////    {
        ////////        for (int currentcol = 7; currentcol <= 8; currentcol++)
        ////////        {
        ////////            if (ObjWorkSheet.Cells[j + 1, currentcol].Text.ToString() != "")
        ////////            {
        ////////                str[j] = ObjWorkSheet.Cells[j + 1, currentcol].Text.ToString();//считываем текст в строку

        ////////                //вытаскиваем регулярками Действующий ГОСТ
        ////////                string input = str[j];
        ////////                string pattern = @"(\bДействует\b)";
        ////////                Regex regex = new Regex(pattern);

        ////////                // Получаем совпадения в экземпляре класса Match
        ////////                Match match = regex.Match(input);

        ////////                if (match.Value == "Действует")
        ////////                {
        ////////                    string[] str2 = new string[lastCell.Row];
        ////////                    str2[j] = ObjWorkSheet.Cells[j + 1, currentcol-2].Text.ToString();//считываем текст в строку

        ////////                    //вытаскиваем регулярками ГОСТ
        ////////                    string input2 = str2[j];
        ////////                    string pattern2 = @"((\bГОСТ\b)(\D)*([^А-я,(])*)";
        ////////                    Regex regex2 = new Regex(pattern2);

        ////////                    // Получаем совпадения в экземпляре класса Match
        ////////                    Match match2 = regex2.Match(input2);
        ////////                    // отображаем все совпадения
        ////////                    //MessageBox.Show("" + match.Value);
        ////////                    //записываем в textBox

        ////////                    //if (match2.Value != null)
        ////////                    //    textBox1.Text = match2.Value;

        ////////                    //MessageBox.Show(ObjWorkSheet.Cells[j + 1, currentcol].Text.ToString() + " || " + match2.Value + " || yeap!");

        ////////                    //поиск в другой таблице
        ////////                    for (int i = 0; i < dataGridView2.RowCount; i++)
        ////////                    {
        ////////                        //MessageBox.Show(dataGridView2[3, i].FormattedValue.ToString() + " || " + match2.Value.Trim());
        ////////                        //MessageBox.Show(match2.Value.Trim() + " заменен " + Convert.ToString(dataGridView2[3, i].Value));
        ////////                        if (dataGridView2[2, i].FormattedValue.ToString().Contains(match2.Value.Trim()))
        ////////                        {
        ////////                            upd = Convert.ToString(dataGridView2[2, i].Value) + " заменен " + Convert.ToString(dataGridView2[3, i].Value);
        ////////                            ObjWorkSheet.Cells[j + 1, currentcol].Value = upd;
        ////////                            ObjWorkSheet.Cells[j + 1, currentcol].Font.Color = Excel.XlRgbColor.rgbGreen;
        ////////                            //MessageBox.Show(Convert.ToString(dataGridView2[2, i].Value) + " заменен " + Convert.ToString(dataGridView2[3, i].Value));
        ////////                            //return;
        ////////                        }
        ////////                    }
        ////////                }
        ////////            }
        ////////        }
        ////////        progressBar1.Maximum = lastCell.Row;
        ////////        progressBar1.Value = j;
        ////////    }

        ////////    progressBar1.Visible = false;
        ////////    label4.Visible = false;
        ////////    ObjWorkExcel.DisplayAlerts = false;
        ////////    ObjWorkSheet.SaveAs(fileName);
        ////////    ObjWorkExcel.Interactive = true;
        ////////    ObjWorkExcel.ScreenUpdating = true;
        ////////    ObjWorkExcel.UserControl = true;


        ////////    dataGridView2.Dispose();
        ////////    GC.Collect();
        ////////    GC.WaitForPendingFinalizers();
        ////////    GC.Collect();
        ////////}
    }
}


    //progressBar1.Value = 0;
    //progressBar1.Visible = true;
    //progressBar1.Maximum = lastCell.Row;
    //progressBar1.Value = j;
    //progressBar1.Visible = false;