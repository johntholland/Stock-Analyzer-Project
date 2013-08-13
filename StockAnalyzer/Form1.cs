using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using HtmlAgilityPack;
//using Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace StockAnalyzer
{
    public partial class Form1 : Form
    {
        //private Microsoft.Office.Interop.Excel.Application excel;
        private bool containsErrors = false;
        static List<string> SymbolsList;

        public Form1()
        {
            InitializeComponent();
            //excel = new Microsoft.Office.Interop.Excel.Application();

            //if (excel == null)
            //{
            //    MessageBox.Show("EXCEL could not be started. Check that your office installation and project references are correct.");
            //    return;
            //}
        }

        public static void parseSymbols()
        {
            try
            {
                SymbolsList = new List<string>();
                using (StreamReader sr = new StreamReader("symbols.txt"))
                {

                    string line;
                    while( (line = sr.ReadLine()) != null)
                    {
                    String[] templist = line.Split('\t');
                    SymbolsList.Add(templist[0]);
                    }

                }
            }
            catch (Exception e)
            {
                Console.WriteLine("The file could not be read:");
                Console.WriteLine(e.Message);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //createTextListing();

            toolStripStatusLabel1.Visible = false;
            String currentPath = Directory.GetCurrentDirectory();

            parseSymbols();
            foreach(string str in SymbolsList)
            {
                Console.WriteLine(str);
            }
            Console.WriteLine("");


            /*

            OpenFileDialog op = new OpenFileDialog();
            op.InitialDirectory = currentPath;
            if (op.ShowDialog() == DialogResult.OK)
                currentPath = op.FileName;
            else
            {
                toolStripStatusLabel1.Text = "Failed to Load Workbook";
                toolStripStatusLabel1.Visible = true;
                return;
            }

            //Workbook wb = excel.Workbooks.Open(currentPath);

            //Worksheet ws = (Worksheet)wb.Worksheets[1];
            //ws.Visible = XlSheetVisibility.xlSheetVisible;

            if (ws == null)
            {
                Console.WriteLine("Worksheet could not be created. Check that your office installation and project references are correct.");
                return;
            }

            String urlbase = "https://www.google.com/finance?q=";


            //if (NasdaqRadioButton.Checked)
            //    urlbase = "http://www.google.com/finance?q=NASDAQ:";
            //else
            //    urlbase = "http://www.google.com/finance?q=NYSE:";
            string url = urlbase + textBox1.Text + "&fstype=ii";
            string currentPriceUrl = urlbase + textBox1.Text;

            var webGet = new HtmlWeb();
            var document = webGet.Load(currentPriceUrl);

            //get current price
            var price = document.DocumentNode.SelectSingleNode("//*[@class='pr']/span[1]");
            this.excel.ActiveSheet.Cells[48, 8] = price.InnerText.Trim();

            //paste desired return %
            this.excel.ActiveSheet.Cells[31, 5] = 10;

            var bankpage = webGet.Load(urlbase);
            var banktest = bankpage.DocumentNode.SelectSingleNode("//*[@id='appbar']/div/div[2]/div[1]/span");

            if (banktest.InnerHtml.Contains("bank"))
            {
                //if Google summary page: //*[@id="related"]/div[4]/div/div[1]/a[2] includes "bank" then
                //go to http://www.reuters.com/finance/stocks/incomeStatement?perType=INT&symbol=BAC ACC

                //sum of Non-Interest Income, Bank and Interest Income, Bank. instead of Total Revenue
                var bankfinancials = webGet.Load("");
            }

            else
            {
                document = webGet.Load(url);

                //get Company's full name
                var htest = document.DocumentNode.SelectSingleNode("//*[@id='appbar']/div/div[2]/div[1]/span");
                if (htest != null)
                    this.excel.ActiveSheet.Cells[3, 3] = htest.InnerText.Trim();
                else
                {
                    errorOccured();
                }

                //paste provided stock symbol 
                this.excel.ActiveSheet.Cells[4, 3] = textBox1.Text.ToUpper();

                //get multiplier
                var multiplier = document.DocumentNode.SelectSingleNode("//*[@id='fs-table']/thead/tr/th[1]");
                if (multiplier != null)
                    this.excel.ActiveSheet.Cells[6, 7] = multiplier.InnerText.Trim();
                else
                {
                    errorOccured();
                }

                //get dates Total Revenue & Net Income (annual)
                var datesAnnual = document.DocumentNode.SelectSingleNode("//div[@id='incannualdiv']/table[@id='fs-table']/thead/tr/th[2]");
                this.excel.ActiveSheet.Cells[8, 4] = datesAnnual.InnerText.Trim();
                var datesAnnual1 = document.DocumentNode.SelectSingleNode("//div[@id='incannualdiv']/table[@id='fs-table']/thead/tr/th[3]");
                this.excel.ActiveSheet.Cells[8, 5] = datesAnnual1.InnerText.Trim();
                var datesAnnual2 = document.DocumentNode.SelectSingleNode("//div[@id='incannualdiv']/table[@id='fs-table']/thead/tr/th[4]");
                this.excel.ActiveSheet.Cells[8, 6] = datesAnnual2.InnerText.Trim();
                var datesAnnual3 = document.DocumentNode.SelectSingleNode("//div[@id='incannualdiv']/table[@id='fs-table']/thead/tr/th[5]");
                this.excel.ActiveSheet.Cells[8, 7] = datesAnnual3.InnerText.Trim();

                //get Total Revenue (annual)
                var totalRevenueAnnual = document.DocumentNode.SelectSingleNode("//div[@id='incannualdiv']/table[@id='fs-table']/tbody/tr[3]/td[2]");
                this.excel.ActiveSheet.Cells[9, 4] = totalRevenueAnnual.InnerText.Trim();
                var totalRevenueAnnual1 = document.DocumentNode.SelectSingleNode("//div[@id='incannualdiv']/table[@id='fs-table']/tbody/tr[3]/td[3]");
                this.excel.ActiveSheet.Cells[9, 5] = totalRevenueAnnual1.InnerText.Trim();
                var totalRevenueAnnual2 = document.DocumentNode.SelectSingleNode("//div[@id='incannualdiv']/table[@id='fs-table']/tbody/tr[3]/td[4]");
                this.excel.ActiveSheet.Cells[9, 6] = totalRevenueAnnual2.InnerText.Trim();
                var totalRevenueAnnual3 = document.DocumentNode.SelectSingleNode("//div[@id='incannualdiv']/table[@id='fs-table']/tbody/tr[3]/td[5]");
                this.excel.ActiveSheet.Cells[9, 7] = totalRevenueAnnual3.InnerText.Trim();

                //*[@id="fs-table"]/tbody/tr[28]/td[2]

                //get Net Income 
                var netIncomeAnnual = document.DocumentNode.SelectSingleNode("//div[@id='incannualdiv']/table[@id='fs-table']/tbody/tr[28]/td[2]");
                this.excel.ActiveSheet.Cells[10, 4] = netIncomeAnnual.InnerText.Trim();
                var netIncomeAnnual1 = document.DocumentNode.SelectSingleNode("//div[@id='incannualdiv']/table[@id='fs-table']/tbody/tr[28]/td[3]");
                this.excel.ActiveSheet.Cells[10, 5] = netIncomeAnnual1.InnerText.Trim();
                var netIncomeAnnual2 = document.DocumentNode.SelectSingleNode("//div[@id='incannualdiv']/table[@id='fs-table']/tbody/tr[28]/td[4]");
                this.excel.ActiveSheet.Cells[10, 6] = netIncomeAnnual2.InnerText.Trim();
                var netIncomeAnnual3 = document.DocumentNode.SelectSingleNode("//div[@id='incannualdiv']/table[@id='fs-table']/tbody/tr[28]/td[5]");
                this.excel.ActiveSheet.Cells[10, 7] = netIncomeAnnual3.InnerText.Trim();

                //get dates Total Revenue & Net Income (Quarterly)
                var datesQuarterly = document.DocumentNode.SelectSingleNode("//div[@id='incinterimdiv']/table[@id='fs-table']/thead/tr/th[2]");
                this.excel.ActiveSheet.Cells[20, 6] = datesQuarterly.InnerText.Trim();
                var datesQuarterly1 = document.DocumentNode.SelectSingleNode("//div[@id='incinterimdiv']/table[@id='fs-table']/thead/tr/th[3]");
                this.excel.ActiveSheet.Cells[20, 7] = datesQuarterly1.InnerText.Trim();
                var datesQuarterly2 = document.DocumentNode.SelectSingleNode("//div[@id='incinterimdiv']/table[@id='fs-table']/thead/tr/th[4]");
                this.excel.ActiveSheet.Cells[20, 8] = datesQuarterly2.InnerText.Trim();
                var datesQuarterly3 = document.DocumentNode.SelectSingleNode("//div[@id='incinterimdiv']/table[@id='fs-table']/thead/tr/th[5]");
                this.excel.ActiveSheet.Cells[20, 9] = datesQuarterly3.InnerText.Trim();

                //get Total Revenue (Quarterly)
                var totalRevenueQuarterly = document.DocumentNode.SelectSingleNode("//div[@id='incinterimdiv']/table[@id='fs-table']/tbody/tr[3]/td[2]");
                this.excel.ActiveSheet.Cells[21, 6] = totalRevenueQuarterly.InnerText.Trim();
                var totalRevenueQuarterly1 = document.DocumentNode.SelectSingleNode("//div[@id='incinterimdiv']/table[@id='fs-table']/tbody/tr[3]/td[3]");
                this.excel.ActiveSheet.Cells[21, 7] = totalRevenueQuarterly1.InnerText.Trim();
                var totalRevenueQuarterly2 = document.DocumentNode.SelectSingleNode("//div[@id='incinterimdiv']/table[@id='fs-table']/tbody/tr[3]/td[4]");
                this.excel.ActiveSheet.Cells[21, 8] = totalRevenueQuarterly2.InnerText.Trim();
                var totalRevenueQuarterly3 = document.DocumentNode.SelectSingleNode("//div[@id='incinterimdiv']/table[@id='fs-table']/tbody/tr[3]/td[5]");
                this.excel.ActiveSheet.Cells[21, 9] = totalRevenueQuarterly3.InnerText.Trim();

                //get Net Income (Quarterly)

                var netIncomeQuarterly = document.DocumentNode.SelectSingleNode("//div[@id='incinterimdiv']/table[@id='fs-table']/tbody/tr[28]/td[2]");
                this.excel.ActiveSheet.Cells[22, 6] = netIncomeQuarterly.InnerText.Trim();
                var netIncomeQuarterly1 = document.DocumentNode.SelectSingleNode("//div[@id='incinterimdiv']/table[@id='fs-table']/tbody/tr[28]/td[3]");
                this.excel.ActiveSheet.Cells[22, 7] = netIncomeQuarterly1.InnerText.Trim();
                var netIncomeQuarterly2 = document.DocumentNode.SelectSingleNode("//div[@id='incinterimdiv']/table[@id='fs-table']/tbody/tr[28]/td[4]");
                this.excel.ActiveSheet.Cells[22, 8] = netIncomeQuarterly2.InnerText.Trim();
                var netIncomeQuarterly3 = document.DocumentNode.SelectSingleNode("//div[@id='incinterimdiv']/table[@id='fs-table']/tbody/tr[28]/td[5]");
                this.excel.ActiveSheet.Cells[22, 9] = netIncomeQuarterly3.InnerText.Trim();

                //get debts Date current as of
                var currentLiabilitiesDateQuarterly = document.DocumentNode.SelectSingleNode("//div[@id='balinterimdiv']/table[@id='fs-table']/thead/tr/th[2]");
                this.excel.ActiveSheet.Cells[37, 7] = currentLiabilitiesDateQuarterly.InnerText.Trim();
                this.excel.ActiveSheet.Cells[37, 8] = currentLiabilitiesDateQuarterly.InnerText.Trim();

                //get current Debts
                var currentLiabilitiesQuarterly = document.DocumentNode.SelectSingleNode("//div[@id='balinterimdiv']/table[@id='fs-table']/tbody/tr[26]/td[2]");
                this.excel.ActiveSheet.Cells[39, 7] = currentLiabilitiesQuarterly.InnerText.Trim();
                var currentOtherLiabilitiesQuarterly = document.DocumentNode.SelectSingleNode("//div[@id='balinterimdiv']/table[@id='fs-table']/tbody/tr[30]/td[2]");
                this.excel.ActiveSheet.Cells[39, 8] = currentOtherLiabilitiesQuarterly.InnerText.Trim();

                //get current outstanding shares
                //*[@id="fs-table"]/tbody/tr[42]/td[2]
                var currentOutstandingShares = document.DocumentNode.SelectSingleNode("//div[@id='balinterimdiv']/table[@id='fs-table']/tbody/tr[42]/td[2]");
                this.excel.ActiveSheet.Cells[47, 5] = currentOutstandingShares.InnerText.Trim();


                excel.Visible = true;
                try
                {
                    wb.SaveAs(textBox1.Text, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                }
                catch
                {
                    toolStripStatusLabel1.Text = "The workbook was unable to be saved. do you have a workbook by the same name open already?";
                    toolStripStatusLabel1.Visible = true;
                }
            }
             */
        }
        //private void createTextListing()
        //{
        //    //*[@id="ctl00_cph1_divSymbols"]/table/tbody/tr[2]/td[1]/a
        //    //*[@id="ctl00_cph1_divSymbols"]/table/tbody/tr[3]/td[1]/a
            
        //    //go and scrape stock symbols from http://www.eoddata.com/symbols.aspx

        //    string url = "http://www.eoddata.com/symbols.aspx";

        //    var webGet = new HtmlWeb();
        //    var document = webGet.Load(url);

        //    //get all the symbols
        //    var symbol = document.DocumentNode.SelectSingleNode("//*[@id=\"ctl00_cph1_divSymbols\"]/table/tbody/tr[3]/td[1]");

        //    MessageBox.Show(symbol.InnerText.Trim());

        //    document = webGet.Load(url);
        //    throw new NotImplementedException();
        //}

        private void errorOccured()
        {
            toolStripStatusLabel1.Text = "The Worksheet for " + textBox1.Text.ToUpper() + " contains Errors. Please report this error";
            toolStripStatusLabel1.Visible = true;
            containsErrors = true;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //  excel.Quit();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
