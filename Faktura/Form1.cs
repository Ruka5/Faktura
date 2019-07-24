using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;
using System.Collections.Specialized;
using MySql.Data.MySqlClient;
using System.Configuration;
using Microsoft.Office.Interop.Excel; 





namespace Faktura
{
    public partial class Form1 : Form
    {
        public string DB_nagios = ConfigurationManager.AppSettings["DB_nagios"];
        public string database_nagios = ConfigurationManager.AppSettings["database_nagios"];
        public string uid_nagios = ConfigurationManager.AppSettings["uid_nagios"];
        public string pwd_nagios = ConfigurationManager.AppSettings["pwd_nagios"];

        public string path = ConfigurationManager.AppSettings["path"];
        public static String STR;
        
        BackgroundWorker bwDBalive;
        public bool DB_alive;

        
        
        
        public Form1()
        {
            InitializeComponent();
            STR = "Server=" + DB_nagios + ";Database=" + database_nagios + ";Uid=" + uid_nagios + ";Pwd=" + pwd_nagios + ";Connect Timeout=10;Pooling=False;";
            bwDBalive = new BackgroundWorker();
            bwDBalive.DoWork += new DoWorkEventHandler(bwDBalive_DoWork);
        }

        void bwDBalive_DoWork(object sender, DoWorkEventArgs e)
        {
            using (MySqlConnection conn = new MySqlConnection(STR))
            {
                try
                {
                    conn.Open();

                    DB_alive = true; // set online

                    this.Invoke(
                      (MethodInvoker)delegate()
                      {
                          label1.BackColor = Color.Lime;
                      }
                     );
                }
                catch
                {
                    DB_alive = false; // set offline

                    this.Invoke(
                      (MethodInvoker)delegate()
                      {
                          label1.BackColor = Color.Red;
                       }
                     );
                }
                finally
                {
                    conn.Close();
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            try
            {
                if (bwDBalive.IsBusy == false)
                    bwDBalive.RunWorkerAsync();

                if (DB_alive == true)
                    label2.Text = "Connections Live";

                else if (DB_alive == false)
                    label2.Text = "Connections Not available";
            }

            catch (System.Exception ex)
            {
                StreamWriter log;
                if (!File.Exists("ErrorLog.txt"))
                {
                    log = new StreamWriter("ErrorLog.txt");
                }
                else
                {
                    log = File.AppendText("ErrorLog.txt");
                }

                log.WriteLine(DateTime.Now);
                log.WriteLine(ex.Message);
                log.WriteLine();
                log.Close();

            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            /*
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            Excel.Range chartRange;
            */
            //xlApp = new Excel.ApplicationClass();
           // xlWorkBook = xlApp.Workbooks.Add(misValue);
           //lWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            
            
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
            worksheet = workbook.ActiveSheet;
            Microsoft.Office.Interop.Excel.Range chartRange;
            

            worksheet.Name = "Faktura";
            app.Visible = true;
            worksheet.PageSetup.Zoom = 80;

            worksheet.PageSetup.BottomMargin = 0.2;
            worksheet.PageSetup.LeftMargin = 0.2;
            worksheet.PageSetup.RightMargin = 0.2;
            worksheet.PageSetup.HeaderMargin = 0.2;
            worksheet.PageSetup.FooterMargin = 0.2;
            worksheet.PageSetup.Orientation = XlPageOrientation.xlPortrait;

            //ramecek danovy doklad
            worksheet.get_Range("A1", "H1").Merge(false); //slucuje bunky
            chartRange = app.get_Range("A1", "H1");
            chartRange.FormulaR1C1 = "DAŇOVÝ DOKLAD - FAKTURA";
            chartRange.RowHeight = 24.75;
            worksheet.Range["A1:H1"].Style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            worksheet.Range["A1:H1"].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            
            worksheet.get_Range("A1", "A1").Cells.Font.Size = 12;
            worksheet.get_Range("A1", "A1").Cells.Font.Name = "Arial";
            worksheet.get_Range("A1", "A1").Cells.Font.Bold = true;

            //konec
            
            //ramecek cislo fakuty
            worksheet.get_Range("L1", "M1").Merge(false);
            chartRange = app.get_Range("L1", "M1");
            chartRange.FormulaR1C1 = "č. 05513";
            worksheet.Range["L1:M1"].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            worksheet.get_Range("L1", "L1").Cells.Font.Size = 12;
            worksheet.get_Range("L1", "L1").Cells.Font.Name = "Arial";
            worksheet.get_Range("L1", "L1").Cells.Font.Bold = true;
            //konec

            //ohraniceni zahlavi
            worksheet.get_Range("A1", "M1").BorderAround(Type.Missing, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Type.Missing);
            //konec

            //maly radek mezi
            chartRange = app.get_Range("A2", "M2");
            chartRange.RowHeight = 6.75;
            //konec

            //ramecek dodavatel
            worksheet.get_Range("A3", "B3").Merge(false); //slucuje bunky
            chartRange = app.get_Range("A3", "B3");
            chartRange.FormulaR1C1 = "Dodavatel:";

            worksheet.get_Range("A3", "A3").Cells.Font.Size = 12;
            worksheet.get_Range("A3", "A3").Cells.Font.Name = "Arial";
            worksheet.get_Range("A3", "A3").Cells.Font.Bold = true;
            worksheet.get_Range("A3", "A3").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            worksheet.get_Range("B5", "E5").Merge(false); //NAZEV FIRMY
            chartRange = app.get_Range("B5", "E5");
            chartRange.FormulaR1C1 = "KOVO-FALCH, s.r.o";
            worksheet.get_Range("B5", "B5").Cells.Font.Size = 12;
            worksheet.get_Range("B5", "B5").Cells.Font.Name = "Arial";
            worksheet.get_Range("B5", "B5").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            worksheet.get_Range("B6", "E6").Merge(false); //CISLO ULICE
            chartRange = app.get_Range("B6", "E6");
            chartRange.FormulaR1C1 = "Kratka 416";
            worksheet.get_Range("B6", "B6").Cells.Font.Size = 12;
            worksheet.get_Range("B6", "B6").Cells.Font.Name = "Arial";
            worksheet.get_Range("B6", "B6").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            worksheet.get_Range("B7", "E7").Merge(false); //SMEROVACKA A MESTO
            chartRange = app.get_Range("B7", "E7");
            chartRange.FormulaR1C1 = "739 25 Sviadnov";
            worksheet.get_Range("B7", "B7").Cells.Font.Size = 12;
            worksheet.get_Range("B7", "B7").Cells.Font.Name = "Arial";
            worksheet.get_Range("B7", "B7").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            worksheet.get_Range("F4", "H4").Merge(false); //ICO
            chartRange = app.get_Range("F4", "H4");
            chartRange.FormulaR1C1 = "IČ: 25858173";
            worksheet.get_Range("F4", "F4").Cells.Font.Size = 12;
            worksheet.get_Range("F4", "F4").Cells.Font.Name = "Arial";
            worksheet.get_Range("F4", "F4").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            worksheet.get_Range("F5", "H5").Merge(false); //DIC
            chartRange = app.get_Range("F5", "H5");
            chartRange.FormulaR1C1 = "DIČ: CZ25858173";
            worksheet.get_Range("F5", "F5").Cells.Font.Size = 12;
            worksheet.get_Range("F5", "F5").Cells.Font.Name = "Arial";
            worksheet.get_Range("F5", "F5").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            worksheet.get_Range("A9", "B9").Merge(false); //PENEZNI USTAV
            chartRange = app.get_Range("A9", "B9");
            chartRange.FormulaR1C1 = "Peněžní ústav:";
            worksheet.get_Range("A9", "A9").Cells.Font.Size = 12;
            worksheet.get_Range("A9", "A9").Cells.Font.Name = "Arial";
            worksheet.get_Range("A9", "A9").Cells.Font.Bold = true;
            worksheet.get_Range("A9", "A9").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            worksheet.get_Range("A10", "C10").Merge(false); //PENEZNI USTAV
            chartRange = app.get_Range("A10", "C10");
            chartRange.FormulaR1C1 = "KB Frýdek-Místek";
            worksheet.get_Range("A10", "A10").Cells.Font.Size = 12;
            worksheet.get_Range("A10", "A10").Cells.Font.Name = "Arial";
            worksheet.get_Range("A10", "A10").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            worksheet.get_Range("D9", "E9").Merge(false); //CISLO UCTU
            chartRange = app.get_Range("D9", "E9");
            chartRange.FormulaR1C1 = "číslo účtu";
            worksheet.get_Range("D9", "D9").Cells.Font.Size = 12;
            worksheet.get_Range("D9", "D9").Cells.Font.Name = "Arial";
            worksheet.get_Range("D9", "D9").Cells.Font.Bold = true;
            worksheet.get_Range("D9", "D9").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            worksheet.get_Range("D10", "E10").Merge(false);
            chartRange = app.get_Range("D10", "E10");
            chartRange.FormulaR1C1 = "123456789123";
            worksheet.get_Range("D10", "D10").Cells.Font.Size = 12;
            worksheet.get_Range("D10", "D10").Cells.Font.Name = "Arial";
            worksheet.get_Range("D10", "D10").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            worksheet.get_Range("D10", "D10").NumberFormat = "0";

            worksheet.get_Range("F9", "F9").Merge(false); //KOD
            chartRange = app.get_Range("F9", "F9");
            chartRange.FormulaR1C1 = "kod";
            worksheet.get_Range("F9", "F9").Cells.Font.Size = 12;
            worksheet.get_Range("F9", "F9").Cells.Font.Name = "Arial";
            worksheet.get_Range("F9", "F9").Cells.Font.Bold = true;
            worksheet.get_Range("F9", "F9").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            worksheet.get_Range("F10", "F10").Merge(false);
            chartRange = app.get_Range("F10", "F10");
            chartRange.FormulaR1C1 = "/100";
            worksheet.get_Range("F10", "F10").Cells.Font.Size = 12;
            worksheet.get_Range("F10", "F10").Cells.Font.Name = "Arial";
            worksheet.get_Range("F10", "F10").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            worksheet.get_Range("A3", "H10").BorderAround(Type.Missing, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Type.Missing);

            worksheet.get_Range("I3", "K3").Merge(false); //CISLO OBJEDNAVKY
            chartRange = app.get_Range("I3", "K3");
            chartRange.FormulaR1C1 = "Objednávka číslo:";
            worksheet.get_Range("I3", "I3").Cells.Font.Size = 12;
            worksheet.get_Range("I3", "I3").Cells.Font.Name = "Arial";
            worksheet.get_Range("I3", "I3").Cells.Font.Bold = true;
            worksheet.get_Range("I3", "I3").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            worksheet.get_Range("I4", "L4").Merge(false);
            chartRange = app.get_Range("I4", "L4");
            chartRange.FormulaR1C1 = "4400001148/N29/854";
            worksheet.get_Range("I4", "I4").Cells.Font.Size = 12;
            worksheet.get_Range("I4", "I4").Cells.Font.Name = "Arial";
            worksheet.get_Range("I4", "I4").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            worksheet.get_Range("I7", "K7").Merge(false); //ODBERATEL ICO DICO
            chartRange = app.get_Range("I7", "K7");
            chartRange.FormulaR1C1 = "Odběratel:";
            worksheet.get_Range("I7", "I7").Cells.Font.Size = 12;
            worksheet.get_Range("I7", "I7").Cells.Font.Name = "Arial";
            worksheet.get_Range("I7", "I7").Cells.Font.Bold = true;
            worksheet.get_Range("I7", "I7").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            worksheet.get_Range("I8", "K8").Merge(false); //ODBERATEL ICO DICO
            chartRange = app.get_Range("I8", "K8");
            chartRange.FormulaR1C1 = "IČ: 25858173";
            worksheet.get_Range("I8", "I8").Cells.Font.Size = 12;
            worksheet.get_Range("I8", "I8").Cells.Font.Name = "Arial";
            worksheet.get_Range("I8", "I8").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            worksheet.get_Range("I9", "K9").Merge(false); //ODBERATEL ICO DICO
            chartRange = app.get_Range("I9", "K9");
            chartRange.FormulaR1C1 = "DIČ: CZ25858173";
            worksheet.get_Range("I9", "I9").Cells.Font.Size = 12;
            worksheet.get_Range("I9", "I9").Cells.Font.Name = "Arial";
            worksheet.get_Range("I9", "I9").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            worksheet.get_Range("I3", "M10").BorderAround(Type.Missing, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Type.Missing);

            //maly radek mezi
            chartRange = app.get_Range("A11", "M11");
            chartRange.RowHeight = 6.75;
            //konec

            worksheet.get_Range("A12", "C12").Merge(false); //ADRESA PRIJEMCE, ZPUSOB DOPRAVY, CISLO DD
            chartRange = app.get_Range("A12", "C12");
            chartRange.FormulaR1C1 = "Adresa příjemce:";
            worksheet.get_Range("A12", "A12").Cells.Font.Size = 12;
            worksheet.get_Range("A12", "A12").Cells.Font.Name = "Arial";
            worksheet.get_Range("A12", "A12").Cells.Font.Bold = true;
            worksheet.get_Range("A12", "A12").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            worksheet.get_Range("A13", "D13").Merge(false); //ADRESA PRIJEMCE, - PRVNI RADEK PAK JESTE DVA
            chartRange = app.get_Range("A13", "D13");
            chartRange.FormulaR1C1 = "skl. 854";
            worksheet.get_Range("A13", "A13").Cells.Font.Size = 12;
            worksheet.get_Range("A13", "A13").Cells.Font.Name = "Arial";
            worksheet.get_Range("A13", "A13").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            worksheet.get_Range("A14", "D14").Merge(false); //ADRESA PRIJEMCE - DRUHY PRAZDNY
            chartRange = app.get_Range("A14", "D14");
            chartRange.FormulaR1C1 = "";
            worksheet.get_Range("A14", "A14").Cells.Font.Size = 12;
            worksheet.get_Range("A14", "A14").Cells.Font.Name = "Arial";
            worksheet.get_Range("A14", "A14").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            worksheet.get_Range("A15", "D15").Merge(false); //ADRESA PRIJEMCE - TRETI PRAZDNY
            chartRange = app.get_Range("A15", "D15");
            chartRange.FormulaR1C1 = "";
            worksheet.get_Range("A15", "A15").Cells.Font.Size = 12;
            worksheet.get_Range("A15", "A15").Cells.Font.Name = "Arial";
            worksheet.get_Range("A15", "A15").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            worksheet.get_Range("A16", "C16").Merge(false); //ZPUSOB DOPRAVY
            chartRange = app.get_Range("A16", "C16");
            chartRange.FormulaR1C1 = "Způsob dopravy:";
            worksheet.get_Range("A16", "A16").Cells.Font.Size = 12;
            worksheet.get_Range("A16", "A16").Cells.Font.Name = "Arial";
            worksheet.get_Range("A16", "A16").Cells.Font.Bold = true;
            worksheet.get_Range("A16", "A16").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            worksheet.get_Range("A17", "C17").Merge(false);
            chartRange = app.get_Range("A17", "C17");
            chartRange.FormulaR1C1 = "";
            worksheet.get_Range("A17", "A17").Cells.Font.Size = 12;
            worksheet.get_Range("A17", "A17").Cells.Font.Name = "Arial";
            worksheet.get_Range("A17", "A17").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            worksheet.get_Range("A19", "B19").Merge(false); //CISLO DODACIHO LISTU
            chartRange = app.get_Range("A19", "B19");
            chartRange.FormulaR1C1 = "Dodací list číslo:";
            worksheet.get_Range("A19", "A19").Cells.Font.Size = 12;
            worksheet.get_Range("A19", "A19").Cells.Font.Name = "Arial";
            worksheet.get_Range("A19", "A19").Cells.Font.Bold = true;
            worksheet.get_Range("A19", "A19").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            worksheet.get_Range("C19", "D19").Merge(false);
            chartRange = app.get_Range("C19", "D19");
            chartRange.FormulaR1C1 = "05413";
            worksheet.get_Range("C19", "C19").Cells.Font.Size = 12;
            worksheet.get_Range("C19", "C19").Cells.Font.Name = "Arial";
            worksheet.get_Range("C19", "C19").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            worksheet.get_Range("A12", "E19").BorderAround(Type.Missing, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Type.Missing);

            //ADRESA ODBERATELE
            worksheet.get_Range("F13", "G13").Merge(false); //ADRESA ODBERATELE
            chartRange = app.get_Range("F13", "G13");
            chartRange.FormulaR1C1 = "Odběratel:";
            worksheet.get_Range("F13", "F13").Cells.Font.Size = 12;
            worksheet.get_Range("F13", "F13").Cells.Font.Name = "Arial";
            worksheet.get_Range("F13", "F13").Cells.Font.Bold = true;
            worksheet.get_Range("F13", "F13").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            worksheet.get_Range("G14", "L14").Merge(false); //NAZEV FIRMY ODBERATELE
            chartRange = app.get_Range("G14", "L14");
            chartRange.FormulaR1C1 = "ArcelorMIttal Tubular Products Ostrava a.s.";
            worksheet.get_Range("G14", "G14").Cells.Font.Size = 12;
            worksheet.get_Range("G14", "G14").Cells.Font.Name = "Arial";
            worksheet.get_Range("G14", "G14").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            worksheet.get_Range("G15", "L15").Merge(false);//ULICE ODBERATELE
            chartRange = app.get_Range("G15", "L15");
            chartRange.FormulaR1C1 = "Vratimovská 689";
            worksheet.get_Range("G15", "G15").Cells.Font.Size = 12;
            worksheet.get_Range("G15", "G15").Cells.Font.Name = "Arial";
            worksheet.get_Range("G15", "G15").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            worksheet.get_Range("G16", "L16").Merge(false); //MESTO ODBERATELE
            chartRange = app.get_Range("G16", "L16");
            chartRange.FormulaR1C1 = "Ostrava-Kunčice";
            worksheet.get_Range("G16", "G16").Cells.Font.Size = 12;
            worksheet.get_Range("G16", "G16").Cells.Font.Name = "Arial";
            worksheet.get_Range("G16", "G16").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            worksheet.get_Range("G17", "L17").Merge(false); //PSC ODBERATELE
            chartRange = app.get_Range("G17", "L17");
            chartRange.FormulaR1C1 = "707 02";
            worksheet.get_Range("G17", "G17").Cells.Font.Size = 12;
            worksheet.get_Range("G17", "G17").Cells.Font.Name = "Arial";
            worksheet.get_Range("G17", "G17").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;      

            worksheet.get_Range("F12", "M19").BorderAround(Type.Missing, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Type.Missing);
            //konec

            //maly radek mezi
            chartRange = app.get_Range("A20", "M20");
            chartRange.RowHeight = 6.75;
            //konec

            //DATUMY
            worksheet.get_Range("A21", "E21").Merge(false); //DATUM USKUTECNENI ZDANITELNEHO PLNENI
            chartRange = app.get_Range("A21", "E21");
            chartRange.FormulaR1C1 = "Datum uskutečnění zdanitelného plnění:";
            worksheet.get_Range("A21", "A21").Cells.Font.Size = 12;
            worksheet.get_Range("A21", "A21").Cells.Font.Name = "Arial";
            worksheet.get_Range("A21", "A21").Cells.Font.Bold = true;
            worksheet.get_Range("A21", "A21").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            worksheet.Range["A21", "E21"].Style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

            worksheet.get_Range("A22", "E22").Merge(false); //DATUM USKUTECNENI ZDANITELNEHO PLNENI
            chartRange = app.get_Range("A22", "E22");
            chartRange.FormulaR1C1 = "22.1.2013";
            worksheet.get_Range("A22", "A22").Cells.Font.Size = 12;
            worksheet.get_Range("A22", "A22").Cells.Font.Name = "Arial";
            worksheet.get_Range("A22", "A22").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            worksheet.get_Range("F21", "J21").Merge(false); //DATUM VYSTAVENI DD
            chartRange = app.get_Range("F21", "J21");
            chartRange.FormulaR1C1 = "Datum vystavení daňového dokladu:";
            worksheet.get_Range("F21", "F21").Cells.Font.Size = 12;
            worksheet.get_Range("F21", "F21").Cells.Font.Name = "Arial";
            worksheet.get_Range("F21", "F21").Cells.Font.Bold = true;
            worksheet.get_Range("F21", "F21").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            worksheet.Range["F21", "F21"].Style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

            worksheet.get_Range("F22", "J22").Merge(false); //DATUM VYSTAVENI DD
            chartRange = app.get_Range("F22", "J22");
            chartRange.FormulaR1C1 = "22.1.2013";
            worksheet.get_Range("F22", "F22").Cells.Font.Size = 12;
            worksheet.get_Range("F22", "F22").Cells.Font.Name = "Arial";
            worksheet.get_Range("F22", "F22").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            worksheet.get_Range("K21", "M21").Merge(false); //DATUM SPLATNOSTI
            chartRange = app.get_Range("K21", "M21");
            chartRange.FormulaR1C1 = "Datum splatnosti";
            worksheet.get_Range("K21", "K21").Cells.Font.Size = 12;
            worksheet.get_Range("K21", "K21").Cells.Font.Name = "Arial";
            worksheet.get_Range("K21", "K21").Cells.Font.Bold = true;
            worksheet.get_Range("K21", "K21").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            worksheet.Range["K21", "K21"].Style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

            worksheet.get_Range("K22", "M22").Merge(false); //DATUM SPLATNOSTI
            chartRange = app.get_Range("K22", "M22");
            chartRange.FormulaR1C1 = "22.3.2013";
            worksheet.get_Range("K22", "K22").Cells.Font.Size = 12;
            worksheet.get_Range("K22", "K22").Cells.Font.Name = "Arial";
            worksheet.get_Range("K22", "K22").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            worksheet.get_Range("A21", "E22").BorderAround(Type.Missing, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Type.Missing);
            worksheet.get_Range("F21", "J22").BorderAround(Type.Missing, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Type.Missing);
            worksheet.get_Range("K21", "M22").BorderAround(Type.Missing, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Type.Missing);
            //konec
            
            //maly radek mezi
            chartRange = app.get_Range("A23", "M23");
            chartRange.RowHeight = 6.75;
            //konec

            //POLOZKY
            worksheet.get_Range("A24", "D24").Merge(false); //OZNACENI DODAVKY
            chartRange = app.get_Range("A24", "D24");
            chartRange.FormulaR1C1 = "Označení dodávky";
            worksheet.get_Range("A24", "A24").Cells.Font.Size = 10;
            worksheet.get_Range("A24", "A24").Cells.Font.Name = "Arial";
            worksheet.get_Range("A24", "A24").Cells.Font.Bold = true;
            worksheet.get_Range("A24", "A24").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

            worksheet.get_Range("A25", "D25").Merge(false); //OZNACENI DODAVKYI
            chartRange = app.get_Range("A25", "D25");
            chartRange.FormulaR1C1 = "";
            worksheet.get_Range("A25", "A25").Cells.Font.Size = 10;
            worksheet.get_Range("A25", "A25").Cells.Font.Name = "Arial";
            worksheet.get_Range("A25", "A25").Cells.Font.Bold = true;
            worksheet.get_Range("A25", "A25").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

            worksheet.get_Range("E24", "F24").Merge(false); //JEDNOTKA MNOZSTVI
            chartRange = app.get_Range("E24", "F24");
            chartRange.FormulaR1C1 = "Jednotka množství";
            worksheet.get_Range("E24", "E24").Cells.Font.Size = 10;
            worksheet.get_Range("E24", "E24").Cells.Font.Name = "Arial";
            worksheet.get_Range("E24", "E24").Cells.Font.Bold = true;
            worksheet.get_Range("E24", "E24").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

            worksheet.get_Range("E25", "F25").Merge(false); //JEDNOTKA MNOZSTVI
            chartRange = app.get_Range("E25", "F25");
            chartRange.FormulaR1C1 = "Množství";
            worksheet.get_Range("E25", "E25").Cells.Font.Size = 10;
            worksheet.get_Range("E25", "E25").Cells.Font.Name = "Arial";
            worksheet.get_Range("E25", "E25").Cells.Font.Bold = true;
            worksheet.get_Range("E25", "E25").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

            worksheet.get_Range("G24", "I24").Merge(false); //CENA
            chartRange = app.get_Range("G24", "I24");
            chartRange.FormulaR1C1 = "Cena za MJ bez DPH";
            worksheet.get_Range("G24", "G24").Cells.Font.Size = 10;
            worksheet.get_Range("G24", "G24").Cells.Font.Name = "Arial";
            worksheet.get_Range("G24", "G24").Cells.Font.Bold = true;
            worksheet.get_Range("G24", "G24").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

            worksheet.get_Range("G25", "I25").Merge(false); //CENA
            chartRange = app.get_Range("G25", "I25");
            chartRange.FormulaR1C1 = "Cena bez DPH celkem";
            worksheet.get_Range("G25", "G25").Cells.Font.Size = 10;
            worksheet.get_Range("G25", "G25").Cells.Font.Name = "Arial";
            worksheet.get_Range("G25", "G25").Cells.Font.Bold = true;
            worksheet.get_Range("G25", "G25").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

            worksheet.get_Range("J24", "K24").Merge(false); //DPH
            chartRange = app.get_Range("J24", "K24");
            chartRange.FormulaR1C1 = "DPH %";
            worksheet.get_Range("J24", "J24").Cells.Font.Size = 10;
            worksheet.get_Range("J24", "J24").Cells.Font.Name = "Arial";
            worksheet.get_Range("J24", "J24").Cells.Font.Bold = true;
            worksheet.get_Range("J24", "J24").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

            worksheet.get_Range("J25", "K25").Merge(false); //DPH
            chartRange = app.get_Range("J25", "K25");
            chartRange.FormulaR1C1 = "DPH Kč";
            worksheet.get_Range("J25", "J25").Cells.Font.Size = 10;
            worksheet.get_Range("J25", "J25").Cells.Font.Name = "Arial";
            worksheet.get_Range("J25", "J25").Cells.Font.Bold = true;
            worksheet.get_Range("J25", "J25").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

            worksheet.get_Range("L24", "M24").Merge(false); //CENA
            chartRange = app.get_Range("L24", "M24");
            chartRange.FormulaR1C1 = "Celkem Kč s DPH";
            worksheet.get_Range("L24", "L24").Cells.Font.Size = 10;
            worksheet.get_Range("L24", "L24").Cells.Font.Name = "Arial";
            worksheet.get_Range("L24", "L24").Cells.Font.Bold = true;
            worksheet.get_Range("L24", "L24").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

            worksheet.get_Range("A24", "D25").BorderAround(Type.Missing, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Type.Missing);
            worksheet.get_Range("E24", "F25").BorderAround(Type.Missing, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Type.Missing);
            worksheet.get_Range("G24", "I25").BorderAround(Type.Missing, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Type.Missing);
            worksheet.get_Range("J24", "K25").BorderAround(Type.Missing, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Type.Missing);
            worksheet.get_Range("L24", "M25").BorderAround(Type.Missing, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Type.Missing);

            //konec

            //TEXTY DOLE PODPIS A REGISTRACE
            worksheet.get_Range("A64", "F64").Merge(false); //DPH
            chartRange = app.get_Range("A64", "F64");
            chartRange.FormulaR1C1 = "Registrace:  OR u KS v Ostravě, oddíl C, vložka 22648";
            worksheet.get_Range("A64", "A64").Cells.Font.Size = 10;
            worksheet.get_Range("A64", "A64").Cells.Font.Name = "Arial";
            worksheet.get_Range("A64", "A64").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            worksheet.get_Range("A66", "C66").Merge(false); //DPH
            chartRange = app.get_Range("A66", "C66");
            chartRange.FormulaR1C1 = "Razítko a podpis:";
            worksheet.get_Range("A66", "A66").Cells.Font.Size = 10;
            worksheet.get_Range("A66", "A66").Cells.Font.Name = "Arial";
            worksheet.get_Range("A66", "A66").Cells.Font.Bold = true;
            worksheet.get_Range("A66", "A66").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            //VYSLEDNE CASTKY
            worksheet.get_Range("J64", "K64").Merge(false); 
            chartRange = app.get_Range("J64", "K64");
            chartRange.FormulaR1C1 = "Celkem bez DPH:";
            worksheet.get_Range("J64", "J64").Cells.Font.Size = 11;
            worksheet.get_Range("J64", "J64").Cells.Font.Name = "Arial";
            worksheet.get_Range("J64", "J64").Cells.Font.Bold = true;
            worksheet.get_Range("J64", "J64").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            worksheet.get_Range("J65", "K65").Merge(false);
            chartRange = app.get_Range("J65", "K65");
            chartRange.FormulaR1C1 = "Celkem DPH:";
            worksheet.get_Range("J65", "J65").Cells.Font.Size = 11;
            worksheet.get_Range("J65", "J65").Cells.Font.Name = "Arial";
            worksheet.get_Range("J65", "J65").Cells.Font.Bold = true;
            worksheet.get_Range("J65", "J65").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            worksheet.get_Range("J66", "K66").Merge(false); 
            chartRange = app.get_Range("J66", "K66");
            chartRange.FormulaR1C1 = "Celkem:";
            worksheet.get_Range("J66", "J66").Cells.Font.Size = 11;
            worksheet.get_Range("J66", "J66").Cells.Font.Name = "Arial";
            worksheet.get_Range("J66", "J66").Cells.Font.Bold = true;
            worksheet.get_Range("J66", "J66").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            worksheet.get_Range("J68", "K68").Merge(false); 
            chartRange = app.get_Range("J68", "K68");
            chartRange.FormulaR1C1 = "Celkem k úhradě:";
            worksheet.get_Range("J68", "J68").Cells.Font.Size = 11;
            worksheet.get_Range("J68", "J68").Cells.Font.Name = "Arial";
            worksheet.get_Range("J68", "J68").Cells.Font.Bold = true;
            worksheet.get_Range("J68", "J68").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;


            //VYSLEDNE CASTKY CASTKY

            worksheet.get_Range("L64", "M64").Merge(false);
            chartRange = app.get_Range("L64", "M64");
            chartRange.FormulaR1C1 = "36000";
            worksheet.get_Range("L64", "L64").Cells.Font.Size = 12;
            worksheet.get_Range("L64", "L64").Cells.Font.Name = "Arial";
            worksheet.get_Range("L64", "L64").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            worksheet.get_Range("L64", "L64").NumberFormat = "0.00";

            worksheet.get_Range("L65", "M65").Merge(false);
            chartRange = app.get_Range("L65", "M65");
            chartRange.FormulaR1C1 = "7560";
            worksheet.get_Range("L65", "L65").Cells.Font.Size = 12;
            worksheet.get_Range("L65", "L65").Cells.Font.Name = "Arial";
            worksheet.get_Range("L65", "L65").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            worksheet.get_Range("L65", "L65").NumberFormat = "0.00";

            worksheet.get_Range("L66", "M66").Merge(false);
            chartRange = app.get_Range("L66", "M66");
            chartRange.FormulaR1C1 = "7560";
            worksheet.get_Range("L66", "L66").Cells.Font.Size = 12;
            worksheet.get_Range("L66", "L66").Cells.Font.Name = "Arial";
            worksheet.get_Range("L66", "L66").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            worksheet.get_Range("L66", "L66").NumberFormat = "0.00";
            //worksheet.get_Range("L66", "L66").Style.SetOutlineBorder(Borders.BottomBorder, CellBorderType.Thick, Color.Blue);

            worksheet.get_Range("L68", "M68").Merge(false);
            chartRange = app.get_Range("L68", "M68");
            chartRange.FormulaR1C1 = "43560";
            worksheet.get_Range("L68", "L68").Cells.Font.Size = 12;
            worksheet.get_Range("L68", "L68").Cells.Font.Name = "Arial";
            worksheet.get_Range("L68", "L68").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            worksheet.get_Range("L68", "L68").NumberFormat = "0.00";

            worksheet.get_Range("A26", "M70").BorderAround(Type.Missing, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Type.Missing);
        }
    }
}
