using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace WorksheetMakerC
{
    public partial class Form1 : Form
    {
        private string m_strSrcPath;
        //private string pathStr = Directory.GetCurrentDirectory();
        private string m_strTempPath = Path.Combine(Environment.CurrentDirectory, @"ASMtemplate.xlsx"); //"C:\\Users\\eric\\Desktop\\CreateCrmAccount\\template\\ClientTemplate_Shipper.xlsx";
        private string m_strFolderPath;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application tXL;
            Excel._Workbook sWB, tWB;
            Excel._Worksheet tSheet;

            
            
            if (textBoxSrcPath.Text == string.Empty || textBoxFolderPath.Text == string.Empty)
            {
                MessageBox.Show("Please choose source file and folder to save output files.");
                return;
            }

            button1.Enabled = false;
            try
            {

                   
                //Start Excel and get Applicaiton object.
                tXL = new Excel.Application
                {
                    Visible = false
                };



                //Get a new workbook on tWB, open workbook on sWB
                //tWB = (Excel._Workbook)(tXL.Workbooks.Add(Missing.Value));


                //tSheet = (Excel._Worksheet)tWB.ActiveSheet;
                sWB = (Excel._Workbook)(tXL.Workbooks.Open(m_strSrcPath));
                //sRng = sSheet.UsedRange;


                foreach (Excel.Worksheet sheet in sWB.Sheets)
                {
                    var v = sheet.Visible;  // check if sheet is hidden, if hidden ignore
                    if (v == Excel.XlSheetVisibility.xlSheetHidden)
                    {
                        continue;
                    }
                    tWB = (Excel._Workbook)tXL.Workbooks.Open(m_strTempPath);
                    tSheet = tWB.Sheets[2];

                    //int fullRow = sheet.Rows.Count;
                    Excel.Range firstCell = sheet.Cells[1, 1];
                    int rowCount = firstCell.get_End(Excel.XlDirection.xlDown).Row;
                    // below two methods get last row doesn't work, the number is too large, not the last cell with value
                    //Excel.Range last = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                    //int rowCount = last.Row;
                    //int colCount = last.Column;
                    //sRng = sheet.UsedRange;
                    //sRng.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    //int rowCount = sRng.Rows.Count;
                    //int colCount = sRng.Columns.Count;
                    //int j = 2;
                    for (int i = 12; i <= rowCount; i++)
                    {
                        tSheet.Cells[i - 10, 5] = "CN";  // column E ----CN
                        tSheet.Cells[i - 10, 6] = sheet.Cells[i, 1];
                        tSheet.Cells[i - 10, 7] = sheet.Cells[i, 8];
                        tSheet.Cells[i - 10, 9] = "100";
                        tSheet.Cells[i - 10, 10] = "4000000";
                        tSheet.Cells[i - 10, 11] = sheet.Cells[i, 4];
                        tSheet.Cells[i - 10, 13] = sheet.Cells[i, 9];
                        tSheet.Cells[i - 10, 14] = sheet.Cells[11, 6];
                        tSheet.Cells[i - 10, 16] = sheet.Cells[i, 7];
                        tSheet.Cells[i - 10, 19] = "B";
                        tSheet.Cells[i - 10, 20] = 0;
                        tSheet.Cells[i - 10, 24] = "S";
                        tSheet.Cells[i - 10, 25] = sheet.Cells[i, 3];
                        tSheet.Cells[i - 10, 26] = "PK";
                        tSheet.Cells[i - 10, 27] = "N/M";
                        tSheet.Cells[i - 10, 53] = "Z";
                        tSheet.Cells[i - 10, 54] = "380";
                        tSheet.Cells[i - 10, 55] = sheet.Cells[6, 2];
                        tSheet.Cells[i - 10, 65] = sheet.Cells[2, 2];

                    }

                    //j +=1
                    //string currentDate = DateTime.Today.ToString("d").Replace("/", "");
                    string fName = m_strFolderPath + "\\" + sheet.Name;

                    tXL.DisplayAlerts = false;
                    tWB.SaveAs(fName);
                    tWB.Close();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(tWB);

                }





                sWB.Close();

                tXL.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(sWB);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(tXL);
                button1.Enabled = true;
            }
            catch (Exception theException)
            {

                string errorMessage;
                errorMessage = "Error: ";
                errorMessage = string.Concat(errorMessage, theException.Message);
                errorMessage = string.Concat(errorMessage, " Line: ");
                errorMessage = string.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            }


                MessageBox.Show("Program finished successfully.");
                //Add table headers going cell by cell.
                //tSheet.Cells[1, 1] = sSheet.Cells[1, 1];
                //tSheet.Cells[1, 2] = "Last Name";
                //tSheet.Cells[1, 3] = "Full Name";
                //tSheet.Cells[1, 4] = "Salary";

                ////Format A1:D1 as bold, vertical alignment = center.

                //tSheet.get_Range("A1", "D1").Font.Bold = true;
                //tSheet.get_Range("A1", "D1").VerticalAlignment =
                //Excel.XlVAlign.xlVAlignCenter;

                ////Create an array to multiple values at once.
                //string[,] saNames = new string[5, 2];

                //saNames[0, 0] = "John";
                //saNames[0, 1] = "Smith";
                //saNames[1, 0] = "Tom";
                //saNames[1, 1] = "Brown";
                //saNames[2, 0] = "Sue";
                //saNames[2, 1] = "Thomas";
                //saNames[3, 0] = "Jane";
                //saNames[3, 1] = "Jones";
                //saNames[4, 0] = "Adam";
                //saNames[4, 1] = "Johnson";

                ////Fill A2:B6 with an array of values (First and Last Names).
                //tSheet.get_Range("A2", "B6").Value2 = saNames;

                ////Fill C2:C6 with a relative formula (=A2 & " " & B2).
                //tRng = tSheet.get_Range("C2", "C6");
                //tRng.Formula = "=A2 & \" \" & B2";

                ////Fill D2:D6 with a formula (=RAND()*100000) and apply format.
                //tRng = tSheet.get_Range("D2", "D6");
                //tRng.Formula = "=RAND()*100000";
                //tRng.NumberFormat = "$0.00";

                ////AutoFit columns A:D.
                //tRng = tSheet.get_Range("A1", "D1");
                //tRng.EntireColumn.AutoFit();

                

                ////Make sure Excel is visible and give the user control
                ////of Microsoft Excel's lifetime.
                //tXL.Visible = true;
                //tXL.UserControl = true;



            
            
        }


        private void btnSelFile_Click(object sender, EventArgs e)
        {
            openFileDialog.Filter = "Excel File(*.xlsx) | *.xlsx;";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                m_strSrcPath = openFileDialog.FileName;
                textBoxSrcPath.AppendText(m_strSrcPath);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog.Filter = "Excel File(*.xlsx) | *.xlsx;";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                m_strTempPath = openFileDialog.FileName;
                textBoxTempPath.AppendText(m_strTempPath);
            }
        }

        private static Random random = new Random();
        public static string RandomString(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            return new string(Enumerable.Repeat(chars, length)
                .Select(s => s[random.Next(s.Length)]).ToArray());
        }

        private void btnFolderPath_Click(object sender, EventArgs e)
        {
            //FolderBrowserDialog fbd = new FolderBrowserDialog();

            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                m_strFolderPath = folderBrowserDialog.SelectedPath;
                textBoxFolderPath.Text = m_strFolderPath;
            }
            else
                m_strFolderPath = string.Empty;
        }

        
    }
    
}
