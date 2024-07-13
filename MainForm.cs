using System;
using System.CodeDom;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using LicenseContext = OfficeOpenXml.LicenseContext;
using Microsoft.Win32;
//using static System.Windows.Forms.VisualStyles.VisualStyleElement;
//using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;

namespace MainMenu001
{
    public partial class MainForm : Form
    {
        //private object pro1param01;
        public MainForm()
        {
            InitializeComponent();
            //textBox1.Visible = false;
            //this.BringToFront();
        }
           private void MainForm_Load(object sender, EventArgs e)
        {
         //   InitializeComponent();
        }
        private void creditWIthToolStripMenuItem_Click(object sender, EventArgs e)
        {
           // textBox1.Visible = false;
        }
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        // option 1 nsdl - normal ca - credit - Entry
        private void entryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FileInfo fi = new FileInfo(@"d:\cafiles\NSDLCA001.xlsx");
            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(@"d:\cafiles\NSDLCA001.xlsx");
            }
            else
            {
                //file doesn't exist
                MessageBox.Show("File Does Not Exits");
            }
        }

        //private void importToolStripMenuItem_Click(object sender, EventArgs e)
        //{ System.Diagnostics.Process.Start(@"d:\cafiles\cnvns01.bat");}
        
        // option 1 nsdl - normal ca - credit - generation of ca file
        private void generationOfUploadFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //copy from here
            
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            string sourceXlsxFilePath = @"d:\\cafiles\\nsdlca001.xlsx";
            string targetCsvFilePath = @"d:\\cafiles\\nsdlca001.csv";

            ConvertXlsxToCsv(sourceXlsxFilePath, targetCsvFilePath);
            //            System.Diagnostics.Process.Start(@"d:\bendem\nsdl\cnvnsbd01.bat");
            Console.WriteLine("Conversion complete.");
        }

        private void ConvertXlsxToCsv(string sourceXlsxFilePath, string targetCsvFilePath)
        {
            using (var excelPackage = new ExcelPackage(new FileInfo(sourceXlsxFilePath)))
            {
                int DATA = 0;
                var worksheet1 = excelPackage.Workbook.Worksheets[DATA];
                int rows = worksheet1.Dimension.Rows;
                int columns = worksheet1.Dimension.Columns;

                using (var streamWriter = new StreamWriter(targetCsvFilePath))
                {
                    // Write data rows
                    for (int i = 1; i <= rows; i++)
                    {
                        for (int j = 1; j <= columns; j++)
                        {
                            if (j > 1 && j <= 9)
                            {
                                streamWriter.Write(",");
                            }
                            var cellValue1 = worksheet1.Cells[i, j].Value?.ToString() ?? "";
                            streamWriter.Write(cellValue1);
                        }
                        streamWriter.WriteLine();
                    }
                }
            }
             // copy to here
            SqlConnection con = new SqlConnection("Data Source=VCCIPL-TECH\\VENTURESQLEXP;Initial Catalog=VCCIPL;Integrated Security=True;");
            con.Open();
            SqlCommand sql_cmnd1 = new SqlCommand("sp_firstsbrdeltab", con);
            sql_cmnd1.CommandType = CommandType.StoredProcedure;
            sql_cmnd1.ExecuteNonQuery();
            con.Close();
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = excelApp.Workbooks.Open(@"d:\cafiles\nsdlca001.xlsx");
            Worksheet worksheet = workbook.Sheets["Parameters"];
            string cellValue = worksheet.Range["D2"].Value.ToString();
            File.WriteAllText(@"d:\cafiles\output\frca001.bat", cellValue);
            System.Diagnostics.Process.Start(@"d:\cafiles\output\frca001.bat").WaitForExit();
            workbook.Close();
            excelApp.Quit();
            MessageBox.Show("Process is over and file [d:][CAFILES][NSDL][gencsvfiles] folder generated successfully!");
           
        }

        private void button1_Click(object sender, EventArgs e)
        {
           
            //textBox1.Visible=false;
        }
        private void entryToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            FileInfo fi1 = new FileInfo(@"d:\cafiles\CDSLCA001.xlsx");
            if (fi1.Exists)
            {
                System.Diagnostics.Process.Start(@"d:\cafiles\CDSLCA001.xlsx");
            }
            else
            {
                //file doesn't exist
                MessageBox.Show("File Does Not Exits");
            }

        }
        
        private void generationOfCAFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            {
                System.Diagnostics.Process.Start(@"D:\vccipl_projects\Deployment_Projects\xlstocsvcdsl01\xlstocsvcdsl01\bin\Debug\xlstocsvcdsl01.exe").WaitForExit();
                System.Diagnostics.Process.Start(@"d:\cafiles\cnvcd01.bat").WaitForExit();
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook workbook = excelApp.Workbooks.Open(@"d:\cafiles\cdslca001.xlsx");
                Worksheet worksheet = workbook.Sheets["Parameters"];
                string cellValue = worksheet.Range["E2"].Value.ToString();
                workbook.Close();
                excelApp.Quit();
                File.WriteAllText(@"d:\cafiles\output\frcd001.bat", cellValue);
                System.Diagnostics.Process.Start(@"d:\cafiles\output\frcd001.bat").WaitForExit();
            }
            MessageBox.Show("Your File has been Generated in [D][CAFILES][CDSL][GENCSVFILES] folder successfully");
        }
        private void requestsToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
        private void openFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "All files (*.*)|*.*";
            //textBox1.Visible = true;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;
                string content = File.ReadAllText(filePath);
                //textBox1.Text = content; 
                //textBox1.WordWrap = true;
            }
        }

        private void exitToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            MessageBox.Show("Bye for Now");
            this.Close();
        }
        
        
        private void entryToolStripMenuItem4_Click(object sender, EventArgs e)
        {

            FileInfo fi = new FileInfo(@"d:\cafiles\NSDLCA1A.xlsx");
            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(@"d:\cafiles\NSDLCA1A.xlsx");
            }
            else
            {
                //file doesn't exist
                MessageBox.Show("File Does Not Exits");
            }

        }

        private void generationOfUplFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
                System.Diagnostics.Process.Start(@"D:\vccipl_projects\Deployment_Projects\xlstocsv1a\xlstocsv1a\bin\Debug\xlstocsv1a.exe").WaitForExit();
                System.Diagnostics.Process.Start(@"d:\cafiles\cnvns1A.bat").WaitForExit();
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook workbook = excelApp.Workbooks.Open(@"d:\cafiles\nsdlca1A.xlsx");
                Worksheet worksheet = workbook.Sheets["Parameters"];
                string cellValue = worksheet.Range["G2"].Value.ToString();
                File.WriteAllText(@"d:\cafiles\output\frca1a.bat", cellValue);
                System.Diagnostics.Process.Start(@"d:\cafiles\output\frca1a.bat").WaitForExit();
                workbook.Close();
                excelApp.Quit();
                
                MessageBox.Show("Your File has been Generated in [d:][CAFILES][gencsvfiles] folder successfully");
        }

        private void entryToolStripMenuItem5_Click(object sender, EventArgs e)
        {
            FileInfo fi4 = new FileInfo(@"d:\cafiles\NSDLCA002.xlsx");
            if (fi4.Exists)
            {
                System.Diagnostics.Process.Start(@"d:\cafiles\NSDLCA002.xlsx");
            }
            else
            {
                //file doesn't exist
                MessageBox.Show("File Does Not Exits");
            }
        }

        private void generationOfCAFileToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            {

                System.Diagnostics.Process.Start(@"D:\vccipl_projects\Deployment_Projects\xlstocsv002\xlstocsv002\bin\Debug\xlstocsv002.exe").WaitForExit();
                System.Diagnostics.Process.Start(@"d:\cafiles\cnvns02.bat").WaitForExit();
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook workbook = excelApp.Workbooks.Open(@"d:\cafiles\nsdlca002.xlsx");
                Worksheet worksheet = workbook.Sheets["Parameters"];
                string cellValue = worksheet.Range["G2"].Value.ToString();
                File.WriteAllText(@"d:\cafiles\output\frca002.bat", cellValue);
                System.Diagnostics.Process.Start(@"d:\cafiles\output\frca002.bat").WaitForExit();
                //SqlConnection con = new SqlConnection("Data Source=VCCIPL-TECH\\VENTURESQLEXP;Initial Catalog=VCCIPL;Integrated Security=True;");
                //con.Open();
                workbook.Close();
                excelApp.Quit();
            }
            MessageBox.Show("Your File has been Generated in [d:][CAFILES][NSDL][gencsvfiles] folder successfully");
        }

        private void conversionToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            //System.Diagnostics.Process.Start(@"d:\cafiles\cnvns02.bat");
        }

        private void entryToolStripMenuItem6_Click(object sender, EventArgs e)
        {
            FileInfo fi5 = new FileInfo(@"d:\cafiles\NSDLCA003.xlsx");
            if (fi5.Exists)
            {
                System.Diagnostics.Process.Start(@"d:\cafiles\NSDLCA003.xlsx");
            }
            else
            {
                //file doesn't exist
                MessageBox.Show("File Does Not Exits");
            }
        }

        private void generationOfCAFileToolStripMenuItem4_Click(object sender, EventArgs e)
        {
            {
                System.Diagnostics.Process.Start(@"D:\vccipl_projects\Deployment_Projects\xlstocsv003\xlstocsv003\bin\Debug\xlstocsv003.exe").WaitForExit();
                System.Diagnostics.Process.Start(@"d:\cafiles\cnvns03.bat").WaitForExit();
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook workbook = excelApp.Workbooks.Open(@"d:\cafiles\nsdlca003.xlsx");
                Worksheet worksheet = workbook.Sheets["Parameters"];
                string cellValue = worksheet.Range["G2"].Value.ToString();
                workbook.Close();
                excelApp.Quit();
                File.WriteAllText(@"d:\cafiles\output\frca003.bat", cellValue);
                System.Diagnostics.Process.Start(@"d:\cafiles\output\frca003.bat").WaitForExit();
                //SqlConnection con = new SqlConnection(@"Data Source=VCCIPL-TECH\VENTURESQLEXP;Initial Catalog=VCCIPL;Integrated Security=True;");
                //con.Open();
            }
            MessageBox.Show("Your File has been Generated in [d:][CAFILES][NSDL][gencsvfiles] folder successfully");
        }

        private void entryToolStripMenuItem7_Click(object sender, EventArgs e)
        {
            FileInfo fi6 = new FileInfo(@"d:\cafiles\NSDLCA004.xlsx");
            if (fi6.Exists)
            {
                System.Diagnostics.Process.Start(@"d:\cafiles\NSDLCA004.xlsx");
            }
            else
            {
                //file doesn't exist
                MessageBox.Show("File Does Not Exits");
            }
        }

        
        private void generationOfCAFileToolStripMenuItem5_Click(object sender, EventArgs e)
        {
            {
                System.Diagnostics.Process.Start(@"D:\vccipl_projects\Deployment_Projects\xlstocsv004\xlstocsv004\bin\Debug\xlstocsv004.exe").WaitForExit();
                System.Diagnostics.Process.Start(@"d:\cafiles\cnvns04.bat").WaitForExit();
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook workbook = excelApp.Workbooks.Open(@"d:\cafiles\nsdlca004.xlsx");
                Worksheet worksheet = workbook.Sheets["Parameters"];
                string cellValue = worksheet.Range["G2"].Value.ToString();
                workbook.Close();
                excelApp.Quit();
                File.WriteAllText(@"d:\cafiles\output\frca004.bat", cellValue);
                System.Diagnostics.Process.Start(@"d:\cafiles\output\frca004.bat").WaitForExit();
                //SqlConnection con = new SqlConnection(@"Data Source=VCCIPL-TECH\VENTURESQLEXP;Initial Catalog=VCCIPL;Integrated Security=True;");
                //con.Open();
            }
            MessageBox.Show("Your File has been Generated in [d:][CAFILES][NSDL][gencsvfiles] folder successfully");
        }

        private void entryToolStripMenuItem8_Click(object sender, EventArgs e)
        {
            FileInfo fi7 = new FileInfo(@"d:\cafiles\NSDLCA005.xlsx");
            if (fi7.Exists)
            {
                System.Diagnostics.Process.Start(@"d:\cafiles\NSDLCA005.xlsx");
            }
            else
            {
                //file doesn't exist
                MessageBox.Show("File Does Not Exits");
            }
        }

        private void generationOfCAFileToolStripMenuItem6_Click(object sender, EventArgs e)
        {
            {
                System.Diagnostics.Process.Start(@"d:\cafiles\cnvns05.bat").WaitForExit();
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook workbook = excelApp.Workbooks.Open(@"d:\cafiles\nsdlca005.xlsx");
                Worksheet worksheet = workbook.Sheets["Parameters"];
                string cellValue = worksheet.Range["J2"].Value.ToString();
                workbook.Close();
                excelApp.Quit();
                File.WriteAllText(@"d:\cafiles\output\frca005.bat", cellValue);
                System.Diagnostics.Process.Start(@"d:\cafiles\output\frca005.bat").WaitForExit();
            }
            MessageBox.Show("Your File has been Generated in [d:][CAFILES][NSDL][SCA] folder successfully");
        }

        private void entryToolStripMenuItem9_Click(object sender, EventArgs e)
        {
            FileInfo fi8 = new FileInfo(@"d:\cafiles\NSDLCA006.xlsx");
            if (fi8.Exists)
            {
                System.Diagnostics.Process.Start(@"d:\cafiles\NSDLCA006.xlsx");
            }
            else
            {
                //file doesn't exist
                MessageBox.Show("File Does Not Exits");
            }
        }

        
        private void generationOfCAFileToolStripMenuItem7_Click(object sender, EventArgs e)
        {
            {
                System.Diagnostics.Process.Start(@"D:\vccipl_projects\Deployment_Projects\xlstocsv006\xlstocsv006\bin\Debug\xlstocsv006.exe").WaitForExit();
                System.Diagnostics.Process.Start(@"d:\cafiles\cnvns06.bat").WaitForExit();
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook workbook = excelApp.Workbooks.Open(@"d:\cafiles\nsdlca006.xlsx");
                Worksheet worksheet = workbook.Sheets["Parameters"];
                string cellValue = worksheet.Range["N2"].Value.ToString();
                workbook.Close();
                excelApp.Quit();
                File.WriteAllText(@"d:\cafiles\output\frca006.bat", cellValue);
                System.Diagnostics.Process.Start(@"d:\cafiles\output\frca006.bat").WaitForExit();
            }
            MessageBox.Show("Your File has been Generated in [d:][CAFILES][NSDL][SCA] folder successfully");
        }

        private void entryToolStripMenuItem10_Click(object sender, EventArgs e)
        {
            FileInfo fi9 = new FileInfo(@"d:\cafiles\NSDLCA05b.xlsx");
            if (fi9.Exists)
            {
                System.Diagnostics.Process.Start(@"d:\cafiles\NSDLCA05b.xlsx");
            }
            else
            {
                //file doesn't exist
                MessageBox.Show("File Does Not Exits");
            }
        }


        private void generationOfCAFileToolStripMenuItem8_Click(object sender, EventArgs e)
        {
            {
                System.Diagnostics.Process.Start(@"D:\vccipl_projects\Deployment_Projects\xlstocsv05b\xlstocsv05b\bin\Debug\xlstocsv05b.exe").WaitForExit();
                System.Diagnostics.Process.Start(@"d:\cafiles\cnvns5b.bat").WaitForExit();
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook workbook = excelApp.Workbooks.Open(@"d:\cafiles\nsdlca05b.xlsx");
                Worksheet worksheet = workbook.Sheets["Parameters"];
                string cellValue = worksheet.Range["J2"].Value.ToString();
                workbook.Close();
                excelApp.Quit();
                File.WriteAllText(@"d:\cafiles\output\frca05b.bat", cellValue);
                System.Diagnostics.Process.Start(@"d:\cafiles\output\frca05b.bat").WaitForExit();
            }
            MessageBox.Show("Your File has been Generated in [d:][CAFILES][NSDL][SCA] folder successfully");
        }

        private void entryToolStripMenuItem11_Click(object sender, EventArgs e)
        {
            FileInfo fi10 = new FileInfo(@"d:\cafiles\CDSLCA002.xlsx");
            if (fi10.Exists)
            {
                System.Diagnostics.Process.Start(@"d:\cafiles\CDSLCA002.xlsx");
            }
            else
            {
                //file doesn't exist
                MessageBox.Show("File Does Not Exits");
            }
        }


        private void geneartionOfCAFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            {
                System.Diagnostics.Process.Start(@"D:\vccipl_projects\Deployment_Projects\xlstocsvcdsl02\xlstocsvcdsl02\bin\Debug\xlstocsvcdsl02.exe").WaitForExit();
                System.Diagnostics.Process.Start(@"d:\cafiles\cnvcd02.bat").WaitForExit();
                Microsoft.Office.Interop.Excel.Application excelAppc = new Microsoft.Office.Interop.Excel.Application();
                Workbook workbookc = excelAppc.Workbooks.Open(@"d:\cafiles\cdslca002.xlsx");
                Worksheet worksheetc = workbookc.Sheets["Parameters"];
                string cellValuec = worksheetc.Range["E2"].Value.ToString();
                workbookc.Close();
                excelAppc.Quit();
                File.WriteAllText(@"d:\cafiles\output\frcd002.bat", cellValuec);
                System.Diagnostics.Process.Start(@"d:\cafiles\output\frcd002.bat").WaitForExit();
            }
            MessageBox.Show("Your File has been Generated in [D][CAFILES][CDSL][GENCSVFILES] folder successfully");

        }

        private void entryToolStripMenuItem12_Click(object sender, EventArgs e)
        {
            FileInfo fi11 = new FileInfo(@"d:\cafiles\CDSLCA003.xlsx");
            if (fi11.Exists)
            {
                System.Diagnostics.Process.Start(@"d:\cafiles\CDSLCA003.xlsx");
            }
            else
            {
                //file doesn't exist
                MessageBox.Show("File Does Not Exits");
            }
        }

        private void generationOfCAFileToolStripMenuItem9_Click(object sender, EventArgs e)
        {
            {
                System.Diagnostics.Process.Start(@"D:\vccipl_projects\Deployment_Projects\xlstocsvcdsl03\xlstocsvcdsl03\bin\Debug\xlstocsvcdsl03.exe").WaitForExit();
                System.Diagnostics.Process.Start(@"d:\cafiles\cnvcd03.bat").WaitForExit();
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook workbook = excelApp.Workbooks.Open(@"d:\cafiles\cdslca003.xlsx");
                Worksheet worksheet = workbook.Sheets["Parameters"];
                string cellValue = worksheet.Range["J2"].Value.ToString();
                workbook.Close();
                excelApp.Quit();
                File.WriteAllText(@"d:\cafiles\output\frcd003.bat", cellValue);
                System.Diagnostics.Process.Start(@"d:\cafiles\output\frcd003.bat").WaitForExit();
            }
            MessageBox.Show("Your File has been Generated in [D][CAFILES][CDSL][GENCSVFILES] folder successfully");
        }

        private void entryToolStripMenuItem13_Click(object sender, EventArgs e)
        {
            FileInfo fi12 = new FileInfo(@"d:\cafiles\CDSLCA004.xlsx");
            if (fi12.Exists)
            {
                System.Diagnostics.Process.Start(@"d:\cafiles\CDSLCA004.xlsx");
            }
            else
            {
                MessageBox.Show("File Does Not Exits");
            }
        }

        private void generationOfCAFileToolStripMenuItem10_Click(object sender, EventArgs e)
        {
            {
                System.Diagnostics.Process.Start(@"D:\vccipl_projects\Deployment_Projects\xlstocsvcdsl04\xlstocsvcdsl04\bin\Debug\xlstocsvcdsl04.exe").WaitForExit();
                System.Diagnostics.Process.Start(@"d:\cafiles\cnvcd04.bat").WaitForExit();
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook workbook = excelApp.Workbooks.Open(@"d:\cafiles\cdslca004.xlsx");
                Worksheet worksheet = workbook.Sheets["Parameters"];
                string cellValue = worksheet.Range["J2"].Value.ToString();
                workbook.Close();
                excelApp.Quit();
                File.WriteAllText(@"d:\cafiles\output\frcd004.bat", cellValue);
                System.Diagnostics.Process.Start(@"d:\cafiles\output\frcd004.bat").WaitForExit();
            }
            MessageBox.Show("Your File has been Generated in [D][CAFILES][CDSL][GENCSVFILES] folder successfully");
        }

        private void notepadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start("notepad++.exe", @"d:\sample.txt" );
        }

        private void calculatorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start("calc.exe");
        }

        private void entryToolStripMenuItem14_Click(object sender, EventArgs e)
        {
            FileInfo fi13 = new FileInfo(@"d:\cafiles\CDSLCA005.xlsx");
            if (fi13.Exists)
            {
                System.Diagnostics.Process.Start(@"d:\cafiles\CDSLCA005.xlsx");
            }
            else
            {
                MessageBox.Show("File Does Not Exits");
            }
        }
        
        private void generationOfToolStripMenuItem_Click(object sender, EventArgs e)
        {
            {

                System.Diagnostics.Process.Start(@"D:\vccipl_projects\Deployment_Projects\xlstocsvcdsl05\xlstocsvcdsl05\bin\Debug\xlstocsvcdsl05.exe").WaitForExit();
                System.Diagnostics.Process.Start(@"d:\cafiles\cnvcd05.bat").WaitForExit();
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook workbook = excelApp.Workbooks.Open(@"d:\cafiles\cdslca005.xlsx");
                Worksheet worksheet = workbook.Sheets["Parameters"];
                string cellValue = worksheet.Range["J2"].Value.ToString();
                File.WriteAllText(@"d:\cafiles\output\frcd005.bat", cellValue);
                System.Diagnostics.Process.Start(@"d:\cafiles\output\frcd005.bat").WaitForExit();
                workbook.Close();
                excelApp.Quit();
            }
            MessageBox.Show("Your File has been Generated in [D][CAFILES][CDSL][GENCSVFILES] folder successfully");
        }

        private void bothDebitCreditWithLockinExpiryDateToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
        private void eNTRYToolStripMenuItem15_Click_1(object sender, EventArgs e)
        {
            FileInfo fi14 = new FileInfo(@"d:\cafiles\CDSLCA006.xlsx");
            if (fi14.Exists)
            {
                System.Diagnostics.Process.Start(@"d:\cafiles\CDSLCA006.xlsx");
            }
            else
            {
                MessageBox.Show("File Does Not Exits");
            }
        }
        
        private void generationOfCAFileToolStripMenuItem11_Click_1(object sender, EventArgs e)
        {
            {
                System.Diagnostics.Process.Start(@"D:\vccipl_projects\Deployment_Projects\xlstocsvcdsl06\xlstocsvcdsl06\bin\Debug\xlstocsvcdsl06.exe").WaitForExit();
                System.Diagnostics.Process.Start(@"d:\cafiles\cnvcd06.bat").WaitForExit();
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook workbook = excelApp.Workbooks.Open(@"d:\cafiles\cdslca006.xlsx");
                Worksheet worksheet = workbook.Sheets["Parameters"];
                string cellValue = worksheet.Range["J2"].Value.ToString();
                File.WriteAllText(@"d:\cafiles\output\frcd006.bat", cellValue);
                System.Diagnostics.Process.Start(@"d:\cafiles\output\frcd006.bat").WaitForExit();
                workbook.Close();
                excelApp.Quit();
            }
            MessageBox.Show("Your File has been Generated in [D][CAFILES][CDSL][GENCSVFILES] folder successfully");
        }
        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            FileInfo fi18 = new FileInfo(@"d:\AIF\AIFNSDLCA01A.xlsx");
            if (fi18.Exists)
            {
                System.Diagnostics.Process.Start(@"d:\AIF\AIFNSDLCA01A.xlsx");
            }
            else
            {
                MessageBox.Show("File Does Not Exits");
            }
        }

        private void entryToolStripMenuItem16_Click(object sender, EventArgs e)
        {
            FileInfo fi16 = new FileInfo(@"D:\AIF\AIFCDSLCA001.xlsx");
            if (fi16.Exists)
            {
                System.Diagnostics.Process.Start(@"D:\AIF\AIFCDSLCA001.xlsx");
            }
            else
            {
                MessageBox.Show("File Does Not Exits");
            }
        }

        private void conversionToCSVToolStripMenuItem4_Click(object sender, EventArgs e)
        {
            //System.Diagnostics.Process.Start(@"D:\AIF\cnvaifcdsl01.bat");
        }
        private void entryToolStripMenuItem17_Click(object sender, EventArgs e)
        {
            FileInfo fi15 = new FileInfo(@"d:\AIF\AIFNSDLCAAA1.xlsx");
            if (fi15.Exists)
            {
                System.Diagnostics.Process.Start(@"d:\AIF\AIFNSDLCAAA1.xlsx");
            }
            else
            {
                MessageBox.Show("File Does Not Exits");
            }
        }

        private void conversionOfCSVToolStripMenuItem1_Click(object sender, EventArgs e)
        {

            //System.Diagnostics.Process.Start(@"d:\AIF\cnvaifnscaa1.bat");
        }

        private void generationOfCAFileToolStripMenuItem13_Click(object sender, EventArgs e)
        {
            {
                System.Diagnostics.Process.Start(@"D:\vccipl_projects\Deployment_Projects\xlstocsvaifn02\xlstocsvaifn02\bin\Debug\xlstocsvaifn02.exe").WaitForExit();
                System.Diagnostics.Process.Start(@"d:\AIF\cnvaifnscaa1.bat").WaitForExit();
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook workbook = excelApp.Workbooks.Open(@"d:\AIF\AIFNSDLCAAA1.xlsx");
                Worksheet worksheet = workbook.Sheets["Parameters"];
                string cellValue = worksheet.Range["d2"].Value.ToString();
                File.WriteAllText(@"d:\cafiles\output\fraifna1.bat", cellValue);
                System.Diagnostics.Process.Start(@"d:\cafiles\output\fraifna1.bat").WaitForExit();
                workbook.Close();
                excelApp.Quit();
            }
            MessageBox.Show("Your File has been Generated in [D][AIF][GENCSVFILES] folder successfully");
        }

        private void entryToolStripMenuItem2_Click_2(object sender, EventArgs e)
        {
            //AIFNREDCA001
            FileInfo fi15 = new FileInfo(@"d:\AIF\AIFNREDCA001.xlsx");
            if (fi15.Exists)
            {
                System.Diagnostics.Process.Start(@"d:\AIF\AIFNREDCA001.xlsx");
            }
            else
            {
                MessageBox.Show("File Does Not Exits");
            }
        }
        private void generationOfCAFileToolStripMenuItem1_Click_1(object sender, EventArgs e)
        {
            {
            System.Diagnostics.Process.Start(@"D:\vccipl_projects\Deployment_Projects\xlstocsvaifn03\xlstocsvaifn03\bin\Debug\xlstocsvaifn03.exe").WaitForExit();
            System.Diagnostics.Process.Start(@"d:\AIF\cnvaifnredca01.bat").WaitForExit();
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = excelApp.Workbooks.Open(@"d:\AIF\AIFNREDCA001.xlsx");
            Worksheet worksheet = workbook.Sheets["Parameters"];
            string cellValue = worksheet.Range["d2"].Value.ToString();
            File.WriteAllText(@"d:\cafiles\output\fraifnred01.bat", cellValue);
            System.Diagnostics.Process.Start(@"d:\cafiles\output\fraifnred01.bat").WaitForExit();
            workbook.Close();
            excelApp.Quit();
            }
            MessageBox.Show("Your File has been Generated in [D][AIF][GENCSVFILES] folder successfully");
        }

        private void entryToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            FileInfo fi17 = new FileInfo(@"d:\AIF\AIFNSDLCA001.xlsx");
            if (fi17.Exists)
            {
                System.Diagnostics.Process.Start(@"d:\AIF\AIFNSDLCA001.xlsx");
            }
            else
            {
                MessageBox.Show("File Does Not Exits");
            }

        }
        private void generationOfCAFileToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            {
                System.Diagnostics.Process.Start(@"D:\vccipl_projects\Deployment_Projects\xlstocsvaifc01\xlstocsvaifc01\bin\Debug\xlstocsvaifc01.exe").WaitForExit();
                System.Diagnostics.Process.Start(@"D:\AIF\cnvaifcdsl01.bat").WaitForExit();
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook workbook = excelApp.Workbooks.Open(@"d:\AIF\AIFCDSLCA001.xlsx");
                Worksheet worksheet = workbook.Sheets["Parameters"];
                string cellValue = worksheet.Range["D2"].Value.ToString();
                File.WriteAllText(@"d:\cafiles\output\fraifccd01.bat", cellValue);
                System.Diagnostics.Process.Start(@"d:\cafiles\output\fraifccd01.bat").WaitForExit();
                workbook.Close();
                excelApp.Quit();
            }
            MessageBox.Show("Your File has been Generated in [D][AIF][CDSL][GENCSVFILES] folder successfully");
        }
        private void toolStripMenuItem5_Click(object sender, EventArgs e)
        {
            {
                System.Diagnostics.Process.Start(@"D:\vccipl_projects\Deployment_Projects\xlstocsvepfn01\xlstocsvepfn01\bin\Debug\xlstocsvepfn01.exe").WaitForExit();
                System.Diagnostics.Process.Start(@"d:\AIF\cnvAIFNSCA1A.bat").WaitForExit();
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook workbook = excelApp.Workbooks.Open(@"D:\AIF\AIFNSDLCA01A.xlsx");
                Worksheet worksheet = workbook.Sheets["Parameters"];
                string cellValue = worksheet.Range["D2"].Value.ToString();
                File.WriteAllText(@"d:\cafiles\output\fraifNEPF01.bat", cellValue);
                System.Diagnostics.Process.Start(@"d:\cafiles\output\fraifNEPF01.bat").WaitForExit();
                workbook.Close();
                excelApp.Quit();
            }
            MessageBox.Show("Your File has been Generated in [D][AIF][CDSL][GENCSVFILES] folder successfully");
        }

        private void entryToolStripMenuItem18_Click(object sender, EventArgs e)
        {
            FileInfo fi18 = new FileInfo(@"d:\AIF\AIFCDSLCA01A.xlsx");
            if (fi18.Exists)
            {
                System.Diagnostics.Process.Start(@"d:\AIF\AIFCDSLCA01A.xlsx");
            }
            else
            {
                MessageBox.Show("File Does Not Exits");
            }
        }
        
        private void generationOfCAFileToolStripMenuItem14_Click(object sender, EventArgs e)
        {
            {
                System.Diagnostics.Process.Start(@"D:\vccipl_projects\Deployment_Projects\xlstocsvepfc01\xlstocsvepfc01\bin\Debug\xlstocsvepfc01.exe").WaitForExit();
                System.Diagnostics.Process.Start(@"d:\AIF\cnvaifcsca1a.bat").WaitForExit();
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook workbook = excelApp.Workbooks.Open(@"d:\AIF\AIFCDSLCA01A.xlsx");
                Worksheet worksheet = workbook.Sheets["Parameters"];
                string cellValue = worksheet.Range["D2"].Value.ToString();
                File.WriteAllText(@"d:\cafiles\output\fraifcfcd01.bat", cellValue);
                System.Diagnostics.Process.Start(@"d:\cafiles\output\fraifcfcd01.bat").WaitForExit();
                workbook.Close();
                excelApp.Quit();
            }
            MessageBox.Show("Your File has been Generated in [D][AIF][CDSL][FRACTION] folder successfully");
        }
        private void helpToolStripMenuItem_Click(object sender, EventArgs e)
        {
                MessageBox.Show("Please bear with us for sometime - Manual on this Application!!!");
        }
        private void generationOfCAFileToolStripMenuItem12_Click(object sender, EventArgs e)
        {
            {
                System.Diagnostics.Process.Start(@"D:\vccipl_projects\Deployment_Projects\xlstocsvaifn01\xlstocsvaifn01\bin\Debug\xlstocsvaifn01.exe").WaitForExit();
                System.Diagnostics.Process.Start(@"d:\aif\cnvaifnsca01.bat").WaitForExit();
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook workbook = excelApp.Workbooks.Open(@"d:\AIF\AIFNSDLCA001.xlsx");
                Worksheet worksheet = workbook.Sheets["Parameters"];
                string cellValue = worksheet.Range["D2"].Value.ToString();
                File.WriteAllText(@"d:\cafiles\output\fraifnsdca01.bat", cellValue);
                System.Diagnostics.Process.Start(@"d:\cafiles\output\fraifnsdca01.bat").WaitForExit();
                workbook.Close();
                excelApp.Quit();
            }
            MessageBox.Show("Your File has been Generated in [D][AIF][CDSL][FRACTION] folder successfully");
        }
        private void notepadToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Process.Start("notepad.exe");
        }
        private void stampDutyCalculatorToolStripMenuItem_Click(object sender, EventArgs e)
        {
        System.Diagnostics.Process.Start("https://nsdl.co.in/stampduty_calculator.php");
        }

        private void entryToolStripMenuItem19_Click(object sender, EventArgs e)
        {
            FileInfo fi13 = new FileInfo(@"d:\cafiles\CDSLCA007.xlsx");
            if (fi13.Exists)
            {
                System.Diagnostics.Process.Start(@"d:\cafiles\CDSLCA007.xlsx");
            }
            else
            {
                MessageBox.Show("File Does Not Exits");
            }
        }

        private void generationOfCAFileToolStripMenuItem15_Click(object sender, EventArgs e)
        {
            // sbr
            {

                System.Diagnostics.Process.Start(@"D:\vccipl_projects\Deployment_Projects\xlstocsvcdsl07\xlstocsvcdsl07\bin\Debug\xlstocsvcdsl07.exe").WaitForExit();
                System.Diagnostics.Process.Start(@"d:\cafiles\cnvcd07.bat").WaitForExit();
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook workbook = excelApp.Workbooks.Open(@"d:\cafiles\cdslca007.xlsx");
                Worksheet worksheet = workbook.Sheets["Parameters"];
                string cellValue = worksheet.Range["J2"].Value.ToString();
                File.WriteAllText(@"d:\cafiles\output\frcd007.bat", cellValue);
                System.Diagnostics.Process.Start(@"d:\cafiles\output\frcd007.bat").WaitForExit();
                workbook.Close();
                excelApp.Quit();
            }
            MessageBox.Show("Your File has been Generated in [D][CAFILES][CDSL][GENCSVFILES] folder successfully");
        }

        private void escrowToClientAccountsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // clear
        }

        private void entryToolStripMenuItem20_Click(object sender, EventArgs e)
        {
            //entry
            FileInfo fi9 = new FileInfo(@"d:\cafiles\NSDLCA05c.xlsx");
            if (fi9.Exists)
            {
                System.Diagnostics.Process.Start(@"d:\cafiles\NSDLCA05C.xlsx");
            }
            else
            {
                //file doesn't exist
                MessageBox.Show("File Does Not Exits");
            }
        }

        private void generationOfSCAFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //generation
            {
                System.Diagnostics.Process.Start(@"D:\vccipl_projects\Deployment_Projects\xlstocsv05c\xlstocsv05c\bin\Debug\xlstocsv05c.exe").WaitForExit();
                System.Diagnostics.Process.Start(@"d:\cafiles\cnvns5c.bat").WaitForExit();
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook workbook = excelApp.Workbooks.Open(@"d:\cafiles\nsdlca05c.xlsx");
                Worksheet worksheet = workbook.Sheets["Parameters"];
                string cellValue = worksheet.Range["J2"].Value.ToString();
                workbook.Close();
                excelApp.Quit();
                File.WriteAllText(@"d:\cafiles\output\frca05c.bat", cellValue);
                System.Diagnostics.Process.Start(@"d:\cafiles\output\frca05c.bat").WaitForExit();
            }
            MessageBox.Show("Your File has been Generated in [d:][CAFILES][NSDL][SCA] folder successfully");
        }

        private void toolStripMenuItem4_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("C:\\Program Files\\Microsoft Office\\root\\Office16\\winword.exe");
        }

        private void toolStripMenuItem7_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("C:\\Program Files\\Microsoft Office\\root\\Office16\\excel.exe");
        }
    }
}

