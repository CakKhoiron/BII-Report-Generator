using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

using Excel = Microsoft.Office.Interop.Excel;

namespace SendEmail
{
    class Program
    {
        //tambahkan cek bila email =0;
        //semua value baru dimulai pada baris kedua
        public static string scoreIM; //kolom 23
        public static string scoreQT; //kolom 1
        public static string scoreQF;//kolom 2
        public static string scoreQD;//kolom 3
        public static string scoreQB;//kolom 4
        public static string scoreQC;//kolom 5
        public static string scoreQP;//kolom 6
        public static string emailTo;//kolom 17
        public static string bodySend;


        public void getExcelFile()
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"G:\Google Drive\Sekolah\Kampus\Kuliah\Semester 8\EBSR401 Metode Riset Sistem Informasi\Tugas\Dropbox\Skripsi_Mochamad Khoiron\Survey\Mengukur Innovation Mindset (Tanggapan) extracted on 31 May-edited 5.xlsx");
            Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Worksheets[2];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!

            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    
                    //write the value to the console
                    if (i > 1 && xlRange.Cells[i, j] != null && ((Excel.Range)xlWorksheet.Cells[i, j]).Value2 != null)
                    {
                        if (j == 1)
                            scoreQT = (decimal.Round(decimal.Parse(((double)((Excel.Range)xlWorksheet.Cells[i, 1]).Value2).ToString("R")), 1)).ToString();
                        else if (j == 2)
                            scoreQF = (decimal.Round(decimal.Parse(((double)((Excel.Range)xlWorksheet.Cells[i, 2]).Value2).ToString("R")), 1)).ToString();
                        else if (j == 3)
                            scoreQD = (decimal.Round(decimal.Parse(((double)((Excel.Range)xlWorksheet.Cells[i, 3]).Value2).ToString("R")), 1)).ToString();
                        else if (j == 4)
                            scoreQB = (decimal.Round(decimal.Parse(((double)((Excel.Range)xlWorksheet.Cells[i, 4]).Value2).ToString("R")), 1)).ToString();
                        else if (j == 5)
                            scoreQC = (decimal.Round(decimal.Parse(((double)((Excel.Range)xlWorksheet.Cells[i, 5]).Value2).ToString("R")), 1)).ToString();
                        else if (j == 6)
                            scoreQP = (decimal.Round(decimal.Parse(((double)((Excel.Range)xlWorksheet.Cells[i, 6]).Value2).ToString("R")), 1)).ToString();
                        else if (j == 17)
                            emailTo = ((Excel.Range)xlWorksheet.Cells[i, 17]).Value2.ToString();
                        else if (j == 23)
                            scoreIM = (decimal.Round(decimal.Parse(((double)((Excel.Range)xlWorksheet.Cells[i, 23]).Value2).ToString("R")), 1)).ToString();

                        //Console.Write(((Excel.Range)xlWorksheet.Cells[i, j]).Value2.ToString() + "\t");

                    }

                }

                //new line
                Console.Write("Data ke-" + (i-1)+"\r\n");//fisrt row is a header

                if (emailTo != "0" && i > 1)
                {
                    SendEmailAlertFile(emailTo);
                    Console.WriteLine("Nilai yang didapat {0} {1} {2} {3} {4} {5} {6} {7}", scoreIM, scoreQT, scoreQF, scoreQD, scoreQB, scoreQC, scoreQP, emailTo);
                    Console.WriteLine("");
                }
            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }


        private void SendEmailAlertFile(string emailTo)
        {
            //ConsoleLogging.WriteToConsole("Entering Method SendEmailAlertFile()", Thread.CurrentThread.Name.ToString());

            try
            {

                emailTo = emailTo.Trim();
                bodySend = HEmail.body.Replace("@IM", scoreIM).Replace("@QT", scoreQT).Replace("@QF", scoreQF).Replace("@QD", scoreQD).Replace("@QB", scoreQB)
                    .Replace("@QC", scoreQC).Replace("@QP", scoreQP);

                //HEmail.EmailIsHTML = Convert.ToBoolean(AppConfigFromDB["EmailHtml"]);

                HEmail.EmailTo = emailTo;

                HEmail.SendEmail();

                //emailTo.Clear();

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static void Main()
        {

            Program jalan = new Program();
            jalan.getExcelFile();

            Console.WriteLine("Silakan tekan sembarang tombol...");
            Console.ReadKey();
        }


    }
}
