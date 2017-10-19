using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace UploadConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            string filename = "";
            char[] fileExt = { '.', 'x', 'l', 's' };
            Application excel = new Application();
            Workbook workbook = null;
            excel.DisplayAlerts = false;

            convertToCSV(readXLSFiles());

            excel.DisplayAlerts = true;
            excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            Console.ReadKey();

            string[] readXLSFiles()
            {
                string[] XLSFiles = Directory.GetFiles("C:\\Users\\AG20375\\Documents\\Uploads", "*.xls*");
                return XLSFiles;
            }

            string[] readCSVFiles()
            {
                string[] CSVFiles = Directory.GetFiles("C:\\Users\\AG20375\\Documents\\Uploads", "*.csv");
                return CSVFiles;
            }
            void convertToCSV(string[] XLSFiles)
            {
                foreach (string file in XLSFiles)
                {
                    filename = file.Trim(fileExt);
                    workbook = excel.Workbooks.Open(file, ReadOnly: false, Editable: true);
                    workbook.SaveAs(Filename: filename + ".csv", FileFormat: XlFileFormat.xlCSV);
                    workbook.Close();
                }
            }
        }
    }
}
