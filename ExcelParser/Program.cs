using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace ExcelParser
{
    class Program
    {        
        static Workbook openFile()
        {
            string mysheet = @"C:\Vasili\base.xlsx";
            var excelApp = new Excel.Application();
            excelApp.Visible = true;
            Excel.Workbooks books = excelApp.Workbooks;
           
            try {
                Workbook book = books.Open(mysheet);
            }
            catch (COMException ex) {
                Console.WriteLine(ex.ErrorCode);
                Console.WriteLine("HR CODE:" + ex.HResult);
                Console.WriteLine(ex.Message);
                Console.Read();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return books.Item[1];

        }
        static void ExcelScanInternal(Workbook wb)
        {
            Worksheet sheet = (Worksheet)wb.Sheets[1];
            Excel.Range xlrange = (Range)sheet.Cells[1, 1];
            
        }
        static void Main(string[] args)
        {
            Workbook book = openFile();
            ExcelScanInternal(book);
        }
    }
}
