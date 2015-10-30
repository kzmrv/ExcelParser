using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using Range = Microsoft.Office.Interop.Excel.Range;
using System.Drawing;
using System.IO;
using Newtonsoft.Json;
using System.Text.RegularExpressions;

namespace ExcelParser
{
    class Program
    {
        /* Example
       {
       "name":"ОКБ Агион",
       "telephones":["73852775951"],
       "address" : "656043 Барнаул, ул. Гоголя, 57",
       "email|":"ig_barnaul@mail.ru",
       "website": "geonsk.ru",
       "city": "Барнаул",
       "tags":["Алтайский край", "Конструкторские бюро"]"
       }
       */
        static readonly string DEFAULT_PATH;
        public const int maxCells = 2500;
        static Program()
        {
            DEFAULT_PATH = System.IO.Directory.GetCurrentDirectory();
        }
        class Tuple
        {
            public string name;
            public List<string> telephones;
            public string address;
            public string email;
            public string website;
            public string city;
            public List<string> tags;
            public Tuple(string _name, List<string> _telephones, string _address,
                  string _email, string _website, string _city, List<string> _tags)
            {
                name = _name;
                telephones = _telephones;
                address = _address;
                email = _email;
                website = _website;
                city = _city;
                tags = _tags;

            }
        }
        static Workbook openFile()
        {
            string mysheet = DEFAULT_PATH + @"\base.xlsx";
            var excelApp = new Excel.Application();
            excelApp.Visible = true;
            Excel.Workbooks books = excelApp.Workbooks;

            try
            {
                Workbook book = books.Open(mysheet);
            }
            catch (COMException ex)
            {
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
            string currentRegion = "";
            string currentDirection = "";

            Worksheet sheet = (Worksheet)wb.Sheets[1];
            List<Tuple> result = new List<Tuple>();
            for (int i = 3; i < maxCells; i++)
            {

                Excel.Range cell = sheet.Cells[i, 1];
                 if (cell.Interior.ColorIndex == 44)
                 {
                     currentDirection = (String)sheet.Cells[i, 1].Value;
                     continue;
                 }
                if (cell.Interior.ColorIndex == 6)
                {
                    currentRegion = (String)sheet.Cells[i, 1].Value;
                    continue;
                }

                string cName = (String)sheet.Cells[i, 1].Value;
                if (cName == null)
                    continue;
                string cPhones = (String)sheet.Cells[i, 2].Value;
                if (cPhones == null)
                    continue;
                List<string> lPhones = (cPhones.Split(',')).ToList();
                if (lPhones.Count > 1)
                {
                    for (int j = 1; j < lPhones.Count; j++)
                    {
                        lPhones[j] = lPhones[j].Remove(0, 1);
                    }
                }
                for (int q = 0; q < lPhones.Count; q++)
                {
                    string str = Regex.Replace(lPhones[q], "-", "");
                    lPhones[q] = str;
                }
                string cAdress = (String)sheet.Cells[i, 3].Value;
                string cEmail = (String)sheet.Cells[i, 4].Value;
                string cSite = (String)sheet.Cells[i, 5].Value;
                string cCity = (String)sheet.Cells[i, 6].Value;

                string cTags = (String)sheet.Cells[i, 7].Value;
                List<string> lTags = new List<string>();
                if (cTags.Length > 0)
                    lTags.Add(cTags);
                if (currentRegion.Length > 0)
                    lTags.Add(currentRegion);
                if (currentDirection.Length > 0)
                    lTags.Add(currentDirection);
                Tuple newTuple = new Tuple(cName, lPhones, cAdress, cEmail, cSite, cCity, lTags);
                result.Add(newTuple);
                // break; // for single result
            }
            using (FileStream fs = File.Open(DEFAULT_PATH + @"\result.json", FileMode.OpenOrCreate))
            using (StreamWriter sw = new StreamWriter(fs, Encoding.UTF8))
            using (JsonWriter jw = new JsonTextWriter(sw))
            {

                jw.Formatting = Formatting.Indented;
                JsonSerializer serializer = new JsonSerializer();
                foreach (Tuple t in result)
                {
                    serializer.Serialize(jw, t);
                }
            }
        }
        static void Main(string[] args)
        {
            Console.WriteLine(System.IO.Directory.GetCurrentDirectory());
            Workbook book = null;
            try {
                book = openFile();
            }
            catch (Exception e)
            {
                Console.WriteLine("File read failure. Error: " + e.Message);
                Console.WriteLine("Error code: " + e.HResult);
            }
            if (book != null)
            {
                try {
                    ExcelScanInternal(book);
                }
                catch(COMException comexc)
                {
                    Console.WriteLine("COM Exception: please remove all processes that use target files and try again");
                }
                catch(Exception e)
                {
                    Console.WriteLine("Unexpected exception: " + e.Message);
                }
            }
        }
    }
}
