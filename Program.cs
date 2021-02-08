using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace TUdata
{
    public class PeopleData
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public double Height { get; set; }
        public double Weight { get; set; }

    }
    class Program
    {
        public static List<PeopleData> People = new List<PeopleData>();
        public enum Categories
        {
            Name = 1,
            Age = 2,
            Height = 3,
            Weight = 4
        }
        public static void GetDataFromExcel(int rows, int cols)
        {
            Excel.Application excelApp = new Excel.Application();

            Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(@"C:\Users\katul\private\katka\programovanie\TUdata\data.xls", 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets[1];

            /*
            Excel.Range excelRange = excelWorksheet.UsedRange;
            int rows = excelRange.Rows.Count;
            int cols = excelRange.Columns.Count;
            */

            for (int r = 2; r < rows+1; r++) 
            {
                for (int c = 1; c < cols; c++) 
                {
                    People.Add(new PeopleData() { Age = 0, Height = 0, Name = "", Weight = 0});

                    Excel.Range range = (Excel.Range)excelWorksheet.Cells[r, c];

                    string value = range.Value.ToString();
                    switch (c)
                    {
                        case (int)Categories.Name:
                            People[r-2].Name = value;
                            break;
                        case (int)Categories.Age:
                            People[r-2].Age = Int32.Parse(value);
                            break;
                        case (int)Categories.Height:
                            People[r-2].Height = Double.Parse(value);
                            break;
                        case (int)Categories.Weight:
                            People[r-2].Weight = Double.Parse(value);
                            break;
                    }
                }
            }

            excelWorkbook.Close();
            excelApp.Quit();
        }
        public static void PrintDataFromExcel(int numOfPeople)
        {
            Console.WriteLine("NAME AGE HEIGHT WEIGHT");
            for (int i = 0; i < numOfPeople; i++)
            {
                Console.WriteLine($"{People[i].Name} - {People[i].Age.ToString()} {People[i].Height.ToString()} {People[i].Weight.ToString()}");
            }
        }

        static void Main(string[] args)
        {
            GetDataFromExcel(29, 14);
            PrintDataFromExcel(28);
        }
    }
}
