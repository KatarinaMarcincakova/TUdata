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
        public int Height { get; set; }
        public int Weight { get; set; }

    }
    class Program
    {
        public List<PeopleData> People { get; set; }
        public enum Categories
        {
            Name = 1,
            Age = 2,
            Height = 3,
            Weight = 4
        }
        public void GetDataFromExcel()
        {
            Excel.Application excelApp = new Excel.Application();

            Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(@"C:\Users\katul\private\katka\programovanie\TUdata\data.xls", 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets[1];

            Excel.Range excelRange = excelWorksheet.UsedRange;
            int rows = excelRange.Rows.Count;
            int cols = excelRange.Columns.Count;

            for (int r = 0; r < rows; r++) 
            {
                for (int c = 1; c < cols; c++) 
                {
                    Excel.Range range = (Excel.Range)excelWorksheet.Cells[r, c];
                    switch (c)
                    {
                        case (int)Categories.Name:
                            People[r].Name = range.Value.ToString();
                            break;
                        case (int)Categories.Age:
                            People[r].Age = range.Value.ToInt32();
                            break;
                        case (int)Categories.Height:
                            People[r].Height = range.Value.ToInt32();
                            break;
                        case (int)Categories.Weight:
                            People[r].Weight = range.Value.ToInt32();
                            break;
                    }
                }
            }

            excelWorkbook.Close();
            excelApp.Quit();
        }
        public void PrintDataFromExcel() { }

        static void Main(string[] args)
        {
            
        }
    }
}
