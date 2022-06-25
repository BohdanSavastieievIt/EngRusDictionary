using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
namespace Dictionary
{
    public class EngRusDictionary
    {
        private string pathToFile = @"B:\Eng.xlsx";

        public List<string?> dictionary = new List<string?>();
        public EngRusDictionary()
        {

        }

        public void ExcelHandler()
        {
            //Создаём приложение.
            Application ObjExcel = new Application();
            //Открываем книгу.                                                                                                                                                        
            Workbook ObjWorkBook = ObjExcel.Workbooks.Open(pathToFile, 0, false, 5, "", "", false, XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //Выбираем таблицу(лист).
            Worksheet ObjWorkSheet;
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];

            // Указываем номер столбца (таблицы Excel) из которого будут считываться данные.
            int numCol = 1;
            
            Excel.Range usedColumn = ObjWorkSheet.UsedRange.Columns[numCol];
            System.Array myvalues = (System.Array)usedColumn.Cells.Value2;
            dictionary = myvalues.OfType<object>().Select(o => o.ToString()).ToList();

            // Выходим из программы Excel.
            ObjExcel.Quit();
        }

    }
}