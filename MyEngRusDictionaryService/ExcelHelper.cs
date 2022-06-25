using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace MyEngRusDictionaryService
{
    class ExcelHelper : IDisposable
    {
        private Application _excel;
        private Workbook _workbook;
        private Worksheet _worksheet;

        private string pathToFile = @"B:\Eng.xlsx";

        private List<string?> engWords = new List<string?>();
        private List<string?> rusWords = new List<string?>();

        public string PathToFile
        {
            get { return pathToFile; }
        }

        public List<string?> EngWords
        {
            get { return engWords; }
        }
        public List<string?> RusWords
        {
            get { return rusWords; }
        }

        public ExcelHelper()
        {
            _excel = new Application();
        }

        internal bool Open(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    _workbook = _excel.Workbooks.Open(pathToFile, 0, false, 5, "", "", false, 
                        XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    Worksheet _worksheet;
                    _worksheet = (Worksheet)_workbook.Sheets[1];

                    int numCol = 1;

                    Excel.Range column = _worksheet.UsedRange.Columns[numCol];
                    System.Array engvalues = (System.Array)column.Cells.Value2;
                    engWords = engvalues.OfType<object>().Select(o => Convert.ToString(o)).ToList();

                    numCol = 2;
                    column = _worksheet.UsedRange.Columns[numCol];
                    System.Array rusvalues = (System.Array)column.Cells.Value2;
                    rusWords = rusvalues.OfType<object>().Select(o => Convert.ToString(o)).ToList();
                    //_workbook.Close(false);
                }
                else throw new Exception("The file does not exist!");

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }

        internal bool Add(string engWord, string rusWord)
        {
            try
            {
                _excel.ActiveSheet.Cells[engWords.Count + 1, 1] = engWord;
                _excel.ActiveSheet.Cells[engWords.Count + 1, 2] = rusWord;
                engWords.Add(engWord);
                rusWords.Add(rusWord);
                _workbook.Save();
                //_workbook.Close(false);

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return false;
        }

        public void Quit()
        {
            _excel.Quit();

            _excel = null;
            _workbook = null;

            System.GC.Collect();

        }

        public void Dispose()
        {
            _workbook.Close(0);
            _excel.Quit();
        }
    }
}
