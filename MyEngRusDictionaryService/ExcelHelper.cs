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
        private Workbook? _workbook;
        private List<string?> engWords = new List<string?>();
        private List<string?> rusWords = new List<string?>();
        private List<int> correctResults = new List<int>();
        private List<int> testAttempts = new List<int>();
        private List<double> scores = new List<double>();


        public string PathToFile { get; } = @"B:\Eng.xlsx";

        public List<string?> EngWords
        {
            get { return engWords; }
        }
        public List<string?> RusWords
        {
            get { return rusWords; }
        }
        public List<int> CorrectResults
        {
            get { return correctResults; }
        }
        public List<int> TestAttempts
        {
            get { return testAttempts; }
        }
        public List<double> Scores
        {
            get { return scores; }
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
                    _workbook = _excel.Workbooks.Open(PathToFile, 0, false, 5, "", "", false, 
                        XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    Worksheet _worksheet;
                    _worksheet = (Worksheet)_workbook.Sheets[1];

                    int numCol = 1;

                    Excel.Range column = _worksheet.UsedRange.Columns[numCol];
                    Array engvalues = (Array)column.Cells.Value2;
                    engWords = engvalues.OfType<object>().Select(o => Convert.ToString(o)).ToList();

                    numCol = 2;
                    column = _worksheet.UsedRange.Columns[numCol];
                    Array rusvalues = (Array)column.Cells.Value2;
                    rusWords = rusvalues.OfType<object>().Select(o => Convert.ToString(o)).ToList();

                    numCol = 3;
                    column = _worksheet.UsedRange.Columns[numCol];
                    Array correctAttempts = (Array)column.Cells.Value2;
                    correctResults = correctAttempts.OfType<object>().Select(o => Convert.ToInt32(o)).ToList();

                    numCol = 4;
                    column = _worksheet.UsedRange.Columns[numCol];
                    Array totalAttempts = (Array)column.Cells.Value2;
                    testAttempts = totalAttempts.OfType<object>().Select(o => Convert.ToInt32(o)).ToList();

                    numCol = 5;
                    column = _worksheet.UsedRange.Columns[numCol];
                    Array score = (Array)column.Cells.Value2;
                    scores = score.OfType<object>().Select(o => Convert.ToDouble(o)).ToList();
                }
                else throw new FileNotFoundException("The file does not exist!");

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
                _excel.ActiveSheet.Cells[engWords.Count + 1, 3] = 0;
                _excel.ActiveSheet.Cells[engWords.Count + 1, 4] = 0;
                _excel.ActiveSheet.Cells[engWords.Count + 1, 5] = 0;
                engWords.Add(engWord);
                rusWords.Add(rusWord);
                correctResults.Add(0);
                testAttempts.Add(0);
                scores.Add(0);
                _workbook.Save();

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return false;
        }
        internal bool WordsInTestUpdate(List<int> isWordInTest, List<int> isCorrect)
        {
            try
            {
                for (int i = 0; i < isCorrect.Count; i++)
                {
                    _excel.ActiveSheet.Cells[isCorrect[i] + 1, 3].Value2 += 1;
                    correctResults[isCorrect[i]] += 1;
                }
                for (int i = 0; i < isWordInTest.Count; i++)
                {
                    _excel.ActiveSheet.Cells[isWordInTest[i] + 1, 4].Value2 += 1;
                    testAttempts[isCorrect[i]] += 1;
                    _excel.ActiveSheet.Cells[isWordInTest[i] + 1, 5].Value2 = 
                        _excel.ActiveSheet.Cells[isWordInTest[i] + 1, 3].Value2 / _excel.ActiveSheet.Cells[isWordInTest[i] + 1, 4].Value2;
                    scores[i] = (double)correctResults[i] / testAttempts[i];

                }
                _workbook.Save();

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
