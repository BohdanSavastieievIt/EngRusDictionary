using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;

namespace MyEngRusDictionaryService
{
    class DictionaryService
    {
        private readonly ExcelHelper helper;
        public DictionaryService() 
        {
            helper = new ExcelHelper();
            
        }
        public void StartProgram()
        {
            if (helper.Open(filePath: helper.PathToFile))
            {
                Console.BackgroundColor = ConsoleColor.Cyan;
                Console.ForegroundColor = ConsoleColor.DarkMagenta;
                Console.Clear();
                Menu();
            }
            else
            {
                Console.WriteLine("The file does not exist!\n");
            }
        }
        private void Menu()
        {
            Console.WriteLine("1. Enter '1' to see the dictionary");
            Console.WriteLine("2. Enter '2' to take the 'From Eng to Rus' test");
            Console.WriteLine("3. Enter '3' to take the 'From Rus to Eng' test");
            Console.WriteLine("4. Enter '4' to add the word to the dictionary");
            Console.WriteLine("5. Enter '5' to see the dictionary in order of test results");
            Console.WriteLine("0. Enter '0' to quit the program");
            Console.WriteLine();
            
            Int32.TryParse(Console.ReadLine(), out int choice);
            Console.WriteLine();

            switch (choice){
                case 1:
                    Show();
                    Menu();
                    break;
                case 2:
                    Test(0);
                    Menu();
                    break;
                case 3:
                    Test(1);
                    Menu();
                    break;
                case 4:
                    AddWord();
                    Menu();
                    break;
                case 5:
                    ShowResultsOrder();
                    Menu();
                    break;
                case 0:
                    helper.Quit();
                    break;
                default:
                    Console.WriteLine("Incorrect number");
                    Menu();
                    break;
            }
        }

        private void AddWord()
        {
            Console.ForegroundColor = ConsoleColor.DarkBlue;
            Console.WriteLine("Enter the word on English");
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.SetCursorPosition(1, Console.GetCursorPosition().Top);
            string? engWord = Console.ReadLine();
            Console.ForegroundColor = ConsoleColor.DarkBlue;
            Console.WriteLine("Enter the word on Russian");
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.SetCursorPosition(1, Console.GetCursorPosition().Top);
            string? rusWord = Console.ReadLine();
            Console.ForegroundColor = ConsoleColor.Red;

            if (engWord?.Length > 2 && rusWord?.Length > 2)
            {
                if (helper.Add(engWord, rusWord))
                {
                    Console.WriteLine("Success!");
                    Console.WriteLine();
                    Console.WriteLine();
                }
                else
                {
                    Console.WriteLine("Something went wrong!");
                    AddWord();
                }
            }
            else
            {
                Console.WriteLine("Incorrect words");
                AddWord();
            }

            Console.ForegroundColor = ConsoleColor.DarkMagenta;
        }

        private void Show()
        {
            var engWords = helper.EngWords;
            var rusWords = helper.RusWords;
            Console.ForegroundColor = ConsoleColor.DarkBlue;
            Console.WriteLine("    English                 Russian\n");
            for (int i = 0; i < engWords.Count; i++)
            {
                string eng = $"{i + 1}. {engWords[i]}";
                Console.WriteLine($"{eng, -25}{rusWords[i]}", Color.PowderBlue);
            }

            Console.WriteLine();
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.DarkMagenta;
        }

        private void ShowResultsOrder()
        {            
            var dict = new Dictionary<int, Tuple<string?, string?, int, int, double>>();
            for(int i = 0; i < helper.EngWords.Count; i++)
            {
                dict.Add(i, new (helper.EngWords[i], helper.RusWords[i], helper.CorrectResults[i], helper.TestAttempts[i], helper.Scores[i]));
            }
            var sortedList = dict.OrderBy(x => x.Value.Item5).ThenByDescending(x => x.Value.Item4).ToList();

            Console.ForegroundColor = ConsoleColor.DarkBlue;
            Console.WriteLine("   English                  Russian                              Total Result\n");
            for (int i = 0; i < sortedList.Count; i++)
            {
                string eng = $"{i + 1}. {sortedList[i].Value.Item1}";
                Console.WriteLine($"{eng,-25}{sortedList[i].Value.Item2,-40}{sortedList[i].Value.Item5 * 100}% " +
                    $"({sortedList[i].Value.Item3} out of {sortedList[i].Value.Item4})", Color.PowderBlue);
            }

            Console.WriteLine();
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.DarkMagenta;
        }

        private void Test(int lang)
        {
            //Creating variables, filling some data
            #region
            Random random = new Random();
            int length = helper.EngWords.Count;
            var words = new List<string?>();
            var testWords = new List<string?>();
            var mistakes = new List<string?>();
            var wrongAnswers = new List<string?>();
            var corrections = new List<string?>();
            var indexesOfWordsInTest = new List<int>();

            if (lang == 0)
            {
                words.AddRange(helper.EngWords);
                words.AddRange(helper.RusWords);
            }
            else
            {
                words.AddRange(helper.RusWords);
                words.AddRange(helper.EngWords);
            }
            #endregion

            //Console interface with user
            #region
            Console.ForegroundColor = ConsoleColor.DarkBlue;
            Console.WriteLine("1. Choose from latest words.\n2. Choose from all words.");

            Int32.TryParse(Console.ReadLine(), out int noveltySelector);
            if (noveltySelector != 1 && noveltySelector != 2)
            {
                Console.WriteLine("Incorrect!\n");
                Test(lang);
            }
            Console.WriteLine("How many words the test should include?");
            Int32.TryParse(Console.ReadLine(), out int amount);
            if (amount < 1)
            {
                Console.WriteLine("Incorrect!\n");
                Test(lang);
            }
            Console.WriteLine();

            if (amount > length) amount = length;
            #endregion

            //Forming the list of test words
            #region
            
            if (noveltySelector == 1)
            {
                testWords.AddRange(words.Skip(length - amount).Take(amount));
                testWords = testWords.OrderBy(x => random.Next()).ToList();
            }
            else
            {
                testWords.AddRange(words.Take(length));
                testWords = testWords.OrderBy(x => random.Next()).Take(amount).ToList();
            }
            foreach (var word in testWords)
            {
                indexesOfWordsInTest.Add(words.IndexOf(word));
            }
            var indexesOfCorrectWords = indexesOfWordsInTest.ToList();
            int tempCount = testWords.Count;
            for (int i = 0; i < tempCount; i++)
            {
                testWords.Add(words[indexesOfWordsInTest[i] + length]);
            }
            #endregion

            //Test and updating test results in Excel file
            #region
            int count = 0;
            for (int i = 0; i < amount; i++)
            {
                Console.WriteLine($" {testWords[i]}");
                Console.SetCursorPosition(1, Console.GetCursorPosition().Top);
                string? word = Console.ReadLine();
                if (word.Length > 2 && testWords[i + amount].ToLower().Contains(word.ToLower()))
                {
                    count++;
                }
                else
                {
                    mistakes.Add(testWords[i]);
                    wrongAnswers.Add(word);
                    corrections.Add(testWords[i + amount]);
                    indexesOfCorrectWords.Remove(indexesOfWordsInTest[i]);
                }
            }
            helper.WordsInTestUpdate(indexesOfWordsInTest, indexesOfCorrectWords);
            #endregion

            //Console display of results
            #region 
            Console.BackgroundColor = ConsoleColor.Red;
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine($"\nYour result is {count} out of {amount}\nIt is {(double)count / amount * 100} %\n"
                + $"You made mistakes in words:\n"
                + "   Word                       Correct translation              Your answer");
            Console.ForegroundColor = ConsoleColor.Yellow;
            for (int i = 0; i < mistakes.Count; i++)
            {
                string first = $"{i + 1}. {mistakes[i]}";
                Console.WriteLine($"{first,-30}{corrections[i],-33}{wrongAnswers[i]}");
            }
            Console.BackgroundColor = ConsoleColor.Cyan;
            Console.ForegroundColor = ConsoleColor.DarkMagenta;
            Console.WriteLine();
            #endregion
        }
    }
}
