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
        private Random random;
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
            Console.WriteLine("0. Enter '0' to quit the program");
            Console.WriteLine();
            
            var isCorrect = Int32.TryParse(Console.ReadLine(), out int choice);
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

            if (engWord.Length > 2 && rusWord.Length > 2)
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



        private void Test(int lang)
        {
            int lang2 = lang == 0 ? 1 : 0;
            int length = helper.EngWords.Count;
            string [,] allWords = new string[length, 2];
            for(int i = 0; i < length; i++)
            {
                allWords[i, 0] = helper.EngWords[i];
                allWords[i, 1] = helper.RusWords[i];
            }

            int count = 0;
            var mistakes = new List<string?>();
            var corrections = new List<string?>();

            random = new Random();

            Console.ForegroundColor = ConsoleColor.DarkBlue;
            Console.WriteLine("How many words the test should include?");
            var isCorrect = Int32.TryParse(Console.ReadLine(), out int amount);
            Console.WriteLine();
            if (isCorrect)
            {
                if (amount > length) amount = length;
                string[,] testWords = new string[amount, 2];
                Array.Copy(allWords, (length - amount) * 2, testWords, 0, amount * 2);

                for (int i = amount - 1; i >= 1; i--)
                {
                    int j = random.Next(i + 1);
                    var temp = testWords[j, 0];
                    testWords[j, 0] = testWords[i, 0];
                    testWords[i, 0] = temp;
                    temp = testWords[j, 1];
                    testWords[j, 1] = testWords[i, 1];
                    testWords[i, 1] = temp;
                }

                for (int i = 0; i < amount; i++)
                {
                    Console.WriteLine($" {testWords[i, lang]}");
                    Console.SetCursorPosition(1, Console.GetCursorPosition().Top);
                    string? word = Console.ReadLine();
                    if (testWords[i, lang2].ToLower().Contains(word.ToLower()))
                    {
                        count++;
                    }
                    else
                    {
                        mistakes.Add(testWords[i, lang]);
                        corrections.Add(testWords[i, lang2]);
                    }
                }
            }
            else
            {
                Console.WriteLine("Incorrect!\n");
                Test(lang);
            }

            Console.BackgroundColor = ConsoleColor.Red;
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine($"\nYour result is {count} out of {amount}\nIt is {(double)count / amount * 100} %\n" +
                $"You made mistakes in words:\n");
            for (int i = 0; i < mistakes.Count; i++)
            {
                string first = $"{i + 1}. {mistakes[i]}";
                Console.WriteLine($"{first,-25}{corrections[i]}");
            }
            Console.BackgroundColor = ConsoleColor.Cyan;
            Console.ForegroundColor = ConsoleColor.DarkMagenta;
            Console.WriteLine();
        }
    }
}
