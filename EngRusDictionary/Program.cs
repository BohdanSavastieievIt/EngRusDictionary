using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dictionary
{
    public class Program
    {  
        static void Main(string[] args)
        {
            var dict = new EngRusDictionary();
            dict.ExcelHandler();
            foreach(var item in dict.dictionary)
            {
                Console.WriteLine(item);
            }
        }
    }
}
