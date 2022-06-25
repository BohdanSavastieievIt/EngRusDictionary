using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyEngRusDictionaryService
{
    public class Program
    {
        static void Main(string[] args)
        {
            var dict = new DictionaryService();
            dict.StartProgram();
        }
    }
}