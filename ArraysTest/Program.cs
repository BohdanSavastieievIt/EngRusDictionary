using System;

namespace Arrays
{
    class Program
    {
        public static void Main(string [] args)
        {
            int[,] a = { { 0, 1 }, { 2, 3 }, { 4, 5 }, { 6, 7 }, { 8, 9} };
            int[,] b = new int[3, 2];

            Array.Copy(a, 4, b, 0, 6);

            foreach (int i in b)
            {
                Console.WriteLine(i);
            }
            Random rnd = new Random();
            for (int i = b.Length / 2 - 1; i >= 1; i--)
            {
                int j = rnd.Next(i + 1);
                // обменять значения data[j] и data[i]
                var temp = b[j, 0];
                b[j, 0] = b[i, 0];
                b[i, 0] = temp;
                temp = b[j, 1];
                b[j, 1] = b[i, 1];
                b[i, 1] = temp;
            }
            foreach (int i in b)
            {
                Console.WriteLine(i);
            }
        }
    }
}