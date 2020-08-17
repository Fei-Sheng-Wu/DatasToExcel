using System;
using System.IO;

namespace DatasToExcel.Test
{
    class Program
    {
        static void Main()
        {
            string[,] datas = new string[,]
            {
                { "Name", "Country", "Age", "Career" },
                { "Helen", "U.S.", "21", "Police" },
                { "Jucia", "Canada", "34", "Dancer" },
                { "Erik", "Canada", "13", "Student" },
                { "Bob", "British", "26", "Business person" },
                { "Nancy", "Russia", "64", "Fisherman" },
            };

            for (int i = 0; i < datas.GetLength(0); i++)
            {
                for (int j = 0; j < datas.GetLength(1); j++)
                {
                    Console.Write(datas[i, j] + "\t");
                }
                Console.Write("\n");
            }

            Console.WriteLine();
            Console.WriteLine("Please enter the output Excel file path:");

            string filename = Console.ReadLine();

            Console.WriteLine();

            try
            {
                datas.GenerateExcel(filename, true);

                Console.WriteLine("Generate Excel file successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error:");
                Console.WriteLine(ex.Message);
            }

            Console.WriteLine();
            Console.WriteLine("Press any key to exit.");

            Console.ReadKey();
        }
    }
}
