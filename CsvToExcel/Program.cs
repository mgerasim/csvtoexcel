using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CsvToExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string path = Directory.GetCurrentDirectory();
                string[] files = System.IO.Directory.GetFiles(path, "*.csv");
                foreach (var file in files)
                {
                    Console.WriteLine("Укажите имя файла (*.xlsx) для сохранения " + file);
                    string filenameExcel = Console.ReadLine();
                    ConvertCsvToExcel(file, filenameExcel);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        static void ConvertCsvToExcel(string filenameCsv, string filenameExcel)
        {

        }
    }
}
