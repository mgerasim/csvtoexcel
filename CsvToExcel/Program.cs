using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
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
                Settings theSetting = new Settings();
                theSetting.Load();
                string path = Directory.GetCurrentDirectory();

                string filename = path + "\\" + theSetting.FILENAME;
                if (File.Exists(filename) == false)
                {
                    throw new Exception("Отсутствует файл " + filename);
                }
                Console.WriteLine("Укажите имя файла (*.xlsx): ");
                string filenameExcel = Console.ReadLine();
                filenameExcel = path + "\\" + theSetting.PREFIX + filenameExcel + ".xlsx";
                ConvertCsvToExcel(filename, filenameExcel, theSetting);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
            }

        }

        static void ConvertCsvToExcel(string filenameCsv, string filenameExcel, Settings theSettings)
        {
            string csvFileName = filenameCsv;
            string excelFileName = filenameExcel;

            string worksheetsName = "TEST";

            bool firstRowIsHeader = false;

            var format = new ExcelTextFormat();
            format.Delimiter = ';';
            format.EOL = "\r";              // DEFAULT IS "\r\n";
            // format.TextQualifier = '"';

            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFileName)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(worksheetsName);
                worksheet.Cells["A1"].LoadFromText(new FileInfo(csvFileName), format, OfficeOpenXml.Table.TableStyles.None, firstRowIsHeader);

                worksheet.Row(1).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                worksheet.Row(1).Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                worksheet.Row(1).Style.WrapText = true;
                worksheet.Row(1).Style.ShrinkToFit = true;
                worksheet.Row(1).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Row(1).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(
                    Convert.ToInt32(theSettings.COLOR_FILL_HEADER.Split(new char[] {'-'}, StringSplitOptions.None)[0]), 
                    Convert.ToInt32(theSettings.COLOR_FILL_HEADER.Split(new char[] {'-'}, StringSplitOptions.None)[1]), 
                    Convert.ToInt32(theSettings.COLOR_FILL_HEADER.Split(new char[] {'-'}, StringSplitOptions.None)[2]) ));

                worksheet.Cells[1, 2, 1, 2].Style.Font.Bold = true;
                worksheet.Cells[1, 17, 1, 17].Style.Font.Bold = true;

                for (int i = 1; i <= 27; i++)
                {
                    worksheet.Column(i).Width = 10;
                }

                worksheet.Column(2).Width = 25;
                worksheet.Column(4).Width = 14;
                worksheet.Column(9).Width = 8;
                worksheet.Column(10).Width = 10;
                worksheet.Column(11).Width = 8;
                worksheet.Column(12).Width = 8;
                worksheet.Column(13).Width = 8;
                worksheet.Column(14).Width = 7;
                worksheet.Column(15).Width = 8;
                worksheet.Column(16).Width = 8;
                worksheet.Column(17).Width = 14;

                foreach(var del in theSettings.DEL_COLUMNS.Split(new char [] {','}, StringSplitOptions.RemoveEmptyEntries).OrderByDescending(x=>Convert.ToInt32(x)))
                {
                    worksheet.DeleteColumn(Convert.ToInt32(del));
                }

                var rowCnt = worksheet.Dimension.End.Row;
                var colCnt = worksheet.Dimension.End.Column;

                string fileTxt = theSettings.USERPATH + "\\AppData\\Local\\CSV-to-Excel\\csv-to-excel.txt";

                if (File.Exists(fileTxt) == false)
                {
                    throw new Exception("Файл с текстовой информацией отсутствует " + fileTxt);
                }

                string readTxt = File.ReadAllText(fileTxt);

                worksheet.Cells[rowCnt + 2, 1, rowCnt + 2, 1].Value = readTxt;

                worksheet.Cells[2, 3, 11, 18].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[2, 3, 11, 18].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(
                    Convert.ToInt32(theSettings.COLOR_FILL_TEXT.Split(new char[] { '-' }, StringSplitOptions.None)[0]),
                    Convert.ToInt32(theSettings.COLOR_FILL_TEXT.Split(new char[] { '-' }, StringSplitOptions.None)[1]),
                    Convert.ToInt32(theSettings.COLOR_FILL_TEXT.Split(new char[] { '-' }, StringSplitOptions.None)[2])));

                package.Save();
            }
        }
    }
}
