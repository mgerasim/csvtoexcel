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
        static string fileLog;
        static Settings theSetting;

        static void Log(string msg)
        {
            if (fileLog.Length > 0 && (theSetting != null && theSetting.DEBUG > 2))
            {
                File.AppendAllText(fileLog, msg);
                File.AppendAllText(fileLog, "\r\n");
            }
        }
        static void Main(string[] args)
        {
            try
            {
                string path = Directory.GetCurrentDirectory();
                fileLog = path + "\\csv-to-excel.log";
                
                theSetting = new Settings(fileLog);
                theSetting.Load();

                string filename = path + "\\" + theSetting.FILENAME;
                if (File.Exists(filename) == false)
                {
                    Log("Not FileExists" );
                    throw new Exception("Отсутствует файл " );
                }
                else {
                    Log("FileExists Ok");
                }
                Console.WriteLine("Укажите имя файла (*.xlsx): ");
                string filenameExcel = Console.ReadLine();
                filenameExcel = path + "\\" + theSetting.PREFIX + filenameExcel + ".xlsx";
                Log("Excel File Name" );
                ConvertCsvToExcel(filename, filenameExcel, theSetting);
                Log("Delete file");
                File.Delete(filename);

            }
            catch (Exception ex)
            {
                Log(ex.Message);
                Log(ex.Source);
                Log(ex.StackTrace);
                Console.WriteLine(ex.Message);
                Console.ReadKey();
            }

        }

        static void ConvertCsvToExcel(string filenameCsv, string filenameExcel, Settings theSettings)
        {
            Log("ConvertCsvToExcel");
            //Log("DateTime Now");
            //Log(DateTime.Now.ToShortDateString());
            //if (DateTime.Now > (new DateTime(2017,05,05))) {
            //    Log("Protect");
            //    return;
            //} else
            //{
            //    Log("Protect Go");
            //}
            string csvFileName = filenameCsv;
            string excelFileName = filenameExcel;

            string worksheetsName = "TEST";

            bool firstRowIsHeader = false;

            var format = new ExcelTextFormat();

            Log("theSettings.DELIMITER");
            Log(theSettings.DELIMITER);
            string DELIMITER = theSettings.DELIMITER;
            DELIMITER = DELIMITER.Trim(new char[] { '\'' });
            try
            {
                format.Delimiter = DELIMITER[0];
            }
            catch
            {
                Log("error delimiter");
                format.Delimiter = ';';
            }
            
            format.EOL = "\n";              // DEFAULT IS "\r\n";
            format.TextQualifier = '"';

            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFileName)))
            {
                Log("ExcelPackage");
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(worksheetsName);
                worksheet.Cells["A1"].LoadFromText(new FileInfo(csvFileName), format, OfficeOpenXml.Table.TableStyles.None, firstRowIsHeader);

                var colCnt = worksheet.Dimension.End.Column;

                worksheet.Row(1).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                worksheet.Row(1).Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                worksheet.Row(1).Style.WrapText = true;
                worksheet.Row(1).Style.ShrinkToFit = true;
                worksheet.Cells[1, 1, 1, colCnt].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;

                Log("COLOR_FILL_HEADER");
                Log(theSettings.COLOR_FILL_HEADER);
                foreach (var token in theSettings.COLOR_FILL_HEADER.Split(new char[] { '-' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    Log(token);
                }

                worksheet.Cells[1, 1, 1, colCnt].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(
                    Convert.ToInt32(theSettings.COLOR_FILL_HEADER.Split(new char[] {'-'}, StringSplitOptions.RemoveEmptyEntries)[0]), 
                    Convert.ToInt32(theSettings.COLOR_FILL_HEADER.Split(new char[] {'-'}, StringSplitOptions.RemoveEmptyEntries)[1]), 
                    Convert.ToInt32(theSettings.COLOR_FILL_HEADER.Split(new char[] {'-'}, StringSplitOptions.RemoveEmptyEntries)[2]) ));

                worksheet.Cells[1, 2, 1, 2].Style.Font.Bold = true;
                worksheet.Cells[1, 17, 1, 17].Style.Font.Bold = true;

                for (int i = 1; i <= 27; i++)
                {
                    worksheet.Column(i).Width = (theSettings.SIZE_DEFAULT_IN_PX + 5 - 12) / 7d + 1; ;
                }

                foreach (var col in theSettings.COLS_SIZE180PX.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    try
                    {
                        worksheet.Column(Convert.ToInt32(col)).Width = (185 - 12) / 7d + 1;
                    }
                    catch (Exception ex)
                    {
                        Log("COLS_SIZE180PX");
                        Log(ex.Message);
                        Log(col);
                        Log(theSettings.COLS_SIZE180PX);
                        Log(ex.StackTrace);
                    }                    
                }

                foreach (var col in theSettings.COLS_SIZE103PX.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    try
                    {
                        worksheet.Column(Convert.ToInt32(col)).Width = (108 - 12) / 7d + 1;
                    }
                    catch (Exception ex)
                    {
                        Log("COLS_SIZE103PX");
                        Log(ex.Message);
                        Log(col);
                        Log(theSettings.COLS_SIZE103PX);
                        Log(ex.StackTrace);
                    }
                }


                foreach (var col in theSettings.COLS_SIZE061PX.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    try
                    {
                        worksheet.Column(Convert.ToInt32(col)).Width = (66 - 12) / 7d + 1;
                    }
                    catch (Exception ex)
                    {
                        Log("COLS_SIZE061PX");
                        Log(ex.Message);
                        Log(col);
                        Log(theSettings.COLS_SIZE061PX);
                        Log(ex.StackTrace);
                    }
                }
                
                Log("del_column");
                Log(theSettings.DEL_COLUMNS);
                foreach (var del in theSettings.DEL_COLUMNS.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).OrderByDescending(x => Convert.ToInt32(x)))
                {
                    var colCnts = worksheet.Dimension.End.Column;
                    Log(del + " - " + colCnts);
                    worksheet.DeleteColumn(Convert.ToInt32(del));
                    Log(Convert.ToInt32(del).ToString());
                    Log("del ok");
                }

                var rowCnt = worksheet.Dimension.End.Row;

                string fileTxt = theSettings.USERPATH + "\\AppData\\Local\\CSV-to-Excel\\csv-to-excel.txt";

                if (File.Exists(fileTxt) == false)
                {
                    throw new Exception("Файл с текстовой информацией отсутствует " + fileTxt);
                }

                Log("COLOR_FILL_TEXT");
                Log(theSettings.COLOR_FILL_TEXT);
                foreach (var token in theSettings.COLOR_FILL_TEXT.Split(new char[] { '-' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    Log(token);
                }

                System.Drawing.Color theColor = Color.FromArgb(
                    Convert.ToInt32(theSettings.COLOR_FILL_TEXT.Split(new char[] { '-' }, StringSplitOptions.None)[0]),
                    Convert.ToInt32(theSettings.COLOR_FILL_TEXT.Split(new char[] { '-' }, StringSplitOptions.None)[1]),
                    Convert.ToInt32(theSettings.COLOR_FILL_TEXT.Split(new char[] { '-' }, StringSplitOptions.None)[2]));



                var readTxt = File.ReadAllLines(fileTxt);

                foreach(var line in readTxt)
                {
                    Uri uriResult;
                    bool result = Uri.TryCreate(line, UriKind.Absolute, out uriResult)
                        && (uriResult.Scheme == Uri.UriSchemeHttp || uriResult.Scheme == Uri.UriSchemeHttps);
                                        
                    if (result)
                    {
                        var er = new ExcelHyperLink(line);
                        worksheet.Cells[rowCnt + 2, 1, rowCnt + 2, 1].Hyperlink = new Uri(line);
                        worksheet.Cells[rowCnt + 2, 1, rowCnt + 2, 1].Value = line;

                        try
                        {
                            var namedStyle = worksheet.Workbook.Styles.CreateNamedStyle("HyperLink");
                            namedStyle.Style.Font.UnderLine = true;
                            namedStyle.Style.Font.Color.SetColor(Color.Blue);
                        }
                        catch(Exception ex)
                        {
                            Log(ex.Message);
                        }
                        
                        worksheet.Cells[rowCnt + 2, 1, rowCnt + 2, 1].StyleName = "HyperLink";
                    }
                    else
                    {
                        worksheet.Cells[rowCnt + 2, 1, rowCnt + 2, 1].Value = line;
                    }

                    worksheet.Cells[rowCnt + 2, 1, rowCnt + 2, 11].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[rowCnt + 2, 1, rowCnt + 2, 11].Style.Fill.BackgroundColor.SetColor(theColor);
                    rowCnt++;
                }

                


                
                package.Save();
            }
        }
    }
}
