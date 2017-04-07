using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace CsvToExcel
{
    public class Settings
    {
        protected string fileLog;
        public string FILENAME;
        public string DEL_COLUMNS;
        public string COLOR_FILL_HEADER;
        public string COLOR_FILL_TEXT;
        public string PREFIX;
        public string COLS_SIZE180PX;
        public string COLS_SIZE061PX;
        public string COLS_SIZE103PX;
        public int SIZE_DEFAULT_IN_PX;
        public string DELIMITER;
        public int DEBUG;

        public Settings(string fileLog)
        {
            this.fileLog = fileLog;
        }
        public void Log(string msg)
        {
            if (this.DEBUG > 0)
            {
                File.AppendAllText(fileLog, msg);
                File.AppendAllText(fileLog, "\r\n");
            }            
        }
        public string USERPATH
        {
            get
            {
                System.Environment.GetEnvironmentVariable("USERPROFILE");
                string pathUser = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                return pathUser;
            }
        }
        public void Load()
        {
            Log("Load");
            System.Environment.GetEnvironmentVariable("USERPROFILE");

            string pathUser = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            Log("GetFolderPath" );
            if (Directory.Exists(pathUser + @"\AppData\Local\CSV-to-Excel\") == false)
            {
                Log("CreateDir ");
                Directory.CreateDirectory(pathUser + @"\AppData\Local\CSV-to-Excel\");
            }
            else
            {
                Log("Already CreateDir ");
            }

            AMS.Profile.Ini Ini = new AMS.Profile.Ini(pathUser + @"\AppData\Local\CSV-to-Excel\csv-to-excel.ini");

            if (File.Exists(pathUser + @"\AppData\Local\CSV-to-Excel\csv-to-excel.ini") == false)
            {
                Log("CreateLogDir set " );
                Ini.SetValue("COMMON", "FILENAME", "batch_analysis_export.utf-16.csv");
                Ini.SetValue("COMMON", "DEL_COLUMNS", "1,3,5,6,7,8");
                Ini.SetValue("COMMON", "COLOR_FILL_HEADER", "200-250-200");
                Ini.SetValue("COMMON", "COLOR_FILL_TEXT", "255-242-204");
                Ini.SetValue("COMMON", "PREFIX", "otchet_summ_");
                Ini.SetValue("COMMON", "COLS_SIZE180PX", "2");
                Ini.SetValue("COMMON", "COLS_SIZE103PX", "4,17");
                Ini.SetValue("COMMON", "COLS_SIZE061PX", "9,10,11,12,13,14,15,16");
                Ini.SetValue("COMMON", "SIZE_DEFAULT_IN_PX", 72);
                Ini.SetValue("COMMON", "DELIMITER", "'\t'");
                Ini.SetValue("COMMON", "DEBUG", 0);
            }
            else
            {
                Log("Alraedy CreateLogDir ");
            }
            this.FILENAME = Ini.GetValue("COMMON", "FILENAME", "sample.csv");
            this.DEL_COLUMNS = Ini.GetValue("COMMON", "DEL_COLUMNS", "1,3,5,6,7,8");
            this.COLOR_FILL_HEADER = Ini.GetValue("COMMON", "COLOR_FILL_HEADER", "200-250-200");
            this.COLOR_FILL_TEXT = Ini.GetValue("COMMON", "COLOR_FILL_TEXT", "255-242-204");
            this.PREFIX = Ini.GetValue("COMMON", "PREFIX", "otchet_summ_");
            this.COLS_SIZE180PX = Ini.GetValue("COMMON", "COLS_SIZE180PX", "2");
            this.COLS_SIZE103PX = Ini.GetValue("COMMON", "COLS_SIZE103PX", "4,17");
            this.COLS_SIZE061PX = Ini.GetValue("COMMON", "COLS_SIZE061PX", "9,10,11,12,13,14,15,16");
            this.SIZE_DEFAULT_IN_PX = Ini.GetValue("COMMON", "SIZE_DEFAULT_IN_PX", 72);
            this.DELIMITER = Ini.GetValue("COMMON", "DELIMITER", "'\t'");
            this.DEBUG = Ini.GetValue("COMMON", "DEBUG", 0);

            Log(this.FILENAME);
            Log(this.DEL_COLUMNS);
            Log(this.COLOR_FILL_HEADER);
            Log(this.COLOR_FILL_TEXT);
            Log(this.PREFIX);
        }

    }
}
