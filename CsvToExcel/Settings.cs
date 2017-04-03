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
        public string FILENAME;
        public string DEL_COLUMNS;
        public string COLOR_FILL_HEADER;
        public string COLOR_FILL_TEXT;
        public string PREFIX;

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
            System.Environment.GetEnvironmentVariable("USERPROFILE");

            string pathUser = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);

            //if (Directory.Exists("pathUser") == false )
            //{
            //    throw new Exception("Отсутствует каталог пользователя");
            //}
            if (Directory.Exists(pathUser + @"\AppData\Local\CSV-to-Excel\") == false)
            {
                Directory.CreateDirectory(pathUser + @"\AppData\Local\CSV-to-Excel\");
            }

            AMS.Profile.Ini Ini = new AMS.Profile.Ini(pathUser + @"\AppData\Local\CSV-to-Excel\csv-to-excel.ini");

            if (File.Exists(pathUser + @"\AppData\Local\CSV-to-Excel\csv-to-excel.ini") == false)
            {
                Ini.SetValue("COMMON", "FILENAME", "sample.csv");
                Ini.SetValue("COMMON", "DEL_COLUMNS", "1,2,3");
                Ini.SetValue("COMMON", "COLOR_FILL_HEADER", "200-250-200");
                Ini.SetValue("COMMON", "COLOR_FILL_TEXT", "200-240-200");
                Ini.SetValue("COMMON", "PREFIX", "otchet_summ_");
            }
            this.FILENAME = Ini.GetValue("COMMON", "FILENAME", "sample.csv");
            this.DEL_COLUMNS = Ini.GetValue("COMMON", "DEL_COLUMNS", "1,2,3");
            this.COLOR_FILL_HEADER = Ini.GetValue("COMMON", "COLOR_FILL_HEADER", "200-250-200");
            this.COLOR_FILL_TEXT = Ini.GetValue("COMMON", "COLOR_FILL_TEXT", "200-240-200");
            this.PREFIX = Ini.GetValue("COMMON", "PREFIX", "otchet_summ_");

        }

    }
}
