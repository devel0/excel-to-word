using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToWord
{
    public class Global
    {

        #region instance

        static Global instance;
        static object lckInstance = new object();

        public static Global Instance
        {
            get
            {
                if (instance == null)
                {
                    lock (lckInstance)
                    {
                        if (instance != null) return instance;
                        instance = new Global();
                    }
                }
                return instance;
            }
        }

        #endregion
        
        public string AppData
        {
            get
            {
                var path = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "ExcelToWord");
                if (!Directory.Exists(path)) Directory.CreateDirectory(path);

                return path;
            }
        }

        public string ConfigPathfilename
        {
            get
            {
                var file = Path.Combine(AppData, "config.json");

                return file;
            }
        }

        Config config;
        public Config Config
        {
            get
            {
                if (config == null)
                {
                    config = Config.Load();
                }
                return config;
            }
        }

    }
}
