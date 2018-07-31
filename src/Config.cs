using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToWord
{

    public class Config
    {
        public ObservableCollection<ExcelToWordConfigItem> Mapping { get; set; } = new ObservableCollection<ExcelToWordConfigItem>();

        public string ExcelPathfilename { get; set; }
        public string WordPathfilename { get; set; }
        public string OutputFolder { get; set; }

        internal Global global { get { return Global.Instance; } }
        internal static string Pathfilename { get { return Global.Instance.ConfigPathfilename; } }

        internal static Config Load()
        {
            Config config = null;

            if (!File.Exists(Pathfilename))
            {
                config = new Config();
                config.Save();
            }
            else
            {
                config = Newtonsoft.Json.JsonConvert.DeserializeObject<Config>(File.ReadAllText(Pathfilename));
            }

            return config;
        }

        public void Save()
        {
            File.WriteAllText(Pathfilename, Newtonsoft.Json.JsonConvert.SerializeObject(this, Newtonsoft.Json.Formatting.Indented));
        }

    }

    public class ExcelToWordConfigItem
    {
        public string ColumnName { get; set; }
        public string TokenToReplace { get; set; }
    }

}
