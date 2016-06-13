using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataCollector
{
    public class GlobalVars
    {
        public const string URL_BING = @"https://www.bing.com/";
        public const string URL_GOV_US = @"https://usr.minjust.gov.ua/ua/freesearch/";
        public const string ID_SEARCH_INPUT = "query";
        public const string TAX_NUMBER = "3403201375";
        public const string FRAME_CLASS_NAME = "extiframe";
        public const string SCREEN_MAX = "--start-maximized";

        public string BaseDir { get; private set; } = AppDomain.CurrentDomain.BaseDirectory;
    }
}