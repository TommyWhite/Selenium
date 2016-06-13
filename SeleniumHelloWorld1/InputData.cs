using System.Configuration;

namespace DataCollector
{
    public class InputData : ConfigurationSection
    {
        [ConfigurationProperty("excelWithTaxNumbers", IsRequired = true)]
        public string ExcelFile
        {
            get
            {
                return (string)this["excelWithTaxNumbers"];
            }
        }
    }
}
