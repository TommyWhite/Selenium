using System.Configuration;

namespace SeleniumHelloWorld
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
