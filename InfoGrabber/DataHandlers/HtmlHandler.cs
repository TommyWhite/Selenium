using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace EmployeeInfoGrabber
{
    public class HtmlHandler
    {
        private string ReadEntireFile(string file)
        {
            string content;
            using (StreamReader reader = new StreamReader(file))
            {
                content = reader.ReadToEnd();
            }

            return content;
        }

        private List<string> RetrievePersonData(string file)
        {
            string content = ReadEntireFile(file);
            List<string> personData = new List<string>();
            string[] parsedData = ParseHtml(content)
                .ToList()
                .Where((x, n) => (n & 1) == 1)
                .ToArray();
            personData.Add(file);
            personData.AddRange(parsedData);

            return personData;
        }

        public string[,] ReadReportData(string htmlReportsPath)
        {
            var files = Directory.GetFiles(htmlReportsPath, "*.html");
            string[,] data = new string[files.Length, 17];
            for (int i = 0; i < files.Length; i++)
            {

                List<string> personData = RetrievePersonData(files[i]);
                for (int j = 0; j < personData.Count; j++)
                {
                    data[i, j] = personData[j].ReplaceMany(new[] { "<b>", "</b>", "<br>" }, "");
                }
            }

            return data;
        }
        
        //TODO: Edit access modifier to 'private' after testing.
        public string[] ParseHtml(string data)
        {
            const string TagPattern = @"<td>(.*?)<\/td>";
            var occurences = Regex.Matches(data, TagPattern);
            string[] matches = new string[occurences.Count];
            for (int i = 0; i < occurences.Count; i++)
            {
                matches[i] = occurences[i].Groups[1].Value.ToString();
            }

            return matches;
        }
    }
}