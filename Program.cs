using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;

namespace MatchTranslation
{
    class Program
    {
        /// <summary>
        /// first run your ==> ngx-translate-extract --input ./src --output ./src/assets/i18n/template.json --key-as-default-value --replace --format json && ngx-translate-extract --input ./src --o ./src/assets/i18n/de.json --o ./src/assets/i18n/en.json --o ./src/assets/i18n/es.json --o ./src/assets/i18n/fr.json --o ./src/assets/i18n/hr.json --o ./src/assets/i18n/it.json --o ./src/assets/i18n/ru.json --key-as-default-value --clean --format json"
        /// from git bash with "npm run extract"
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            string excelPath = args[0];
            string frontendi18nPath = args[1];

            string connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelPath + ";Extended Properties=Excel 12.0;";

            StringBuilder stbQuery = new StringBuilder();
            stbQuery.Append("SELECT * FROM [A1:M9999]");
            OleDbDataAdapter adp = new OleDbDataAdapter(stbQuery.ToString(), connStr);

            DataTable dt = new DataTable();
            adp.Fill(dt);

            var excelData = JsonConvert.DeserializeObject<List<Sheet1>>(JsonConvert.SerializeObject(dt));

            foreach (var lang in new List<string> { /*"en", */"de", "hr", "es", "it", "fr", "ru" })
            {
                var jsonData = JsonConvert.DeserializeObject<IDictionary<string, string>>(File.ReadAllText(frontendi18nPath + lang + ".json"));

                foreach (var keyValue in jsonData)
                {
                    var translated = excelData.FirstOrDefault(x => x != null && x.English != null && x.English.ToLower() == keyValue.Key.ToLower());
                    if (translated == null)
                    {
                        Console.WriteLine($"Not found '{keyValue.Key}' for language {lang}");
                    }

                    else
                        switch (lang)
                        {
                            case "de":
                                jsonData[keyValue.Key] = translated.German;
                                break;
                            case "hr":
                                jsonData[keyValue.Key] = translated.Croatian;
                                break;
                            case "es":
                                jsonData[keyValue.Key] = translated.Spanish;
                                break;
                            case "it":
                                jsonData[keyValue.Key] = translated.Italien;
                                break;
                            case "fr":
                                jsonData[keyValue.Key] = translated.French;
                                break;
                            case "ru":
                                jsonData[keyValue.Key] = translated.Russian;
                                break;
                            default:
                                break;
                        }
                }

                File.WriteAllText(frontendi18nPath + lang + ".json", JsonConvert.SerializeObject(jsonData, Formatting.Indented));
            }

            Console.ReadLine();

        }
    }

    public partial class Sheet1
    {
        [JsonProperty("English")]
        public string English { get; set; }

        [JsonProperty("German", NullValueHandling = NullValueHandling.Ignore)]
        public string German { get; set; }

        [JsonProperty("Croatian", NullValueHandling = NullValueHandling.Ignore)]
        public string Croatian { get; set; }

        [JsonProperty("Spanish", NullValueHandling = NullValueHandling.Ignore)]
        public string Spanish { get; set; }

        [JsonProperty("Italien", NullValueHandling = NullValueHandling.Ignore)]
        public string Italien { get; set; }

        [JsonProperty("French", NullValueHandling = NullValueHandling.Ignore)]
        public string French { get; set; }

        [JsonProperty("Russian", NullValueHandling = NullValueHandling.Ignore)]
        public string Russian { get; set; }
    }
}
