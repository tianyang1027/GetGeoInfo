using Newtonsoft.Json.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;
using System.Net;
using System.Threading.Tasks;

namespace GetGeoInfo
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            string path = @"C:\Users\v-yangtian\Desktop\门店信息汇总-20230208.xlsx";
            var fs = File.OpenRead(path);
            var wb = new XSSFWorkbook(fs);
            var sheet1 = wb.GetSheetAt(0);
            for (int i = 0; i <= sheet1.LastRowNum; i++)
            {
                IRow row = sheet1.GetRow(i);
                if (row != null && row.RowNum != 0)
                {
                    string address =string.Empty;
                    for (int j = 1; j < 6; j++)
                    {
                        address += row.GetCell(j).ToString().Trim();
                    }
                    var location = await GeocodingMap.GeLocation(address);
                    row.Cells[6].SetCellValue(location);
                }
            }
            // overwrite the workbook using a new stream
            using (FileStream fileStream = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                wb.Write(fileStream);
            }
            Console.WriteLine("Processed successfully");
            Console.ReadKey();
        }
    }


    public class Config
    {
        public static string Ak { get; set; } = "ab4877be5ea6b7f0e491ade12dbcefd3";
    }

    public class HttpRequestHelper
    {
        public static async Task<string> RequestUrl(string url)
        {
            string content = string.Empty;
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            {
                using (StreamReader sr = new StreamReader(response.GetResponseStream()))
                {
                    content = await sr.ReadToEndAsync();
                }
            }
            return content;
        }
    }

    public class GeocodingMap
    {
        public static async Task<string> GeLocation(string address)
        {
            //API documents：https://lbs.amap.com/api/webservice/guide/api/georegeo

            string url = $"https://restapi.amap.com/v3/geocode/geo?key={Config.Ak}&address={address}";
            string strJson = await HttpRequestHelper.RequestUrl(url);
            if (JObject.Parse(strJson)["infocode"].ToString() == "10000")
            {
                return JObject.Parse(strJson)["geocodes"][0]["location"].ToString();
            }
            return string.Empty;
        }
    }

}
