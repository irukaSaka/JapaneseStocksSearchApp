using AngleSharp;
using AngleSharp.Dom;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;

namespace JapaneseStocksSearchApp
{
    class Program
    {
        private static double BASELINE_EQUITY_RATIO = 50;
        private static double BASELINE_OPERATING_PROFIT_MARGIN = 10;

        private static int PAGE_NUM_RANKING = 15;
        
        private static string URL_RANKING = "https://finance.yahoo.co.jp/stocks/ranking/dividendYield?market=all&term=daily";
        //private static string URL_INDEX = "https://www.nikkei.com/markets/kabu/nidxprice/";

        static void Main(string[] args)
        {
            List<string> rowList = new List<string>();
            Dictionary<string, Company> code2name = GetCode2Name();         
            foreach (KeyValuePair<string, Company> kv in code2name)
            {
                bool operatingProfitMargin = false;
                bool equityRatio = false;

                string url = $"https://www.nikkei.com/nkd/company/kessan/?scode={kv.Key}&ba=1";
                IDocument document = BrowsingContext.New(Configuration.Default.WithDefaultLoader()).OpenAsync(url).Result;
                IHtmlCollection<IElement> elements = document.QuerySelectorAll("tr");
                foreach (IElement element in elements)
                {
                    try
                    {
                        string label = element.QuerySelector("th").Text().Trim();
                        if (label.Contains("営業利益率"))
                        {
                            operatingProfitMargin = IsAboveBaseline(element, BASELINE_OPERATING_PROFIT_MARGIN);
                        }
                        else if (label.Contains("自己資本比率"))
                        {
                            equityRatio = IsAboveBaseline(element, BASELINE_EQUITY_RATIO);
                        }
                    }
                    catch (Exception e)
                    {
                        //Console.WriteLine(e);
                    }
                }

                if (operatingProfitMargin && equityRatio)
                {
                    double pbr = GetPBR(kv.Value.Url);
                    if (0.5 < pbr && pbr < 1)
                    {
                        rowList.Add($"{kv.Value.Name},{kv.Value.Code},{pbr},{kv.Value.Url},{kv.Value.DividendYield}");
                    }
                }

                Thread.Sleep(1000);
            }

            using (StreamWriter writer = new StreamWriter($@"C:\Users\ruka\Desktop\Result_{DateTime.Now.ToString("yyyyMMdd")}.csv", false, Encoding.UTF8))
            {
                writer.WriteLine("Name,Code,PBR,Url,DividendYield"); // header line

                foreach (string row in rowList)
                {
                    writer.WriteLine(row);
                }
            }

        }

        private static double GetPBR(string url)
        {
            IDocument document = BrowsingContext.New(Configuration.Default.WithDefaultLoader()).OpenAsync(url).Result;
            IHtmlCollection<IElement> elements = document.QuerySelectorAll("#referenc li");
            foreach (IElement element in elements)
            {
                string label = element.QuerySelector("._30FbHvZq").Text();
                if (label.Contains("PBR"))
                {
                    string val = element.QuerySelector("._11kV6f2G").Text();
                    return String2Double(val);
                }
            }
            return 100;
        }

        private static bool IsAboveBaseline(IElement element, double baseLine)
        {
            bool flag = true;
            IHtmlCollection<IElement> values = element.QuerySelectorAll("td");

            double total = 0;
            foreach (IElement value in values)
            {
                double val = String2Double(value.Text());
                total += val;

                flag = baseLine < val;
            }
            return baseLine < total / 5 && flag;
        }

        private static double String2Double(string str)
        {
            if (str == "--")
            {
                return 0;
            }

            double number;   
            str = str.Replace("－", "-").Trim();

            if (double.TryParse(str, out number))
            {
                return number;
            }
            else
            {
                Console.WriteLine("Invalid string");
                Console.WriteLine(str);
            }
            return number;
        }

        private static Dictionary<string, Company> GetCode2Name()
        {
            Dictionary<string, Company> dic = new Dictionary<string, Company>();
            for (int i = 1; i <= PAGE_NUM_RANKING; i++)
            {
                IDocument document = BrowsingContext.New(Configuration.Default.WithDefaultLoader()).OpenAsync($"{URL_RANKING}&page={i}").Result;
                IHtmlCollection <IElement> elements = document.QuerySelectorAll("._1GwpkGwB");
                foreach (IElement element in elements)
                {
                    string code = element.QuerySelector(".vv_mrYM6").Text();
                    string dividendYield = element.QuerySelector("._3rXWJKZF").Text();

                    Company company = new Company();
                    company.Url = element.QuerySelector("a").GetAttribute("href");
                    company.Name = element.QuerySelector("a").Text();
                    company.Code = code;
                    company.DividendYield = String2Double(dividendYield);

                    dic.Add(code, company);
                }
                Thread.Sleep(1000);
            }
            return dic;
        }
    }
}
