using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JapaneseStocksSearchApp
{
    public class Company
    {
        public string Name { get; set; }
        public string Code { get; set; }
        public string Url { get; set; }

        public double DividendYield { get; set; }
    }
}
