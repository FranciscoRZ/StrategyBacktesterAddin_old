using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Configuration;
using ThreeFourteen.AlphaVantage;

namespace StrategyBacktesterAddin
{
    public static class AlphaVantageDataImporter
    {
        private static string _key;

        private static void ReadKey()
        {
            _key = ConfigurationManager.AppSettings.Get("AVKey");
        }
        public static async void Import(string ticker, DateTime startDate, DateTime endDate)
        {
            ReadKey();
            var alphaVantage = new AlphaVantage(_key);
            var stockData = await alphaVantage.Stocks.Daily(ticker)
                                                     .SetOutputSize(OutputSize.Full)
                                                     .GetAsync();

        }
    }
}
