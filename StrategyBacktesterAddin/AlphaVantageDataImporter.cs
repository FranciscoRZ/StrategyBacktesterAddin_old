using System;
using System.Linq;

using System.Windows.Forms;
using System.Configuration;
using ThreeFourteen.AlphaVantage;
using Model = ThreeFourteen.AlphaVantage.Model;

namespace StrategyBacktesterAddin
{
    public static class AlphaVantageDataImporter
    {
        private static string _key;
        private static Model.TimeSeriesEntry[] _data;

        private static void ReadKey()
        {
            _key = ConfigurationManager.AppSettings.Get("AVKey");
        }

        public static Model.TimeSeriesEntry[] GetData()
        {
            return _data;
        }

        public static async void ImportData(string ticker, DateTime startDate, DateTime endDate)
        {
            ReadKey();
            var alphaVantage = new AlphaVantage(_key);
            var stockData = await alphaVantage.Stocks.Daily(ticker)
                                                     .SetOutputSize(OutputSize.Full)
                                                     .GetAsync();
            try
            {
                var queryData = from stock in stockData.Data
                                where (DateTime.Compare(stock.Timestamp.Date, endDate.Date) <= 0 &&
                                       DateTime.Compare(stock.Timestamp.Date, startDate.Date) >= 0)
                                select stock;
                 _data = queryData.Cast<Model.TimeSeriesEntry>().ToArray();   
            }
            catch (InvalidOperationException e)
            {
                MessageBox.Show(e.Message);
            }
        }
    }
}
