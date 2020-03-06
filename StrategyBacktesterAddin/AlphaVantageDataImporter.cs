using System;
using System.Linq;

using System.Windows.Forms;
using System.Configuration;
using ThreeFourteen.AlphaVantage;
using Model = ThreeFourteen.AlphaVantage.Model;

namespace StrategyBacktesterAddin
{
    public sealed class AlphaVantageDataImporter
    {
        private string _key;
        private Model.TimeSeriesEntry[] _data;

        private void ReadKey()
        {
            this._key = ConfigurationManager.AppSettings.Get("AVKey");
        }

        public Model.TimeSeriesEntry[] GetData()
        {
            return this._data;
        }

        public async void ImportData(string ticker, DateTime startDate, DateTime endDate)
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
                 this._data = queryData.Cast<Model.TimeSeriesEntry>().ToArray();   
            }
            catch (InvalidOperationException e)
            {
                MessageBox.Show(e.Message);
            }
        }

        /// <summary>
        /// Private constructor
        /// </summary>
        private AlphaVantageDataImporter()
        {
        }

        public static AlphaVantageDataImporter Instance { get { return NestedAVDataImporter.instance; } }
        
        private class NestedAVDataImporter
        {
            /// <summary>
            /// Explicit static constructor to tell C# compiler not to mark as before as beforefieldinit
            /// </summary>
            static NestedAVDataImporter()
            {
            }

            /// <summary>
            /// Instanciation of AlphaVantageDataImporter happens on first call to instance, and never again.
            /// </summary>
            internal static readonly AlphaVantageDataImporter instance = new AlphaVantageDataImporter();
        }

    }
}
