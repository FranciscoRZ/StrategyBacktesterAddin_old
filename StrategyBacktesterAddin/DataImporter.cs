using System;
using System.Collections.Generic;
using System.Linq;

using System.Configuration;
using System.Windows.Forms;
using ServiceStack;

namespace StrategyBacktesterAddin
{
    public sealed class DataImporter
    {
        private string _key;        
        private List<AlphaVantageData> _data;
        private static readonly DataImporter instance = new DataImporter();

        public List<AlphaVantageData> GetData()
        {
            return this._data;
        }

        public void ImportData(string symbol, DateTime startDate, DateTime endDate)
        {
            ReadConfig();
            string apiKey = this._key;
            string request = $"https://www.alphavantage.co/query?function=TIME_SERIES_DAILY&symbol={symbol}&apikey={apiKey}&outputsize=full&datatype=csv";
            try
            {
                List<AlphaVantageData> dailyPrices = request.GetStringFromUrl()
                                                            .FromCsv<List<AlphaVantageData>>();
                var queryData = from stock in dailyPrices
                                where (DateTime.Compare(stock.Timestamp.Date, endDate.Date) <= 0 &&
                                       DateTime.Compare(stock.Timestamp.Date, startDate.Date) >= 0)
                                select stock;
                this._data = queryData.Cast<AlphaVantageData>().ToList();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        private void ReadConfig()
        {
            this._key = ConfigurationManager.AppSettings.Get("AVKey2");

        }

        private DataImporter()
        {
        }

        static DataImporter()
        {
        }

        public static DataImporter Instance
        {
            get
            {
                return instance;
            }
        }
    }
}
