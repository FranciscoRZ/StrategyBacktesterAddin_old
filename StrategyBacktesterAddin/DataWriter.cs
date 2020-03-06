using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
using Model = ThreeFourteen.AlphaVantage.Model;

namespace StrategyBacktesterAddin
{
    public static class DataWriter
    {
        public static void WriteStockData(string ticker, Model.TimeSeriesEntry[] data)
        {
            Excel.Worksheet ws = XLSingleton.Instance.XLWorkbook.Worksheets.Add(Type: Excel.XlSheetType.xlWorksheet);
            ws.Name = ticker + " data" ;
            ws.Range["A1", "E1"].Font.Bold = true;
            ws.Range["A1", "E1"].Value2 = new string[] { "Date", "Open", "Close", "High", "Low" };
            
        }
    }
}
