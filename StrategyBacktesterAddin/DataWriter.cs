using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;
using Model = ThreeFourteen.AlphaVantage.Model;

    namespace StrategyBacktesterAddin
{
    public static class DataWriter
    {
        private static Excel.Application _xlApp = (Excel.Application)ExcelDnaUtil.Application;
        private static Excel.Workbook _cntWb = _xlApp.ActiveWorkbook;

        public static void WriteStockData(string ticker, Model.TimeSeriesEntry[] data)
        {
            Excel.Worksheet ws = _cntWb.Worksheets.Add(Type: Excel.XlSheetType.xlWorksheet);
            
            ws.Name = ticker + " data" ;
            ws.Range["A1", "B1"].Font.Bold = true;
            ws.Range["A1", "E1"].Value2 = new string[] { "Date", "Open", "Close", "High", "Low" };
        }
    }
}
