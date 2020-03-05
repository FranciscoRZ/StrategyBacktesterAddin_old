using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
using ExcelDna;
using Model = ThreeFourteen.AlphaVantage.Model;

    namespace StrategyBacktesterAddin
{
    public class DataWriter
    {
        private Model.TimeSeriesEntry[] _data;

        public void Write()
        {

        }

        public DataWriter(Model.TimeSeriesEntry[] data)
        {
            this._data = data;
        }
    }
}
