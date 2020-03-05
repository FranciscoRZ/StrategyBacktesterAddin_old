using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Runtime.InteropServices;
using CustomUI = ExcelDna.Integration.CustomUI;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;

namespace StrategyBacktesterAddin
{
    [ComVisible(true)]
    public class RibbonControler: CustomUI.ExcelRibbon
    {
        private string _ticker;
        private string _startDate;
        private string _endDate;

        public void OnImportDataPress(CustomUI.IRibbonControl control)
        {
            var application = (Excel.Application)ExcelDnaUtil.Application;
            object selection = application.Selection;
            if (selection is Excel.Range)
            {
                var selectedRange = (Excel.Range)selection;
                int firstCol = selectedRange.Column;
                int firstRow = selectedRange.Row;

                selectedRange.Worksheet.Cells[firstRow, firstCol] = _ticker;
                selectedRange.Worksheet.Cells[firstRow, firstCol + 1] = _startDate;
                selectedRange.Worksheet.Cells[firstRow, firstCol + 2] = _endDate;
            }
        }

        public void GetTickerValue(CustomUI.IRibbonControl control, string text)
        {
            if (!string.IsNullOrEmpty(text))
            {
                _ticker = text;
            }
        }

        public void GetStartDateValue(CustomUI.IRibbonControl control, string text)
        {
            if (!string.IsNullOrEmpty(text))
            {
                _startDate = text;
            }
        }

        public void GetEndDateValue(CustomUI.IRibbonControl control, string text)
        {
            if (!string.IsNullOrEmpty(text))
            {
                _endDate = text;
            }
        }
    }
}
