using System;

using System.Windows.Forms;
using System.Globalization;
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
        private DateTime _startDate;
        private DateTime _endDate;
        private CultureInfo _culture = new CultureInfo("fr-FR");

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

            AlphaVantageDataImporter.Import(_ticker, _startDate, _endDate);
        }

        public void GetTickerValue(CustomUI.IRibbonControl control, string text)
        {
            if (!string.IsNullOrEmpty(text))
            {
                this._ticker = text;
            }
        }

        public void GetStartDateValue(CustomUI.IRibbonControl control, string text)
        {   
            try
            {
                this._startDate = DateTime.Parse(text, this._culture);
            }
            catch (FormatException e)
            {
                if (!string.IsNullOrEmpty(text))
                {
                    MessageBox.Show(e.Message);
                }
            }
        }

        public void GetEndDateValue(CustomUI.IRibbonControl control, string text)
        {
            try
            {
                this._endDate = DateTime.Parse(text, this._culture);
            }
            catch (FormatException e)
            {
                if (!string.IsNullOrEmpty(text))
                {
                    MessageBox.Show(e.Message);
                }
            }
        }
    }
}
