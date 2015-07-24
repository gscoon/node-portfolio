using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using Office = NetOffice.OfficeApi;
using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;
using NetOffice.ExcelApi.Tools.Utils;
using ExcelDna.Integration;

namespace GSEXCEL
{
    public class ExcelClass
    {
        private int uniqueKey = 0;

        public async Task<object> Invoke(object input) {
            int v = (int)input;
            Excel.Application excelApplication = new Excel.Application();
            excelApplication.Visible = true;

            Excel.Workbook workBook = excelApplication.Workbooks.Add();
            excelApplication.Calculation = XlCalculation.xlCalculationManual; // have to open workbook before setting calculation
            Excel.Worksheet workSheet = (Excel.Worksheet)workBook.Worksheets[1];
            workSheet.Range("A1").Value = "1";
            workSheet.Range("A2").Value = "2";
            workSheet.Range("A3").Value = "=a1+a2";
            //workBook.SaveAs("Example01.xlsx");
            // close excel and dispose reference
            //excelApplication.Quit();
            //excelApplication.Dispose();
            //return Helper.AddSeven(v);
            return XlCalculation.xlCalculationManual;

        }
    }

    static class Helper
    {
        public static int AddSeven(int v)
        {
            return v + 7;
        }
    }

}
