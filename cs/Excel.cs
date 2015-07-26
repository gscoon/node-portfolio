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
        private bool isAppSet = false;
        private Excel.Application excelApplication;
        private Excel.Workbook workBook;

        public async Task<object> Invoke(dynamic input) {
            var functionName = (string)input.fn;
            //call the function with the file function
            return this.GetType().InvokeMember(functionName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase | BindingFlags.InvokeMethod, null, this, new object[] {input});
        }

        public object SetExcelObject(IDictionary<string, object> parameters){
            this.excelApplication = new Excel.Application();
            this.excelApplication.Visible = true;
            this.workBook = excelApplication.Workbooks.Add();

            //excelApplication.Calculation = XlCalculation.xlCalculationManual; // have to open workbook before setting calculation
            this.isAppSet = true;
            return true;
        }

        public object PopulateSheet(dynamic parameters){
            Excel.Worksheet workSheet = (Excel.Worksheet)this.workBook.Worksheets[1];
            workSheet.Range("A1").Value = "1";
            workSheet.Range("A2").Value = "2";
            workSheet.Range("A3").Value = "=a1+a2";
            return null;
        }

        public object PopulateDataSheet(dynamic parameters){
            var data = (object[])parameters.data;
            foreach(object[] dTable in data) {
                foreach(dynamic row in dTable){
                    foreach (var item in row){
                        return item.Key;
                        break;
                    }
                    break;
                }
            }
            return data;
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
