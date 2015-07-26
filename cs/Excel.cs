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

namespace GSEXCEL {
    public class ExcelClass {

        private int uniqueKey = 0;
        private bool isAppSet = false;
        private Excel.Application excelApplication;
        private Excel.Workbook workBook;
        private Excel.Worksheet ws1;
        object oOpt = System.Reflection.Missing.Value; //for optional arguments

        public async Task<object> Invoke(dynamic input) {
            var functionName = (string)input.fn;
            //call the function with the file function
            return this.GetType().InvokeMember(functionName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase | BindingFlags.InvokeMethod, null, this, new object[] {input});
        }

        public object SetExcelObject(IDictionary<string, object> parameters){
            this.excelApplication = new Excel.Application();
            this.excelApplication.Visible = true;
            this.workBook = excelApplication.Workbooks.Add();
            this.ws1 = (Excel.Worksheet)this.workBook.Worksheets[1];
            //excelApplication.Calculation = XlCalculation.xlCalculationManual; // have to open workbook before setting calculation
            this.isAppSet = true;
            return true;
        }

        public object PopulateSheet(dynamic parameters){

            this.ws1.Range("A1").Value = "1";
            this.ws1.Range("A2").Value = "2";
            this.ws1.Range("A3").Value = "=a1+a2";
            return null;
        }

        public object PopulateDataSheet(dynamic parameters){
            // loop through each data table
            int offset = 0;

            foreach(var tableRows in parameters.data) {
                int numberOfRows = tableRows.Length;
                // set the range

                //return (tableRows[0].Count());
                int numberOfColumns = this.getDynLength(tableRows[0]);

                // set dumping range
                var startCell = (Excel.Range) this.ws1.Cells[1, 1 + offset];
                var endCell = (Excel.Range) this.ws1.Cells[numberOfRows, numberOfColumns + offset];
                Excel.Range oRng = this.ws1.get_Range(startCell, endCell);

                // set the object that will populate the range
                object[,] outputArray = new object[numberOfRows, numberOfColumns];

                // loop through each row in the table
                int r = 0;
                foreach(var row in tableRows){
                    // loop through each item in the row
                    int c = 0;
                    foreach (var item in row){
                        outputArray.SetValue(item.Value, r, c);
                        c++;
                    }
                    r++;
                }

                oRng.set_Value(this.oOpt, outputArray);

                // add headerRow to dataTable

                offset = offset + numberOfColumns + 1;

            }

            return "Something";

        }

        private int getDynLength(dynamic dList){
            int i = 0;
            foreach(var d in dList){
                i++;
            }
            return i;
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
