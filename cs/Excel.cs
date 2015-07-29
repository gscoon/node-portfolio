using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Dynamic;
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
        private Excel.Application excelApp;
        private Dictionary<string, object> wb = new Dictionary<string, object>();
        object oOpt = System.Reflection.Missing.Value; //for optional arguments

        public async Task<object> Invoke(dynamic input) {
            var functionName = (string)input.func;
            //call the function with the file function
            return this.GetType().InvokeMember(functionName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase | BindingFlags.InvokeMethod, null, this, new object[] {input});
        }

        public object SetExcelApplication(dynamic parameters){
            this.excelApp = new Excel.Application();
            this.excelApp.Visible = true;
            this.excelApp.DisplayAlerts = false;

            //
            this.isAppSet = true;
            return true;
        }

        public object PopulateDataSheet(dynamic p){
            // set workbook and worksheet object
            this.wb[p.template.id] = (Excel.Workbook) this.excelApp.Workbooks.Add(p.template.templatePath + p.template.name);
            var dataSheet = (Excel.Worksheet)this.wb[p.template.id].Worksheets[p.template.sheet];

            this.excelApp.ScreenUpdating = false;
            this.excelApp.Calculation = XlCalculation.xlCalculationManual; // have to open workbook before setting calculation

            // unhide hidden sheets but remember them
            List<int> hiddenSheetList = this.unhideHiddenSheets(this.wb[p.template.id].Worksheets);

            // column spacing between each dump
            int offset = 0;

            // loop through each data table
            foreach(var dataTable in p.data) {
                int numberOfRows = dataTable.results.Length;
                int numberOfColumns = this.ReturnColumnsCount(dataTable.results[0]);

                // set dumping range
                var startCell = (Excel.Range) dataSheet.Cells[p.template.dataStart[0], p.template.dataStart[1] + offset];
                var endCell = (Excel.Range) dataSheet.Cells[p.template.dataStart[0] + numberOfRows - 1, numberOfColumns + p.template.dataStart[1] + offset - 1];


                // determine range and set named range
                Excel.Range oRng = dataSheet.get_Range(startCell, endCell);
                var nrName = (string) p.template.nrPrefix.data + "." + dataTable.name; // name of named range
                //dataSheet.Names.Item(nrName, Type.Missing, Type.Missing).Delete();
                this.wb[p.template.id].Names.Add(nrName, oRng);

                // set the object that will populate the range
                object[,] outputArray = new object[numberOfRows, numberOfColumns];

                // loop through each row in the table
                int r = 0;
                foreach(var row in dataTable.results){
                    int c = 0;
                    // loop through each item in the row
                    foreach (var item in row){
                        // if this is the first row, handle field named ranges
                        if(r == 0){
                            var fieldNR = p.template.nrPrefix.mapping + "." + dataTable.name + "." + item.Key;
                            var fieldCell = dataSheet.Cells[p.template.fieldLabelStart[0], p.template.fieldLabelStart[1] + offset + c];
                            fieldCell.Value = c;
                            try {
                                this.wb[p.template.id].Names.Add(fieldNR, fieldCell);
                            }
                            catch (Exception ex){
                                //return fieldNR + " | " + GetExceptionDetails(ex);
                            }
                        }
                        outputArray.SetValue(item.Value, r, c);
                        c++;  // next column
                    }
                    r++;  // next row
                }
                oRng.set_Value(this.oOpt, outputArray);
                offset = offset + numberOfColumns + 1;
            }

            this.excelApp.Calculation = XlCalculation.xlCalculationAutomatic;
            this.PasteSheetValues(this.wb[p.template.id], p.template.pasteValSheets);

            // handle pushes and pulls within template
            this.HandleTemplatePointers(this.wb[p.template.id], p.template.nrPrefix);
            this.rehideHiddenSheets(this.wb[p.template.id].Worksheets, hiddenSheetList);

            this.excelApp.ScreenUpdating = true;

            var savePath = ((string)p.template.savePath + p.template.id + ".xlsx").Replace("/", "\\");
            try {
                this.wb[p.template.id].SaveAs(@savePath);
            }
            catch (Exception ex){
                return GetExceptionDetails(ex);
            }
            return savePath;

        }

        private void PasteSheetValues(Excel.Workbook wb, dynamic snArr){
            foreach(var sn in snArr){
                var ws = (Excel.Worksheet) wb.Sheets[sn];
                ws.Cells.Copy();
                ws.Cells.PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
            }
        }

        private void HandleTemplatePointers(Excel.Workbook wb, dynamic nrPrefix) {
            var nameList = new Dictionary<string, List<Dictionary<string, dynamic>>>();
            nameList["push"] = new List<Dictionary<string, dynamic>>();
            nameList["pull"] = new List<Dictionary<string, dynamic>>();
            char delim = '.';

            foreach (Excel.Name name in wb.Names){
                string[] nameSplit = name.Name.Split(delim);
                if(nameSplit.Length > 1){
                    var d = new Dictionary<string, dynamic>()
                    {
                        {"id", nameSplit[1]},
                        {"nr", name}
                    };

                    if(nameSplit[0] == nrPrefix.push)
                        nameList["push"].Add(d);
                    else if(nameSplit[0] == nrPrefix.pull)
                        nameList["pull"].Add(d);
                }
            }
            foreach(var pull in nameList["pull"]){
                foreach(var push in nameList["push"]){
                    if((string)pull["id"] == (string)push["id"]){
                        var pushNR = (Excel.Name) push["nr"];
                        var pullNR = (Excel.Name) pull["nr"];

                        var pushSheet = (Excel.Worksheet)wb.Sheets[pushNR.RefersToRange.Cells.Worksheet.Name];
                        var pullSheet = (Excel.Worksheet)wb.Sheets[pullNR.RefersToRange.Cells.Worksheet.Name];

                        var rPush = pushSheet.get_Range(pushNR.Name);
                        var rPull = pullSheet.get_Range(pullNR.Name);

                        rPull.Value = rPush.Value;

                        break;
                    }
                }
            }

        }

        private int ReturnColumnsCount(dynamic dList){
            int i = 0;
            foreach(var d in dList){
                i++;
            }
            return i;
        }

        public static string GetExceptionDetails(Exception exception) {
            PropertyInfo[] properties = exception.GetType()
                                    .GetProperties();
            List<string> fields = new List<string>();
            foreach(PropertyInfo property in properties) {
                object value = property.GetValue(exception, null);
                fields.Add(String.Format(
                                 "{0} = {1}",
                                 property.Name,
                                 value != null ? value.ToString() : String.Empty
                ));
            }
            return String.Join("\n", fields.ToArray());
        }

        private List<int> unhideHiddenSheets(Excel.Sheets sheets) {
            List<int> indexes = new List<int>();
            int index = 1;
            foreach (Excel.Worksheet sheet in sheets){
                if (sheet.Visible == Excel.Enums.XlSheetVisibility.xlSheetHidden){
                    sheet.Visible = Excel.Enums.XlSheetVisibility.xlSheetVisible;
                    indexes.Add(index);
                }
                index++;
            }
            return indexes;
        }

        private void rehideHiddenSheets(Excel.Sheets sheets, List<int> indexes) {

            foreach (int index in indexes){
                int i = 1;
                foreach (Excel.Worksheet sheet in sheets){
                    if (index == i){
                        sheet.Visible = Excel.Enums.XlSheetVisibility.xlSheetHidden;
                    }
                    i++;
                }
            }
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
