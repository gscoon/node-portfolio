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
        private Dictionary<string, Excel.Workbook> wb = new Dictionary<string, Excel.Workbook>();
        object oMissing = System.Reflection.Missing.Value; //for optional arguments

        public async Task<object> Invoke(dynamic input) {
            var functionName = (string)input.func;
            //call the function with the file function
            return this.GetType().InvokeMember(functionName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase | BindingFlags.InvokeMethod, null, this, new object[] {input});
        }

        public object SetExcelApplication(dynamic p){
            this.excelApp = new Excel.Application();
            this.excelApp.Visible = false;
            this.excelApp.DisplayAlerts = false;
            this.isAppSet = true;
            p.success = true;
            return p;
        }

        private void CheckOnApp(dynamic p){
            if(!this.isAppSet){
                var setSuccess = this.SetExcelApplication(p);
            }
        }

        public object ShowHideExcelApp(dynamic p){
            excelApp.Visible = p.isVisible;
            p.success = true;
            return p;
        }

        public object GetVisibleSheets(dynamic p){
            var sheetList = new List<string>();
            foreach(var ws in wb[p.wbID].Worksheets){
                if(ws.Visible == XlSheetVisibility.xlSheetVisible){
                    sheetList.Add(ws.Name);
                }
            }
            p.results = sheetList;
            return p;
        }

        public object GetAllSheets(dynamic p){
            var sheetList = new List<Dictionary<string, dynamic>>();
            foreach(var ws in wb[p.wbID].Worksheets){
                Dictionary<string, dynamic> dic = new Dictionary<string, dynamic>();
                dic.Add("name", (string) ws.Name);
                dic.Add("isVisible", (bool) false);
                if(ws.Visible == Excel.Enums.XlSheetVisibility.xlSheetVisible)
                    dic["isVisible"] = true;

                sheetList.Add(dic);
            }
            p.results = sheetList;
            return p;
        }

        public object PopulateDataSheet(dynamic p){
            this.CheckOnApp(p);
            // set workbook and worksheet object
            this.wb[p.wbID] = (Excel.Workbook) this.excelApp.Workbooks.Add(p.template.src);
            var dataSheet = (Excel.Worksheet)this.wb[p.wbID].Worksheets[p.template.dataSheet];
            this.excelApp.ScreenUpdating = false;
            this.excelApp.Calculation = XlCalculation.xlCalculationManual; // have to open workbook before setting calculation
            // unhide hidden sheets but remember them
            List<int> hiddenSheetList = this.unhideHiddenSheets(this.wb[p.wbID].Worksheets);

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
                var nrName = (string) p.template.nrPrefix.data + "." + dataTable.query_id; // name of named range
                //dataSheet.Names.Item(nrName, Type.Missing, Type.Missing).Delete();

                this.wb[p.wbID].Names.Add(nrName, oRng);

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
                            var fieldNR = p.template.nrPrefix.mapping + "." + dataTable.query_id + "." + item.Key;
                            var fieldCell = dataSheet.Cells[p.template.fieldLabelStart[0], p.template.fieldLabelStart[1] + offset + c];
                            fieldCell.Value = c;
                            try {
                                this.wb[p.wbID].Names.Add(fieldNR, fieldCell);
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
                oRng.set_Value(this.oMissing, outputArray);
                offset = offset + numberOfColumns + 1;
            }

            this.excelApp.Calculation = XlCalculation.xlCalculationAutomatic;
            this.PasteSheetValues(this.wb[p.wbID], p.template.pasteValSheets);

            // handle pushes and pulls within template
            this.HandleTemplatePointers(this.wb[p.wbID], p.template.nrPrefix);
            this.rehideHiddenSheets(this.wb[p.wbID].Worksheets, hiddenSheetList);

            this.excelApp.ScreenUpdating = true;

            var savePath = ((string)p.template.savePath + p.template.saveName + ".xlsx").Replace("/", "\\");
            try {
                Console.WriteLine(savePath);
                this.wb[p.wbID].SaveAs(@savePath);
                p.success = true;
            }
            catch (Exception ex){
                p.success = false;
                p.fail = GetExceptionDetails(ex);
            }
            return p;
        }

        public object OpenExcelFile(dynamic p){
            this.CheckOnApp(p);

            if(p.openType == "open"){
                this.wb.Add(p.wbID, this.excelApp.Workbooks.Open(@p.src));
            }
            else{
                this.wb.Add(p.wbID, this.excelApp.Workbooks.Add(@p.src));
            }

            var pp = GetAllSheets(p);
            p.results = pp.results;
            return p;
        }


        public object SetSheetProperties(dynamic p){
            this.CheckOnApp(p);

            this.wb[p.wbID] = (Excel.Workbook) this.excelApp.Workbooks.Add();

            object oDocCustomProps = this.wb[p.wbID].CustomDocumentProperties;
            object[] oArgs = {p.prop.key, false, Office.Enums.MsoDocProperties.msoPropertyTypeString, p.prop.val};

            oDocCustomProps.GetType().InvokeMember("Add", BindingFlags.Default |
                BindingFlags.InvokeMethod, null,
                oDocCustomProps, oArgs);

            p.results = "success";
            return p;
        }


        public object GetSheetProperties(dynamic p){
            if(!this.isAppSet)
                return "Excel app not set";

            dynamic oDocCustomProps = this.wb[p.wbID].CustomDocumentProperties;
            p.results = new Dictionary<string, string>();
            foreach (dynamic prop in oDocCustomProps){
                p.results.Add(prop.Name, prop.Value);
            }

            return p;
        }

        public object GetWorkbookNames(dynamic p){
            try {
                var nameList = new Dictionary<string, Dictionary<string, string>>();
                foreach (Excel.Name nr in wb[p.wbID].Names){
                    foreach(var targetName in p.names){
                        if(nr.Name == targetName){
                            nameList[nr.Name] = new Dictionary<string, string>();
                            nameList[nr.Name]["value"] = nr.Value;
                            nameList[nr.Name]["comment"] = nr.Comment;
                        }
                    }
                }
                p.results = nameList;
            }
            catch (Exception ex){
                p.ex = GetExceptionDetails(ex);
            }

            return p;
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

        public object AddNewWorksheet(dynamic p){
            Excel.Worksheet ws = wb[p.wbID].Worksheets.Add();
            ws.Name = p.worksheetName;
            p.success = true;
            return p;
        }

        // sent from javascript
        public object HideUnhideSheets(dynamic p){
            int i = 0;
            foreach(Excel.Worksheet sheet in wb[p.wbID].Sheets){
                foreach(var givenSheet in p.sheetArray){
                    if(sheet.Name == givenSheet.name){
                        i++;
                        if(givenSheet.type == "visible")
                            sheet.Visible = Excel.Enums.XlSheetVisibility.xlSheetVisible;
                        else if(givenSheet.type == "hidden")
                            sheet.Visible = Excel.Enums.XlSheetVisibility.xlSheetHidden;
                        else if(givenSheet.type == "veryhidden")
                            sheet.Visible = Excel.Enums.XlSheetVisibility.xlSheetVeryHidden;
                        else
                            i--;
                    }
                }
            }
            p.successCount = i;
            return p;
        }

        public object CreateNewWorkbook(dynamic p){
            this.CheckOnApp(p);
            this.wb[p.wbID] = (Excel.Workbook) this.excelApp.Workbooks.Add();
            p.success = true;
            return p;
        }

        public object CloseWorkbook(dynamic p){
            wb[p.wbID].Save();
            wb[p.wbID].Close(0);
            p.success = true;
            return p;
        }

        public object SaveAndCloseWorkbook(dynamic p){
            this.wb[p.wbID].SaveAs(@p.src);
            wb[p.wbID].Close(0);
            p.success = true;
            return p;
        }

        public object CloseExcelApp(dynamic p){
            foreach(var w in this.wb){
                w.Value.Close(0);
            }
            this.excelApp.Quit();
            this.isAppSet = false;
            return p;
        }

        public object GetSelectedRangeAddress(dynamic p){
            wb[p.wbID].Activate();
            Excel.Application app = excelApp;
            var rng = (Excel.Range)app.Selection;
            p.results = rng.get_AddressLocal(rng.Rows.Count, rng.Columns.Count, XlReferenceStyle.xlA1, oMissing, oMissing);
            return p;
        }

        public object BringToFront(dynamic p){
            // works when there's a workbook
            excelApp.ActiveWindow.Activate();
            p.success = true;
            return p;
        }

        public object ShowExcelRangePrompt(dynamic p){
            var rng = excelApp.InputBox(
                                 p.promptText,
                                 "Field Selection",
                                 oMissing,
                                 oMissing,
                                 oMissing,
                                 oMissing,
                                 oMissing, 8);
             if(rng != null){
                 string rngAddress = rng.get_AddressLocal(false, false, XlReferenceStyle.xlA1, oMissing, oMissing);
                 p.results = new Dictionary<string, string>();
                 p.results.Add("address", rngAddress);
                 p.results.Add("sheet", rng.Worksheet.Name);
             }

             return p;
        }

        public object ReturnSelectedRangeAsArray(dynamic p){
            Excel.Worksheet ws = wb[p.wbID].Sheets[p.sheetName];
            Excel.Range rng = ws.get_Range(p.rangeAddress);
            object rangeVal = rng.Value;
            p.results = rangeVal;
            return p;
        }

        public object ReturnNamedRangeAsArray(dynamic p){
            p.results = new Dictionary<string, Dictionary<string, dynamic>>();
            foreach (var pnr in p.nrArray){
                foreach (Excel.Name nr in wb[p.wbID].Names){
                    if(nr.Name == pnr){
                        Dictionary<string, dynamic> n = new Dictionary<string, dynamic>();
                        n.Add("value", nr.Value);
                        n.Add("comment", nr.Comment);
                        n.Add("name", nr.Name);
                        Excel.Range rng = nr.RefersToRange;
                        n.Add("rng", (object) rng.Value2);
                        p.results.Add(nr.Name, n);
                    }
                }

            }

            return p;
        }

        // reutrns data array based on fields and dates
        public object ReturnDataValsByDims(dynamic p){
            Excel.Worksheet ws = wb[p.wbID].Sheets[p.sheetName];

            // usally the date
            Excel.Range rowRange = ws.get_Range(p.rowAddress);
            // usually the field
            Excel.Range colRange = ws.get_Range(p.colAddress);

            Excel.Range start = ws.Cells[colRange.Row, rowRange.Column];

            Excel.Range end = ws.Cells[(colRange.Row + colRange.Rows.Count - 1), (rowRange.Column + rowRange.Columns.Count - 1)];

            Excel.Range valueRange = ws.get_Range(start, end);
            object[,] valArray = (object[,]) valueRange.get_Value(oMissing);
            List<List<object>> list = new List<List<object>>();
            int size1 = valArray.GetLength(1);
            //Console.WriteLine("Size1: " + size1.ToString());
            int size0 = valArray.GetLength(0);
            //Console.WriteLine("Size0: " + size0.ToString());

            // Loop through each row in within each column
            for(int j=0; j<size1; j++){
                List<object> subList = new List<object>();
                    for(int i=0; i<size0; i++){
                    subList.Add(valArray[i+1, j+1]);
                    // excel array starts at 1
                }
                list.Add(subList);
            }

            p.results = list;

            string rngAddress = valueRange.get_AddressLocal(false, false, XlReferenceStyle.xlA1, oMissing, oMissing);
            p.valueAddr = rngAddress;
            //Console.WriteLine("Rank: " + valArray.Rank.ToString());
            return p;
        }

    }

}
