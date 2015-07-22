var win32ole = require('win32ole');

var xl = win32ole.client.Dispatch('Excel.Application');
xl.Visible = true;
var book = xl.Workbooks.Add();
var sheet = book.Worksheets(1);
var result = book.SaveAs('testfileutf8.xls');
xl.Workbooks.Close();
xl.Quit();
