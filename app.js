var edge = require('edge');
var excel = edge.func({
    source: __dirname + "/cs/Excel.cs",
    references: [
        'System.Data.dll',
        __dirname + '/lib/NetOffice.dll',
        __dirname + '/lib/VBIDEApi.dll',
        __dirname + '/lib/ExcelApi.dll',
        __dirname + '/lib/OfficeApi.dll',
        __dirname + '/lib/exceldna/ExcelDna.Integration.dll'
    ],
    typeName: 'GSEXCEL.ExcelClass',
    methodName: 'Invoke' // This must be Func<object,Task<object>>
});

excel(3, function(error, result){
    console.log(result);
    console.log(error);
});
