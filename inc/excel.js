var edge = require('electron-edge2');
var baseDir = './';
var excelClass = function(){
    this.simpleMath = function(i, func){
        this.call(i, func);
    }

    this.call = edge.func({
        source: baseDir + "cs/Excel.cs",
        references: [
            'System.Data.dll',
            baseDir + 'lib/NetOffice.dll',
            baseDir + 'lib/VBIDEApi.dll',
            baseDir + 'lib/ExcelApi.dll',
            baseDir + 'lib/OfficeApi.dll',
            baseDir + 'lib/exceldna/ExcelDna.Integration.dll'
        ],
        typeName: 'GSEXCEL.ExcelClass',
        methodName: 'Invoke' // This must be Func<object,Task<object>>
    });
}

module.exports = new excelClass();
