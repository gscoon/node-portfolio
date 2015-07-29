var remote = require('remote');
$ = require('jquery');
require('./../config/global.js');

site = {
    config: require(__dirname + '/../config/config'),
    request: require('request'),
    async: require('async'),
    excel: require(__dirname + '/excel'),
    templateFolder: "C:/projects/node-portfolio/files/templates/",
    saveFolder: "C:/projects/node-portfolio/files/saved/",
    setUp: function(){
        this.w = remote.getCurrentWindow();

        $('#x').on('click', function(){
            site.w.close();
        });

        this.excel.call({func:'SetExcelApplication'}, function(error, result){
            var message = (typeof result === 'undefined')?'excel app NOT set': 'excel app set';
            console.log(message);
        });

        this.analysisTemplateQueries(2);
    },
    rand: function(len, charset){
        charset = charset || "abcdefghijklmnopqrstuvwxyz0123456789";
        var str = "";
        for (var i=0; i < len; i++)
            str += charset.charAt(Math.floor(Math.random() * charset.length));

        return str;
    }
};

site.db = require(__dirname + '/../config/db')

site.analysisTemplateQueries = function(mfiID){
    console.log('analysisTemplateQueries');
    site.db.getAnalysisTemplateData(mfiID, function(err, data){
        var excelObj = {
            func: 'PopulateDataSheet',
            template: {
                name: 'Analysis Template.xltx',
                sheet:'DB_LOAD',
                dataStart:[2,2], // cell B2
                fieldLabelStart:[1,2], // cell B1
                templatePath: site.templateFolder,
                savePath : site.saveFolder,
                id: site.rand(5),
                pasteValSheets:['DATA_FORMATTING'],
                nrPrefix:{
                    mapping:'mapping',
                    push: 'push',
                    pull: 'pull',
                    data: 'data'
                }
            },
            data: data,
            fieldMapping: site.returnFieldMapping(data)
        };

        console.log(excelObj);

        site.excel.call(excelObj, function(error, result){
            console.log(error?'getAnalysisTemplateData NOT set':'getAnalysisTemplateData is set');
            console.log(error || result);
        });
    });
}

site.returnFieldMapping = function(data){
    var fieldMapping = [];
    data.forEach(function(q, i){
        var fArray = [];
        q.fields.forEach(function(field){
            fArray.push(field.name);
        });
        fieldMapping.push(fArray);
    });
    return fieldMapping;
}

$(function(){
    site.setUp();
});
