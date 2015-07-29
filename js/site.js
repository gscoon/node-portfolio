var remote = require('remote');
var dialog = remote.require('dialog');

$ = require('jquery');
require('./../config/global.js');

site = {
    config: require(__dirname + '/../config/config.json'),
    request: require('request'),
    async: require('async'),
    excel: require(__dirname + '/excel'),
    templateFolder: "C:/projects/node-portfolio/files/templates/",
    saveFolder: "C:/projects/node-portfolio/files/saved/",

    setUp: function(){
        var me = this;
        this.currentBookID = null;
        this.w = remote.getCurrentWindow();
        this.setEventHandlers();

        this.excel.call({func:'SetExcelApplication'}, function(error, result){
            var message = (typeof result === 'undefined')?'excel app NOT set': 'excel app set';
            console.log(message);
        });
    },

    rand: function(len, charset){
        charset = charset || "abcdefghijklmnopqrstuvwxyz0123456789";
        var str = "";
        for (var i=0; i < len; i++)
            str += charset.charAt(Math.floor(Math.random() * charset.length));

        return str;
    },

    close: function(){
        site.excel.call({func: 'CloseExcelApp'}, this.excelReturn);
        //site.w.close();
    },

    setDocumentProperties: function(){
        this.currentBookID = site.rand(5);
        site.excel.call({
            func: 'SetSheetProperties',
            wbID: this.currentBookID,
            prop:{
                val: 'np-config',
                key: 'Somesing'
            }
        }, this.excelReturn);
    },

    getDocumentProperties: function(){
        site.excel.call({
            func: 'GetSheetProperties',
            wbID: this.currentBookID,
            prop:{
                val: 'np-config',
                name: 'Somesing'
            }
        }, this.excelReturn);
    },

    inspectFile: function(){
        me = this;
        dialog.showOpenDialog({
                title: "Select an Excel file",
                filters: [{name:'Excel File', extensions:['xls', 'xlsx', 'xlsm', 'xltx']}],
                properties: ['openFile']
            }, function(fileReturnArray){
                if(!fileReturnArray || fileReturnArray.length != 1) return false;
                var fileSrc = fileReturnArray[0];
                // do something now that you have this file
                me.currentBookID = site.rand(5);
                site.excel.call({
                    func: 'OpenExcelFile',
                    wbID: me.currentBookID,
                    openType: 'open',
                    src: fileSrc
                }, me.excelReturn);
            })
    },

    excelReturn: function(error, result){
        console.log(error);
        console.log(result);
    },

    setEventHandlers: function(){
        var me = this;

        $('#x').on('click', function(){
            site.w.close();
        });

        $('#reload_template_button').on('click', function(){
            me.analysisTemplateQueries(2);
        });

        $('#inner_close_button').on('click', function(){
            me.close();
        });

        $('#set_property_button').on('click', function(){
            me.setDocumentProperties();
        });

        $('#get_property_button').on('click', function(){
            me.getDocumentProperties();
        });

        $('#file_inspector_button').on('click', function(){
            me.inspectFile();
        });
    }

};

site.db = require(__dirname + '/../config/db');

site.analysisTemplateQueries = function(mfiID){
    console.log('analysisTemplateQueries');
    site.db.getAnalysisTemplateData(mfiID, function(err, data){
        var excelObj = {
            func: 'PopulateDataSheet',
            template: {
                name: 'Analysis Template.xltx',
                sheet: 'DB_LOAD',
                dataStart: [2,2], // cell B2
                fieldLabelStart: [1,2], // cell B1
                templatePath: site.templateFolder,
                savePath : site.saveFolder,
                pasteValSheets: ['DATA_FORMATTING'],
                nrPrefix:{
                    mapping: 'mapping',
                    push: 'push',
                    pull: 'pull',
                    data: 'data'
                }
            },
            wbID: site.rand(5),
            data: data,
            fieldMapping: site.returnFieldMapping(data)
        };

        site.excel.call(excelObj, this.excelReturn);
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
