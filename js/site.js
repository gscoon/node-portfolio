var hasWindow = false;

if(hasWindow){
    var remote = require('remote');
    var dialog = remote.require('dialog');
    var atomScreen = remote.require('screen');
    $ = require('jquery');
}
//
require(__dirname + '/global');
config = require(__dirname + '/../config/config.json');

var appClass = function() {
    var self = this;
    this.mousePosition = null;
    this.content = require(__dirname + '/content');
    this.request = require('request');
    this.async = require('async');
    this.excel = require(__dirname + '/excel');
    this.templateFolder = "C:/projects/node-portfolio/files/templates/";
    this.saveFolder = "C:/projects/node-portfolio/files/processed/";
    this.db = require(__dirname + '/db');
    this.moment = require('moment');

    if(hasWindow){
        this.w = remote.getCurrentWindow();
        this.subWindow = [];
    }

    var __construct = function() {
        self.setUp();
    }()
};



appClass.prototype.setUp = function(){
    var self = this;
    if(hasWindow){
        self.w = remote.getCurrentWindow();
        $(self.setEventHandlers.bind(this)); // when page has been loaded
        setInterval(function(){
            self.mousePosition = atomScreen.getCursorScreenPoint();
        }, 5000);
    }

    //self.excel.call({func:'SetExcelApplication'}, self.excelReturn);
}

appClass.prototype.createNewWorkbook = function(){
    console.log('createNewWorkbook');
    var wbID = self.rand(5);
    this.excel.call({func: 'CreateNewWorkbook', wbID: wbID}, this.excelReturn);
}


appClass.prototype.close = function(){
    this.excel.call({func: 'CloseExcelApp'}, this.excelReturn);
    //self.w.close();
}

appClass.prototype.setDocumentProperties = function(){
    var wbID = this.rand(5);

    this.excel.call({
        func: 'SetSheetProperties',
        wbID: wbID,
        prop:{
            val: 'np-config',
            key: 'Somesing'
        }
    }, this.excelReturn);
}

appClass.prototype.getDocumentProperties = function(wbID){
    this.excel.call({
        func: 'GetSheetProperties',
        wbID: wbID,
        prop:{
            val: 'np-config',
            name: 'Somesing'
        }
    }, this.excelReturn);
}

appClass.prototype.getAllSheets = function(wbID, callback){
    this.excel.call({func: 'GetAllSheets', wbID: wbID}, callback);
}

appClass.prototype.showFileDialog = function(callback){
    console.log("showFileDialog");
    var self = this;
    // dialog is a module included above
    if(!callback) callback = self.excelReturn;
    dialog.showOpenDialog({
            title: "Select an Excel file",
            filters: [{name:'Excel File', extensions:['xls', 'xlsx', 'xlsm', 'xltx']}],
            properties: ['openFile']
        }, function(fileReturnArray){
            if(!fileReturnArray || fileReturnArray.length != 1) return false;
            var fileSrc = fileReturnArray[0];
            // do something now that you have this file
            self.openExistingFile(fileSrc, callback);
        });
}

appClass.prototype.openExistingFile = function(fileSrc, callback){
    var self = this;
    var wbID = self.rand(5);
    self.excel.call({
        func: 'OpenExcelFile',
        wbID: wbID,
        openType: 'add',
        src: fileSrc
    }, callback);
}

appClass.prototype.closeWorkbook = function(wbID, callback){
    this.excel.call({
        func: 'CloseWorkbook',
        wbID: wbID
    }, callback);
}

appClass.prototype.returnTemplateConfig = function(){
    var tc = {
        importSheets: [],
        reportingSheet: [],
        pointers: []
    };
}

appClass.prototype.hideUnhideSheets = function(wbID, sheetArray, callback){
    if(typeof callback != 'function') callback = this.excelReturn;
    var o = {func: 'HideUnhideSheets', sheetArray: sheetArray, wbID: wbID};
    this.excel.call(o, callback);
}

appClass.prototype.showExcelRangePrompt = function(wbID, callback){
    var self = this;
    if(hasWindow) self.w.hide();
    self.showHideExcel(true);
    self.bringToFront(function(){
        self.excel.call({
            func: 'ShowExcelRangePrompt',
            wbID: wbID,
            promptText: 'Select fields'
        }, function(err, ret){
            self.showHideExcel(false);
            if(hasWindow) self.w.show();
            if(typeof callback == 'function') callback(ret);
        });
    });
    //ShowExcelRangePrompt
}

appClass.prototype.findDataValuesByDim = function(p, callback){
    //sheet, rowAddr, colAddr
    var self = this;
    self.excel.call({
        func: 'ReturnDataValsByDims',
        wbID: p.bookID,
        sheetName: p.sheet,
        rowAddress: p.rowAddr,
        colAddress: p.colAddr
    }, callback);
}

appClass.prototype.returnRangeValues = function(wbID, sheet, address, callback){
    var self = this;
    self.excel.call({
        func: 'ReturnSelectedRangeAsArray',
        wbID: wbID,
        sheetName: sheet,
        rangeAddress: address
    }, callback);
}

appClass.prototype.returnNamedRangeValues = function(p, callback){
    var self = this;
    self.excel.call({
        func: 'ReturnNamedRangeAsArray',
        wbID: p.bookID,
        nrArray: p.nrArray
    }, callback);

}

appClass.prototype.showHideExcel = function(isVisible){
    this.excel.call({
        func: 'ShowHideExcelApp',
        isVisible: isVisible
    }, this.excelReturn);
}

appClass.prototype.bringToFront = function(callback){
    if(typeof callback != 'function') callback = this.excelReturn;
    this.excel.call({
        func: 'BringToFront'
    }, callback);
}

appClass.prototype.getHighlightedText = function(wbID){
    var self = this;

    var currentText = null;
    var maxInterval = 60, count = 0;
    var intervalID = setInterval(function(){

        self.excel.call({
            func: 'GetSelectedRangeAddress',
            wbID: wbID
        }, function(error, result){
            count++;
            console.log({result: result, error: error});
            if(count == maxInterval){
                console.log('max time reached');
                clearInterval(intervalID);
            }
        });
    }, 1000);
}

appClass.prototype.excelReturn = function(error, result){
    console.log({result: result, err: error});
}

appClass.prototype.setEventHandlers = function(){

    console.log('page loaded');
    var self = this;
    $('#x').on('click', function(){
        self.w.close();
    });

    $('#reload_template_button').on('click', this.analysisTemplateQueries.bind(this, 2));

    $('#inner_close_button').on('click', this.close.bind(this));

    $('#set_property_button').on('click', this.setDocumentProperties.bind(this));

    $('#get_property_button').on('click', this.getDocumentProperties.bind(this));

    $('#open_existing_button').on('click', this.openExistingFile);

    $('#highlighted_text_button').on('click', this.getHighlightedText.bind(this));

    $('#new_wb_button').on('click', this.createNewWorkbook.bind(this));

    $('#range_prompt_button').on('click', this.showExcelRangePrompt.bind(this));

    $('#bring_front_button').on('click', this.bringToFront.bind(this));

    $('#process_template_button').on('click', this.processTemplate.bind(this));
}

appClass.prototype.analysisTemplateQueries = function(mfiID){
    var self = this;
    console.log('analysisTemplateQueries');
    self.db.getAnalysisTemplateData(mfiID, function(err, data){
        var excelObj = {
            func: 'PopulateDataSheet',
            template: {
                name: 'Analysis Template.xltx',
                sheet: 'DB_LOAD',
                dataStart: [2,2], // cell B2
                fieldLabelStart: [1,2], // cell B1
                templatePath: self.templateFolder,
                savePath : self.saveFolder,
                pasteValSheets: ['DATA_FORMATTING'],
                nrPrefix:{
                    mapping: 'mapping',
                    push: 'push',
                    pull: 'pull',
                    data: 'data',
                    uploadField: 'field'
                }
            },
            wbID: self.rand(5),
            data: data,
            fieldMapping: self.returnFieldMapping(data)
        };

        self.excel.call(excelObj, self.excelReturn);
    });
}

appClass.prototype.returnFieldMapping = function(data){
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

appClass.prototype.rand = function(len, charset){
    charset = charset || "abcdefghijklmnopqrstuvwxyz0123456789";
    var str = "";
    for (var i=0; i < len; i++)
        str += charset.charAt(Math.floor(Math.random() * charset.length));

    return str;
}

site = new appClass();
