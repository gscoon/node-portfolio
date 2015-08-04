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
    this.currentBookID = null;
    this.mousePosition = null;
    this.content = require(__dirname + '/content');
    this.request = require('request');
    this.async = require('async');
    this.excel = require(__dirname + '/excel');
    this.templateFolder = "C:/projects/node-portfolio/files/templates/";
    this.saveFolder = "C:/projects/node-portfolio/files/saved/";
    this.db = require(__dirname + '/db');
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

    self.currentBookID = null;
    if(hasWindow){
        self.w = remote.getCurrentWindow();
        $(self.setEventHandlers.bind(this)); // when page has been loaded
        setInterval(function(){
            self.mousePosition = atomScreen.getCursorScreenPoint();
        }, 5000);
    }

    self.excel.call({func:'SetExcelApplication'}, self.excelReturn);
}

appClass.prototype.createNewWorkbook = function(){
    console.log('createNewWorkbook');
    this.currentBookID = self.rand(5);
    this.excel.call({func: 'CreateNewWorkbook', wbID: this.currentBookID}, this.excelReturn);
}


appClass.prototype.close = function(){
    this.excel.call({func: 'CloseExcelApp'}, this.excelReturn);
    //self.w.close();
}

appClass.prototype.setDocumentProperties = function(){
    this.currentBookID = this.rand(5);

    this.excel.call({
        func: 'SetSheetProperties',
        wbID: this.currentBookID,
        prop:{
            val: 'np-config',
            key: 'Somesing'
        }
    }, this.excelReturn);
}

appClass.prototype.getDocumentProperties = function(){
    this.excel.call({
        func: 'GetSheetProperties',
        wbID: this.currentBookID,
        prop:{
            val: 'np-config',
            name: 'Somesing'
        }
    }, this.excelReturn);
}

appClass.prototype.openExistingFile = function(callback){
    console.log("openExistingFile");
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
            self.currentBookID = self.rand(5);
            self.excel.call({
                func: 'OpenExcelFile',
                wbID: self.currentBookID,
                openType: 'add',
                src: fileSrc
            }, callback);
        });
}

// process whatever template is provided
appClass.prototype.processTemplate = function(){
    var self = this;
    self.async.waterfall([
        //func 1
        // open an exisitng file / template
        function(callback){
            self.openExistingFile(function(err, oResults){
                console.log('Opened the template file, now callback...');
                callback(null, err, oResults);
            });
        },
        //func 2
        // return all sheets and look specifically for the config sheet
        function(err, oResults, callback){
            var o = oResults;
            o.func = 'GetAllSheets';
            self.excel.call(o, function(sErr, sResults){
                if(sErr != null){ console.log({error:sErr}); return false;}
                console.log(sResults);
                // look for config sheet
                var allSheets = sResults.results;
                var sheetMatchArr = allSheets.filter(function( obj ) {
                    return obj.name == config.configSheet.name;
                });
                if(sheetMatchArr.length > 0){
                    // it already exists, so just unhide it
                    var sArr = [{name: config.configSheet.name, 'type':'visible'}];
                    self.hideUnhideSheets(sArr, function(){
                        callback(null, allSheets);
                    });
                }

                var o = {func:'AddNewWorksheet', wbID: self.currentBookID, worksheetName: config.configSheet.name}
                self.excel.call(o, function(){
                    callback(null, allSheets);
                });
            });
        },
        //func 3
        // now you know the config sheet is there, finally do some processing
        function(allSheets, callback){

            // loop through each visible sheet and ask whether input or output
            // get date
            // get fields
            // figure out custom fields

            var sw = self.content.createNewSubwindow();
            self.content.showSubWindow(sw);
            self.content.displaySheetMenu(sw, allSheets);
        }
    ]); // end of async

}

appClass.prototype.returnTemplateConfig = function(){
    var tc = {
        importSheets: [],
        reportingSheet: [],
        pointers: []
    };
}

appClass.prototype.hideUnhideSheets = function(sheetArray, callback){
    if(typeof callback != 'function') callback = this.excelReturn;
    var o = {func: 'HideUnhideSheets', sheetArray: sheetArray, wbID: this.currentBookID};
    this.excel.call(o, callback);
}

appClass.prototype.showExcelRangePrompt = function(callback){
    var self = this;
    if(hasWindow) self.w.hide();
    self.showHideExcel(true);
    self.bringToFront(function(){
        self.excel.call({
            func: 'ShowExcelRangePrompt',
            wbID: self.currentBookID,
            promptText: 'Select fields'
        }, function(err, ret){
            self.showHideExcel(false);
            if(hasWindow) self.w.show();
            if(typeof callback == 'function') callback(ret);
        });
    });
    //ShowExcelRangePrompt
}

appClass.prototype.returnRangeValues = function(callback){
    var self = this;
    self.showExcelRangePrompt(function(ret){
        if(typeof ret == 'object'){
            if(typeof callback != 'function') callback = self.excelReturn;
            self.excel.call({
                func: 'ReturnSelectedRangeAsArray',
                wbID: self.currentBookID,
                sheetName: ret.results.sheet,
                rangeAddress: ret.results.address
            }, callback);
        }
    });
}
//ReturnSelectedRangeAsArray

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

appClass.prototype.getHighlightedText = function(){
    var self = this;

    var currentText = null;
    var maxInterval = 60, count = 0;
    var intervalID = setInterval(function(){

        self.excel.call({
            func: 'GetSelectedRangeAddress',
            wbID: self.currentBookID
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
