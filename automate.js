require('./js/site');

var auto = new function(){
    var inProgress = false;

    var startTemplateProcess = function(){
        var row = null;
        console.log('started ' + site.moment().format("YYYY-MM-DD HH:mm:ss"));
        site.async.waterfall([
            // 1. check refresh table
            function(callback){
                site.db.checkRefreshTable(callback);
            },
            // 2. make sure everything is set correctly
            function(ret, fields, callback){
                if(ret.length == 0)
                    return callback("No refresh found");

                row = ret[0];

                if(row.config == null || !row.config.isJSONString())
                    return callback("Config not JSON");

                row.configObj = JSON.parse(row.config);
                callback(null);
            },
            // 3. get and process sql queries
            function(callback){
                var dArray = row.configObj.reportTemplate.dataNeeded;
                site.db.getReportingQueries(dArray, row.mfi_id, callback);
            },
            // 5. add data to object and update refresh table
            function(data, callback){
                row.data = data;
                site.db.updateRefreshTable(row.refresh_id, 'processing', callback);
            },
            // 6.
            function(ret, fields, callback){
                row.bookID = site.rand(5);
                processReportingTemplate(row, callback);
            }
        ], function(err, ret){
            // FInal callback
            setTimeout(startTemplateProcess, 1000 * 5);
            console.log('Full process completed');
            console.log({ret:ret, err:err});

            if(err !== null && err !== 'undefined') // not running
                return;

            // success
            site.closeWorkbook(row.bookID, emptyFunc);
            site.db.updateRefreshTable(row.refresh_id, 'completed', function(){
                console.log('ended ' + site.moment().format("YYYY-MM-DD HH:mm:ss"));
            });
        });
    }


    var processReportingTemplate = function(rObj, callback){
        console.log('processReportingTemplate');
        var mfiID = rObj.mfi_id;
        var excelObj = {
            func: 'PopulateDataSheet',
            template: {
                dataSheet: 'DB_LOAD',
                dataStart: [2,2], // cell B2
                fieldLabelStart: [1,2], // cell B1
                src: rObj.src,
                savePath: rObj.dest,
                saveName: rObj.mfi_id.toString(),
                pasteValSheets: rObj.configObj.reportTemplate.pasteValsSheet,
                nrPrefix: {
                    mapping: 'mapping',
                    push: 'push',
                    pull: 'pull',
                    data: 'data'
                }
            },
            wbID: rObj.bookID,
            data: rObj.data,
            fieldMapping: site.returnFieldMapping(rObj.data)
        };
        site.excel.call(excelObj, function(err, ret){
            console.log("reporting template complete");
            //do aggregated financials
            processUploadSheetSheet(rObj, callback);
        });
    }


    var processUploadSheetSheet = function(rObj, finalCallback){
        console.log("processUploadSheetSheet");
        var mfiID = rObj.mfi_id;
        var sheetConfig = rObj.configObj;
        site.async.each(sheetConfig.importTemplate, function(t, callback){
            // yep, async within an async
            site.async.waterfall([
                // 4a.
                function(cback){
                    // get the named ranges of dates, fields, and custom dimensions if available
                    var nrArray = [t.date.nr, t.field.nr];
                    if(t.customDim.nr != null)
                        nrArray.push(t.customField.nr);
                    site.returnNamedRangeValues({nrArray: nrArray, bookID: rObj.bookID}, cback);
                },
                // 4b.
                function(ret, cback){
                    var nrObj = ret.results;
                    // get field ids for each field based on matched aliases
                    mapTemplateFields(t.sheetID, nrObj[t.field.nr].rng, function(idArray){
                        nrObj.fieldIDArray = idArray;
                        cback(null, nrObj);
                    });
                },
                // 4c find the data already
                function(nrObj, cback){
                    var columnAddr = nrObj[t.field.nr].value;
                    var rowAddr = nrObj[t.date.nr].value;
                    // pull all data values into a 2 dim array [column][row]
                    var p = {sheet: t.sheet, rowAddr: rowAddr, colAddr: columnAddr, bookID: rObj.bookID};
                    site.findDataValuesByDim(p,  function(err, ret){
                        // loop through each data column
                        site.async.each(ret.results, function(colData, cb){
                            var i = ret.results.indexOf(colData);
                            var fsDate = nrObj[t.date.nr].rng[i];
                            // handle custom dimension
                            var customDim = (t.customDim.nr == null) ? t.customDim.value : nrObj[t.customDim.nr].rng[i];
                            if(typeof t.customDim.convert[customDim] !== 'undefined')
                                customDim = t.customDim.convert[customDim];

                            // create new field set
                            addFieldSet({
                                mfiID: mfiID,
                                sheetID: t.sheetID,
                                fsDate: fsDate,
                                customDim: customDim,
                                vals: colData,
                                ids: nrObj.fieldIDArray
                            }, cb);
                        }, cback);
                    });
                }
            ], callback); // end of waterfall
        }, finalCallback); // end of loop throught input sheets
    }

    function addFieldSet(fObj, callback){
        site.db.addReportingFieldSet(fObj, function(err, results){
            if(err != null) return false;
            //console.log(results);

            var fieldsetID = results.insertId;
            var insertArray = [];
            for(var j = 0; j < fObj.vals.length; j++){
                var currentVal = fObj.vals[j];
                if(typeof currentVal != 'number')
                    currentVal = 0;
                var currentFieldID = fObj.ids[j];
                insertArray.push([currentFieldID, currentVal, fieldsetID]);
            }
            site.db.insertDataValues(insertArray, callback);
        });
    }

    function mapTemplateFields(sheetID, fieldArray, callback){
        site.db.getReportingLabelsBySheet(sheetID, function(err, results){
            var idArray = [];
            for(f = 0; f < fieldArray.length; f++){
                var i = idArray.length;
                idArray[i] = null;
                for(r = 0; r < results.length; r++){
                    if(results[r].alias == fieldArray[f])
                        idArray[i] = results[r].field_label_id
                }
            }
            callback(idArray);
        });
    }

    var emptyFunc = function(){}

    var __construct = function() {
        startTemplateProcess();
    }()

}
