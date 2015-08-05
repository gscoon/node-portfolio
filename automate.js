require('./js/site');

var mfiID = 2;
site.async.waterfall([
    //1. open workbook
    function(callback){
        var fileSrc = 'C:/projects/node-portfolio/files/templates/filled/Access Bank_2406478.xlsx';
        site.openExistingFile(fileSrc, callback);
    },
    //2. find config sheet and config variable
    function(ret, callback){
        var sheetArray = ret.results;
        var sheetMatchArr = sheetArray.filter(function( sht ) {
            return sht.name == config.configSheet.name;
        });
        // only proceed if config sheet is found
        if(sheetMatchArr.length > 0)
            site.returnRangeValues(config.configSheet.name, "A1", callback)
    },
    //3. get config object
    //4. process template
    function(ret, callback){
        if(!ret.results.isJSONString) return false;
        var sheetConfig = JSON.parse(ret.results);
        //loop through each sheet
        site.async.each(sheetConfig.importTemplate, function(t){
            // yep, async within an async
            site.async.waterfall([
                // 4a.
                function(cback){
                    // get the named ranges of dates, fields, and custom dimensions if available
                    var nrArray = [t.date.nr, t.field.nr];
                    if(t.customDim.nr != null) nrArray.push(t.customField.nr);
                    site.returnNamedRangeValues(nrArray, cback);
                },
                // 4b.
                function(ret, cback){
                    var nrObj = ret.results;
                    // get field ids for each field based on matched aliases
                    mapTemplateFields(t.id, nrObj[t.field.nr].rng, function(idArray){
                        nrObj.fieldIDArray = idArray;
                        cback(null, nrObj, cback);
                    });
                },
                // 4c find the data already
                function(nrObj, cback){

                    var columnAddr = nrObj[t.field.nr].value;
                    var rowAddr = nrObj[t.date.nr].value;
                    site.findDataValuesByDim(t.sheet, rowAddr, columnAddr,  function(err, ret){
                        // loop through each data column
                        for(var i = 0; i < ret.results.length; i++){
                            var colData = ret.results[i];
                            var fsDate = nrObj[t.date.nr].rng[i];
                            // handle custom dimension
                            var customDim = (t.customDim.nr == null) ? t.customDim.value : nrObj[t.customDim.nr].rng[i];
                            if(typeof t.customDim.convert[customDim] !== 'undefined')
                                customDim = t.customDim.convert[customDim];

                            // create new field set
                            addFieldSet({
                                mfiID: mfiID,
                                sheetID: t.id,
                                fsDate: fsDate,
                                customDim: customDim,
                                vals: colData,
                                ids: nrObj.fieldIDArray
                            });
                        }
                    });
                }

            ]);
        }); // end of loop throught input sheets
    }
]);

function addFieldSet(fObj){
    site.db.addReportingFieldSet(fObj, function(err, results){
        var fieldsetID = results.insertId;
        var insertArray = [];
        for(var j = 0; j < fObj.vals.length; j++){
            var currentVal = fObj.vals[j];
            if(typeof currentVal != 'number') currentVal = 0;
            var currentFieldID = fObj.ids[j];
            insertArray.push([currentFieldID, currentVal, fieldsetID]);
        }
        site.db.insertDataValues(insertArray, function(err, ret){console.log(ret)});
    });
}

function mapTemplateFields(sheetID, fieldArray, callback){
    site.db.getReportingLabelsBySheet(sheetID, function(err, results){
        var idArray = [];
        for(f = 0; f < fieldArray.length; f++){
            var i = idArray.length;
            idArray[i] = null;
            for(r = 0; r < results.length; r++){
                if(results[r].alias == fieldArray[f]){
                    idArray[i] = results[r].field_label_id;
                }
            }
        }
        callback(idArray);
    });
}
