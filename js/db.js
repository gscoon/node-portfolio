var mysql = require('mysql');

var dbClass = function(){
    console.log('MYSQL obj created');
    this.conn = mysql.createConnection({
        host: config.db.server,
        port: config.db.port,
        user: config.db.user,
        password: config.db.password,
        database: config.db.name,
        dateStrings: 'DATE' // most annoying feature that I had to account for
    });

    this.getReportingQueries = function(qKeys, mfiID, callBack){
        var inString = '"' + qKeys.join('", "') + '"';
        var q = 'SELECT * FROM reporting_query WHERE query_id IN ({0})'.format(inString);
        runQuery(q, [], function(err, ret){
            var params = [mfiID];
            site.async.map(qKeys, function(key, mapCallBack){
                var isMatched = false;
                // map each query and it's results to the keys you passed, maintaining order
                ret.some(function(q){
                    if(q.query_id == key){
                        isMatched = true;
                        runQuery(q.query_string, params, function(qErr, qResults, fields){
                            q.results = qResults;
                            q.fields = fields;
                            mapCallBack(qErr, q);
                        });
                        return true;
                    }
                });
                if(!isMatched) mapCallBack(null, q);
            }, callBack);
        });
    }

    this.getMFIQueries = function(qArray){
        var retArray = [];
        var me = this;
        qArray.forEach(function(name, index){
            // return object with query string and name
            if(typeof me.qs.mfi[name] !== 'undefined')
                retArray.push({
                    name: name,
                    qStr: me.qs.mfi[name]
                });
        });
        return retArray;
    }


    this.getMFIDetails = function(mfiID, callBack){
        var qArray = this.getMFIQueries(['mfi_detail']);
        var params = [mfiID];
        runQuery(qArray[0].qStr, params, callBack);
    }

    this.getTemplateData = function(qArray, mfiID, callBack){
        var params = [mfiID];
        site.async.map(qArray, function(q, mapCallBack){
            runQuery(q.query_string, params, function(err, results, fields){
                q.results = results;
                q.fields = fields;
                mapCallBack(err, q);
            });
        }, callBack);
    }

    this.getReportingLabelsBySheet = function(sheetID, callback){
        var q = "SELECT fl.sheet_id, fl.field_label_id, a.alias FROM reporting_field_label_alias a JOIN reporting_field_label fl ON fl.field_label_id = a.field_label_id WHERE fl.sheet_id = ?";
        runQuery(q, [sheetID], callback);
    }

    this.addReportingFieldSet = function(fObj, callback){
        var q = 'INSERT INTO reporting_field_set (mfi_id, sheet_id, field_set_date, is_audited) VALUES (?, ?,  DATE(DATE_ADD("1899-12-30", INTERVAL ? day)), ?)';
        // 1899-12-30 because excel addes leep year in 1900
        var params = [fObj.mfiID, fObj.sheetID, fObj.fsDate, fObj.customDim];
        runQuery(q, params, callback);
    }

    this.insertDataValues = function(valueArray, callback){
        var q = "INSERT INTO reported_financials (field_label_id, field_value, field_set_id) VALUES ?";
        runQuery(q, [valueArray], callback);
    }

    this.checkRefreshTable = function(callback){
        // status: pending|processing|completed
        // LIMIT to 1
        var q = "SELECT * FROM reporting_refresh rr LEFT JOIN reporting_custom_template rct ON rct.template_id = rr.template_id WHERE rr.status = 'pending' LIMIT 1";
        runQuery(q, [], callback);
    }

    this.updateRefreshTable = function(id, status, callback){
        var ts = site.moment().format("YYYY-MM-DD HH:mm:ss");
        var q = 'UPDATE reporting_refresh SET status = ?, update_ts = ? WHERE refresh_id = ?';
        var params = [status, ts, id];
        runQuery(q, params, callback);
    }

    this.insertRefresh = function(mfiID, templateID, status){
        var ts = site.moment().format("YYYY-MM-DD HH:mm:ss");
        if(typeof status == 'undefined') status = 'pending';
        var q = "INSERT INTO reporting_refresh (mfi_id, template_id, request_ts, status) VALES (?, ?, ?, ?)";
        var params = [mfiID, templateID, ts, id];
    }

    this.getAllMFIsWithData = function(callback){
        var q = "SELECT DISTINCT rfs.mfi_id, m.mfi_name FROM reporting_field_set rfs LEFT JOIN mfi m ON m.mfi_id = rfs.mfi_id ORDER BY rfs.mfi_id ASC;";
        runQuery(q, [valueArray], callback);
    }

    function runQuery(q, params, callback){
        site.db.conn.query(q, params, callback);
    }

}

module.exports = new dbClass();
