var remote = require('remote');
$ = require('jquery');
require('./../config/global.js');

site = {
    config: require(__dirname + '/../config/config'),
    request: require('request'),
    async: require('async'),
    excel: require(__dirname + '/excel'),
    setUp: function(){
        this.w = remote.getCurrentWindow();

        $('#x').on('click', function(){
            site.w.close();
        });

        this.excel.call({fn:'SetExcelObject'}, function(error, result){
            var message = (typeof result == 'undefined')?'excel app NOT set': 'excel app set';
            console.log(message);
        });

        this.analysisTemplateQueries(2);
    }
};

site.db = require(__dirname + '/../config/db')

/*setTimeout(function(){
    site.excel.call({fn:'PopulateSheet'}, function(error, result){
        if(typeof result == 'undefined'){
            console.log('populatesheet NOT set');
            return false;
        }

        console.log('populatesheet is set');
        console.log(result);
    });
},5000);*/
site.analysisTemplateQueries = function(mfiID){
    site.db.getAnalysisTemplateData(mfiID, function(err, data){
        site.excel.call({
            fn:'PopulateDataSheet',
            data: data
        },function(error, result){
            if(typeof result == 'undefined'){
                console.log('getAnalysisTemplateData NOT set');
                console.log(error);
                return false;
            }

            console.log('getAnalysisTemplateData is set');
            console.log(result);
        });
    });
}

$(function(){
    site.setUp();
});
