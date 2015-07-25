var remote = require('remote');

var siteClass = function(){

    this.setUp = function(){
        this.excel = require(__dirname + '/inc/excel');  // excel module
        this.w = remote.getCurrentWindow();

        $('#x').on('click', function(){
            this.w = close();
        });

        this.excel.call({fn:'SetExcelObject'}, function(error, result){
            console.log(error);
            console.log(result);
        });

        setTimeout(function(){
            site.excel.call({fn:'PopulateSheet'}, function(error, result){
                console.log(error);
                console.log(result);
            })
        },5000)
    }

}

site = new siteClass();
