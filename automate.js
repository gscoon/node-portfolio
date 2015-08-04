require('./js/site');

// open a file first

//C:/projects/node-portfolio/files/templates/filled/Access Bank_2406478.xlsx
site.currentBookID = site.rand(5);
site.excel.call({
    func: 'OpenExcelFile',
    wbID: site.currentBookID,
    openType: 'add',
    src: 'C:/projects/node-portfolio/files/templates/filled/Access Bank_2406478.xlsx'
}, function(err, results){
    site.returnRangeValues(function(rErr, rResults){
        console.log('you should have an array now');
        console.log(rResults);
        console.log(rErr);
    });
});




/*

{
    importTemplates: [{
    sheetName: "AGGREGATED_DATA",
    id:17,
    date: multiple,
    custom_field:[
        {
            name: isAudited,
            convert:{
                "UA": 0,
                "Audited": 1
            }
        }
    ],
    fieldRange:"",
    dateRange:"",
    fieldList:[
        alias:[],
        id
    ]
}],

}
*/
