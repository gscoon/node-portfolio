var config = require('./config/config.json');
var mysql = require('./js/db');
var async = require('async');

var conn = mysql.createConnection({
    host: config.db.server,
    port: config.db.port,
    user: config.db.user,
    password: config.db.password,
    database: config.db.name,
    dateStrings: 'DATE' // most annoying feature that I had to account for
});

var q = "SELECT * FROM reporting_field_label";

conn.query(q, [], function(err, results){
    console.log(results[0]);
    var params = [];
    results.forEach(function(row){
        var id = parseInt(row.field_label_id);
        params.push([id, row.field_label_english.trim()]);

        if(row.field_label_spanish != null)
            params.push([id, row.field_label_spanish.trim()]);

        if(row.field_label_portuguese != null)
            params.push([id, row.field_label_portuguese.trim()]);
    });

    var insertq = "INSERT INTO reporting_field_label_alias (field_label_id, alias) VALUES ?";

    conn.query(insertq, [params], function(err2, results2){
        console.log(results2);
        console.log(err2);
    });
});
