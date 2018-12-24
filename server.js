
var http = require("http");
var request = require("request");
var url = require("url");
var express    = require("express");
var mysql      = require('mysql');
var xlsx = require('node-xlsx');

var connection = mysql.createConnection({
  host     : 'kamuna.org',
  user     : 'vsc',
  password : 'uispl@123',
  database : 'kchr'
});

var app = express();

connection.connect(function(err){
  if(!err) {
    console.log("Database is connected ... nn");
  } else {
    console.log("Error connecting database ... nn");
  }
});

app.get("/api/excel/:table",function(req, res) {
  connection.query('SELECT * from empmaster where workingstatus="ACTIVE"', [req.params.table], function(err, rows, fields) {
    if (!err) {
      res.setHeader('Content-disposition', 'attachment; filename=' + req.params.table + '.xlsx');
      rows = JSON.stringify(rows);
      rows = JSON.parse(rows.toString());
      var data = [];
      for (var index = 0; index < rows.length; index++) {
        var dataRow = [];
        for(var value in rows[index]) {
          dataRow.push(rows[index][value]);
        }
        data.push(dataRow);
      }
      var buffer = xlsx.build([{name: "main", data: data}]);
      res.send(buffer);
      console.log('Table %s to excel', req.params.table);
    }
    else {
      console.log('Error while performing Query.');
    }  
  });
});
app.listen(3000);