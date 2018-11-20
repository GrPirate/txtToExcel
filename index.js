const fs = require('fs')
const path = require('path')
const excelPort = require('excel-export')


exports.write = function(req, res) {
  console.log(req)
  var conf = {};
  conf.name = "mysheet"
  conf.cols = [
    {caption: 'gene_id1', type: 'string', width: 20},
    {caption: 'gene_id2', type: 'string', width: 40},
    {caption: 'gene_id3', type: 'string', width: 40},
    {caption: 'gene_id4', type: 'string', width: 40},
    {caption: 'gene_id5', type: 'string', beforeCellWrite:function(row, cellData){
      return cellData ? cellData : '';
    }, width: 40
  }]
  conf.rows = req
  var result = excelPort.execute(conf)
  var filePath = "./result.xlsx";
  fs.writeFile(filePath, result, 'binary', function(err) {
    if (err) {
        console.log(err);
    }
    console.log("success!");
  });

};


fs.readFile(path.join(__dirname, 'SummaryReport01.txt'), "utf-8", function(error, config) {
  if (error) {
    console.log(error)
    console.log("config文件读入出错")
  }
  let items = config.toString().trim().split('\r\n')
  const results = []
  for (let i = 0, len = items.length; i < len; i++) {
    let temp = items[i].split(/\s+/)
    results.push(temp)
  }
  exports.write(results)
})

