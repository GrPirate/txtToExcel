const fs = require('fs')
const path = require('path')
const express = require('express')
const app = express()
const excelPort = require('excel-export')

app.get('/', (req, res) => {
  res.send('Hello world')
})

app.get('/Excel', (req, res) => {
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
    console.log(results)
    const conf = {};
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
    conf.rows = results
    const result = excelPort.execute(conf)
    res.setHeader('Content-Type', 'application/vnd.openxmlformats');
    res.setHeader("Content-Disposition", "attachment; filename=" + "Report.xlsx");
    res.end(result, 'binary');
  })
  
})

app.listen(3000, () => console.log('Example app listening on port 3000'))