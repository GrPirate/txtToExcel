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
    // 按行切割文件内容
    let items = config.toString().trim().split('\r\n')
    const results = [] // ["1 2 3", "1,2,3", ....]
    let maxLength = 0 // 记录最长的数组长度
    // 按空格切割每行数据生成每列
    for (let i = 0, len = items.length; i < len; i++) {
      let temp = items[i].split(/\s+/)
      if (temp.length > maxLength) maxLength = temp.length
      results.push(temp)
    }
    console.log(results)
    const conf = {};
    conf.name = "mysheet"
    conf.cols = []
    for(let i = 0; i < maxLength; i++) {
      let item = {
        caption: `header${i + 1}`,
        type: 'string',
        beforeCellWrite:function(row, cellData){
          return cellData ? cellData : '';
        },
        width: 40
      }
      config.cols.push(item)
    }
    conf.rows = results
    const result = excelPort.execute(conf)
    res.setHeader('Content-Type', 'application/vnd.openxmlformats');
    res.setHeader("Content-Disposition", "attachment; filename=" + "Report.xlsx");
    res.end(result, 'binary');
  })
  
})

app.listen(3000, () => console.log('Example app listening on port 3000'))