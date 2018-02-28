# PHPExcelJs
Node version of PHPExcel
<br/><br/>
In this repository I'm using PHPExcel library for recreating xlsx files with nodejs
<br/>
As this precess requires interation between PHP and nodejs, php must be installed in your server <br/>
```sudo apt-get install php7-cli```
<br>
I facilitate this communication by creating a radom json file (in nodejs) that I save in temp folder, next I execute a child process 
```"php  convertor.php " + jsonFile``` for calling php.
<br> ```jsonFile``` help php to know where to get informations
<br> Once php get the file content, the file is diretly deleted
<br/><br/>
Usage

```js
var express = require('express');

const WorkBook = require("./lib/workBook").WorkBook;
let app = express();

app.use(express.static(__dirname));
app.get('/excel', (req, res, next) => {


  const styleFormatedNumber = {
    format: "#.##",
    fontSize: 19,
    fill: 'FF0000'
  }

  const styleHeader = {
    alignment: {
      key: 'horizontal',
      value: 'center'
    },
    font: {
      color: {
        'argb': "FFFFFF"
      }
    },
    fill: '0066ff'

  }
  const dateType = {
    type: 'date',
    format: "M/D/YYYY"
  }
  const styleBorder = {
    border: {
      style: 'thin',
      color: 'FFFF0000',
      //position :  'right', 'left', 'top', 'bottom' . default 'allborders'
    },
  }
  const dateStyle = {
    type: 'date',
    format: "M/D/YYYY",
    border: {
      style: 'thin',
      color: 'FFFF0000',
      //position : 'allborders'
    },
    font: {
      bold: true,
      color: {
        'argb': 'FFFF0000'
      },
      size: 9,
      name: 'Vardana'
    },
    alignment: {
      key: 'horizontal',
      value: 'center',
      rotation: 45
    }
  }

  const textRotale = {
    alignment: {
      rotation: 45
    }
  }

  const bold = {
    font: {
      bold: true
    }
  }


  let wb = new WorkBook('Myfile.xlsx');

  var ws = wb.addWorksheet("Sheet1");

  // A1
  ws.cell(1, 1).value("JEREMIE LODI").style(styleHeader);
  ws.cell(5, 1).value("Bad"); //A5
  ws.setCellValue("A3", 1200.8747).style(styleFormatedNumber);
  ws.setCellValue("A2", "Works");
  ws.cell('A2').style(bold);

  ws.freezePane("A1");
  ws.mergeCells("C4", "F12").value("Super long test underlined and rotated")
    .freeze()
    .style(styleBorder, textRotale)
    .setUnderline(true);

  ws.col("A").setWidth(30);

  ws.cell("D1").value("1992-03-13").style(dateType);

  var ws2 = wb.addWorksheet("sheet 2");
  ws2.col("A").setWidth(15);
  ws2.col("B").setWidth(15);
  ws2.setCellValue("A1", "First name").style(styleHeader).setUnderline(true);

  ws2.setCellValue("B1", "Last name").style(styleHeader).setUnderline(true);
  for (var i = 2; i < 5000; i++) {
    ws2.setCellValue("A" + i, "Alice" + i);
    ws2.setCellValue("B" + i, "Bob" + i);
  }

  wb.render().then(result => {
    res.set(result.headers);
    res.send(result.report); //report is excel's stream
  }).catch(err => {
    console.log(err);
  }).done();

});

app.listen(8181);
console.log('app run on 8181');
```