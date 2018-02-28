var express = require('express');
let app = express();

app.use(express.static(__dirname));
app.get('/excel', (req, res, next) => {


  const style1 = {
    format: "#.##",
    fontSize: 19,
    fill: 'FF0000'
  }

  const styleHeader = {
    alignment: {
      key: 'horizontal',
      value: 'center'
    }
  }
  const dateStyle = {
    format: "m/d/yyyy",
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




  const WorkBook = require('./lib/workBook').WorkBook;

  let wb = new WorkBook('Myfile.xlsx');

  var ws = wb.addWorksheet("S1");
  ws.cell(1, 1).value("JEREMIE LODI").style(styleHeader, dateStyle);
  ws.cell(3, 1).value("Bad");
  ws.setCellValue("A3", 1200.8747).style(style1);
  ws.setCellValue("A2", "Works");

  var ws2 = wb.addWorksheet("classe 2");
  ws2.setCellValue("A1", '2017-03-09').style(dateStyle);
  ws.cell('A2').style({
    font: {
      bold: true
    }
  });

  ws.freezePane("A1");
  ws.mergeCells("C4", "F4").value("Super long test").freeze();
  ws.cell("C4").style(style1);

  ws.col("A").setWidth(30);



  wb.render().then(result => {

    res.set(result.headers);
    res.send(result.report); //result is excel's stream
  }).catch(err => {
    console.log(err);
  }).done();

});

app.listen(8181);
console.log('app run on 8181');