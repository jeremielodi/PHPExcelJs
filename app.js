const express = require('express');

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
      /*
        dashDot, dashDotDot, dashed, dotted, double, hair,mediumDashDot
        mediumDashDotDot, mediumDashed, slantDashDot, thin, thick
      */
      style: 'medium',
      color: 'FFFF0000'
    },
  }
  const dateStyle = {
    type: 'date',
    format: "M/D/YYYY",
    border: {
      style: 'thin',
      color: 'FFFF0000',
      //position : 'allborders', left, outline, right,bottom, top, vertical, diagonal
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

  var ws = wb.addWorksheet("sheet1");

  // A1
  ws.cell(1, 1).value("JEREMIE LODI").style(styleHeader);
  ws.cell(5, 1).value("Bad"); //A5
  ws.setCellValue("A3", 1200.8747).style(styleFormatedNumber);
  ws.setCellValue("A2", "Works");
  ws.cell('A2').style(bold);

  ws.freezePane("A1");
  ws.mergeCells("C4", "F12").value("Super long text underlined and rotated")
    .freeze()
    .style(styleBorder, textRotale)
    .setUnderline(true);

  ws.col("A").setWidth(30);

  ws.cell("D1").value("1992-03-13").style(dateType);

  // new worksheet
  var ws2 = wb.addWorksheet("sheet 2");
  ws2.col("A").setWidth(15);
  ws2.col("B").setWidth(15);
  ws2.setCellValue("A1", "First name").style(styleHeader).setUnderline(true);

  ws2.setCellValue("B1", "Last name").style(styleHeader).setUnderline(true);

  for (var i = 2; i < 5000; i++) {
    ws2.setCellValue("A" + i, "Alice" + i);
    ws2.setCellValue("B" + i, "Bob" + i);
  }


  const logoPath = __dirname + "/images/phpexcel_logo.gif";
  let logo = ws.addImage(logoPath);
  logo.setName("Logo")
    .setDescription('Adding image')
    .setCoordinates("A8")
    .setOffsetX(20)
    .setRotation(5)
    .setHeight(40);



  // new worksheet
  var ws3 = wb.addWorksheet('Worksheet');

  ws3.fromArray(
    [
      ['', 2010, 2011, 2012],
      ['Q1', 12, 15, 21],
      ['Q2', 56, 73, 86],
      ['Q3', 52, 61, 69],
      ['Q4', 30, 32, 0],
    ]
  );


  let chart = ws3.addChart('barChart');
  chart.setTitle("Chart title")
    .setSeriesLables(['$B$1', '$C$1', '$D$1'], 'Number', 1) //2010 2011, 2012, 2013
    .setXAxisTickValues(['$A$2:$A$5'], 'String', 4) //Q1 to Q4
    .setSeriesValues(['$B$2:$B$5', '$C$2:$C$5', '$D$2:$D$5'], 'Number', 4)
    .setValueTitle("Value ($k)")
    .setTopLeftPosition("A7")
    .setBottomRightPosition("H20");


  wb.render().then(result => {
    res.set(result.headers);
    res.send(result.report); //report is excel's stream
  }).catch(err => {
    console.log(err);
  }).done();

});

app.listen(8181);
console.log('app run on 8181');