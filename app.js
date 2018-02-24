const phpExcel = require('./index');

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

const WorkBook = {
  name: 'Myfile.xlsx',
  sheets: [{
    name: 'My sheet',
    columns: [{
        key: 'A',
        width: 12
      },
      {
        key: 'B',
        width: 10
      },
      {
        key: 'C',
        width: 15
      }
    ],
    rows: [
      [{
          key: 'A5',
          value: "Nom",
          style: styleHeader
        },
        {
          key: 'B5',
          value: "dob",
          style: styleHeader
        },
        {
          key: 'C5',
          value: "prenom",
          style: styleHeader
        }
      ],
      [{
          key: 'A6',
          value: 23.546,
          style: style1
        },
        {
          key: 'B6',
          value: '2018-01-03',
          style: dateStyle
        },
        {
          key: 'C6',
          value: 'Jeremie Lodi'
        }
      ],
      [{
          key: 'A7',
          value: 10
        },
        {
          key: 'B7',
          value: 20
        },
        {
          key: 'C7',
          value: '=SUM(A7:B7)'
        }
      ],

    ],

    drawings: [{
        name: 'logo',
        cordonnates: 'A2',
        path: __dirname + '\\images\\phpexcel_logo.gif',
        height: 40,
        offsetX: 20
      },
      {
        name: 'logo',
        cordonnates: 'L2',
        path: __dirname + '\\images\\phpexcel_logo.gif',
        height: 40,
        offsetX: 20
      }
    ]
  },

{
  name :"sheet2",
  rows : [
    [
      {key:'A1', value:"First name", style: styleHeader},
      {key:'B2', value:"dob", style: styleHeader},
      {key:'C3', value:"Last name", style: styleHeader}
    ]]
}
]
};

var express = require('express');
let app = express();

app.use(express.static(__dirname));
app.get('/excel', (req, res, next) => {

  phpExcel.render(WorkBook).then(ExcelBuffer => {

    res.writeHead(200, {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': 'attachment; filename=file.xlsx',
      'Content-Length': ExcelBuffer.length
    });
    res.end(ExcelBuffer);
  }).catch(err => {
    console.log(err);
  }).done();

});

app.listen(8181);
console.log('app run on 8181');