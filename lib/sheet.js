const Column = require('./col').Column;
const Cell = require('./cell').Cell;
const Chart = require('./chart').Chart;
//
class Sheet {

  constructor(_name) {
    this.name = _name,
    this.columns = {};
    this.freezePanes = [];
    this.cells = {};
    this.merges = [];
    this.drawings = [],
    this.alphabet = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];
    this.charts = [];
  };

  addChart(type){
    var chart = new Chart(type);
    this.charts.push(chart);
    return chart;
  }
  //return A1 for col=1, row = 1
  cellIndex(col, row) {
    return this.getCol(row).concat(col);
  }

  /**
   * fill excel's cells from an array
   * @param {array[][]} array 
   */
  fromArray(array){
    array.forEach((row, i) => {
      row.forEach((col, j) => {
        this.cell(i+1,j+1).value(col);
      });
    });
  }
  
  /**
   * 
   * @param {*} key cell's key(eg. A1)
   * @param {*} val value(eg. 10)
   */
  setCellValue(key, val) {
    const cell = this.cell(key);
    cell.value(val);
    return cell;
  }

  freezePane(...key) {
    key = [].concat(key);

    key.forEach(k => {
      Object.assign(this.freezePanes, k);
    });

    return this;
  }

  col(key) {
    if (this.columns[key]) {
      return this.columns[key];
    } else {
      const col = new Column();
      this.columns[key] = col;
      return this.columns[key];
    }
  }
  cell(row, col) {
    var key = '';
    if (!col && (typeof (row) === 'string')) {
      key = row;
    } else {
      key = this.getCol(col).concat(row);
    }

    if (this.cells[key]) {
      return this.cells[key];
    } else {
      var cell = new Cell();
      this.cells[key] = cell;
      return this.cells[key];
    }
  }


  mergeCells(keyFrom, keyTo) {
    this.merges.push(''.concat(keyFrom, ':', keyTo));
    //linking styles for all merged cells
    const cellFrom = this.cell(keyFrom);
    const cel1 = this.coordonates(keyFrom);
    const cel2 = this.coordonates(keyTo);
    for (var i = cel1.row; i <= cel2.row; i++) {
      for (var j = cel1.col; j <= cel2.col; j++) {
        this.cell(i, j).styles = cellFrom.styles;
      }
    }
    return cellFrom;
  }
  //returns column letter eg. A for index = 0, C for index =2
  getCol(index) {
    index = parseInt(index) - 1;
    var n = this.alphabet.length;
    var nb_div = index / (n);

    var nTurn = Math.floor(nb_div);
    var mod = nb_div % 2;
    var intRight = parseInt(mod, 10);
    var right = (intRight === mod) ? mod : (intRight + 1)


    if (nTurn === 0) {
      return this.alphabet[index];
    }

    if (nTurn < 26) {
      console.log(nb_div);
      return this.alphabet[nTurn - 1].concat(this.alphabet[right]);
    } else {
      var div = Math.floor(nTurn / n);
      return this.getCol(nTurn / n).concat(this.getCol(div));
    }
  }

  //return [1,1] for key = A1
  coordonates(key) {
    key = key.split('');
    var r = {
      row: '',
      col: ''
    }
    key.forEach((c, i) => {

      const ch = parseInt(c, 10);
      if (isNaN(ch)) {
        r.col = r.col.concat(c);
      } else {
        r.row = r.row.concat(c);
      }
    });

    if (r.col.length > 1) {
      r.col = r.col.split('');
      var n = r.col.length;
      var sum = 0;
      for (var i = 0; i < n - 1; i++) {
        var indice = (this.alphabet.indexOf(r.col[i]) + 1);

        if (n > 1) {
          indice = indice * 26;
        }
        sum += indice;
      }
      var lastIndice = (this.alphabet.indexOf(r.col[r.col.length - 1]) + 1);

      r.col = sum + lastIndice;
    } else {
      r.col = (this.alphabet.indexOf(r.col[0]) + 1);
    }
    r.row = parseInt(r.row, 10);
    return r;
  }
}




module.exports.Sheet = Sheet;