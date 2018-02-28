const Column = require('./col').Column;
const Cell = require('./cell').Cell;
//
class Sheet {

  constructor(_name) {
    this.name = _name,
    this.columns = {};
    this.freezePanes = [];
    this.cells = {};
    this.merges = [];
    this.drawings = []
  };

  //return A1 for col=1, row = 1
  cellIndex(col, row){
    return this.getCol(row - 1).concat(col);
  }

  setCellValue(key, val) {
    const cell =  this.cell(key);
    cell.value(val);
    return cell;
  }

  freezePane(...key) {

   /*if( isArray(key)){
     console.log(key);
   }*/

    if (this.freezePanes[key]) {
      this.freezePanes[key];
    } else {
      this.freezePanes.push(key);
    }
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
  cell(col, row) {
    var key = "";
    if (!row && (typeof (col) === 'string')) {
      key = col;
    } else {
      key = this.getCol(row - 1).concat(col);
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
    return this.cell(keyFrom);
  }
  //returns column number eg. A for index = 0, C for index =2
  getCol(index) {
    var cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];
    if (index > cols.length - 1) {
      var nb_div = index / (cols.length);
      var nTurn = Math.floor(nb_div);
      var str = "";
      for (var i = 0; i < nTurn; i++) {
        str += cols[0];
      }
      var mod = index % (cols.length);
      return str + cols[mod];
    } else {
      return cols[index];
    }
  }
}




module.exports.Sheet = Sheet;