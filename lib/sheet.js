const Column = require('./col').Column;
const Cell = require('./cell').Cell;
const Chart = require('./chart').Chart;
const Util = require('./util').Util;
const  util = new Util();
//
class Sheet {

  constructor(_name) {
    this.name = _name,
      this.columns = {};
    this.freezePanes = [];
    this.cells = {};
    this.merges = [];
    this.drawings = [],
      this.charts = [];
  };

  addChart(type) {
    var chart = new Chart(type);
    this.charts.push(chart);
    return chart;
  }
  //return A1 for col=1, row = 1
  cellIndex(col, row) {
    return this.getExcelAlpha(row).concat(col);
  }

  /**
   * fill excel's cells from an array
   * @param {array[][]} array 
   */
  fromArray(array) {
    array.forEach((row, i) => {
      row.forEach((col, j) => {
        this.cell(i + 1, j + 1).value(col);
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
      key = this.getExcelAlpha(col).concat(row);
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
    const cel1 = this.getExcelRowCol(keyFrom);
    const cel2 = this.getExcelRowCol(keyTo);
    for (var i = cel1.row; i <= cel2.row; i++) {
      for (var j = cel1.col; j <= cel2.col; j++) {
        this.cell(i, j).styles = cellFrom.styles;
      }
    }
    return cellFrom;
  }

  getExcelAlpha(colNum) {
    return  util.getExcelAlpha(colNum);
  }
  getExcelCellRef(rowNum, colNum) {
    return util.getExcelCellRef(rowNum, colNum);
  }

  getExcelRowCol(str){
    return util.getExcelRowCol(str);
  }
}



module.exports.Sheet = Sheet;