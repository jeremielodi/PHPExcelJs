class Cell {
  constructor() {
    this.val = '';
  }
  value(val) {
    this.val = val;
    return this;
  }
  style(styl){
    this.style = styl;
    return this;
  }
}

class Column {
  constructor() {
    this.width = 10;
  }
  setWidth(width){
    this.width = width;
    return this;
  }
}

class Sheet {
  constructor(_name) {
    this.name = _name,
      this.columns = {};
      this.freezePanes = {};
      this.cells = {};
      this.merges = [];
      this.drawings = []
  };

  setCellValue(key, val){
    return this.cell(key).value(val);
  }

  freezePane(key){
    if (this.freezePanes[key]) {
      this.freezePanes[key];
    } else {
      this.freezePanes.push(key);
    }
    return this;
  }

  col(key){
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
    if (!row && (typeof(col) === 'string')) {
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

  mergeCells(keyFrom, keyTo){
    this.merges.push("".concat(keyFrom,":", keyTo));
    return  this.cell(keyFrom);
  }
  //returns column number
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

class WorkBook {

  constructor() {
    this.sheets = [];
  }

  addWorksheet(name) {
    const n = this.sheets.length;
    this.sheets[n] = new Sheet(name);
    return this.sheets[n];
  }


}


let test = new WorkBook();
var sh1 = test.addWorksheet("S1");
sh1.cell(1, 1).value("JEREMIE LODI");
sh1.cell(3, 1).value("Bad");
sh1.setCellValue("A2", "Works");
sh1.cell('A2').style({
  font :{bold:true}
});

sh1.mergeCells("C4", "C8").value("Super long test");
sh1.col("A").setWidth(30);
console.log(JSON.stringify(test));