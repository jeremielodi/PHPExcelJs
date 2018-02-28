
const Sheet = require('./sheet').Sheet;
const phpExcel = require('./phpExcel');

class WorkBook {

  constructor(name) {
    this.name = name;
    this.sheets = [];
  }

  addWorksheet(name) {
    const workSheet = new Sheet(name);
    this.sheets.push(workSheet);
    return workSheet;
  }

  render(){
    return phpExcel.render(this);
  }
}


module.exports.WorkBook =  WorkBook;
