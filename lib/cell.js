class Cell {
  constructor() {
    this.val = '';
    this.styles = {};
  }
  value(val) {
    this.val = val;
    return this;
  }
  style(...styl) {
    styl = [].concat(styl);

    styl.forEach(stl => {
      Object.assign(this.styles, stl);
    });

    return this;
  }

  setUnderline(flag) {
    this.underline = flag;
  }

  freeze() {
    this.freezePan = true;
    return this;
  }
}
module.exports.Cell = Cell;