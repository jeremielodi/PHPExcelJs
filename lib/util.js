class Util {
  /**
   * Translates a Excel cell represenation into row and column numerical equivalents 
   * @function getExcelRowCol
   * @param {String} str Excel cell representation
   * @returns {Object} Object keyed with row and col
   * @example
   * // returns {row: 2, col: 3}
   * getExcelRowCol('C2')
   */
  getExcelRowCol(str) {
    var numeric = str.split(/\D/).filter(function (el) {
      return el !== '';
    })[0];
    var alpha = str.split(/\d/).filter(function (el) {
      return el !== '';
    })[0];
    var row = parseInt(numeric, 10);
    var col = alpha.toUpperCase().split('').reduce(function (a, b, index, arr) {
      return a + (b.charCodeAt(0) - 64) * Math.pow(26, arr.length - index - 1);
    }, 0);
    return {
      row: row,
      col: col
    };
  };

  /**
   * Translates a column number into the Alpha equivalent used by Excel
   * @function getExcelAlpha
   * @param {Number} rowNum Row number that is to be transalated
   * @param {Number} colNum Column number that is to be transalated
   * @returns {String} The Excel alpha representation of the column number
   * @example
   * // returns B1
   * getExcelCellRef(1, 2);
   */
  getExcelCellRef(rowNum, colNum) {
    var remaining = colNum;
    var aCharCode = 65;
    var columnName = '';
    while (remaining > 0) {
      var mod = (remaining - 1) % 26;
      columnName = String.fromCharCode(aCharCode + mod) + columnName;
      remaining = (remaining - 1 - mod) / 26;
    }
    return columnName + rowNum;
  };
    //returns column letter eg. A for colNum = 1, C for colNum =3
    getExcelAlpha(colNum) {
      var remaining = colNum;
      var aCharCode = 65;
      var columnName = '';
      while (remaining > 0) {
        var mod = (remaining - 1) % 26;
        columnName = String.fromCharCode(aCharCode + mod) + columnName;
        remaining = (remaining - 1 - mod) / 26;
      }
      return columnName;
    }
}

module.exports.Util = Util;