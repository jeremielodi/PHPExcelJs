class Chart {

  constructor(type){
    this.type = type;
    this.seriesLabels = [];
    this.xAxisTickValues = [];
    this.seriesValues = [];
    this.valueTitle = '';
    this.topLeftPosition = 'A1';
    this.bottomRightPosition = 'H13';
  }
  setTitle(t){
    this.title = t;
    return this;
  }
  setValueTitle(t){
    this.valueTitle = t;
    return this;
  }
  
  setTopLeftPosition(key){
    this.topLeftPosition = key;
    return this;
  }
  setBottomRightPosition(key){
    this.bottomRightPosition = key;
    return this;
  }
  /**@param valDef array
   * @param rowNumber , number of row
   * @param dataType is series' savlue data type
   */

  setSeriesValues(valDef, dataType, rowNumber){
    this.seriesValues = valDef;
    this.seriesValuesRowNumber = rowNumber;
    this.seriesValuesDataType = dataType;
    return this;
  }

  setXAxisTickValues(v, dataType, rowNumber){
    this.xAxisTickValues = v;
    this.xAxisTickValuesRowNumber = rowNumber;
    this.xAxisTickValuesDataType = dataType;
    return this;
  }
  setSeriesLables(s, dataType, rowNumber){
    this.seriesLabels = s;
    this.seriesLabelsRowNumber = rowNumber;
    this.seriesLabelsDataType = dataType;
    return this;
  }
}


module.exports.Chart = Chart;