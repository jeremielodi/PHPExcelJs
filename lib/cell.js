

class Cell {
  constructor() {
    this.val = '';
    this.styles = {};
  }
  value(val) {
    this.val = val;
    return this;
  }
  style(...styl){
    if(Array.isArray(styl)){
      styl.forEach(stl=>{
        Object.assign(this.styles , stl);
      });
    }else{
      Object.assign(this.styles , styl);
    }
    
    return this;
  }

  freeze(){
    this.freezePan =  true;
    return this;
  }
}
module.exports.Cell = Cell;