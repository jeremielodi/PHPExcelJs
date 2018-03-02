
class Drawing {

  constructor(path) {
    this.path = path;
  }
  setName(name) {
    this.name = name;
    return this;
  }

  setCoordinates(cellKey) {
    this.cordonates = cellKey;
    return this;
  }
  setHeight(h) {
    this.height = h;
    return this;
  }
  setRotation(r) {
    this.rotation = r;
    return this;
  }

  setDescription(d) {
    this.description = d;
    return this;
  }
  setDirection(direction) {
    this.direction = direction;
    return this;
  }

  setOffsetX(offsetx){
    this.offsetX = offsetx;
    return this;
  }
}
module.exports.Drawing = Drawing;
