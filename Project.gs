function Project(rowRange) {
  Row.call(this, rowRange)
  
  this.markAsComplete = function() {
    this.status = CONSTS.STATUS.COMPLETED
    this.commitToRow()
    this._moveToSheet(CONSTS.SHEETS.COMPLETED_PROJECTS)
  }
  
  this.moveToProjects = function() {
    this._moveToSheet(CONSTS.SHEETS.PROJECTS)
  }
}

Project.prototype = Object.create(Row.prototype)
Project.prototype.constructor = Project
Project.prototype.hydrate = function() {
  var row = this.rowRange.getValues()
  this.title = row[0][0]
  this.description = row[0][1]
  this.type = row[0][2]
  this.status = row[0][3]
}
Project.prototype.commitToRow = function() {
  var values = this.rowRange.getValues()
  values[0][0] = this.title
  values[0][1] = this.description
  values[0][2] = this.type
  values[0][3] = this.status
  this.rowRange.setValues(values)
}
