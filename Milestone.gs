function Milestone(rowRange) {
  Row.call(this, rowRange)
  
  this.markAsComplete = function() {
    this.status = CONSTS.STATUS.COMPLETED
    this.commitToRow()
    this._moveToSheet(CONSTS.SHEETS.COMPLETED_MILESTONES)
  }
  
  this.markAsIncomplete = function() {
    this._moveToSheet(CONSTS.SHEETS.MILESTONES)
    this.status = CONSTS.STATUS.ACTIVE
    this.commitToRow()
  }
}

Milestone.prototype = Object.create(Row.prototype)
Milestone.prototype.constructor = Milestone
Milestone.prototype.hydrate = function() {
  var row = this.rowRange.getValues()
  this.project = row[0][0]
  this.title = row[0][1]
  this.description = row[0][2]
  this.status = row[0][3]
  this.priority = row[0][4]
}
Milestone.prototype.commitToRow = function() {
  var values = this.rowRange.getValues()
  values[0][0] = this.project
  values[0][1] = this.title
  values[0][2] = this.description
  values[0][3] = this.status
  values[0][4] = this.priority
  this.rowRange.setValues(values)
}
Milestone.prototype.generateMilestoneProject = function() {
  return this.project + ': ' + this.title
}
Milestone.prototype.deleteRow = function() {
  this._deleteRow()
}
