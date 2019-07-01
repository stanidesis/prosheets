function Row(rowRange) {
  this.rowRange = rowRange
  this.hydrate()
}

Row.prototype.hydrate = function() {}

Row.prototype.commitToRow = function() {}

Row.prototype._moveToSheet = function(sheetName) {
  var newRange = newRowIn(sheetName)
  this.rowRange.copyTo(newRange)
  removeRow(this.rowRange.getSheet().getSheetName(), this.rowRange.getRow())
  this.rowRange = newRange
}

Row.prototype.isBlank = function() {
  var values = this.rowRange.getValues()
  for (var i = 0; i < values[0].length; i++) {
    if (typeof values[0][i] !== 'undefined') {
      return false
    }
  }
  return true
}

Row.prototype._deleteRow = function() {
  removeRow(this.rowRange.getSheet().getSheetName(), this.rowRange.getRow())
}
