function Task (rowRange) {
  Row.call(this, rowRange)
  
  this.markAsIncomplete = function() {
    this.completedOn = ''
    var today = new Date()
    today.setHours(0,0,0)
    this.estimate = daysBetweenDates(this.startDate, today) + 1
    this.commitToRow()
    this.moveToTasks()
  }
  
  this.markAsCompleted = function(onDate) {
    if (!onDate) {
      onDate = new Date()
      onDate.setHours(0,0,0)
    }
    this.completedOn = onDate
    this.estimate = daysBetweenDates(this.startDate, onDate) + 1
    this.commitToRow()
    this.moveToCompletedTasks()
  }
  
  this.moveToCompletedTasks = function() {
    this._moveToSheet(CONSTS.SHEETS.COMPLETED_TASKS)
  }
  
  this.moveToTasks = function() {
    this._moveToSheet(CONSTS.SHEETS.TASKS)
  }
  
  this.getId = function() {
    return this.getJSON()[CONSTS.TASK.JSON_KEYS.ID]
  }
  
  this.setId = function(id) {
    var jsonData = this.getJSON()
    
    if (typeof id === undefined || id === null) {
      delete jsonData.id
    } else {
      jsonData[CONSTS.TASK.JSON_KEYS.ID] = id
    }
    this.setJSON(jsonData)
  }
  
  this.hasStartTime = function() {
    return typeof this.getJSON()[CONSTS.TASK.JSON_KEYS.START_TIMESTAMP] !== typeof undefined
  }
  
  this.getStartTime = function() {
    if (!this.hasStartTime()) {
      return undefined
    }
    return new Date(this.getJSON()[CONSTS.TASK.JSON_KEYS.START_TIMESTAMP])
  }
  
  this.setStartTime = function(date) {
    var jsonData = this.getJSON()
    if (typeof date === typeof undefined || date === null) {
      delete jsonData[CONSTS.TASK.JSON_KEYS.START_TIMESTAMP]
    } else {
      jsonData[CONSTS.TASK.JSON_KEYS.START_TIMESTAMP] = date.toString()
    }
    this.setJSON(jsonData)
  }
  
  this.addMinutesToTimeSpent = function(minutes) {
    if (this.timeSpent === '' || typeof this.timeSpent === typeof undefined) {
      this.timeSpent = new Date(Date.parse(CONSTS.TASK.TIME_SPENT_STARTING_POINT))
    }
    this.timeSpent = new Date(this.timeSpent.getTime() + minutes * 60000)
  }
  
  this.getTimeSpentDisplayValue = function() {
    return this.rowRange.getDisplayValues()[0][5]
  }
  
  this.getTaskSymbol = function() {
    if (this.completedOn === '') {
      return this.hasStartTime() ? CONSTS.TASK.CHAR_IN_PROGRESS : CONSTS.TASK.CHAR_OPEN
    }
    return CONSTS.TASK.CHAR_CLOSED
  }
  
  this.getCalendarTitle = function() {
    return this.getTaskSymbol() + ' ' + this.title
  }
  
  this.getCalendarDescription = function() {
    return '<table><tr><th>' + CONSTS.TASK.PROJ_MILESTONE_HEADER + '</th><td>' + this.projectMilestone + '</td></tr>' +
      '<tr><th>' + CONSTS.TASK.TIME_SPENT_HEADER + '</th><td>' + this.getTimeSpentDisplayValue() + '</td></tr>' +
      '<tr><th>' + CONSTS.TASK.DESCRIPTION_HEADER + '</th><td>' + this.description + '</td></tr></table>' +
        CONSTS.TASK.DESCRIPTION_FOOTER.replace('STAT_1',CONSTS.TASK.CHAR_OPEN).replace('STAT_2', CONSTS.TASK.CHAR_CLOSED).replace('STAT_3', CONSTS.TASK.CHAR_IN_PROGRESS) +
          getAllProjectMilestones().reduce(function(final, currentValue) {
          return final + CONSTS.TASK.PROJECT_MILESTONE_TEMPLATE.replace('%s',currentValue)
        }, '') + CONSTS.TASK.PROJECT_MILESTONE_END
  }
  
  this.getEventColor = function() {
    if (this.completedOn === '') {
      return this.hasStartTime() ? CONSTS.TASK.COLOR_IN_PROGRESS : CONSTS.TASK.COLOR_OPEN
    }
    return CONSTS.TASK.COLOR_CLOSED
  }
  
  this.getJSON = function() {
    var notes = this.rowRange.getNotes()
    var jsonString = notes[0][CONSTS.TASK.JSON_NOTE_INDEX] || '{}'
    var json = {}
    try {
      json = JSON.parse(jsonString)
    } catch (e) {
      // Backwards compatibility: they probably have an ID stored here
      json[CONSTS.TASK.JSON_KEYS.ID] = jsonString
    }
    return json
  }
  
  this.setJSON = function(jsonData) {
    var notes = this.rowRange.getNotes()
    notes[0][CONSTS.TASK.JSON_NOTE_INDEX] = JSON.stringify(jsonData)
    this.rowRange.setNotes(notes)
  }
}

Task.prototype = Object.create(Row.prototype)
Task.prototype.constructor = Task
Task.prototype.hydrate = function() {
  var row = this.rowRange.getValues()
  this.projectMilestone = row[0][0]
  this.title = row[0][1]
  this.description = row[0][2]
  this.estimate = row[0][3]
  this.startDate = row[0][4]
  this.timeSpent = row[0][5]
  this.etaDate = row[0][6]
  this.completedOn = row[0][7]
}
Task.prototype.commitToRow = function() {
  var values = this.rowRange.getValues()
  values[0][0] = this.projectMilestone
  values[0][1] = this.title
  values[0][2] = this.description
  values[0][3] = this.estimate
  values[0][4] = this.startDate
  values[0][5] = this.timeSpent  
  values[0][6] = ''
  values[0][7] = this.completedOn
  this.rowRange.setValues(values)
}
Task.prototype.hydrateFromCalendarEvent = function(event) {
  var projectMilestone = CONSTS.NA
  // Add logic to pull project+milestone from the description itself?
  if (projectMilestone === '') {
    throw 'Failed to find Project/Milestone'
  }
  this.projectMilestone = projectMilestone
  this.title = event.summary
  this.description = event.description
  if (!this.description) {
    this.description = ''
  }
  if (event.start.date) {
    this.startDate = parseDate(event.start.date)
  } else {
    this.startDate = parseDate(event.start.dateTime)
  }
  if (event.end.date) {
    this.estimate = daysBetweenDates(this.startDate, parseDate(event.end.date)) + 1
  } else {
    this.estimate = daysBetweenDates(this.startDate, parseDate(event.end.dateTime)) + 1
  }
  this.addMinutesToTimeSpent(0)
  this.setId(event.id)
}

Task.prototype.toString = function() {
  var tostring = ''
  if (this.rowRange) {
    tostring += 'Row: ' + this.rowRange.getSheet().getName() + ':' + this.rowRange.getA1Notation() + '\n'
  } else {
    tostring += 'Row: N/A\n'
  }
  tostring += 'ID: ' + (this.getId() ? this.getId() : 'N/A') + '\n'
  tostring += 'Project/Milestone: ' + (this.projectMilestone ? this.projectMilestone : 'N/A') + '\n'
  tostring += 'Title: ' + (this.title ? this.title : 'N/A') + '\n'
  tostring += 'Description: ' + (this.description ? this.description : 'N/A') + '\n'
  tostring += 'Start Date: ' + (this.startDate ? this.startDate.toLocaleString() : 'N/A') + '\n'
  tostring += 'Time Spent: ' + (this.rowRange.getDisplayValues()[0][6]) + '\n'
  tostring += 'Estimate: ' + (this.estimate ? this.estimate : 'N/A') + '\n'
  tostring += 'Completed On: ' + (this.completedOn ? this.completedOn.toLocaleString() : 'N/A')
  return tostring
}
