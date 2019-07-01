function getCalId() {
  var calId = PropertiesService.getUserProperties().getProperty(CONSTS.PROPERTIES.CALENDAR_ID)
  if (calId === '' || calId === null) {
    setCalId(true)
    throw('Error: Calendar ID not set! Use Project Menu to set the calendar ID')
  }
  return calId
}

function getCalendar() {
  return CalendarApp.getCalendarById(getCalId())
}

function onUserEdit(event) {
  var cell = event.range
  var sheetName = cell.getSheet().getName()
  if (sheetName === CONSTS.SHEETS.TASKS) {
    var task = getTaskAtLocation(cell)
    // New task?
    if (!task.getId()) {
      if (task.projectMilestone && task.title && task.etaDate) {
        createOrUpdateTaskEvent(task)
      }
      return
    }
    if (task.completedOn) {
      // User set a completion date
      task.markAsCompleted(task.completedOn)
      createOrUpdateTaskEvent(task)
    } else if (!task.etaDate) {
      // User erased either the estimate or the start date
      deleteTaskEvent(task)
    } else if (task.projectMilestone && task.title) {
      createOrUpdateTaskEvent(task)
    }
  } else if (sheetName === CONSTS.SHEETS.MILESTONES) {
    var milestone = getMilestoneAtLocation(cell)
    if (cell.getValue() === CONSTS.STATUS.MARK_AS_COMPLETE) {
      milestone.markAsComplete()
      return
    } else if (cell.getValue() === CONSTS.STATUS.DELETE) {
      milestone.deleteRow()
      return
    }
  } else if (sheetName === CONSTS.SHEETS.PROJECTS) {
    var project = getProjectFromLocation(cell)
    if (cell.getValue() === CONSTS.STATUS.MARK_AS_COMPLETE) {
      project.markAsComplete()
    } else if (cell.getValue() === CONSTS.STATUS.DELETE) {
      project._deleteRow()
    }
  }
}

// https://developers.google.com/apps-script/reference/calendar/
function createOrUpdateTaskEvent(task) {
  var taskCalendar = getCalendar()
  var calenderEvent
  if (task.getId()) {
    // Update situation!
    calendarEvent = taskCalendar.getEventById(task.getId())
    calendarEvent.setTitle(task.getCalendarTitle())
    calendarEvent.setColor(task.getEventColor())
    calendarEvent.setDescription(task.getCalendarDescription())
    calendarEvent.setAllDayDates(task.startDate, 
                                 getRelativeDate(task.startDate, task.estimate))
  } else {
    // Create situation!
    calendarEvent = taskCalendar.createAllDayEvent(
      task.getCalendarTitle(),
      task.startDate,
      getRelativeDate(task.startDate, task.estimate),
      {description: task.getCalendarDescription()})
    
    calendarEvent.setColor(task.getEventColor())
    task.setId(calendarEvent.getId())
  }
}

function deleteTaskEvent(task) {
  var taskCalendar = getCalendar()
  try {
    taskCalendar.getEventById(task.getId()).deleteEvent()
  } catch (error) {
    // This event was deleted on the calendar but failed to sync to ProSheets
    console.log(error)
  }
  task.setId(null)
}

function taskFromId(id) {
  var taskSheet = SpreadsheetApp.getActive().getSheetByName(CONSTS.SHEETS.TASKS)
  var allTasks = taskSheet.getDataRange()
  var allNotes = allTasks.getNotes()
  // Find the right note!
  var noteIndex = CONSTS.TASK.JSON_NOTE_INDEX 
  for (var i = CONSTS.TEMPLATE_ROW_IDX; i < allNotes.length; i++) {
    if (allNotes[i][noteIndex] && allNotes[i][noteIndex].indexOf(id) > -1) {
      // Found it!
      return new Task(allTasks.offset(i, 0, 1))
    }
  }
  // Try Completed Tasks
  taskSheet = SpreadsheetApp.getActive().getSheetByName(CONSTS.SHEETS.COMPLETED_TASKS)
  allTasks = taskSheet.getDataRange()
  allNotes = allTasks.getNotes()
  for (var i = CONSTS.TEMPLATE_ROW_IDX; i < allNotes.length; i++) {
    if (allNotes[i][noteIndex] && allNotes[i][noteIndex].indexOf(id) > -1) {
      // Found the completed task!
      return new Task(allTasks.offset(i, 0, 1))
    }
  }
  return null
}

function syncOnce() {
  syncEvents(true)
}

function onCalendarUpdate(event) {
  syncEvents(false)
}

/**
 * https://developers.google.com/calendar/v3/reference/
 *
 * Retrieve and log events from the given calendar that have been modified
 * since the last sync. If the sync token is missing or invalid, log all
 * events from up to a month ago (a full sync).
 *
 * @param {string} calendarId The ID of the calender to retrieve events from.
 * @param {boolean} fullSync If true, throw out any existing sync token and
 *        perform a full sync; if false, use the existing sync token if possible.
 */
function syncEvents(fullSync) {  
  var lock = LockService.getScriptLock()
  lock.tryLock(1000)
  if (!lock.hasLock()) {
    return
  }
  
  var calendarId = getCalId()
  var properties = PropertiesService.getUserProperties()
  var options = {
    maxResults: 100
  }
  var syncToken = properties.getProperty('syncToken')
  if (syncToken && !fullSync) {
    options.syncToken = syncToken
  } else {
    // Sync events up to thirty days in the past.
    options.timeMin = getRelativeDateOffset(-30, 0).toISOString()
  }

  // Retrieve events one page at a time.
  var events
  var pageToken
  do {
    try {
      options.pageToken = pageToken
      events = Calendar.Events.list(calendarId, options)
    } catch (e) {
      // Check to see if the sync token was invalidated by the server;
      // if so, perform a full sync instead.
      if (e.message === "Sync token is no longer valid, a full sync is required.") {
        properties.deleteProperty(CONSTS.PROPERTIES.SYNC_TOKEN)
        lock.releaseLock()
        syncEvents(true)
        return
      } else {
        lock.releaseLock()
        throw new Error(e.message)
      }
    }
    
    if (!events.items || events.items.length === 0) {
      pageToken = events.nextPageToken
      continue
    }
    
    for (var i = 0; i < events.items.length; i++) {
      var event = events.items[i]
      var task = taskFromId(event.id)
      if (!task) {
        if (event.status === CONSTS.EVENT_STATUS.CANCELLED) {
          // User deleted a task we didn't know about
          continue
        }
        // This task is new!
        task = new Task(newRowIn(CONSTS.SHEETS.TASKS))
        // Create a backup of the event in case things go wrong
        var backupEvent = {}
        try {
          task.hydrateFromCalendarEvent(event)
          task.commitToRow()
          
          // Back up event data
          backupEvent.summary = event.summary
          backupEvent.description = event.description
          backupEvent.colorId = event.colorId
          backupEvent.start = {}
          backupEvent.start.date = event.start.date
          backupEvent.start.dateTime = event.start.dateTime
          backupEvent.start.timeZone = event.start.timeZone
          backupEvent.end = {}
          backupEvent.end.date = event.end.date
          backupEvent.end.dateTime = event.end.dateTime
          backupEvent.end.timeZone = event.end.timeZone
          
          
          event.summary = task.getCalendarTitle()
          event.description = task.getCalendarDescription()
          event.colorId = task.getEventColor()
          event.start.date = task.startDate.toISOString().slice(0,10)
          event.start.dateTime = null
          event.end.date = getRelativeDate(task.startDate, task.estimate).toISOString().slice(0,10)
          event.end.dateTime = null
          
          Calendar.Events.update(event, calendarId, event.id)
        } catch (e) {
          console.log("❌: Error (" + e + ")")
          // Remove the added row on failure
          removeNewRow(CONSTS.SHEETS.TASKS)
          event.summary = '[❌ Error : See Event Description] ' + backupEvent.summary
          event.description = CONSTS.APP_NAME + ' Error: ' + e + '\n\n' + backupEvent.description
          event.colorId = backupEvent.colorId
          event.start.date = backupEvent.start.date
          event.start.dateTime = backupEvent.start.dateTime
          event.start.timeZone = backupEvent.start.timeZone
          event.end.date = backupEvent.end.date
          event.end.dateTime = backupEvent.end.dateTime
          event.end.timeZone = backupEvent.end.timeZone
          Calendar.Events.update(event, calendarId, event.id)
        }
        continue
      } else if (event.status === CONSTS.EVENT_STATUS.CANCELLED) {
        // Event deleted!
        task.setId(null)
        task.startDate = ''
        task.completedOn = ''
        task.commitToRow()
        task._deleteRow()
        continue
      }
      // All other updates
      if (event.summary.indexOf(CONSTS.ACTIONS.CLOSE) > -1) {
        task.markAsCompleted()
        createOrUpdateTaskEvent(task)
      } else if (event.summary.indexOf(CONSTS.ACTIONS.OPEN) > -1) {
        // Not quite done 
        task.markAsIncomplete()
        createOrUpdateTaskEvent(task)
      } else if (event.summary.indexOf(CONSTS.ACTIONS.START) > -1) {
        task.setStartTime(new Date())
        createOrUpdateTaskEvent(task)
      } else if (event.summary.indexOf(CONSTS.ACTIONS.STOP) > -1) {
        var rawMinutesBetween = minutesBetweenDates(task.getStartTime(), new Date())
        var proratePref = PropertiesService.getUserProperties().getProperty(CONSTS.PROPERTIES.PRORATE)
        if (proratePref === CONSTS.TASK.PRORATES[0]) { // 'None' no prorate
          task.addMinutesToTimeSpent(rawMinutesBetween)
        } else {
          // This rounds up to the prorate
          var prorateChoice = parseInt(proratePref)
          task.addMinutesToTimeSpent(prorateChoice * Math.ceil(rawMinutesBetween/prorateChoice))
        }
        task.setStartTime()
        createOrUpdateTaskEvent(task)
      } else {
        var shouldUpdateEvent = false
        // Update the title in case it changed
        var newTitle = event.summary.slice(task.getTaskSymbol().length).trim()
        if (newTitle.toLowerCase() !== task.title.toLowerCase()) {
          task.title = newTitle
          shouldUpdateEvent = true
        }
        // Update description in case it changed and is still valid
        var fullDesc = event.description
        var beginningOfTaskDesc = fullDesc.lastIndexOf('<td>') + '<td>'.length
        var endOfTaskDesc = fullDesc.lastIndexOf('</td>')
        if (beginningOfTaskDesc > -1 && beginningOfTaskDesc < endOfTaskDesc) {
          task.description = fullDesc.substring(beginningOfTaskDesc, endOfTaskDesc).trim()
        }
        // Update Project+Milestone if necessary
        var beginningOfProjMilestone = fullDesc.indexOf('<td>') + '<td>'.length
        var endOfProjMilestone = fullDesc.indexOf('</td>')
        var projMilestoneFromEvent = fullDesc.substring(beginningOfProjMilestone, endOfProjMilestone).toLowerCase().trim()
        if (task.projectMilestone.toLowerCase() !== projMilestoneFromEvent) {
          // New project+milestone detected (possibly)
          var closestMatch = '' 
          var closestMatchDistance = 999999999
          var allProjMilestones = getAllProjectMilestones()
          for (var i = 0; i < allProjMilestones.length; i++) {
            var distance = levenshtein(projMilestoneFromEvent, allProjMilestones[i])
            if (distance < closestMatchDistance) {
              closestMatch = allProjMilestones[i]
              closestMatchDistance = distance
            }
          }
          if (task.projectMilestone !== closestMatch) {
            task.projectMilestone = closestMatch
            shouldUpdateEvent = true
          }
        }
        if (shouldUpdateEvent) {
          createOrUpdateTaskEvent(task)
        }
      }
      // Update estimate if available
      if (event.start.date && event.end.date) {
        task.startDate = parseDate(event.start.date)
        task.estimate = daysBetweenDates(task.startDate, parseDate(event.end.date))
      }
      task.commitToRow()
    }
    pageToken = events.nextPageToken
  } while (pageToken)
  properties.setProperty(CONSTS.PROPERTIES.SYNC_TOKEN, events.nextSyncToken)
  lock.releaseLock()
}

function itsABrandNewDay() {
  var allTasks = getAllDataRows(CONSTS.SHEETS.TASKS)
  var today = new Date()
  for (var i = CONSTS.TEMPLATE_ROW_IDX; i < allTasks.getHeight(); i++) {
    var task = new Task(allTasks.offset(i, 0, 1))
    if (task.isBlank()) {continue}
    if (task.getId() && task.startDate && task.etaDate <= today && task.completedOn === '') {
      task.estimate = 1 + daysBetweenDates(today, task.startDate)
      task.commitToRow()
      task.hydrate()
      createOrUpdateTaskEvent(task)
    }
  }
}
