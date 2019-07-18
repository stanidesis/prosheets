function onOpen() {
  createMenu()
  // upgradeIfNecessary()
}

function upgradeIfNecessary() {
  var upgradeAvailable = parseFloat(PropertiesService.getUserProperties().getProperty(CONSTS.PROPERTIES.VERSION) || '1.0') < 
    parseFloat(CONSTS.VERSION)
  
  if (!upgradeAvailable) {
    return
  }
  var ui = SpreadsheetApp.getUi()
  var response = ui.alert(CONSTS.UPGRADES.PROMPT, ui.ButtonSet.OK)
  if (response !== ui.Button.OK) {
    SpreadsheetApp.getUi().alert(CONSTS.UPGRADES.PROMPT_REJECTED)
    return
  }
  upgradeProSheets()
}

function createMenu() {
  var ui = SpreadsheetApp.getUi() 
  
  // Grab the current prorate
  var currentProrate = PropertiesService.getUserProperties().getProperty(CONSTS.PROPERTIES.PRORATE) || CONSTS.TASK.PRORATES[0]
  var prorateMenu = ui.createMenu(CONSTS.MENU.PRORATE_TITLE)
  for (var i = 0; i < CONSTS.TASK.PRORATES.length; i++) {
    prorateMenu.addItem(CONSTS.TASK.PRORATES[i] + (i > 0 && CONSTS.MENU.PRORATE_MINUTES || '') +
      (currentProrate === CONSTS.TASK.PRORATES[i] && CONSTS.MENU.CHECK || ''), 'setProrate' + CONSTS.TASK.PRORATES[i])
  }
  
  var timeTrackingMenu = ui.createMenu(CONSTS.MENU.TIME_TRACKING_TITLE)
  timeTrackingMenu.addSubMenu(prorateMenu)
  
  var menu = ui.createMenu(CONSTS.MENU.TITLE)
  .addItem(CONSTS.MENU.ABOUT, 'showAbout')
  .addSubMenu(timeTrackingMenu)
  
  if (!PropertiesService.getUserProperties().getProperty(CONSTS.PROPERTIES.FTUE)) {
    menu.addItem(CONSTS.MENU.SETUP, 'ftue')
  } else {
    menu.addItem(CONSTS.MENU.SET_CAL_ID, 'setCalId')
  }
  
  menu.addToUi()
}

function ftue() {
  var ui = SpreadsheetApp.getUi()
  var ftueCompleted = PropertiesService.getUserProperties().getProperty(CONSTS.PROPERTIES.FTUE)
  if (ftueCompleted) {
    ui.alert('All Set!', 
             'You’ve completed the setup process! If you feel you have made a mistake and need help, contact me at contact@stanleyidesis.com',
             ui.ButtonSet.OK)
    return
  }
  var response1 = ui.alert(
    'Welcome to ProSheets!',
    'This step-by-step setup will guide you through well, setting up ProSheets!\n\nYou will need:\n\n- Access to Google Calendar\n- Access to Cloud Console Projects\n\nYou may restart this tutorial at any time!\n\nClick OK to proceed to Step 1', 
    ui.ButtonSet.OK_CANCEL)
  if (response1 != ui.Button.OK) {
    return
  }
  var response4 = ui.alert(
    'Step 1: Enable Calendar API V3',
    'Befor you proceed, you must grant additional access to Google Calendar before ProSheets can synchronize your Events.\n\nFollow the instructions in the FAQ section or watch this: https://youtu.be/v84x2bxw0HU?t=48\n\n\n\nWhen you complete this step, restart Setup and click Ok',
    ui.ButtonSet.OK_CANCEL)
  if (response4 != ui.Button.OK) {
    return
  }
  var response2 = ui.alert(
    'Step 2: Create a Google Calendar', 
    'In Google Calendar, create (or choose an existing) Calendar to synchronize all of your Task Events with (we recommend a fresh start).\n\nWhen you\'ve chosen a Calendar, click Okay',
    ui.ButtonSet.OK_CANCEL)
  if (response2 != ui.Button.OK) {
    return
  }
  var stringOfCals = getAvailableCalendars().reduce(function(total, current) {
    return total + '- ' + current.summary + ' (id: ' + current.id + ')\n'
  },'')
  var response3 = ui.prompt(
    'Step 3: Provide Your Calendar’s ID', 
    'Find your calendar’s ID in the Calendar settings under the label, ’Calendar ID.’\nIt resembles an email address: abcdefghijklmnopqrstuvwxyz@group.calendar.google.com.\n\nHere are some calendars and IDs ProSheets was able to pull back:\n\n' + stringOfCals + '\n',
    ui.ButtonSet.OK_CANCEL)
  if (response3.getSelectedButton() != ui.Button.OK) {
    return
  }
  PropertiesService.getUserProperties().setProperty(
    CONSTS.PROPERTIES.CALENDAR_ID, response3.getResponseText())
  var calId = PropertiesService.getUserProperties().getProperty(CONSTS.PROPERTIES.CALENDAR_ID)
  var calendar
  try {
    calendar = getCalendar()
  } catch (e) {}
  if (!calendar) {
    ui.alert('Whoops! We failed to recover a Calendar for the provided ID: ' +
             calId +
             '\n\nPlease look the ID up once more and restart the Setup process.')
    PropertiesService.getUserProperties().deleteProperty(CONSTS.PROPERTIES.CALENDAR_ID)
    return
  }
  var response5 = ui.alert(
    'Step 4: Last Step!', 
    'ProSheets will now set up some triggers! Three to be precise:\n\n1. Automatically update Task Events as you edit the sheet\n2. Incorporate changes as they are received from your Calendar\n3. Automatically extend unfinished tasks to the next day\n\nClick Ok to set these snazzy triggers! (you can edit them under Tools > Script Editor ... Edit > Current project’s triggers)',
    ui.ButtonSet.OK_CANCEL)
  if (response5 != ui.Button.OK) {
    return
  }
  // Delete all triggers before making new ones!
  var allTriggers = ScriptApp.getProjectTriggers()
  for (var i = 0; i < allTriggers.length; i++) {
    ScriptApp.deleteTrigger(allTriggers[i])
  }
  var newDayTrigger = ScriptApp.newTrigger('itsABrandNewDay').timeBased().everyDays(1).atHour(1).create()
  var onEditTrigger = ScriptApp.newTrigger('onUserEdit').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create()
  var calendarUpdateTrigger = ScriptApp.newTrigger('onCalendarUpdate').forUserCalendar(getCalId()).onEventUpdated().create()
  // This should work now...
  syncOnce()
  PropertiesService.getUserProperties().setProperty(CONSTS.PROPERTIES.FTUE, 'true')
  var response6 = ui.alert(
    'Done!', 'You’re ready to use ProSheets!', 
    ui.ButtonSet.OK)
}

function showAbout() {
  var ui = SpreadsheetApp.getUi()
  ui.alert(
    'ProSheets is a personal project management tool designed and developed by Stanley Idesis.\n\nFor all inquiries: contact@stanleyidesis.com', 
    ui.ButtonSet.OK)
}

function setCalId(shouldRepeatAction) {
  var calId = PropertiesService.getUserProperties().getProperty(CONSTS.PROPERTIES.CALENDAR_ID)
  var ui = SpreadsheetApp.getUi()
  var prompt = 'Find your calendar’s ID in calendar settings under the label, ‘Calendar ID.’ It looks like abcdefghijklmnopqrstuvwxyz@group.calendar.google.com'
  if (calId === '' || calId === null) {
    prompt += '\n\nCurrent ID: unset'
  } else {
    prompt += '\n\nCurrent ID: ' + calId
  }
  if (shouldRepeatAction) {
    prompt += '\n\nAfter setting your ID, retry your last action 😊'
  }
  prompt += '\n\nHere are some calendars ProSheets found:\n\n' + getAvailableCalendars().reduce(function(total, current) {
    return total + '- ' + current.summary + ' (id: ' + current.id + ')\n'
  },'')
  
  var response = ui.prompt(CONSTS.MENU.SET_CAL_ID, 
            prompt,
            ui.ButtonSet.OK_CANCEL)
  if (response.getSelectedButton() == ui.Button.OK) {
    PropertiesService.getUserProperties()
       .setProperty(CONSTS.PROPERTIES.CALENDAR_ID, response.getResponseText())
  }
}

function setProrate(rate) {
  PropertiesService.getUserProperties().setProperty(CONSTS.PROPERTIES.PRORATE, rate)
  SpreadsheetApp.getUi().alert('Prorate updated to ' + rate + '! Refresh page to see an updated menu')
}

function setProrateNone() {
  setProrate(CONSTS.TASK.PRORATES[0])
}

function setProrate5() {
  setProrate(CONSTS.TASK.PRORATES[1])
}

function setProrate10() {
  setProrate(CONSTS.TASK.PRORATES[2])  
}

function setProrate15() {
  setProrate(CONSTS.TASK.PRORATES[3])
}

function setProrate20() {
  setProrate(CONSTS.TASK.PRORATES[4])
}

function setProrate30() {
  setProrate(CONSTS.TASK.PRORATES[5])
}

function getAvailableCalendars() {
  var calendarList = [];
  var calendars;
  var pageToken;
  do {
    calendars = Calendar.CalendarList.list({
      maxResults: 100,
      pageToken: pageToken
    });
    if (calendars.items && calendars.items.length > 0) {
      for (var i = 0; i < calendars.items.length; i++) {
        var calendar = calendars.items[i];
        calendarList.push({summary: calendar.summary, id: calendar.id})
      }
    } else {
      Logger.log('No calendars found.');
    }
    pageToken = calendars.nextPageToken;
  } while (pageToken);
  return calendarList;
}
  
