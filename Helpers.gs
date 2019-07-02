/**
 * Parses an RFC 3339 date or datetime string and returns a corresponding Date
 * object. This function is provided as a workaround until Apps Script properly
 * supports RFC 3339 dates. For more information, see
 * https://code.google.com/p/google-apps-script-issues/issues/detail?id=3860
 * @param {string} string The RFC 3339 string to parse.
 * @return {Date} The parsed date.
 */
function parseDate(string) {
  var parts = string.split('T');
  parts[0] = parts[0].replace(/-/g, '/');
  return getRelativeDate(new Date(parts.join(' ')), 0);
}

/**
 * Helper function to get a new Date object relative to the current date.
 * @param {number} daysOffset The number of days in the future for the new date.
 * @param {number} hour The hour of the day for the new date, in the time zone
 *     of the script.
 * @return {Date} The new date.
 */
function getRelativeDateOffset(daysOffset, hour) {
  var date = new Date();
  date.setDate(date.getDate() + daysOffset);
  date.setHours(hour);
  date.setMinutes(0);
  date.setSeconds(0);
  date.setMilliseconds(0);
  return date;
}

/**
 * Helper function to get a new Date object relative to the provided date.
 * @param {number} daysOffset The number of days in the future for the new date.
 * @return {Date} The new date.
 */
function getRelativeDate(startDate, daysOffset) {
  var date = new Date(startDate);
  date.setDate(date.getDate() + daysOffset);
  date.setHours(0);
  date.setMinutes(0);
  date.setSeconds(0);
  date.setMilliseconds(0);
  return date;
}

/**
 * Helper function to retreive the number of whole days between two dates.
 * @param {Date} date1 The first date
 * @param {Date} date2 The second date
 * @return {number} The number of days between date1 and date2.
 */
function daysBetweenDates(date1, date2) {
  var oneDay = 24*60*60*1000; // hours*minutes*seconds*milliseconds  
  return Math.round(Math.abs((date2.getTime() - date1.getTime())/(oneDay)));
}

/**
 * Helper function to retreive the number of minutes (rounded) between two dates.
 * @param {Date} date1 The first date
 * @param {Date} date2 The second date
 * @return {number} The number of minutes between date1 and date2.
 */
function minutesBetweenDates(date1, date2) {
  return Math.floor((Math.abs(date2 - date1)/1000)/60)
}

function getRowFromCell(cell) {
  var sheet = cell.getSheet()
  return sheet.getRange(cell.getRow(), 1, 1, sheet.getMaxColumns())
}

/**
 * Helper function that pulls a Task from any cell as long as the cell
 * is found at the right sheet
 * @param {Range} cell A cell found somewhere on the Task
 * @return {Task} the Task representing the full row
 */
function getTaskAtLocation(cell) {
  return new Task(getRowFromCell(cell))
}

/**
 * Helper function that pulls a Milestone from any cell as long as the cell
 * is found at the right sheet
 * @param {Range} cell A cell found somewhere on the Task
 * @return {Milestone} the Milestone representing the full row
 */
function getMilestoneAtLocation(cell) {
  return new Milestone(getRowFromCell(cell))
}

/**
 * Helper function that pulls a Milestone from any cell as long as the cell
 * is found at the right sheet
 * @param {Range} cell A cell found somewhere on the Task
 * @return {Project} the Project representing the full row
 */
function getProjectFromLocation(cell) {
  return new Project(getRowFromCell(cell))
}

/**
 * Levenshtein distance between two strings
 */
function levenshtein(a, b) {
	var tmp;
	if (a.length === 0) { return b.length; }
	if (b.length === 0) { return a.length; }
	if (a.length > b.length) { tmp = a; a = b; b = tmp; }

	var i, j, res, alen = a.length, blen = b.length, row = Array(alen);
	for (i = 0; i <= alen; i++) { row[i] = i; }

	for (i = 1; i <= blen; i++) {
		res = i;
		for (j = 1; j <= alen; j++) {
			tmp = row[j - 1];
			row[j - 1] = res;
			res = b[i - 1] === a[j - 1] ? tmp : Math.min(tmp + 1, Math.min(res + 1, row[j] + 1));
		}
	}
	return res;
}

function newRowIn(sheetName) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName)
  return sheet.insertRowBefore(CONSTS.TEMPLATE_ROW)
              .getRange(CONSTS.TEMPLATE_ROW, 1, 1, sheet.getMaxColumns())
}

function removeNewRow(sheetName) {
  removeRow(sheetName, CONSTS.TEMPLATE_ROW)
}

function removeRow(sheetName, row) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName)
  if (sheet.getMaxRows() === CONSTS.TEMPLATE_ROW) {
    var range = sheet.getRange(row, 1, 1, sheet.getMaxColumns())
    range.clearContent()
    range.clearNote()
    return
  }
  sheet.deleteRow(row)
}

function getAllDataRows(sheetName) {
  return SpreadsheetApp.getActive().getSheetByName(sheetName).getDataRange()
}

function getAllProjectMilestones() {
  var allProjMilestones = SpreadsheetApp.getActive().getRangeByName(CONSTS.NAMED_RANGES.PROJECT_MILESTONE_LIST).getValues()
  var projMilestoneArray = []
  for (var i = 0; i < allProjMilestones.length; i++) {
    if (allProjMilestones[i][0] === '') {
      break
    }
    projMilestoneArray.push(allProjMilestones[i][0])
  }
  return projMilestoneArray
}

function upgradeProSheets() {
  var currentVersion = PropertiesService.getUserProperties().getProperty(CONSTS.PROPERTIES.VERSION) || '1.0'
  switch (currentVersion) {
    case '1.0':
    default:
      // Upgrade to v1.1
      var taskSheet = SpreadsheetApp.getActive().getSheetByName(CONSTS.SHEETS.TASKS)
      taskSheet.insertColumnAfter(CONSTS.UPGRADES.V1_1.INSERT_COL_AFTER_IDX)
      taskSheet.setColumnWidth(CONSTS.UPGRADES.V1_1.INSERT_COL_AFTER_IDX,
                               CONSTS.UPGRADES.V1_1.TIME_SPENT_COL_WIDTH)
      
      var newColumnRange = taskSheet.getRange(CONSTS.UPGRADES.V1_1.TIME_SPENT_COL_FULL_A1)
      .clearDataValidations()
      
      // Cache then reset formatting
      var fontColors = newColumnRange.getFontColors()
      var fontFamilies = newColumnRange.getFontFamilies()
      var fontSizes = newColumnRange.getFontSizes()
      var fontStyles = newColumnRange.getFontStyles()
      var horizontalAlignments = newColumnRange.getHorizontalAlignments()
      var textStyles = newColumnRange.getTextStyles()
      var verticalAlignments = newColumnRange.getVerticalAlignments()
      newColumnRange.clear()
      newColumnRange.setFontColors(fontColors)
      .setFontFamilies(fontFamilies)
      .setFontSizes(fontSizes)
      .setFontStyles(fontStyles)
      .setHorizontalAlignments(horizontalAlignments)
      .setTextStyles(textStyles)
      .setVerticalAlignments(verticalAlignments)
      
      taskSheet.getRange(CONSTS.UPGRADES.V1_1.TIME_SPENT_COL_ALL_BUT_FIRST_A1)
      .setValue(0).setNumberFormat(CONSTS.UPGRADES.V1_1.TIME_SPENT_COL_NUM_FORMAT)
      
      taskSheet.getRange(CONSTS.UPGRADES.V1_1.TIME_SPENT_COL_FIRST_ROW_A1)
      .setValue(CONSTS.UPGRADES.V1_1.TIME_SPENT_COL_HEADER)
      
      // Repeat for completed tasks
      var completedTaskSheet = SpreadsheetApp.getActive().getSheetByName(CONSTS.SHEETS.COMPLETED_TASKS)
      completedTaskSheet.insertColumnAfter(CONSTS.UPGRADES.V1_1.INSERT_COL_AFTER_IDX)
      completedTaskSheet.setColumnWidth(CONSTS.UPGRADES.V1_1.INSERT_COL_AFTER_IDX,
                               CONSTS.UPGRADES.V1_1.TIME_SPENT_COL_WIDTH)
      
      newColumnRange = completedTaskSheet.getRange(CONSTS.UPGRADES.V1_1.TIME_SPENT_COL_FULL_A1)
      .clearDataValidations()
      
      // Cache then reset formatting
      var fontColors = newColumnRange.getFontColors()
      var fontFamilies = newColumnRange.getFontFamilies()
      var fontSizes = newColumnRange.getFontSizes()
      var fontStyles = newColumnRange.getFontStyles()
      var horizontalAlignments = newColumnRange.getHorizontalAlignments()
      var textStyles = newColumnRange.getTextStyles()
      var verticalAlignments = newColumnRange.getVerticalAlignments()
      newColumnRange.clear()
      newColumnRange.setFontColors(fontColors)
      .setFontFamilies(fontFamilies)
      .setFontSizes(fontSizes)
      .setFontStyles(fontStyles)
      .setHorizontalAlignments(horizontalAlignments)
      .setTextStyles(textStyles)
      .setVerticalAlignments(verticalAlignments)
      
      completedTaskSheet.getRange(CONSTS.UPGRADES.V1_1.TIME_SPENT_COL_ALL_BUT_FIRST_A1)
      .setValue(0).setNumberFormat(CONSTS.UPGRADES.V1_1.TIME_SPENT_COL_NUM_FORMAT)
      
      completedTaskSheet.getRange(CONSTS.UPGRADES.V1_1.TIME_SPENT_COL_FIRST_ROW_A1)
      .setValue(CONSTS.UPGRADES.V1_1.TIME_SPENT_COL_HEADER)
  }
  PropertiesService.getUserProperties().setProperty(CONSTS.PROPERTIES.VERSION, 
                                                    CONSTS.VERSION)
  SpreadsheetApp.getUi().alert('ProSheets upgraded to version ' + CONSTS.VERSION + '!\n\nPlease refresh this page to continue.')
}
