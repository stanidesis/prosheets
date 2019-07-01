var CONSTS = {
  MENU: {
    TITLE: 'ProSheets Menu',
    TIME_TRACKING_TITLE: 'Time Tracking',
    PRORATE_TITLE: 'Prorate',
    PRORATE_MINUTES: ' min',
    ABOUT: 'About',
    SETUP: 'Setup',
    SET_CAL_ID: 'Set Calendar ID',
    CHECK: ' ✓'
  },
  TASK: {
    JSON_NOTE_INDEX: 1,
    CHAR_OPEN: '⭕',
    CHAR_CLOSED: '✅',
    CHAR_IN_PROGRESS: '🏁',
    COLOR_OPEN: CalendarApp.EventColor.CYAN,
    COLOR_CLOSED: CalendarApp.EventColor.PALE_GREEN,
    COLOR_IN_PROGRESS: CalendarApp.EventColor.YELLOW,
    TIME_SPENT_STARTING_POINT: 'Sat Dec 30 1899 00:00:00 GMT-0600 (CST)',
    JSON_KEYS: {
      ID: 'id',
      START_TIMESTAMP: 'start_timestamp'
    },
    PRORATES: ['None', '5', '10', '15', '20', '30'],
    PROJ_MILESTONE_HEADER: 'Project+Milestone',
    TIME_SPENT_HEADER: 'Time Spent',
    DESCRIPTION_HEADER: 'Description',
    DESCRIPTION_FOOTER:     
    '<h2>Instructions</h2>' +
    
    '<h3>Editing Tasks</h3>' +
    
    'You can perform an action (start/stop/open/close) or edit the task details, not both at once.\n\n' +
    'So follow this flow: add an action to the title, save event. Edit the title or description, save event.\n' +
    
    '<h3>Edit Task Title</h3>' +
    
    'Change the summary of this event to change the title of your task.\n\n' +
    'Do not modify or delete the status symbol (STAT_1, STAT_2, STAT_3).\n' +

    '<h3>Edit Task Description</h3>' +
    
    'Edit the description cell above to change the description for your task.\n' +
    
    '<h3>Perform an Action</h3>' +
    '<ul>' +
    '<li><strong>Begin Time Tracking</strong>: Add "(start)" to the summary, then save</li>' +
    '<li><strong>Stop Time Tracking</strong>: Add "(stop)" to the summary, then save</li>' +
    '<li><strong>Mark as Complete</strong>: Add "(close)" to the summary, then save</li>' +
    '<li><strong>Mark as Incomplete</strong>: Add "(open)" to the summary, then save</li>' +
    '</ul>' +
    
    '<blockquote><em>Remember, you can perform an action or edit the task from your calendar, but not at the same time!</em></blockquote>' +
    
    '<h3>Reassign Project and Milestone</h3>' +
    'Change the <strong>Project+Milestone</strong> cell above to one of the following:\n' +
    
    '<ul>',
    PROJECT_MILESTONE_TEMPLATE: '<li>%s</li>',
    PROJECT_MILESTONE_END: '</ul>'
  },
  SHEETS: {
    PROJECTS: 'Projects',
    MILESTONES: 'Milestones',
    TASKS: 'Tasks',
    COMPLETED_PROJECTS: 'Completed Projects',
    COMPLETED_MILESTONES: 'Completed Milestones',
    COMPLETED_TASKS: 'Completed Tasks',
    RANGES: 'Ranges'
  },
  ACTIONS: {
    CLOSE: '(close)',
    OPEN: '(open)',
    START: '(start)',
    STOP: '(stop)'
  },
  STATUS: {
    ACTIVE: 'Active',
    BACKLOG: 'Backlog',
    BLOCKED: 'Blocked',
    COMPLETED: 'Completed',
    DELETE: 'Delete',
    MARK_AS_COMPLETE: 'Mark as Complete'
  },
  PROPERTIES: {
    CALENDAR_ID: 'calID',
    SYNC_TOKEN: 'syncToken',
    FTUE: 'ftue',
    VERSION: 'version',
    PRORATE: 'prorate'
  },
  EVENT_STATUS: {
    CANCELLED: 'cancelled',
    CONFIRMED: 'confirmed'
  },
  NAMED_RANGES: {
    PRIORITY_MATRIX: 'PriorityMatrix',
    PROJECT_LIST: 'ProjectList',
    PROJECT_MILESTONE_LIST: 'ProjectMilestoneList',
    PROJECT_TYPE: 'ProjectType',
    STATUS: 'Status'
  },
  UPGRADES: {
    PROMPT: 'You must update ProSheets before you continue!\n\nPress \'Ok\' to perform this quick update now.',
    PROMPT_REJECTED: '☠: Proceed at your own risk, ProSheets may not work as expected...',
    V1_1: {
      INSERT_COL_AFTER_IDX: 5,
      TIME_SPENT_COL_WIDTH: 110,
      TIME_SPENT_COL_HEADER: 'Time Spent',
      TIME_SPENT_COL_NUM_FORMAT: 'hh"h" mm"m"',
      TIME_SPENT_COL_FULL_A1: 'F:F',
      TIME_SPENT_COL_FIRST_ROW_A1: 'F1',
      TIME_SPENT_COL_ALL_BUT_FIRST_A1: 'F2:F'
    }
  },
  NA: 'N/A',
  TEMPLATE_ROW: 2,
  TEMPLATE_ROW_IDX: 1,
  APP_NAME: 'ProSheets',
  VERSION: '1.1'
}
