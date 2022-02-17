// setup
function setup() {
  const groupUrl = '?????@googlegroups.com'; // replace this line
  if (groupUrl.includes('?')) { // detect default value
    throw new Error('ERROR: change "?????@googlegroups.com" to your google group name');
  }
  createSpreadsheet(18); // create spreadsheet for 18 users
  createCalendars(18, groupUrl); // create 19 read + 18 write calendars
  createTriggers();
}

// set row and column count of sheet
function changeSheetSize(sheet, rows, columns) {
  // column count will most likely be decreased
  // we will change columns first to prevent cell count to exceed 10000000 cell limit
  if (columns !== 0) { 
    var sheetColumns = sheet.getMaxColumns();
    if (columns < sheetColumns) {
      sheet.deleteColumns(columns+1, sheetColumns-columns);
    }
    else if (columns > sheetColumns) {
      sheet.insertColumnsAfter(sheetColumns,columns-sheetColumns);
    }
  }
  if (rows !== 0) {
    var sheetRows = sheet.getMaxRows();
    if (rows < sheetRows) {
      sheet.deleteRows(rows+1, sheetRows-rows);
    }
    else if (rows > sheetRows) {
      sheet.insertRowsAfter(sheetRows, rows-sheetRows);
    }
  }
  sheet.getRange(1, 1, rows, columns).setWrap(true); // wrap overflowing text
}

function arrayFill2d(rows, columns, value) { // create 2d array filled with value
  return Array(rows).fill().map(() => Array(columns).fill(42));
}

// creates spreadsheet for {userCount} users
function createSpreadsheet(userCount) {
  // create spreadsheet for configuration
  const equipmentCount = 50; // number of equipments
  const experimentConditionCount = 20 // number of experiment conditions for a single equipment
  const experimentConditionRows = 6000 // number of rows in experiment condition
  var configSpreadsheet = SpreadsheetApp.create('configSpreadsheet');
  // users sheet
  var activeSheet = configSpreadsheet.insertSheet('users');
  configSpreadsheet.deleteSheet(configSpreadsheet.getSheetByName('Sheet1'));
  changeSheetSize(activeSheet, userCount+2, equipmentCount+9);
  activeSheet.hideColumns(2, 6); // hide columns used for debug
  // draw borders
  activeSheet.getRange(1, 1, userCount+2, equipmentCount+9).setBorder(true, true, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
  activeSheet.getRange(1, 1, userCount+2, equipmentCount+9).setBorder(true, true, true, true, null, true, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  activeSheet.getRange(1, 1, 1, equipmentCount+9).setBorder(null, null, true, null, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
  activeSheet.getRange(1, 1, userCount+2, 1).setBorder(null, null, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
  activeSheet.getRange(1, 5, userCount+2, 1).setBorder(null, null, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  activeSheet.getRange(1, 7, userCount+2, 1).setBorder(null, null, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  activeSheet.getRange(1, 9, userCount+2, 1).setBorder(null, null, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  activeSheet.getRange(1, 1, 1, 9).setValues(
    [['Full Name (EDIT this line)', 'Last Name', 'First Name', 'User Name 1', 'User Name 2', 'Read CalendarId', 'Write CalendarId', 'Read Calendar URL', 'Write Calendar URL']]
  );
  // normal user row
  activeSheet.getRange(2, 10, userCount, equipmentCount).insertCheckboxes('no'); // create unchecked checkbox for 100 columns (equipments)
  var filledArray = arrayFill2d(userCount, 1, 'First Last');
  activeSheet.getRange(2, 1, userCount).setValues(filledArray);
  // 'ALL EVENTS' user row
  activeSheet.getRange(2+userCount, 10, 1, equipmentCount).insertCheckboxes('yes'); // create checked checkbox for 'ALL EVENTS'
  activeSheet.getRange(2+userCount, 1).setValue('ALL EVENTS');
  // copy equipments name from properties sheet
  var filledArray = arrayFill2d(1, equipmentCount, `=properties!R${2+i}C${1}`); // refer to sheet 'properties' for equipment name
  activeSheet.getRange(1, 10, 1, equipmentCount).setFormulas(filledArray);

  // properties sheet
  var activeSheet = configSpreadsheet.insertSheet('properties');
  changeSheetSize(activeSheet, equipmentCount+1, experimentConditionCount+1;
  // draw borders
  activeSheet.getRange(1, 1, equipmentCount+1, experimentConditionCount).setBorder(true, true, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
  activeSheet.getRange(1, 1, 1, experimentConditionCount).setBorder(null, null, true, null, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
  activeSheet.getRange(1, 1, equipmentCount+1, 1).setBorder(null, null, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
  var filledArray = [[]];
  filledArray[0] = ['equipmentName', 'sheetName', 'Properties ->'];
  for (var i = 0; i < equipmentCount; i++) {
    filledArray[i+1] = ['', `equipment ${i+1}`, ''];
  }
  activeSheet.getRange(1, 1, equipmentCount+1, 3).setValues(filledArray);
  activeSheet.hideColumns(2); // hide columns used for debug

  // create spreadsheet for experiment condition logging
  var experimentConditionSpreadsheet = SpreadsheetApp.create('experimentConditionSpreadsheet');
  var activeSheet = experimentConditionSpreadsheet.insertSheet('eventLog');
  for (var i = 0; i < equipmentCount; i++) { // create sheet for each equipment
    Utilities.sleep(1000);
    var activeSheet = experimentConditionSpreadsheet.insertSheet(`equipment ${i+1}`);
    if (i === 0) {        
      experimentConditionSpreadsheet.deleteSheet(experimentConditionSpreadsheet.getSheetByName('Sheet1'));
    }
    changeSheetSize(activeSheet, 10000, 15)
    activeSheet.getRange(1, 1, 1, 12).setValues(
      [['startTime', 'endTime', 'name', 'equipment', 'status', 'description', 'isAllDayEvent', 'isRecurringEvent', 'action', 'executionTime', 'id', 'eventExists']]
    );
    activeSheet.getRange(1, 13).setValue(sheet.getName()); // set sheet name
    var filledArray = [[]];
    for (var j = 0; j < 10000-1; j++) {
      filledArray[j] = [];
      // see if event exists (if it is 1[unmodified(is the last entry with the same id)] and 2[not canceled])
      filledArray[j][12] = `=AND(COUNTIF(INDIRECT("R[1]C[-1]", FALSE):INDIRECT("R[1]C[-1]", FALSE),B4)=0, INDIRECT("R[0]C[-3]", FALSE)="add")`;
    }
    activeSheet.getRange(2, 1, experimentConditionRows, 12).setFormulas(filledArray);
    sheet.setColumnFilterCriteria(12, ); // todo
  }
  
  // create spreadsheet for finalized logging
  var loggingSpreadsheet = SpreadsheetApp.create('loggingSpreadsheet');
  loggingSpreadsheet.insertSheet('finalLog');
  loggingSpreadsheet.deleteSheet(loggingSpreadsheet.getSheetByName('Sheet1'));
  var activeSheet = loggingSpreadsheet.getSheetByName('finalLog');  
  changeSheetSize(activeSheet, 100000, 11);
  activeSheet.getRange(1, 1, 1, 11).setValues(
    [['startTime', 'endTime', 'name', 'equipment', 'status', 'description', 'isAllDayEvent', 'isRecurringEvent', 'action', 'executionTime', 'id']]
  );

  var property = {
    configSpreadsheetId : configSpreadsheet.getId(),
    experimentConditionSpreadsheetId : experimentConditionSpreadsheet.getId(),
    loggingSpreadsheetId : loggingSpreadsheet.getId(),
  };
  setIds(property);
}

// creates calendars for {userCount} users
function createCalendars(userCount, groupUrl) {
  const properties = PropertiesService.getUserProperties();
  const configSpreadsheet = SpreadsheetApp.openById(properties.getProperty('configSpreadsheetId'));
  const resource = { // used to add google group as guest
    'scope': {
      'type': 'group',
      'value': groupUrl,
    },
    'role': 'writer',
  }
  // create {userCount+1} read calendars
  for (var i = 0; i < userCount+1; i++){
    Utilities.sleep(3000);
    var calendar = CalendarApp.createCalendar(`read ${i+1}`);
    var readCalendarId = calendar.getId();
    Calendar.Acl.insert(resource, readCalendarId); // add access permission to google group
    var activeSheet = configSpreadsheet.getSheetByName('users');
    activeSheet.getRange(2+i, 6).setValue(readCalendarId);
    activeSheet.getRange(2+i, 8).setValue(`https://calendar.google.com/calendar/u/0?cid=${readCalendarId}`);
    Logger.log(`Created read calendar ${calendar.getName()}, with the ID ${readCalendarId}.`);
  }
  // create {userCount} write calendars
  for (var i = 0; i < userCount; i++){
    Utilities.sleep(3000);
    var calendar = CalendarApp.createCalendar(`write ${i+1}`);
    var writeCalendarId = calendar.getId();
    Calendar.Acl.insert(resource, writeCalendarId); // add access permission to google group
    var activeSheet = configSpreadsheet.getSheetByName('users');
    activeSheet.getRange(2+i, 7).setValue(writeCalendarId);
    activeSheet.getRange(2+i, 9).setValue(`https://calendar.google.com/calendar/u/0?cid=${writeCalendarId}`);
    Logger.log(`Created write calendar ${calendar.getName()}, with the ID ${writeCalendarId}.`);
  }
}

// set ids
function setIds(property) {
  const properties = PropertiesService.getUserProperties();
  properties.setProperty('configSpreadsheetId', property.configSpreadsheetId);
  properties.setProperty('experimentConditionSpreadsheetId', property.experimentConditionSpreadsheetId);
  properties.setProperty('loggingSpreadsheetId', property.loggingSpreadsheetId);
}

function setIdsManual() {
  const properties = PropertiesService.getUserProperties();
  //property = {
  //  configSpreadsheetId : ;
  //  experimentConditionSpreadsheetId : ;
  //  loggingSpreadsheetId : ;
  //}
  properties.setProperty('configSpreadsheetId', property.configSpreadsheetId);
  properties.setProperty('experimentConditionSpreadsheetId', property.experimentConditionSpreadsheetId);
  properties.setProperty('loggingSpreadsheetId', property.loggingSpreadsheetId);
}

// delete all triggers for this script
function deleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++){
    trigger = triggers[i];
    ScriptApp.deleteTrigger(trigger);
  }
}

// create triggers
// only 20 triggers can be made for single script
// we will use 18 for write calendars, 1 for daily logging, 1 for spreadsheet
function createTriggers() {
  const properties = PropertiesService.getUserProperties();
  const sheet = SpreadsheetApp.openById(properties.getProperty('configSpreadsheetId')).getSheetByName('users');
  const writeCalendarIds = getWriteCalendarIds(sheet);
  
  // create trigger for each of the 18 write calendars
  // (calls function 'onCalendarEdit' on trigger)
  Logger.log(writeCalendarIds.length + ' calendar triggers will be created');
  for (var i = 0; i < writeCalendarIds.length; i++){
    const writeCalendarId = writeCalendarIds[i];
    ScriptApp.newTrigger('onCalendarEdit')
      .forUserCalendar(writeCalendarId)
      .onEventUpdated()
      .create(); 
  }
  // create 1 Sheets trigger (calls function 'onSheetsEdit' on trigger)
  ScriptApp.newTrigger('onSheetsEdit')
      .forSpreadsheet(properties.getProperty('configSpreadsheetId'))
      .onEdit()
      .create();
  // create 1 Sheets trigger for daily logging past events
  ScriptApp.newTrigger('finalLogging')
    .timeBased()
    .atHour(4) // 4:00
    .nearMinute(0)  
    .everyDays(1) 
    .create();
}

// when calendar gets edited
function onCalendarEdit(e) {
  const properties = PropertiesService.getUserProperties();
  const sheet = SpreadsheetApp.openById(properties.getProperty('configSpreadsheetId')).getSheetByName('users');
  const calendarId = e.calendarId;
  const writeCalendarIds = getWriteCalendarIds(sheet);
  const index = writeCalendarIds.indexOf(calendarId);
  const fullSync = false;
  eventLoggingSetup(); // setup admin logging
  writeEventsToReadCalendar(sheet, calendarId, index, fullSync);
}

// when sheets gets edited
function onSheetsEdit(e) {
  const properties = PropertiesService.getUserProperties();
  const book = SpreadsheetApp.openById(properties.getProperty('configSpreadsheetId'));
  const sheet = e.source.getActiveSheet();
  const cell = e.source.getActiveRange();
  const newValue = e.value;
  const row = cell.getRow();
  const column = cell.getColumn();
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const users = getUsers(sheet);
  const index = row-2;
  const readUser = users[index];
  const writeCalendarIds = getWriteCalendarIds(sheet);
  const calendarId = writeCalendarIds[index]
  const fullSync = true;
  
  eventLoggingSetup(); // setup admin logging
  if (row == 1 || row > lastRow || column > lastColumn || (column > 1 && column < 6)){
    action = 'edited invalid area';
  } else if (column === 1) {
    action = 'edited name';
  } else if (column === 6) {
    action = 'edited read calendarId';
  } else if (column === 7) {
    action = 'edited write calendarId';
  } else if (column > 7) {
    action = 'edited equipment';
  } 
  const executionTime = new Date(); // current time
  eventLoggingStoreData({
    executionTime: executionTime,
    name: readUser,
    action: `sheets (row, col) = (${row}, ${column}): ${action}`,
  });  
  // when the checkbox (H2~nm) is edited in sheets on sheet 'users'
  // update corresponding user's subscribed equipments
  if (sheet.getName() === 'users' && row > 1 && column > 7){ 
    changeSubscribedEquipments(sheet, readUser, index, users);
  }
  // when the full name (A2~An) is edited in sheets on sheet 'users'
  // update all of the corresponding user's event title
  else if (sheet.getName() === 'users' && row > 1 && column === 1){
    updateCalendarUserName(sheet, cell, newValue);
    writeEventsToReadCalendar(sheet, calendarId, index, fullSync);
  }
}

function eventLoggingSetup() { // prepares constant for eventLogging
  const properties = PropertiesService.getUserProperties();
  book = SpreadsheetApp.openById(properties.getProperty('experimentConditionSpreadsheetId')); // spreadsheet for loggingrepares constant for eventLogging
  const lastRow = book.getSheetByName('eventLog').getLastRow();
  const row = lastRow + 1; // write on new row
  properties.setProperty('row', row.toString());
  properties.setProperty('eventLoggingData', JSON.stringify({}));
}

function eventLoggingStoreData(logObj) { // set data for logging
  const properties = PropertiesService.getUserProperties();
  var eventLoggingData = JSON.parse(properties.getProperty('eventLoggingData'));
  for (const key in logObj) { // iterate through log object
    eventLoggingData[key] = value;
  }
  properties.setProperty('eventLoggingData', JSON.stringify(eventLoggingData));
}

function eventLoggingExecute(logObj) { // execute logging to sheets
  const properties = PropertiesService.getUserProperties();
  const columnDescriptions = { // shows which description corresponds to which column
    startTime: 1,
    endTime: 2,
    name: 3,
    equipment: 4,
    status: 5,
    description: 6,
    isAllDayEvent: 7,
    isRecurringEvent: 8,
    action: 9,
    executionTime: 10,
    id: 11,
  };
  const row = parseInt(properties.getProperty('row'));
  var eventLoggingData = JSON.parse(properties.getProperty('eventLoggingData'));
  const eventLogSheet = SpreadsheetApp.openById(properties.getProperty('SpreadsheetId')).getSheetByName('eventLog'); // spreadsheet for logging
  var filledArray = [[]];
  for (const key in logObj) { // iterate through log object
    var value = logObj[key];
    var col = columnDescriptions[key];
    filledArray[0][col-1] = value;
  }
  eventLogSheet.getRange(row, 1, 1, 11).setValue(filledArray);
}

function finalLogging() { // logs just the necessary data
  const properties = PropertiesService.getUserProperties();
  const finalLogSheet = SpreadsheetApp.openById(properties.getProperty('loggingSpreadsheetId')).getSheetByName('finalLog')
  const lastRow = finalLogSheet.getLastRow();
  const row = lastRow + 1; // write on new row
  const columnDescriptions = { // shows which description corresponds to which column
    startTime: 1,
    endTime: 2,
    name: 3,
    equipment: 4,
    status: 5,
    description: 6,
    isAllDayEvent: 7,
    isRecurringEvent: 8,
    action: 9,
    executionTime: 10,
    id: 11,
  };
  const sheet = SpreadsheetApp.openById(properties.getProperty('configSpreadsheetId')).getSheetByName('users');
  const writeCalendarIds = getWriteCalendarIds(sheet);
  const users = getUsers(sheet);
  // get events from 2~3 days ago
  options = {}
  options.timeMin = getRelativeDate(-3, 0).toISOString(); // 3 days ago
  options.timeMax = getRelativeDate(-2, 0).toISOString(); // 2 days ago
  for (var i = 0; i < writeCalendarIds.length; i++){ // iterate through every write calendar id
    Utilities.sleep(100);
    const writeCalendarId = writeCalendarIds[i];
    const writeUser = users[i];
    const eventsList = Calendar.Events.list(writeCalendarId, options);
    var events = []; 
    // get events from eventsList format
    if (eventsList.items && eventsList.items.length > 0) {
      for (var j = 0; j < eventsList.items.length; j++) {
        const event = eventsList.items[j];
        if (event.status === 'cancelled') {
          Logger.log(`Event id ${event.id} was cancelled.`);
        } else{
          events.push(event);
        }
      }
    } 

    if (events.length === 0) {
      Logger.log(`No events found for: ${writeCalendarId}`);
    } else {
      Logger.log(`${events.length} events found for: ${writeCalendarId}`);
    }
    
    // log each event
    for (var j = 0; j < events.length; j++){
      Utilities.sleep(100);
      var event = events[j];
      const equipmentStateFromEvent = getEquipmentStateFromEvent(event); // this must be called first before getCalendarById
      const equipment = equipmentStateFromEvent.equipment;
      const state = equipmentStateFromEvent.state;
      const eid = event.iCalUID;
      event = CalendarApp.getCalendarById(writeCalendarId).getEventById(eid);
      var logObj = {
          startTime: event.getStartTime(),
          endTime: event.getEndTime(),
          name: writeUser,
          equipment: equipment,
          status: state,
          description: event.getDescription(),
          isAllDayEvent: event.isAllDayEvent(),
          isRecurringEvent: event.isRecurringEvent(),
          action: '',
          executionTime: '',
          id: eid,
      }
      var filledArray = [[]];
      for (const key in logObj) { // iterate through log object
        var value = logObj[key];
        var col = columnDescriptions[key];
        filledArray[0][col-1] = value;
      }
      finalLogSheet.getRange(row, 1, 1, 11).setValue(filledArray);
    }
  }
}

// filter readUsers who are not writeUser and have the equipment
function filterUsers(writeUser, event, readCalendarIds, users, enabledEquipmentsList) {
  var filteredReadCalendarIds = []; // readCalendarIds excluding the same user as writeCalendarId
  var filteredReadUsers = []; // readUsers excluding the same user as writeCalendarId
  const equipment = getEquipmentStateFromEvent(event).equipment; // equipment used in the event
  for (var i = 0; i < readCalendarIds.length; i++){
    const readCalendarId = readCalendarIds[i];
    const readUser = users[i];
    const enabledEquipments = enabledEquipmentsList[i];
    if (enabledEquipments.includes(equipment) === true){ // if equipment is enabled for readUser
      if (readUser != writeUser) { // avoid duplicating event for same user's read and write calendars      
        filteredReadCalendarIds.push(readCalendarId);
        filteredReadUsers.push(readUser);
      }
    }
  }
  return {filteredReadCalendarIds, filteredReadUsers}
}

// write events to read calendar based on updated events in write calendar
function writeEventsToReadCalendar(sheet, writeCalendarId, index, fullSync) {
  const readCalendarIds = getReadCalendars(sheet).readCalendarIds;
  const enabledEquipmentsList = getReadCalendars(sheet).enabledEquipmentsList;
  const users = getUsers(sheet);
  const writeUser = users[index];
  const events = getEvents(writeCalendarId, fullSync);
  const eid = events.getId();
  Logger.log(`${readCalendarIds.length} read calendars`);
  for (var i = 0; i < events.length; i++){
    const event = events[i];
    const filteredReadCalendarIds = filterUsers(writeUser, event, readCalendarIds, users, enabledEquipmentsList).filteredReadCalendarIds;
    Logger.log(`writing event no.${i+1} to [ ${filteredReadCalendarIds} ]`);
    writeEvent(event, writeCalendarId, writeUser, filteredReadCalendarIds); // create event in write calendar and add read calendars as guests
    if (Calendar.Events.get(writeCalendarId, eid).status === 'cancelled') {
      var action = 'cancel';
    } else {
      var action = 'add';
    } // todo: add modified status
    eventLoggingStoreData({ // log event
      startTime: event.getStartTime(),
      endTime: event.getEndTime(),
      name: writeUser,
      equipment: equipment,
      status: state,
      description: event.getDescription(),
      isAllDayEvent: event.isAllDayEvent(),
      isRecurringEvent: event.isRecurringEvent(),
      action: action, 
      executionTime: new Date(), // current time
      id: eid, 
    });
    eventLoggingExecute();
    Logger.log('event logging done');
  }
  updateSyncToken(writeCalendarId); // renew sync token after adding guest
  Logger.log(`Wrote updated events to read calendar. Fullsync = ${fullSync}`);
}

// update corresponding user's subscribed equipments 
function changeSubscribedEquipments(sheet, readUser, index, users){
  const fullSync = true;
  const readCalendarIds = getReadCalendars(sheet).readCalendarIds;
  const enabledEquipmentsList = getReadCalendars(sheet).enabledEquipmentsList;
  const writeCalendarIds = getWriteCalendarIds(sheet);
  const readCalendarId = readCalendarIds[index];
  const enabledEquipments = enabledEquipmentsList[index];
  Logger.log(`${writeCalendarIds.length} write calendars`);
  for (var i = 0; i < writeCalendarIds.length; i++){
    const writeUser = users[i];
    const writeCalendarId = writeCalendarIds[i];
    const events = getEvents(writeCalendarId, fullSync);
    for (var j = 0; j < events.length; j++){
      const event = events[j];
      const filteredReadCalendarIds = filterUsers(writeUser, event, [readCalendarId], users, enabledEquipmentsList).filteredReadCalendarIds;
      writeEvent(event, writeCalendarId, writeUser, filteredReadCalendarIds); // create event in write calendar and add read calendars as guests
    }
    updateSyncToken(writeCalendarId);
  }
  Logger.log('Changed subscribed equipments');
}
  
// update calendar's User Name based on full name input
function updateCalendarUserName(sheet, cell, newValue){  
  setFirstLastNames(sheet, cell, newValue); // get last and first names from full name and set User Name 1
  setUserNames(sheet); // set User Name 2 using User Name 1
  setCheckboxes(sheet, cell); // create checkboxes for selecting equipment
  setCalendars(sheet, cell); // set read calendar and write calendar for created user
  Logger.log('Updated user name');
}

// get all the User Names
function getUsers(sheet) {
  const users = [];
  const lastRow = sheet.getLastRow();
  for (var row = 2; row < lastRow+1; row++) { 
    const user = sheet.getRange(row, 5).getValue();
    users.push(user);
  }
  return users;
}

// get all the write calendar's calendarIds
function getWriteCalendarIds(sheet) {
  const calendarIds = [];
  const lastRow = sheet.getLastRow();
  for (var row = 2; row < lastRow+1; row++) { 
    calendarId = sheet.getRange(row, 7).getValue();
    calendarIds.push(calendarId);
  }
  return calendarIds;
}

// get all the read calendar's calendarIds
function getReadCalendars(sheet) {
  const calendarIds = [];  
  const enabledEquipmentsList = [];
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  // get calendarId and enabledEquipment
  for (var row = 2; row < lastRow+1; row++) { 
    // get calendarId and add to calendarIds
    const calendarId = sheet.getRange(row, 6).getValue();
    calendarIds.push(calendarId);
    // get enabledEquipment and add to enabledEquipments
    const enabledEquipments = [];
    for (var column = 10; column < lastColumn+1; column++) { 
      const equipment = sheet.getRange(1, column).getValue();
      const checked = sheet.getRange(row, column).isChecked();
      if (checked === true){
        enabledEquipments.push(equipment);
      }
    }
    enabledEquipmentsList.push(enabledEquipments);
  }
  return {
    readCalendarIds: calendarIds,
    enabledEquipmentsList: enabledEquipmentsList,
  };
}

// get events from the given calendar that have been modified since the last sync.
// if the sync token is missing or invalid, log all events from up to a ten days ago (a full sync).
function getEvents(calendarId, fullSync) {
  const properties = PropertiesService.getUserProperties();
  const options = {
    maxResults: 100
  };
  const syncToken = properties.getProperty(`syncToken ${calendarId}`);
  Logger.log(`Current sync token: ${syncToken}`);
  if (syncToken && !fullSync) {
    options.syncToken = syncToken;
  } else {
    // Sync events up to ten days in the past.
    options.timeMin = getRelativeDate(-10, 0).toISOString();
  }

  // Retrieve events one page at a time.
  var eventsList;
  var pageToken = null;
  var events = [];
  do {
    try {
      if (pageToken === null) { // first page
        delete options.pageToken; // delete key 'pageToken'
      } else {
        options.pageToken = pageToken;
      }
      eventsList = Calendar.Events.list(calendarId, options);
    } catch (e) {
      // Check to see if the sync token was invalidated by the server;
      // if so, perform a full sync instead.
      if (e.message === 'API call to calendar.events.list failed with error: Sync token is no longer valid, a full sync is required.') {
        Logger.log('Sync token invalidated -> Full sync initiated');
        properties.deleteProperty(`syncToken ${calendarId}`);
        events = getEvents(calendarId, true);
        return events;
      } else {
        throw new Error(e.message);
      }
    }
    if (eventsList.items && eventsList.items.length > 0) {
      for (var i = 0; i < eventsList.items.length; i++) {
        const event = eventsList.items[i];
        if (event.status === 'cancelled') {
          Logger.log('Event id %s was cancelled.', event.id);
        } else{
          events.push(event);
        }
      }
    } else {
      Logger.log('No events found.');
    }
    pageToken = eventsList.nextPageToken;
  } while (pageToken);
  properties.setProperty(`syncToken ${calendarId}`, eventsList.nextSyncToken);
  return events;
}

// get equipment and state from event summary
function getEquipmentStateFromEvent(event){
  const summary = event.summary;
  const status = summary.split(' '); // split to equipment and state
  if (status.length === 1) { // just the equipment name (state is 'use')
    var equipment = status[0];
    var state = 'use';
  } else if (status.length === 2 || status.length === 3) { // (User Name) + equipment + state
    var equipment = status[status.length-2];
    var state = status[status.length-1];
  }
  return {equipment, state};
}

// rename and write event in write calendar and add read calendars as guests
function writeEvent(event, writeCalendarId, writeUser, readCalendarIds) {
  Utilities.sleep(100);
  const equipmentStateFromEvent = getEquipmentStateFromEvent(event);
  const equipment = equipmentStateFromEvent.equipment;
  const state = equipmentStateFromEvent.state;
  const eid = event.iCalUID;
  var event = CalendarApp.getCalendarById(writeCalendarId).getEventById(eid);
  // if equipment is enabled in sheets, add to guest subscription
  // change title from '(User Name) + equipment + state' to 'User Name + equipment + state'
  const summary = `${writeUser} ${equipment} ${state}`;
  event.setTitle(summary);
  // add read calendars as guests
  for (var i = 0; i < readCalendarIds.length; i++) {
    const readCalendarId = readCalendarIds[i];
    event.addGuest(readCalendarId);
  }
}

function getRelativeDate(daysOffset, hour) {
  const date = new Date();
  date.setDate(date.getDate() + daysOffset);
  date.setHours(hour);
  date.setMinutes(0);
  date.setSeconds(0);
  date.setMilliseconds(0);
  return date;
}

// get last and first names from full name
function setFirstLastNames(sheet, cell, newValue){
  const names = newValue.split(' ', 2); // split to last and first name
  const lastName = names[1];
  const firstName = names[0];
  sheet.getRange(cell.getRow(), 2).setValue(lastName);
  sheet.getRange(cell.getRow(), 3).setValue(firstName);
  // set User Name 1 using last and first name
  // User Name 1 = {Last Name up to 4 letters}.{First Name up to 1 letter}
  sheet.getRange(cell.getRow(), 4).setValue(lastName.slice(0,4)+'.'+firstName.slice(0,1));
}

// set User Name 2 using User Name 1
// User Name 2 = {User Name 1}{unique identifier 1~9}
//             = {Last Name up to 4 letters}.{First Name up to 1 letter}{unique identifier 1,2,3,...}
function setUserNames(sheet){
  const lastRow = sheet.getLastRow();
  // update User Name 2 for row0 = 2~lastRow
  for (var row0 = 2; row0 < lastRow+1; row0++) {
    const value0 = sheet.getRange(row0, 4).getValue();
    var count = 0;
    // check the duplicate count of value0 for row1 = 2~row0
    for (var row1 = 2; row1 < row0+1; row1++) {
      const value1 = sheet.getRange(row1, 4).getValue();
      if (value0 === value1){
        count += 1;
      }
    }
    // use count as unique identifier (1,2,3,...)
    sheet.getRange(row0, 5).setValue(value0+count); 
  }
}

// create checkboxes for selecting which equipment to show in the calendar
function setCheckboxes(sheet, cell) {
  const lastColumn = sheet.getLastColumn();
  if (sheet.getRange(cell.getRow(), 10).isChecked() == null){ // if cell is not a checkbox
    // create checkboxes
    for (var column = 10; column < lastColumn+1; column++) { 
      sheet.getRange(cell.getRow(), column).insertCheckboxes();  
    }
  }
  Logger.log('Created checkboxes');
}

// set read calendar and write calendar for created user
function setCalendars(sheet, cell) {
  const userName = sheet.getRange(cell.getRow(), 5).getValue();
  const readCalendarId = sheet.getRange(cell.getRow(), 6).getValue();
  const writeCalendarId = sheet.getRange(cell.getRow(), 7).getValue();
  changeCalendarName(readCalendarId, userName, 'Read');
  changeCalendarName(writeCalendarId, userName, 'Write');
}

// update calendar name and description
function changeCalendarName(calendarId, userName, readOrWrite) {
  // calendar name (summary) and description changes for read and write calendar
  if (readOrWrite === 'Read') { 
    var summary = `Read ${userName}`;
    var description = '装置の予約状況\n' +
      'schedule for selected equipments';  
  } else if (readOrWrite === 'Write') { 
    var summary = `Write ${userName}`;
    var description = '装置を予約する\n' +
      'Reserve equipments\n' +
      'Formatting: [Equipment] [State]\n' +
      'Equipments: rie, nrie(new RIE), cvd, ncvd(new CVD), pvd, fts\n' +
      'States: evac(evacuation), use(or no entry), cool(cooldown), o2(RIE O2 ashing)\n';  
  } else {
    Logger.log('readOrWrite has to be \'Read\' or \'Write\'');
  };
  const calendar = CalendarApp.getCalendarById(calendarId);
  calendar.setName(summary);
  calendar.setDescription(description);
  Logger.log('Updated calendar name');
}

// update the sync token after adding and deleting guests
function updateSyncToken(calendarId) {
  const properties = PropertiesService.getUserProperties();
  const options = {
    maxResults: 1000 // suppress nextPageToken which supresses nextSyncToken by fitting all events in one page
  };
  var eventsList;
  eventsList = Calendar.Events.list(calendarId, options);
  properties.setProperty(`syncToken ${calendarId}`, eventsList.nextSyncToken);
  Logger.log('Updated sync token. New sync token: ' + eventsList.nextSyncToken);
}
