// setup
function setup() {
  const groupUrl = '?????@googlegroups.com'; // replace this line
  if (groupUrl.includes('?')) { // detect default value
    throw new Error('ERROR: change "?????@googlegroups.com" to your google group name');
  }
  createSpreadsheet(1); // create spreadsheet for 18 users
  createCalendars(1, groupUrl); // create 19 read + 18 write calendars
  createTriggers();
}

// creates spreadsheet for {count} users
function createSpreadsheet(count) {
  // create spreadsheet for configuration
  var configSpreadsheet = SpreadsheetApp.create('configSpreadsheet');
  configSpreadsheet.insertSheet('users');
  configSpreadsheet.insertSheet('properties');
  configSpreadsheet.deleteSheet(configSpreadsheet.getSheetByName('Sheet1'));
  var activeSheet = configSpreadsheet.getSheetByName('users');
  activeSheet.getRange(1, 1, 200, 200).setWrap(true); // wrap overflowing text
  activeSheet.getRange(1, 1, 1, 9).setValues(
    [['Full Name (EDIT this line)', 'Last Name', 'First Name', 'User Name 1', 'User Name 2', 'Read calendarId', 'Write calendarId', 'Read Cal URL', 'Write Cal URL']]
  );
  activeSheet.hideColumns(2, 6); // hide columns used for debug
  activeSheet.getRange(2, 10, count, 100).insertCheckboxes('no'); // create unchecked checkbox for 100 columns (devices)
  var fillValue = [];
  for (var i = 0; i < count; i++) {
    fillValue[i] = ['First Last'];
  }
  activeSheet.getRange(2, 1, count).setValues(fillValue);
  activeSheet.getRange(2+count+1, 10, 1, 100).insertCheckboxes('yes'); // create checked checkbox for "ALL EVENTS"
  activeSheet.getRange(2+count+1, 1).setValue('ALL EVENTS');
  var fillValue = [[]];
  for (var i = 0; i < 100; i++) {
    fillValue[0][i] = `=properties!R${1+i}C${1}`; // refer to sheet "properties" for device name
  }
  activeSheet.getRange(10, 1, 1, 100).setFormulas(fillValue); // copy device name

  var activeSheet = configSpreadsheet.getSheetByName('properties');
  activeSheet.getRange(1, 1, 200, 200).setWrap(true); // wrap overflowing text
  activeSheet.getRange(1, 1, 1, 2).setValues(
    [['Equipment', 'Properties (ex. temp, pressure, time) ->']]
  );

  // create spreadsheet for logging
  var loggingSpreadsheet = SpreadsheetApp.create('loggingSpreadsheet');
  // todo: add header

  var property = {
    configSpreadsheetId : configSpreadsheet.getId(),
    loggingSpreadsheetId : loggingSpreadsheet.getId(),
  };
  setIds(property);
}

// creates calendars for {count} users
function createCalendars(count, groupUrl) {
  const properties = PropertiesService.getUserProperties();
  const configSpreadsheet = SpreadsheetApp.openById(properties.getProperty('configSpreadsheetId'));
  const resource = { // used to add google group as guest
    'scope': {
      'type': 'group',
      'value': groupUrl,
    },
    'role': 'writer',
  }
  // create {count+1} read calendars
  for (var i = 0; i < count+1; i++){
    Utilities.sleep(3000);
    var calendar = CalendarApp.createCalendar(`read ${i+1}`);
    var readCalendarId = calendar.getId();
    Calendar.Acl.insert(resource, readCalendarId); // add access permission to google group
    var activeSheet = configSpreadsheet.getSheetByName('users');
    activeSheet.getRange(2+i, 6).setValue(readCalendarId);
    activeSheet.getRange(2+i, 8).setValue(`https://calendar.google.com/calendar/u/0?cid=${readCalendarId}`);
    Logger.log(`Created read calendar ${calendar.getName()}, with the ID ${readCalendarId}.`);
  }
  // create {count} write calendars
  for (var i = 0; i < count; i++){
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
  properties.setProperty('loggingSpreadsheetId', property.loggingSpreadsheetId);
}

function setIdsManual(y) {
  //const configSpreadsheetId = ;
  //const loggingSpreadsheetId = ;
  properties.setProperty('configSpreadsheetId', property.configSpreadsheetId);
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
  ScriptApp.newTrigger("eventLogging")
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
  const users = getUsers(sheet);
  const calendarId = e.calendarId;
  const writeCalendarIds = getWriteCalendarIds(sheet);
  const index = writeCalendarIds.indexOf(calendarId);
  const fullSync = false;
  adminLoggingSetup(); // setup admin logging
  adminLogging({ 
    name: users[index],
    action: "add/move/del event", 
  });
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
  
  adminLoggingSetup(); // setup admin logging
  if (row == 1 || row > lastRow || column > lastColumn || (column > 1 && column < 6)){
    action = "edited invalid area";
  } else if (column === 1) {
    action = "edited name";
  } else if (column === 6) {
    action = "edited read calendarId";
  } else if (column === 7) {
    action = "edited write calendarId";
  } else if (column > 7) {
    action = "edited device";
  } 
  const executionDateTime = new Date(); // current time
    adminLogging({
      executionYear: executionDateTime.getFullYear(),
      executionMonth: executionDateTime.getMonth()+1, // Jan. is 0, Dec is 11 --change--> 1~12
      executionDay: executionDateTime.getDate(),
      executionHour: executionDateTime.getHours(),
      executionMin: executionDateTime.getMinutes(),
      executionSec: executionDateTime.getSeconds(),
      executionMilliSec: executionDateTime.getMilliseconds(),
    name: readUser,
    action: 'sheets (row, col) = (' + row + ', ' + column + '): ' + action,
  });  
  // when the checkbox (H2~nm) is edited in sheets on sheet 'users'
  // update corresponding user's subscribed devices
  if (sheet.getName() === 'users' && row > 1 && column > 7){ 
    changeSubscribedDevices(sheet, readUser, index, users);
  }
  // when the full name (A2~An) is edited in sheets on sheet 'users'
  // update all of the corresponding user's event title
  else if (sheet.getName() === 'users' && row > 1 && column === 1){
    updateCalendarUserName(sheet, cell, newValue);
    writeEventsToReadCalendar(sheet, calendarId, index, fullSync);
  }
}

function adminLoggingSetup() { // prepares constant for adminLogging
  const properties = PropertiesService.getUserProperties();
  book = SpreadsheetApp.openById(properties.getProperty('loggingSpreadsheetId')); // spreadsheet for loggingrepares constant for adminLogging
  const lastRow = book.getSheetByName('admin_log').getLastRow();
  const row = lastRow + 1; // write on new row
  properties.setProperty('row', row.toString());
  properties.setProperty('adminLoggingDone', 'false'); // reset execution state of admin logging
}

function adminLogging(logObj) { // logs everything
  const properties = PropertiesService.getUserProperties();
  const columnDescriptions = { // shows which description corresponds to which column
    executionYear: 1,
    executionMonth: 2,
    executionDay: 3,
    executionHour: 4,
    executionMin: 5,
    executionSec: 6,
    executionMilliSec: 7,
    startYear: 8,
    startMonth: 9,
    startDay: 10,
    startHour: 11,
    startMin: 12,
    startSec: 13,
    startMilliSec: 14,
    endYear: 15,
    endMonth: 16,
    endDay: 17,
    endHour: 18,
    endMin: 19,
    endSec: 20,
    endMilliSec: 21,
    durationYear: 22,
    durationMonth: 23,
    durationDay: 24,
    durationHour: 25,
    durationMin: 26,
    durationSec: 27,
    durationMilliSec: 28,
    specialAllDay: 29,
    specialReccuring: 30,
    specialRecurringDays: 31,
    name: 32,
    device: 33,
    status: 34,
    comment: 35,
    action: 36,
  }; // didn't put this in adminLoggingSetup because it cannot hold js objects
  const row = parseInt(properties.getProperty('row'));
  const adminLogSheet = SpreadsheetApp.openById(properties.getProperty('loggingSpreadsheetId')).getSheetByName('admin_log'); // spreadsheet for logging
  for (const key in logObj) { // iterate through log object
    var value = logObj[key];
    var col = columnDescriptions[key];
    adminLogSheet.getRange(row,col).setValue(value);
  }
}

function eventLogging() { // logs just the necessary data
  const properties = PropertiesService.getUserProperties();
  const eventLogSheet = SpreadsheetApp.openById(properties.getProperty('loggingSpreadsheetId')).getSheetByName('event_log')
  const lastRow = eventLogSheet.getLastRow();
  const row = lastRow + 1; // write on new row
  const columnDescriptions = { // shows which description corresponds to which column
    startYear: 1,
    startMonth: 2,
    startDay: 3,
    startHour: 4,
    startMin: 5,
    endYear: 6,
    endMonth: 7,
    endDay: 8,
    endHour: 9,
    endMin: 10,
    durationYear: 11,
    durationMonth: 12,
    durationDay: 13,
    durationHour: 14,
    durationMin: 15,
    specialAllDay: 16,
    specialReccuring: 17,
    specialRecurringDays: 18,
    name: 19,
    device: 20,
    status: 21,
    comment: 22,
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
          Logger.log('Event id %s was cancelled.', event.id);
        } else{
          events.push(event);
        }
      }
    } 

    if (events.length === 0) {
      Logger.log('No events found for: ' + writeCalendarId);
    } else {
      Logger.log(events.length + ' events found for: ' + writeCalendarId);
    }
    
    // log each event
    for (var j = 0; j < events.length; j++){
      Utilities.sleep(100);
      var event = events[j];
      const deviceStateFromEvent = getDeviceStateFromEvent(event); // this must be called first before getCalendarById
      const device = deviceStateFromEvent.device;
      const state = deviceStateFromEvent.state;
      const eid = event.iCalUID;
      event = CalendarApp.getCalendarById(writeCalendarId).getEventById(eid);
      const startDateTime = event.getStartTime();
      const endDateTime = event.getEndTime();
      const durationDateTime = new Date(endDateTime - startDateTime);
      var logObj = {
          startYear: startDateTime.getFullYear(),
          startMonth: startDateTime.getMonth()+1,
          startDay: startDateTime.getDate(),
          startHour: startDateTime.getHours(),
          startMin: startDateTime.getMinutes(),
          endYear: endDateTime.getFullYear(),
          endMonth: endDateTime.getMonth()+1,
          endDay: endDateTime.getDate(),
          endHour: endDateTime.getHours(),
          endMin: endDateTime.getMinutes(),
          durationYear: durationDateTime.getFullYear() - 1970, // starts at 1970
          durationMonth: durationDateTime.getMonth(), // dont add 1 to difference
          durationDay: durationDateTime.getDate()-1, // dont know why it shows 1 when it should be 0
          durationHour: durationDateTime.getHours(),
          durationMin: durationDateTime.getMinutes(),
          specialAllDay: event.isAllDayEvent(),
          specialReccuring: event.isRecurringEvent(),
          specialRecurringDays: "",
          name: writeUser,
          device: device,
          status: state,
          comment: "",
      }
      for (const key in logObj) { // iterate through log object
        var value = logObj[key];
        var col = columnDescriptions[key];
        eventLogSheet.getRange(row,col).setValue(value);
      }
    }
  }
}

// filter readUsers who are not writeUser and have the device
function filterUsers(writeUser, event, readCalendarIds, users, enabledDevicesList) {
  var filteredReadCalendarIds = []; // readCalendarIds excluding the same user as writeCalendarId
  var filteredReadUsers = []; // readUsers excluding the same user as writeCalendarId
  const device = getDeviceStateFromEvent(event).device; // device used in the event
  for (var i = 0; i < readCalendarIds.length; i++){
    const readCalendarId = readCalendarIds[i];
    const readUser = users[i];
    const enabledDevices = enabledDevicesList[i];
    if (enabledDevices.includes(device) === true){ // if device is enabled for readUser
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
  const enabledDevicesList = getReadCalendars(sheet).enabledDevicesList;
  const writeCalendarIds = getWriteCalendarIds(sheet);
  const users = getUsers(sheet);
  const writeUser = users[index];
  const events = getEvents(writeCalendarId, fullSync);
  Logger.log(readCalendarIds.length + ' read calendars');
  for (var i = 0; i < events.length; i++){
    const event = events[i];
    const filteredReadCalendarIds = filterUsers(writeUser, event, readCalendarIds, users, enabledDevicesList).filteredReadCalendarIds;
    Logger.log('writing event no.' + (i+1).toString() +  ' to [ ' + filteredReadCalendarIds + ' ]');
    writeEvent(event, writeCalendarId, writeUser, filteredReadCalendarIds); // create event in write calendar and add read calendars as guests
  }
  updateSyncToken(writeCalendarId); // renew sync token after adding guest
  Logger.log('Wrote updated events to read calendar. Fullsync = ' + fullSync);
}

// update corresponding user's subscribed devices 
function changeSubscribedDevices(sheet, readUser, index, users){
  const fullSync = true;
  const readCalendarIds = getReadCalendars(sheet).readCalendarIds;
  const enabledDevicesList = getReadCalendars(sheet).enabledDevicesList;
  const writeCalendarIds = getWriteCalendarIds(sheet);
  const readCalendarId = readCalendarIds[index];
  const enabledDevices = enabledDevicesList[index];
  Logger.log(writeCalendarIds.length + ' write calendars');
  for (var i = 0; i < writeCalendarIds.length; i++){
    const writeUser = users[i];
    const writeCalendarId = writeCalendarIds[i];
    const events = getEvents(writeCalendarId, fullSync);
    for (var j = 0; j < events.length; j++){
      const event = events[j];
      const filteredReadCalendarIds = filterUsers(writeUser, event, [readCalendarId], users, enabledDevicesList).filteredReadCalendarIds;
      writeEvent(event, writeCalendarId, writeUser, filteredReadCalendarIds); // create event in write calendar and add read calendars as guests
    }
    updateSyncToken(writeCalendarId);
  }
  Logger.log('Changed subscribed devices');
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
  const enabledDevicesList = [];
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  // get calendarId and enabledDevice
  for (var row = 2; row < lastRow+1; row++) { 
    // get calendarId and add to calendarIds
    const calendarId = sheet.getRange(row, 6).getValue();
    calendarIds.push(calendarId);
    // get enabledDevice and add to enabledDevices
    const enabledDevices = [];
    for (var column = 10; column < lastColumn+1; column++) { 
      const device = sheet.getRange(1, column).getValue();
      const checked = sheet.getRange(row, column).isChecked();
      if (checked === true){
        enabledDevices.push(device);
      }
    }
    enabledDevicesList.push(enabledDevices);
  }
  return {
    readCalendarIds: calendarIds,
    enabledDevicesList: enabledDevicesList,
  };
}

// get events from the given calendar that have been modified since the last sync.
// if the sync token is missing or invalid, log all events from up to a ten days ago (a full sync).
function getEvents(calendarId, fullSync) {
  const properties = PropertiesService.getUserProperties();
  const options = {
    maxResults: 100
  };
  const syncToken = properties.getProperty('syncToken'+calendarId);
  Logger.log('Current sync token: ' + syncToken);
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
        delete options.pageToken; // delete key "pageToken"
      } else {
        options.pageToken = pageToken;
      }
      eventsList = Calendar.Events.list(calendarId, options);
    } catch (e) {
      // Check to see if the sync token was invalidated by the server;
      // if so, perform a full sync instead.
      if (e.message === 'API call to calendar.events.list failed with error: Sync token is no longer valid, a full sync is required.') {
        Logger.log('Sync token invalidated -> Full sync initiated');
        properties.deleteProperty('syncToken'+calendarId);
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
  properties.setProperty('syncToken'+calendarId, eventsList.nextSyncToken);
  return events;
}

// get device and state from event summary
function getDeviceStateFromEvent(event){
  const summary = event.summary;
  const status = summary.split(' '); // split to device and state
  if (status.length === 1) { // just the device name (state is 'use')
    var device = status[0];
    var state = 'use';
  } else if (status.length === 2 || status.length === 3) { // (User Name) + device + state
    var device = status[status.length-2];
    var state = status[status.length-1];
  }
  return {device, state};
}

// rename and write event in write calendar and add read calendars as guests
function writeEvent(event, writeCalendarId, writeUser, readCalendarIds) {
  Utilities.sleep(100);
  const deviceStateFromEvent = getDeviceStateFromEvent(event);
  const device = deviceStateFromEvent.device;
  const state = deviceStateFromEvent.state;
  const eid = event.iCalUID;
  var event = CalendarApp.getCalendarById(writeCalendarId).getEventById(eid);
  // if device is enabled in sheets, add to guest subscription
  // change title from '(User Name) + device + state' to 'User Name + device + state'
  const summary = writeUser  + ' ' + device + ' ' + state;
  event.setTitle(summary);
  // add read calendars as guests
  for (var i = 0; i < readCalendarIds.length; i++) {
    const readCalendarId = readCalendarIds[i];
    event.addGuest(readCalendarId);
  }
  const executionDateTime = new Date(); // current time
  const startDateTime = event.getStartTime();
  const endDateTime = event.getEndTime();
  const durationDateTime = new Date(endDateTime - startDateTime);
  const properties = PropertiesService.getUserProperties();
  var adminLoggingDone = properties.getProperty('adminLoggingDone');
  if (adminLoggingDone === 'false'){ // do admin logging only once
    adminLogging({
      executionYear: executionDateTime.getFullYear(),
      executionMonth: executionDateTime.getMonth()+1, // Jan. is 0, Dec is 11 --change--> 1~12
      executionDay: executionDateTime.getDate(),
      executionHour: executionDateTime.getHours(),
      executionMin: executionDateTime.getMinutes(),
      executionSec: executionDateTime.getSeconds(),
      executionMilliSec: executionDateTime.getMilliseconds(),
      startYear: startDateTime.getFullYear(),
      startMonth: startDateTime.getMonth()+1,
      startDay: startDateTime.getDate(),
      startHour: startDateTime.getHours(),
      startMin: startDateTime.getMinutes(),
      startSec: startDateTime.getSeconds(),
      startMilliSec: startDateTime.getMilliseconds(),
      endYear: endDateTime.getFullYear(),
      endMonth: endDateTime.getMonth()+1,
      endDay: endDateTime.getDate(),
      endHour: endDateTime.getHours(),
      endMin: endDateTime.getMinutes(),
      endSec: endDateTime.getSeconds(),
      endMilliSec: endDateTime.getMilliseconds(),
      durationYear: durationDateTime.getFullYear() - 1970, // starts at 1970
      durationMonth: durationDateTime.getMonth(), // dont add 1 to difference
      durationDay: durationDateTime.getDate()-1, // dont know why it shows 1 when it should be 0
      durationHour: durationDateTime.getHours(),
      durationMin: durationDateTime.getMinutes(),
      durationSec: durationDateTime.getSeconds(),
      durationMilliSec: durationDateTime.getMilliseconds(),
      specialAllDay: event.isAllDayEvent(),
      specialReccuring: event.isRecurringEvent(),
      specialRecurringDays: "",
      device: device,
      status: state,
      comment: "",
    });
    Logger.log('admin logging done');
    properties.setProperty('adminLoggingDone', 'true');
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
    var summary = 'Read' + ' ' + userName;
    var description = '装置の予約状況\n' +
      'schedule for selected devices';  
  } else if (readOrWrite === 'Write') { 
    var summary = 'Write' + ' ' + userName;
    var description = '装置を予約する\n' +
      'Reserve devices\n' +
      'Formatting: [Device] [State]\n' +
      'Devices: rie, nrie(new RIE), cvd, ncvd(new CVD), pvd, fts\n' +
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
  properties.setProperty('syncToken'+calendarId, eventsList.nextSyncToken);
  Logger.log('Updated sync token. New sync token: ' + eventsList.nextSyncToken);
}
