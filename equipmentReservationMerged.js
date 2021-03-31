// delete all triggers for this script
function deleteTriggers(){
  const triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++){
    trigger = triggers[i];
    ScriptApp.deleteTrigger(trigger);
  }
}

// create triggers
// only 20 triggers can be made for single script
// we will use 19 for write calendars, 1 for spreadsheet
function createTriggers() {
  const sheet = SpreadsheetApp.openById('{spreadsheetid}').getSheetByName('users');
  const writeCalendarIds = getWriteCalendarIds(sheet);
  
  // create trigger for each of the 19 write calendars
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
      .forSpreadsheet('{spreadsheetid}')
      .onEdit()
      .create();
}

// when calendar gets edited
function onCalendarEdit(e) {
  const sheet = SpreadsheetApp.openById('{spreadsheetid}').getSheetByName('users');
  const calendarId = e.calendarId;
  const writeCalendarIds = getWriteCalendarIds(sheet);
  const index = writeCalendarIds.indexOf(calendarId);
  const fullSync = false;
  writeEventsToReadCalendar(sheet, calendarId, index, fullSync);
}

// when sheets gets edited
function onSheetsEdit(e) {
  const sheet = e.source.getActiveSheet();
  const cell = e.source.getActiveRange();
  const newValue = e.value;
  const row = cell.getRow();
  const column = cell.getColumn();
  const users = getUsers(sheet);
  const index = row-2;
  const readUser = users[index];
  const writeCalendarIds = getWriteCalendarIds(sheet);
  const calendarId = writeCalendarIds[index]
  const fullSync = true;
  
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

// write events to read calendar based on updated events in write calendar
function writeEventsToReadCalendar(sheet, calendarId, index, fullSync) {
  const readCalendarIds = getReadCalendars(sheet).readCalendarIds;
  const enabledDevicesList = getReadCalendars(sheet).enabledDevicesList;
  const writeCalendarIds = getWriteCalendarIds(sheet);
  const users = getUsers(sheet);
  const writeUser = users[index];
  const events = getEvents(calendarId, fullSync);
  updateSyncToken(calendarId);
  Logger.log(readCalendarIds.length + ' read calendars');
  for (var i = 0; i < readCalendarIds.length; i++){
    const readUser = users[i];
    if (readUser != writeUser) { // avoid duplicating event for same user's read and write calendars
      const readCalendarId = readCalendarIds[i];
      const enabledDevices = enabledDevicesList[i];
      writeEvents(events, calendarId, readCalendarId, enabledDevices, writeUser);
      Logger.log(events.length + ' events for CalendarId ' + readCalendarId);
    }
  }
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
    const writeCalendarId = writeCalendarIds[i];
    const events = getEvents(writeCalendarId, fullSync);
    const writeUser = users[i];
    if (readUser != writeUser) { // avoid duplicating event for same user's read and write calendars
      writeEvents(events, writeCalendarId, readCalendarId, enabledDevices, writeUser);
    }
    Logger.log(events.length + ' events for CalendarId ' + writeCalendarId);
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
    for (var column = 8; column < lastColumn+1; column++) { 
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

// update the sync token because updating once does not work
function updateSyncToken(calendarId) {
  const properties = PropertiesService.getUserProperties();
  const options = {
    maxResults: 100
  };
  const syncToken = properties.getProperty('syncToken'+calendarId);
  options.syncToken = syncToken;
  // Retrieve events one page at a time.
  var eventsList;
  eventsList = Calendar.Events.list(calendarId, options);
  properties.setProperty('syncToken'+calendarId, eventsList.nextSyncToken);
  Logger.log('Updated sync token. New sync token: ' + eventsList.nextSyncToken);
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
  var pageToken;
  var events = [];
  do {
    try {
      options.pageToken = pageToken;
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
  Logger.log('New sync token: ' + eventsList.nextSyncToken);
  return events;
}

// write events from write calendar to read calendar
function writeEvents(events, writeCalendarId, readCalendarId, enabledDevices, user) {
  if (events.length > 0) {
    for (var i = 0; i < events.length; i++) {
      Utilities.sleep(1000);
      var event = events[i];
      const summary = event.summary;
      const status = summary.split(' '); // split to device and state
      if (status.length === 1) { // just the device name (state is 'use')
        var device = status[0];
        var state = 'use';
      } else if (status.length === 2 || status.length === 3) { // (User Name) + device + state
        var device = status[status.length-2];
        var state = status[status.length-1];
      }
      const eid = event.iCalUID;
      var event = CalendarApp.getCalendarById(writeCalendarId).getEventById(eid);
      if (enabledDevices.includes(device) === true){ 
        // if device is enabled in sheets, add to guest subscription
        // change title from '(User Name) + device + state' to 'User Name + device + state'
        const summary = user  + ' ' + device + ' ' + state;
        event.setTitle(summary);
        event.addGuest(readCalendarId);
      }
      else { // if device is enabled in sheets, remove guest subscription
        event.removeGuest(readCalendarId);
      }
    }
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
  if (sheet.getRange(cell.getRow(), 8).isChecked() == null){ // if cell is not a checkbox
    // create checkboxes
    for (var column = 8; column < lastColumn+1; column++) { 
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