// ===============================================================================================
// ======================================= SETUP FUNCTIONS ======================================= 
// ===============================================================================================

// define constants used over several scripts
function defineConstants() {
  Logger.log('Defining constants');
  const properties = PropertiesService.getUserProperties();
  properties.setProperty('groupUrl', '?????@googlegroups.com'); // groupURL
  properties.setProperty('timeZone', 'Asia/Tokyo'); // set timezone
  properties.setProperty('userCount', 18); // number of users
  properties.setProperty('equipmentCount', 50); // number of equipments
  properties.setProperty('experimentConditionCount', 20); // number of experiment conditions for a single equipment
  properties.setProperty('experimentConditionRows', 5000); // number of rows in experiment condition
  properties.setProperty('experimentConditionBackupRows', 4500); // number of rows to backup and delete in case of overflow of sheets
  properties.setProperty('finalLoggingRows', 1000000); // number of rows in final logging
  properties.setProperty('finalLoggingBackupRows', 990000); // number of rows in final logging
  properties.setProperty('backgroundColor', '#bbbbbb'); // background color of the uneditable cells (gray)
  if (properties.getProperty('groupUrl').includes('?')) { // detect default value and throw error
    throw new Error('ERROR: change "?????@googlegroups.com" to your google group name');
  }
}

// setup: split into 4 parts to avoid execution time limit
function setup() {
  Logger.log('Running setup. Dont touch any files and wait for **15** minutes`');
  defineConstants(); // define constants used over several scripts
  createSpreadsheets1(); // create spreadsheet for 18 users
  timedTrigger('setup2'); // create spreadsheet for 18 users
}
function setup2() {
  createSpreadsheets2();
  timedTrigger('setup3'); 
}
function setup3() {
  createSpreadsheets3();
  timedTrigger('setup4'); // create 19 read + 18 write calendars
}
function setup4() {
  createCalendars();
  deleteTriggers(); // delete timed triggers and previous triggers
  getAndStoreObjects();
  createTriggers();
}

// creates spreadsheet for {userCount} users
function createSpreadsheets1() {
  const properties = PropertiesService.getUserProperties();
  const userCount = parseInt(properties.getProperty('userCount'));
  const timeZone = properties.getProperty('timeZone');
  const groupUrl = properties.getProperty('groupUrl');
  const equipmentCount = parseInt(properties.getProperty('equipmentCount'));
  const experimentConditionCount = parseInt(properties.getProperty('experimentConditionCount'));
  const experimentConditionRows = parseInt(properties.getProperty('experimentConditionRows'));
  const finalLoggingRows = parseInt(properties.getProperty('finalLoggingRows'));

  // create workbooks(spreadsheets) and sheets
  var experimentConditionSpreadsheet = SpreadsheetApp.create('experimentConditionSpreadsheet');
  experimentConditionSpreadsheet.setSpreadsheetTimeZone(timeZone);
  experimentConditionSpreadsheet.insertSheet('users'); 
  experimentConditionSpreadsheet.insertSheet('properties');
  experimentConditionSpreadsheet.insertSheet('allEquipments');
  var loggingSpreadsheet = SpreadsheetApp.create('loggingSpreadsheet');
  loggingSpreadsheet.setSpreadsheetTimeZone(timeZone);
  loggingSpreadsheet.insertSheet('finalLog');
  loggingSpreadsheet.deleteSheet(loggingSpreadsheet.getSheetByName('Sheet1'));
  // get ids
  const experimentConditionSpreadsheetId = experimentConditionSpreadsheet.getId();
  const loggingSpreadsheetId = loggingSpreadsheet.getId();
  var property = { // store spreadsheetIds
    experimentConditionSpreadsheetId : experimentConditionSpreadsheetId,
    loggingSpreadsheetId : loggingSpreadsheetId,
  };
  setIds(property);
  // share to google groups
  DriveApp.getFileById(experimentConditionSpreadsheetId).addEditor(groupUrl);
  DriveApp.getFileById(loggingSpreadsheetId).addEditor(groupUrl);

  // create spreadsheet for experiment condition logging
  Logger.log('Creating experiment condition spreadsheet');
  var sheetIds = [];
  var filledArrayBatch = []; // list of filledArray
  var activeSheet = experimentConditionSpreadsheet.getSheetByName('eventLog');
  var firstSheet = '';
  for (var i = 0; i < equipmentCount; i++) { // create sheet for each equipment
    Logger.log(`Creating equipmentSheet ${i+1}/${equipmentCount}`);
    Utilities.sleep(100);
    if (i === 0) { // create first sheet
      var activeSheet = experimentConditionSpreadsheet.insertSheet(`equipment${i+1}`);
      firstSheet = experimentConditionSpreadsheet.getSheetByName(`equipment${i+1}`);
      sheetIds[i] = activeSheet.getSheetId();
      experimentConditionSpreadsheet.deleteSheet(experimentConditionSpreadsheet.getSheetByName('Sheet1'));
      changeSheetSize(activeSheet, experimentConditionRows, 12+experimentConditionCount);
      activeSheet.hideColumns(6, 7); // hide columns used for debug
      const filledArray = [['startTime', 'endTime', 'name', 'equipment', 'state', 'description', 'isAllDayEvent', 'isRecurringEvent', 'action', 'executionTime', 'id', 'eventExists']];
      setValues(filledArray, `equipment${i+1}!${R1C1RangeToA1Range(1, 1, 1, 12)}`, experimentConditionSpreadsheetId);
      activeSheet.getRange(1, 1, 1, 12+experimentConditionCount).setBorder(null, null, true, null, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
      activeSheet.getRange(1, 5, experimentConditionRows, 1).setBorder(null, null, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
      activeSheet.getRange(1, 12, experimentConditionRows, 1).setBorder(null, null, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
      // protect range
      protectRange(activeSheet.getRange(1, 1, 1, 12+experimentConditionCount));
      protectRange(activeSheet.getRange(2, 6, experimentConditionRows-1, 7));
      // set headers
      var filledArray = [[]];
      for (var j = 0; j < experimentConditionCount; j++) {
        filledArray[0][j] = `=INDIRECT(\"properties!R${2+i}C${4+j}\", FALSE)`;
      }
      setValues(filledArray, `equipment${i+1}!${R1C1RangeToA1Range(1, 13, 1, experimentConditionCount)}`, experimentConditionSpreadsheetId);
      var filledArray = arrayFill2d(experimentConditionRows, 12, '');
      for (var j = 0; j < experimentConditionRows; j++) {
        filledArray[j][11] = `=INDIRECT(\"allEquipments!R\" & 1+MATCH(\"equipment${i+1}!R\" & ROW(), INDIRECT(\"allEquipments!E2:E\"), 0) & \"C6\", FALSE)`; // ADDRESS(row, col)
      }
      setValues(filledArray, `equipment${i+1}!${R1C1RangeToA1Range(2, 1, experimentConditionRows, 12)}`, experimentConditionSpreadsheetId);

      Logger.log('Creating filters for hiding canceled and modified events');

      // addFilterView -> adds filter view -> updates automatically -> apply manually after select
      // setBasicFilter -> adds filter (same as spreadsheetApp filter) -> doesn't update automatically
      addFilterViewRequest = {
        'addFilterView': {
          'filter': {
            "filterViewId": i+1, // 1~equipmentCount
            'title': 'hide canceled or modified events',
            'sortSpecs': [ // sort doesn't include header row
              {'dimensionIndex': 0, 'sortOrder': 'ASCENDING'}, // sort by startTime
              {'dimensionIndex': 9, 'sortOrder': 'ASCENDING'}, // sort by executionTime if startTime is same
            ], 
            "range": {
              "sheetId": sheetIds[i],
              "startRowIndex": 0,
              "endRowIndex": experimentConditionRows,
              "startColumnIndex": 0,
              "endColumnIndex": 12+experimentConditionCount,
            },
            'criteria': { // when column 12 is FALSE, hide row
              11: { 'hiddenValues': ['FALSE'] },
            },
          }
        }
      }
      Sheets.Spreadsheets.batchUpdate(
        {'requests': [addFilterViewRequest]}, experimentConditionSpreadsheetId
      );
    } else { // copy most contents from first sheet
      var activeSheet = firstSheet.copyTo(experimentConditionSpreadsheet).setName(`equipment${i+1}`);
      sheetIds[i] = activeSheet.getSheetId();
      // set headers
      var filledArray = [[]];
      for (var j = 0; j < experimentConditionCount; j++) {
        filledArray[0][j] = `=INDIRECT("properties!R${2+i}C${4+j}", FALSE)`;
      }
      filledArrayBatch.push({"majorDimension": "ROWS", "values": filledArray, "range": `equipment${i+1}!${R1C1RangeToA1Range(1, 13, 1, experimentConditionCount)}`});
      var filledArray = arrayFill2d(experimentConditionRows, 1, '');
      for (var j = 0; j < experimentConditionRows; j++) {
        filledArray[j][0] = `=INDIRECT("allEquipments!R" & 1+MATCH("equipment${i+1}!R" & ROW(), INDIRECT("allEquipments!E2:E"), 0) & "C6", FALSE)`; // ADDRESS(row, col)
      }
      filledArrayBatch.push({"majorDimension": "ROWS", "values": filledArray, "range": `equipment${i+1}!${R1C1RangeToA1Range(2, 12, experimentConditionRows, 1)}`});
    }
  }
  setValuesBatch(filledArrayBatch, experimentConditionSpreadsheetId);
  filledArrayBatch = [];
  properties.setProperty('sheetIds', JSON.stringify(sheetIds));
}

// creates spreadsheet for {userCount} users
function createSpreadsheets2() {
  const properties = PropertiesService.getUserProperties();
  const userCount = parseInt(properties.getProperty('userCount'));
  const timeZone = properties.getProperty('timeZone');
  const groupUrl = properties.getProperty('groupUrl');
  const equipmentCount = parseInt(properties.getProperty('equipmentCount'));
  const experimentConditionCount = parseInt(properties.getProperty('experimentConditionCount'));
  const experimentConditionRows = parseInt(properties.getProperty('experimentConditionRows'));
  const experimentConditionSpreadsheetId = properties.getProperty('experimentConditionSpreadsheetId');
  const loggingSpreadsheetId = properties.getProperty('loggingSpreadsheetId');
  const finalLoggingRows = parseInt(properties.getProperty('finalLoggingRows'));
  const experimentConditionSpreadsheet = SpreadsheetApp.openById(experimentConditionSpreadsheetId);
  const loggingSpreadsheet = SpreadsheetApp.openById(loggingSpreadsheetId);
  const sheetIds = JSON.parse(properties.getProperty('sheetIds'));
  
  // allEquipment sheet
  Logger.log('Creating allEquipment sheet');
  var activeSheet = experimentConditionSpreadsheet.getSheetByName('allEquipments'); 
  changeSheetSize(activeSheet, experimentConditionRows*equipmentCount+1, 6);
  // draw borders
  Logger.log('Drawing borders');
  activeSheet.getRange(1, 1, 1, 6).setBorder(null, null, true, null, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
  // protect range
  Logger.log('Protecting range');
  protectRange(activeSheet.getRange(1, 1, experimentConditionRows*equipmentCount+1, 6));
  // set headers
  var filledArray = [['startTime','executionTime','id','action','originalAddress','eventExists']];
  setValues(filledArray, `allEquipments!${R1C1RangeToA1Range(1, 1, 1, 6)}`, experimentConditionSpreadsheetId);
  var filledArray1 = [];
  var filledArray2 = [];
  var row = 0;
  for (var i = 0; i < equipmentCount; i++) {
    for (var j = 0; j < experimentConditionRows; j++) {
      row = i*experimentConditionRows + j;
      filledArray1[row] = [ // column A~C
        `=INDIRECT(\"equipment${i+1}!R${2+row}C1\", FALSE)`, // C1
        `=INDIRECT(\"equipment${i+1}!R${2+row}C10\", FALSE)`, // C10
        `=INDIRECT(\"equipment${i+1}!R${2+row}C11\", FALSE)`, // C11
      ]; // refer to sheet 'properties' for equipment name
      filledArray2[row] = [ // column D~E
        `=INDIRECT(\"equipment${i+1}!R${2+row}C9\", FALSE)`, // C9
        `'equipment${i+1}!R${2+j}`,
      ]; // refer to sheet 'properties' for equipment name
    }
  }
  // column A~C
  Logger.log('Settings formulas for columns A~C');
  setValues(filledArray1, `allEquipments!A2:C${experimentConditionRows*equipmentCount+1}`, experimentConditionSpreadsheetId);
  // column D~E
  Logger.log('Settings formulas for columns D~E');
  setValues(filledArray2, `allEquipments!D2:E${experimentConditionRows*equipmentCount+1}`, experimentConditionSpreadsheetId);
  
  // column F (could not be set with sheetsAPI for some reason...)
  //see if event exists (if it is 1[unmodified(is the last entry with the same id)] and 2[not canceled]) or 3[cell is empty]
  Logger.log('Settings formulas for column F');
  experimentConditionSpreadsheet
    .getSheetByName('allEquipments')
    .getRange(2, 6, experimentConditionRows*equipmentCount, 1)
    .setFormula(`=OR(AND(COUNTIF(INDIRECT("R[1]C[-3]:R${experimentConditionRows*equipmentCount+1}C[-3]", FALSE), INDIRECT("R[0]C[-3]", FALSE))=0, INDIRECT("R[0]C[-2]", FALSE)="add"), INDIRECT("R[0]C[-4]", FALSE)="")`);

  addFilterViewRequest = {
    'addFilterView': {
      'filter': {
        "filterViewId": 0, 
        'title': 'sort events by date and time',
        'sortSpecs': [ // sort doesn't include header row
          {'dimensionIndex': 0, 'sortOrder': 'ASCENDING'}, // sort by startTime
          {'dimensionIndex': 1, 'sortOrder': 'ASCENDING'}, // sort by executionTime if startTime is same
        ], 
        "range": {
          "sheetId": activeSheet.getSheetId(),
          "startRowIndex": 0,
          "endRowIndex": experimentConditionRows*equipmentCount+1,
          "startColumnIndex": 0,
          "endColumnIndex": 6,
        },
      }
    }
  }
  Logger.log('Adding sort filter view');
  Sheets.Spreadsheets.batchUpdate(
    {'requests': [addFilterViewRequest]}, experimentConditionSpreadsheetId
  );
}

// creates spreadsheet for {userCount} users
function createSpreadsheets3() {
  const properties = PropertiesService.getUserProperties();
  const userCount = parseInt(properties.getProperty('userCount'));
  const timeZone = properties.getProperty('timeZone');
  const groupUrl = properties.getProperty('groupUrl');
  const equipmentCount = parseInt(properties.getProperty('equipmentCount'));
  const experimentConditionCount = parseInt(properties.getProperty('experimentConditionCount'));
  const experimentConditionRows = parseInt(properties.getProperty('experimentConditionRows'));
  const experimentConditionSpreadsheetId = properties.getProperty('experimentConditionSpreadsheetId');
  const loggingSpreadsheetId = properties.getProperty('loggingSpreadsheetId');
  const finalLoggingRows = parseInt(properties.getProperty('finalLoggingRows'));
  const experimentConditionSpreadsheet = SpreadsheetApp.openById(experimentConditionSpreadsheetId);
  const loggingSpreadsheet = SpreadsheetApp.openById(loggingSpreadsheetId);
  const sheetIds = JSON.parse(properties.getProperty('sheetIds'));
  
  // users sheet
  Logger.log('Creating users sheet');
  var activeSheet = experimentConditionSpreadsheet.getSheetByName('users'); 
  changeSheetSize(activeSheet, userCount+2, equipmentCount+9);
  activeSheet.hideColumns(2, 6); // hide columns used for debug
  activeSheet.getRange(2, 6, userCount+1, 4).setHorizontalAlignment("left"); // show "https://..." not the center of url
  activeSheet.getRange(2, 6, userCount+1, 4).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); // link is too long -> clip
  // draw borders
  activeSheet.getRange(1, 1, userCount+2, equipmentCount+9).setBorder(true, true, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
  activeSheet.getRange(1, 1, userCount+2, equipmentCount+9).setBorder(true, true, true, true, null, true, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  activeSheet.getRange(1, 1, 1, equipmentCount+9).setBorder(null, null, true, null, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
  activeSheet.getRange(1, 1, userCount+2, 1).setBorder(null, null, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
  activeSheet.getRange(1, 5, userCount+2, 1).setBorder(null, null, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  activeSheet.getRange(1, 7, userCount+2, 1).setBorder(null, null, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  activeSheet.getRange(1, 9, userCount+2, 1).setBorder(null, null, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  // protect range
  protectRange(activeSheet.getRange(1, 1, 1, equipmentCount+9));
  protectRange(activeSheet.getRange(2, 2, userCount+1, 8));
  // set headers
  var filledArray = [['Full Name (EDIT this line)', 'Last Name', 'First Name', 'User Name 1', 'User Name 2', 'Read CalendarId', 'Write CalendarId', 'Read Calendar URL', 'Write Calendar URL']];
  setValues(filledArray, `users!${R1C1RangeToA1Range(1, 1, 1, 9)}`, experimentConditionSpreadsheetId);
  // normal user row
  activeSheet.getRange(2, 10, userCount, equipmentCount).insertCheckboxes(); // create unchecked checkbox for 100 columns (equipments)
  var filledArray = arrayFill2d(userCount, 1, 'First Last');
  setValues(filledArray, `users!${R1C1RangeToA1Range(2, 1, userCount, 1)}`, experimentConditionSpreadsheetId);
  // 'ALL EVENTS' user row
  activeSheet.getRange(2+userCount, 10, 1, equipmentCount).insertCheckboxes(); // create checked checkbox for 'ALL EVENTS'
  activeSheet.getRange(2+userCount, 1).setValue('ALL EVENTS');
  // copy equipments name from properties sheet
  var filledArray = [[]];
  for (var i = 0; i < equipmentCount; i++) {
    filledArray[0][i] = `=INDIRECT("properties!R${2+i}C1", FALSE)`; // refer to sheet 'properties' for equipment name
  }
  activeSheet.getRange(1, 10, 1, equipmentCount).setFormulas(filledArray);

  // properties sheet
  Logger.log('Creating properties sheet');
  Utilities.sleep(1000);
  var activeSheet = experimentConditionSpreadsheet.getSheetByName('properties');
  changeSheetSize(activeSheet, equipmentCount+1, experimentConditionCount+1);
  // draw borders
  activeSheet.getRange(1, 1, equipmentCount+1, experimentConditionCount+1).setBorder(true, true, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
  activeSheet.getRange(1, 1, 1, experimentConditionCount+1).setBorder(null, null, true, null, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
  activeSheet.getRange(1, 1, equipmentCount+1, 1).setBorder(null, null, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
  activeSheet.getRange(1, 3, equipmentCount+1, 1).setBorder(null, null, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
  // protect range
  protectRange(activeSheet.getRange(1, 1, 1, experimentConditionCount+1));
  protectRange(activeSheet.getRange(2, 2, equipmentCount, 2));
  // set headers
  var filledArray = [[]];
  filledArray[0] = ['equipmentName', 'sheetId', 'sheetUrl', 'Properties ->'];
  for (var i = 0; i < equipmentCount; i++) {
    filledArray[i+1] = ['', sheetIds[i], `https://docs.google.com/spreadsheets/d/${experimentConditionSpreadsheetId}/edit#gid=${sheetIds[i]}`, ''];
  }
  setValues(filledArray, `properties!${R1C1RangeToA1Range(1, 1, equipmentCount+1, 4)}`, experimentConditionSpreadsheetId);
  activeSheet.getRange(1, 2, equipmentCount+1, 2).setHorizontalAlignment("left"); // show https://... not the center of url
  activeSheet.getRange(1, 2, equipmentCount+1, 2).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); // link is too long -> clip
  activeSheet.hideColumns(2); // hide columns used for debug
  
  // create spreadsheet for finalized logging
  Logger.log('Creating logging spreadsheet');
  Utilities.sleep(1000);
  Logger.log('Creating final log sheet');
  var activeSheet = loggingSpreadsheet.getSheetByName('finalLog');
  changeSheetSize(activeSheet, finalLoggingRows, 8);
  // draw borders
  activeSheet.getRange(1, 1, 1, 8).setBorder(null, null, true, null, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
  // protect range
  protectRange(activeSheet.getRange(1, 1, finalLoggingRows, 8));
  // set headers
  const filledArray = [['startTime', 'endTime', 'name', 'equipment', 'state', 'description', 'isAllDayEvent', 'isRecurringEvent']];
  setValues(filledArray, `finalLog!${R1C1RangeToA1Range(1, 1, 1, 8)}`, loggingSpreadsheetId);
}

// creates calendars for {userCount} users
function createCalendars() {
  Logger.log('Creating calendars');
  const properties = PropertiesService.getUserProperties();
  const userCount = parseInt(properties.getProperty('userCount'));
  const groupUrl = properties.getProperty('groupUrl');
  const experimentConditionSpreadsheet = SpreadsheetApp.openById(properties.getProperty('experimentConditionSpreadsheetId'));
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
    var activeSheet = experimentConditionSpreadsheet.getSheetByName('users');
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
    var activeSheet = experimentConditionSpreadsheet.getSheetByName('users');
    activeSheet.getRange(2+i, 7).setValue(writeCalendarId);
    activeSheet.getRange(2+i, 9).setValue(`https://calendar.google.com/calendar/u/0?cid=${writeCalendarId}`);
    Logger.log(`Created write calendar ${calendar.getName()}, with the ID ${writeCalendarId}.`);
  }
}

// set ids
function setIds(property) {
  const properties = PropertiesService.getUserProperties();
  properties.setProperty('experimentConditionSpreadsheetId', property.experimentConditionSpreadsheetId);
  properties.setProperty('loggingSpreadsheetId', property.loggingSpreadsheetId);
}

// set ids manually
function setIdsManual() {
  Logger.log('manually setting spreadsheet ids');
  const properties = PropertiesService.getUserProperties();
  //property = {
  //  experimentConditionSpreadsheetId : ;
  //  loggingSpreadsheetId : ;
  //}
  properties.setProperty('experimentConditionSpreadsheetId', property.experimentConditionSpreadsheetId);
  properties.setProperty('loggingSpreadsheetId', property.loggingSpreadsheetId);
}

// delete all triggers for this script
function deleteTriggers() {
  Logger.log('deleting triggers');
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
  const writeCalendarIds = JSON.parse(properties.getProperty('writeCalendarIds'));
  
  // create trigger for each of the 18 write calendars
  // (calls function 'onCalendarEdit' on trigger)
  Logger.log(`Creating triggers. ${writeCalendarIds.length} calendar(s) triggers will be created`);
  for (var i = 0; i < writeCalendarIds.length; i++){
    const writeCalendarId = writeCalendarIds[i];
    ScriptApp.newTrigger('onCalendarEdit')
      .forUserCalendar(writeCalendarId)
      .onEventUpdated()
      .create(); 
  }
  // create 1 Sheets trigger (calls function 'onSheetsEdit' on trigger)
  ScriptApp.newTrigger('onSheetsEdit')
      .forSpreadsheet(properties.getProperty('experimentConditionSpreadsheetId'))
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

// ==============================================================================================
// ======================================= MAIN FUNCTIONS =======================================
// ==============================================================================================

// when calendar gets edited
function onCalendarEdit(e) {
  Logger.log('Calendar edit trigger');
  getAndStoreObjects(); // get sheets, calendars and store them in properties
  importMomentJS();
  const properties = PropertiesService.getUserProperties();
  const writeCalendarIds = JSON.parse(properties.getProperty('writeCalendarIds'));
  const calendarId = e.calendarId;
  const index = writeCalendarIds.indexOf(calendarId);
  const fullSync = false;
  writeEventsToReadCalendar(calendarId, index, fullSync);
}

// when sheets gets edited
function onSheetsEdit(e) {
  Logger.log('Sheets edit trigger');
  const properties = PropertiesService.getUserProperties();
  getAndStoreObjects(); // get sheets, calendars and store them in properties
  importMomentJS();
  const writeCalendarIds = JSON.parse(properties.getProperty('writeCalendarIds'));
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  const cell = e.source.getActiveRange();
  const newValue = e.value;
  const row = cell.getRow();
  const column = cell.getColumn();
  const index = row-2;
  const writeCalendarId = writeCalendarIds[index]
  const fullSync = true;
  const equipmentSheetNameFromEquipmentName = JSON.parse(properties.getProperty('equipmentSheetNameFromEquipmentName'));
  
  // when the 'users' sheet is edited
  if (sheetName === 'users') {
    onUsersSheetEdit();
    // when the checkbox (H2~nm) is edited in sheets on sheet 'users'
    // update corresponding user's subscribed equipments
    if (row > 1 && column > 9){ 
      changeSubscribedEquipments(index);
    }
    // when the full name (A2~An) is edited in sheets on sheet 'users'
    // update all of the corresponding user's event title
    else if (row > 1 && column === 1){
      updateCalendarUserName(sheet, cell, newValue);
      writeEventsToReadCalendar(writeCalendarId, index, fullSync);
    }
  }
  // when the 'properties' sheet is edited
  else if (sheetName === 'properties') {
    onPropertiesSheetEdit();
  }
  // if equipment sheet is edited
  else if (Object.values(equipmentSheetNameFromEquipmentName).includes(sheet.getName()) && row > 1){ 
    onEquipmentConditionEdit(sheet, row);
  }
}

function onUsersSheetEdit() {
  const properties = PropertiesService.getUserProperties();
  var lastRow = '';
  var lastColumn = '';
  var values = '';

  // get sheets
  const experimentConditionSpreadsheet = SpreadsheetApp.openById(properties.getProperty('experimentConditionSpreadsheetId'));
  const usersSheet = experimentConditionSpreadsheet.getSheetByName('users');

  // get all the write and read calendar's calendarIds
  lastRow = usersSheet.getLastRow();
  lastColumn = usersSheet.getLastColumn();
  values = usersSheet.getRange(2, 5, lastRow-1, 3).getValues();
  var readCalendarIds = []; // get read calendars' Ids
  var writeCalendarIds = []; // get write calendars' Ids
  var users = []; // get user names
  for (var i = 0; i < lastRow-1; i++) {
    users[i] = values[i][0];
    readCalendarIds[i] = values[i][1];
    if (i !== lastRow-2) {
      writeCalendarIds[i] = values[i][2]; // writeCalendarId in the last row is blank
    }
  }

  // get equipments that are enabled by user
  var enabledEquipmentsList = []; 
  var equipmentValues = usersSheet.getRange(1, 10, 1, lastColumn-9).getValues();
  var checkedValues = usersSheet.getRange(2, 10, lastRow-1, lastColumn-9).getValues();
  for (var i = 0; i < lastRow-1; i++) {
    enabledEquipmentsList[i] = [];
    for (var j = 0; j < lastColumn-9; j++) {
      if (checkedValues[i][j] === true) {
        enabledEquipmentsList[i].push(equipmentValues[0][j]);
      }
    }
  }

  // store objects in property
  properties.setProperty('writeCalendarIds', JSON.stringify(writeCalendarIds));
  properties.setProperty('readCalendarIds', JSON.stringify(readCalendarIds));
  properties.setProperty('usersSheet', JSON.stringify(usersSheet));
  properties.setProperty('users', JSON.stringify(users));
  properties.setProperty('enabledEquipmentsList', JSON.stringify(enabledEquipmentsList));
}

function onPropertiesSheetEdit() {  
  const properties = PropertiesService.getUserProperties();
  var lastRow = '';
  var lastColumn = '';
  var values = '';

  // get sheets
  const experimentConditionSpreadsheet = SpreadsheetApp.openById(properties.getProperty('experimentConditionSpreadsheetId'));
  const sheets = experimentConditionSpreadsheet.getSheets();
  const sheetNames = sheets.map((sheet) => {return sheet.getName()});
  const propertiesSheet = experimentConditionSpreadsheet.getSheetByName('properties');

  // get equipment sheet id for each equipment name
  var equipmentSheetIdFromEquipmentName = {};
  var equipmentSheetNameFromEquipmentName = {};
  lastRow = propertiesSheet.getLastRow();
  values = propertiesSheet.getRange(2, 1, lastRow-1, 2).getValues();

  var sheetId = '';
  var sheetName = '';
  var equipmentName = '';

  // convert sheetId to sheetName
  for (var i = 0; i < sheetNames.length; i++) {
    sheetId = experimentConditionSpreadsheet.getSheetByName(sheetNames[i]).getSheetId();
    for (var j = 0; j < lastRow-1; j++){
      equipmentName = values[j][0];
      if (sheetId === values[j][1]) {
        sheetName = sheetNames[i];
        break;  
      }
    }
    equipmentSheetIdFromEquipmentName[equipmentName] = sheetId; // get sheetId
    equipmentSheetNameFromEquipmentName[equipmentName] = sheetName; // get sheetName
  }

  // store objects in property
  properties.setProperty('propertiesSheet', JSON.stringify(propertiesSheet));
  properties.setProperty('equipmentSheetIdFromEquipmentName', JSON.stringify(equipmentSheetIdFromEquipmentName));
  properties.setProperty('equipmentSheetNameFromEquipmentName', JSON.stringify(equipmentSheetNameFromEquipmentName));
}

function onEquipmentConditionEdit(equipmentSheet, row) {
  Logger.log('Equipment condition edit trigger');
  const properties = PropertiesService.getUserProperties();
  // 4: equipment, 6: description, 7: isAllDayEvent, 8: isRecurringEvent, 9: action, 10: executionTime, 11: id, are protected from being edited
  const lastColumn = equipmentSheet.getLastColumn();
  const values = equipmentSheet.getRange(row, 1, 1, lastColumn).getValues();
  const headers = equipmentSheet.getRange(1, 1, 1, lastColumn).getValues();
  // when experiment condition gets edited -> change event summary 
  const startTime = values[0][0];
  const endTime = values[0][1];
  const user = values[0][2];
  const state = values[0][4];
  const id = values[0][10];
  const users = JSON.parse(properties.getProperty('users'));
  const writeCalendarIds = JSON.parse(properties.getProperty('writeCalendarIds'));
  const equipmentSheetIdFromEquipmentName = JSON.parse(properties.getProperty('equipmentSheetIdFromEquipmentName'));
  const experimentConditionRows = parseInt(properties.getProperty('experimentConditionRows'));
  const experimentConditionCount = parseInt(properties.getProperty('experimentConditionCount'));
  const equipmentCount = parseInt(properties.getProperty('equipmentCount'));
  const equipment = Object.keys(equipmentSheetIdFromEquipmentName).filter( (key) => { 
    return equipmentSheetIdFromEquipmentName[key] === equipmentSheet.getSheetId();
  });
  if (users.includes(user)) { // if user exists
    var writeCalendarId = writeCalendarIds[users.indexOf(user)]; // get writeCalendarId for specified user
    var writeCalendar = CalendarApp.getCalendarById(writeCalendarId);
  } else {
    Logger.log('the specified user does not exist');
    return;
  }
  if (startTime === '' || endTime === '' || user === '') {
    Logger.log('startTime, endTime, user cannot be empty')
    return;
  }
  var experimentCondition = {};
  for (var i = 0; i < lastColumn-12; i++) {
    if (headers[0][12+i] !== '') { // if condition is not ''
      experimentCondition[headers[0][12+i]] = values[0][12+i];
    }
  }
  equipmentSheet.getRange(row, 6).setValue(JSON.stringify(experimentCondition)); // set description in sheets
  if (id === '') { // 1. when experiment id doesn't exist -> add event 
    if (state === ''){
      var title = `${user} ${equipment}`;
    } else {
      var title = `${user} ${equipment} ${state}`;
    }
    var event = writeCalendar.createEvent(title, localTimeToUTC(startTime), localTimeToUTC(endTime), {description: experimentCondition});
  } else { // 2. when experiment id exists -> modify event
    var event = writeCalendar.getEventById(id);
    event.setDescription(JSON.stringify(experimentCondition)); // save experiment condition as stringified JSON
    if (event === null) {
      Logger.log('the specified event id does not exist');
    } else {
      event.setTime(localTimeToUTC(startTime), localTimeToUTC(endTime))// edit start and end time
      event.setTitle(`${user} ${equipment} ${state}`); // edit equipmentName, name ,state
      event.setDescription(JSON.stringify(experimentCondition)); // write experiment condition in details
    }
  }
  Logger.log('Sorting events');
  const allEquipmentsSheet = SpreadsheetApp.openById(properties.getProperty('experimentConditionSpreadsheetId')).getSheetByName('allEquipments');
}  

// write events to read calendar based on updated events in write calendar
function writeEventsToReadCalendar(writeCalendarId, index, fullSync) {
  const properties = PropertiesService.getUserProperties();
  const readCalendarIds = JSON.parse(properties.getProperty('readCalendarIds'));
  const enabledEquipmentsList = JSON.parse(properties.getProperty('enabledEquipmentsList'));
  const users = JSON.parse(properties.getProperty('users'));
  const writeUser = users[index];
  const allEvents = getEvents(writeCalendarId, fullSync);
  const events = allEvents.events;
  const canceledEvents = allEvents.canceledEvents;
  Logger.log(`${readCalendarIds.length} read calendars`);
  const equipmentSheetNameFromEquipmentName = JSON.parse(properties.getProperty('equipmentSheetNameFromEquipmentName'));
  for (var i = 0; i < events.length; i++){
    var event = events[i];
    const filteredReadCalendarIds = filterUsers(writeUser, event, readCalendarIds, enabledEquipmentsList).filteredReadCalendarIds;
    Logger.log(`writing event no.${i+1} to [ ${filteredReadCalendarIds} ]`);
    writeEvent(event, writeCalendarId, writeUser, filteredReadCalendarIds); // create event in write calendar and add read calendars as guests
  }
  const writeCalendar = CalendarApp.getCalendarById(writeCalendarId);
  for (var i = 0; i < events.length; i++) { // log events
    var event = events[i];
    const equipmentStatus = getEquipmentNameAndStateFromEvent(event);
    const equipmentName = equipmentStatus.equipmentName;
    const state = equipmentStatus.state;
    const action = 'add';
    const eid = event.getId();
    event = writeCalendar.getEventById(eid);
    eventLoggingStoreData({ // log event
      startTime: UTCToLocalTime(event.getStartTime()),
      endTime: UTCToLocalTime(event.getEndTime()),
      name: writeUser,
      equipmentName: equipmentName,
      state: state,
      description: event.getDescription(),
      isAllDayEvent: event.isAllDayEvent(),
      isRecurringEvent: event.isRecurringEvent(),
      action: action, 
      executionTime: UTCToLocalTime(new Date()), // current time
      id: eid, 
    });
    eventLoggingExecute(equipmentSheetNameFromEquipmentName[equipmentName]);
  }
  for (var i = 0; i < canceledEvents.length; i++) { // log canceled events
    var event = canceledEvents[i];
    const equipmentStatus = getEquipmentNameAndStateFromEvent(event);
    const equipmentName = equipmentStatus.equipmentName;
    const state = equipmentStatus.state;
    const action = 'cancel';
    const eid = event.getId();
    event = writeCalendar.getEventById(eid);
    eventLoggingStoreData({ // log event
      startTime: UTCToLocalTime(event.getStartTime()),
      endTime: UTCToLocalTime(event.getEndTime()),
      name: writeUser,
      equipmentName: equipmentName,
      state: state,
      description: event.getDescription(),
      isAllDayEvent: event.isAllDayEvent(),
      isRecurringEvent: event.isRecurringEvent(),
      action: action, 
      executionTime: UTCToLocalTime(new Date()), // current time
      id: eid, 
    });
    eventLoggingExecute(equipmentSheetNameFromEquipmentName[equipmentName]);
  }
  Logger.log('event logging done');
  updateSyncToken(writeCalendarId); // renew sync token after adding guest
  Logger.log(`Wrote updated events to read calendar. Fullsync = ${fullSync}`);
}

// update corresponding user's subscribed equipments 
function changeSubscribedEquipments(index){
  const properties = PropertiesService.getUserProperties();
  const readCalendarIds = JSON.parse(properties.getProperty('readCalendarIds'));
  const enabledEquipmentsList = JSON.parse(properties.getProperty('enabledEquipmentsList'));
  const fullSync = true;
  const writeCalendarIds = JSON.parse(properties.getProperty('writeCalendarIds'));
  const readCalendarId = readCalendarIds[index];
  const users = JSON.parse(properties.getProperty('users'));
  Logger.log(`${writeCalendarIds.length} write calendars`);
  for (var i = 0; i < writeCalendarIds.length; i++){
    const writeUser = users[i];
    const writeCalendarId = writeCalendarIds[i];
    const allEvents = getEvents(writeCalendarId, fullSync);
    const events = allEvents.events;
    for (var j = 0; j < events.length; j++){
      const event = events[j];
      const filteredReadCalendarIds = filterUsers(writeUser, event, [readCalendarId], enabledEquipmentsList).filteredReadCalendarIds;
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

// =================================================================================================
// ======================================= LOGGING FUNCTIONS ======================================= 
// =================================================================================================

// set data for logging
function eventLoggingStoreData(logObj) { 
  const properties = PropertiesService.getUserProperties();
  if (properties.getProperty('eventLoggingData') === null) { // if key value pair not defined
    var eventLoggingData = {}; // initialize
  } else {
    var eventLoggingData = JSON.parse(properties.getProperty('eventLoggingData'));    
  }
  for (const key in logObj) { // iterate through log object
    eventLoggingData[key] = logObj[key];
  }
  properties.setProperty('eventLoggingData', JSON.stringify(eventLoggingData));
}

// execute logging to sheets
function eventLoggingExecute(equipmentSheetName) { 
  Logger.log('Logging event');
  const properties = PropertiesService.getUserProperties();
  const experimentConditionCount = parseInt(properties.getProperty('experimentConditionCount'));
  const experimentConditionRows = parseInt(properties.getProperty('experimentConditionRows'));
  const equipmentCount = parseInt(properties.getProperty('equipmentCount'));
  const experimentConditionBackupRows = parseInt(properties.getProperty('experimentConditionBackupRows'));
  // spreadsheet for experiment condition logging
  const experimentConditionSpreadsheetId = properties.getProperty('experimentConditionSpreadsheetId');
  const equipmentSheet = SpreadsheetApp.openById(experimentConditionSpreadsheetId).getSheetByName(equipmentSheetName);
  const allEquipmentsSheet = SpreadsheetApp.openById(properties.getProperty('experimentConditionSpreadsheetId')).getSheetByName('allEquipments');
  const eventLoggingData = JSON.parse(properties.getProperty('eventLoggingData'));
  properties.deleteProperty('eventLoggingData');
  var row = equipmentSheet.getRange("A1:A").getValues().filter(String).length + 1; // get last row of first column
  if (row > experimentConditionBackupRows) { // prevent overflow of spreadsheet data by backing up and deleting it
    backupAndDeleteOverflownLoggingData(equipmentSheet) 
    row = equipmentSheet.getRange("A1:A").getValues().filter(String).length + 1; // get last row of first column
  }
  const columnDescriptions = { // shows which description corresponds to which column
    startTime: 1,
    endTime: 2,
    name: 3,
    equipmentName: 4,
    state: 5,
    description: 6,
    isAllDayEvent: 7,
    isRecurringEvent: 8,
    action: 9,
    executionTime: 10,
    id: 11,
  };
  var filledArray = [[]];
  for (const key in eventLoggingData) { // iterate through log object
    var value = eventLoggingData[key];
    var col = columnDescriptions[key];
    filledArray[0][col-1] = value;
  }
  Logger.log(eventLoggingData);
  setValues(filledArray, `${equipmentSheetName}!${R1C1RangeToA1Range(row, 1, 1, 11)}`, experimentConditionSpreadsheetId); 

  // get experiment condition from description
  try {
    const description = JSON.parse(eventLoggingData['description']);
    const descriptionHeaders = equipmentSheet.getRange(1, 13, 1, experimentConditionCount).getValues();
    var descriptionHeader = '';
    var filledArray = [[]];
    for (var i = 0; i < experimentConditionCount; i++){
      descriptionHeader = descriptionHeaders[0][i]; // experiment condition header
      if (descriptionHeader in description) { // if value exists for key 'descriptionHeader'
        filledArray[0][i] = description[descriptionHeader];;
      } else {
        filledArray[0][i] = '';
      }
    }
    setValues(filledArray, `${equipmentSheetName}!${R1C1RangeToA1Range(row, 13, 1, experimentConditionCount)}`, experimentConditionSpreadsheetId); // experiment condition
  } catch (e) {}

  Logger.log('Sorting events');
}

// logs just the necessary data
function finalLogging() { 
  Logger.log('Daily logging of event');
  getAndStoreObjects(); // get sheets, calendars and store them in properties
  const properties = PropertiesService.getUserProperties();
  const finalLoggingBackupRows = parseInt(properties.getProperty('finalLoggingBackupRows'));
  const loggingSpreadsheetId = properties.getProperty('loggingSpreadsheetId');
  const finalLogSheet = SpreadsheetApp.openById(loggingSpreadsheetId).getSheetByName('finalLog');
  const lastRow = finalLogSheet.getLastRow();
  if (lastRow > finalLoggingBackupRows) {
    backupAndDeleteOverflownLoggingData(finalLogSheet) // prevent overflow of spreadsheet data by backing up and deleting it
  }
  const row = lastRow + 1; // write on new row
  const columnDescriptions = { // shows which description corresponds to which column
    startTime: 1,
    endTime: 2,
    name: 3,
    equipmentName: 4,
    state: 5,
    description: 6,
    isAllDayEvent: 7,
    isRecurringEvent: 8,
    action: 9,
    executionTime: 10,
    id: 11,
  };
  const writeCalendarIds = JSON.parse(properties.getProperty('writeCalendarIds'));
  const users = JSON.parse(properties.getProperty('users'));
  // get events from 2~3 days ago
  options = {
    timeMin : getRelativeDate(-3, 0).toISOString(), // 3 days ago
    timeMax : getRelativeDate(-2, 0).toISOString(), // 2 days ago
    showDeleted : true,
  }
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
      const equipmentStateFromEvent = getEquipmentNameAndStateFromEvent(event); // this must be called first before getCalendarById
      const equipmentName = equipmentStateFromEvent.equipmentName;
      const state = equipmentStateFromEvent.state;
      const eid = event.iCalUID;
      event = CalendarApp.getCalendarById(writeCalendarId).getEventById(eid);
      var logObj = {
        startTime: UTCToLocalTime(event.getStartTime()),
        endTime: UTCToLocalTime(event.getEndTime()),
          name: writeUser,
          equipmentName: equipmentName,
          state: state,
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
      setValues(filledArray, `finalLog!${R1C1RangeToA1Range(row, 1, 1, 11)}`, loggingSpreadsheetId);
    }
  }
}

// ===============================================================================================
// ======================================= OTHER FUNCTIONS ======================================= 
// ===============================================================================================

// prevent overflow of spreadsheet data by backing up and deleting it
function backupAndDeleteOverflownEquipmentData(equipmentSheet) {
  const properties = PropertiesService.getUserProperties();
  // backup rows
  const equipment = equipmentSheet.getRange(2, 4).getValue();
  const startTime = localTimeToUTC(equipmentSheet.getRange(2, 1).getValue());
  const endTime = localTimeToUTC(equipmentSheet.getRange(2+experimentConditionBackupRows-1, 1).getValue());
  const experimentConditionBackupRows = parseInt(properties.getProperty('experimentConditionBackupRows'));
  const backupColumns = equipmentSheet.getLastColumn();
  equipmentSheet.getRange(1, 1, experimentConditionBackupRows+1, backupColumns).copyTo(
    SpreadsheetApp.create(`BACKUP_${equipment}_${startTime}-${endTime}`),
    SpreadsheetApp.CopyPasteType.PASTE_VALUES, 
    false
  )
  // delete rows
  var filledArray = [];
  filledArray = arrayFill2d(experimentConditionBackupRows, 11, '');
  equipmentSheet.getRange(2, 1, experimentConditionBackupRows, 11).setValues(filledArray);
  filledArray = arrayFill2d(experimentConditionBackupRows, experimentConditionCount, '');
  equipmentSheet.getRange(13, 1, experimentConditionBackupRows, experimentConditionCount).setValues(filledArray);
}

// prevent overflow of spreadsheet data by backing up and deleting it
function backupAndDeleteOverflownLoggingData(finalLogSheet) {
  // backup rows
  const backupRows = finalLogSheet.getLastRow()-1;
  const backupColumns = finalLogSheet.getLastColumn();
  const startTime = localTimeToUTC(finalLogSheet.getRange(2, 1).getValue());
  const endTime = localTimeToUTC(finalLogSheet.getRange(2+backupRows-1, 1).getValue());
  finalLogSheet.getRange(1, 1, backupRows+1, backupColumns).copyTo(
    SpreadsheetApp.create(`BACKUP_LOG_${startTime}-${endTime}`),
    SpreadsheetApp.CopyPasteType.PASTE_VALUES, 
    false
  )
  // delete rows
  var filledArray = [];
  filledArray = arrayFill2d(backupRows, 8, '');
  finalLogSheet.getRange(2, 1, backupRows, 8).setValues(filledArray);
}

// ================================================================================================
// ======================================= HELPER FUNCTIONS ======================================= 
// ================================================================================================

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
  sheet.getRange(1, 1, rows, columns).setHorizontalAlignment("center"); // center text
  sheet.getRange(1, 1, rows, columns).setVerticalAlignment("middle"); // center text
}

// protect and color the specified range
function protectRange(range) {
  const properties = PropertiesService.getUserProperties();
  range.protect().setDescription('Protected Range').addEditor(Session.getEffectiveUser());
  range.setBackground(properties.getProperty('backgroundColor'));
}

// create 2d array filled with value
function arrayFill2d(rows, columns, value) { 
  return Array(rows).fill().map(() => Array(columns).fill(value));
}

// fill array with value in sheetsAPI
function setValues(filledArray, range, spreadsheetId) {
  Sheets.Spreadsheets.Values.update(
    {"majorDimension": "ROWS", "values": filledArray},
    spreadsheetId, 
    range,
    {"valueInputOption": "USER_ENTERED"},
  );
}

// fill array with value in sheetsAPI 
function setValuesBatch(filledArrayBatch, spreadsheetId) {
  Sheets.Spreadsheets.Values.batchUpdate(
    {"valueInputOption": "USER_ENTERED", "data": filledArrayBatch},
    spreadsheetId, 
  );
}

// update the sync token after adding and deleting guests
function updateSyncToken(calendarId) {
  const properties = PropertiesService.getUserProperties();
  const options = {
    maxResults: 1000, // suppress nextPageToken which supresses nextSyncToken by fitting all events in one page
    showDeleted: true,
  };
  var eventsList;
  eventsList = Calendar.Events.list(calendarId, options);
  properties.setProperty(`syncToken ${calendarId}`, eventsList.nextSyncToken);
  Logger.log('Updated sync token. New sync token: ' + eventsList.nextSyncToken);
}

// get sheets and store them in properties
function getAndStoreObjects() {
  onUsersSheetEdit();
  onPropertiesSheetEdit();
}

// filter readUsers who are not writeUser and have the equipment
function filterUsers(writeUser, event, readCalendarIds, enabledEquipmentsList) {
  const properties = PropertiesService.getUserProperties();
  const users = JSON.parse(properties.getProperty('users'));
  var filteredReadCalendarIds = []; // readCalendarIds excluding the same user as writeCalendarId
  var filteredReadUsers = []; // readUsers excluding the same user as writeCalendarId
  const equipmentName = getEquipmentNameAndStateFromEvent(event).equipmentName; // equipment used in the event
  for (var i = 0; i < readCalendarIds.length; i++){
    const readCalendarId = readCalendarIds[i];
    const readUser = users[i];
    const enabledEquipments = enabledEquipmentsList[i];
    if (enabledEquipments.includes(equipmentName) === true){ // if equipment is enabled for readUser
      if (readUser != writeUser) { // avoid duplicating event for same user's read and write calendars      
        filteredReadCalendarIds.push(readCalendarId);
        filteredReadUsers.push(readUser);
      }
    }
  }
  return {filteredReadCalendarIds, filteredReadUsers}
}

// get events from the given calendar that have been modified since the last sync.
// if the sync token is missing or invalid, log all events from up to a ten days ago (a full sync).
function getEvents(calendarId, fullSync) {
  const properties = PropertiesService.getUserProperties();
  const options = {
    maxResults: 100,
    showDeleted : true,
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
  var canceledEvents = [];
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
          Logger.log(`Event id ${event.id} was cancelled.`);
          canceledEvents.push(event);
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
  return {canceledEvents, events};
}

// get equipment name and state from event summary
function getEquipmentNameAndStateFromEvent(event){
  const summary = event.summary;  
  const status = summary.split(' '); // split to equipment and state
  if (status.length === 1) { // just the equipment name (state is 'use')
    var equipmentName = status[0];
    var state = 'use';
  } else if (status.length === 2 || status.length === 3) { // (User Name) + equipment + state
    var equipmentName = status[status.length-2];
    var state = status[status.length-1];
  }
  return {equipmentName, state};
}

// rename and write event in write calendar and add read calendars as guests
function writeEvent(event, writeCalendarId, writeUser, readCalendarIds) {
  Utilities.sleep(100);
  const equipmentStateFromEvent = getEquipmentNameAndStateFromEvent(event);
  const equipmentName = equipmentStateFromEvent.equipmentName;
  const state = equipmentStateFromEvent.state;
  const eid = event.iCalUID;
  var event = CalendarApp.getCalendarById(writeCalendarId).getEventById(eid);
  // if equipment is enabled in sheets, add to guest subscription
  // change title from '(User Name) + equipment + state' to 'User Name + equipment + state'
  const summary = `${writeUser} ${equipmentName} ${state}`;
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
  const properties = PropertiesService.getUserProperties();
  const experimentConditionSpreadsheetId = properties.getProperty('experimentConditionSpreadsheetId');
  const names = newValue.split(' ', 2); // split to last and first name
  const lastName = names[1];
  const firstName = names[0];
  const row = cell.getRow();
  // set User Name 1 using last and first name
  // User Name 1 = {Last Name up to 4 letters}.{First Name up to 1 letter}
  const filledArray = [[lastName, firstName, lastName.slice(0,4)+'.'+firstName.slice(0,1)]];
  setValues(filledArray, `users!${R1C1RangeToA1Range(row, 2, 1, 3)}`, experimentConditionSpreadsheetId);
}

// set User Name 2 using User Name 1
// User Name 2 = {User Name 1}{unique identifier 1~9}
//             = {Last Name up to 4 letters}.{First Name up to 1 letter}{unique identifier 1,2,3,...}
function setUserNames(sheet){
  const properties = PropertiesService.getUserProperties();
  const experimentConditionSpreadsheetId = properties.getProperty('experimentConditionSpreadsheetId');
  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(2, 4, lastRow-1).getValues();
  const filledArray = [];
  // update User Name 2 for row0 = 2~lastRow
  for (var i = 0; i < lastRow-1; i++){
    // check the duplicate count of value0 for row1 = 2~row0
    var count = 0; // duplicate count
    for (var j = 0; j < i + 1; j++){
      if (values[i][0] === values[j][0]){
        count += 1;
      }
    }
    filledArray[i] = [`${values[i][0]}${count}`]; // name + unique number
  }
  // use count as unique identifier (1,2,3,...)
  setValues(filledArray, `users!${R1C1RangeToA1Range(2, 5, lastRow-1)}`, experimentConditionSpreadsheetId);
}

// create checkboxes for selecting which equipment to show in the calendar
function setCheckboxes(sheet, cell) {
  const lastColumn = sheet.getLastColumn();
  const row = cell.getRow();
  if (sheet.getRange(row, 10).isChecked() == null){ // if cell is not a checkbox
    // create checkboxes
    sheet.getRange(row, column, 1, lastColumn-9).insertCheckboxes();
  }
  Logger.log('Created checkboxes');
}

// set read calendar and write calendar for created user
function setCalendars(sheet, cell) {
  const row = cell.getRow();
  const values = sheet.getRange(row, 5, 1, 3).getValues();
  const userName = values[0][0];
  const readCalendarId = values[0][1];
  const writeCalendarId = values[0][2];
  changeCalendarName(readCalendarId, userName, 'Read');
  changeCalendarName(writeCalendarId, userName, 'Write');
}

// update calendar name and description
function changeCalendarName(calendarId, userName, readOrWrite) {
  // calendar name (summary) and description changes for read and write calendar
  if (calendarId !== '') { // write calendarId is empty for "all event" calendar
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
    calendar.setDescription(JSON.stringify(description));
    Logger.log('Updated calendar name');
  }
  Logger.log('Skipped update of calendar name because calendarId is empty');
}

function importMomentJS() {
  Logger.log('Start importing Moment, Moment-timezone, Moment-timezone-with-data');
  eval(UrlFetchApp.fetch('https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.29.1/moment.min.js').getContentText());
  eval(UrlFetchApp.fetch('https://cdnjs.cloudflare.com/ajax/libs/moment-timezone/0.5.34/moment-timezone.min.js').getContentText());
  eval(UrlFetchApp.fetch('https://cdnjs.cloudflare.com/ajax/libs/moment-timezone/0.5.34/moment-timezone-with-data.js').getContentText());
  Logger.log('Done importing');
}

// parse MM/DD/YY HH:mm (local time) to 0000-00-00T00:00:00.000Z (ISO-8601) UTC
// string -> Date object
function localTimeToUTC(inputDate) {
  const properties = PropertiesService.getUserProperties();
  const timeZone = properties.getProperty('timeZone');
  const outputDate = moment.tz(inputDate, "MM/DD/YY HH:mm", timeZone).toDate(); // toDate outputs utc time
  return outputDate
}

// parse 0000-00-00T00:00:00.000Z (ISO-8601) UTC to MM/DD/YY HH:mm (local time)
// Date object -> string
function UTCToLocalTime(inputDate) {
  const properties = PropertiesService.getUserProperties();
  const timeZone = properties.getProperty('timeZone');
  const outputDate = moment(inputDate).tz(timeZone).format("MM/DD/YY HH:mm");
  return outputDate
}

// timed trigger after 10 seconds
function timedTrigger(functionName) {  
  ScriptApp
    .newTrigger(functionName)
    .timeBased()
    .at(new Date((new Date()).getTime()+10000))
    .create();
  Logger.log(`trigger set for ${functionName}`);
}

// convert R1C1 range notation to A1 range string notation
function R1C1RangeToA1Range(row, column, rowCount, columnCount) {
  return `${R1C1ToA1(row, column)}:${R1C1ToA1(row+rowCount-1, column+columnCount-1)}`
}

// convert R1C1 notation to A1 notation
function R1C1ToA1(row, column) {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  let columnRef = '';
  column -= 1;

  if (column < 1) {
    columnRef = chars[0];
  }
  else {
    const base26 = column.toString(26);
    const digits = base26.split('');
    columnRef = digits.map((digit, i) => chars[parseInt(digit, 26) - (i === digits.length-1?0:1)]).join('');
  }

  return `${columnRef}${Math.max(row, 1)}`;
};
