function isValidDate(d) {
  if ( Object.prototype.toString.call(d) !== "[object Date]" )
    return false;
  return !isNaN(d.getTime());
}

function syncTabs() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  if (sheets.length == 0) {
    return;
  }

  for (var index = 0; index < sheets.length; index++) {
    var sheet = sheets[index];
    var curYear = new Date().getUTCFullYear()
    var name = sheet.getName();
    if (containsDate(name)) {
      var curDate = new Date(name)
      curDate.setUTCFullYear(curYear)
      curDate.setHours(6)
      syncCalendar(sheet, curDate, isSunday(name));
    }
  }
}

function containsDate(name) {
  var re = /^\w{3} \d+\/\d+$/;
  return re.test(name);
}

function isSunday(name) {
  var re = /^SUN \d+\/\d+$/;
  return re.test(name)
}

function syncCalendar(sheet, curDateValue, isSunday) {
  var myCalendar = CalendarApp.getCalendarsByName("ucb_klesis_calendar")[0]; // Put the Calendar name here
  var studentCalendar = CalendarApp.getCalendarsByName("Berkeley Klesis")[0];
  if( sheet == null || myCalendar == null) {return;}
  var firstColumn = "A";
  var lastColumn = isSunday ? "K" : "I";
  var startTimeIndex = 0;
  var endTimeIndex = 1;
  var eventNameIndex = 2;
  var eventLocationIndex = 3;
  var eventInChargeIndex = 4;
  var eventHelpersIndex = 5;
  var eventNoteIndex = 6;
  var shouldPutInWUDIndex = isSunday ? 8: 7;
  var shouldPutInStudentCalIndex = isSunday ? 9 : 8;

  var hasExtraColumn1 = true;
  var descriptionText1 = "In Charge: "; // Modify these as needed for each column
  var hasExtraColumn2 = true;
  var descriptionText2 = "Helpers: ";
  var hasExtraColumn3 = true;
  var descriptionText3 = "Notes: ";

  var columnRange = firstColumn + ":" + lastColumn;
  var WUDRange = sheet.getRange(columnRange);
  var allCellsInWUD = WUDRange.getValues();
  var allFormulas = WUDRange.getFormulasR1C1();
  
  var numRows = sheet.getLastRow();
  var index = 2;

  var eventName = null;
  var eventStart = null;
  var eventEnd = null;
  var eventLoc = null;
  var events = null;
  var toCreate = {};
  var studentToCreate = {};
  var missingEnd = false;
  var toDelete = {};
  var toStudentDelete = {};
  var descriptionText; 
  var today = new Date();
  var key;
  
  while( index < numRows) {
    eventStart = allCellsInWUD[index][startTimeIndex];
    eventEnd = allCellsInWUD[index][endTimeIndex];
    eventName = allCellsInWUD[index][eventNameIndex];      
    
    if(eventStart == "") {
      index++;
      continue;
    }
      
    // should we put on gcal
    row = allCellsInWUD[index];
    shouldWePutOnGcal = allCellsInWUD[index][shouldPutInWUDIndex];
    shouldWePutOnStudentCal = allCellsInWUD[index][shouldPutInStudentCalIndex];
    if (!shouldWePutOnGcal && !shouldWePutOnStudentCal) {
      index++;
      continue;
    }
      

      
    if (eventEnd == "" || !isValidDate(eventEnd)) {
      eventEnd = new Date(eventStart);
      missingEnd = true;
    }
    // sometimes these have issues with timezones. You can use getUTCHours() and getUTCMinutes() if you're running into timezone issues
    var isEasternTime = eventStart.getHours() >= 0 && eventStart.getHours() <= 3; // timezone issues with EST..
    eventStart = new Date(curDateValue.getUTCFullYear(), curDateValue.getUTCMonth(), isEasternTime ? curDateValue.getUTCDate() + 1 : curDateValue.getUTCDate(), eventStart.getHours(), eventStart.getMinutes());
    eventEnd = new Date(curDateValue.getUTCFullYear(), curDateValue.getUTCMonth(), isEasternTime ?curDateValue.getUTCDate() + 1 : curDateValue.getUTCDate(), eventEnd.getHours(), eventEnd.getMinutes());
    if (missingEnd) {
      eventEnd.setTime(eventStart.getTime() + 60*60*1000); // add an hour if no end time is set
      missingEnd = false;
    }
    if (eventEnd.getTime() < eventStart.getTime()) { //hack fix for when the end time is 12AM or something. Add 24 hours to move it to next day.
      eventEnd.setTime(eventEnd.getTime() + 1000*60*60*24);
    }
    eventLoc = allCellsInWUD[index][eventLocationIndex];
    if (allFormulas[index][eventLocationIndex]) {
      // Zoom link or online link provided
      var formula = allFormulas[index][4];
      var linkRegex = "=HYPERLINK\\(\"(.*)\",\"(.*)\"\\)";
      var link = formula.match(linkRegex)[1];
      eventLoc = link
    }
    descriptionText = "";
    studentDescriptionText = ""
    if (hasExtraColumn1) {
      descriptionText += "<b>" + descriptionText1 + "</b>" + allCellsInWUD[index][eventInChargeIndex];
    }
    if (hasExtraColumn2) {
      descriptionText += "\n\n<b>" + descriptionText2 + "</b>" + allCellsInWUD[index][eventHelpersIndex];
    }
    if (hasExtraColumn3) {
      descriptionText += "\n\n<b>" + descriptionText3 + "</b>" + allCellsInWUD[index][eventNoteIndex];
      studentDescriptionText += "<b>" + descriptionText3 + "</b>" + allCellsInWUD[index][eventNoteIndex];
    }
    if (allFormulas[index][eventLocationIndex]) {
      descriptionText += "\n\n<b>Link for meeting: </b>" + eventLoc;
    }
    key = "" + eventStart.getTime() + ";" + eventEnd.getTime() + ";" + eventName;
    if (shouldWePutOnGcal) {
      toCreate[key] = {name: eventName, start: eventStart, end: eventEnd, 
                    options: {location: eventLoc, 
                              description: descriptionText}, date: curDateValue, allDay: false};
    }
    if (shouldWePutOnStudentCal) {
      studentToCreate[key] = {name: eventName, start: eventStart, end: eventEnd, 
                    options: {location: eventLoc, 
                              description: studentDescriptionText}, date: curDateValue, allDay: false};
    }

    index++;
  }
  

  events = myCalendar.getEventsForDay(new Date(curDateValue.getUTCFullYear(), curDateValue.getUTCMonth(), curDateValue.getUTCDate() + 1));
  testingDate = curDateValue.getDate()
  studentEvents = studentCalendar.getEventsForDay(curDateValue);
  for(var e in events){
    if (events[e].isAllDayEvent())
      key = "allday;" + events[e].getAllDayStartDate().getDate() + ";" + events[e].getTitle();
    else
      key = "" + events[e].getStartTime().getTime() + ";" + events[e].getEndTime().getTime() + ";" + events[e].getTitle();
    if (!(key in toCreate) && (events[e].getStartTime().valueOf() >= curDateValue.getTime().valueOf() || events[e].isAllDayEvent())) {
      // can update who's in charge, location, and who's involved if needed, then pop the event
      toDelete[events[e].getId()] = events[e];
    } else if (toCreate[key]) {
      if (events[e].getLocation() != toCreate[key]['options']['location'] || events[e].getDescription() != toCreate[key]['options']['description']) {
        events[e].setLocation(toCreate[key]['options']['location']);
        events[e].setDescription(toCreate[key]['options']['description']);
      }
      delete toCreate[key];
    }

  }

  for(var e in studentEvents){
    if (studentEvents[e].isAllDayEvent())
      key = "allday;" + studentEvents[e].getAllDayStartDate().getDate() + ";" + studentEvents[e].getTitle();
    else
      key = "" + studentEvents[e].getStartTime().getTime() + ";" + studentEvents[e].getEndTime().getTime() + ";" + studentEvents[e].getTitle();
    if (!(key in studentToCreate) && (studentEvents[e].getStartTime().valueOf() >= curDateValue.getTime().valueOf() || studentEvents[e].isAllDayEvent())) {
      // can update who's in charge, location, and who's involved if needed, then pop the event
      toStudentDelete[studentEvents[e].getId()] = studentEvents[e];
    } else if (studentToCreate[key]) {
      if (studentEvents[e].getLocation() != studentToCreate[key]['options']['location'] || studentEvents[e].getDescription() != studentToCreate[key]['options']['description']) {
        studentEvents[e].setLocation(studentToCreate[key]['options']['location']);
        studentEvents[e].setDescription(studentToCreate[key]['options']['description']);
      }
      delete studentToCreate[key];
    }

  }
    
  var counter = 0;
  for (var e in toCreate) {
    if (toCreate[e]['allDay']) {
      myCalendar.createAllDayEvent(toCreate[e]['name'], toCreate[e]['date'], toCreate[e]['end'], toCreate[e]['options']);
      counter++;
      if (counter % 20 == 0)
        Utilities.sleep(1000);
    } else {
      myCalendar.createEvent(toCreate[e]['name'], toCreate[e]['start'], toCreate[e]['end'], toCreate[e]['options']);
      counter++;
      if (counter % 20 == 0)
        Utilities.sleep(1000);
    }
  }

  for (var e in studentToCreate) {
    if (studentToCreate[e]['allDay']) {
      studentCalendar.createAllDayEvent(studentToCreate[e]['name'], studentToCreate[e]['date'], studentToCreate[e]['end'], studentToCreate[e]['options']);
      counter++;
      if (counter % 20 == 0)
        Utilities.sleep(1000);
    } else {
      studentCalendar.createEvent(studentToCreate[e]['name'], studentToCreate[e]['start'], studentToCreate[e]['end'], studentToCreate[e]['options']);
      counter++;
      if (counter % 20 == 0)
        Utilities.sleep(1000);
    }
  }
  for(var key in toDelete){
    toDelete[key].deleteEvent();
    counter++;
    if (counter % 20 == 0)
      Utilities.sleep(1000);
  }

  for (var key in toStudentDelete) {
    toStudentDelete[key].deleteEvent();
    counter++;
    if (counter % 20 == 0)
      Utilities.sleep(1000);
  }
}

// for debugging purposes. Delete all events in the last 8 days and 20 days from now
function clearAllEvents() {
  var myCalendar = CalendarApp.getCalendarsByName("test_aym_calendar")[0];
  var now = new Date();
  events = myCalendar.getEvents(new Date(now.getTime() - 8*24*60*60*1000), new Date(now.getTime() + 20*24*60*60*1000));
  for(var e in events){
    events[e].deleteEvent();
  }
}