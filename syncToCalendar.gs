function isValidDate(d) {
  if ( Object.prototype.toString.call(d) !== "[object Date]" )
    return false;
  return !isNaN(d.getTime());
}

function syncCalendarScript() { 
  function convertDateToUTC(date) { return new Date(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate(), date.getUTCHours(), date.getUTCMinutes(), date.getUTCSeconds()); }

  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("WUD"); // Put the tab name here
  let myCalendar = CalendarApp.getCalendarsByName("ucb_klesis_calendar")[0]; // Put the Calendar name here
  let studentCalendar = CalendarApp.getCalendarsByName("Berkeley Klesis")[0];
  if( sheet == null || myCalendar == null) {return;}
  let firstColumn = "A";
  let lastColumn = "J";

  let hasExtraColumn1 = true;
  let descriptionText1 = "In Charge: "; // Modify these as needed for each column
  let hasExtraColumn2 = true;
  let descriptionText2 = "Helpers: ";
  let hasExtraColumn3 = true;
  let descriptionText3 = "Notes: ";

  let columnRange = firstColumn + ":" + lastColumn;
  let WUDRange = sheet.getRange(columnRange);
  let allCellsInWUD = WUDRange.getValues();
  let allFormulas = WUDRange.getFormulasR1C1();
  
  let numRows = sheet.getLastRow();
  let index = 0;
  let dateOnCell = null;
  let curDateValue = null;
  let eventName = null;
  let eventStart = null;
  let eventEnd = null;
  let eventLoc = null;
  let events = null;
  let toCreate = {};
  let studentToCreate = {};
  let firstDateValue = null;
  let missingEnd = false;
  let toDelete = {};
  let toStudentDelete = {};
  let descriptionText; 
  let today = new Date();
  let threeDaysAgo = today.getTime() - 1000*60*60*24*3;
  let twoWeeksAfter = today.getTime() + 1000 * 60 * 60 * 24 * 24;
  let key;
  let previousEventName = "";
  let previousAllDate;
  let previousKey;
  
  while (index < numRows) {
    dateOnCell = allCellsInWUD[index][0];
    if(isValidDate(dateOnCell)){
      let tempDate = convertDateToUTC(dateOnCell);
      if (tempDate.getTime() >= threeDaysAgo) {
        break;
      }
    }
    index++;
  }
  
  while( index < numRows)
  {
    dateOnCell = allCellsInWUD[index][0];
    if(curDateValue != null || isValidDate(dateOnCell)) {
      if(isValidDate(dateOnCell)){
        curDateValue = convertDateToUTC(dateOnCell);
        if (firstDateValue == null) {
          firstDateValue = new Date(curDateValue);
          firstDateValue.setHours(0,0,0,0);
        }
        
        if (curDateValue.getTime() > twoWeeksAfter) {
          break;
        }
      }
      
      eventStart = allCellsInWUD[index][1];
      eventEnd = allCellsInWUD[index][2];
      eventName = allCellsInWUD[index][3];      
      
      if(eventStart == "") {
        index++;
        continue;
      }
      
      // should we put on gcal
      row = allCellsInWUD[index];
      shouldWePutOnGcal = allCellsInWUD[index][8];
      shouldWePutOnStudentCal = allCellsInWUD[index][9];
      if (!shouldWePutOnGcal && !shouldWePutOnStudentCal) {
        index++;
        continue;
      }
      
      if (!isValidDate(eventStart)) {
        if (eventStart && (eventStart.toString().toLowerCase().trim().indexOf("all day") == 0 || eventStart.toString().trim().toLowerCase().indexOf("allday") == 0)) {
          eventLoc = allCellsInWUD[index][4];
          descriptionText = "";
          if (hasExtraColumn1) {
            descriptionText += descriptionText1 + allCellsInWUD[index][5];
          }
          if (hasExtraColumn2) {
            descriptionText += "\n" + descriptionText2 + allCellsInWUD[index][6];
          }
          if (hasExtraColumn3) {
            //descriptionText += "\n" + descriptionText3 + allCellsInWUD[index][7];
          }
          eventEnd = new Date(curDateValue.getUTCFullYear(), curDateValue.getUTCMonth(), curDateValue.getUTCDate()+1, curDateValue.getHours(), 0);
          
          if (previousEventName.indexOf(eventName) == 0 && (!previousAllDate || previousAllDate.getDate() === curDateValue.getDate())) {
            key = previousKey;
            if (toCreate[key]) {
              toCreate[key].end = eventEnd;
              previousAllDate = eventEnd;
            } else {
              Logger.log("Key doesn't exist: " + key);
            }

            if (studentToCreate[key]) {
              studentToCreate[key].end = eventEnd;
              previousAllDate = eventEnd;
            } else {
              Logger.log("Key doesn't exist: " + key);
            }
          } else {
            key = "allday;" + curDateValue + ";" + eventName;
            if (shouldWePutOnGcal) {
              toCreate[key] = {name: eventName, start: curDateValue, end: eventEnd, 
                              options: {location: eventLoc, 
                                        description: descriptionText}, date: curDateValue, allDay: true};
            }
            if (shouldWePutOnStudentCal) {
              studentToCreate[key] = {name: eventName, start: curDateValue, end: eventEnd, 
                             options: {location: eventLoc, 
                                       description: descriptionText}, date: curDateValue, allDay: true};
            }
            previousEventName = eventName;
            previousAllDate = eventEnd;
            previousKey = key;
          }
        }
        index++;
        continue;
      }

      if (eventEnd == "" || !isValidDate(eventEnd)) {
        eventEnd = new Date(eventStart);
        missingEnd = true;
      }
      // sometimes these have issues with timezones. You can use getUTCHours() and getUTCMinutes() if you're running into timezone issues
      let isEasternTime = eventStart.getHours() >= 0 && eventStart.getHours() <= 3; // timezone issues with EST..
      eventStart = new Date(curDateValue.getUTCFullYear(), curDateValue.getUTCMonth(), isEasternTime ? curDateValue.getUTCDate() + 1 : curDateValue.getUTCDate(), eventStart.getHours(), eventStart.getMinutes());
      eventEnd = new Date(curDateValue.getUTCFullYear(), curDateValue.getUTCMonth(), isEasternTime ?curDateValue.getUTCDate() + 1 : curDateValue.getUTCDate(), eventEnd.getHours(), eventEnd.getMinutes());
      if (missingEnd) {
        eventEnd.setTime(eventStart.getTime() + 60*60*1000); // add an hour if no end time is set
        missingEnd = false;
      }
      if (eventEnd.getTime() < eventStart.getTime()) { //hack fix for when the end time is 12AM or something. Add 24 hours to move it to next day.
        eventEnd.setTime(eventEnd.getTime() + 1000*60*60*24);
      }
      eventLoc = allCellsInWUD[index][4];
      if (allFormulas[index][4]) {
        // Zoom link or online link provided
        let formula = allFormulas[index][4];
        let linkRegex = "=HYPERLINK\\(\"(.*)\",\"(.*)\"\\)";
        let link = formula.match(linkRegex)[1];
        eventLoc = link
      }
      descriptionText = "";
      studentDescriptionText = ""
      if (hasExtraColumn1) {
        descriptionText += "<b>" + descriptionText1 + "</b>" + allCellsInWUD[index][5];
      }
      if (hasExtraColumn2) {
        descriptionText += "\n\n<b>" + descriptionText2 + "</b>" + allCellsInWUD[index][6];
      }
      if (hasExtraColumn3) {
        descriptionText += "\n\n<b>" + descriptionText3 + "</b>" + allCellsInWUD[index][7];
        studentDescriptionText += "<b>" + descriptionText3 + "</b>" + allCellsInWUD[index][7];
      }
      if (allFormulas[index][4]) {
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
    }
    index++;
  }

  while (firstDateValue.getTime() <= curDateValue.getTime()) {
    events = myCalendar.getEventsForDay(firstDateValue);
    studentEvents = studentCalendar.getEventsForDay(firstDateValue);
    for(let e in events){
      matching = null;
      if (events[e].isAllDayEvent())
        key = "allday;" + events[e].getAllDayStartDate().getDate() + ";" + events[e].getTitle();
      else
        key = "" + events[e].getStartTime().getTime() + ";" + events[e].getEndTime().getTime() + ";" + events[e].getTitle();
      if (!(key in toCreate) && (events[e].getStartTime().valueOf() >= firstDateValue.getTime().valueOf() || events[e].isAllDayEvent())) {
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

    for(let e in studentEvents){
      matching = null;
      if (studentEvents[e].isAllDayEvent())
        key = "allday;" + studentEvents[e].getAllDayStartDate().getDate() + ";" + studentEvents[e].getTitle();
      else
        key = "" + studentEvents[e].getStartTime().getTime() + ";" + studentEvents[e].getEndTime().getTime() + ";" + studentEvents[e].getTitle();
      if (!(key in studentToCreate) && (studentEvents[e].getStartTime().valueOf() >= firstDateValue.getTime().valueOf() || studentEvents[e].isAllDayEvent())) {
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
    firstDateValue.setDate(firstDateValue.getUTCDate() + 1);
  }
  
  let counter = 0;
  for (let e in toCreate) {
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

  for (let e in studentToCreate) {
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
  for(let key in toDelete){
    toDelete[key].deleteEvent();
    counter++;
    if (counter % 20 == 0)
      Utilities.sleep(1000);
  }

  for (let key in toStudentDelete) {
    toStudentDelete[key].deleteEvent();
    counter++;
    if (counter % 20 == 0)
      Utilities.sleep(1000);
  }
}

// for debugging purposes. Delete all events in the last 8 days and 20 days from now
function clearAllEvents() {
  let myCalendar = CalendarApp.getCalendarsByName("test_aym_calendar")[0];
  let now = new Date();
  events = myCalendar.getEvents(new Date(now.getTime() - 8*24*60*60*1000), new Date(now.getTime() + 20*24*60*60*1000));
  for(let e in events){
    events[e].deleteEvent();
  }
}

