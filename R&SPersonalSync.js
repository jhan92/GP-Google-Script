function onOpen() {
    var ui = SpreadsheetApp.getUi();
    var mainMenu = ui.createMenu("R&S Functions");
    mainMenu.addItem("Sync to Ray's Calendar", "syncToCalendar");
    mainMenu.addToUi();  
};

DAYS_OF_WEEK = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];

function getPreviousMonday(date) {
    const day = date.getDay();
    const prevMonday = new Date();
    if(date.getDay() == 0){
        prevMonday.setDate(date.getDate() - 7);
    }
    else{
        prevMonday.setDate(date.getDate() - (day-1));
    }

    return prevMonday;
}

function getTimeRange(timeStr, currDate) {
  const timeRegex = /(\d{1,2}:?\d{0,2})(pm|p|am|a)?-(\d{1,2}:?\d{0,2})(pm|p|am|a)?/
  const matched = timeStr.match(timeRegex);
  if (matched) {
    const startTime = matched[1];
    let isStartAm = true;
    const endTime = matched[3];
    const endAMPM = matched[4];
    let isEndAM = true;

    if (endAMPM == 'p' || endAMPM == 'pm') {
      isEndAM = false;
      isStartAm = false;

      if (endTime == 12) {
        isStartAm = true;
        isEndAM = true;
      }

      if (startTime == 12) {
        isStartAm = true;
      }
    }

    const startTimeHoursMin = startTime.split(':');
    const endTimeHoursMin = endTime.split(':');
    if (!isEndAM) {
      endTimeHoursMin[0] = Number(endTimeHoursMin[0]) + 12;
    }

    if (!isStartAm) {
      startTimeHoursMin[0] = Number(startTimeHoursMin[0]) + 12;
    }

    const startDate = new Date(currDate);
    startDate.setHours(0, 0, 0, 0)
    startDate.setHours(startTimeHoursMin[0])
    if (startTimeHoursMin.length == 2) {
      startDate.setMinutes(startTimeHoursMin[1]);
    }

    const endDate = new Date(currDate)
    endDate.setHours(0, 0, 0, 0)
    endDate.setHours(endTimeHoursMin[0]);
    if (endTimeHoursMin.length == 2) {
      endDate.setMinutes(endTimeHoursMin[1]);
    }
    return {
      startDate, endDate
    };
  }
}

function findValueOfMerged(allMerged, row, col) {
  for (mergedRange of allMerged) {
    const lastRow = mergedRange.getLastRow();
    const firstRow = mergedRange.getRow();
    const lastColumn = mergedRange.getLastColumn();
    const firstColumn = mergedRange.getColumn();
    if (row >= firstRow && row <= lastRow && col >= firstColumn && col <= lastColumn) {
      return mergedRange.getValue();
    }
  }

  return '';
}

function syncToCalendar() { 
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Temp"); // Put the tab name here
  let myCalendar = CalendarApp.getOwnedCalendarsByName('RaySieun Couples Calendar')[0]; // Put the Calendar name here
  if( sheet == null || myCalendar == null) {return;}
  let firstColumn = "A";
  let lastColumn = "K";

  let columnRange = firstColumn + ":" + lastColumn;
  let WUDRange = sheet.getRange(columnRange);
  let allCells = WUDRange.getValues();
  let allMerged = WUDRange.getMergedRanges();
  let allFormulas = WUDRange.getRichTextValues();
  allFormulas = allFormulas.map(row => {
    return row.map(val => {
      return val.getLinkUrl();
    });
  });
  const toCreate = {}
  let firstDateValue;
  
  let numRows = Math.min(sheet.getLastRow() + 1, 40);

  let currDate = getPreviousMonday(new Date());
  firstDateValue = currDate;
  for (let i = 1; i <= 5; i++) {
    let dayCol = i*2 - 1;
    const toCreateEvents = {}
    currDate.setHours(0, 0, 0, 0);

    let index = 3;
    let previousEventName = '';
    let previousTimeRanges;
    let previousFormula;
    while (index < numRows) {
      const currTimeRange = allCells[index][0];
      const timeRanges = getTimeRange(currTimeRange, currDate);
      let eventName = allCells[index][dayCol];
      let formula = allFormulas[index][dayCol];
      if (!eventName) {
        eventName = findValueOfMerged(allMerged, index + 1, dayCol + 1);
      }

      if (previousEventName !== eventName) {

        if (previousEventName && previousTimeRanges) {
          const eventKey = `${previousTimeRanges.startDate.getHours()}:${previousTimeRanges.startDate.getMinutes()}-${previousTimeRanges.endDate.getHours()}:${previousTimeRanges.endDate.getMinutes()};${previousEventName}`;
          toCreateEvents[eventKey] = {
            name: previousEventName,
            start: previousTimeRanges.startDate,
            end: previousTimeRanges.endDate, 
            date: currDate,
            options: {
              description: previousFormula
            },
            allDay: false
          };
        }

        previousEventName = eventName;
        previousTimeRanges = timeRanges;
        previousFormula = formula;
      } else if (previousEventName == eventName && eventName) {
        if (!previousTimeRanges) {
          previousTimeRanges = timeRanges;
        }

        previousTimeRanges.endDate = timeRanges.endDate;
      }

      index++;
    }



    toCreate[DAYS_OF_WEEK[i-1]] = toCreateEvents;
    currDate = new Date(currDate);
    currDate.setDate(currDate.getDate() + 1);
  }

  let i = 0;
  const toDelete = {};
  while (firstDateValue.getTime() <= currDate.getTime()) {
    let dayOfWeek = DAYS_OF_WEEK[i]
    let deletePerDay = {}
    toDelete[dayOfWeek] = deletePerDay;
    const events = myCalendar.getEventsForDay(firstDateValue);
    for (let e of events) {
      const key = `${e.getStartTime().getHours()}:${e.getStartTime().getMinutes()}-${e.getEndTime().getHours()}:${e.getEndTime().getMinutes()};${e.getTitle()}`
      if (!(key in toCreate[dayOfWeek])) {
        e.deleteEvent();
      } else {
        delete toCreate[dayOfWeek][key]
      }
      
    }

    i++;
    firstDateValue.setDate(firstDateValue.getDate() + 1);
  }
  
  let counter = 0;
  for (let toCreatePerDay in toCreate) {
    for (let key in toCreate[toCreatePerDay]) {
      let e= toCreate[toCreatePerDay][key]
      myCalendar.createEven
      myCalendar.createEvent(e['name'], e['start'], e['end'], e['options']);
      counter++;
      if (counter % 20 == 0) {
        Utilities.sleep(1000);
      }
    }
  }

  for (let day in toDelete) {
    for (let key in toDelete[day]) {
      let e = toDelete[day][key];
      e.deleteEvent();
    }
  }
}
