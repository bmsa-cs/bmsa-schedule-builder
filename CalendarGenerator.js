/*
  Name:       BMSA Calendar Generator
  Author:     Alex Wolfe
              "What Color Is It?" by Lisa Berry
  Date:       11/26/2017
  
  License:    Released under GNU GPLv3 (https://www.gnu.org/licenses/gpl-3.0.en.html)
  
  Description: Uses spreadsheet data to build a (custom) calendar for the BMSA Upper Academy SY 2017-2018.
  Features:
    -Generate a personalized (class names and locations) or general calendar for the school year.
    -Delete your calendar from Google Sheets.
    -Colors the events based on the color of the days.
    -Only updates days that don't match the color/number of events.
    -Supports the addition of more types of schedules as long as a sheet exists in the same format as the others.

  Version History:
    1.84 
    Restored writing a super short description to the calendar.
    
    1.83
    Forces deletion of a day if it's red2018 or purple2018.
    
    1.82
    Removed personalization saving due to Google severely limiting how much info can go in the calendar description.
    
    1.81
    Forced regeneration of all dates due to changes in the schedule.
    
    1.8
    Added color support for 2018 revisions of blue, yellow, red, and purple.
    
    1.7
    Added support for special Accelerated Term customization.
    
    1.6
    Disables using regular class titles for "Accelerated Term."
    
    1.512
    Minor fixes to custom naming.
    
    1.511
    Removed forceAllDates and added an option to generate without reminder alerts.
    
    1.51
    Reverted current month limit.
  
    1.5
    Alphabetized most functions.
    Limited the calendar to generating the current month only to avoid timeout errors.
    Made generateCalendar() take optional arguments calendarName, and forceAllDates.

    1.42
    Added Accelerated Term color.

    1.41
    fixPurples got removed in 1.4.

    1.4
    Added a few more comments.
    Personalization info is now saved in the XML for future updates to calendar. (No more making separate copies of the document).
    Can now check personalization info and ask it to be updated from the spreadsheet.

    1.31
    Purple days need regenerated due to a schedule error.

    1.3
    Fixed an error when a calendar had an empty description or no XML.

    1.2
    Color check is now based on event color instead of the description.
    Added support for locations in personalization.
    Now only updates present and future dates.
    Separated whatColor() and description XML functions into their own files.
    Uses XML formatting in calendar description to more elegantly keep track of things like version and date updated.

    1.1
    All personalization data is loaded at the start of building the calendar, so that the sheet isn't read over and over.
    Renamed variables and cleaned up code.

    1.0
    Initial release.

*/

function onOpen() {
  //Add a menu to the spreadsheet.
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('BMSA Calendar Generator')
  .addItem('Generate/Update Calendar', 'generateCalendar')
  .addItem('Delete Schedule Calendar', 'deleteCalendar')
  .addSeparator()
  .addItem('Check saved personalizations', 'checkPersonalizations')
  .addItem('Update saved personalizations', 'updatePersonalizations')
  .addSeparator()
  .addItem("What Color Is It?", 'whatColor')
  .addToUi();
  whatColor();
}

function checkPersonalizations() {
  // DOESN'T WORK. GOOGLE SHORTENED MAX CALENDAR DESCRIPTION LENGTH.
  var ui = SpreadsheetApp.getUi();
  var CALENDAR_NAME = "BMSA Schedule";
  var colorCal = CalendarApp.getOwnedCalendarsByName(CALENDAR_NAME)[0];
  var pArray = getPersonalizationArrayfromXML(colorCal.getDescription());
  if (pArray==null) {
    ui.alert("No Personalizations", "No valid preferences saved. They will be created from the 'Personal Schedule' sheet next time you update the calendar.", ui.ButtonSet.OK);
  } else {
    var pOutput = "";
    for (i=0; i<pArray.length; i++) {
      for (j=0; j<pArray[i].length; j++) {
        if (pArray[i][j] == "") {pArray[i][j]="(blank)";}
      }
      pOutput = pOutput + "Core: " + pArray[i][0] + " - Class: " + pArray[i][1] + " - Location: " + pArray[i][2] +"\n";
    }
    ui.alert("Saved Personalizations: ", pOutput, ui.ButtonSet.OK);
  }
}

function updatePersonalizations() {
  // DOESN'T WORK. GOOGLE SHORTENED MAX CALENDAR DESCRIPTION LENGTH.
  var ui = SpreadsheetApp.getUi();
  var CALENDAR_NAME = "BMSA Schedule";
  var VERSION = getVersion();
  var TODAY = new Date();
  var NICE_TODAY = (TODAY.getMonth()+1)+"/"+TODAY.getDate()+"/"+TODAY.getYear();
  var colorCal = CalendarApp.getOwnedCalendarsByName(CALENDAR_NAME)[0];
  var ccDesc = colorCal.getDescription();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pArray = buildPersonalizationArray(ss);
  Logger.log("Writing updated personalizations to calendar.");
  writeCalendarDescription(colorCal, VERSION, NICE_TODAY, pArray);
  ui.alert("Personalizations Updated","All done. Use 'Check saved personalizations' to verify. \nPLEASE NOTE: that this will only affect days that haven't already been added to the calendar.", ui.ButtonSet.OK);

}

function deleteCalendar() {
  var ui = SpreadsheetApp.getUi();
  var CALENDAR_NAME = "BMSA Schedule";
  var colorCal = CalendarApp.getOwnedCalendarsByName(CALENDAR_NAME)[0];

  var response = ui.alert("Delete Calendar", "Are you sure you want to DELETE 'BMSA Schedule?'", ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {colorCal.deleteCalendar();}
}

function getVersion() {
  var VERSION = "1.83";
  return VERSION;
}

function generateCalendar(calendarName, noReminders) {
  if (arguments.length==2) {
    var CALENDAR_NAME = calendarName;
    var noReminders = noReminders;
  } else if (arguments.length==1) {
    var CALENDAR_NAME = calendarName;
    var noReminders = false;
  } else {
    var CALENDAR_NAME = "BMSA Schedule";
    var noReminders = false;
  }
    
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var VERSION = getVersion();
  var TODAY = new Date();
  var NICE_TODAY = (TODAY.getMonth()+1)+"/"+TODAY.getDate()+"/"+TODAY.getYear();
  var NOT_SCHEDULES = ["Calculating", "Date Not Found", "Days Til School", "PD", "No School", "Dates", "Past Dates", "Personal Schedule"]; //Sheets that aren't schedules saves time by not reading them in buildScheduleArray.
  var MILLISECS_IN_DAY = 24 * 60 * 60 * 1000;
  var colorCal = CalendarApp.getOwnedCalendarsByName(CALENDAR_NAME)[0];
  var dateArray = new Array();
  var scheduleArray = new Array();
  var personalizationArray = new Array();

  if (colorCal==null) {
    Logger.log("Calendar doesn't exist yet. Creating "+CALENDAR_NAME);
    colorCal = CalendarApp.createCalendar(CALENDAR_NAME);
    colorCal.setSelected(true);
    colorCal.setColor("#29527A"); //SEA_BLUE
    colorCal.setTimeZone("America/New_York");
    personalizationArray = buildPersonalizationArray(ss);
    writeCalendarDescription(colorCal, VERSION, 0, personalizationArray); //Setup the initial XML description.
  } else if (getDateUpdated(colorCal.getDescription()) == NICE_TODAY) {
    var response = ui.alert("Calendar Already Updated", 'Calendar was already updated today.\n Are you sure you want to continue?', ui.ButtonSet.YES_NO);

    if (response == ui.Button.YES) {
      Logger.log("Continuing despite being updated.");
    } else {
      throw("Calendar already updated today. Stopping.");
    }
  }

  buildDateArray(ss);
  buildScheduleArray(ss);
  personalizationArray = buildPersonalizationArray(ss, colorCal);
  if (getVersionUsed(colorCal.getDescription()) < "1.41" && getVersionUsed(colorCal.getDescription()) != "1.31") {
    Logger.log("Version was less than 1.41, purple days need regenerated.");
    dateArray.forEach(fixPurples);
  }

  ui.alert("Patience","There may be a lot of events to create. If the script might gives an error about 'maximum execution time exceeded,' don't panic. Just click on 'Generate/Update Calendar' again and it'll eventually finish.", ui.ButtonSet.OK);

  scheduleArray.forEach(getAndPushSchedules);
  dateArray.forEach(createCalendarForDay);

  Logger.log("All done. Updating description.");
  writeCalendarDescription(colorCal, VERSION, NICE_TODAY, personalizationArray);

  function buildDateArray(spreadsheet) {
    var dateSheet = spreadsheet.getSheetByName("Dates");
    var maxRows = dateSheet.getLastRow();
    /* read dates and colors (B and C) into array */
    dateArray = dateSheet.getRange(1,2, maxRows, 2).getValues();
    //TO-DO: Only add current and future dates to array.
  }

  function buildScheduleArray(spreadsheet) {
    //Builds scheduleArray from the current Spreadsheet
    //Get all sheets and find the named schedules in it.
    //Exclusions come from NOT_SCHEDULES.
    var allSheets = spreadsheet.getSheets();
    for (i=0; i<allSheets.length; i++) {
      var tempName = allSheets[i].getName()
      if (NOT_SCHEDULES.indexOf(tempName) == -1) {
        scheduleArray.push([tempName])
      }
    }
  }

  function createCalendarForDay(item) {
    //Creates the calendar events for the given day and color.
    var tempDate = item[0];
    var tempColor = item[1];
    var tempEvents = colorCal.getEventsForDay(tempDate);
    var colorIndex = -1;

    var eventColor = getClosestColor(tempColor);

    if (tempDate.getTime() >= TODAY.getTime()) {
      //If our date is in the future OR is the exact same date, build the calendar.
      //Find the index of the color in scheduleArray.
      for (i=0; i<scheduleArray.length; i++ ) {
        if (scheduleArray[i][0] == tempColor) {
          colorIndex = i;
          break;
        }
      }
      if (colorIndex == -1) {
        Logger.log(tempDate+" does NOT have a schedule. Type: " + tempColor);
        if (tempEvents.length == 1) {
          Logger.log("One event for today already detected. Skipping.");
        } else {
          Logger.log("Does not match expected number of events.");
          deleteAllEventsForDay(colorCal, tempDate);
          var newEvent = colorCal.createAllDayEvent(tempColor, tempDate);
          if (eventColor!=null) {newEvent.setColor(eventColor);}

        }
      } else {
        Logger.log(tempDate+" does have a schedule. Type: " + tempColor);
        if (tempEvents.length == scheduleArray[colorIndex][1].length && getClosestColor(tempColor) == tempEvents[0].getColor() && tempColor != "red2018" && tempColor != "purple2018") {
          //Correct number of events, correct color, and not red2018 or purple2018 (same # of events, but different order).
          Logger.log("The correct number of events exist and color already matches. Skipping.");
        } else {
          Logger.log("Mismatch in number of events or color. Deleting all of today's events and recreating.");
          deleteAllEventsForDay(colorCal, tempDate);

          for (i=0; i<scheduleArray[colorIndex][1].length; i++) {
            var eventName = scheduleArray[colorIndex][1][i][0];

            if (tempColor!="Accelerated Term") {
              var class = getPersonalTitle(eventName);
            } else {
              var class = getAccTermTitle(eventName);
            }
            
            var eventLocation = String(getLocation(eventName));
            
            if (eventName <= 7) {eventName = "Core " + eventName;}
            if (class!=null && class!="") {eventName = eventName + ": " + class;}
            
            var startTime = new Date(tempDate);
            startTime.setHours(scheduleArray[colorIndex][1][i][1].getHours());
            startTime.setMinutes(scheduleArray[colorIndex][1][i][1].getMinutes());

            var endTime = new Date(tempDate);
            endTime.setHours(scheduleArray[colorIndex][1][i][2].getHours());
            endTime.setMinutes(scheduleArray[colorIndex][1][i][2].getMinutes());

            var desc = "";

            Logger.log("Creating Event:---");
            Logger.log("Name: " + eventName);
            Logger.log("Start Time: " + startTime);
            Logger.log("End Time: " + endTime);
            Logger.log("Event Color:" + eventColor);
            Logger.log("Location: " + eventLocation);
            Logger.log("Description: " + desc);

            var newEvent = colorCal.createEvent(eventName, startTime, endTime, {description: desc, location: eventLocation});
            if (eventColor!=null) {newEvent.setColor(eventColor);}
            if (!noReminders) {
              newEvent.addPopupReminder(3);
              newEvent.addPopupReminder(0);
            }

            Utilities.sleep(1 * 1000); //Wait 1 second after creating each event.
          }
        }
      }
    } else {Logger.log("Date is not this month or is in the past. Skipping.");}
  }

  function deleteAllEventsForDay(calendar, date) {
    //Delete all events in the calendar on the given day.
    var tempEvents = calendar.getEventsForDay(date);
    for (i=0; i<tempEvents.length; i++) {
      tempEvents[i].deleteEvent();
      Utilities.sleep(1 * 1000); //Wait 1 second after deleting each event.
    }
  }

  function fixPurples(item) {
    var tempDate = item[0];
    var tempColor = item[1];
    if (tempColor == 'purple') {deleteAllEventsForDay(colorCal, tempDate);}
  }

  function getAndPushSchedules(item) {
    //Gets the schedule information for each named/colored schedule in the array and pushes it back to the item.
    tempSheet = ss.getSheetByName(item);
    item.push(tempSheet.getRange(4,2,tempSheet.getLastRow()-3,3).getValues());
  }

  function getClosestColor(color) {
    if (color == "pink") {return "4";} //PALE_RED
    else if (color == "green") {return "10";} //GREEN
    else if (color == "blue" || color == "blue2018") {return "9";} //BLUE
    else if (color == "red" || color == "red2018") {return "11";} //RED
    else if (color == "purple" || color == "purple2018") {return "3";} //MAUVE
    else if (color == "yellow" || color == "yellow2018") {return "5";} //YELLOW
    else if (color == "orange") {return "6";} //ORANGE
    else if (color == "Accelerated Term") {return "7";} //CYAN
    else return null;
  }

  function getLocation(period) {
    //Returns any location information for the schedule.
    if (!isNaN(Number(period))) {return personalizationArray[Number(period)-1][2];}
    else if (period=="Family Group") {return personalizationArray[7][2];}
    else {return ""};
  }

  function getPersonalTitle(period) {
    //Returns any personal titles the user has added to their schedule.
    if (!isNaN(Number(period))) {return personalizationArray[Number(period)-1][1];}
    else if (period=="Family Group") {return personalizationArray[7][1];}
    else {return ""};
  }
  
  function getAccTermTitle(period) {
    //Returns any Accelerated Term titles the user has added to their schedule.
    if (!isNaN(Number(period))) {return personalizationArray[8+Number(period)-1][1];}
    else {return ""};
  }

}

function buildPersonalizationArray(spreadsheet, calendar) {
  // XML DOESN'T WORK. GOOGLE SHORTENED MAX CALENDAR DESCRIPTION LENGTH.
  if (arguments.length == 2) {
    //Checks XML for personalizations, otherwise reads in the personalizations from SS and returns them.
    Logger.log("Attempting to get personalizations from XML description.");
    xmlPArray = getPersonalizationArrayfromXML(calendar.getDescription());
    if (xmlPArray != null) {
      return xmlPArray;
    } else {
      Logger.log("Failed. Building personalizations from sheet.");
      var cSheet = spreadsheet.getSheetByName("Personal Schedule");
      return cSheet.getRange(2,1,cSheet.getLastRow()-1,3).getValues();
    }
    
  } else { //only spreadsheet was given
    Logger.log("Building personalizations from sheet.");
    var cSheet = spreadsheet.getSheetByName("Personal Schedule");
    return cSheet.getRange(2,1,cSheet.getLastRow()-1,3).getValues();
  }
}
