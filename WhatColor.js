//"What Color Is It?" by Lisa Berry

function whatColor() {


  var TODAY = new Date();
  var tempDate;
  var startDate;
  var row = -1;
  var daysTilSchool = -1;
  var colorSheetName;

  var MILLISECS_IN_DAY = 24 * 60 * 60 * 1000;

  /* Display "Calculating..." while we work */

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName("Calculating"));
  SpreadsheetApp.flush();


  /* set up spreadsheet handle and other variables */
  var calendarSheet = ss.getSheetByName("Dates");
  var maxRows = calendarSheet.getLastRow();

  /* read dates (column B) into array */
  var dateArray = new Array(maxRows);
  var range = calendarSheet.getRange(1,2, maxRows, 1);
  dateArray = range.getValues();


  /* make sure school has actually started */

  startDate = dateArray[0][0];
  if (startDate.valueOf() > TODAY.valueOf()) {
      colorSheetName = "Days Til School";
      daysTilSchool = 1 + Math.floor(((startDate - TODAY) / MILLISECS_IN_DAY))
    }
  else {

    /* search for TODAY's date */

    var i=-1;
    for(i=0; i<=maxRows-1; i++) {
      tempDate = dateArray[i][0];
      if (TODAY.getMonth() == tempDate.getMonth() && TODAY.getDate() == tempDate.getDate()) {
        row = i+1;      /* array index starts at zero, rows start at one */
        break;
      }
    }

    /* get appropriate sheet name */

    if (row == -1) {
      colorSheetName = "Date Not Found";
    }
    else {
      range = calendarSheet.getRange(row,3);
      colorSheetName = range.getValue();
    }
  }

  /* Display whatever is stored in "colorSheetName" */

  var colorSheet = ss.getSheetByName(colorSheetName);
  ss.setActiveSheet(colorSheet);
  if (colorSheetName == "Days Til School") {
    range = colorSheet.getRange(6,3);
    range.setValue(daysTilSchool);
  }

}
