function printEvents(events) {
    getSpreadsheets()
    showDialog("Loading", 200, 100, "Getting Events")
  
    const titleRow = source1.getRange(3, 1, 1, 13)
    const titleRowContent = ["Customer", "Loc ID", "Date", "Start", "End", "Crew", "Description", "Time Onsite", "# of Men", "MHRS", "Include in Calculations", "CalendarID", "EventID"]
    titleRow.setValues([titleRowContent])
  
    if (events === undefined) {
      events = getEvents()
    } else {
      getSpreadsheets()
    }
  
    removeCheckBoxes()
  
    //Delete the range between A1 and G1000 run to allow it to be written over
  
    source1.deleteRows(4, source1.getMaxRows() - 4)
    console.log("Clear old data complete")
  
    let detailsToPrint = []
  
    let crewName = originalId => {
      switch (originalId) {
        case "randall@plm-llc.com":
          return "Randall"
        case "caleb@plm-llc.com":
          return "Caleb"
        case "crew1@plm-llc.com":
          return "Crew 1"
        case "crew2@plm-llc.com":
          return "Crew 2"
        case "crew3@plm-llc.com":
          return "Crew 3"
        case "crew4@plm-llc.com":
          return "Crew 4"
        default:
          return originalId
  
      }
    }
  
    let hoursOnsite = (startTime, endTime) => {
      const start = new Date(startTime)
      const end = new Date(endTime)
      const startHour = start / 3600000
      const endHour = end / 3600000
      return endHour - startHour
    }
    console.log("Getting details of events Started")
    events.forEach(array => {
      array.forEach(event => {
  
        let details = [event.getTitle(), event.getLocation(), event.getEndTime(), event.getStartTime(), event.getEndTime(), crewName(event.getOriginalCalendarId()), event.getDescription(), hoursOnsite(event.getStartTime(), event.getEndTime()), "", null, null, event.getOriginalCalendarId(), event.getId()];
        detailsToPrint.push(details);
  
      })
    })
    console.log("Getting details completed")
  
    const printPoint = source1.getLastRow();
    const printPointRange = source1.getRange(printPoint + 1, 1, detailsToPrint.length, detailsToPrint[0].length);
    printPointRange.setValues(detailsToPrint);
    setFormula(detailsToPrint.length)
    insertCheckboxes(detailsToPrint.length)
  
    console.log("Details printed")
  
    sortSheet(source1, 3);
  
    // showAlert("Get Events complete", false)
    showDialog("Loading", 200, 100, "Preparing Sheet")
    //resetAltScheme(source1.getLastRow())
  
  }// END OF: printEvents function
  
  
  function formatSheet() {
    getSpreadsheets()
    // unhighlightCells()
    source1.getRange(1, 1, sheet.getMaxRows(), sheet.getLastColumn()).setHorizontalAlignment("center") // Set entire sheet with center align
    source1.getRange(3, 1, 1, source1.getLastColumn()).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP) // Set row three (column labels) to have text wrapping
  
    const textStyle = SpreadsheetApp.newTextStyle()
      .setBold(false)
      .setFontSize(10)
      .setForegroundColor("black")
      .build()
  
    source1.getRange(4, 1, source1.getLastRow(), source1.getLastColumn())
      .setTextStyle(textStyle)
    source1.getRange(4, 10, source1.getLastRow(), source1.getLastColumn()).setNumberFormat("0.00")
  
  }
  
  
  function sortSheet(sheet, columnPosition) {
    sheet.sort(columnPosition, true)
    console.log("Events sorted complete")
  } // END OF: sortSheet
  
  
  function setFormula(qty) {
    getSpreadsheets()
    let row = 4
    for (i = 0; i < qty; i++) {
      let formula = String(`=I${row}*H${row}`)
      const rowToPrint = source1.getRange(row, 10)
      rowToPrint.setValue(formula)
      row += 1
  
    }
  }// END OF: setFormula function
  
  
  function addCalculationFormulas() {
    getSpreadsheets()
    const lastRow = source1.getLastRow()
    source1.getRange(lastRow, 1, 1, 11).setBorder(false, false, true, false, false, true, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  
    const count = lastRow - 3
    const totalFormula = `=sum(J4:J${lastRow})`
    const avgFormula = `=sum(J4:J${lastRow})/${count}`
    const countFormula = `${count}`
  
    const textStyle = SpreadsheetApp.newTextStyle()
      .setBold(true)
      .setFontSize(15)
      .build()
  
    source1.getRange(lastRow + 1, 11, 6, 1)
      .setValues([["Total MHRS"], ["AVG MHRS"], ["Number of trips"], [""], [""], [""]])
      .clearDataValidations()
      .setTextStyle(textStyle)
  
    source1.getRange(lastRow + 1, 10, 3, 1)
      .setValues([[totalFormula], [avgFormula], [countFormula]])
      .setNumberFormat("0.00")
      .setTextStyle(textStyle)
  
    highlightDuplicateDates()
    highlightHighHours()
    checkDateGaps()
  }
  
  function removeCalculationFormulas() {
    getSpreadsheets()
    let maxRow = source1.getMaxRows()
    let lastCol = source1.getLastColumn()
    source1.getRange(4, 1, maxRow, lastCol).setBorder(false, false, false, false, false, false)
    addCalculationFormulas()
  }
  
  function updateCalculationFormulas() {
    getSpreadsheets()
  
    // Get values of all the cells in the J column (mhrs values) in two dimensional array
    const mhrsValues = source1.getRange(4, 10, source1.getLastRow() - 6, 1).getValues()
  
    // Get values of all the cells in the K column (checkboxes) in two dimensional array
    const checkMarkValues = source1.getRange(4, 11, source1.getLastRow() - 6, 1).getValues()
  
    // compile an array of all rows containing a check mark
    let rowsWithChecks = []
    checkMarkValues.forEach((array, index) => {
      if (array[0] === true) {
        rowsWithChecks.push(index + 4)
      }
    })
  
    // Loop through each row on the J col that is included in the "checked" array and add it to Total Mhrs cell
    let total = 0
    mhrsValues.forEach((array, index) => {
      if (rowsWithChecks.includes(index + 4)) {
        total += array[0]
      }
    })
  
    // Calculate average from total figured with "checked" array. Set AVG MHRS value
    const newAverage = total / rowsWithChecks.length
  
    source1.getRange(source1.getLastRow() - 2, 10).setValue(total)
    source1.getRange(source1.getLastRow(), 10).setValue(rowsWithChecks.length)
    source1.getRange(source1.getLastRow() - 1, 10).setValue(newAverage)
  }
  
  function insertCheckboxes(qty) {
    getSpreadsheets()
    const range = source1.getRange(4, 11, qty, 1)
    range.insertCheckboxes().check()
  
    const textStyle = SpreadsheetApp.newTextStyle()
      .setBold(false)
      .setFontSize(10)
      .setForegroundColorObject(SpreadsheetApp.newColor().setRgbColor("#073763").build())
      .build()
  
    range.setTextStyle(textStyle)
  }
  
  function removeCheckBoxes() {
    getSpreadsheets()
    SpreadsheetApp.getActiveSheet().getRange("K4:K").removeCheckboxes();
  }
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  