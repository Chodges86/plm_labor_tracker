function didEdit(e) {
    getSpreadsheets()
    const range = e.range
    const row = range.getRow()
    const col = range.getColumn()
    const currentSheet = e.source.getActiveSheet().getIndex()
  
    if (currentSheet === 2) {
      switch (col) {
        case 2: // Edited location of event.  Send info to calendar
          const startTime = source1.getRange(row, 4).getValue()
          const endTime = source1.getRange(row, 5).getValue()
          const calID = source1.getRange(row, 12).getValue()
          const eventID = source1.getRange(row, 13).getValue()
          const locationNumber = e.range.getValue()
  
          let ui = SpreadsheetApp.getUi()
          let result = ui.alert(`Would you like to update the calendar event for ${source1.getRange(row, 1).getValue()} with location ID of: ${locationNumber}`, ui.ButtonSet.OK_CANCEL)
          if (result === ui.Button.OK) {
            updateLocationInCalendar(startTime, endTime, calID, eventID, locationNumber)
          } else {
            ui.alert("Update calendar event location canceled")
            range.setValue("")
          }
          break
        case 3: // Edited Date. Set it back bc that is in error. Needed special case to format cell
          range.setValue(e.oldValue)
          range.setNumberFormat("M/dd/yyyy")
          break
        case 4: // Edited Start Time. Set it back bc that is in error. Needed special case to format cell
  
          if (range.getRow() !== 2) {
            range.setValue(e.oldValue)
            range.setNumberFormat('h:mm:ss A/P"M"')
          } else {
            const textStyle = SpreadsheetApp.newTextStyle()
              .setBold(false)
              .setFontSize(12)
              .setForegroundColorObject(SpreadsheetApp.newColor().setRgbColor("white").build())
              .build()
            source1.getRange(2, 3).setValue("Edit year:").setTextStyle(textStyle)
          }
  
          break
        case 5: // Edited End Time. Set it back bc that is in error. Needed special case to format cell
          range.setValue(e.oldValue)
          range.setNumberFormat('h:mm:ss A/P"M"')
          break
        case 9: // Edit on Number of Men. Do nothing except update final calculations as this is needed for calculations
          updateCalculationFormulas()
          break
        case 10: // Edited MHRS formula. Set it back bc that is in error.  resetFormula function to determine which formula was changed
          range.setValue(resetFormula(row, col))
          break
        case 11: // Changed status of checkbox.  Update Total/ AVG/ Number of trips calculated
  
          updateCalculationFormulas()
          break
        default:
          range.setValue(e.oldValue)
      }
    }
  
    if (currentSheet === 1) {
      organizeSheet(e)
    }
  
  } // End didEdit function
  
  function updateLocationInCalendar(startTime, endTime, calID, eventId, locationNumber) {
    const cal = CalendarApp.getCalendarById(calID)
    const events = cal.getEvents(startTime, endTime)
    if (events.length > 1) {
      events.forEach(event => {
        if (event.getId() === eventId) {
          event.setLocation(locationNumber)
        }
      })
    } else if (events.length === 1) {
      events[0].setLocation(locationNumber)
    } else {
      // This means that no event was returned.  An ideal place for an error message alert
      console.log("Could not update location because no event was returned")
      SpreadsheetApp.getUi().alert("Could not update location because no event was returned. Please check calendar to set date manually")
    }
  }
  
  function resetFormula(row, col) {
    const value = source1.getRange(row, col + 1, 1, 1).getValue()
    console.log(value)
    let formula;
  
    switch (value) {
      case "Total MHRS":
        formula = `=sum(J4:J${row - 2})`
        return formula
        break
      case "AVG MHRS":
        const numberOfEvents = row - 3
        formula = `=sum(J4:J${row - 3})/${numberOfEvents}`
        return formula
        break
      case "Number of trips":
        return row - 7
        break
      default:
        formula = `=I${row}*H${row}`
        return formula
        break
    }
  }
  
  function testGetEventbyId() {
    const cal = CalendarApp.getCalendarById("crew3@plm-llc.com")
    const event = cal.getEvents(new Date("4/2/2022 0:00:00"), new Date("4/3/2022 0:00:00"))
  
    // console.log("Copied", "796347d8324f4ef3a1d23948850c752a")
    console.log("Retrieved", event[0].getId())
    // console.log("Retrieved Title", event[0].getTitle())
    // event[0].setLocation("174")
  
    // console.log("IS SAME", "796347d8324f4ef3a1d23948850c752a" === event[0].getId())
  
    // const eventFromID = cal.getEventById("796347d8324f4ef3a1d23948850c752a")
    // eventFromID.setTitle("NP Reading garden")
  
    // console.log(eventFromID)
  }
  
  
  
  
  
  
  
  
  
  
  