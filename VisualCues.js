function getDates() {
    getSpreadsheets()
    const allDates = source1.getRange(4, 3, source1.getLastRow()-6, 1).getValues()
    return allDates
  }
  
  function highlightDuplicateDates() {
    getSpreadsheets()
  
    // Build an array (formattedDates) that has the dates of each event listed in a comparable format to find duplicate days without factoring in time comparison
    let formattedDates = []
  
    const allDates = getDates()
  
    allDates.forEach(array => {
      const month = array[0].getMonth()
      const day = array[0].getDate()
      let date = `${month}-${day}`
      formattedDates.push(date)
    })
    const duplicates = (array) => {
      const uniqueDates = new Set(array)
      const filteredDates = array.filter((item, index) => {
        if (uniqueDates.has(item)) {
          uniqueDates.delete(item)
        } else {
          return item
        }
      })
      return filteredDates
    }
  
    const duplicateArray = duplicates(formattedDates)
    console.log(duplicateArray)
  
    const highlight = SpreadsheetApp.newTextStyle()
      .setForegroundColor("#ffff00")
      .setBold(true)
      .build()
  
    formattedDates.forEach((date, index) => {
      if (duplicateArray.includes(date)) {
        const range = source1.getRange(index + 4, 3)
        range.setTextStyle(highlight).setBackgroundColor("#073763")
      }
    })
    console.log("highlight complete")
    showAlert("Filter Events Complete", false)
   
  }
  
  function highlightHighHours() {
    getSpreadsheets()
    const highlight = SpreadsheetApp.newTextStyle()
      .setForegroundColor("#ffff00")
      .setBold(true)
      .build()
  
    const hours = source1.getRange(4, 8, source1.getLastRow() - 6, 1).getValues()
    hours.forEach((hour, index) => {
      if (hour[0] > 12) {
        const range = source1.getRange(index + 4, 8)
        range.setTextStyle(highlight).setBackgroundColor("#073763")
      }
    })
  }
  
  function checkDateGaps() {
    getSpreadsheets()
    const allDates = getDates()
    const formattedDates = allDates.map(date => {
      const formattedDate = new Date(date[0])
      return formattedDate
    })
    for (i=0; i<formattedDates.length-1; i++) {
      const msDates = formattedDates.map(date => date.getTime())
      const difference = Math.ceil((msDates[i + 1] - msDates[i]) / (1000*60*60*24))
      const month = formattedDates[i].getMonth()
      console.log(`${month+1}`, difference)
      if (month > 2 && month < 10) {
        if (difference > 20) {
          console.log("Border set")
          const cell = source1.getRange(i+4, 3, 2, 1)
        cell.setBorder(true, true, true, true, false, false, "red", SpreadsheetApp.BorderStyle.SOLID_THICK)
        }
      } 
      
    }
  }
  
  
  // The analyzeDescriptions function below is to allow for special processes based on the descriptions given.  Still under Development
  
  // function analyzeDescriptions() {
  //   getSpreadsheets()
  
  //   const descriptions = source1.getRange(4, 7, source1.getLastRow()-6, 1).getValues()
  
  //   descriptions.forEach((array, index) => {
  //     if (array[0].includes("Mulch")) {
  //       console.log(`row ${index + 4} has Mulch in it`)
  //     }
  //   })
  
  // }