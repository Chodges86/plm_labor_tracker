function organizeSheet(e) { // This function adjusts the layout for the Contract Data sheet
    getSpreadsheets()
    const sheetRange = source0.getRange(1, 1, source0.getMaxRows(), source0.getMaxColumns())
    
    const a2Value = source0.getRange("A2").getValue()
    const c2Value = source0.getRange("C2").getValue()
    const allRecords = source0.getRange(3, 1, source0.getLastRow(), source0.getLastColumn()).getValues()
  
    
  
  
    const isAFilterEdit = () => {
      if (e.range.getRow() === 2 && e.range.getColumn() === 1 || e.range.getRow() === 2 && e.range.getColumn() === 3) {
        return true
      } else {
        return false
      }
    }
    
    if (isAFilterEdit()) {
      source0.insertRowsAfter(source0.getMaxRows(), 5)
      source0.unhideRow(sheetRange)
      allRecords.forEach((row, index) => {
        // console.log(row[0], row[2], row[0] === a2Value, row[2] === c2Value)
        const rowRange = source0.getRange(index + 3, 1)
  
        if (a2Value === "All" && c2Value === "All") {
          source0.unhideRow(sheetRange)
        } else if (a2Value === "All" && c2Value != "All") {
          if (c2Value != row[2]) {
            source0.hideRow(rowRange)
          }
        } else if (c2Value === "All" && a2Value != "All") {
          if (a2Value != row[0]) {
            source0.hideRow(rowRange)
          }
        } else if (row[0] != a2Value || row[2] != c2Value) {
          source0.hideRow(rowRange)
        }
      })
        source0.deleteRows(source0.getMaxRows()-4, 5)
  
    }
  
    if (e.range.getColumn() === 1 && e.range.getRow() !== 2 && a2Value !== "All") {
      const editValue = e.range.getValue()
      if (editValue != a2Value) {
        const range = source0.getRange(e.range.getRow(), e.range.getColumn())
        source0.hideRow(range)
      }
    }
  
     if (e.range.getColumn() === 3 && e.range.getRow() !== 2 && c2Value !== "All") {
      const editValue = e.range.getValue()
      if (editValue != c2Value) {
        const range = source0.getRange(e.range.getRow(), e.range.getColumn())
        source0.hideRow(range)
      }
    }
  
  }
  
  function addNewContract(customer) {
    const { name, loc, date, status, type } = customer
  
    getSpreadsheets()
  
    console.log(name, loc, date, status, type)
  
    source0.insertRowsAfter(source0.getLastRow(), 1)
    
    source0.getRange(source0.getLastRow()+1, 1, 1, 15).setValues([["Active", name, type, loc, "", "", "", "", "", "", "", "", "", date, status]])
  
  }
  
  function getNextLocation() {
    console.log("Next loc ran")
    return "450"
  }
  
  
  
  
  
  
  
  
  