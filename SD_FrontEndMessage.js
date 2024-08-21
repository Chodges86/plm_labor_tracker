function showDialog(htmlFile, htmlWidth, htmlHeight, htmlDialog) {

    const html = HtmlService.createHtmlOutputFromFile(htmlFile)
      .setWidth(htmlWidth)
      .setHeight(htmlHeight);
    SpreadsheetApp.getUi()
      .showModalDialog(html, htmlDialog);
  }
  
  function showAlert(message, withOkayCancel) {
    const ui = SpreadsheetApp.getUi()
    if (withOkayCancel) {
      ui.alert(message, ui.ButtonSet.OK_CANCEL)
    } else {
      ui.alert(message)
    }
  
  }
  
  function showContractAddForm() {
    const html = HtmlService.createHtmlOutputFromFile("ContractForm.html")
    .setWidth(500)
    .setHeight(500)
  SpreadsheetApp.getUi()
  .showModalDialog(html, "Enter New Contract Info")
  }