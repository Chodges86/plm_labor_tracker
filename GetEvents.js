function getEvents() {
  getSpreadsheets();

  const getCalendars = () => CalendarApp.getAllOwnedCalendars();
  const allCalendars = getCalendars(); // Get all owned calendars
  const unneededCalendars = ["Transferred from jason@plm-llc.com"]; // An array containing any calendars that will not be needed. ADD UN-NEEDED CALENDARS HERE
  for (i in allCalendars) {
    if (unneededCalendars.includes(allCalendars[i].getName())) {
      allCalendars.splice(i, 1);
    }
  }

  let yearToRetrieve = source1.getRange(2, 4).getValue();
  if (yearToRetrieve === "Enter Year (yyyy)") {
    const prompt = SpreadsheetApp.getUi().prompt(
      "Please enter the year of events you want to retrieve.\n Year format must be yyyy"
    );
    yearToRetrieve = prompt.getResponseText();
    source1.getRange(2, 4).setValue(yearToRetrieve);
  }

  let events = [];
  const startDate = new Date(`January 1, ${yearToRetrieve} 00:00:00 CST`);
  const endDate = new Date(`December 31, ${yearToRetrieve} 23:59:59 CST`);
  allCalendars.forEach((calendar) =>
    events.push(calendar.getEvents(startDate, endDate))
  );
  console.log("getEvents Complete");
  return events;
} // END OF: getEvents function
