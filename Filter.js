function filter(type) {
  console.log("filter running");
  showDialog("Loading", 200, 100, "Filtering Events");
  getSpreadsheets();

  let filteredEvents = [];
  let events;

  if (source1.getRange(2, 4).getValue() === "") {
    events = getEvents();
  } else {
    events = getEvents(source1.getRange(2, 4).getValue());
  }

  switch (type) {
    case "Location":
      console.log("Filtered by location");
      const filterPrompt = SpreadsheetApp.getUi().prompt(
        "Please enter Location ID"
      );
      // Filter events and return results to printEventsDetail function
      const locationID = String(filterPrompt.getResponseText());
      showDialog("Loading", 200, 100, "Filtering Events");
      events.forEach((array) => {
        array.forEach((event) => {
          let eventLocation = String(event.getLocation());
          let pattern = new RegExp(locationID);
          if (locationID === "") {
            if (eventLocation === "") {
              filteredEvents.push(event);
            }
          } else {
            if (eventLocation !== "" && pattern.test(eventLocation)) {
              filteredEvents.push(event);
            }
          }
        });
      });
      break;
    case "Title":
      console.log("Filtered by title");
      const titlePrompt = SpreadsheetApp.getUi().prompt(
        "Please enter the Title for events"
      );
      // Filter events and return results to printEventDetail function
      const string = String(titlePrompt.getResponseText());
      showDialog("Loading", 200, 100, "Filtering Events");
      events.forEach((array) => {
        array.forEach((event) => {
          let eventTitle = event.getTitle();
          if (compareStringWithTitle(string, eventTitle)) {
            filteredEvents.push(event);
          }
        });
      });
      break;
  }

  // Check to see if filter type provided returned events.  Else show alert and stop running script
  if (filteredEvents.length === 0) {
    let result = SpreadsheetApp.getUi().alert(
      `The value provided has no events in the specified year.\n\n Please check the value provided and try again.`
    );
    if (result === SpreadsheetApp.getUi().Button.OK) {
    }
    // Will not continue with process of printing Events
  } else {
    printEventsDetails([filteredEvents]);
  }
} // END OF: filter function

// This function is meant to return events with mistakes in the title, but still filter results enough to make going through the list easy
function compareStringWithTitle(string, title) {
  let isMatch = false;
  const lowerCaseString = string.toLowerCase();
  const lowerCaseTitle = title.toLowerCase();
  const arrayString = Array.from(lowerCaseString);
  const arrayTitle = Array.from(lowerCaseTitle);

  let longerArray;
  let shorterArray;

  if (arrayString.length >= arrayTitle.length) {
    longerArray = arrayString;
    shorterArray = arrayTitle;
  } else {
    longerArray = arrayTitle;
    shorterArray = arrayString;
  }

  // Test for a 2/3 match of letters between given string and event title
  let matchCount = 0;
  for (i in longerArray) {
    if (shorterArray.includes(longerArray[i])) {
      matchCount += 1;
    }
  }
  if (
    matchCount >= longerArray.length * 0.666667 &&
    shorterArray[0] === longerArray[0]
  ) {
    isMatch = true;
  } else {
    isMatch = false;
  }

  if (isMatch) {
    // Two thirds of the letters in the given string match the event title and first letter of each matches
    return true;
  } else {
    // Check two word entries and test if given string is one of those words
    const titleSplitArray = lowerCaseTitle.split(" ");
    if (titleSplitArray.includes(lowerCaseString)) {
      return true;
    } else {
      return false;
    }
  }
} // END OF: compareStringWithTitle function
