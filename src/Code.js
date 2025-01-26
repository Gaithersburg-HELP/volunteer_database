// Using Google Rhino with v8 runtime disabled due to this bug with multi-login when exporting contacts and deleting the sheet automatically: https://stackoverflow.com/questions/60230776/were-sorry-a-server-error-occurred-while-reading-from-storage-error-code-perm
// Google Rhino supports ES5 only, so ES6 features like let and for..of are not available

// 1-based index of data columns after group/application columns
var DATA_COLUMNS = Object.freeze({
  Name: 1,
  FirstName: 2,
  LastName: 3,
  Title: 4,
  Address: 5,
  City: 6,
  State: 7,
  Zip: 8,
  Phone: 9,
  PhoneType: 10,
  AltPhone: 11,
  Email1: 12,
  Email2: 13,
  Position: 14,
  Years: 15,
  Notes: 16,
  Availability: 17,
  Frequency: 18,
  OtherSkills: 19
});

// 1-based index of columns specific to Pending/Accepted/Rejected sheets
var APPLICATION_COLUMNS = Object.freeze({
  Date: 1,
  Accepted: 2,
  Tracking: 22
});

var VOL_DATA_START_COLUMN = 16; // First column after Group columns on Volunteers/Ex Volunteers
var APP_DATA_START_COLUMN = 3; // First column after initial application columns on Pending/Accepted/Rejected
var DATA_LENGTH = Object.keys(DATA_COLUMNS).length; // Number of columns other than Group columns on Volunteers/Ex Volunteers
var CONTACT_DATA_LENGTH = 13; // Number of columns containing contact export data. Assumes this is contiguous
var APPLICATION_ADDITIONAL_DATA_LENGTH = Object.keys(APPLICATION_COLUMNS).length;

var HEADERS = Object.freeze({
  Title: "Title: ",
  FirstName: "First-name: ",
  LastName: "Last-name: ",
  Email: "Email-address: ",
  StreetAddr: "Street-address: ",
  City: "City: ",
  State: "State: ",
  Zip: "Zipcode: ",
  Phone: "Phone Number: ",
  PhoneType: /(?=(\b((Home)|(Cell)) Phone\b))/,
  Age: "Age: ",
  Availability: "Availability: ",
  Frequency: "Frequency: ",
  VolPositions: "Volunteer-positions: ",
  OtherPositions: "Other-position: ",
  OtherSkills: "Other-skills: ",
  Referral: "Referral: ",
  WhyInterest: "Why-interested: ",
  Signature: "Electronic-signature: "
});

var HEADERS_INDEX = [
  "Title: ",
  "First-name: ",
  "Last-name: ",
  "Email-address: ",
  "Street-address: ",
  "City: ",
  "State: ",
  "Zipcode: ",
  "Phone Number: ",
  /(?=(\b((Home)|(Cell)) Phone\b))/,
  "Age: ",
  "Availability: ",
  "Frequency: ",
  "Volunteer-positions: ",
  "Other-position: ",
  "Other-skills: ",
  "Referral: ",
  "Why-interested: ",
  "Electronic-signature: "
];

// Returns next header in HEADERS enum given a header
function getNextHeader(header) {
  // findIndex not available in ES5 so basic implementation using reduce
  var headerIndex = HEADERS_INDEX.reduce(function (returnIndex, element, currentIndex) {
    var found = false;
    if (header instanceof RegExp) {
      found = element.toString() === header.toString();
    } else {
      found = element === header;
    }
    if (found) {
      return currentIndex;
    } else {
      return returnIndex;
    }
  }, -1);
  var nextHeader = HEADERS[Object.keys(HEADERS)[headerIndex + 1]];
  return nextHeader;
}

function capitalize(uncapitalized) {
  return uncapitalized.substring(0, 1).toUpperCase().concat(uncapitalized.substring(1, uncapitalized.length));
}

// Function to return field value following header (Header: value)
function FieldValue(data, header) {
  var found = new RegExp(header).exec(data); // 0-based
  if (found === null) {
    return "";
  }

  var headerStart = found.index;

  // includes starting space when value exists, trim it later
  // space does not exist when value does not exist
  var valueStart = headerStart + found[0].length;

  // ui.prompt replaces line breaks with a space, so we only see headers in the user input string
  var value = "";

  if (header === HEADERS.Signature) {
    value = data.substring(valueStart);
  } else {
    var nextHeaderStart = data.search(getNextHeader(header));
    value = data.substring(valueStart, nextHeaderStart);
  }

  return value.trim();
}

function formatPhone(phoneUnformatted) {
  // Remove all user entered separation characters (spaces, hyphens, periods, parens), country code (+digits)
  var phone = phoneUnformatted.replace(/[ .\-()+]/g, "");
  // Country code can have 1-6 digits at beginning, so grab last 10 digits
  phone = phone.substring(phone.length, phone.length - 10);

  if (!phone) {
    return phone;
  }

  // Format as ###-###-####
  phone = phone.substring(0, 3).concat("-", phone.substring(3, 6), "-", phone.substring(6, 10));

  return phone;
}

function addToValue(value, addValue) {
  if (value === "") {
    return addValue;
  } else {
    return value.concat("; ", addValue);
  }
}

function InsertVolData(data) {
  var firstName = FieldValue(data, HEADERS.FirstName);
  if (firstName === "") {
    throw Error("Missing required heading First-name in copied text");
  }
  firstName = capitalize(firstName);

  var lastName = FieldValue(data, HEADERS.LastName);
  if (lastName === "") {
    throw Error("Missing required heading Last-name in copied text");
  }
  lastName = capitalize(lastName);

  var email = FieldValue(data, HEADERS.Email);
  if (email === "") {
    throw Error("Missing required heading Email-address in copied text");
  }
  email = email.toLowerCase();

  var street = FieldValue(data, HEADERS.StreetAddr);
  if (street === "") {
    throw Error("Missing required heading Street-address in copied text");
  }
  street = capitalize(street);

  var city = FieldValue(data, HEADERS.City);
  if (city === "") {
    throw Error("Missing required heading City in copied text");
  }
  city = capitalize(city);

  var state = FieldValue(data, HEADERS.State);
  if (state === "") {
    throw Error("Missing required heading State in copied text");
  }
  state = state.toUpperCase();
  if (state === "MARYLAND") {
    state = "MD";
  }

  var zip = FieldValue(data, HEADERS.Zip);
  if (zip === "") {
    throw Error("Missing required heading Zipcode in copied text");
  }

  var phone = FieldValue(data, HEADERS.Phone);
  phone = formatPhone(phone);
  if (phone === "") {
    throw Error("Missing required heading Phone Number in copied text");
  }

  var phoneType = FieldValue(data, HEADERS.PhoneType);
  if (phoneType === "") {
    throw Error("Missing required Phone Type in copied text");
  }
  phoneType = phoneType.replace(" Phone", "");

  var availability = FieldValue(data, HEADERS.Availability);
  if (availability === "") {
    throw Error("Missing required heading Availability in copied text");
  }
  availability = availability.replace(/,/g, ";");

  var frequency = FieldValue(data, HEADERS.Frequency);
  if (frequency === "") {
    throw Error("Missing required heading Frequency in copied text");
  }

  var positions = FieldValue(data, HEADERS.VolPositions);
  if (positions === "") {
    throw Error("Missing required heading Volunteer-positions in copied text");
  }

  var notes = "";
  if (positions.search("Call Center - Food/Infant") !== -1) {
    notes = addToValue(notes, "Food Coord");
  }
  if (positions.search("Call Center - Prescription Financial") !== -1) {
    notes = addToValue(notes, "Rx Coord");
  }
  if (positions.search("Call Center - Bilingual Assistance") !== -1) {
    notes = addToValue(notes, "Bilingual Assistant");
  }
  if (positions.search("Call Center - Transportation Requests") !== -1) {
    notes = addToValue(notes, "Trans Sched");
  }
  if (positions.search("Driver - Client Transportation") !== -1) {
    notes = addToValue(notes, "Trans Driver");
  }
  if (positions.search("Driver - Donation Pickup") !== -1) {
    notes = addToValue(notes, "Donation Pickup");
  }
  if (positions.search("Driver - Food Delivery") !== -1) {
    notes = addToValue(notes, "Food Driver");
  }
  if (positions.search("IT Assistant") !== -1) {
    notes = addToValue(notes, "IT Assistant");
  }
  if (positions.search("Pantry - Food Bagger") !== -1) {
    notes = addToValue(notes, "Food Bagger");
  }
  if (positions.search("Pantry - Food Distributor") !== -1) {
    notes = addToValue(notes, "Pantry Worker");
  }
  if (positions.search("Pantry - Food Receiver") !== -1) {
    notes = addToValue(notes, "Food Receiver");
  }
  if (positions.search("Pantry - Food Stocker") !== -1) {
    notes = addToValue(notes, "Pantry Stocker");
  }
  if (positions.search("Seasonal Events") !== -1) {
    notes = addToValue(notes, "Seasonal Events");
  }
  if (positions.search("Social Media Coordinator") !== -1) {
    notes = addToValue(notes, "Social Media Coordinator");
  }
  if (positions.search("Sponsor a Food Drive") !== -1) {
    notes = addToValue(notes, "Sponsor a Food Drive");
  }
  if (positions.search("Groups") !== -1) {
    notes = addToValue(notes, "Groups");
  }

  if (!FieldValue(data, HEADERS.Signature)) {
    throw Error("Missing required heading Electronic-signature in copied text");
  }

  var title = FieldValue(data, HEADERS.Title);
  // clear title if set to default value of ' Click here to choose title'
  if (title.search("Click") !== -1) {
    title = "";
  }
  title = capitalize(title);
  // keep clergy titles only
  switch (title) {
    case "Mr.":
    case "Mrs.":
    case "Ms.":
      title = "";
      break;
    default:
      break;
  }

  var ageUnder18 = FieldValue(data, HEADERS.Age);
  if (ageUnder18) {
    notes = addToValue(notes, ageUnder18.concat(" years old on start year"));
  }
  var otherSkills = FieldValue(data, HEADERS.OtherSkills);
  otherSkills = otherSkills.replace(/,/g, ";");
  var otherPosition = FieldValue(data, HEADERS.OtherPositions);

  //var referral = FieldValue(data, HEADERS.Referral);
  //var whyInterested = FieldValue(data, HEADERS.WhyInterest);

  var year = new Date().toJSON().slice(0, 4);
  var currentDate = new Date().toJSON().slice(0, 10);

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pending");
  var lastRow = sheet.getLastRow() + 1;

  var range = sheet.getRange(lastRow, 1, 1, DATA_LENGTH + APPLICATION_ADDITIONAL_DATA_LENGTH);

  var valuesToSet = [];
  valuesToSet[APPLICATION_COLUMNS.Date - 1] = currentDate;
  valuesToSet[APPLICATION_COLUMNS.Accepted - 1] = "Pending";
  valuesToSet[APP_DATA_START_COLUMN - 1 + DATA_COLUMNS.Name - 1] = firstName.concat(" ", lastName);
  valuesToSet[APP_DATA_START_COLUMN - 1 + DATA_COLUMNS.FirstName - 1] = firstName;
  valuesToSet[APP_DATA_START_COLUMN - 1 + DATA_COLUMNS.LastName - 1] = lastName;
  valuesToSet[APP_DATA_START_COLUMN - 1 + DATA_COLUMNS.Title - 1] = title;
  valuesToSet[APP_DATA_START_COLUMN - 1 + DATA_COLUMNS.Address - 1] = street;
  valuesToSet[APP_DATA_START_COLUMN - 1 + DATA_COLUMNS.City - 1] = city;
  valuesToSet[APP_DATA_START_COLUMN - 1 + DATA_COLUMNS.State - 1] = state;
  valuesToSet[APP_DATA_START_COLUMN - 1 + DATA_COLUMNS.Zip - 1] = zip;
  valuesToSet[APP_DATA_START_COLUMN - 1 + DATA_COLUMNS.Phone - 1] = phone;
  valuesToSet[APP_DATA_START_COLUMN - 1 + DATA_COLUMNS.PhoneType - 1] = phoneType;
  valuesToSet[APP_DATA_START_COLUMN - 1 + DATA_COLUMNS.AltPhone - 1] = "";
  valuesToSet[APP_DATA_START_COLUMN - 1 + DATA_COLUMNS.Email1 - 1] = email;
  valuesToSet[APP_DATA_START_COLUMN - 1 + DATA_COLUMNS.Email2 - 1] = "";
  valuesToSet[APP_DATA_START_COLUMN - 1 + DATA_COLUMNS.Position - 1] = otherPosition;
  valuesToSet[APP_DATA_START_COLUMN - 1 + DATA_COLUMNS.Years - 1] = year;
  valuesToSet[APP_DATA_START_COLUMN - 1 + DATA_COLUMNS.Notes - 1] = notes;
  valuesToSet[APP_DATA_START_COLUMN - 1 + DATA_COLUMNS.Availability - 1] = availability;
  valuesToSet[APP_DATA_START_COLUMN - 1 + DATA_COLUMNS.Frequency - 1] = frequency;
  valuesToSet[APP_DATA_START_COLUMN - 1 + DATA_COLUMNS.OtherSkills - 1] = otherSkills;
  valuesToSet[APPLICATION_COLUMNS.Tracking - 1] = "";

  range.setValues([valuesToSet]);

  range.activate();
}

function AddVol() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt("Paste application email here", ui.ButtonSet.OK_CANCEL);

  var button = result.getSelectedButton();
  if (button === ui.Button.CANCEL) {
    return;
  }

  var text = result.getResponseText().trim();
  if (!text) {
    throw new Error("Invalid data pasted");
  }

  InsertVolData(text);
}

function getSelectedRows() {
  var ranges = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveRangeList().getRanges();
  var selectedRows = [];

  for (var i = 0; i < ranges.length; i++) {
    var rng = ranges[i];
    var topRow = rng.getRow();
    for (var j = topRow; j < topRow + rng.getHeight(); j++) {
      if (j === 1) {
        throw Error("Cannot select header row, deselect it and try again");
      }
      if (selectedRows.indexOf(j) === -1) {
        selectedRows.push(j);
      }
    }
  }
  return selectedRows;
}

function emptyGroupColArray() {
  // Array.fill not supported in Rhino
  var rowToAppend = Array(VOL_DATA_START_COLUMN - 1);
  var i = 0;
  while (i < rowToAppend.length) {
    rowToAppend[i] = "";
    i = i + 1;
  }

  return rowToAppend;
}

function MoveBetweenVolSheets(sourceSheet, destSheet) {
  var rowsToMove = getSelectedRows();

  rowsToMove.sort(function (a, b) {
    return b - a;
  });

  var startCol = VOL_DATA_START_COLUMN;
  var length = DATA_LENGTH;
  if (sourceSheet.getName() === "Pending") {
    startCol = 1;
    length = DATA_LENGTH + APPLICATION_ADDITIONAL_DATA_LENGTH;
  }

  rowsToMove.forEach(function (row) {
    var rowToAppend = [];
    if (sourceSheet.getName() !== "Pending") {
      // Include Group columns when moving from Volunteer/Ex Volunteer sheet
      rowToAppend = sourceSheet.getRange(row, 1, 1, VOL_DATA_START_COLUMN - 1).getValues()[0];
    }
    rowToAppend = rowToAppend.concat(sourceSheet.getRange(row, startCol, 1, length).getValues()[0]);
    destSheet.appendRow(rowToAppend);
    sourceSheet.deleteRow(row);
  });

  if (sourceSheet.getName() !== "Pending") {
    destSheet.sort(VOL_DATA_START_COLUMN + DATA_COLUMNS.Name - 1);
  }
}

function AcceptVol() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert("Accept selected volunteers?", ui.ButtonSet.OK_CANCEL);

  if (result === ui.Button.CANCEL) {
    return;
  }

  var rowsToAccept = getSelectedRows();
  var volSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Volunteers");

  rowsToAccept.forEach(function (rowNum, index, array) {
    SpreadsheetApp.getActiveSheet().getRange(rowNum, APPLICATION_COLUMNS.Accepted).setValue("Yes");

    // Skip Application Date, Accepted, Tracking columns
    rowToAppend = emptyGroupColArray().concat(
      SpreadsheetApp.getActiveSheet().getRange(rowNum, APP_DATA_START_COLUMN, 1, DATA_LENGTH).getValues()[0]
    );

    volSheet.appendRow(rowToAppend);
  });

  volSheet.sort(VOL_DATA_START_COLUMN + DATA_COLUMNS.Name - 1);

  MoveBetweenVolSheets(
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pending"),
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Accepted")
  );
}

function RejectVol() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert("Reject selected volunteers?", ui.ButtonSet.OK_CANCEL);

  if (result === ui.Button.CANCEL) {
    return;
  }

  getSelectedRows().forEach(function (rowNum, index, array) {
    SpreadsheetApp.getActiveSheet().getRange(rowNum, APPLICATION_COLUMNS.Accepted).setValue("No");
  });
  MoveBetweenVolSheets(
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pending"),
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Rejected")
  );
}

function DeleteSelected() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    "Delete selected applications from this sheet? Accepted/Rejected volunteers are kept in the database. Deleted applicants will not appear in the Reports sheet, so backup any Reports data before continuing.",
    ui.ButtonSet.OK_CANCEL
  );

  if (result === ui.Button.CANCEL) {
    return;
  }

  var rowsToDelete = getSelectedRows();
  rowsToDelete.sort(function (a, b) {
    return b - a;
  });

  rowsToDelete.forEach(function (rowNum, index, array) {
    SpreadsheetApp.getActiveSheet().deleteRow(rowNum);
  });
}

function deleteExportSheet(sheetName) {
  // save selection, selection gets deleted after deleteSheet
  var ranges = SpreadsheetApp.getActiveRangeList();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheet);
  SpreadsheetApp.getActiveSheet().setActiveRangeList(ranges);
}

function showExportDialog(exportSheetName) {
  // from https://stackoverflow.com/questions/63265983/google-apps-script-how-to-export-specific-sheet-as-csv-format
  var ssID = SpreadsheetApp.getActiveSpreadsheet().getId();
  var sheetId = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(exportSheetName).getSheetId().toString();
  var params = ssID + "/export?gid=" + sheetId + "&format=csv";
  var url = "https://docs.google.com/spreadsheets/d/" + params;

  // from https://stackoverflow.com/questions/10744760/google-apps-script-to-open-a-url
  var html = HtmlService.createHtmlOutput(
    "<html>" +
      // Offer URL as clickable link in case code fails.
      "<head><style>" +
      "#container {" +
      "position: relative;" +
      "width: 50%;" +
      "margin-left: auto;" +
      "margin-right: auto;" +
      "}" +
      "#loading, #close {" +
      "width: 100%;" +
      "}" +
      "#close {" +
      "position: absolute;" +
      "top: 0;" +
      "min-width: 64px;" +
      "border: none;" +
      "outline: none;" +
      "background: rgba(0,0,0,0);" +
      "-webkit-font-smoothing: antialiased;" +
      "transition: box-shadow 280ms cubic-bezier(0.4, 0, 0.2, 1);" +
      "padding: 0 16px 0 16px;" +
      "will-change: transform,opacity;" +
      "height: 36px;" +
      "border-radius: 4px;" +
      "border-style: none;" +
      "-webkit-tap-highlight-color: rgba(232, 230, 227, 0);" +
      "box-shadow: 0px 3px 1px -2px rgba(0, 0, 0, 0.2), 0px 2px 2px 0px rgba(0, 0, 0, 0.14), 0px 1px 5px 0px rgba(0, 0, 0, 0.12);" +
      "background-color: #4e00be;" +
      "color: #e8e6e3;" +
      "}" +
      // From https://codepen.io/mandelid/pen/kNBYLJ
      "#loading {" +
      "position: relative;" +
      "margin-left: auto;" +
      "margin-right: auto;" +
      "visibility: hidden;" +
      "width: 50px;" +
      "height: 50px;" +
      "border: 3px solid #4e00be;" +
      "border-radius: 50%;" +
      "border-top-color: #fff;" +
      "animation: spin 1s ease-in-out infinite;" +
      "-webkit-animation: spin 1s ease-in-out infinite;" +
      "}" +
      "@keyframes spin {" +
      "to { -webkit-transform: rotate(360deg); }" +
      "}" +
      "@-webkit-keyframes spin {" +
      "to { -webkit-transform: rotate(360deg); }" +
      "}" +
      "</style></head>" +
      '<body style="word-break:break-word;font-family:sans-serif;"><p><a href="' +
      url +
      '" target="_blank" id="link" onclick="window.close()" download>Right click this link</a> and click Save Link As to save the CSV.</p>' +
      '<div id="container">' +
      '<div id="loading"></div>' +
      '<button id="close" onClick="closeDelete()">Close and delete sheet</button>' +
      "</div>" +
      '<p id="error"></p>' +
      "</body><script>" +
      "function closeDelete() {" +
      'var exportSheetName = "' +
      exportSheetName +
      '";' +
      'document.getElementById("close").style.visibility = "hidden";' +
      'document.getElementById("loading").style.visibility = "visible";' +
      "google.script.run" +
      ".withSuccessHandler(() => {" +
      "google.script.host.close()" +
      "})" +
      ".withFailureHandler(() => {" +
      'document.getElementById("loading").style.visibility = "hidden";' +
      'document.getElementById("error").append("Failed to delete sheet, close this dialog in the top right and delete sheet manually.");' +
      "})" +
      ".deleteExportSheet(exportSheetName);" +
      "}" +
      "</script>" +
      "</html>"
  )
    .setWidth(450)
    .setHeight(170);

  SpreadsheetApp.getUi().showModalDialog(html, "Exporting ...");
}

function getSelectedGroupColumns() {
  var ranges = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveRangeList().getRanges();
  var selectedColumns = [];

  for (var i = 0; i < ranges.length; i++) {
    var rng = ranges[i];
    var leftmostCol = rng.getColumn();
    for (var j = leftmostCol; j < leftmostCol + rng.getWidth(); j++) {
      if (j >= VOL_DATA_START_COLUMN) {
        throw Error("Cannot select data columns, deselect them and try again");
      }
      if (selectedColumns.indexOf(j) === -1) {
        selectedColumns.push(j);
      }
    }
  }
  return selectedColumns;
}

// As of 9/24/24:
// Google Groups CSVs can bulk upload duplicate emails of emails already in group, they are ignored
// Google Groups CSVs cannot upload to multiple groups at the same time, all emails must have same group email
function ExportGroup() {
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = activeSheet.getDataRange().getValues();
  var selectedGroupColumns = getSelectedGroupColumns();

  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt("Enter group email to bulk upload to here", ui.ButtonSet.OK_CANCEL);

  var button = result.getSelectedButton();
  if (button === ui.Button.CANCEL) {
    return;
  }

  var groupEmail = result.getResponseText().trim();

  var exportSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
  var exportSheetName = exportSheet.getSheetName();

  var rowsToAppend = [];

  rowsToAppend.push(["Group Email [Required]", "Member Email", "Member Type", "Member Role"]);

  var emailsToExport = [];
  for (var i in data) {
    if (i !== "0") {
      // continue does not work in Rhino
      for (var j in selectedGroupColumns) {
        if (data[i][selectedGroupColumns[j] - 1]) {
          emailsToExport.push(data[i][VOL_DATA_START_COLUMN - 1 + DATA_COLUMNS.Email1 - 1]);
          break;
        }
      }
    }
  }

  emailsToExport.forEach(function (email, index, array) {
    var rowToAppend = [groupEmail];
    rowToAppend = rowToAppend.concat(email);
    rowToAppend = rowToAppend.concat(["USER", "MEMBER"]);
    rowsToAppend.push(rowToAppend);
  });

  exportSheet.getRange(1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);

  showExportDialog(exportSheetName);
}

// From https://stackoverflow.com/a/25868785/13342792
function orderedHash() {
  var keys = [];
  var vals = {};
  return {
    push: function (k, v) {
      if (!Object.prototype.hasOwnProperty.call(vals, k)) keys.push(k);
      vals[k] = v;
    },
    insert: function (pos, k, v) {
      if (!Object.prototype.hasOwnProperty.call(vals, k)) {
        keys.splice(pos, 0, k);
        vals[k] = v;
      }
    },
    val: function (k) {
      return vals[k];
    },
    length: function () {
      return keys.length;
    },
    keys: function () {
      return keys;
    },
    values: function () {
      return vals;
    }
  };
}

function ExportContact() {
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rowsToExport = getSelectedRows();

  var exportSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
  var exportSheetName = exportSheet.getSheetName();

  var startColumn = VOL_DATA_START_COLUMN;
  if (activeSheet.getName() === "Pending") {
    startColumn = APP_DATA_START_COLUMN;
  }

  // Google Contacts expects certain header values, see template here:
  // https://support.google.com/contacts/answer/15147365?hl=en&p=contacts_template&rd=1#contacts_template
  // Key: Header
  // Value: zero-based column index from database starting from startColumn. CONTACT_DATA_LENGTH is ""
  var CONTACTS_HEADERS = orderedHash();
  CONTACTS_HEADERS.push("Name Prefix", DATA_COLUMNS.Title - 1);
  CONTACTS_HEADERS.push("First Name", DATA_COLUMNS.FirstName - 1);
  CONTACTS_HEADERS.push("Middle Name", CONTACT_DATA_LENGTH);
  CONTACTS_HEADERS.push("Last Name", DATA_COLUMNS.LastName - 1);
  CONTACTS_HEADERS.push("Name Suffix", CONTACT_DATA_LENGTH);
  CONTACTS_HEADERS.push("Phonetic First Name", CONTACT_DATA_LENGTH);
  CONTACTS_HEADERS.push("Phonetic Middle Name", CONTACT_DATA_LENGTH);
  CONTACTS_HEADERS.push("Phonetic Last Name", CONTACT_DATA_LENGTH);
  CONTACTS_HEADERS.push("Nickname", CONTACT_DATA_LENGTH);
  CONTACTS_HEADERS.push("E-mail 1 - Label", CONTACT_DATA_LENGTH);
  CONTACTS_HEADERS.push("E-mail 1 - Value", DATA_COLUMNS.Email1 - 1);
  CONTACTS_HEADERS.push("E-mail 2 - Value", DATA_COLUMNS.Email2 - 1);
  CONTACTS_HEADERS.push("Phone 1 - Label", DATA_COLUMNS.PhoneType - 1);
  CONTACTS_HEADERS.push("Phone 1 - Value", DATA_COLUMNS.Phone - 1);
  CONTACTS_HEADERS.push("Phone 2 - Value", DATA_COLUMNS.AltPhone - 1);
  CONTACTS_HEADERS.push("Address 1 - Label", CONTACT_DATA_LENGTH);
  CONTACTS_HEADERS.push("Address 1 - Country", CONTACT_DATA_LENGTH);
  CONTACTS_HEADERS.push("Address 1 - Street", DATA_COLUMNS.Address - 1);
  CONTACTS_HEADERS.push("Address 1 - Extended Address", CONTACT_DATA_LENGTH);
  CONTACTS_HEADERS.push("Address 1 - City", DATA_COLUMNS.City - 1);
  CONTACTS_HEADERS.push("Address 1 - Region", DATA_COLUMNS.State - 1);
  CONTACTS_HEADERS.push("Address 1 - Postal Code", DATA_COLUMNS.Zip - 1);

  exportSheet.appendRow(CONTACTS_HEADERS.keys());

  rowsToExport.forEach(function (rowNum, value, set) {
    var dataToAppend = activeSheet.getRange(rowNum, startColumn, 1, CONTACT_DATA_LENGTH).getValues()[0];
    dataToAppend.push(""); // so below loop can reference empty string
    var arrayToAppend = [];

    CONTACTS_HEADERS.keys().forEach(function (key, index) {
      arrayToAppend.push(dataToAppend[CONTACTS_HEADERS.val(key)]);
    });
    exportSheet.appendRow(arrayToAppend);
  });

  showExportDialog(exportSheetName);
}

function RetireVol() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert("Retire selected volunteers?", ui.ButtonSet.OK_CANCEL);

  if (result === ui.Button.CANCEL) {
    return;
  }

  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var exVolSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ExVolunteers");
  MoveBetweenVolSheets(activeSheet, exVolSheet);
}

function ReinstateVol() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert("Reinstate selected volunteers?", ui.ButtonSet.OK_CANCEL);

  if (result === ui.Button.CANCEL) {
    return;
  }

  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var volSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Volunteers");
  MoveBetweenVolSheets(activeSheet, volSheet);
}
