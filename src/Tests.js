function logTest(cell, correctValue) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pending");
  var testResult = sheet.getRange(cell).getDisplayValue() === correctValue;

  console.log(cell, ": ", sheet.getRange(cell).getValue(), " = ", correctValue, ": ", testResult);
  return testResult;
}

function testAddApplication() {
  console.log("TEST RESULTS");

  var testData =
    "Title: Dr.\n" +
    "First-name: TEST IGNORE First\n" +
    "Last-name: Last\n" +
    "Email-address: email@example.com\n" +
    "Street-address: 311 High Gables Drive, Unit 310\n" +
    "City: Gaithersburg\n" +
    "State: MD\n" +
    "Zipcode: 12345\n" +
    "Phone Number: 111-222-3333\n" +
    "Cell Phone\n" +
    "Age: 15\n" +
    "Availability: I am flexible, Weekdays 9:00am - 1:00pm, Weekdays 1:00pm - 5:00pm, Weekdays 5:00pm - 8:00pm\n" +
    "Frequency: 1-5 hours per week\n" +
    "Volunteer-positions: Call Center - Food/Infant, Call Center - Prescription Financial, Call Center - Bilingual Assistance, Call Center - Transportation Requests, Driver - Client Transportation *, Driver - Donation Pickup *, Driver - Food Delivery *, IT Assistant, Pantry - Food Bagger **, Pantry - Food Distributor **, Pantry - Food Receiver **, Pantry - Food Stocker **, Seasonal Events, Social Media Coordinator, Sponsor a Food Drive, Groups - No opportunities for groups of more than 2-3 people available, Other position not listed - please enter below\n" +
    "Other-position: other position\n" +
    "Other-skills: Accounting, Administration, Fundraising, Grant Writing, Information Technology, Management\n" +
    "Referral: hear about volunteer\n" +
    "Why-interested: interest in volunteering\n" +
    "Electronic-signature: Applicant Authorization\n";

  InsertVolData(testData);

  var lastRow = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pending").getLastRow();

  var allResults =
    logTest("A".concat(lastRow), new Date().toJSON().slice(0, 10)) &&
    logTest("B".concat(lastRow), "Pending") &&
    logTest("C".concat(lastRow), "TEST IGNORE First Last") &&
    logTest("D".concat(lastRow), "TEST IGNORE First") &&
    logTest("E".concat(lastRow), "Last") &&
    logTest("F".concat(lastRow), "Dr.") &&
    logTest("G".concat(lastRow), "311 High Gables Drive, Unit 310") &&
    logTest("H".concat(lastRow), "Gaithersburg") &&
    logTest("I".concat(lastRow), "MD") &&
    logTest("J".concat(lastRow), "12345") &&
    logTest("K".concat(lastRow), "111-222-3333") &&
    logTest("L".concat(lastRow), "Cell") && // Phone type
    logTest("M".concat(lastRow), "") && //Alt phone
    logTest("N".concat(lastRow), "email@example.com") &&
    logTest("O".concat(lastRow), "") && //2nd email
    logTest("P".concat(lastRow), "other position") && //Position
    logTest("Q".concat(lastRow), new Date().getFullYear().toString()) &&
    logTest(
      "R".concat(lastRow),
      "Food Coord; Rx Coord; Bilingual Assistant; Trans Sched; Trans Driver; Donation Pickup; Food Driver; IT Assistant; Food Bagger; Pantry Worker; Food Receiver; Pantry Stocker; Seasonal Events; Social Media Coordinator; Sponsor a Food Drive; Groups; 15 years old on start year"
    ) && //Notes
    logTest(
      "S".concat(lastRow),
      "I am flexible; Weekdays 9:00am - 1:00pm; Weekdays 1:00pm - 5:00pm; Weekdays 5:00pm - 8:00pm"
    ) && //Availability
    logTest("T".concat(lastRow), "1-5 hours per week") && //Frequency
    logTest(
      "U".concat(lastRow),
      "Accounting; Administration; Fundraising; Grant Writing; Information Technology; Management"
    ); //Other skills

  //logTest(testData, HEADERS.OtherSkills, "Administration, Grant Writing") &&
  //logTest(testData, HEADERS.Referral, "City of Gaithersburg website") &&
  //logTest(testData, HEADERS.WhyInterest, "I have been a resident of Gaithersburg since 2017, and have lived in the area almost my entire life. I would like to feel more connected to my local community and give back, as well as meet like-minded individuals who are interested in doing the same.") &&

  var homeTestData = testData.replace("Cell ", "Home ");

  InsertVolData(homeTestData);

  lastRow = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pending").getLastRow();

  allResults = allResults && logTest("L".concat(lastRow), "Home");

  console.log(allResults ? "PASSING" : "FAILING");
}
