/***** DO NOT CHANGE!!!. Well don't if you don't know what you are doing 😊 *****/
const SETTINGS =
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
const LOG = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Log");
const STATS = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Stats");
const DOC =
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Documentation");
const ALL_UNRETURNS =
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("All Unreturns");
const ACTIVESHEET = SpreadsheetApp.getActiveSheet();
const VALIDATION = SpreadsheetApp.newDataValidation();
const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
optimizeScan = SETTINGS.getRange("OptimizeScan").getValue();

/***** Admin credential for AdminDirectory API *****/
const MY_CUSTOMER = "my_customer";

/***** All the statuses *****/
const STATUS = [
  "Awaiting Return",
  "Do Not Disable",
  "Did Not Return",
  "Returned",
];

/***** Assign each state *****/
const [AR, DND, DNR, R] = STATUS;

/***** Student and faculty domain *****/
const STUDENT_DOMAIN = "@learn.hohschools.org";
const FACULTY_DOMAIN = "@hohschools.org";

/***** Colors *****/
const SUCESS_COLOR = "#2fda4e";
const ERROR_COLOR = "#fc5736";
const RESET_COLOR = "#fff";
const WARNING_COLOR = "#e98a15";
const NEUTRAL_COLOR = "#34f6f2";

/********************************************
 * Reasons for borrowing a chrome book.      *
 * You can add more reasons if you want.     *
 * Or even change them.                      *
 * All reasons should be separated by comma. *
 *********************************************/
const REASONS = [
  "Left my chromebook home",
  "My chromebook is missing / damaged",
  "My chromebook is in for repairs",
  "Other",
];

/***** Setup the first reason to be the default reason for borrowing a chromebook *****/
const [DEFAULT_REASON] = REASONS;

const tester = () => {
  PropertiesService.getDocumentProperties().getKeys().forEach(console.log);
  console.log(CacheService.getScriptCache());
};

/***** Open the LOG sheet when the sheet is and set current cell to the next available cell when opened *****/
const onStart = () => {
  menu();
  LOG.activate();
  SpreadsheetApp.flush();
  LOG.getCurrentCell()
    .offset(LOG.getLastRow() - LOG.getCurrentCell().getLastRow() + 1, 0)
    .activateAsCurrentCell();
  SpreadsheetApp.flush();
};

/**** Create the menu when the sheet is opened ****/
const menu = () => {
  const UI = SpreadsheetApp.getUi();
  UI.createMenu("Loaner")
    .addItem("Disable A Chromebook by Tag", "disableChromebookByTag")
    .addItem("Sidebar", "showSideBar")
    .addToUi();
};

/***** User information helper functions *****/
const getUser = (email) => {
  let user = AdminDirectory.Users.get(email);
  return user;
};

/****** Get the grade from each student's OU path ********/
const getGradeFromOU = (path) => {
  path = path.toString().toLowerCase();
  if (path.includes("/k")) return "K";
  if (path.includes("01")) return 1;
  if (path.includes("02")) return 2;
  if (path.includes("03")) return 3;
  if (path.includes("04")) return 4;
  if (path.includes("05")) return 5;
  if (path.includes("06")) return 6;
  if (path.includes("07")) return 7;
  if (path.includes("08")) return 8;
  if (path.includes("09")) return 9;
  if (path.includes("10")) return 10;
  if (path.includes("11")) return 11;
  if (path.includes("12")) return 12;
  else return "Faculty/Staff";
};

/***** Chromebook information helper functions *****/
const findChromebookByTag = (tag) => {
  tag = tag.toString().trim();
  try {
    let chromebook = AdminDirectory.Chromeosdevices.list(MY_CUSTOMER, {
      query: `id:${tag}`,
    });
    if (chromebook.chromeosdevices.length === 1) {
      let [device] = chromebook.chromeosdevices;
      return device;
    } else {
      chromebook.chromeosdevices.forEach((cb) => {
        let { serialNumber } = cb;
        if (serialNumber.toLowerCase() === tag.toLowerCase()) return cb;
      });
    }
  } catch (e) {
    SPREADSHEET.toast(`Couldn't find chromebook: ${tag}, ${e}`);
  }
};

const enableChromebook = (serviceTag) => {
  let { deviceId, status } = findChromebookByTag(serviceTag);
  if (status === "DISABLED" && deviceId)
    AdminDirectory.Chromeosdevices.action(
      { action: "reenable" },
      MY_CUSTOMER,
      deviceId
    );
};

const disableChromebook = (serviceTag) => {
  let { deviceId, status } = findChromebookByTag(serviceTag);
  if (status !== "DISABLED" && deviceId)
    AdminDirectory.Chromeosdevices.action(
      { action: "disable" },
      MY_CUSTOMER,
      deviceId
    );
};

/***** Log sheet helper functions *****/
const getActualLastRow = (range) => {
  let actualRow = range.length;
  for (let i = range.length - 1; i >= 0; i--) {
    if (range[i][0] == "" || range[i][3] == "") actualRow = i;
    if (range[i][0] != "" && range[i][3] != "") break;
  }
  return actualRow;
};

const removeEmptyRows = (sheet) => {
  let range = sheet.getRange("A:D").getValues();
  let lastRow = getActualLastRow(range) + 2;
  let maxRows = sheet.getMaxRows();
  if (lastRow < maxRows) sheet.deleteRows(lastRow, maxRows - lastRow);
  sheet.getRange(lastRow - 1, 1, 2, sheet.getLastColumn()).clearContent();
};

const removeAboveInvalidRows = (sheet, endIndex) => {
  let rowsToDelete = 0;
  let startIndex = 0;
  for (let i = endIndex; i >= 1; i--) {
    let currentRowAbove = sheet.getRange(i, 1, 1, sheet.getLastColumn());
    const [values] = currentRowAbove.getValues();
    const [email, , , cbTag, ...rest] = values;
    if (email == "" || cbTag == "") {
      rowsToDelete += 1;
      startIndex = i;
    }
    if (email != "" && cbTag != "") break;
  }
  sheet.deleteRows(startIndex, rowsToDelete);
};

const updateWithTodaysDate = () => {
  removeEmptyRows(LOG);
  for (let i = LOG.getLastRow(); i > 1; i--) {
    let todayDate = new Date().getDate();
    let currentRow = LOG.getRange(i, 1, 1, LOG.getLastColumn());
    let status = LOG.getRange(i, LOG.getLastColumn() - 1);
    let chromebookTag = LOG.getRange(i, 4).getValue().toString();
    let borrowedDate = new Date(LOG.getRange(i, 7).getValue()).getDate();
    if (status.getValue() == AR && borrowedDate >= todayDate) {
      status.setValue(DNR);
      currentRow.setBackground(ERROR_COLOR);
      disableChromebook(chromebookTag);
    }
    if (status.getValue() == DND) continue;
    if (new Date(borrowedDate).getTime() < new Date(todayDate).getTime()) break;
  }
};

const updateFromStartToFinish = () => {
  removeEmptyRows(LOG);
  for (let i = 1; i <= LOG.getLastRow(); i++) {
    let currentRow = LOG.getRange(i, 1, 1, LOG.getLastColumn());
    let status = LOG.getRange(i, LOG.getLastColumn() - 1);
    let chromebookTag = LOG.getRange(i, 4).getValue().toString();
    if (status.getValue() == AR) {
      status.setValue(DNR);
      currentRow.setBackground(ERROR_COLOR);
      disableChromebook(chromebookTag);
    }
  }
};

const updateLog = () =>
  optimizeScan ? updateWithTodaysDate() : updateFromStartToFinish();

const didNotReturn = (email) => {
  let unreturnedData = [];
  let unreturns = ALL_UNRETURNS.createTextFinder(email).findAll();
  unreturns.forEach((row) => {
    let eachRow = row.getRow();
    let [tagAndName] = ALL_UNRETURNS.getRange(eachRow, 4, 1, 2).getValues();
    let [tag, name] = tagAndName;
    unreturnedData.push(`${tag} (${name})`);
  });
  return unreturnedData;
};

/***** Log sheet actions *****/

const onEmailInput = () => {
  const UI = SpreadsheetApp.getUi();
  if (ACTIVESHEET.getName() === LOG.getName()) {
    let currentCell = LOG.getActiveCell();
    let previousRow = LOG.getRange(
      currentCell.offset(-1, 0).getRowIndex(),
      1,
      1,
      LOG.getLastColumn()
    );
    let currentRow = currentCell.getRowIndex();
    let input = currentCell.getValue().toString().toLowerCase().trim();
    let email = input;
    let nameCol = currentCell.offset(0, 1);
    let gradeCol = currentCell.offset(0, 2);
    try {
      if (
        currentCell.getColumn() === 1 &&
        currentCell.getRowIndex() > 1 &&
        !currentCell.isBlank()
      ) {
        if (email.includes(STUDENT_DOMAIN) || email.includes(FACULTY_DOMAIN)) {
          email = email;
        } else email += STUDENT_DOMAIN;
        let user = getUser(email);
        console.log(user);
        let { primaryEmail, name, orgUnitPath } = user;
        let { fullName } = name;
        console.log(primaryEmail, name, orgUnitPath);
        currentCell.setValue(primaryEmail);
        nameCol.setValue(fullName);
        gradeCol.setValue(getGradeFromOU(orgUnitPath));
        let isUserOwingReturn = didNotReturn(email);
        if (isUserOwingReturn.length >= 1)
          throw `${email} has not returned ${isUserOwingReturn.toString()}`;
      }
    } catch (e) {
      if (e.toString().includes("has not returned")) UI.alert(e);
      else {
        try {
          email = input + FACULTY_DOMAIN;
          let user = getUser(email);
          console.log(user);
          let { primaryEmail, name, orgUnitPath } = user;
          let { fullName } = name;
          currentCell.setValue(primaryEmail);
          nameCol.setValue(fullName);
          gradeCol.setValue(getGradeFromOU(orgUnitPath));
          let isUserOwingReturn = didNotReturn(email);
          if (isUserOwingReturn.length >= 1)
            throw `${email} has not returned ${isUserOwingReturn.toString()}`;
        } catch (e) {
          if (e.toString().includes("has not returned")) UI.alert(e);
          else {
            currentCell.setValue("Oops we couldn't find your email. Try again");
            gradeCol.clearContent();
            UI.alert(
              `
          User: ${input} could not be found.
          Please try again with the correct username.
          ${e}
          `
            );
            throw `Couldn't find ${input}`;
          }
        }
      }
    }
  }
};

const onTagInput = () => {
  if (ACTIVESHEET.getName() === LOG.getName()) {
    let currentCell = LOG.getActiveCell();
    let previousRow = LOG.getRange(
      currentCell.offset(-1, 0).getRowIndex(),
      1,
      1,
      LOG.getLastColumn()
    );
    let currentRow = currentCell.getRowIndex();
    let input = currentCell.getValue().toString().trim();
    let chromebookName = currentCell.offset(0, 1);
    let reason = currentCell.offset(0, 2);
    let borrowed = currentCell.offset(0, 3);
    let exceptioned = currentCell.offset(0, 4);
    let returned = currentCell.offset(0, 5);
    let status = currentCell.offset(0, 7);
    let email = currentCell.offset(0, -3);
    let enforceCheckbox = VALIDATION.requireCheckbox()
      .setAllowInvalid(false)
      .build();
    let reasons = VALIDATION.requireValueInList(REASONS, false)
      .setAllowInvalid(true)
      .build();
    try {
      if (
        currentCell.getColumn() === 4 &&
        currentCell.getRowIndex() > 1 &&
        !currentCell.isBlank()
      ) {
        if (email.isBlank()) {
          SPREADSHEET.toast(
            "Make sure the email is not blank.",
            "Empty Email",
            5
          );
          currentCell.clearContent();
        } else {
          let chromebook = findChromebookByTag(input);
          let { serialNumber, annotatedAssetId, annotatedUser } = chromebook;
          currentCell.setValue(serialNumber);
          chromebookName.setValue(
            `${annotatedAssetId ? annotatedAssetId : ""} ${
              annotatedAssetId ? "/ " : ""
            }${annotatedUser ? annotatedUser : ""}`
          );
          reason.setDataValidation(reasons);
          reason.setValue(DEFAULT_REASON);
          borrowed.setValue(new Date());
          exceptioned.insertCheckboxes();
          returned.insertCheckboxes();
          exceptioned.setDataValidation(enforceCheckbox);
          returned.setDataValidation(enforceCheckbox);
          status.setValue(AR);
        }
      }
      if (currentRow >= LOG.getMaxRows()) {
        LOG.insertRowsAfter(currentRow, 2);
        LOG.getRange(
          currentRow + 1,
          1,
          LOG.getMaxRows() - currentRow + 1,
          LOG.getLastColumn()
        ).removeCheckboxes();
      }
      if (
        previousRow.getValues()[0][0] == "" &&
        previousRow.getValues()[0][4] == ""
      )
        removeAboveInvalidRows(LOG, previousRow.getRowIndex());
    } catch (e) {
      currentCell.setValue(
        `Couldn't find ${input}.Try again or manually check in G-Suites.`
      );
      throw `Couldn't find ${input}`;
    }
  }
};

const whenReturned = () => {
  if (
    ACTIVESHEET.getName() === LOG.getName() &&
    LOG.getActiveCell().getColumn() === 9
  ) {
    let currentCell = LOG.getActiveCell();
    let currentRow = currentCell.getRowIndex();
    let lastCol = LOG.getLastColumn();
    let currentRange = LOG.getRange(currentRow, 1, 1, lastCol);
    let email = currentCell.offset(0, -8);
    let tag = currentCell.offset(0, -5);
    let returnedDate = currentCell.offset(0, 1);
    let status = currentCell.offset(0, 2);
    let exceptioned = currentCell.offset(0, -1);
    if (currentCell.getColumn() === 9) {
      if (!email.isBlank()) {
        if (currentCell.isChecked() && !email.isBlank() && !tag.isBlank()) {
          currentRange.setBackground(SUCESS_COLOR);
          returnedDate.setValue(new Date());
          status.setValue(R);
          if (exceptioned.isChecked()) exceptioned.uncheck();
          enableChromebook(tag.getValue().toString());
        } else if (!currentCell.isChecked() && exceptioned.isChecked()) {
          status.setValue(DND);
          currentRange.setBackground(NEUTRAL_COLOR);
        } else if (!currentCell.isChecked() && !exceptioned.isChecked()) {
          status.setValue(AR);
          currentRange.setBackground(RESET_COLOR);
          returnedDate.clearContent();
        }
      } else {
        status.clearContent();
        returnedDate.clearContent();
        currentCell.removeCheckboxes();
      }
    }
  }
};

const whenExceptioned = () => {
  if (
    ACTIVESHEET.getName() === LOG.getName() &&
    LOG.getActiveCell().getColumn() === 8
  ) {
    let currentCell = LOG.getActiveCell();
    let currentRow = currentCell.getRowIndex();
    let lastCol = LOG.getLastColumn();
    let currentRange = LOG.getRange(currentRow, 1, 1, lastCol);
    let email = currentCell.offset(0, -7);
    let tag = currentCell.offset(0, -4);
    let returnedDate = currentCell.offset(0, 2);
    let status = currentCell.offset(0, 3);
    let returned = currentCell.offset(0, 1);
    if (currentCell.getColumn() === 8) {
      if (!email.isBlank()) {
        if (!email.isBlank() && currentCell.isChecked() && !tag.isBlank()) {
          if (returned.isChecked()) {
            SPREADSHEET.toast(
              "Can't exception a returned chromebook. Please check if the chromebook returned or uncheck the return checkbox to exception it.",
              "Can't Exception.",
              6
            );
            currentCell.uncheck();
          } else {
            currentRange.setBackground(NEUTRAL_COLOR);
            status.setValue(DND);
            returnedDate.clearContent();
            enableChromebook(tag.getValue().toString());
          }
        } else if (!currentCell.isChecked()) {
          currentRange.setBackground(RESET_COLOR);
          status.setValue(AR);
        }
      } else {
        status.clearContent();
        returnedDate.clearContent();
        currentCell.removeCheckboxes();
      }
    }
  }
};

const showSideBar = () => {
  let widget = HtmlService.createHtmlOutput(`<h1>Sidebar</h1>`);
  SpreadsheetApp.getUi().showSidebar(widget);
};
