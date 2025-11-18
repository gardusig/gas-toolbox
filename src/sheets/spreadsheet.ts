function isSpreadsheetUrl(value: string): boolean {
  return value.includes("spreadsheets.google.com");
}

export function getSpreadsheet(
  spreadsheetIdOrURL?: string
): GoogleAppsScript.Spreadsheet.Spreadsheet | null {
  if (spreadsheetIdOrURL === null || spreadsheetIdOrURL === undefined) {
    return SpreadsheetApp.getActiveSpreadsheet();
  }
  if (typeof spreadsheetIdOrURL !== "string") {
    throw new Error("Spreadsheet ID or URL must be a string");
  }
  if (spreadsheetIdOrURL.trim() === "") {
    return SpreadsheetApp.getActiveSpreadsheet();
  }
  try {
    if (isSpreadsheetUrl(spreadsheetIdOrURL)) {
      return SpreadsheetApp.openByUrl(spreadsheetIdOrURL);
    }
    return SpreadsheetApp.openById(spreadsheetIdOrURL);
  } catch (error) {
    Logger.log(`Failed to open spreadsheet: ${(error as Error).message}`);
    return null;
  }
}

export function createSheet(
  sheetName: string,
  header: string[],
  spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet
): GoogleAppsScript.Spreadsheet.Sheet {
  if (!sheetName || typeof sheetName !== "string") {
    throw new Error("Sheet name must be a non-empty string");
  }
  if (!header || !Array.isArray(header)) {
    throw new Error("Header must be an array");
  }
  if (header.length === 0) {
    throw new Error("Header must not be empty");
  }
  if (!spreadsheet) {
    throw new Error("Spreadsheet is required");
  }
  const sheet =
    spreadsheet.getSheetByName(sheetName) ?? spreadsheet.insertSheet(sheetName);
  sheet.clear();
  sheet.appendRow(header);
  Logger.log(`Sheet "${sheetName}" created or updated with header`);
  return sheet;
}

export function getSheet(
  sheetName: string,
  spreadsheetIdOrURL?: string
): GoogleAppsScript.Spreadsheet.Sheet | null {
  if (!sheetName || typeof sheetName !== "string") {
    throw new Error("Sheet name must be a non-empty string");
  }
  const spreadsheet = getSpreadsheet(spreadsheetIdOrURL);
  if (spreadsheet === null) {
    return null;
  }
  const sheet = spreadsheet.getSheetByName(sheetName);
  return sheet;
}
