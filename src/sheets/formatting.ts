import { getSheet } from "./spreadsheet";

export function trim(sheetName: string, spreadsheetIdOrURL?: string): void {
  trimColumns(sheetName, spreadsheetIdOrURL);
  trimRows(sheetName, spreadsheetIdOrURL);
}

export function trimRows(sheetName: string, spreadsheetIdOrURL?: string): void {
  if (!sheetName || typeof sheetName !== "string") {
    throw new Error("Sheet name must be a non-empty string");
  }
  const sheet = getSheet(sheetName, spreadsheetIdOrURL);
  if (sheet === null) {
    throw new Error(
      `Sheet '${sheetName}' not found${spreadsheetIdOrURL ? ` in spreadsheet '${spreadsheetIdOrURL}'` : ""}`
    );
  }
  const lastRowWithData = sheet.getLastRow();
  const maxRows = sheet.getMaxRows();
  const numRowsToRemove = maxRows - lastRowWithData;
  if (numRowsToRemove > 0) {
    sheet.deleteRows(lastRowWithData + 1, numRowsToRemove);
  }
  Logger.log(
    `Trimmed ${numRowsToRemove} empty row(s) from sheet "${sheetName}"`
  );
}

export function trimColumns(
  sheetName: string,
  spreadsheetIdOrURL?: string
): void {
  if (!sheetName || typeof sheetName !== "string") {
    throw new Error("Sheet name must be a non-empty string");
  }
  const sheet = getSheet(sheetName, spreadsheetIdOrURL);
  if (sheet === null) {
    throw new Error(
      `Sheet '${sheetName}' not found${spreadsheetIdOrURL ? ` in spreadsheet '${spreadsheetIdOrURL}'` : ""}`
    );
  }
  const lastColumnWithData = sheet.getLastColumn();
  const maxColumns = sheet.getMaxColumns();
  const columnsToRemove = maxColumns - lastColumnWithData;
  if (columnsToRemove > 0) {
    sheet.deleteColumns(lastColumnWithData + 1, columnsToRemove);
  }
  Logger.log(
    `Trimmed ${columnsToRemove} empty column(s) from sheet "${sheetName}"`
  );
}
