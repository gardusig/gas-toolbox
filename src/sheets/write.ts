import { getSpreadsheet, createSheet, getSheet } from "./spreadsheet";
import type { GenericObject, HeaderMap } from "./types";

function getSerializedObject(obj: GenericObject, header: string[]): any[] {
  const serializedObject: any[] = [];
  header.forEach((key) => {
    serializedObject.push(obj[key]);
  });
  return serializedObject;
}

function getSheetWithHeader(
  sheetName: string,
  header: string[],
  spreadsheetIdOrURL?: string,
  preserveData: boolean = false,
): { sheet: GoogleAppsScript.Spreadsheet.Sheet; header: string[] } {
  if (!sheetName || typeof sheetName !== "string") {
    throw new Error("Sheet name must be a non-empty string");
  }
  if (!header || !Array.isArray(header)) {
    throw new Error("Header must be an array");
  }
  if (header.length === 0) {
    throw new Error("Header must not be empty");
  }
  const spreadsheet = getSpreadsheet(spreadsheetIdOrURL);
  if (spreadsheet === null) {
    throw new Error(
      `Failed to find spreadsheet${spreadsheetIdOrURL ? ` for ID or URL '${spreadsheetIdOrURL}'` : ""}`,
    );
  }
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    sheet.appendRow(header);
  } else if (!preserveData) {
    // Only clear if not preserving data
    sheet.clear();
    sheet.appendRow(header);
  } else {
    // Check if header matches
    const lastColumn = sheet.getLastColumn();
    if (lastColumn === 0 || lastColumn < header.length) {
      // Sheet is empty or doesn't have enough columns, set header
      if (lastColumn === 0) {
        sheet.appendRow(header);
      } else {
        sheet.clear();
        sheet.appendRow(header);
      }
    } else {
      try {
        const existingHeader = sheet.getRange(1, 1, 1, header.length).getValues()[0];
        const headerMatches = existingHeader.length === header.length &&
          existingHeader.every((val, i) => val === header[i]);
        if (!headerMatches) {
          // Header doesn't match, need to recreate
          sheet.clear();
          sheet.appendRow(header);
        }
      } catch (error) {
        // If we can't read the header, clear and set it
        sheet.clear();
        sheet.appendRow(header);
      }
    }
  }
  Logger.log(`Sheet "${sheetName}" ready with header`);
  return { sheet, header };
}

export function appendObject(
  sheetName: string,
  header: string[],
  obj: GenericObject,
  spreadsheetIdOrURL?: string,
): void {
  if (!obj || typeof obj !== "object") {
    throw new Error("Object must be a valid object");
  }
  const { sheet, header: sheetHeader } = getSheetWithHeader(
    sheetName,
    header,
    spreadsheetIdOrURL,
  );
  const serializedObject = getSerializedObject(obj, sheetHeader);
  sheet.appendRow(serializedObject);
  Logger.log(`Object appended to sheet "${sheetName}"`);
}

export function appendObjects(
  sheetName: string,
  header: string[],
  objs: GenericObject[],
  spreadsheetIdOrURL?: string,
): void {
  if (!objs || !Array.isArray(objs)) {
    throw new Error("Objects must be an array");
  }
  objs.forEach((obj) => {
    appendObject(sheetName, header, obj, spreadsheetIdOrURL);
  });
}

export function updateObject(
  sheetName: string,
  header: string[],
  rowIndex: number,
  obj: GenericObject,
  spreadsheetIdOrURL?: string,
): void {
  if (rowIndex === null || rowIndex === undefined || typeof rowIndex !== "number") {
    throw new Error("Row index must be a number");
  }
  if (rowIndex < 0) {
    throw new Error(
      "Row index must be >= 0 (0-based indexing, 0 = first data row)",
    );
  }
  if (!obj || typeof obj !== "object") {
    throw new Error("Object must be a valid object");
  }
  const { sheet, header: sheetHeader } = getSheetWithHeader(
    sheetName,
    header,
    spreadsheetIdOrURL,
  );
  const serializedObject = getSerializedObject(obj, sheetHeader);
  // rowIndex 0 = first data row = row 2 in sheet (row 1 is header)
  const range = sheet.getRange(rowIndex + 2, 1, 1, sheetHeader.length);
  range.setValues([serializedObject]);
  Logger.log(`Object at row ${rowIndex} updated in sheet "${sheetName}"`);
}

export function updateObjects(
  sheetName: string,
  header: string[],
  updates: Array<{ rowIndex: number; obj: GenericObject }>,
  spreadsheetIdOrURL?: string,
): void {
  if (!updates || !Array.isArray(updates)) {
    throw new Error("Updates must be an array");
  }
  updates.forEach(({ rowIndex, obj }) => {
    updateObject(sheetName, header, rowIndex, obj, spreadsheetIdOrURL);
  });
}

export function deleteObject(
  sheetName: string,
  header: string[],
  rowIndex: number,
  spreadsheetIdOrURL?: string,
): void {
  if (rowIndex === null || rowIndex === undefined || typeof rowIndex !== "number") {
    throw new Error("Row index must be a number");
  }
  if (rowIndex < 0) {
    throw new Error(
      "Row index must be >= 0 (0-based indexing, 0 = first data row)",
    );
  }
  const { sheet } = getSheetWithHeader(sheetName, header, spreadsheetIdOrURL);
  // rowIndex 0 = first data row = row 2 in sheet (row 1 is header)
  sheet.deleteRow(rowIndex + 2);
  Logger.log(`Row ${rowIndex} deleted from sheet "${sheetName}"`);
}

export function deleteObjects(
  sheetName: string,
  header: string[],
  rowIndices: number[],
  spreadsheetIdOrURL?: string,
): void {
  if (!rowIndices || !Array.isArray(rowIndices)) {
    throw new Error("Row indices must be an array");
  }
  const sortedIndices = [...rowIndices].sort((a, b) => b - a);
  sortedIndices.forEach((rowIndex) => {
    deleteObject(sheetName, header, rowIndex, spreadsheetIdOrURL);
  });
}

export function deleteObjectsByFilter(
  sheetName: string,
  header: string[],
  predicate: (obj: GenericObject, rowIndex: number) => boolean,
  spreadsheetIdOrURL?: string,
): number {
  if (!predicate || typeof predicate !== "function") {
    throw new Error("Predicate must be a function");
  }
  const { sheet } = getSheetWithHeader(sheetName, header, spreadsheetIdOrURL, true);
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  if (values.length < 2) {
    return 0; // Only header row or empty
  }

  const indicesToDelete: number[] = [];
  const headerRow = values[0];

  // Create header map
  const headerMap: HeaderMap = {};
  headerRow.forEach((headerValue, index) => {
    if (typeof headerValue === "string") {
      headerMap[headerValue] = index;
    }
  });

  // Check each data row
  for (let rowIndex = 1; rowIndex < values.length; rowIndex++) {
    const row = values[rowIndex];
    const obj: GenericObject = {};

    for (const [columnName, columnIndex] of Object.entries(headerMap)) {
      const cellContent = row[columnIndex];
      if (
        cellContent !== undefined &&
        cellContent !== null &&
        cellContent !== ""
      ) {
        obj[columnName] = cellContent;
      }
    }

    if (predicate(obj, rowIndex - 1)) {
      indicesToDelete.push(rowIndex - 1);
    }
  }

  deleteObjects(sheetName, header, indicesToDelete, spreadsheetIdOrURL);
  Logger.log(`Deleted ${indicesToDelete.length} object(s) from sheet "${sheetName}"`);
  return indicesToDelete.length;
}

export function clearAll(
  sheetName: string,
  header: string[],
  spreadsheetIdOrURL?: string,
): void {
  const { sheet } = getSheetWithHeader(sheetName, header, spreadsheetIdOrURL);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
  Logger.log(`All data cleared from sheet "${sheetName}"`);
}

export function upsertObject(
  sheetName: string,
  header: string[],
  keyColumn: string,
  obj: GenericObject,
  spreadsheetIdOrURL?: string,
): number {
  if (!keyColumn || typeof keyColumn !== "string") {
    throw new Error("Key column must be a non-empty string");
  }
  if (!header.includes(keyColumn)) {
    throw new Error(`Key column '${keyColumn}' not found in header`);
  }
  if (!obj || typeof obj !== "object") {
    throw new Error("Object must be a valid object");
  }
  const keyValue = obj[keyColumn];
  if (keyValue === undefined || keyValue === null) {
    throw new Error(`Key value is required for column '${keyColumn}'`);
  }

  const { sheet, header: sheetHeader } = getSheetWithHeader(
    sheetName,
    header,
    spreadsheetIdOrURL,
    true, // Preserve existing data
  );
  const keyColumnIndex = sheetHeader.indexOf(keyColumn);
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  // Check if object with this key already exists
  for (let rowIndex = 1; rowIndex < values.length; rowIndex++) {
    if (values[rowIndex][keyColumnIndex] === keyValue) {
      // Update existing row
      updateObject(sheetName, header, rowIndex - 1, obj, spreadsheetIdOrURL);
      return rowIndex - 1;
    }
  }

  // Append new row
  appendObject(sheetName, header, obj, spreadsheetIdOrURL);
  return values.length - 1; // Return the new row index (0-based)
}

export function upsertObjects(
  sheetName: string,
  header: string[],
  keyColumn: string,
  objs: GenericObject[],
  spreadsheetIdOrURL?: string,
): number {
  if (!keyColumn || typeof keyColumn !== "string") {
    throw new Error("Key column must be a non-empty string");
  }
  if (!objs || !Array.isArray(objs)) {
    throw new Error("Objects must be an array");
  }
  let upsertedCount = 0;
  objs.forEach((obj) => {
    upsertObject(sheetName, header, keyColumn, obj, spreadsheetIdOrURL);
    upsertedCount++;
  });
  Logger.log(`Upserted ${upsertedCount} object(s) to sheet "${sheetName}"`);
  return upsertedCount;
}

export function replaceAll(
  sheetName: string,
  header: string[],
  objs: GenericObject[],
  spreadsheetIdOrURL?: string,
): void {
  if (!objs || !Array.isArray(objs)) {
    throw new Error("Objects must be an array");
  }
  clearAll(sheetName, header, spreadsheetIdOrURL);
  if (objs.length > 0) {
    appendObjects(sheetName, header, objs, spreadsheetIdOrURL);
  }
  Logger.log(`Replaced all data in sheet "${sheetName}" with ${objs.length} object(s)`);
}

