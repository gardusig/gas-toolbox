import { getSheet } from "./spreadsheet";
import type {
  GenericObject,
  HeaderMap,
  SheetRow,
  SheetCellValue,
} from "./types";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function deepCopy<T extends Record<string, any> | any[]>(
  obj: T | null
): T | null {
  if (obj === null || typeof obj !== "object") {
    return obj;
  }
  // This path is defensive and not reachable through public API (HeaderMap is Record<string, number>)
  /* istanbul ignore if */
  if (Array.isArray(obj)) {
    // eslint-disable-next-line @typescript-eslint/no-unsafe-return
    return obj.map(element => deepCopy(element)) as T;
  }
  const newObj = {} as T;
  for (const key in obj) {
    if (Object.prototype.hasOwnProperty.call(obj, key)) {
      const value = (obj as Record<string, unknown>)[key];
      // eslint-disable-next-line @typescript-eslint/no-unsafe-member-access, @typescript-eslint/no-unsafe-assignment, @typescript-eslint/no-explicit-any
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment
      newObj[key] = deepCopy(
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        value as Record<string, any> | any[] | null
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
      ) as any;
    }
  }
  return newObj;
}

function createHeaderMap(sheetHeaderRow: SheetRow): HeaderMap {
  if (!sheetHeaderRow || sheetHeaderRow.length === 0) {
    throw new Error("Empty header row");
  }
  const headerMap: HeaderMap = {};
  sheetHeaderRow.forEach((sheetCellValue, columnIndex) => {
    if (typeof sheetCellValue !== "string") {
      throw new Error(`Unexpected column name type at index ${columnIndex}`);
    }
    headerMap[sheetCellValue] = columnIndex;
  });
  return headerMap;
}

function getSheetData(
  sheetName: string,
  spreadsheetIdOrURL?: string
): {
  sheetRows: SheetRow[];
  headerMap: HeaderMap;
} {
  if (!sheetName || typeof sheetName !== "string") {
    throw new Error("Sheet name must be a non-empty string");
  }
  const sheet = getSheet(sheetName, spreadsheetIdOrURL);
  if (sheet === null) {
    throw new Error(
      `Sheet '${sheetName}' not found${spreadsheetIdOrURL ? ` in spreadsheet '${spreadsheetIdOrURL}'` : ""}`
    );
  }
  const sheetRows = sheet.getDataRange().getValues() as SheetRow[];
  if (sheetRows.length < 1) {
    throw new Error("Empty sheet");
  }
  // eslint-disable-next-line @typescript-eslint/no-unsafe-return
  const headerMap = createHeaderMap(sheetRows[0]);
  return { sheetRows, headerMap };
}

function createObject(sheetRow: SheetRow, headerMap: HeaderMap): GenericObject {
  const genericObject: GenericObject = {};
  for (const [columnName, columnIndex] of Object.entries(headerMap)) {
    const cellContent = sheetRow[columnIndex];
    if (
      cellContent !== undefined &&
      cellContent !== null &&
      cellContent !== ""
    ) {
      genericObject[columnName] = cellContent;
    }
  }
  return genericObject;
}

function createObjectList(
  sheetRows: SheetRow[],
  headerMap: HeaderMap,
  startRowIndex?: number,
  finishRowIndex?: number
): GenericObject[] {
  startRowIndex = startRowIndex ?? Number.MIN_SAFE_INTEGER;
  finishRowIndex = finishRowIndex ?? Number.MAX_SAFE_INTEGER;
  const objectList: GenericObject[] = [];
  startRowIndex = Math.max(startRowIndex, 1);
  finishRowIndex = Math.min(finishRowIndex, sheetRows.length);
  for (let rowIndex = startRowIndex; rowIndex < finishRowIndex; rowIndex++) {
    const genericObject = createObject(sheetRows[rowIndex], headerMap);
    objectList.push(genericObject);
  }
  return objectList;
}

export function getAllObjects(
  sheetName: string,
  spreadsheetIdOrURL?: string
): GenericObject[] {
  const { sheetRows, headerMap } = getSheetData(sheetName, spreadsheetIdOrURL);
  return createObjectList(sheetRows, headerMap);
}

export function getObject(
  sheetName: string,
  rowIndex: number,
  spreadsheetIdOrURL?: string
): GenericObject | null {
  if (
    rowIndex === null ||
    rowIndex === undefined ||
    typeof rowIndex !== "number"
  ) {
    throw new Error("Row index must be a number");
  }
  if (rowIndex < 0) {
    throw new Error("Row index must be >= 0");
  }
  const { sheetRows, headerMap } = getSheetData(sheetName, spreadsheetIdOrURL);
  const objectList = createObjectList(
    sheetRows,
    headerMap,
    rowIndex + 1,
    rowIndex + 2
  );
  if (objectList.length === 0) {
    return null;
  }
  return objectList[0];
}

export function getObjectBatch(
  sheetName: string,
  startRowIndex: number,
  finishRowIndex: number,
  spreadsheetIdOrURL?: string
): GenericObject[] {
  if (
    startRowIndex === null ||
    startRowIndex === undefined ||
    typeof startRowIndex !== "number"
  ) {
    throw new Error("Start row index must be a number");
  }
  if (
    finishRowIndex === null ||
    finishRowIndex === undefined ||
    typeof finishRowIndex !== "number"
  ) {
    throw new Error("Finish row index must be a number");
  }
  if (startRowIndex < 0) {
    throw new Error("Start row index must be >= 0");
  }
  if (finishRowIndex < startRowIndex) {
    throw new Error("Finish row index must be >= start row index");
  }
  const { sheetRows, headerMap } = getSheetData(sheetName, spreadsheetIdOrURL);
  return createObjectList(
    sheetRows,
    headerMap,
    startRowIndex + 1,
    finishRowIndex + 1
  );
}

export function getHeaderMap(
  sheetName: string,
  spreadsheetIdOrURL?: string
): HeaderMap | null {
  const { headerMap } = getSheetData(sheetName, spreadsheetIdOrURL);
  return deepCopy(headerMap);
}

export function filterObjects(
  sheetName: string,
  predicate: (obj: GenericObject, rowIndex: number) => boolean,
  spreadsheetIdOrURL?: string
): GenericObject[] {
  if (!predicate || typeof predicate !== "function") {
    throw new Error("Predicate must be a function");
  }
  const { sheetRows, headerMap } = getSheetData(sheetName, spreadsheetIdOrURL);
  const objectList: GenericObject[] = [];
  for (let rowIndex = 1; rowIndex < sheetRows.length; rowIndex++) {
    const genericObject = createObject(sheetRows[rowIndex], headerMap);
    if (predicate(genericObject, rowIndex - 1)) {
      objectList.push(genericObject);
    }
  }
  return objectList;
}

export function findObject(
  sheetName: string,
  predicate: (obj: GenericObject, rowIndex: number) => boolean,
  spreadsheetIdOrURL?: string
): GenericObject | null {
  if (!predicate || typeof predicate !== "function") {
    throw new Error("Predicate must be a function");
  }
  const { sheetRows, headerMap } = getSheetData(sheetName, spreadsheetIdOrURL);
  for (let rowIndex = 1; rowIndex < sheetRows.length; rowIndex++) {
    const genericObject = createObject(sheetRows[rowIndex], headerMap);
    if (predicate(genericObject, rowIndex - 1)) {
      return genericObject;
    }
  }
  return null;
}

export function findObjectIndex(
  sheetName: string,
  predicate: (obj: GenericObject, rowIndex: number) => boolean,
  spreadsheetIdOrURL?: string
): number | null {
  if (!predicate || typeof predicate !== "function") {
    throw new Error("Predicate must be a function");
  }
  const { sheetRows, headerMap } = getSheetData(sheetName, spreadsheetIdOrURL);
  for (let rowIndex = 1; rowIndex < sheetRows.length; rowIndex++) {
    const genericObject = createObject(sheetRows[rowIndex], headerMap);
    if (predicate(genericObject, rowIndex - 1)) {
      return rowIndex - 1;
    }
  }
  return null;
}

export function countObjects(
  sheetName: string,
  spreadsheetIdOrURL?: string
): number {
  const { sheetRows } = getSheetData(sheetName, spreadsheetIdOrURL);
  return Math.max(0, sheetRows.length - 1);
}

export function getFirst(
  sheetName: string,
  spreadsheetIdOrURL?: string
): GenericObject | null {
  return getObject(sheetName, 0, spreadsheetIdOrURL);
}

export function getLast(
  sheetName: string,
  spreadsheetIdOrURL?: string
): GenericObject | null {
  const count = countObjects(sheetName, spreadsheetIdOrURL);
  if (count === 0) {
    return null;
  }
  return getObject(sheetName, count - 1, spreadsheetIdOrURL);
}

export function exists(
  sheetName: string,
  predicate: (obj: GenericObject, rowIndex: number) => boolean,
  spreadsheetIdOrURL?: string
): boolean {
  return findObject(sheetName, predicate, spreadsheetIdOrURL) !== null;
}

export function sortObjects(
  sheetName: string,
  sortBy: string | string[],
  ascending: boolean = true,
  spreadsheetIdOrURL?: string
): GenericObject[] {
  if (!sortBy) {
    throw new Error("Sort by must be a string or array of strings");
  }
  if (typeof sortBy !== "string" && !Array.isArray(sortBy)) {
    throw new Error("Sort by must be a string or array of strings");
  }
  const objects = getAllObjects(sheetName, spreadsheetIdOrURL);
  const sortColumns = Array.isArray(sortBy) ? sortBy : [sortBy];

  return objects.sort((a, b) => {
    for (const column of sortColumns) {
      const aVal = a[column];
      const bVal = b[column];

      if (aVal === undefined || aVal === null) {
        return ascending ? 1 : -1;
      }
      if (bVal === undefined || bVal === null) {
        return ascending ? -1 : 1;
      }

      let comparison = 0;
      if (typeof aVal === "number" && typeof bVal === "number") {
        comparison = aVal - bVal;
      } else if (aVal instanceof Date && bVal instanceof Date) {
        comparison = aVal.getTime() - bVal.getTime();
      } else {
        comparison = String(aVal).localeCompare(String(bVal));
      }

      if (comparison !== 0) {
        return ascending ? comparison : -comparison;
      }
    }
    return 0;
  });
}

export function getObjectsPaginated(
  sheetName: string,
  page: number,
  pageSize: number,
  spreadsheetIdOrURL?: string
): {
  data: GenericObject[];
  total: number;
  page: number;
  pageSize: number;
  totalPages: number;
} {
  if (page === null || page === undefined || typeof page !== "number") {
    throw new Error("Page must be a number");
  }
  if (page < 1) {
    throw new Error("Page must be >= 1");
  }
  if (
    pageSize === null ||
    pageSize === undefined ||
    typeof pageSize !== "number"
  ) {
    throw new Error("Page size must be a number");
  }
  if (pageSize < 1) {
    throw new Error("Page size must be >= 1");
  }

  const allObjects = getAllObjects(sheetName, spreadsheetIdOrURL);
  const total = allObjects.length;
  const totalPages = Math.ceil(total / pageSize);
  const startIndex = (page - 1) * pageSize;
  const endIndex = startIndex + pageSize;

  return {
    data: allObjects.slice(startIndex, endIndex),
    total,
    page,
    pageSize,
    totalPages,
  };
}

export function sum(
  sheetName: string,
  column: string,
  spreadsheetIdOrURL?: string
): number {
  if (!column || typeof column !== "string") {
    throw new Error("Column must be a non-empty string");
  }
  const objects = getAllObjects(sheetName, spreadsheetIdOrURL);
  return objects.reduce((sum, obj) => {
    const value = obj[column];
    if (typeof value === "number") {
      return sum + value;
    }
    return sum;
  }, 0);
}

export function average(
  sheetName: string,
  column: string,
  spreadsheetIdOrURL?: string
): number {
  if (!column || typeof column !== "string") {
    throw new Error("Column must be a non-empty string");
  }
  const objects = getAllObjects(sheetName, spreadsheetIdOrURL);
  const numericObjects = objects.filter(obj => typeof obj[column] === "number");
  if (numericObjects.length === 0) {
    return 0;
  }
  return sum(sheetName, column, spreadsheetIdOrURL) / numericObjects.length;
}

export function min(
  sheetName: string,
  column: string,
  spreadsheetIdOrURL?: string
): SheetCellValue | null {
  if (!column || typeof column !== "string") {
    throw new Error("Column must be a non-empty string");
  }
  const objects = getAllObjects(sheetName, spreadsheetIdOrURL);
  if (objects.length === 0) {
    return null;
  }

  const values = objects
    .map(obj => obj[column])
    .filter(val => val !== undefined && val !== null);

  if (values.length === 0) {
    return null;
  }

  if (typeof values[0] === "number") {
    return Math.min(...(values as number[]));
  }
  if (values[0] instanceof Date) {
    return new Date(Math.min(...(values as Date[]).map(d => d.getTime())));
  }
  return values.sort()[0];
}

export function max(
  sheetName: string,
  column: string,
  spreadsheetIdOrURL?: string
): SheetCellValue | null {
  if (!column || typeof column !== "string") {
    throw new Error("Column must be a non-empty string");
  }
  const objects = getAllObjects(sheetName, spreadsheetIdOrURL);
  if (objects.length === 0) {
    return null;
  }

  const values = objects
    .map(obj => obj[column])
    .filter(val => val !== undefined && val !== null);

  if (values.length === 0) {
    return null;
  }

  if (typeof values[0] === "number") {
    return Math.max(...(values as number[]));
  }
  if (values[0] instanceof Date) {
    return new Date(Math.max(...(values as Date[]).map(d => d.getTime())));
  }
  return values.sort().reverse()[0];
}

export function groupBy(
  sheetName: string,
  column: string,
  spreadsheetIdOrURL?: string
): Record<string, GenericObject[]> {
  if (!column || typeof column !== "string") {
    throw new Error("Column must be a non-empty string");
  }
  const objects = getAllObjects(sheetName, spreadsheetIdOrURL);
  const grouped: Record<string, GenericObject[]> = {};

  objects.forEach(obj => {
    const key = String(obj[column] ?? "");
    if (!grouped[key]) {
      grouped[key] = [];
    }
    grouped[key].push(obj);
  });

  return grouped;
}

export function getDistinctValues(
  sheetName: string,
  column: string,
  spreadsheetIdOrURL?: string
): SheetCellValue[] {
  if (!column || typeof column !== "string") {
    throw new Error("Column must be a non-empty string");
  }
  const objects = getAllObjects(sheetName, spreadsheetIdOrURL);
  const values = new Set<SheetCellValue>();

  objects.forEach(obj => {
    const value = obj[column];
    if (value !== undefined && value !== null && value !== "") {
      values.add(value);
    }
  });

  return Array.from(values);
}

export function filterByColumn(
  sheetName: string,
  column: string,
  value: SheetCellValue,
  spreadsheetIdOrURL?: string
): GenericObject[] {
  if (!column || typeof column !== "string") {
    throw new Error("Column must be a non-empty string");
  }
  return filterObjects(
    sheetName,
    obj => obj[column] === value,
    spreadsheetIdOrURL
  );
}
