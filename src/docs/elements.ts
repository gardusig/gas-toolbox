import { findFile } from "../drive";

export function insertTable(
  folderPath: string,
  fileName: string,
  rows: number,
  columns: number,
  cellValues?: string[][]
): GoogleAppsScript.Document.Table {
  if (!folderPath || typeof folderPath !== "string") {
    throw new Error("Folder path must be a non-empty string");
  }
  if (!fileName || typeof fileName !== "string") {
    throw new Error("File name must be a non-empty string");
  }
  if (rows === null || rows === undefined || typeof rows !== "number") {
    throw new Error("Rows must be a number");
  }
  if (rows < 1) {
    throw new Error("Rows must be >= 1");
  }
  if (
    columns === null ||
    columns === undefined ||
    typeof columns !== "number"
  ) {
    throw new Error("Columns must be a number");
  }
  if (columns < 1) {
    throw new Error("Columns must be >= 1");
  }
  if (
    cellValues !== undefined &&
    (!Array.isArray(cellValues) || !Array.isArray(cellValues[0]))
  ) {
    throw new Error("Cell values must be a 2D array");
  }
  const file = findFile(folderPath, fileName);
  const doc = DocumentApp.openById(file.getId());
  const body = doc.getBody();

  const table = body.appendTable();

  // Create rows and columns
  for (let i = 0; i < rows; i += 1) {
    const row = table.appendTableRow();
    for (let j = 0; j < columns; j += 1) {
      const cell = row.appendTableCell();
      if (cellValues && cellValues[i] && cellValues[i][j] !== undefined) {
        cell.setText(cellValues[i][j]);
      }
    }
  }

  doc.saveAndClose();
  Logger.log(
    `Table with ${rows}x${columns} inserted into document "${fileName}" in folder "${folderPath}"`
  );

  return table;
}

export function insertImage(
  folderPath: string,
  fileName: string,
  imageBlob: GoogleAppsScript.Base.Blob,
  width?: number,
  height?: number
): GoogleAppsScript.Document.InlineImage {
  if (!folderPath || typeof folderPath !== "string") {
    throw new Error("Folder path must be a non-empty string");
  }
  if (!fileName || typeof fileName !== "string") {
    throw new Error("File name must be a non-empty string");
  }
  if (!imageBlob) {
    throw new Error("Image blob is required");
  }
  const file = findFile(folderPath, fileName);
  const doc = DocumentApp.openById(file.getId());
  const body = doc.getBody();

  const image = body.appendImage(imageBlob);

  if (
    width !== undefined &&
    width !== null &&
    typeof width === "number" &&
    width > 0
  ) {
    image.setWidth(width);
  }
  if (
    height !== undefined &&
    height !== null &&
    typeof height === "number" &&
    height > 0
  ) {
    image.setHeight(height);
  }

  doc.saveAndClose();
  Logger.log(
    `Image inserted into document "${fileName}" in folder "${folderPath}"`
  );

  return image;
}
