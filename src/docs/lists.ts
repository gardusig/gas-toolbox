import { findFile } from "../drive";
import { formatParagraph } from "./formatting";

export function appendBulletedListToFile(
  folderPath: string,
  fileName: string,
  items: string[]
): GoogleAppsScript.Document.ListItem[] {
  if (!folderPath || typeof folderPath !== "string") {
    throw new Error("Folder path must be a non-empty string");
  }
  if (!fileName || typeof fileName !== "string") {
    throw new Error("File name must be a non-empty string");
  }
  if (!items || !Array.isArray(items)) {
    throw new Error("Items must be an array");
  }
  if (items.length === 0) {
    Logger.log(
      `No items provided for bulleted list in document "${fileName}". Skipping.`
    );
    return [];
  }

  const file = findFile(folderPath, fileName);
  const doc = DocumentApp.openById(file.getId());
  const body = doc.getBody();
  const listItems: GoogleAppsScript.Document.ListItem[] = [];

  items.forEach(item => {
    if (typeof item !== "string") {
      Logger.log(`Skipping non-string item: ${String(item)}`);
      return;
    }
    const listItem = body.appendListItem(item);
    listItem.setGlyphType(DocumentApp.GlyphType.BULLET);
    formatParagraph(listItem);
    listItems.push(listItem);
  });

  doc.saveAndClose();
  Logger.log(
    `Bulleted list appended to document "${fileName}" in folder "${folderPath}"`
  );

  return listItems;
}

export function appendNumberedListToFile(
  folderPath: string,
  fileName: string,
  items: string[]
): GoogleAppsScript.Document.ListItem[] {
  if (!folderPath || typeof folderPath !== "string") {
    throw new Error("Folder path must be a non-empty string");
  }
  if (!fileName || typeof fileName !== "string") {
    throw new Error("File name must be a non-empty string");
  }
  if (!items || !Array.isArray(items)) {
    throw new Error("Items must be an array");
  }
  if (items.length === 0) {
    Logger.log(
      `No items provided for numbered list in document "${fileName}". Skipping.`
    );
    return [];
  }

  const file = findFile(folderPath, fileName);
  const doc = DocumentApp.openById(file.getId());
  const body = doc.getBody();
  const listItems: GoogleAppsScript.Document.ListItem[] = [];

  items.forEach(item => {
    if (typeof item !== "string") {
      Logger.log(`Skipping non-string item: ${String(item)}`);
      return;
    }
    const listItem = body.appendListItem(item);
    listItem.setGlyphType(DocumentApp.GlyphType.NUMBER);
    formatParagraph(listItem);
    listItems.push(listItem);
  });

  doc.saveAndClose();
  Logger.log(
    `Numbered list appended to document "${fileName}" in folder "${folderPath}"`
  );

  return listItems;
}
