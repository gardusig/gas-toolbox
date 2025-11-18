import { findFile } from "../drive";
import { formatParagraph } from "./formatting";

export function appendParagraphToFile(
  folderPath: string,
  fileName: string,
  content: string,
  heading?: GoogleAppsScript.Document.ParagraphHeading,
): GoogleAppsScript.Document.Paragraph {
  if (!folderPath || typeof folderPath !== "string") {
    throw new Error("Folder path must be a non-empty string");
  }
  if (!fileName || typeof fileName !== "string") {
    throw new Error("File name must be a non-empty string");
  }
  if (content === null || content === undefined || typeof content !== "string") {
    throw new Error("Content must be a string");
  }
  const file = findFile(folderPath, fileName);
  const doc = DocumentApp.openById(file.getId());
  const body = doc.getBody();

  // Check if the last paragraph is empty - if so, replace it instead of appending
  let paragraph: GoogleAppsScript.Document.Paragraph;
  const numChildren = body.getNumChildren();

  if (numChildren > 0) {
    const lastChild = body.getChild(numChildren - 1);
    if (lastChild.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const lastParagraph = lastChild.asParagraph();
      // If last paragraph is empty, replace it instead of appending
      if (lastParagraph.getText().trim() === "") {
        paragraph = lastParagraph;
        paragraph.setText(content);
      } else {
        // Append new paragraph
        paragraph = body.appendParagraph(content);
      }
    } else {
      // Last child is not a paragraph, append new one
      paragraph = body.appendParagraph(content);
    }
  } else {
    // Document is empty, insert at start
    paragraph = body.insertParagraph(0, content);
  }

  if (heading !== undefined) {
    paragraph = paragraph.setHeading(heading);
  }

  formatParagraph(paragraph);
  doc.saveAndClose();
  Logger.log(
    `Content appended to document "${fileName}" in folder "${folderPath}"`,
  );

  return paragraph;
}

export function insertParagraphAtPosition(
  folderPath: string,
  fileName: string,
  content: string,
  position: number,
  heading?: GoogleAppsScript.Document.ParagraphHeading,
): GoogleAppsScript.Document.Paragraph {
  if (!folderPath || typeof folderPath !== "string") {
    throw new Error("Folder path must be a non-empty string");
  }
  if (!fileName || typeof fileName !== "string") {
    throw new Error("File name must be a non-empty string");
  }
  if (content === null || content === undefined || typeof content !== "string") {
    throw new Error("Content must be a string");
  }
  if (position === null || position === undefined || typeof position !== "number") {
    throw new Error("Position must be a number");
  }
  if (position < 0) {
    throw new Error("Position must be >= 0");
  }
  const file = findFile(folderPath, fileName);
  const doc = DocumentApp.openById(file.getId());
  const body = doc.getBody();
  const numChildren = body.getNumChildren();
  
  // Clamp position to valid range
  const insertIndex = Math.min(position, numChildren);
  
  const paragraph = body.insertParagraph(insertIndex, content);
  
  if (heading !== undefined) {
    paragraph.setHeading(heading);
  }

  formatParagraph(paragraph);
  doc.saveAndClose();
  Logger.log(
    `Paragraph inserted at position ${insertIndex} in document "${fileName}" in folder "${folderPath}"`,
  );

  return paragraph;
}

export function deleteParagraph(
  folderPath: string,
  fileName: string,
  position: number,
): void {
  if (!folderPath || typeof folderPath !== "string") {
    throw new Error("Folder path must be a non-empty string");
  }
  if (!fileName || typeof fileName !== "string") {
    throw new Error("File name must be a non-empty string");
  }
  if (position === null || position === undefined || typeof position !== "number") {
    throw new Error("Position must be a number");
  }
  if (position < 0) {
    throw new Error("Position must be >= 0");
  }
  const file = findFile(folderPath, fileName);
  const doc = DocumentApp.openById(file.getId());
  const body = doc.getBody();
  const numChildren = body.getNumChildren();
  
  if (position >= numChildren) {
    throw new Error(`Position ${position} is out of bounds (document has ${numChildren} elements)`);
  }
  
  const child = body.getChild(position);
  if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
    body.removeChild(child);
    doc.saveAndClose();
    Logger.log(`Paragraph at position ${position} deleted from document "${fileName}"`);
  } else {
    doc.saveAndClose();
    throw new Error(`Element at position ${position} is not a paragraph`);
  }
}

export function getParagraphAtPosition(
  folderPath: string,
  fileName: string,
  position: number,
): GoogleAppsScript.Document.Paragraph | null {
  if (!folderPath || typeof folderPath !== "string") {
    throw new Error("Folder path must be a non-empty string");
  }
  if (!fileName || typeof fileName !== "string") {
    throw new Error("File name must be a non-empty string");
  }
  if (position === null || position === undefined || typeof position !== "number") {
    throw new Error("Position must be a number");
  }
  if (position < 0) {
    throw new Error("Position must be >= 0");
  }
  const file = findFile(folderPath, fileName);
  const doc = DocumentApp.openById(file.getId());
  const body = doc.getBody();
  const numChildren = body.getNumChildren();
  
  if (position >= numChildren) {
    doc.saveAndClose();
    return null;
  }
  
  const child = body.getChild(position);
  if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
    const paragraph = child.asParagraph();
    doc.saveAndClose();
    return paragraph;
  }
  
  doc.saveAndClose();
  return null;
}

export function getParagraphCount(
  folderPath: string,
  fileName: string,
): number {
  if (!folderPath || typeof folderPath !== "string") {
    throw new Error("Folder path must be a non-empty string");
  }
  if (!fileName || typeof fileName !== "string") {
    throw new Error("File name must be a non-empty string");
  }
  const file = findFile(folderPath, fileName);
  const doc = DocumentApp.openById(file.getId());
  const body = doc.getBody();
  let count = 0;
  const numChildren = body.getNumChildren();
  for (let i = 0; i < numChildren; i += 1) {
    const child = body.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      count += 1;
    }
  }
  doc.saveAndClose();
  Logger.log(`Document "${fileName}" has ${count} paragraph(s)`);
  return count;
}

