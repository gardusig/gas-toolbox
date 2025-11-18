import {
  getOrCreateFolderByPath,
  getFileByName,
  findFile,
} from "../drive";

export function ensureFolder(
  folderPath: string,
): GoogleAppsScript.Drive.Folder {
  if (!folderPath || typeof folderPath !== "string") {
    throw new Error("Folder path must be a non-empty string");
  }
  return getOrCreateFolderByPath(folderPath);
}

export function createDocument(
  folderPath: string,
  docName: string,
): GoogleAppsScript.Document.Document {
  if (!folderPath || typeof folderPath !== "string") {
    throw new Error("Folder path must be a non-empty string");
  }
  if (!docName || typeof docName !== "string") {
    throw new Error("Document name must be a non-empty string");
  }
  const targetFolder = ensureFolder(folderPath);
  const existingFile = getFileByName(docName, targetFolder);

  if (existingFile) {
    Logger.log(
      `Document "${docName}" already exists in folder: "${targetFolder.getName()}"`,
    );
    return DocumentApp.openById(existingFile.getId());
  }

  const doc = DocumentApp.create(docName);
  const file = DriveApp.getFileById(doc.getId());
  file.moveTo(targetFolder);
  Logger.log(
    `Document created successfully:` +
      `\n\tURL: ${doc.getUrl()}` +
      `\n\tName: ${doc.getName()}` +
      `\n\tLocation: ${targetFolder.getName()}`,
  );
  return doc;
}

export function clearDocument(
  folderPath: string,
  fileName: string,
): void {
  if (!folderPath || typeof folderPath !== "string") {
    throw new Error("Folder path must be a non-empty string");
  }
  if (!fileName || typeof fileName !== "string") {
    throw new Error("File name must be a non-empty string");
  }
  const file = findFile(folderPath, fileName);
  const doc = DocumentApp.openById(file.getId());
  const body = doc.getBody();
  body.clear();
  doc.saveAndClose();
  Logger.log(`Document "${fileName}" in folder "${folderPath}" cleared`);
}

export function getDocumentContent(
  folderPath: string,
  fileName: string,
): string {
  if (!folderPath || typeof folderPath !== "string") {
    throw new Error("Folder path must be a non-empty string");
  }
  if (!fileName || typeof fileName !== "string") {
    throw new Error("File name must be a non-empty string");
  }
  const file = findFile(folderPath, fileName);
  const doc = DocumentApp.openById(file.getId());
  const body = doc.getBody();
  const content = body.getText();
  doc.saveAndClose();
  Logger.log(`Retrieved content from document "${fileName}" in folder "${folderPath}"`);
  return content;
}

