import { getOrCreateFolderByPath } from "./folders";

export function getFileByName(
  fileName: string,
  targetFolder: GoogleAppsScript.Drive.Folder,
): GoogleAppsScript.Drive.File | null {
  if (!targetFolder) {
    throw new Error("Target folder is required");
  }
  if (!fileName || typeof fileName !== "string") {
    return null;
  }
  const files = targetFolder.getFilesByName(fileName);
  if (files.hasNext()) {
    const existingFile = files.next();
    Logger.log(`Document already exists: ${existingFile.getUrl()}`);
    return existingFile;
  }
  return null;
}

export function findFile(
  folderPath: string,
  fileName: string,
): GoogleAppsScript.Drive.File {
  if (!folderPath || typeof folderPath !== "string") {
    throw new Error("Folder path must be a non-empty string");
  }
  if (!fileName || typeof fileName !== "string") {
    throw new Error("File name must be a non-empty string");
  }
  const targetFolder = getOrCreateFolderByPath(folderPath);
  const existingFile = getFileByName(fileName, targetFolder);
  if (!existingFile) {
    throw new Error(
      `Document "${fileName}" not found in folder: "${folderPath}"`,
    );
  }
  Logger.log(`Document "${fileName}" found in folder "${folderPath}"`);
  return existingFile;
}

export function getAllFilesInFolder(
  folder: GoogleAppsScript.Drive.Folder,
): GoogleAppsScript.Drive.File[] {
  if (!folder) {
    throw new Error("Folder is required");
  }
  const files: GoogleAppsScript.Drive.File[] = [];
  const fileIterator = folder.getFiles();
  while (fileIterator.hasNext()) {
    files.push(fileIterator.next());
  }
  Logger.log(`Found ${files.length} file(s) in folder "${folder.getName()}"`);
  return files;
}

export function deleteFile(
  file: GoogleAppsScript.Drive.File,
): void {
  if (!file) {
    throw new Error("File is required");
  }
  const fileName = file.getName();
  DriveApp.removeFile(file);
  Logger.log(`File "${fileName}" deleted successfully`);
}

export function checkFileExists(
  folderPath: string,
  fileName: string,
): boolean {
  if (!folderPath || typeof folderPath !== "string") {
    return false;
  }
  if (!fileName || typeof fileName !== "string") {
    return false;
  }
  try {
    const file = findFile(folderPath, fileName);
    return file !== null;
  } catch (error) {
    return false;
  }
}

export function getFileById(
  fileId: string,
): GoogleAppsScript.Drive.File | null {
  if (!fileId || typeof fileId !== "string") {
    return null;
  }
  try {
    const file = DriveApp.getFileById(fileId);
    return file;
  } catch (error) {
    Logger.log(`File with ID "${fileId}" not found`);
    return null;
  }
}

export function getFilesByType(
  folder: GoogleAppsScript.Drive.Folder,
  mimeType: string,
): GoogleAppsScript.Drive.File[] {
  if (!folder) {
    throw new Error("Folder is required");
  }
  if (!mimeType || typeof mimeType !== "string") {
    throw new Error("MIME type must be a non-empty string");
  }
  const files: GoogleAppsScript.Drive.File[] = [];
  const fileIterator = folder.getFilesByType(mimeType);
  while (fileIterator.hasNext()) {
    files.push(fileIterator.next());
  }
  Logger.log(`Found ${files.length} file(s) of type "${mimeType}" in folder "${folder.getName()}"`);
  return files;
}

