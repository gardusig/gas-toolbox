function findOrCreateFolder(
  parentFolder: GoogleAppsScript.Drive.Folder,
  folderName: string
): GoogleAppsScript.Drive.Folder {
  const folderIterator = parentFolder.getFoldersByName(folderName);
  if (folderIterator.hasNext()) {
    const existingFolder = folderIterator.next();
    Logger.log(`Folder "${existingFolder.getName()}" found`);
    return existingFolder;
  }
  const newFolder = parentFolder.createFolder(folderName);
  Logger.log(`Folder "${newFolder.getName()}" created`);
  return newFolder;
}

export function getOrCreateFolderByPath(
  path: string
): GoogleAppsScript.Drive.Folder {
  if (path === null || path === undefined) {
    throw new Error("Path must be a non-empty string");
  }
  if (typeof path !== "string") {
    throw new Error("Path must be a non-empty string");
  }
  const parts = path.split("/").filter(part => part.trim() !== "");
  if (parts.length === 0) {
    return DriveApp.getRootFolder();
  }
  let currentFolder = DriveApp.getRootFolder();
  parts.forEach(part => {
    currentFolder = findOrCreateFolder(currentFolder, part.trim());
  });
  Logger.log(`Folder path "${path}" retrieved or created successfully`);
  return currentFolder;
}

export function getAllFoldersInFolder(
  folder: GoogleAppsScript.Drive.Folder
): GoogleAppsScript.Drive.Folder[] {
  if (!folder) {
    throw new Error("Folder is required");
  }
  const folders: GoogleAppsScript.Drive.Folder[] = [];
  const folderIterator = folder.getFolders();
  while (folderIterator.hasNext()) {
    folders.push(folderIterator.next());
  }
  Logger.log(
    `Found ${folders.length} folder(s) in folder "${folder.getName()}"`
  );
  return folders;
}

export function deleteFolder(folder: GoogleAppsScript.Drive.Folder): void {
  if (!folder) {
    throw new Error("Folder is required");
  }
  const folderName = folder.getName();
  DriveApp.removeFolder(folder);
  Logger.log(`Folder "${folderName}" deleted successfully`);
}

export function checkFolderExists(folderPath: string): boolean {
  if (!folderPath || typeof folderPath !== "string") {
    return false;
  }
  try {
    const folder = getOrCreateFolderByPath(folderPath);
    return folder !== null;
  } catch (error) {
    return false;
  }
}

export function getFolderById(
  folderId: string
): GoogleAppsScript.Drive.Folder | null {
  if (!folderId || typeof folderId !== "string") {
    return null;
  }
  try {
    const folder = DriveApp.getFolderById(folderId);
    return folder;
  } catch (error) {
    Logger.log(`Folder with ID "${folderId}" not found`);
    return null;
  }
}

export function getFolderByName(
  folderName: string,
  parentFolder: GoogleAppsScript.Drive.Folder
): GoogleAppsScript.Drive.Folder | null {
  if (!parentFolder) {
    throw new Error("Parent folder is required");
  }
  if (!folderName || typeof folderName !== "string") {
    return null;
  }
  const folderIterator = parentFolder.getFoldersByName(folderName);
  if (folderIterator.hasNext()) {
    const folder = folderIterator.next();
    Logger.log(
      `Folder "${folderName}" found in folder "${parentFolder.getName()}"`
    );
    return folder;
  }
  return null;
}

export function createFolder(
  folderName: string,
  parentFolder?: GoogleAppsScript.Drive.Folder
): GoogleAppsScript.Drive.Folder {
  if (!folderName || typeof folderName !== "string") {
    throw new Error("Folder name must be a non-empty string");
  }
  const targetFolder = parentFolder || DriveApp.getRootFolder();
  const newFolder = targetFolder.createFolder(folderName);
  Logger.log(
    `Folder "${folderName}" created in folder "${targetFolder.getName()}"`
  );
  return newFolder;
}
