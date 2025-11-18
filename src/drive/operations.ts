export function copyFile(
  file: GoogleAppsScript.Drive.File,
  destinationFolder: GoogleAppsScript.Drive.Folder,
  newName?: string
): GoogleAppsScript.Drive.File {
  if (!file) {
    throw new Error("File is required");
  }
  if (!destinationFolder) {
    throw new Error("Destination folder is required");
  }
  const fileName = newName || file.getName();
  const copiedFile = file.makeCopy(fileName, destinationFolder);
  Logger.log(
    `File "${file.getName()}" copied to folder "${destinationFolder.getName()}" as "${fileName}"`
  );
  return copiedFile;
}

export function moveFile(
  file: GoogleAppsScript.Drive.File,
  destinationFolder: GoogleAppsScript.Drive.Folder
): void {
  if (!file) {
    throw new Error("File is required");
  }
  if (!destinationFolder) {
    throw new Error("Destination folder is required");
  }
  const fileName = file.getName();
  file.moveTo(destinationFolder);
  Logger.log(
    `File "${fileName}" moved to folder "${destinationFolder.getName()}"`
  );
}

export function renameFile(
  file: GoogleAppsScript.Drive.File,
  newName: string
): void {
  if (!file) {
    throw new Error("File is required");
  }
  if (!newName || typeof newName !== "string") {
    throw new Error("New file name must be a non-empty string");
  }
  const oldName = file.getName();
  file.setName(newName);
  Logger.log(`File renamed from "${oldName}" to "${newName}"`);
}

export function renameFolder(
  folder: GoogleAppsScript.Drive.Folder,
  newName: string
): void {
  if (!folder) {
    throw new Error("Folder is required");
  }
  if (!newName || typeof newName !== "string") {
    throw new Error("New folder name must be a non-empty string");
  }
  const oldName = folder.getName();
  folder.setName(newName);
  Logger.log(`Folder renamed from "${oldName}" to "${newName}"`);
}

export function moveFolder(
  folder: GoogleAppsScript.Drive.Folder,
  destinationFolder: GoogleAppsScript.Drive.Folder
): void {
  if (!folder) {
    throw new Error("Folder is required");
  }
  if (!destinationFolder) {
    throw new Error("Destination folder is required");
  }
  const folderName = folder.getName();
  folder.moveTo(destinationFolder);
  Logger.log(
    `Folder "${folderName}" moved to folder "${destinationFolder.getName()}"`
  );
}
