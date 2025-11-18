import {
  getOrCreateFolderByPath,
  getFileByName,
  findFile,
  getAllFilesInFolder,
  getAllFoldersInFolder,
  deleteFile,
  deleteFolder,
  copyFile,
  moveFile,
  checkFolderExists,
  checkFileExists,
  getFolderById,
  getFolderByName,
  createFolder,
  getFileById,
  getFilesByType,
  renameFile,
  renameFolder,
  moveFolder,
} from "../src/drive";
import {
  createMockFolder,
  createMockFile,
} from "./helpers/appsScriptEnv";

describe("Drive Module", () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  describe("getOrCreateFolderByPath", () => {
    it("should create a single folder if it doesn't exist", () => {
      const rootFolder = createMockFolder("root");
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const folder = getOrCreateFolderByPath("Projects");

      expect(global.DriveApp.getRootFolder).toHaveBeenCalled();
      expect(rootFolder.getFoldersByName).toHaveBeenCalledWith("Projects");
      expect(rootFolder.createFolder).toHaveBeenCalledWith("Projects");
      expect(folder).toBeDefined();
    });

    it("should return existing folder if it exists", () => {
      const rootFolder = createMockFolder("root");
      const existingFolder = createMockFolder("Projects");
      rootFolder._addFolder(existingFolder);

      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const folder = getOrCreateFolderByPath("Projects");

      expect(rootFolder.getFoldersByName).toHaveBeenCalledWith("Projects");
      expect(rootFolder.createFolder).not.toHaveBeenCalled();
      expect(folder).toBe(existingFolder);
    });

    it("should create nested folder path", () => {
      const rootFolder = createMockFolder("root");
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const folder = getOrCreateFolderByPath("Projects/2024/January");

      expect(rootFolder.getFoldersByName).toHaveBeenCalledWith("Projects");
      expect(rootFolder.createFolder).toHaveBeenCalledWith("Projects");
      expect(folder).toBeDefined();
    });

    it("should handle empty path by returning root folder", () => {
      const rootFolder = createMockFolder("root");
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const folder = getOrCreateFolderByPath("");

      expect(folder).toBe(rootFolder);
      expect(rootFolder.createFolder).not.toHaveBeenCalled();
    });

    it("should handle path with multiple slashes", () => {
      const rootFolder = createMockFolder("root");
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const folder = getOrCreateFolderByPath("Projects//2024///January");

      expect(folder).toBeDefined();
      expect(rootFolder.getFoldersByName).toHaveBeenCalled();
    });

    it("should handle path with leading/trailing slashes", () => {
      const rootFolder = createMockFolder("root");
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const folder = getOrCreateFolderByPath("/Projects/2024/");

      expect(folder).toBeDefined();
    });

    it("should throw error if path is null", () => {
      expect(() => {
        getOrCreateFolderByPath(null as any);
      }).toThrow("Path must be a non-empty string");
    });

    it("should throw error if path is undefined", () => {
      expect(() => {
        getOrCreateFolderByPath(undefined as any);
      }).toThrow("Path must be a non-empty string");
    });

    it("should throw error if path is not a string", () => {
      expect(() => {
        getOrCreateFolderByPath(123 as any);
      }).toThrow("Path must be a non-empty string");
    });
  });

  describe("getFileByName", () => {
    it("should return file if found", () => {
      const folder = createMockFolder("Projects");
      const file = createMockFile("report.docx");
      folder._addFile(file);

      const result = getFileByName("report.docx", folder as any);

      expect(folder.getFilesByName).toHaveBeenCalledWith("report.docx");
      expect(result).toBe(file);
    });

    it("should return null if file not found", () => {
      const folder = createMockFolder("Projects");

      const result = getFileByName("nonexistent.docx", folder as any);

      expect(folder.getFilesByName).toHaveBeenCalledWith("nonexistent.docx");
      expect(result).toBeNull();
    });

    it("should return null if fileName is null", () => {
      const folder = createMockFolder("Projects");

      const result = getFileByName(null as any, folder as any);

      expect(result).toBeNull();
    });

    it("should return null if fileName is undefined", () => {
      const folder = createMockFolder("Projects");

      const result = getFileByName(undefined as any, folder as any);

      expect(result).toBeNull();
    });

    it("should return null if fileName is not a string", () => {
      const folder = createMockFolder("Projects");

      const result = getFileByName(123 as any, folder as any);

      expect(result).toBeNull();
    });

    it("should return null if fileName is empty string", () => {
      const folder = createMockFolder("Projects");

      const result = getFileByName("", folder as any);

      expect(result).toBeNull();
    });

    it("should throw error if folder is null", () => {
      expect(() => {
        getFileByName("test.docx", null as any);
      }).toThrow("Target folder is required");
    });

    it("should throw error if folder is undefined", () => {
      expect(() => {
        getFileByName("test.docx", undefined as any);
      }).toThrow("Target folder is required");
    });
  });

  describe("findFile", () => {
    it("should return file if found", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("report.docx");
      folder._addFile(file);
      rootFolder._addFolder(folder);

      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const result = findFile("Projects", "report.docx");

      expect(result).toBe(file);
    });

    it("should throw error if file not found", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      rootFolder._addFolder(folder);

      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      expect(() => {
        findFile("Projects", "nonexistent.docx");
      }).toThrow('Document "nonexistent.docx" not found in folder: "Projects"');
    });

    it("should create folder path if it doesn't exist", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("report.docx");
      folder._addFile(file);

      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);
      rootFolder.createFolder = jest.fn(() => folder);

      const result = findFile("Projects", "report.docx");

      expect(rootFolder.createFolder).toHaveBeenCalledWith("Projects");
      expect(result).toBe(file);
    });

    it("should throw error if folderPath is null", () => {
      expect(() => {
        findFile(null as any, "test.docx");
      }).toThrow("Folder path must be a non-empty string");
    });

    it("should throw error if folderPath is undefined", () => {
      expect(() => {
        findFile(undefined as any, "test.docx");
      }).toThrow("Folder path must be a non-empty string");
    });

    it("should throw error if fileName is null", () => {
      expect(() => {
        findFile("Projects", null as any);
      }).toThrow("File name must be a non-empty string");
    });

    it("should throw error if fileName is undefined", () => {
      expect(() => {
        findFile("Projects", undefined as any);
      }).toThrow("File name must be a non-empty string");
    });

    it("should throw error if fileName is empty string", () => {
      expect(() => {
        findFile("Projects", "");
      }).toThrow("File name must be a non-empty string");
    });

    it("should handle nested folder paths", () => {
      const rootFolder = createMockFolder("root");
      const projectsFolder = createMockFolder("Projects");
      const yearFolder = createMockFolder("2024");
      const file = createMockFile("report.docx");
      yearFolder._addFile(file);
      projectsFolder._addFolder(yearFolder);
      rootFolder._addFolder(projectsFolder);

      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const result = findFile("Projects/2024", "report.docx");

      expect(result).toBe(file);
    });
  });

  describe("getAllFilesInFolder", () => {
    it("should return all files in folder", () => {
      const folder = createMockFolder("Projects");
      const file1 = createMockFile("file1.docx");
      const file2 = createMockFile("file2.docx");
      folder._addFile(file1);
      folder._addFile(file2);

      const files = getAllFilesInFolder(folder as any);

      expect(files).toHaveLength(2);
      expect(files).toContain(file1);
      expect(files).toContain(file2);
    });

    it("should return empty array if folder is empty", () => {
      const folder = createMockFolder("Projects");

      const files = getAllFilesInFolder(folder as any);

      expect(files).toEqual([]);
      expect(files).toHaveLength(0);
    });

    it("should throw error if folder is null", () => {
      expect(() => {
        getAllFilesInFolder(null as any);
      }).toThrow("Folder is required");
    });

    it("should throw error if folder is undefined", () => {
      expect(() => {
        getAllFilesInFolder(undefined as any);
      }).toThrow("Folder is required");
    });
  });

  describe("getAllFoldersInFolder", () => {
    it("should return all folders in folder", () => {
      const parentFolder = createMockFolder("Projects");
      const folder1 = createMockFolder("2024");
      const folder2 = createMockFolder("2025");
      parentFolder._addFolder(folder1);
      parentFolder._addFolder(folder2);

      const folders = getAllFoldersInFolder(parentFolder as any);

      expect(folders).toHaveLength(2);
      expect(folders).toContain(folder1);
      expect(folders).toContain(folder2);
    });

    it("should return empty array if folder has no subfolders", () => {
      const folder = createMockFolder("Projects");

      const folders = getAllFoldersInFolder(folder as any);

      expect(folders).toEqual([]);
      expect(folders).toHaveLength(0);
    });

    it("should throw error if folder is null", () => {
      expect(() => {
        getAllFoldersInFolder(null as any);
      }).toThrow("Folder is required");
    });

    it("should throw error if folder is undefined", () => {
      expect(() => {
        getAllFoldersInFolder(undefined as any);
      }).toThrow("Folder is required");
    });
  });

  describe("deleteFile", () => {
    it("should delete file successfully", () => {
      const file = createMockFile("test.docx");
      (global.DriveApp as any).removeFile = jest.fn();

      deleteFile(file as any);

      expect(global.DriveApp.removeFile).toHaveBeenCalledWith(file);
    });

    it("should throw error if file is null", () => {
      expect(() => {
        deleteFile(null as any);
      }).toThrow("File is required");
    });

    it("should throw error if file is undefined", () => {
      expect(() => {
        deleteFile(undefined as any);
      }).toThrow("File is required");
    });
  });

  describe("deleteFolder", () => {
    it("should delete folder successfully", () => {
      const folder = createMockFolder("Projects");
      (global.DriveApp as any).removeFolder = jest.fn();

      deleteFolder(folder as any);

      expect(global.DriveApp.removeFolder).toHaveBeenCalledWith(folder);
    });

    it("should throw error if folder is null", () => {
      expect(() => {
        deleteFolder(null as any);
      }).toThrow("Folder is required");
    });

    it("should throw error if folder is undefined", () => {
      expect(() => {
        deleteFolder(undefined as any);
      }).toThrow("Folder is required");
    });
  });

  describe("copyFile", () => {
    it("should copy file to destination folder", () => {
      const file = createMockFile("test.docx");
      const destinationFolder = createMockFolder("Backup");
      const copiedFile = createMockFile("test.docx", "copied-id");
      file.makeCopy = jest.fn(() => copiedFile) as any;

      const result = copyFile(file as any, destinationFolder as any);

      expect(file.makeCopy).toHaveBeenCalledWith("test.docx", destinationFolder);
      expect(result).toBe(copiedFile);
    });

    it("should copy file with new name", () => {
      const file = createMockFile("test.docx");
      const destinationFolder = createMockFolder("Backup");
      const copiedFile = createMockFile("test-copy.docx", "copied-id");
      file.makeCopy = jest.fn(() => copiedFile) as any;

      const result = copyFile(file as any, destinationFolder as any, "test-copy.docx");

      expect(file.makeCopy).toHaveBeenCalledWith("test-copy.docx", destinationFolder);
      expect(result).toBe(copiedFile);
    });

    it("should throw error if file is null", () => {
      const folder = createMockFolder("Backup");
      expect(() => {
        copyFile(null as any, folder as any);
      }).toThrow("File is required");
    });

    it("should throw error if file is undefined", () => {
      const folder = createMockFolder("Backup");
      expect(() => {
        copyFile(undefined as any, folder as any);
      }).toThrow("File is required");
    });

    it("should throw error if destinationFolder is null", () => {
      const file = createMockFile("test.docx");
      expect(() => {
        copyFile(file as any, null as any);
      }).toThrow("Destination folder is required");
    });

    it("should throw error if destinationFolder is undefined", () => {
      const file = createMockFile("test.docx");
      expect(() => {
        copyFile(file as any, undefined as any);
      }).toThrow("Destination folder is required");
    });
  });

  describe("moveFile", () => {
    it("should move file to destination folder", () => {
      const file = createMockFile("test.docx");
      const destinationFolder = createMockFolder("Archive");
      file.moveTo = jest.fn() as any;

      moveFile(file as any, destinationFolder as any);

      expect(file.moveTo).toHaveBeenCalledWith(destinationFolder);
    });

    it("should throw error if file is null", () => {
      const folder = createMockFolder("Archive");
      expect(() => {
        moveFile(null as any, folder as any);
      }).toThrow("File is required");
    });

    it("should throw error if file is undefined", () => {
      const folder = createMockFolder("Archive");
      expect(() => {
        moveFile(undefined as any, folder as any);
      }).toThrow("File is required");
    });

    it("should throw error if destinationFolder is null", () => {
      const file = createMockFile("test.docx");
      expect(() => {
        moveFile(file as any, null as any);
      }).toThrow("Destination folder is required");
    });

    it("should throw error if destinationFolder is undefined", () => {
      const file = createMockFile("test.docx");
      expect(() => {
        moveFile(file as any, undefined as any);
      }).toThrow("Destination folder is required");
    });
  });

  describe("checkFolderExists", () => {
    it("should return true if folder exists", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const exists = checkFolderExists("Projects");

      expect(exists).toBe(true);
    });

    it("should return true if folder path is created", () => {
      const rootFolder = createMockFolder("root");
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const exists = checkFolderExists("NewFolder");

      expect(exists).toBe(true);
    });

    it("should return false if folderPath is null", () => {
      const exists = checkFolderExists(null as any);

      expect(exists).toBe(false);
    });

    it("should return false if folderPath is undefined", () => {
      const exists = checkFolderExists(undefined as any);

      expect(exists).toBe(false);
    });

    it("should return false if folderPath is empty string", () => {
      const exists = checkFolderExists("");

      expect(exists).toBe(false);
    });

    it("should return false if folderPath is not a string", () => {
      const exists = checkFolderExists(123 as any);

      expect(exists).toBe(false);
    });

    it("should return false on error", () => {
      const rootFolder = createMockFolder("root");
      (global.DriveApp as any).getRootFolder = jest.fn(() => {
        throw new Error("Access denied");
      });

      const exists = checkFolderExists("Invalid/Path");

      expect(exists).toBe(false);
    });
  });

  describe("checkFileExists", () => {
    it("should return true if file exists", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("report.docx");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const exists = checkFileExists("Projects", "report.docx");

      expect(exists).toBe(true);
    });

    it("should return false if file does not exist", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const exists = checkFileExists("Projects", "nonexistent.docx");

      expect(exists).toBe(false);
    });

    it("should return false if folderPath is null", () => {
      const exists = checkFileExists(null as any, "file.docx");

      expect(exists).toBe(false);
    });

    it("should return false if folderPath is undefined", () => {
      const exists = checkFileExists(undefined as any, "file.docx");

      expect(exists).toBe(false);
    });

    it("should return false if fileName is null", () => {
      const exists = checkFileExists("Projects", null as any);

      expect(exists).toBe(false);
    });

    it("should return false if fileName is undefined", () => {
      const exists = checkFileExists("Projects", undefined as any);

      expect(exists).toBe(false);
    });

    it("should return false if fileName is empty string", () => {
      const exists = checkFileExists("Projects", "");

      expect(exists).toBe(false);
    });

    it("should return false if folderPath is not a string", () => {
      const exists = checkFileExists(123 as any, "file.docx");

      expect(exists).toBe(false);
    });

    it("should return false if fileName is not a string", () => {
      const exists = checkFileExists("Projects", 123 as any);

      expect(exists).toBe(false);
    });

    it("should return false on error", () => {
      const rootFolder = createMockFolder("root");
      (global.DriveApp as any).getRootFolder = jest.fn(() => {
        throw new Error("Access denied");
      });

      const exists = checkFileExists("Invalid/Path", "file.docx");

      expect(exists).toBe(false);
    });
  });

  describe("getFolderById", () => {
    it("should return folder if found", () => {
      const folder = createMockFolder("Projects", "folder-id-123");
      (global.DriveApp as any).getFolderById = jest.fn(() => folder);

      const result = getFolderById("folder-id-123");

      expect(global.DriveApp.getFolderById).toHaveBeenCalledWith("folder-id-123");
      expect(result).toBe(folder);
    });

    it("should return null if folderId is null", () => {
      const result = getFolderById(null as any);

      expect(result).toBeNull();
    });

    it("should return null if folderId is undefined", () => {
      const result = getFolderById(undefined as any);

      expect(result).toBeNull();
    });

    it("should return null if folderId is empty string", () => {
      const result = getFolderById("");

      expect(result).toBeNull();
    });

    it("should return null if folderId is not a string", () => {
      const result = getFolderById(123 as any);

      expect(result).toBeNull();
    });

    it("should return null if folder not found", () => {
      (global.DriveApp as any).getFolderById = jest.fn(() => {
        throw new Error("Folder not found");
      });

      const result = getFolderById("invalid-id");

      expect(result).toBeNull();
    });
  });

  describe("getFolderByName", () => {
    it("should return folder if found", () => {
      const parentFolder = createMockFolder("Projects");
      const folder = createMockFolder("2024");
      parentFolder._addFolder(folder);

      const result = getFolderByName("2024", parentFolder as any);

      expect(parentFolder.getFoldersByName).toHaveBeenCalledWith("2024");
      expect(result).toBe(folder);
    });

    it("should return null if folder not found", () => {
      const parentFolder = createMockFolder("Projects");

      const result = getFolderByName("2024", parentFolder as any);

      expect(parentFolder.getFoldersByName).toHaveBeenCalledWith("2024");
      expect(result).toBeNull();
    });

    it("should return null if folderName is null", () => {
      const parentFolder = createMockFolder("Projects");

      const result = getFolderByName(null as any, parentFolder as any);

      expect(result).toBeNull();
    });

    it("should return null if folderName is undefined", () => {
      const parentFolder = createMockFolder("Projects");

      const result = getFolderByName(undefined as any, parentFolder as any);

      expect(result).toBeNull();
    });

    it("should return null if folderName is empty string", () => {
      const parentFolder = createMockFolder("Projects");

      const result = getFolderByName("", parentFolder as any);

      expect(result).toBeNull();
    });

    it("should return null if folderName is not a string", () => {
      const parentFolder = createMockFolder("Projects");

      const result = getFolderByName(123 as any, parentFolder as any);

      expect(result).toBeNull();
    });

    it("should throw error if parentFolder is null", () => {
      expect(() => {
        getFolderByName("2024", null as any);
      }).toThrow("Parent folder is required");
    });

    it("should throw error if parentFolder is undefined", () => {
      expect(() => {
        getFolderByName("2024", undefined as any);
      }).toThrow("Parent folder is required");
    });
  });

  describe("createFolder", () => {
    it("should create folder in root if parent not specified", () => {
      const rootFolder = createMockFolder("root");
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const folder = createFolder("NewFolder");

      expect(rootFolder.createFolder).toHaveBeenCalledWith("NewFolder");
      expect(folder).toBeDefined();
    });

    it("should create folder in specified parent", () => {
      const parentFolder = createMockFolder("Projects");
      const newFolder = createMockFolder("2024");
      parentFolder.createFolder = jest.fn(() => newFolder) as any;

      const folder = createFolder("2024", parentFolder as any);

      expect(parentFolder.createFolder).toHaveBeenCalledWith("2024");
      expect(folder).toBe(newFolder);
    });

    it("should throw error if folderName is null", () => {
      expect(() => {
        createFolder(null as any);
      }).toThrow("Folder name must be a non-empty string");
    });

    it("should throw error if folderName is undefined", () => {
      expect(() => {
        createFolder(undefined as any);
      }).toThrow("Folder name must be a non-empty string");
    });

    it("should throw error if folderName is empty string", () => {
      expect(() => {
        createFolder("");
      }).toThrow("Folder name must be a non-empty string");
    });

    it("should throw error if folderName is not a string", () => {
      expect(() => {
        createFolder(123 as any);
      }).toThrow("Folder name must be a non-empty string");
    });
  });

  describe("getFileById", () => {
    it("should return file if found", () => {
      const file = createMockFile("test.docx", "file-id-123");
      (global.DriveApp as any).getFileById = jest.fn(() => file);

      const result = getFileById("file-id-123");

      expect(global.DriveApp.getFileById).toHaveBeenCalledWith("file-id-123");
      expect(result).toBe(file);
    });

    it("should return null if fileId is null", () => {
      const result = getFileById(null as any);

      expect(result).toBeNull();
    });

    it("should return null if fileId is undefined", () => {
      const result = getFileById(undefined as any);

      expect(result).toBeNull();
    });

    it("should return null if fileId is empty string", () => {
      const result = getFileById("");

      expect(result).toBeNull();
    });

    it("should return null if fileId is not a string", () => {
      const result = getFileById(123 as any);

      expect(result).toBeNull();
    });

    it("should return null if file not found", () => {
      (global.DriveApp as any).getFileById = jest.fn(() => {
        throw new Error("File not found");
      });

      const result = getFileById("invalid-id");

      expect(result).toBeNull();
    });
  });

  describe("getFilesByType", () => {
    it("should return files of specified type", () => {
      const folder = createMockFolder("Projects");
      const file1 = createMockFile("doc1.pdf", "file1-id");
      const file2 = createMockFile("doc2.pdf", "file2-id");
      folder._addFile(file1);
      folder._addFile(file2);
      
      const fileIterator: any = {
        hasNext: jest.fn(() => {
          let callCount = 0;
          return () => {
            callCount++;
            if (callCount === 1) return true;
            if (callCount === 2) return true;
            return false;
          };
        })(),
        next: jest.fn(() => {
          let callCount = 0;
          return () => {
            callCount++;
            if (callCount === 1) return file1;
            return file2;
          };
        })(),
      };
      folder.getFilesByType = jest.fn(() => fileIterator) as any;

      const files = getFilesByType(folder as any, "application/pdf");

      expect(folder.getFilesByType).toHaveBeenCalledWith("application/pdf");
      expect(files).toHaveLength(2);
    });

    it("should return empty array if no files of type", () => {
      const folder = createMockFolder("Projects");
      const fileIterator: any = {
        hasNext: jest.fn(() => false),
        next: jest.fn(),
      };
      folder.getFilesByType = jest.fn(() => fileIterator) as any;

      const files = getFilesByType(folder as any, "application/pdf");

      expect(files).toEqual([]);
      expect(files).toHaveLength(0);
    });

    it("should throw error if folder is null", () => {
      expect(() => {
        getFilesByType(null as any, "application/pdf");
      }).toThrow("Folder is required");
    });

    it("should throw error if folder is undefined", () => {
      expect(() => {
        getFilesByType(undefined as any, "application/pdf");
      }).toThrow("Folder is required");
    });

    it("should throw error if mimeType is null", () => {
      const folder = createMockFolder("Projects");
      expect(() => {
        getFilesByType(folder as any, null as any);
      }).toThrow("MIME type must be a non-empty string");
    });

    it("should throw error if mimeType is undefined", () => {
      const folder = createMockFolder("Projects");
      expect(() => {
        getFilesByType(folder as any, undefined as any);
      }).toThrow("MIME type must be a non-empty string");
    });

    it("should throw error if mimeType is empty string", () => {
      const folder = createMockFolder("Projects");
      expect(() => {
        getFilesByType(folder as any, "");
      }).toThrow("MIME type must be a non-empty string");
    });

    it("should throw error if mimeType is not a string", () => {
      const folder = createMockFolder("Projects");
      expect(() => {
        getFilesByType(folder as any, 123 as any);
      }).toThrow("MIME type must be a non-empty string");
    });
  });

  describe("renameFile", () => {
    it("should rename file successfully", () => {
      const file = createMockFile("old.docx");
      file.setName = jest.fn() as any;

      renameFile(file as any, "new.docx");

      expect(file.setName).toHaveBeenCalledWith("new.docx");
    });

    it("should throw error if file is null", () => {
      expect(() => {
        renameFile(null as any, "new.docx");
      }).toThrow("File is required");
    });

    it("should throw error if file is undefined", () => {
      expect(() => {
        renameFile(undefined as any, "new.docx");
      }).toThrow("File is required");
    });

    it("should throw error if newName is null", () => {
      const file = createMockFile("old.docx");
      expect(() => {
        renameFile(file as any, null as any);
      }).toThrow("New file name must be a non-empty string");
    });

    it("should throw error if newName is undefined", () => {
      const file = createMockFile("old.docx");
      expect(() => {
        renameFile(file as any, undefined as any);
      }).toThrow("New file name must be a non-empty string");
    });

    it("should throw error if newName is empty string", () => {
      const file = createMockFile("old.docx");
      expect(() => {
        renameFile(file as any, "");
      }).toThrow("New file name must be a non-empty string");
    });

    it("should throw error if newName is not a string", () => {
      const file = createMockFile("old.docx");
      expect(() => {
        renameFile(file as any, 123 as any);
      }).toThrow("New file name must be a non-empty string");
    });
  });

  describe("renameFolder", () => {
    it("should rename folder successfully", () => {
      const folder = createMockFolder("OldFolder");
      folder.setName = jest.fn() as any;

      renameFolder(folder as any, "NewFolder");

      expect(folder.setName).toHaveBeenCalledWith("NewFolder");
    });

    it("should throw error if folder is null", () => {
      expect(() => {
        renameFolder(null as any, "NewFolder");
      }).toThrow("Folder is required");
    });

    it("should throw error if folder is undefined", () => {
      expect(() => {
        renameFolder(undefined as any, "NewFolder");
      }).toThrow("Folder is required");
    });

    it("should throw error if newName is null", () => {
      const folder = createMockFolder("OldFolder");
      expect(() => {
        renameFolder(folder as any, null as any);
      }).toThrow("New folder name must be a non-empty string");
    });

    it("should throw error if newName is undefined", () => {
      const folder = createMockFolder("OldFolder");
      expect(() => {
        renameFolder(folder as any, undefined as any);
      }).toThrow("New folder name must be a non-empty string");
    });

    it("should throw error if newName is empty string", () => {
      const folder = createMockFolder("OldFolder");
      expect(() => {
        renameFolder(folder as any, "");
      }).toThrow("New folder name must be a non-empty string");
    });

    it("should throw error if newName is not a string", () => {
      const folder = createMockFolder("OldFolder");
      expect(() => {
        renameFolder(folder as any, 123 as any);
      }).toThrow("New folder name must be a non-empty string");
    });
  });

  describe("moveFolder", () => {
    it("should move folder to destination", () => {
      const folder = createMockFolder("Projects");
      const destinationFolder = createMockFolder("Archive");
      folder.moveTo = jest.fn() as any;

      moveFolder(folder as any, destinationFolder as any);

      expect(folder.moveTo).toHaveBeenCalledWith(destinationFolder);
    });

    it("should throw error if folder is null", () => {
      const destinationFolder = createMockFolder("Archive");
      expect(() => {
        moveFolder(null as any, destinationFolder as any);
      }).toThrow("Folder is required");
    });

    it("should throw error if folder is undefined", () => {
      const destinationFolder = createMockFolder("Archive");
      expect(() => {
        moveFolder(undefined as any, destinationFolder as any);
      }).toThrow("Folder is required");
    });

    it("should throw error if destinationFolder is null", () => {
      const folder = createMockFolder("Projects");
      expect(() => {
        moveFolder(folder as any, null as any);
      }).toThrow("Destination folder is required");
    });

    it("should throw error if destinationFolder is undefined", () => {
      const folder = createMockFolder("Projects");
      expect(() => {
        moveFolder(folder as any, undefined as any);
      }).toThrow("Destination folder is required");
    });
  });
});
