import {
  ensureFolder,
  createDocument,
  appendParagraphToFile,
  appendBulletedListToFile,
  appendNumberedListToFile,
  insertParagraphAtPosition,
  replaceTextInFile,
  clearDocument,
  getDocumentContent,
  getParagraphCount,
  insertTable,
  insertImage,
  deleteParagraph,
  getParagraphAtPosition,
  formatParagraph,
} from "../src/docs";
import {
  createMockFolder,
  createMockFile,
  createMockDocument,
  createMockParagraph,
} from "./helpers/appsScriptEnv";

describe("Docs Module", () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  describe("ensureFolder", () => {
    it("should create folder if it doesn't exist", () => {
      const rootFolder = createMockFolder("root");
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const folder = ensureFolder("Projects/2024");

      expect(folder).toBeDefined();
      expect(rootFolder.getFoldersByName).toHaveBeenCalled();
    });

    it("should return existing folder if it exists", () => {
      const rootFolder = createMockFolder("root");
      const existingFolder = createMockFolder("Projects");
      rootFolder._addFolder(existingFolder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const folder = ensureFolder("Projects");

      expect(folder).toBe(existingFolder);
      expect(rootFolder.createFolder).not.toHaveBeenCalled();
    });

    it("should throw error if folderPath is null", () => {
      expect(() => {
        ensureFolder(null as any);
      }).toThrow("Folder path must be a non-empty string");
    });

    it("should throw error if folderPath is undefined", () => {
      expect(() => {
        ensureFolder(undefined as any);
      }).toThrow("Folder path must be a non-empty string");
    });

    it("should throw error if folderPath is empty string", () => {
      expect(() => {
        ensureFolder("");
      }).toThrow("Folder path must be a non-empty string");
    });

    it("should throw error if folderPath is not a string", () => {
      expect(() => {
        ensureFolder(123 as any);
      }).toThrow("Folder path must be a non-empty string");
    });

    it("should handle nested folder paths", () => {
      const rootFolder = createMockFolder("root");
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const folder = ensureFolder("Projects/2024/Q1");

      expect(folder).toBeDefined();
    });
  });

  describe("createDocument", () => {
    it("should create a new document", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report");
      const mockFile = createMockFile("Report", mockDoc.getId());

      (global.DocumentApp as any).create = jest.fn(() => mockDoc);
      (global.DriveApp as any).getFileById = jest.fn(() => mockFile);

      const doc = createDocument("Projects", "Report");

      expect(global.DocumentApp.create).toHaveBeenCalledWith("Report");
      expect(global.DriveApp.getFileById).toHaveBeenCalledWith(mockDoc.getId());
      expect(mockFile.moveTo).toHaveBeenCalledWith(folder);
      expect(doc).toBe(mockDoc);
    });

    it("should return existing document if it exists", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const existingFile = createMockFile("Report", "doc-id");
      folder._addFile(existingFile);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      const doc = createDocument("Projects", "Report");

      expect(global.DocumentApp.create).not.toHaveBeenCalled();
      expect(global.DocumentApp.openById).toHaveBeenCalledWith("doc-id");
      expect(doc).toBe(mockDoc);
    });

    it("should throw error if folderPath is null", () => {
      expect(() => {
        createDocument(null as any, "Report");
      }).toThrow("Folder path must be a non-empty string");
    });

    it("should throw error if folderPath is undefined", () => {
      expect(() => {
        createDocument(undefined as any, "Report");
      }).toThrow("Folder path must be a non-empty string");
    });

    it("should throw error if folderPath is empty string", () => {
      expect(() => {
        createDocument("", "Report");
      }).toThrow("Folder path must be a non-empty string");
    });

    it("should throw error if docName is null", () => {
      expect(() => {
        createDocument("Projects", null as any);
      }).toThrow("Document name must be a non-empty string");
    });

    it("should throw error if docName is undefined", () => {
      expect(() => {
        createDocument("Projects", undefined as any);
      }).toThrow("Document name must be a non-empty string");
    });

    it("should throw error if docName is empty string", () => {
      expect(() => {
        createDocument("Projects", "");
      }).toThrow("Document name must be a non-empty string");
    });

    it("should throw error if docName is not a string", () => {
      expect(() => {
        createDocument("Projects", 123 as any);
      }).toThrow("Document name must be a non-empty string");
    });
  });

  describe("appendParagraphToFile", () => {
    it("should append paragraph to document", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      const body = mockDoc.getBody();
      const paragraph = createMockParagraph();
      body.getNumChildren = jest.fn(() => 0); // Empty document
      body.appendParagraph = jest.fn((_text: string) => paragraph) as any;
      body.insertParagraph = jest.fn(
        (_index: number, _text: string) => paragraph
      ) as any;

      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      const result = appendParagraphToFile("Projects", "Report", "Hello World");

      expect(global.DocumentApp.openById).toHaveBeenCalledWith("doc-id");
      expect(body.insertParagraph).toHaveBeenCalledWith(0, "Hello World");
      expect(result).toBe(paragraph);
      expect(mockDoc.saveAndClose).toHaveBeenCalled();
    });

    it("should replace empty last paragraph instead of appending", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      const emptyParagraph = createMockParagraph("");
      const body = mockDoc.getBody();
      body.getNumChildren = jest.fn(() => 1);
      body.getChild = jest.fn((_index: number) => emptyParagraph) as any;
      emptyParagraph.getType = jest.fn(() => "PARAGRAPH");

      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      const result = appendParagraphToFile("Projects", "Report", "Hello World");

      expect(emptyParagraph.setText).toHaveBeenCalledWith("Hello World");
      expect(body.appendParagraph).not.toHaveBeenCalled();
      expect(result).toBe(emptyParagraph);
    });

    it("should apply heading if provided", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      const body = mockDoc.getBody();
      const paragraph = createMockParagraph();
      body.getNumChildren = jest.fn(() => 0); // Empty document
      paragraph.setHeading = jest.fn((_heading: any) => paragraph) as any;
      body.insertParagraph = jest.fn(
        (_index: number, _text: string) => paragraph
      ) as any;

      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      appendParagraphToFile(
        "Projects",
        "Report",
        "Heading",
        (global.DocumentApp as any).ParagraphHeading.HEADING1
      );

      expect(paragraph.setHeading).toHaveBeenCalledWith("HEADING1");
    });

    it("should throw error if folderPath is null", () => {
      expect(() => {
        appendParagraphToFile(null as any, "Report", "Content");
      }).toThrow("Folder path must be a non-empty string");
    });

    it("should throw error if folderPath is undefined", () => {
      expect(() => {
        appendParagraphToFile(undefined as any, "Report", "Content");
      }).toThrow("Folder path must be a non-empty string");
    });

    it("should throw error if fileName is null", () => {
      expect(() => {
        appendParagraphToFile("Projects", null as any, "Content");
      }).toThrow("File name must be a non-empty string");
    });

    it("should throw error if fileName is undefined", () => {
      expect(() => {
        appendParagraphToFile("Projects", undefined as any, "Content");
      }).toThrow("File name must be a non-empty string");
    });

    it("should throw error if content is null", () => {
      expect(() => {
        appendParagraphToFile("Projects", "Report", null as any);
      }).toThrow("Content must be a string");
    });

    it("should throw error if content is undefined", () => {
      expect(() => {
        appendParagraphToFile("Projects", "Report", undefined as any);
      }).toThrow("Content must be a string");
    });

    it("should throw error if content is not a string", () => {
      expect(() => {
        appendParagraphToFile("Projects", "Report", 123 as any);
      }).toThrow("Content must be a string");
    });

    it("should handle empty content string", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      const body = mockDoc.getBody();
      const paragraph = createMockParagraph();
      body.getNumChildren = jest.fn(() => 0);
      body.insertParagraph = jest.fn(() => paragraph) as any;

      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      const result = appendParagraphToFile("Projects", "Report", "");

      expect(result).toBeDefined();
    });

    it("should handle document with non-paragraph last child", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      const body = mockDoc.getBody();
      const nonParagraph = { getType: jest.fn(() => "TABLE") };
      body.getNumChildren = jest.fn(() => 1);
      body.getChild = jest.fn(() => nonParagraph) as any;
      const paragraph = createMockParagraph();
      body.appendParagraph = jest.fn(() => paragraph) as any;

      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      const result = appendParagraphToFile("Projects", "Report", "Content");

      expect(result).toBe(paragraph);
    });

    it("should append paragraph when last paragraph is not empty", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      const body = mockDoc.getBody();
      const lastParagraph = createMockParagraph("Existing content");
      body.getNumChildren = jest.fn(() => 1);
      body.getChild = jest.fn((_index: number) => lastParagraph) as any;
      lastParagraph.getType = jest.fn(() => "PARAGRAPH");
      const newParagraph = createMockParagraph();
      body.appendParagraph = jest.fn(() => newParagraph) as any;

      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      const result = appendParagraphToFile("Projects", "Report", "New Content");

      expect(body.appendParagraph).toHaveBeenCalledWith("New Content");
      expect(result).toBe(newParagraph);
    });
  });

  describe("appendBulletedListToFile", () => {
    it("should append bulleted list to document", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      const body = mockDoc.getBody();
      const listItem1: any = {
        setGlyphType: jest.fn(),
        setAttributes: jest.fn(),
        setAlignment: jest.fn(),
        getText: jest.fn(),
        setText: jest.fn(),
        setHeading: jest.fn(),
        getType: jest.fn(() => "LIST_ITEM"),
        asParagraph: jest.fn(function () {
          return this;
        }),
      };
      const listItem2: any = {
        setGlyphType: jest.fn(),
        setAttributes: jest.fn(),
        setAlignment: jest.fn(),
        getText: jest.fn(),
        setText: jest.fn(),
        setHeading: jest.fn(),
        getType: jest.fn(() => "LIST_ITEM"),
        asParagraph: jest.fn(function () {
          return this;
        }),
      };
      body.appendListItem = jest.fn((text: string) => {
        if (text === "Item 1") return listItem1;
        return listItem2;
      }) as any;

      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      const result = appendBulletedListToFile("Projects", "Report", [
        "Item 1",
        "Item 2",
      ]);

      expect(body.appendListItem).toHaveBeenCalledTimes(2);
      expect(listItem1.setGlyphType).toHaveBeenCalledWith("BULLET");
      expect(listItem2.setGlyphType).toHaveBeenCalledWith("BULLET");
      expect(result).toHaveLength(2);
      expect(mockDoc.saveAndClose).toHaveBeenCalled();
    });

    it("should return empty array if no items provided", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      const result = appendBulletedListToFile("Projects", "Report", []);

      expect(result).toEqual([]);
      expect(mockDoc.getBody().appendListItem).not.toHaveBeenCalled();
    });

    it("should throw error if folderPath is null", () => {
      expect(() => {
        appendBulletedListToFile(null as any, "Report", ["Item"]);
      }).toThrow("Folder path must be a non-empty string");
    });

    it("should throw error if folderPath is undefined", () => {
      expect(() => {
        appendBulletedListToFile(undefined as any, "Report", ["Item"]);
      }).toThrow("Folder path must be a non-empty string");
    });

    it("should throw error if fileName is null", () => {
      expect(() => {
        appendBulletedListToFile("Projects", null as any, ["Item"]);
      }).toThrow("File name must be a non-empty string");
    });

    it("should throw error if items is null", () => {
      expect(() => {
        appendBulletedListToFile("Projects", "Report", null as any);
      }).toThrow("Items must be an array");
    });

    it("should throw error if items is undefined", () => {
      expect(() => {
        appendBulletedListToFile("Projects", "Report", undefined as any);
      }).toThrow("Items must be an array");
    });

    it("should throw error if items is not an array", () => {
      expect(() => {
        appendBulletedListToFile("Projects", "Report", "not an array" as any);
      }).toThrow("Items must be an array");
    });

    it("should handle single item list", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      const body = mockDoc.getBody();
      const listItem: any = {
        setGlyphType: jest.fn(),
        setAttributes: jest.fn(),
        setAlignment: jest.fn(),
      };
      body.appendListItem = jest.fn(() => listItem) as any;

      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      const result = appendBulletedListToFile("Projects", "Report", [
        "Single Item",
      ]);

      expect(result).toHaveLength(1);
      expect(body.appendListItem).toHaveBeenCalledTimes(1);
    });

    it("should skip non-string items", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      const body = mockDoc.getBody();
      const listItem: any = {
        setGlyphType: jest.fn(),
        setAttributes: jest.fn(),
        setAlignment: jest.fn(),
      };
      body.appendListItem = jest.fn(() => listItem) as any;

      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      const result = appendBulletedListToFile("Projects", "Report", [
        "Item",
        123,
        null,
        "Another Item",
      ] as any);

      expect(body.appendListItem).toHaveBeenCalledTimes(2);
      expect(result).toHaveLength(2);
    });
  });

  describe("appendNumberedListToFile", () => {
    it("should append numbered list to document", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      const body = mockDoc.getBody();
      const listItem1: any = {
        setGlyphType: jest.fn(),
        setAttributes: jest.fn(),
        setAlignment: jest.fn(),
      };
      const listItem2: any = {
        setGlyphType: jest.fn(),
        setAttributes: jest.fn(),
        setAlignment: jest.fn(),
      };
      body.appendListItem = jest.fn((text: string) => {
        if (text === "First") return listItem1;
        return listItem2;
      }) as any;

      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      const result = appendNumberedListToFile("Projects", "Report", [
        "First",
        "Second",
      ]);

      expect(body.appendListItem).toHaveBeenCalledTimes(2);
      expect(listItem1.setGlyphType).toHaveBeenCalledWith("NUMBER");
      expect(listItem2.setGlyphType).toHaveBeenCalledWith("NUMBER");
      expect(result).toHaveLength(2);
      expect(mockDoc.saveAndClose).toHaveBeenCalled();
    });

    it("should return empty array if no items provided", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      const result = appendNumberedListToFile("Projects", "Report", []);

      expect(result).toEqual([]);
    });

    it("should throw error if folderPath is null", () => {
      expect(() => {
        appendNumberedListToFile(null as any, "Report", ["Item"]);
      }).toThrow("Folder path must be a non-empty string");
    });

    it("should throw error if items is not an array", () => {
      expect(() => {
        appendNumberedListToFile("Projects", "Report", "not an array" as any);
      }).toThrow("Items must be an array");
    });

    it("should throw error if fileName is null", () => {
      expect(() => {
        appendNumberedListToFile("Projects", null as any, ["Item"]);
      }).toThrow("File name must be a non-empty string");
    });

    it("should throw error if fileName is undefined", () => {
      expect(() => {
        appendNumberedListToFile("Projects", undefined as any, ["Item"]);
      }).toThrow("File name must be a non-empty string");
    });

    it("should skip non-string items", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      const body = mockDoc.getBody();
      const listItem: any = {
        setGlyphType: jest.fn(),
        setAttributes: jest.fn(),
        setAlignment: jest.fn(),
      };
      body.appendListItem = jest.fn(() => listItem) as any;

      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      const result = appendNumberedListToFile("Projects", "Report", [
        "Item",
        123,
      ] as any);

      expect(body.appendListItem).toHaveBeenCalledTimes(1);
      expect(result).toHaveLength(1);
    });
  });

  describe("insertParagraphAtPosition", () => {
    it("should insert paragraph at specified position", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      const body = mockDoc.getBody();
      const paragraph = createMockParagraph();
      body.getNumChildren = jest.fn(() => 2);
      body.insertParagraph = jest.fn(() => paragraph) as any;

      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      const result = insertParagraphAtPosition(
        "Projects",
        "Report",
        "New Content",
        1
      );

      expect(body.insertParagraph).toHaveBeenCalledWith(1, "New Content");
      expect(result).toBe(paragraph);
      expect(mockDoc.saveAndClose).toHaveBeenCalled();
    });

    it("should clamp position to valid range", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      const body = mockDoc.getBody();
      const paragraph = createMockParagraph();
      body.getNumChildren = jest.fn(() => 2);
      body.insertParagraph = jest.fn(() => paragraph) as any;

      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      const result = insertParagraphAtPosition(
        "Projects",
        "Report",
        "Content",
        100
      );

      expect(body.insertParagraph).toHaveBeenCalledWith(2, "Content");
      expect(result).toBe(paragraph);
    });

    it("should throw error if folderPath is null", () => {
      expect(() => {
        insertParagraphAtPosition(null as any, "Report", "Content", 0);
      }).toThrow("Folder path must be a non-empty string");
    });

    it("should throw error if fileName is null", () => {
      expect(() => {
        insertParagraphAtPosition("Projects", null as any, "Content", 0);
      }).toThrow("File name must be a non-empty string");
    });

    it("should throw error if fileName is undefined", () => {
      expect(() => {
        insertParagraphAtPosition("Projects", undefined as any, "Content", 0);
      }).toThrow("File name must be a non-empty string");
    });

    it("should throw error if content is null", () => {
      expect(() => {
        insertParagraphAtPosition("Projects", "Report", null as any, 0);
      }).toThrow("Content must be a string");
    });

    it("should throw error if content is undefined", () => {
      expect(() => {
        insertParagraphAtPosition("Projects", "Report", undefined as any, 0);
      }).toThrow("Content must be a string");
    });

    it("should throw error if position is null", () => {
      expect(() => {
        insertParagraphAtPosition("Projects", "Report", "Content", null as any);
      }).toThrow("Position must be a number");
    });

    it("should throw error if position is negative", () => {
      expect(() => {
        insertParagraphAtPosition("Projects", "Report", "Content", -1);
      }).toThrow("Position must be >= 0");
    });

    it("should throw error if position is not a number", () => {
      expect(() => {
        insertParagraphAtPosition(
          "Projects",
          "Report",
          "Content",
          "not a number" as any
        );
      }).toThrow("Position must be a number");
    });

    it("should apply heading if provided", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      const body = mockDoc.getBody();
      const paragraph = createMockParagraph();
      paragraph.setHeading = jest.fn(() => paragraph) as any;
      body.getNumChildren = jest.fn(() => 1);
      body.insertParagraph = jest.fn(() => paragraph) as any;

      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      insertParagraphAtPosition(
        "Projects",
        "Report",
        "Heading",
        0,
        (global.DocumentApp as any).ParagraphHeading.HEADING2
      );

      expect(paragraph.setHeading).toHaveBeenCalledWith("HEADING2");
    });
  });

  describe("replaceTextInFile", () => {
    it("should replace text in document", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      const body = mockDoc.getBody();
      const paragraph = createMockParagraph("Hello {{name}}");
      paragraph.setText = jest.fn();
      body.getNumChildren = jest.fn(() => 1);
      body.getChild = jest.fn((_index: number) => paragraph) as any;
      paragraph.getType = jest.fn(() => "PARAGRAPH");

      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      const count = replaceTextInFile(
        "Projects",
        "Report",
        "{{name}}",
        "World"
      );

      expect(paragraph.getText).toHaveBeenCalled();
      expect(paragraph.setText).toHaveBeenCalledWith("Hello World");
      expect(count).toBe(1);
      expect(mockDoc.saveAndClose).toHaveBeenCalled();
    });

    it("should handle multiple replacements", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      const body = mockDoc.getBody();
      const paragraph = createMockParagraph("Hello {{name}}, hello {{name}}");
      paragraph.setText = jest.fn();
      body.getNumChildren = jest.fn(() => 1);
      body.getChild = jest.fn((_index: number) => paragraph) as any;
      paragraph.getType = jest.fn(() => "PARAGRAPH");

      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      const count = replaceTextInFile(
        "Projects",
        "Report",
        "{{name}}",
        "World"
      );

      expect(paragraph.setText).toHaveBeenCalledWith(
        "Hello World, hello World"
      );
      expect(count).toBe(2);
    });

    it("should throw error for invalid regex pattern", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      expect(() => {
        replaceTextInFile("Projects", "Report", "[invalid", "replacement");
      }).toThrow("Invalid search pattern");
    });

    it("should throw error if folderPath is null", () => {
      expect(() => {
        replaceTextInFile(null as any, "Report", "pattern", "replacement");
      }).toThrow("Folder path must be a non-empty string");
    });

    it("should throw error if folderPath is undefined", () => {
      expect(() => {
        replaceTextInFile(undefined as any, "Report", "pattern", "replacement");
      }).toThrow("Folder path must be a non-empty string");
    });

    it("should throw error if fileName is null", () => {
      expect(() => {
        replaceTextInFile("Projects", null as any, "pattern", "replacement");
      }).toThrow("File name must be a non-empty string");
    });

    it("should throw error if fileName is undefined", () => {
      expect(() => {
        replaceTextInFile(
          "Projects",
          undefined as any,
          "pattern",
          "replacement"
        );
      }).toThrow("File name must be a non-empty string");
    });

    it("should throw error if searchPattern is null", () => {
      expect(() => {
        replaceTextInFile("Projects", "Report", null as any, "replacement");
      }).toThrow("Search pattern must be a string");
    });

    it("should throw error if replacementText is null", () => {
      expect(() => {
        replaceTextInFile("Projects", "Report", "pattern", null as any);
      }).toThrow("Replacement text must be a string");
    });

    it("should handle empty replacement text", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      const body = mockDoc.getBody();
      const paragraph = createMockParagraph("Hello World");
      paragraph.setText = jest.fn();
      body.getNumChildren = jest.fn(() => 1);
      body.getChild = jest.fn(() => paragraph) as any;
      paragraph.getType = jest.fn(() => "PARAGRAPH");

      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      const count = replaceTextInFile("Projects", "Report", "World", "");

      expect(paragraph.setText).toHaveBeenCalledWith("Hello ");
      expect(count).toBe(1);
    });

    it("should handle document with no matching patterns", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      const body = mockDoc.getBody();
      const paragraph = createMockParagraph("Hello World");
      paragraph.setText = jest.fn();
      body.getNumChildren = jest.fn(() => 1);
      body.getChild = jest.fn(() => paragraph) as any;
      paragraph.getType = jest.fn(() => "PARAGRAPH");

      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      const count = replaceTextInFile(
        "Projects",
        "Report",
        "NotFound",
        "Replacement"
      );

      expect(count).toBe(0);
    });

    it("should handle list items", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      const body = mockDoc.getBody();
      const listItem = createMockParagraph("Item {{value}}");
      listItem.setText = jest.fn();
      listItem.getType = jest.fn(() => "LIST_ITEM");
      body.getNumChildren = jest.fn(() => 1);
      body.getChild = jest.fn(() => listItem) as any;

      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      const count = replaceTextInFile("Projects", "Report", "{{value}}", "123");

      expect(listItem.setText).toHaveBeenCalledWith("Item 123");
      expect(count).toBe(1);
    });
  });

  describe("clearDocument", () => {
    it("should clear document content", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      const body = mockDoc.getBody();
      body.clear = jest.fn();

      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      clearDocument("Projects", "Report");

      expect(body.clear).toHaveBeenCalled();
      expect(mockDoc.saveAndClose).toHaveBeenCalled();
    });

    it("should throw error if folderPath is null", () => {
      expect(() => {
        clearDocument(null as any, "Report");
      }).toThrow("Folder path must be a non-empty string");
    });

    it("should throw error if fileName is null", () => {
      expect(() => {
        clearDocument("Projects", null as any);
      }).toThrow("File name must be a non-empty string");
    });

    it("should throw error if folderPath is empty string", () => {
      expect(() => {
        clearDocument("", "Report");
      }).toThrow("Folder path must be a non-empty string");
    });
  });

  describe("getDocumentContent", () => {
    it("should return document content as string", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      const body = mockDoc.getBody();
      body.getText = jest.fn(() => "Hello World\nThis is content");

      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      const content = getDocumentContent("Projects", "Report");

      expect(body.getText).toHaveBeenCalled();
      expect(content).toBe("Hello World\nThis is content");
      expect(mockDoc.saveAndClose).toHaveBeenCalled();
    });

    it("should return empty string for empty document", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      const body = mockDoc.getBody();
      body.getText = jest.fn(() => "");

      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      const content = getDocumentContent("Projects", "Report");

      expect(content).toBe("");
    });

    it("should throw error if folderPath is null", () => {
      expect(() => {
        getDocumentContent(null as any, "Report");
      }).toThrow("Folder path must be a non-empty string");
    });

    it("should throw error if fileName is null", () => {
      expect(() => {
        getDocumentContent("Projects", null as any);
      }).toThrow("File name must be a non-empty string");
    });
  });

  describe("getParagraphCount", () => {
    it("should return count of paragraphs", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      const body = mockDoc.getBody();
      const para1 = createMockParagraph("First");
      const para2 = createMockParagraph("Second");
      const table = { getType: jest.fn(() => "TABLE") };
      body.getNumChildren = jest.fn(() => 3);
      body.getChild = jest.fn((index: number) => {
        if (index === 0) return para1;
        if (index === 1) return para2;
        return table;
      }) as any;
      para1.getType = jest.fn(() => "PARAGRAPH");
      para2.getType = jest.fn(() => "PARAGRAPH");

      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      const count = getParagraphCount("Projects", "Report");

      expect(count).toBe(2);
      expect(mockDoc.saveAndClose).toHaveBeenCalled();
    });

    it("should return 0 for empty document", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      const body = mockDoc.getBody();
      body.getNumChildren = jest.fn(() => 0);

      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      const count = getParagraphCount("Projects", "Report");

      expect(count).toBe(0);
    });

    it("should throw error if folderPath is null", () => {
      expect(() => {
        getParagraphCount(null as any, "Report");
      }).toThrow("Folder path must be a non-empty string");
    });

    it("should throw error if fileName is null", () => {
      expect(() => {
        getParagraphCount("Projects", null as any);
      }).toThrow("File name must be a non-empty string");
    });
  });

  describe("insertTable", () => {
    it("should insert table with specified dimensions", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      const body = mockDoc.getBody();
      const tableRow: any = {
        appendTableCell: jest.fn((_text?: string) => ({
          setText: jest.fn(),
        })),
      };
      const table: any = {
        appendTableRow: jest.fn(() => tableRow),
      };
      body.appendTable = jest.fn(() => table) as any;

      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      const result = insertTable("Projects", "Report", 2, 3);

      expect(body.appendTable).toHaveBeenCalled();
      expect(table.appendTableRow).toHaveBeenCalledTimes(2);
      expect(tableRow.appendTableCell).toHaveBeenCalledTimes(6); // 2 rows * 3 columns
      expect(result).toBe(table);
      expect(mockDoc.saveAndClose).toHaveBeenCalled();
    });

    it("should insert table with cell values", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      const body = mockDoc.getBody();
      const cell1: any = { setText: jest.fn() };
      const _cell2: any = { setText: jest.fn() };
      const tableRow: any = {
        appendTableCell: jest.fn((text?: string) => {
          if (!text) return cell1;
          const cell = { setText: jest.fn() };
          cell.setText(text);
          return cell;
        }),
      };
      const table: any = {
        appendTableRow: jest.fn(() => tableRow),
      };
      body.appendTable = jest.fn(() => table) as any;

      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      const result = insertTable("Projects", "Report", 1, 2, [["A", "B"]]);

      expect(tableRow.appendTableCell).toHaveBeenCalledTimes(2);
      expect(result).toBe(table);
    });

    it("should throw error if folderPath is null", () => {
      expect(() => {
        insertTable(null as any, "Report", 2, 3);
      }).toThrow("Folder path must be a non-empty string");
    });

    it("should throw error if fileName is null", () => {
      expect(() => {
        insertTable("Projects", null as any, 2, 3);
      }).toThrow("File name must be a non-empty string");
    });

    it("should throw error if fileName is undefined", () => {
      expect(() => {
        insertTable("Projects", undefined as any, 2, 3);
      }).toThrow("File name must be a non-empty string");
    });

    it("should throw error if rows is null", () => {
      expect(() => {
        insertTable("Projects", "Report", null as any, 3);
      }).toThrow("Rows must be a number");
    });

    it("should throw error if rows is undefined", () => {
      expect(() => {
        insertTable("Projects", "Report", undefined as any, 3);
      }).toThrow("Rows must be a number");
    });

    it("should throw error if rows is less than 1", () => {
      expect(() => {
        insertTable("Projects", "Report", 0, 3);
      }).toThrow("Rows must be >= 1");
    });

    it("should throw error if columns is null", () => {
      expect(() => {
        insertTable("Projects", "Report", 2, null as any);
      }).toThrow("Columns must be a number");
    });

    it("should throw error if columns is undefined", () => {
      expect(() => {
        insertTable("Projects", "Report", 2, undefined as any);
      }).toThrow("Columns must be a number");
    });

    it("should throw error if columns is less than 1", () => {
      expect(() => {
        insertTable("Projects", "Report", 2, 0);
      }).toThrow("Columns must be >= 1");
    });

    it("should throw error if cellValues is not a 2D array", () => {
      expect(() => {
        insertTable("Projects", "Report", 2, 3, ["not", "2d", "array"] as any);
      }).toThrow("Cell values must be a 2D array");
    });

    it("should handle undefined cellValues", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      const body = mockDoc.getBody();
      const tableRow: any = {
        appendTableCell: jest.fn(() => ({ setText: jest.fn() })),
      };
      const table: any = {
        appendTableRow: jest.fn(() => tableRow),
      };
      body.appendTable = jest.fn(() => table) as any;

      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      const result = insertTable("Projects", "Report", 1, 1);

      expect(result).toBe(table);
    });
  });

  describe("insertImage", () => {
    it("should insert image without dimensions", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      const body = mockDoc.getBody();
      const image: any = {
        setWidth: jest.fn(),
        setHeight: jest.fn(),
      };
      body.appendImage = jest.fn(() => image) as any;

      const imageBlob: any = { getName: () => "test.png" };

      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      const result = insertImage("Projects", "Report", imageBlob);

      expect(body.appendImage).toHaveBeenCalledWith(imageBlob);
      expect(result).toBe(image);
      expect(mockDoc.saveAndClose).toHaveBeenCalled();
    });

    it("should insert image with width and height", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      const body = mockDoc.getBody();
      const image: any = {
        setWidth: jest.fn(),
        setHeight: jest.fn(),
      };
      body.appendImage = jest.fn(() => image) as any;

      const imageBlob: any = { getName: () => "test.png" };

      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      const result = insertImage("Projects", "Report", imageBlob, 200, 150);

      expect(image.setWidth).toHaveBeenCalledWith(200);
      expect(image.setHeight).toHaveBeenCalledWith(150);
      expect(result).toBe(image);
    });

    it("should throw error if folderPath is null", () => {
      expect(() => {
        insertImage(null as any, "Report", {} as any);
      }).toThrow("Folder path must be a non-empty string");
    });

    it("should throw error if fileName is null", () => {
      expect(() => {
        insertImage("Projects", null as any, {} as any);
      }).toThrow("File name must be a non-empty string");
    });

    it("should throw error if fileName is undefined", () => {
      expect(() => {
        insertImage("Projects", undefined as any, {} as any);
      }).toThrow("File name must be a non-empty string");
    });

    it("should throw error if imageBlob is null", () => {
      expect(() => {
        insertImage("Projects", "Report", null as any);
      }).toThrow("Image blob is required");
    });

    it("should ignore invalid width", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      const body = mockDoc.getBody();
      const image: any = {
        setWidth: jest.fn(),
        setHeight: jest.fn(),
      };
      body.appendImage = jest.fn(() => image) as any;

      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      insertImage("Projects", "Report", {} as any, -100, 150);

      expect(image.setWidth).not.toHaveBeenCalled();
      expect(image.setHeight).toHaveBeenCalledWith(150);
    });
  });

  describe("deleteParagraph", () => {
    it("should delete paragraph at position", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      const body = mockDoc.getBody();
      const paragraph = createMockParagraph("To delete");
      paragraph.getType = jest.fn(() => "PARAGRAPH");
      body.getNumChildren = jest.fn(() => 3);
      body.getChild = jest.fn(() => paragraph) as any;
      body.removeChild = jest.fn();

      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      deleteParagraph("Projects", "Report", 1);

      expect(body.removeChild).toHaveBeenCalledWith(paragraph);
      expect(mockDoc.saveAndClose).toHaveBeenCalled();
    });

    it("should throw error if position is out of bounds", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      const body = mockDoc.getBody();
      body.getNumChildren = jest.fn(() => 2);

      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      expect(() => {
        deleteParagraph("Projects", "Report", 5);
      }).toThrow("Position 5 is out of bounds");
    });

    it("should throw error if element is not a paragraph", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      const body = mockDoc.getBody();
      const table = { getType: jest.fn(() => "TABLE") };
      body.getNumChildren = jest.fn(() => 2);
      body.getChild = jest.fn(() => table) as any;

      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      expect(() => {
        deleteParagraph("Projects", "Report", 1);
      }).toThrow("Element at position 1 is not a paragraph");
    });

    it("should throw error if folderPath is null", () => {
      expect(() => {
        deleteParagraph(null as any, "Report", 0);
      }).toThrow("Folder path must be a non-empty string");
    });

    it("should throw error if fileName is null", () => {
      expect(() => {
        deleteParagraph("Projects", null as any, 0);
      }).toThrow("File name must be a non-empty string");
    });

    it("should throw error if fileName is undefined", () => {
      expect(() => {
        deleteParagraph("Projects", undefined as any, 0);
      }).toThrow("File name must be a non-empty string");
    });

    it("should throw error if position is null", () => {
      expect(() => {
        deleteParagraph("Projects", "Report", null as any);
      }).toThrow("Position must be a number");
    });

    it("should throw error if position is undefined", () => {
      expect(() => {
        deleteParagraph("Projects", "Report", undefined as any);
      }).toThrow("Position must be a number");
    });

    it("should throw error if position is negative", () => {
      expect(() => {
        deleteParagraph("Projects", "Report", -1);
      }).toThrow("Position must be >= 0");
    });
  });

  describe("getParagraphAtPosition", () => {
    it("should return paragraph at position", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      const body = mockDoc.getBody();
      const paragraph = createMockParagraph("Content");
      paragraph.getType = jest.fn(() => "PARAGRAPH");
      body.getNumChildren = jest.fn(() => 2);
      body.getChild = jest.fn(() => paragraph) as any;

      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      const result = getParagraphAtPosition("Projects", "Report", 1);

      expect(result).toBe(paragraph);
      expect(mockDoc.saveAndClose).toHaveBeenCalled();
    });

    it("should return null if position is out of bounds", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      const body = mockDoc.getBody();
      body.getNumChildren = jest.fn(() => 2);

      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      const result = getParagraphAtPosition("Projects", "Report", 10);

      expect(result).toBeNull();
    });

    it("should return null if element is not a paragraph", () => {
      const rootFolder = createMockFolder("root");
      const folder = createMockFolder("Projects");
      const file = createMockFile("Report", "doc-id");
      folder._addFile(file);
      rootFolder._addFolder(folder);
      (global.DriveApp as any).getRootFolder = jest.fn(() => rootFolder);

      const mockDoc = createMockDocument("Report", "doc-id");
      const body = mockDoc.getBody();
      const table = { getType: jest.fn(() => "TABLE") };
      body.getNumChildren = jest.fn(() => 2);
      body.getChild = jest.fn(() => table) as any;

      (global.DocumentApp as any).openById = jest.fn(() => mockDoc);

      const result = getParagraphAtPosition("Projects", "Report", 1);

      expect(result).toBeNull();
    });

    it("should throw error if folderPath is null", () => {
      expect(() => {
        getParagraphAtPosition(null as any, "Report", 0);
      }).toThrow("Folder path must be a non-empty string");
    });

    it("should throw error if fileName is null", () => {
      expect(() => {
        getParagraphAtPosition("Projects", null as any, 0);
      }).toThrow("File name must be a non-empty string");
    });

    it("should throw error if fileName is undefined", () => {
      expect(() => {
        getParagraphAtPosition("Projects", undefined as any, 0);
      }).toThrow("File name must be a non-empty string");
    });

    it("should throw error if position is null", () => {
      expect(() => {
        getParagraphAtPosition("Projects", "Report", null as any);
      }).toThrow("Position must be a number");
    });

    it("should throw error if position is undefined", () => {
      expect(() => {
        getParagraphAtPosition("Projects", "Report", undefined as any);
      }).toThrow("Position must be a number");
    });

    it("should throw error if position is negative", () => {
      expect(() => {
        getParagraphAtPosition("Projects", "Report", -1);
      }).toThrow("Position must be >= 0");
    });
  });

  describe("formatParagraph", () => {
    it("should format paragraph with default font", () => {
      const paragraph = createMockParagraph();

      formatParagraph(paragraph as any);

      expect(paragraph.setAttributes).toHaveBeenCalledWith({
        FONT_FAMILY: "Roboto",
      });
      expect(paragraph.setAlignment).toHaveBeenCalledWith("JUSTIFY");
    });

    it("should format paragraph with custom font", () => {
      const paragraph = createMockParagraph();

      formatParagraph(paragraph as any, "Arial");

      expect(paragraph.setAttributes).toHaveBeenCalledWith({
        FONT_FAMILY: "Arial",
      });
      expect(paragraph.setAlignment).toHaveBeenCalledWith("JUSTIFY");
    });

    it("should throw error if paragraph is null", () => {
      expect(() => {
        formatParagraph(null as any);
      }).toThrow();
    });

    it("should throw error if paragraph is undefined", () => {
      expect(() => {
        formatParagraph(undefined as any);
      }).toThrow();
    });

    it("should handle empty fontFamily string", () => {
      const paragraph = createMockParagraph();

      formatParagraph(paragraph as any, "");

      expect(paragraph.setAttributes).toHaveBeenCalledWith({
        FONT_FAMILY: "",
      });
    });

    it("should format list item", () => {
      const listItem: any = {
        setAttributes: jest.fn(),
        setAlignment: jest.fn(),
      };

      formatParagraph(listItem, "Arial");

      expect(listItem.setAttributes).toHaveBeenCalledWith({
        FONT_FAMILY: "Arial",
      });
      expect(listItem.setAlignment).toHaveBeenCalledWith("JUSTIFY");
    });
  });
});
