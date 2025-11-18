// Jest setup file - mocks Google Apps Script APIs

// Mock Logger
global.Logger = {
  log: jest.fn(),
} as any;

// Mock DriveApp
global.DriveApp = {
  getRootFolder: jest.fn(),
  getFileById: jest.fn(),
  getFolderById: jest.fn(),
  removeFile: jest.fn(),
  removeFolder: jest.fn(),
} as any;

// Mock DocumentApp
global.DocumentApp = {
  create: jest.fn(),
  openById: jest.fn(),
  Attribute: {
    FONT_FAMILY: "FONT_FAMILY",
  },
  HorizontalAlignment: {
    JUSTIFY: "JUSTIFY",
  },
  ElementType: {
    PARAGRAPH: "PARAGRAPH",
    LIST_ITEM: "LIST_ITEM",
    TABLE: "TABLE",
  },
  ParagraphHeading: {
    HEADING1: "HEADING1",
    HEADING2: "HEADING2",
    HEADING3: "HEADING3",
    NORMAL: "NORMAL",
  },
  GlyphType: {
    BULLET: "BULLET",
    NUMBER: "NUMBER",
  },
} as any;

// Mock SpreadsheetApp
global.SpreadsheetApp = {
  getActiveSpreadsheet: jest.fn(),
  openById: jest.fn(),
  openByUrl: jest.fn(),
} as any;

// Sheets module uses function-based API, no namespaces needed
