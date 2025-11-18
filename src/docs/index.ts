// Document management
export {
  ensureFolder,
  createDocument,
  clearDocument,
  getDocumentContent,
} from "./document";

// Paragraph operations
export {
  appendParagraphToFile,
  insertParagraphAtPosition,
  deleteParagraph,
  getParagraphAtPosition,
  getParagraphCount,
} from "./paragraphs";

// List operations
export {
  appendBulletedListToFile,
  appendNumberedListToFile,
} from "./lists";

// Text manipulation
export {
  replaceTextInFile,
} from "./text";

// Element operations
export {
  insertTable,
  insertImage,
} from "./elements";

// Formatting
export {
  formatParagraph,
} from "./formatting";

