// Helper functions to create mock objects for Google Apps Script APIs

export function createMockFolder(
  name: string,
  id: string = `folder-${name}`
): any {
  const folders: any[] = [];
  const files: any[] = [];

  return {
    getId: jest.fn(() => id),
    getName: jest.fn(() => name),
    getUrl: jest.fn(() => `https://drive.google.com/drive/folders/${id}`),
    getFoldersByName: jest.fn((folderName: string) => {
      const matching = folders.filter(f => f.getName() === folderName);
      let index = 0;
      return {
        hasNext: jest.fn(() => index < matching.length),
        next: jest.fn(() => matching[index++]),
      };
    }),
    createFolder: jest.fn((folderName: string): any => {
      const newFolder: any = createMockFolder(
        folderName,
        `folder-${folderName}-${Date.now()}`
      );
      folders.push(newFolder);
      return newFolder;
    }),
    getFilesByName: jest.fn((fileName: string) => {
      const matching = files.filter(f => f.getName() === fileName);
      let index = 0;
      return {
        hasNext: jest.fn(() => index < matching.length),
        next: jest.fn(() => matching[index++]),
      };
    }),
    getFiles: jest.fn(() => {
      let index = 0;
      return {
        hasNext: jest.fn(() => index < files.length),
        next: jest.fn(() => files[index++]),
      };
    }),
    getFolders: jest.fn(() => {
      let index = 0;
      return {
        hasNext: jest.fn(() => index < folders.length),
        next: jest.fn(() => folders[index++]),
      };
    }),
    setName: jest.fn(),
    moveTo: jest.fn(),
    _addFolder: (folder: any) => folders.push(folder),
    _addFile: (file: any) => files.push(file),
  };
}

export function createMockFile(
  name: string,
  id: string = `file-${name}`,
  url: string = `https://drive.google.com/file/d/${id}`
) {
  return {
    getId: jest.fn(() => id),
    getName: jest.fn(() => name),
    getUrl: jest.fn(() => url),
    moveTo: jest.fn(),
    setName: jest.fn(),
    makeCopy: jest.fn((newName: string, _destination: any): any => {
      return createMockFile(newName, `copied-${id}`, url);
    }),
  };
}

export function createMockDocument(
  name: string,
  id: string = `doc-${name}`,
  url: string = `https://docs.google.com/document/d/${id}`
) {
  const paragraphs: any[] = [];

  const mockDoc: any = {
    getId: jest.fn(() => id),
    getName: jest.fn(() => name),
    getUrl: jest.fn(() => url),
    saveAndClose: jest.fn(),
    _getParagraphs: () => paragraphs,
  };

  const body: any = {
    appendParagraph: jest.fn((text: string) => {
      const paragraph = createMockParagraph(text);
      paragraph.getParent = jest.fn(() => body);
      paragraphs.push(paragraph);
      return paragraph;
    }),
    appendListItem: jest.fn((text: string) => {
      const listItem = createMockListItem(text);
      listItem.getParent = jest.fn(() => body);
      paragraphs.push(listItem);
      return listItem;
    }),
    insertParagraph: jest.fn((index: number, text: string) => {
      const paragraph = createMockParagraph(text);
      paragraph.getParent = jest.fn(() => body);
      paragraphs.splice(index, 0, paragraph);
      return paragraph;
    }),
    appendTable: jest.fn(() => {
      const table: any = {
        appendTableRow: jest.fn(() => ({
          appendTableCell: jest.fn(() => ({
            setText: jest.fn(),
          })),
        })),
      };
      paragraphs.push(table);
      return table;
    }),
    appendImage: jest.fn((_blob: any) => {
      const image: any = {
        setWidth: jest.fn(),
        setHeight: jest.fn(),
      };
      paragraphs.push(image);
      return image;
    }),
    removeChild: jest.fn((child: any) => {
      const index = paragraphs.indexOf(child);
      if (index !== -1) {
        paragraphs.splice(index, 1);
      }
    }),
    clear: jest.fn(() => {
      paragraphs.length = 0;
    }),
    getNumChildren: jest.fn(() => paragraphs.length),
    getChild: jest.fn((index: number) => {
      const child = paragraphs[index];
      if (child && typeof child.getParent !== 'function') {
        child.getParent = jest.fn(() => body);
      }
      return child;
    }),
    getText: jest.fn(() =>
      paragraphs
        .map(p => {
          if (p && typeof p.getText === "function") {
            return p.getText();
          }
          return "";
        })
        .join("\n")
    ),
    getType: jest.fn(() => "BODY_SECTION"),
    asBody: jest.fn(() => body),
    getParent: jest.fn(() => mockDoc),
  };

  mockDoc.getBody = jest.fn(() => body);
  
  return mockDoc;
}

export function createMockParagraph(text: string = "") {
  const mockParagraph: any = {
    getText: jest.fn(() => text),
    setText: jest.fn((newText: string) => {
      text = newText;
    }),
    setHeading: jest.fn(function (_heading: any) {
      return this;
    }),
    setAttributes: jest.fn(),
    setAlignment: jest.fn(),
    getType: jest.fn(() => "PARAGRAPH"),
    asParagraph: jest.fn(function () {
      return this;
    }),
    getParent: jest.fn(),
  };
  return mockParagraph;
}

export function createMockListItem(text: string = "") {
  const mockListItem: any = {
    getText: jest.fn(() => text),
    setText: jest.fn((newText: string) => {
      text = newText;
    }),
    setGlyphType: jest.fn(),
    setHeading: jest.fn(function (_heading: any) {
      return this;
    }),
    setAttributes: jest.fn(),
    setAlignment: jest.fn(),
    getType: jest.fn(() => "LIST_ITEM"),
    asParagraph: jest.fn(function () {
      return this;
    }),
    asListItem: jest.fn(function () {
      return this;
    }),
    getParent: jest.fn(),
  };
  return mockListItem;
}

export function createMockSheet(name: string, header: string[] = []) {
  const rows: any[][] = header.length > 0 ? [header] : [];
  let lastRow = rows.length;
  let lastColumn = header.length;

  return {
    getName: jest.fn(() => name),
    appendRow: jest.fn((values: any[]) => {
      rows.push(values);
      lastRow = rows.length;
      lastColumn = Math.max(lastColumn, values.length);
    }),
    getDataRange: jest.fn(() => ({
      getValues: jest.fn(() => rows),
    })),
    getRange: jest.fn(
      (row: number, col: number, numRows?: number, _numCols?: number) => {
        const rangeRows = rows.slice(row - 1, row - 1 + (numRows || 1));
        return {
          setValues: jest.fn((values: any[][]) => {
            values.forEach((rowValues, i) => {
              if (rows[row - 1 + i]) {
                rowValues.forEach((val, j) => {
                  rows[row - 1 + i][col - 1 + j] = val;
                });
              }
            });
          }),
          getValues: jest.fn(() => rangeRows),
        };
      }
    ),
    deleteRow: jest.fn((row: number) => {
      rows.splice(row - 1, 1);
      lastRow = rows.length;
    }),
    deleteRows: jest.fn((startRow: number, numRows: number) => {
      rows.splice(startRow - 1, numRows);
      lastRow = rows.length;
    }),
    deleteColumns: jest.fn((startCol: number, numCols: number) => {
      rows.forEach(row => row.splice(startCol - 1, numCols));
      lastColumn = Math.max(0, lastColumn - numCols);
    }),
    getLastRow: jest.fn(() => lastRow),
    getLastColumn: jest.fn(() => lastColumn),
    getMaxRows: jest.fn(() => 1000),
    getMaxColumns: jest.fn(() => 100),
    clear: jest.fn(() => {
      rows.length = 0;
      lastRow = 0;
      lastColumn = 0;
    }),
    _getRows: () => rows,
    _setRows: (newRows: any[][]) => {
      rows.length = 0;
      rows.push(...newRows);
      lastRow = rows.length;
      if (rows.length > 0) {
        lastColumn = Math.max(...rows.map(r => r.length));
      }
    },
  };
}

export function createMockSpreadsheet(name: string = "Test Spreadsheet") {
  const sheets: any[] = [];

  return {
    getId: jest.fn(() => `spreadsheet-${name}`),
    getName: jest.fn(() => name),
    getSheetByName: jest.fn((sheetName: string) => {
      return sheets.find(s => s.getName() === sheetName) || null;
    }),
    insertSheet: jest.fn((sheetName: string) => {
      const sheet = createMockSheet(sheetName);
      sheets.push(sheet);
      return sheet;
    }),
    _addSheet: (sheet: any) => sheets.push(sheet),
    _getSheets: () => sheets,
  };
}
