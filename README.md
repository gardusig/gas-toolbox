# GAS Toolbox

[![Coverage](https://img.shields.io/badge/coverage-100%25-brightgreen.svg)](https://github.com/gardusig/gas-toolbox/actions/workflows/update-coverage-badge.yml)

A comprehensive Google Apps Script library providing utilities for Drive, Docs, and Sheets.

## Setup

### Prerequisites

- Node.js 18+ and npm
- Google account with Apps Script access
- `clasp` (Google Apps Script CLI tool)

### First-Time Setup with Clasp

1. **Clone and install:**

   ```bash
   git clone https://github.com/gardusig/gas-toolbox.git
   cd gas-toolbox
   npm install
   ```

2. **Install clasp globally:**

   ```bash
   npm install -g @google/clasp
   ```

3. **Login to Google:**

   ```bash
   clasp login
   ```

   This will open your browser to authorize clasp with your Google account.

4. **Create a new Apps Script project:**

   ```bash
   clasp create --type standalone --title "GAS Toolbox"
   ```

   This will:
   - Create a new standalone Apps Script project in your Google Drive
   - Generate a `.clasp.json` file with your project ID (this file is gitignored as it's user-specific)

5. **Build the project:**

   ```bash
   npm run build
   ```

   This compiles TypeScript to JavaScript in the `dist/` directory and removes ES module syntax for Apps Script compatibility.

6. **Push to Apps Script:**

   ```bash
   npm run deploy
   ```

   The deployment script (`scripts/deploy.js`) will:
   - Validate that clasp is installed and configured
   - Build the project (TypeScript → JavaScript)
   - Verify build output exists
   - Check what files will be pushed
   - Deploy to Apps Script with force overwrite

   **Available deployment commands:**
   - `npm run deploy` - Build and deploy (force overwrite)
   - `npm run deploy:check` - Build and deploy (ask for confirmation)
   - `npm run deploy:skip-build` - Deploy without rebuilding (use with caution)
   - `npm run deploy:tests` - Deploy integration test functions

7. **Publish a version (optional):**

   ```bash
   # Publish next incremental version
   npm run publish

   # Publish a specific version number
   npm run publish 5

   # Publish a specific version with custom description
   npm run publish 5 "v5.0.0 - Added new features"
   ```

   The publish script will:
   - Get the latest version number from Apps Script
   - Build and deploy the code
   - Create a new version with the specified number (or auto-increment if not specified)

   The `.claspignore` file ensures only the compiled `dist/` files and `appsscript.json` are pushed to Apps Script, excluding:
   - Source TypeScript files (`src/`, `tests/`, `integration-tests/`)
   - Build configuration files
   - Node modules and dependencies
   - CI/CD configurations
   - Documentation and other development files

### Configuration Files

- **`.clasp.json`** - Contains your Apps Script project ID (user-specific, gitignored)
- **`.claspignore`** - Specifies what files clasp should exclude when pushing (already configured)
- **`appsscript.json`** - Apps Script manifest file (pushed to Apps Script)

### Using as a Library

1. Get your Script ID from `.clasp.json` or from the Apps Script editor URL:

   ```
   https://script.google.com/home/projects/YOUR_SCRIPT_ID/edit
   ```

2. In your Apps Script project:
   - Go to **Extensions** → **Apps Script library**
   - Click **Add a library**
   - Paste your Script ID
   - Set a library identifier (e.g., `GasToolbox`)
   - Click **Add**

3. Use the library in your code:
   ```typescript
   const folder = GasToolbox.getOrCreateFolderByPath("MyFolder");
   ```

### Integration Testing

The `integration-tests/` folder contains test functions that can be run in your Apps Script project to verify the deployed library works correctly.

**Setting up integration tests:**

1. Deploy your library: `npm run deploy`
2. Open your Apps Script project: `clasp open`
3. Copy test functions from `integration-tests/test-functions.js` into your Apps Script editor
4. Run test functions in the editor or via `clasp run functionName`

**Available test functions:**
- `testDriveFolderOperations()` - Tests Drive folder operations
- `testDriveFileOperations()` - Tests Drive file operations
- `testDocsOperations()` - Tests Docs operations
- `testSheetsOperations()` - Tests Sheets operations (requires active spreadsheet)
- `runAllIntegrationTests()` - Runs all integration tests

For detailed instructions, see `integration-tests/README.md` and `integration-tests/manual-tests.md`.

## Usage Examples

### Drive Functions

#### Folders

```typescript
// Create or get folder by path (creates nested structure automatically)
const folder = GasToolbox.getOrCreateFolderByPath("Projects/2024/January");

// Get folder by ID
const folderById = GasToolbox.getFolderById("folder-id-123");

// Get folder by name (searches from root)
const folderByName = GasToolbox.getFolderByName("MyFolder");

// Create folder in a specific location
const newFolder = GasToolbox.createFolder("NewFolder", folder);

// Get all folders in a folder
const folders = GasToolbox.getAllFoldersInFolder(folder);

// Check if folder exists
const exists = GasToolbox.checkFolderExists("Projects/2024");

// Delete folder
GasToolbox.deleteFolder(folder);

// Rename folder
GasToolbox.renameFolder(folder, "NewName");

// Move folder
const targetFolder = GasToolbox.getFolderById("target-folder-id");
GasToolbox.moveFolder(folder, targetFolder);
```

#### Files

```typescript
// Find file by name in a folder
const file = GasToolbox.getFileByName("Report.docx", folder);

// Find file by path (throws error if not found)
const fileByPath = GasToolbox.findFile("Projects/2024", "Report.docx");

// Get file by ID
const fileById = GasToolbox.getFileById("file-id-123");

// Get files by type (MIME type)
const pdfFiles = GasToolbox.getFilesByType("application/pdf", folder);

// Get all files in a folder
const files = GasToolbox.getAllFilesInFolder(folder);

// Check if file exists
const fileExists = GasToolbox.checkFileExists("Projects/2024", "Report.docx");

// Copy file
const copiedFile = GasToolbox.copyFile(file, folder, "Copy of Report.docx");

// Move file
GasToolbox.moveFile(file, folder);

// Rename file
GasToolbox.renameFile(file, "NewName.docx");

// Delete file
GasToolbox.deleteFile(file);
```

### Docs Functions

#### Document Management

```typescript
// Create document (creates folders automatically)
const doc = GasToolbox.createDocument("Reports/2024", "Monthly Summary");

// Get document content as text
const content = GasToolbox.getDocumentContent(
  "Reports/2024",
  "Monthly Summary"
);

// Clear all document content
GasToolbox.clearDocument("Reports/2024", "Monthly Summary");

// Ensure folder exists
const folder = GasToolbox.ensureFolder("Projects/2024/Q1");
```

#### Paragraphs

```typescript
// Append paragraph
GasToolbox.appendParagraphToFile(
  "Reports/2024",
  "Monthly Summary",
  "Hello World!"
);

// Append paragraph with heading
GasToolbox.appendParagraphToFile(
  "Reports/2024",
  "Monthly Summary",
  "Title",
  DocumentApp.ParagraphHeading.HEADING1
);

// Insert paragraph at position
GasToolbox.insertParagraphAtPosition(
  "Reports/2024",
  "Monthly Summary",
  0,
  "First paragraph"
);

// Get paragraph at position
const paragraph = GasToolbox.getParagraphAtPosition(
  "Reports/2024",
  "Monthly Summary",
  0
);

// Get paragraph count
const count = GasToolbox.getParagraphCount("Reports/2024", "Monthly Summary");

// Delete paragraph at position
GasToolbox.deleteParagraph("Reports/2024", "Monthly Summary", 0);
```

#### Lists

```typescript
// Append bulleted list
GasToolbox.appendBulletedListToFile("Reports/2024", "Monthly Summary", [
  "Item 1",
  "Item 2",
  "Item 3",
]);

// Append numbered list
GasToolbox.appendNumberedListToFile("Reports/2024", "Monthly Summary", [
  "First",
  "Second",
  "Third",
]);
```

#### Text & Elements

```typescript
// Replace text (regex-compatible)
const count = GasToolbox.replaceTextInFile(
  "Templates",
  "Invoice Template",
  "{{customer_name}}",
  "Acme Corp"
);

// Insert table
const table = GasToolbox.insertTable(
  "Reports/2024",
  "Monthly Summary",
  [
    ["Name", "Age", "Email"],
    ["John", "30", "john@example.com"],
  ],
  200,
  100
);

// Insert image
const image = GasToolbox.insertImage(
  "Reports/2024",
  "Monthly Summary",
  "https://example.com/image.png",
  300,
  200
);

// Format paragraph
GasToolbox.formatParagraph(paragraph, "Arial");
```

### Sheets Functions

#### Spreadsheet Utilities

```typescript
// Get active spreadsheet
const spreadsheet = GasToolbox.getSpreadsheet();

// Get spreadsheet by ID or URL
const spreadsheetById = GasToolbox.getSpreadsheet("spreadsheet-id-123");
const spreadsheetByUrl = GasToolbox.getSpreadsheet(
  "https://docs.google.com/spreadsheets/d/..."
);

// Create sheet with header
const sheet = GasToolbox.createSheet(
  "MySheet",
  ["name", "age", "email"],
  spreadsheet
);

// Get sheet by name
const sheet = GasToolbox.getSheet("MySheet", "spreadsheet-id-123");
```

#### Write Operations

```typescript
const header = ["name", "age", "email"];

// Append single object
GasToolbox.appendObject("MySheet", header, {
  name: "John Doe",
  age: 30,
  email: "john@example.com",
});

// Append multiple objects
GasToolbox.appendObjects("MySheet", header, [
  { name: "Alice", age: 28, email: "alice@example.com" },
  { name: "Bob", age: 35, email: "bob@example.com" },
]);

// Update object at row index (0-based, 0 = first data row)
GasToolbox.updateObject("MySheet", header, 0, {
  name: "John Doe",
  age: 31,
  email: "john@example.com",
});

// Update multiple objects
GasToolbox.updateObjects("MySheet", header, [
  { rowIndex: 0, obj: { name: "John", age: 31, email: "john@example.com" } },
  { rowIndex: 1, obj: { name: "Jane", age: 29, email: "jane@example.com" } },
]);

// Delete object at row index
GasToolbox.deleteObject("MySheet", header, 0);

// Delete multiple objects
GasToolbox.deleteObjects("MySheet", header, [0, 1, 2]);

// Delete objects by filter
const deleted = GasToolbox.deleteObjectsByFilter(
  "MySheet",
  header,
  obj => obj.age < 30
);

// Upsert object (update if exists, insert if new)
const rowIndex = GasToolbox.upsertObject("MySheet", header, "email", {
  email: "john@example.com",
  name: "John Doe",
  age: 31,
});

// Upsert multiple objects
const upserted = GasToolbox.upsertObjects("MySheet", header, "email", [
  { email: "alice@example.com", name: "Alice", age: 28 },
  { email: "bob@example.com", name: "Bob", age: 35 },
]);

// Replace all data
GasToolbox.replaceAll("MySheet", header, [
  { name: "New User 1", age: 25, email: "user1@example.com" },
  { name: "New User 2", age: 30, email: "user2@example.com" },
]);

// Clear all data
GasToolbox.clearAll("MySheet", header);
```

#### Read Operations

```typescript
// Get all objects
const allUsers = GasToolbox.getAllObjects("MySheet");

// Get object at row index
const user = GasToolbox.getObject("MySheet", 0);

// Get batch of objects
const batch = GasToolbox.getObjectBatch("MySheet", 0, 10);

// Get header map
const headerMap = GasToolbox.getHeaderMap("MySheet");

// Filter objects
const activeUsers = GasToolbox.filterObjects(
  "MySheet",
  obj => obj.active === true
);

// Find first matching object
const user = GasToolbox.findObject(
  "MySheet",
  obj => obj.email === "john@example.com"
);

// Find index of matching object
const index = GasToolbox.findObjectIndex(
  "MySheet",
  obj => obj.email === "john@example.com"
);

// Count objects
const count = GasToolbox.countObjects("MySheet");

// Get first object
const first = GasToolbox.getFirst("MySheet");

// Get last object
const last = GasToolbox.getLast("MySheet");

// Check if object exists
const exists = GasToolbox.exists(
  "MySheet",
  obj => obj.email === "john@example.com"
);

// Sort objects
const sorted = GasToolbox.sortObjects("MySheet", "age", true); // ascending
const sortedDesc = GasToolbox.sortObjects("MySheet", "age", false); // descending
const sortedMulti = GasToolbox.sortObjects("MySheet", ["status", "age"], true);

// Get paginated results
const page = GasToolbox.getObjectsPaginated("MySheet", 1, 10); // page 1, 10 per page
console.log(
  `Page ${page.page} of ${page.totalPages}, showing ${page.data.length} of ${page.total}`
);

// Quick filter by column
const activeUsers = GasToolbox.filterByColumn("MySheet", "status", "active");
```

#### Aggregations

```typescript
// Sum column values
const totalRevenue = GasToolbox.sum("Sales", "amount");

// Average column values
const avgSale = GasToolbox.average("Sales", "amount");

// Min value
const minSale = GasToolbox.min("Sales", "amount");

// Max value
const maxSale = GasToolbox.max("Sales", "amount");

// Group by column
const byCategory = GasToolbox.groupBy("Sales", "category");
Object.keys(byCategory).forEach(category => {
  console.log(`${category}: ${byCategory[category].length} items`);
});

// Get distinct values
const categories = GasToolbox.getDistinctValues("Sales", "category");
```

#### Formatting

```typescript
// Trim empty rows and columns
GasToolbox.trim("MySheet");

// Trim empty rows only
GasToolbox.trimRows("MySheet");

// Trim empty columns only
GasToolbox.trimColumns("MySheet");
```

## License

MIT License - Copyright (c) 2024 Gustavo Gardusi
