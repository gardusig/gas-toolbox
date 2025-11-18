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
   - Source TypeScript files (`src/`, `tests/`)
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
   - Set a library identifier (e.g., `Toolbox`)
   - Click **Add**

3. Use the library in your code:
   ```typescript
   const folder = Toolbox.getOrCreateFolderByPath("MyFolder");
   ```

## Usage Examples

### 1. Folders & Files Examples

Working with Google Drive folders and files:

```typescript
function foldersAndFilesExample() {
  // Create a folder (creates nested structure automatically)
  const folder = Toolbox.getOrCreateFolderByPath("MyProjects/2024/Reports");
  console.log("✅ Folder created:", folder.getName());

  // Get folder by name (searches from root)
  const folderByName = Toolbox.getFolderByName("MyProjects");
  
  // Create a subfolder
  const subFolder = Toolbox.createFolder("Archive", folder);
  
  // Create a document to work with (Google Docs files don't have .docx extension)
  const doc = Toolbox.createDocument("MyProjects/2024/Reports", "My Report");
  console.log("✅ Document created:", doc.getName());

  // Get all files in a folder
  const files = Toolbox.getAllFilesInFolder(folder);
  console.log("📁 Files in folder:", files.length);

  // Find a file by name (no extension needed for Google Docs)
  const file = Toolbox.getFileByName("My Report", folder);
  
  // Find file by path
  const fileByPath = Toolbox.findFile("MyProjects/2024/Reports", "My Report");
  
  // Get files by type (e.g., PDFs) - safe to use even if none exist
  try {
    const pdfFiles = Toolbox.getFilesByType(folder, "application/pdf");
    console.log("📄 PDF files found:", pdfFiles.length);
  } catch (error) {
    console.log("ℹ️ No PDF files or error:", error.message);
  }
  
  // Check if file exists
  const fileExists = Toolbox.checkFileExists("MyProjects/2024/Reports", "My Report");
  console.log("📋 File exists:", fileExists);
  
  // Copy a file
  if (file) {
    const copiedFile = Toolbox.copyFile(file, folder, "Copy of My Report");
    console.log("📋 File copied:", copiedFile.getName());
  }
  
  // Move a file to another folder
  if (file && subFolder) {
    Toolbox.moveFile(file, subFolder);
    console.log("📦 File moved to Archive folder");
  }
  
  // Rename a file
  if (file) {
    Toolbox.renameFile(file, "Renamed Report");
    console.log("✏️ File renamed");
  }
  
  // Delete a file (optional cleanup)
  // if (file) {
  //   Toolbox.deleteFile(file);
  // }
}
```

### 2. Docs Examples

Creating and manipulating Google Documents:

```typescript
function docsExample() {
  // Create a document (creates folders automatically)
  const doc = Toolbox.createDocument("MyProjects/2024/Reports", "Monthly Report");
  console.log("✅ Document created:", doc.getName());

  // Add a heading
  Toolbox.appendParagraphToFile(
    "MyProjects/2024/Reports",
    "Monthly Report",
    "Monthly Report",
    DocumentApp.ParagraphHeading.HEADING1
  );

  // Add a paragraph
  Toolbox.appendParagraphToFile(
    "MyProjects/2024/Reports",
    "Monthly Report",
    "This is a sample report created with Toolbox."
  );

  // Add a bulleted list
  Toolbox.appendBulletedListToFile("MyProjects/2024/Reports", "Monthly Report", [
    "Item 1: First task completed",
    "Item 2: Second task in progress",
    "Item 3: Third task planned"
  ]);

  // Add a numbered list
  Toolbox.appendNumberedListToFile("MyProjects/2024/Reports", "Monthly Report", [
    "First priority",
    "Second priority",
    "Third priority"
  ]);

  // Add more content with text we'll replace later
  Toolbox.appendParagraphToFile(
    "MyProjects/2024/Reports",
    "Monthly Report",
    "Task: Review quarterly results"
  );

  // Insert paragraph at specific position
  Toolbox.insertParagraphAtPosition(
    "MyProjects/2024/Reports",
    "Monthly Report",
    "Introduction paragraph",
    1  // position (0-based index)
  );

  // Get document content
  const content = Toolbox.getDocumentContent("MyProjects/2024/Reports", "Monthly Report");
  console.log("📄 Document content:\n", content);

  // Get paragraph count
  const paragraphCount = Toolbox.getParagraphCount("MyProjects/2024/Reports", "Monthly Report");
  console.log("📊 Paragraph count:", paragraphCount);

  // Replace text in document (regex-compatible)
  const replacedCount = Toolbox.replaceTextInFile(
    "MyProjects/2024/Reports",
    "Monthly Report",
    "Task",
    "Action Item"
  );
  console.log("🔄 Text replaced:", replacedCount, "times");

  // Get a specific paragraph
  const paragraph = Toolbox.getParagraphAtPosition(
    "MyProjects/2024/Reports",
    "Monthly Report",
    0
  );
  
  // Format paragraph
  if (paragraph) {
    Toolbox.formatParagraph(paragraph, "Arial"); // font family
    console.log("✅ Paragraph formatted");
  }

  // Print summary
  console.log("\n📋 Summary:");
  console.log("- Document:", doc.getName());
  console.log("- Paragraphs:", paragraphCount);
  console.log("- Text replacements:", replacedCount);
}
```

### 3. Sheets Examples

Working with Google Sheets data:

```typescript
function sheetsExample() {
  // Get active spreadsheet (or use: Toolbox.getSpreadsheet("spreadsheet-id"))
  const spreadsheet = Toolbox.getSpreadsheet();
  if (!spreadsheet) {
    console.log("⚠️ No active spreadsheet. Open a Google Sheet and try again.");
    return;
  }

  // Create a sheet with header
  const header = ["name", "age", "email", "status"];
  const sheet = Toolbox.createSheet("Employees", header, spreadsheet);
  console.log("✅ Sheet created:", sheet.getName());

  // Add a single employee
  Toolbox.appendObject("Employees", header, {
    name: "John Doe",
    age: 30,
    email: "john@example.com",
    status: "active"
  });

  // Add multiple employees
  Toolbox.appendObjects("Employees", header, [
    { name: "Alice Smith", age: 28, email: "alice@example.com", status: "active" },
    { name: "Bob Johnson", age: 35, email: "bob@example.com", status: "inactive" },
    { name: "Carol Brown", age: 32, email: "carol@example.com", status: "active" }
  ]);

  // Read all data
  const allEmployees = Toolbox.getAllObjects("Employees");
  console.log("👥 All employees:", JSON.stringify(allEmployees, null, 2));

  // Filter data
  const activeEmployees = Toolbox.filterObjects(
    "Employees",
    emp => emp.status === "active"
  );
  console.log("✅ Active employees:", activeEmployees.length);

  // Count records
  const totalEmployees = Toolbox.countObjects("Employees");
  console.log("📊 Total employees:", totalEmployees);

  // Find a specific employee
  const employee = Toolbox.findObject(
    "Employees",
    emp => emp.email === "john@example.com"
  );
  console.log("🔍 Found employee:", employee);

  // Update an employee (find index first, then update)
  if (employee) {
    const index = Toolbox.findObjectIndex("Employees", emp => emp.email === "john@example.com");
    if (index !== -1) {
      Toolbox.updateObject("Employees", header, index, {
        name: "John Doe",
        age: 31, // Updated age
        email: "john@example.com",
        status: "active"
      });
      console.log("✅ Employee updated at index", index);
    }
  }

  // Aggregate data
  const averageAge = Toolbox.average("Employees", "age");
  console.log("📈 Average age:", averageAge);

  const maxAge = Toolbox.max("Employees", "age");
  console.log("📈 Max age:", maxAge);

  // Group by status
  const byStatus = Toolbox.groupBy("Employees", "status");
  console.log("📊 Employees by status:", byStatus);

  // Get distinct values
  const statuses = Toolbox.getDistinctValues("Employees", "status");
  console.log("📋 Unique statuses:", statuses);

  // Print summary
  console.log("\n📋 Summary:");
  console.log("- Total employees:", totalEmployees);
  console.log("- Active employees:", activeEmployees.length);
  console.log("- Average age:", averageAge);
  console.log("- Statuses:", statuses.join(", "));
}
```

### Complete API Reference

#### Drive Functions

##### Folders

```typescript
// Create or get folder by path (creates nested structure automatically)
const folder = Toolbox.getOrCreateFolderByPath("Projects/2024/January");

// Get folder by ID
const folderById = Toolbox.getFolderById("folder-id-123");

// Get folder by name (searches from root)
const folderByName = Toolbox.getFolderByName("MyFolder");

// Create folder in a specific location
const newFolder = Toolbox.createFolder("NewFolder", folder);

// Get all folders in a folder
const folders = Toolbox.getAllFoldersInFolder(folder);

// Check if folder exists
const exists = Toolbox.checkFolderExists("Projects/2024");

// Delete folder
Toolbox.deleteFolder(folder);

// Rename folder
Toolbox.renameFolder(folder, "NewName");

// Move folder
const targetFolder = Toolbox.getFolderById("target-folder-id");
if (targetFolder) {
  Toolbox.moveFolder(folder, targetFolder);
}
```

##### Files

```typescript
// First, get or create a folder
const folder = Toolbox.getOrCreateFolderByPath("Projects/2024");

// Create a document first (Google Docs files don't have .docx extension)
const doc = Toolbox.createDocument("Projects/2024", "My Report");

// Find file by name in a folder (no extension needed for Google Docs)
const file = Toolbox.getFileByName("My Report", folder);

// Find file by path (throws error if not found)
const fileByPath = Toolbox.findFile("Projects/2024", "My Report");

// Get file by ID
const fileById = Toolbox.getFileById("file-id-123");

// Get files by type (MIME type)
const pdfFiles = Toolbox.getFilesByType(folder, "application/pdf");

// Get all files in a folder
const files = Toolbox.getAllFilesInFolder(folder);

// Check if file exists
const fileExists = Toolbox.checkFileExists("Projects/2024", "My Report");

// Copy file
const copiedFile = Toolbox.copyFile(file, folder, "Copy of My Report");

// Move file
Toolbox.moveFile(file, folder);

// Rename file
Toolbox.renameFile(file, "Renamed Report");

// Delete file
Toolbox.deleteFile(file);
```

#### Docs Functions

##### Document Management

```typescript
// Create document (creates folders automatically)
const doc = Toolbox.createDocument("Reports/2024", "Monthly Summary");

// Get document content as text
const content = Toolbox.getDocumentContent(
  "Reports/2024",
  "Monthly Summary"
);

// Clear all document content
Toolbox.clearDocument("Reports/2024", "Monthly Summary");

// Ensure folder exists
const folder = Toolbox.ensureFolder("Projects/2024/Q1");
```

##### Paragraphs

```typescript
// Append paragraph
Toolbox.appendParagraphToFile(
  "Reports/2024",
  "Monthly Summary",
  "Hello World!"
);

// Append paragraph with heading
Toolbox.appendParagraphToFile(
  "Reports/2024",
  "Monthly Summary",
  "Title",
  DocumentApp.ParagraphHeading.HEADING1
);

// Insert paragraph at position
Toolbox.insertParagraphAtPosition(
  "Reports/2024",
  "Monthly Summary",
  "First paragraph",
  0  // position (0-based index)
);

// Get paragraph at position
const paragraph = Toolbox.getParagraphAtPosition(
  "Reports/2024",
  "Monthly Summary",
  0
);

// Get paragraph count
const count = Toolbox.getParagraphCount("Reports/2024", "Monthly Summary");

// Delete paragraph at position
Toolbox.deleteParagraph("Reports/2024", "Monthly Summary", 0);
```

##### Lists

```typescript
// Append bulleted list
Toolbox.appendBulletedListToFile("Reports/2024", "Monthly Summary", [
  "Item 1",
  "Item 2",
  "Item 3",
]);

// Append numbered list
Toolbox.appendNumberedListToFile("Reports/2024", "Monthly Summary", [
  "First",
  "Second",
  "Third",
]);
```

##### Text & Elements

```typescript
// Replace text (regex-compatible)
const count = Toolbox.replaceTextInFile(
  "Templates",
  "Invoice Template",
  "{{customer_name}}",
  "Acme Corp"
);

// Insert table
const table = Toolbox.insertTable(
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
const image = Toolbox.insertImage(
  "Reports/2024",
  "Monthly Summary",
  "https://example.com/image.png",
  300,
  200
);

// Format paragraph
Toolbox.formatParagraph(paragraph, "Arial");
```

#### Sheets Functions

##### Spreadsheet Utilities

```typescript
// Get active spreadsheet
const spreadsheet = Toolbox.getSpreadsheet();

// Get spreadsheet by ID or URL
const spreadsheetById = Toolbox.getSpreadsheet("spreadsheet-id-123");
const spreadsheetByUrl = Toolbox.getSpreadsheet(
  "https://docs.google.com/spreadsheets/d/..."
);

// Create sheet with header
const sheet = Toolbox.createSheet(
  "MySheet",
  ["name", "age", "email"],
  spreadsheet
);

// Get sheet by name
const sheet = Toolbox.getSheet("MySheet", "spreadsheet-id-123");
```

##### Write Operations

```typescript
const header = ["name", "age", "email"];

// Append single object
Toolbox.appendObject("MySheet", header, {
  name: "John Doe",
  age: 30,
  email: "john@example.com",
});

// Append multiple objects
Toolbox.appendObjects("MySheet", header, [
  { name: "Alice", age: 28, email: "alice@example.com" },
  { name: "Bob", age: 35, email: "bob@example.com" },
]);

// Update object at row index (0-based, 0 = first data row)
Toolbox.updateObject("MySheet", header, 0, {
  name: "John Doe",
  age: 31,
  email: "john@example.com",
});

// Update multiple objects
Toolbox.updateObjects("MySheet", header, [
  { rowIndex: 0, obj: { name: "John", age: 31, email: "john@example.com" } },
  { rowIndex: 1, obj: { name: "Jane", age: 29, email: "jane@example.com" } },
]);

// Delete object at row index
Toolbox.deleteObject("MySheet", header, 0);

// Delete multiple objects
Toolbox.deleteObjects("MySheet", header, [0, 1, 2]);

// Delete objects by filter
const deleted = Toolbox.deleteObjectsByFilter(
  "MySheet",
  header,
  obj => obj.age < 30
);

// Upsert object (update if exists, insert if new)
const rowIndex = Toolbox.upsertObject("MySheet", header, "email", {
  email: "john@example.com",
  name: "John Doe",
  age: 31,
});

// Upsert multiple objects
const upserted = Toolbox.upsertObjects("MySheet", header, "email", [
  { email: "alice@example.com", name: "Alice", age: 28 },
  { email: "bob@example.com", name: "Bob", age: 35 },
]);

// Replace all data
Toolbox.replaceAll("MySheet", header, [
  { name: "New User 1", age: 25, email: "user1@example.com" },
  { name: "New User 2", age: 30, email: "user2@example.com" },
]);

// Clear all data
Toolbox.clearAll("MySheet", header);
```

##### Read Operations

```typescript
// Get all objects
const allUsers = Toolbox.getAllObjects("MySheet");

// Get object at row index
const user = Toolbox.getObject("MySheet", 0);

// Get batch of objects
const batch = Toolbox.getObjectBatch("MySheet", 0, 10);

// Get header map
const headerMap = Toolbox.getHeaderMap("MySheet");

// Filter objects
const activeUsers = Toolbox.filterObjects(
  "MySheet",
  obj => obj.active === true
);

// Find first matching object
const user = Toolbox.findObject(
  "MySheet",
  obj => obj.email === "john@example.com"
);

// Find index of matching object
const index = Toolbox.findObjectIndex(
  "MySheet",
  obj => obj.email === "john@example.com"
);

// Count objects
const count = Toolbox.countObjects("MySheet");

// Get first object
const first = Toolbox.getFirst("MySheet");

// Get last object
const last = Toolbox.getLast("MySheet");

// Check if object exists
const exists = Toolbox.exists(
  "MySheet",
  obj => obj.email === "john@example.com"
);

// Sort objects
const sorted = Toolbox.sortObjects("MySheet", "age", true); // ascending
const sortedDesc = Toolbox.sortObjects("MySheet", "age", false); // descending
const sortedMulti = Toolbox.sortObjects("MySheet", ["status", "age"], true);

// Get paginated results
const page = Toolbox.getObjectsPaginated("MySheet", 1, 10); // page 1, 10 per page
console.log(
  `Page ${page.page} of ${page.totalPages}, showing ${page.data.length} of ${page.total}`
);

// Quick filter by column
const activeUsers = Toolbox.filterByColumn("MySheet", "status", "active");
```

##### Aggregations

```typescript
// Sum column values
const totalRevenue = Toolbox.sum("Sales", "amount");

// Average column values
const avgSale = Toolbox.average("Sales", "amount");

// Min value
const minSale = Toolbox.min("Sales", "amount");

// Max value
const maxSale = Toolbox.max("Sales", "amount");

// Group by column
const byCategory = Toolbox.groupBy("Sales", "category");
Object.keys(byCategory).forEach(category => {
  console.log(`${category}: ${byCategory[category].length} items`);
});

// Get distinct values
const categories = Toolbox.getDistinctValues("Sales", "category");
```

##### Formatting

```typescript
// Trim empty rows and columns
Toolbox.trim("MySheet");

// Trim empty rows only
Toolbox.trimRows("MySheet");

// Trim empty columns only
Toolbox.trimColumns("MySheet");
```

## License

MIT License - Copyright (c) 2024 Gustavo Gardusi
