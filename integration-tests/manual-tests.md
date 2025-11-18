# Manual Integration Testing Guide

This guide provides step-by-step instructions for manually testing the deployed Apps Script library.

## Prerequisites

1. Library deployed to Apps Script (run `npm run deploy`)
2. Access to the Apps Script editor
3. Google Drive access for testing file/folder operations
4. Access to a Google Sheets document for testing Sheets operations

## Test Setup

### 1. Open Apps Script Editor

```bash
clasp open
```

### 2. Add Test Functions

Copy the test functions from `integration-tests/test-functions.js` into your Apps Script editor, or deploy them separately.

### 3. Set Library Identifier

In your Apps Script project:

1. Go to **Extensions** → **Apps Script library**
2. Add the library with identifier: `GASToolbox`
3. Use your Script ID from `.clasp.json`

## Running Tests

### Option 1: Using Apps Script Editor

1. Open the Apps Script editor
2. Select a test function from the dropdown (e.g., `testDriveFolderOperations`)
3. Click the play button ▶️
4. Check the execution log for results

### Option 2: Using clasp CLI

```bash
# Run a specific test function
clasp run testDriveFolderOperations

# Run all tests
clasp run runAllIntegrationTests
```

## Test Checklist

### Drive Operations

- [ ] Create folder by path
- [ ] Get folder by ID
- [ ] Get folder by name
- [ ] Check folder exists
- [ ] Create file
- [ ] Get file by ID
- [ ] Find file by path
- [ ] Delete file
- [ ] Delete folder

### Docs Operations

- [ ] Create document
- [ ] Append paragraph
- [ ] Get document content
- [ ] Get paragraph count
- [ ] Append bulleted list
- [ ] Append numbered list
- [ ] Replace text
- [ ] Clear document

### Sheets Operations

- [ ] Create sheet with header
- [ ] Append object
- [ ] Get all objects
- [ ] Get object by index
- [ ] Update object
- [ ] Delete object
- [ ] Count objects
- [ ] Filter objects
- [ ] Sort objects

## Troubleshooting

### Library Not Found

If you get errors about `GASToolbox` not being defined:

1. Ensure the library is added to your Apps Script project
2. Check that the identifier matches (case-sensitive)
3. Verify the Script ID is correct
4. Wait a few seconds for the library to sync

### Permission Errors

Some tests require specific permissions:

- **Drive**: Full access to create/modify/delete files and folders
- **Docs**: Access to create and modify documents
- **Sheets**: Access to create and modify spreadsheets

Grant permissions when prompted, or check **View** → **Show manifest file** for required scopes.

### Test Data Cleanup

Tests create files/folders in a `TEST/Integration/` folder path. You may need to manually clean these up if tests fail partway through.

## Cleanup Script

Run this function to clean up all test data:

```javascript
function cleanupTestData() {
  try {
    const testFolder = GASToolbox.getFolderByName("TEST");
    if (testFolder) {
      GASToolbox.deleteFolder(testFolder);
      console.log("✅ Test data cleaned up");
    }
  } catch (error) {
    console.log("⚠️ No test data found or cleanup failed:", error);
  }
}
```
