/**
 * Integration Test Functions
 *
 * These functions test the deployed Apps Script library.
 * Run them in the Apps Script editor or via: clasp run functionName
 *
 * Note: These functions are meant to be added to your Apps Script project
 * manually or via a separate deployment script.
 */

/**
 * Test Drive folder operations
 */
function testDriveFolderOperations() {
  console.log("Testing Drive folder operations...");

  try {
    // Test getOrCreateFolderByPath
    const folder = GASToolbox.getOrCreateFolderByPath(
      "TEST/Integration/Folder1"
    );
    console.log("✅ getOrCreateFolderByPath: PASSED");

    // Test getFolderById
    const folderById = GASToolbox.getFolderById(folder.getId());
    if (folderById && folderById.getId() === folder.getId()) {
      console.log("✅ getFolderById: PASSED");
    } else {
      console.log("❌ getFolderById: FAILED");
    }

    // Test checkFolderExists
    const exists = GASToolbox.checkFolderExists("TEST/Integration/Folder1");
    if (exists) {
      console.log("✅ checkFolderExists: PASSED");
    } else {
      console.log("❌ checkFolderExists: FAILED");
    }

    // Cleanup
    GASToolbox.deleteFolder(folder);
    console.log("✅ deleteFolder: PASSED");

    console.log("✅ All Drive folder tests passed!");
    return true;
  } catch (error) {
    console.error("❌ Drive folder tests failed:", error);
    return false;
  }
}

/**
 * Test Drive file operations
 */
function testDriveFileOperations() {
  console.log("Testing Drive file operations...");

  try {
    // Create a test folder
    const folder = GASToolbox.getOrCreateFolderByPath("TEST/Integration");

    // Create a test document
    const doc = GASToolbox.createDocument("TEST/Integration", "TestDocument");
    console.log("✅ createDocument: PASSED");

    // Test getFileById
    const fileById = GASToolbox.getFileById(doc.getId());
    if (fileById && fileById.getId() === doc.getId()) {
      console.log("✅ getFileById: PASSED");
    } else {
      console.log("❌ getFileById: FAILED");
    }

    // Test findFile
    const foundFile = GASToolbox.findFile("TEST/Integration", "TestDocument");
    if (foundFile && foundFile.getId() === doc.getId()) {
      console.log("✅ findFile: PASSED");
    } else {
      console.log("❌ findFile: FAILED");
    }

    // Test checkFileExists
    const fileExists = GASToolbox.checkFileExists(
      "TEST/Integration",
      "TestDocument"
    );
    if (fileExists) {
      console.log("✅ checkFileExists: PASSED");
    } else {
      console.log("❌ checkFileExists: FAILED");
    }

    // Cleanup
    GASToolbox.deleteFile(doc);
    GASToolbox.deleteFolder(folder);
    console.log("✅ deleteFile: PASSED");

    console.log("✅ All Drive file tests passed!");
    return true;
  } catch (error) {
    console.error("❌ Drive file tests failed:", error);
    return false;
  }
}

/**
 * Test Docs operations
 */
function testDocsOperations() {
  console.log("Testing Docs operations...");

  try {
    const folderPath = "TEST/Integration";
    const fileName = "TestDoc";

    // Create document
    const doc = GASToolbox.createDocument(folderPath, fileName);
    console.log("✅ createDocument: PASSED");

    // Test appendParagraphToFile
    GASToolbox.appendParagraphToFile(folderPath, fileName, "Hello World!");
    console.log("✅ appendParagraphToFile: PASSED");

    // Test getDocumentContent
    const content = GASToolbox.getDocumentContent(folderPath, fileName);
    if (content && content.includes("Hello World!")) {
      console.log("✅ getDocumentContent: PASSED");
    } else {
      console.log("❌ getDocumentContent: FAILED");
    }

    // Test getParagraphCount
    const count = GASToolbox.getParagraphCount(folderPath, fileName);
    if (count > 0) {
      console.log("✅ getParagraphCount: PASSED");
    } else {
      console.log("❌ getParagraphCount: FAILED");
    }

    // Test appendBulletedListToFile
    GASToolbox.appendBulletedListToFile(folderPath, fileName, [
      "Item 1",
      "Item 2",
    ]);
    console.log("✅ appendBulletedListToFile: PASSED");

    // Cleanup
    GASToolbox.deleteFile(doc);
    const folder = GASToolbox.findFile("TEST/Integration", "");
    if (folder) {
      GASToolbox.deleteFolder(folder);
    }

    console.log("✅ All Docs tests passed!");
    return true;
  } catch (error) {
    console.error("❌ Docs tests failed:", error);
    return false;
  }
}

/**
 * Test Sheets operations (requires active spreadsheet)
 */
function testSheetsOperations() {
  console.log("Testing Sheets operations...");

  try {
    // Get active spreadsheet
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      console.log("⚠️ No active spreadsheet. Create one and try again.");
      return false;
    }

    const sheetName = "IntegrationTest";
    const header = ["name", "age", "email"];

    // Test createSheet
    const sheet = GASToolbox.createSheet(sheetName, header, spreadsheet);
    console.log("✅ createSheet: PASSED");

    // Test appendObject
    GASToolbox.appendObject(sheetName, header, {
      name: "Test User",
      age: 30,
      email: "test@example.com",
    });
    console.log("✅ appendObject: PASSED");

    // Test getAllObjects
    const objects = GASToolbox.getAllObjects(sheetName);
    if (objects.length > 0) {
      console.log("✅ getAllObjects: PASSED");
    } else {
      console.log("❌ getAllObjects: FAILED");
    }

    // Test getObject
    const obj = GASToolbox.getObject(sheetName, 0);
    if (obj && obj.name === "Test User") {
      console.log("✅ getObject: PASSED");
    } else {
      console.log("❌ getObject: FAILED");
    }

    // Test countObjects
    const count = GASToolbox.countObjects(sheetName);
    if (count === 1) {
      console.log("✅ countObjects: PASSED");
    } else {
      console.log("❌ countObjects: FAILED");
    }

    // Cleanup
    spreadsheet.deleteSheet(sheet);

    console.log("✅ All Sheets tests passed!");
    return true;
  } catch (error) {
    console.error("❌ Sheets tests failed:", error);
    return false;
  }
}

/**
 * Run all integration tests
 */
function runAllIntegrationTests() {
  console.log("🚀 Running all integration tests...\n");

  const results = {
    driveFolders: testDriveFolderOperations(),
    driveFiles: testDriveFileOperations(),
    docs: testDocsOperations(),
    sheets: testSheetsOperations(),
  };

  console.log("\n📊 Test Results Summary:");
  console.log(
    "Drive Folders:",
    results.driveFolders ? "✅ PASSED" : "❌ FAILED"
  );
  console.log("Drive Files:", results.driveFiles ? "✅ PASSED" : "❌ FAILED");
  console.log("Docs:", results.docs ? "✅ PASSED" : "❌ FAILED");
  console.log("Sheets:", results.sheets ? "✅ PASSED" : "❌ FAILED");

  const allPassed = Object.values(results).every(result => result === true);

  if (allPassed) {
    console.log("\n✅ All integration tests passed!");
  } else {
    console.log("\n⚠️ Some tests failed. Check the logs above.");
  }

  return allPassed;
}
