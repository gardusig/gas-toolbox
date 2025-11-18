#!/usr/bin/env node

/**
 * Deploy integration test functions to Apps Script
 * This allows the test functions to be run in the Apps Script environment
 */

const { execSync } = require("child_process");
const fs = require("fs");
const path = require("path");

const INTEGRATION_TESTS_DIR = path.join(__dirname, "..", "integration-tests");
const TEST_FUNCTIONS_FILE = path.join(INTEGRATION_TESTS_DIR, "test-functions.js");

function log(message, type = "info") {
  const icons = { info: "ℹ️", success: "✅", error: "❌", warning: "⚠️" };
  const icon = icons[type] || icons.info;
  console.log(`${icon} ${message}`);
}

function error(message) {
  log(message, "error");
  process.exit(1);
}

function checkFileExists(filePath) {
  if (!fs.existsSync(filePath)) {
    error(`Test file not found: ${filePath}`);
  }
}

function deployTestFunctions() {
  log("Deploying integration test functions to Apps Script...");

  // Check if test file exists
  checkFileExists(TEST_FUNCTIONS_FILE);

  // Create a temporary directory for test files
  const tempDir = path.join(__dirname, "..", ".temp-tests");
  if (!fs.existsSync(tempDir)) {
    fs.mkdirSync(tempDir, { recursive: true });
  }

  // Copy test functions to temp directory
  const tempTestFile = path.join(tempDir, "test-functions.js");
  fs.copyFileSync(TEST_FUNCTIONS_FILE, tempTestFile);

  log("Test functions copied to temporary directory");

  // Note: This is a simplified version. In practice, you might want to:
  // 1. Use clasp push to push the test file directly
  // 2. Or append it to an existing file in the Apps Script project
  // 3. Or use clasp push with a custom .claspignore that includes test files

  log(
    "\n⚠️ Manual step required:",
    "warning"
  );
  log(
    "Copy the contents of integration-tests/test-functions.js into your Apps Script editor.",
    "warning"
  );
  log(
    "Or run: clasp open and manually add the test functions.",
    "warning"
  );

  // Cleanup
  fs.rmSync(tempDir, { recursive: true, force: true });

  log("\n📝 Test functions are ready to be added to your Apps Script project", "success");
}

if (require.main === module) {
  deployTestFunctions();
}

module.exports = { deployTestFunctions };

