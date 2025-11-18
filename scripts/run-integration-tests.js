#!/usr/bin/env node

/**
 * Run integration tests in Apps Script environment
 * This script pushes test functions to Apps Script and runs them
 */

const { execSync } = require("child_process");
const fs = require("fs");
const path = require("path");

const TEST_FUNCTIONS_FILE = path.join(
  __dirname,
  "..",
  "integration-tests",
  "test-functions.js"
);
const TEMP_TEST_FILE = path.join(__dirname, "..", "test-functions.gs");

function log(message, type = "info") {
  const icons = { info: "ℹ️", success: "✅", error: "❌", warning: "⚠️" };
  const icon = icons[type] || icons.info;
  console.log(`${icon} ${message}`);
}

function error(message) {
  log(message, "error");
  process.exit(1);
}

function checkClaspConfig() {
  const claspJsonPath = path.join(process.cwd(), ".clasp.json");
  if (!fs.existsSync(claspJsonPath)) {
    error(
      ".clasp.json not found. Run 'clasp create' to set up your Apps Script project."
    );
  }
  return true;
}

function pushTestFunctions() {
  log("Pushing test functions to Apps Script...");

  // Check if test file exists
  if (!fs.existsSync(TEST_FUNCTIONS_FILE)) {
    error(`Test file not found: ${TEST_FUNCTIONS_FILE}`);
  }

  // Check if library is deployed first
  log("Checking if library is deployed...");
  try {
    execSync("clasp status", { stdio: "pipe" });
  } catch (err) {
    log(
      "Library might not be deployed. Run 'npm run deploy' first.",
      "warning"
    );
  }

  // Ensure we're in the project root
  process.chdir(path.join(__dirname, ".."));

  // Copy test functions to root for clasp to pick up
  // Use .gs extension for Apps Script compatibility
  const testFileName = "test-functions.gs";
  const tempTestFile = path.join(__dirname, "..", testFileName);
  fs.copyFileSync(TEST_FUNCTIONS_FILE, tempTestFile);
  log("Test functions copied to project root as test-functions.gs");

  // Temporarily allow test-functions.gs in .claspignore
  const claspignorePath = path.join(__dirname, "..", ".claspignore");
  let claspignoreContent = fs.readFileSync(claspignorePath, "utf8");
  const originalContent = claspignoreContent;

  // Remove test-functions.js from ignore list if commented
  claspignoreContent = claspignoreContent.replace(
    /# test-functions\.js/g,
    "test-functions.js"
  );

  // Add exception for test file if not present
  if (!claspignoreContent.includes(`!${testFileName}`)) {
    claspignoreContent += `\n# Allow test functions for integration tests\n!${testFileName}\n`;
  }

  fs.writeFileSync(claspignorePath, claspignoreContent);

  try {
    // Push with the test file
    execSync(`clasp push -f`, { stdio: "inherit" });
    log("Test functions pushed successfully", "success");
  } catch (err) {
    // Restore original .claspignore
    fs.writeFileSync(claspignorePath, originalContent);
    // Cleanup on error
    if (fs.existsSync(tempTestFile)) {
      fs.unlinkSync(tempTestFile);
    }
    error(`Failed to push test functions: ${err.message}`);
  }

  // Restore original .claspignore
  fs.writeFileSync(claspignorePath, originalContent);

  // Cleanup temp file
  if (fs.existsSync(tempTestFile)) {
    fs.unlinkSync(tempTestFile);
  }
}

function runTestFunction(functionName) {
  log(`Running test function: ${functionName}...`);
  try {
    const output = execSync(`clasp run ${functionName}`, {
      encoding: "utf8",
      stdio: "pipe",
    });
    console.log(output);
    return true;
  } catch (err) {
    console.error(err.stdout || err.message);
    return false;
  }
}

function runAllTests() {
  log("🚀 Running all integration tests...\n");

  const testFunctions = [
    "testDriveFolderOperations",
    "testDriveFileOperations",
    "testDocsOperations",
    "testSheetsOperations",
  ];

  const results = {};
  let allPassed = true;

  for (const func of testFunctions) {
    log(`\n📋 Testing: ${func}`, "info");
    const passed = runTestFunction(func);
    results[func] = passed;
    if (!passed) {
      allPassed = false;
    }
  }

  // Run the summary function if it exists
  try {
    log("\n📊 Getting test summary...", "info");
    runTestFunction("runAllIntegrationTests");
  } catch (err) {
    // Summary function might not be available, that's okay
  }

  console.log("\n📊 Test Results:");
  Object.entries(results).forEach(([func, passed]) => {
    console.log(
      `  ${func}: ${passed ? "✅ PASSED" : "❌ FAILED"}`
    );
  });

  if (allPassed) {
    log("\n✅ All integration tests passed!", "success");
  } else {
    log("\n⚠️ Some tests failed. Check the output above.", "warning");
    process.exit(1);
  }
}

function main() {
  const args = process.argv.slice(2);
  const skipPush = args.includes("--skip-push");
  const functionName = args.find(arg => !arg.startsWith("--"));

  checkClaspConfig();

  if (!skipPush) {
    pushTestFunctions();
  }

  if (functionName) {
    // Run a specific test function
    runTestFunction(functionName);
  } else {
    // Run all tests
    runAllTests();
  }
}

if (require.main === module) {
  main();
}

module.exports = { runTestFunction, runAllTests, pushTestFunctions };

