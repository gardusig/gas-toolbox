# Integration Tests

Integration tests for the deployed Google Apps Script library.

These tests verify that the code works correctly after being deployed to Apps Script using clasp.

## Setup

1. Ensure you have deployed the library to Apps Script:

   ```bash
   npm run deploy
   ```

2. Get your Script ID from `.clasp.json`:

   ```bash
   cat .clasp.json
   ```

3. In your Apps Script project, go to **Project Settings** and enable:
   - Show "appsscript.json" manifest file in editor

## Running Tests

### Manual Testing

1. Open your Apps Script project:

   ```bash
   clasp open
   ```

2. In the Apps Script editor, go to **Extensions** → **Apps Script** → **Test functions**

3. Run test functions manually or create test triggers

### Automated Testing with clasp

```bash
# Run a specific function
clasp run functionName

# Example: Run a test function
clasp run testDriveFunctions
```

## Test Structure

```
integration-tests/
├── README.md              # This file
├── test-functions.js      # Test functions to be run in Apps Script
├── test-helpers.js        # Helper utilities for tests
└── manual-tests.md        # Manual testing instructions
```

## Test Files

### `test-functions.js`

This file contains test functions that can be executed in the Apps Script editor or via `clasp run`. These functions test the actual deployed code.

### `test-helpers.js`

Helper functions and utilities to support integration testing.

### `manual-tests.md`

Step-by-step instructions for manually verifying the library works correctly.

## Notes

- Integration tests run against the actual deployed code in Google Apps Script
- These tests require appropriate permissions and Google account access
- Some tests may create or modify files/folders in your Google Drive
- Always test in a development/test Google account before production use
