#!/usr/bin/env node

/**
 * Post-build script to remove ES module syntax from compiled JavaScript
 * for Google Apps Script compatibility.
 * 
 * Apps Script doesn't support ES modules, so we need to:
 * 1. Remove 'export' keywords
 * 2. Remove 'import' statements (since everything is in global scope)
 */

const fs = require('fs');
const path = require('path');

const distDir = path.join(__dirname, '..', 'dist');

// Recursively get all .js files in dist (excluding .map files)
function getAllJsFiles(dir, fileList = []) {
  const files = fs.readdirSync(dir);
  files.forEach(file => {
    const filePath = path.join(dir, file);
    if (fs.statSync(filePath).isDirectory()) {
      getAllJsFiles(filePath, fileList);
    } else if (file.endsWith('.js') && !file.endsWith('.map')) {
      fileList.push(filePath);
    }
  });
  return fileList;
}

const jsFiles = getAllJsFiles(distDir);

jsFiles.forEach(filePath => {
  const relativePath = path.relative(distDir, filePath);
  let content = fs.readFileSync(filePath, 'utf8');
  
  // Remove export keywords (but keep the declarations)
  // Match: export var/const/function/namespace/class
  content = content.replace(/export\s+/g, '');
  
  // Remove export statements like: export { ... };
  content = content.replace(/export\s*\{[^}]*\}\s*;?\s*/g, '');
  
  // Remove import/export statements with from clauses
  // Match: { ... } from "..."
  content = content.replace(/^\s*\{[^}]*\}\s+from\s+['"][^'"]+['"];?\s*$/gm, '');
  
  // Remove standalone object literal statements left after removing export
  // Match: { Identifier }; on its own line
  content = content.replace(/^\s*\{[^}]*\}\s*;\s*$/gm, '');
  
  // Remove import statements
  // Match: import ... from "..."
  content = content.replace(/import\s+.*?\s+from\s+['"][^'"]+['"];?\s*/g, '');
  content = content.replace(/import\s+['"][^'"]+['"];?\s*/g, '');
  
  // Remove source map references (Apps Script doesn't need them)
  content = content.replace(/\/\/# sourceMappingURL=.*$/gm, '');
  
  // Fix namespace initialization for Apps Script library compatibility
  // Change "var Namespace;" to "var Namespace = Namespace || {};"
  // This ensures the namespace is properly initialized when used as a library
  content = content.replace(/^var (Mapper|Util);$/m, 'var $1 = $1 || {};');
  
  fs.writeFileSync(filePath, content, 'utf8');
  console.log(`Processed: ${relativePath}`);
});

console.log('✓ Removed ES module syntax from all files');

