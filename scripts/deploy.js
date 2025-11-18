#!/usr/bin/env node

/**
 * Deployment script for Google Apps Script with validation
 * - Validates build output
 * - Checks clasp configuration
 * - Deploys with proper error handling
 */

const { execSync } = require("child_process");
const fs = require("fs");
const path = require("path");

function log(message, type = "info") {
  const icons = { info: "ℹ️", success: "✅", error: "❌", warning: "⚠️" };
  const icon = icons[type] || icons.info;
  console.log(`${icon} ${message}`);
}

function error(message) {
  log(message, "error");
  process.exit(1);
}

function checkClaspInstalled() {
  try {
    execSync("clasp --version", { stdio: "pipe" });
    return true;
  } catch {
    return false;
  }
}

function checkClaspConfig() {
  const claspJsonPath = path.join(process.cwd(), ".clasp.json");
  if (!fs.existsSync(claspJsonPath)) {
    error(
      ".clasp.json not found. Run 'clasp create' to set up your Apps Script project."
    );
  }

  try {
    const claspConfig = JSON.parse(fs.readFileSync(claspJsonPath, "utf8"));
    if (!claspConfig.scriptId) {
      error(".clasp.json is missing scriptId.");
    }
    return claspConfig;
  } catch (err) {
    error(`Failed to read .clasp.json: ${err.message}`);
  }
}

function checkBuildOutput() {
  const distPath = path.join(process.cwd(), "dist");
  if (!fs.existsSync(distPath)) {
    error("dist/ directory not found. Run 'npm run build' first.");
  }

  const appsscriptJsonPath = path.join(distPath, "appsscript.json");
  if (!fs.existsSync(appsscriptJsonPath)) {
    error("dist/appsscript.json not found. Build may have failed.");
  }

  // Check for JS files
  const jsFiles = [];
  function findJsFiles(dir) {
    const files = fs.readdirSync(dir);
    files.forEach(file => {
      const filePath = path.join(dir, file);
      const stat = fs.statSync(filePath);
      if (stat.isDirectory()) {
        findJsFiles(filePath);
      } else if (file.endsWith(".js") && !file.endsWith(".map")) {
        jsFiles.push(filePath);
      }
    });
  }
  findJsFiles(distPath);

  if (jsFiles.length === 0) {
    error("No JavaScript files found in dist/. Build may have failed.");
  }

  log(`Found ${jsFiles.length} JavaScript files to deploy`, "success");
  return jsFiles;
}

function validateClaspStatus() {
  try {
    log("Checking clasp status...");
    const status = execSync("clasp status", { encoding: "utf8", stdio: "pipe" });
    
    // Check for unwanted files
    const coverageCount = (status.match(/coverage\//g) || []).length;
    const mapFiles = (status.match(/\.map/g) || []).length;
    const dtsFiles = (status.match(/\.d\.ts/g) || []).length;
    const srcFiles = (status.match(/^└─ src\//gm) || []).length;
    const testFiles = (status.match(/^└─ tests\//gm) || []).length;
    
    // Count dist JS files only (these should be pushed)
    const distJsFiles = (status.match(/^└─ dist\/.*\.js$/gm) || []).length;
    const appsscriptJson = status.includes("dist/appsscript.json") ? 1 : 0;
    const expectedFileCount = distJsFiles + appsscriptJson;
    
    // Count all files
    const totalFileCount = (status.match(/^└─/g) || []).length;
    
    if (coverageCount > 0) {
      log(
        `Warning: coverage/ files detected (${coverageCount}). Check .claspignore configuration.`,
        "warning"
      );
    }
    if (mapFiles > 0 || dtsFiles > 0) {
      log(
        `Warning: source maps (${mapFiles}) or .d.ts files (${dtsFiles}) detected. Check .claspignore configuration.`,
        "warning"
      );
    }
    if (srcFiles > 0 || testFiles > 0) {
      log(
        `Warning: source files detected (src: ${srcFiles}, tests: ${testFiles}). Check .claspignore configuration.`,
        "warning"
      );
    }
    
    if (totalFileCount > expectedFileCount + 1) {
      // +1 for .clasp.json which is always included
      log(
        `Warning: ${totalFileCount} files will be pushed, but only ~${expectedFileCount} expected. Check .claspignore configuration.`,
        "warning"
      );
    } else {
      log(`Will push ${distJsFiles} JS file(s) + appsscript.json`, "success");
    }
    return true;
  } catch (err) {
    log("Could not check clasp status, continuing anyway...", "warning");
    return false;
  }
}

function build() {
  log("Building project...");
  try {
    execSync("npm run build", { stdio: "inherit" });
    log("Build completed successfully", "success");
  } catch (err) {
    error("Build failed. Fix errors before deploying.");
  }
}

function deploy(force = false) {
  const flag = force ? "-f" : "";
  log(`Deploying to Apps Script${force ? " (force)" : ""}...`);
  try {
    execSync(`clasp push ${flag}`, { stdio: "inherit" });
    log("Deployment completed successfully", "success");
  } catch (err) {
    error(`Deployment failed: ${err.message}`);
  }
}

function main() {
  const args = process.argv.slice(2);
  const force = args.includes("--force") || args.includes("-f");
  const skipBuild = args.includes("--skip-build");
  const skipValidation = args.includes("--skip-validation");

  log("Starting deployment process...\n");

  // Pre-flight checks
  if (!checkClaspInstalled()) {
    error(
      "clasp is not installed. Run 'npm install -g @google/clasp' to install it."
    );
  }

  checkClaspConfig();

  if (!skipBuild) {
    build();
  }

  if (!skipValidation) {
    checkBuildOutput();
    validateClaspStatus();
  }

  // Deploy
  deploy(force);

  log("\n✅ Deployment complete!", "success");
}

if (require.main === module) {
  main();
}

module.exports = { deploy, build, checkClaspConfig, checkBuildOutput };

