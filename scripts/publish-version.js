#!/usr/bin/env node

/**
 * Publishes a new version of the library:
 * 1. Gets the latest version number
 * 2. Increments it
 * 3. Builds the project
 * 4. Deploys to Apps Script
 * 5. Creates a new version
 */

const { execSync } = require('child_process');
const path = require('path');

function getLatestVersion() {
  try {
    const output = execSync('clasp versions', { encoding: 'utf-8' });
    const lines = output.trim().split('\n').filter(line => line.trim());
    
    if (lines.length === 0 || lines[0].includes('No deployed versions')) {
      return 0;
    }
    
    // Parse version number from lines like "1 - Initial version" or "3 - 3"
    const versionLines = lines.filter(line => /^\d+\s+-/.test(line));
    if (versionLines.length === 0) {
      return 0;
    }
    
    // Get the last line (highest version number)
    const lastLine = versionLines[versionLines.length - 1];
    const match = lastLine.match(/^(\d+)/);
    return match ? parseInt(match[1], 10) : 0;
  } catch (error) {
    console.error('Error getting versions:', error.message);
    return 0;
  }
}

function build() {
  console.log('📦 Building project...');
  execSync('npm run build', { stdio: 'inherit' });
}

function deploy() {
  console.log('🚀 Deploying to Apps Script...');
  execSync('clasp push', { stdio: 'inherit' });
}

function createVersion(versionNumber, description) {
  console.log(`📌 Creating version ${versionNumber}...`);
  const versionDescription = description || `Version ${versionNumber}`;
  execSync(`clasp version "${versionDescription}"`, { stdio: 'inherit' });
}

// Main execution
const args = process.argv.slice(2);
const customDescription = args.join(' ');

const latestVersion = getLatestVersion();
const nextVersion = latestVersion + 1;

console.log(`\n📋 Current latest version: ${latestVersion}`);
console.log(`✨ Next version will be: ${nextVersion}\n`);

build();
deploy();
createVersion(nextVersion, customDescription || `Version ${nextVersion}`);

console.log(`\n✅ Successfully published version ${nextVersion}!`);

