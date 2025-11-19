#!/usr/bin/env node

/**
 * Updates all badges in README.md:
 * - Coverage (from coverage-summary.json)
 * - Version (from package.json)
 * - TypeScript version (from package.json)
 * - Node.js version (from package.json engines or default)
 * - License (from package.json)
 */

const fs = require('fs');
const path = require('path');

const coverageSummaryPath = path.join(__dirname, '..', 'coverage', 'coverage-summary.json');
const packageJsonPath = path.join(__dirname, '..', 'package.json');
const readmePath = path.join(__dirname, '..', 'README.md');

try {
  // Read package.json
  const packageJson = JSON.parse(fs.readFileSync(packageJsonPath, 'utf8'));
  const version = packageJson.version || '1.0.0';
  const license = packageJson.license || 'MIT';
  
  // Get TypeScript version from devDependencies
  const typescriptVersion = packageJson.devDependencies?.typescript?.replace(/[\^~]/, '') || '5.0';
  
  // Get Node.js version from package.json (check engines field or default to 18+)
  let nodeVersion = '18+';
  if (packageJson.engines?.node) {
    nodeVersion = packageJson.engines.node.replace(/[\^~>=]/, '');
  } else if (packageJson.engines?.['node-version']) {
    nodeVersion = packageJson.engines['node-version'].replace(/[\^~>=]/, '');
  }

  // Read coverage summary
  let coverage = '100.00';
  let coverageColor = 'green';
  
  try {
    const coverageSummary = JSON.parse(fs.readFileSync(coverageSummaryPath, 'utf8'));
    coverage = coverageSummary.total.lines.pct.toFixed(2);
    
    // Determine badge color based on coverage
    if (coverage >= 80) {
      coverageColor = 'green';
    } else if (coverage >= 60) {
      coverageColor = 'yellow';
    } else if (coverage >= 40) {
      coverageColor = 'orange';
    } else {
      coverageColor = 'red';
    }
  } catch (error) {
    console.log('⚠️  Coverage file not found, using default coverage');
  }

  // Generate badge markdown
  const badges = [
    `[![Coverage](https://img.shields.io/badge/coverage-${coverage}%25-${coverageColor}.svg)](https://github.com/gardusig/gas-toolbox)`,
    `[![License](https://img.shields.io/badge/license-${license}-blue.svg)](LICENSE)`,
    `[![Version](https://img.shields.io/badge/version-${version}-blue.svg)](package.json)`,
    `[![TypeScript](https://img.shields.io/badge/TypeScript-${typescriptVersion}-blue.svg)](https://www.typescriptlang.org/)`,
    `[![Node.js](https://img.shields.io/badge/Node.js-${nodeVersion}-green.svg)](https://nodejs.org/)`,
  ];
  
  const badgeLine = badges.join('\n');

  // Read README
  let readme = fs.readFileSync(readmePath, 'utf8');
  
  // Replace badge section (everything between title and description)
  const badgeRegex = /# GAS Toolbox\n\n([!\[][^\n]+\n)*\nA comprehensive Google Apps Script/;
  
  if (badgeRegex.test(readme)) {
    readme = readme.replace(badgeRegex, `# GAS Toolbox\n\n${badgeLine}\n\nA comprehensive Google Apps Script`);
  } else {
    // Insert after title if pattern doesn't match
    readme = readme.replace(/(# GAS Toolbox\n)/, `$1\n${badgeLine}\n`);
  }
  
  fs.writeFileSync(readmePath, readme, 'utf8');
  
  console.log('✓ Updated badges:');
  console.log(`  - Coverage: ${coverage}%`);
  console.log(`  - Version: ${version}`);
  console.log(`  - License: ${license}`);
  console.log(`  - TypeScript: ${typescriptVersion}`);
  console.log(`  - Node.js: ${nodeVersion}`);
  
} catch (error) {
  console.error('Error updating badges:', error.message);
  process.exit(1);
}

