#!/usr/bin/env node

/**
 * Updates the coverage badge in README.md with the current coverage percentage
 */

const fs = require('fs');
const path = require('path');

const coverageSummaryPath = path.join(__dirname, '..', 'coverage', 'coverage-summary.json');
const readmePath = path.join(__dirname, '..', 'README.md');

try {
  const coverageSummary = JSON.parse(fs.readFileSync(coverageSummaryPath, 'utf8'));
  const coverage = coverageSummary.total.lines.pct.toFixed(2);
  
  // Determine badge color based on coverage
  let color = 'red';
  if (coverage >= 80) {
    color = 'green';
  } else if (coverage >= 60) {
    color = 'yellow';
  } else if (coverage >= 40) {
    color = 'orange';
  }
  
  const badgeUrl = `https://img.shields.io/badge/coverage-${coverage}%25-${color}.svg`;
  const badgeMarkdown = `[![Coverage](${badgeUrl})](https://github.com/gardusig/gas-toolbox)`;
  
  // Read README
  let readme = fs.readFileSync(readmePath, 'utf8');
  
  // Replace coverage badge
  const badgeRegex = /\[!\[Coverage\]\([^)]+\)\]\([^)]+\)/;
  if (badgeRegex.test(readme)) {
    readme = readme.replace(badgeRegex, badgeMarkdown);
  } else {
    // Insert after title if not found
    readme = readme.replace(/(# GAS Toolbox\n)/, `$1\n${badgeMarkdown}\n`);
  }
  
  fs.writeFileSync(readmePath, readme, 'utf8');
  console.log(`✓ Updated coverage badge to ${coverage}%`);
} catch (error) {
  console.error('Error updating coverage badge:', error.message);
  process.exit(1);
}

