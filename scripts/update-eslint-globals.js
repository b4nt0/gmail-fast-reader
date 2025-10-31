#!/usr/bin/env node
/**
 * Update .eslintrc.js with automatically extracted globals from .js files
 */

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const projectRoot = path.join(__dirname, '..');
const eslintrcPath = path.join(projectRoot, '.eslintrc.js');

// Run the extract-globals script and parse output
const globalsJson = execSync('node scripts/extract-globals.js', { 
  cwd: projectRoot,
  encoding: 'utf8' 
});

const globals = JSON.parse(globalsJson);
const globalNames = Object.keys(globals).sort();

// Read current ESLint config
const eslintConfigText = fs.readFileSync(eslintrcPath, 'utf8');

// Find the globals section and replace it
// Look for the pattern between "globals: {" and the closing brace
const globalsStartMarker = /(\s+globals:\s+\{)/;

if (!globalsStartMarker.test(eslintConfigText)) {
  console.error('Could not find globals section in .eslintrc.js');
  process.exit(1);
}

// Generate new globals list with proper indentation
const globalsList = globalNames
  .map(name => `        '${name}': 'readonly'`)
  .join(',\n');

// Replace the globals section
// Match the closing brace with its indentation, but capture the newline separately
const updatedConfig = eslintConfigText.replace(
  /(\s+globals:\s+\{)[\s\S]*?(\n\s+\},)/,
  `$1\n${globalsList}$2`
);

fs.writeFileSync(eslintrcPath, updatedConfig, 'utf8');

console.log(`✅ Updated .eslintrc.js with ${globalNames.length} globals from .js files`);
console.log(`   Found ${globalNames.length} functions and constants`);
console.log(`   Run 'npm run lint' to verify`);
