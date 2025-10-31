#!/usr/bin/env node
/**
 * Validate that function names referenced as strings in setFunctionName() calls
 * actually exist as functions in the codebase
 */

const fs = require('fs');
const path = require('path');

const addonDir = path.join(__dirname, '..', 'addon');
const definedFunctions = new Set();
const referencedFunctions = [];

// Extract all function names from .gs files
const files = fs.readdirSync(addonDir)
  .filter(file => file.endsWith('.gs'))
  .map(file => path.join(addonDir, file));

files.forEach(file => {
  const content = fs.readFileSync(file, 'utf8');
  
  // Extract all function declarations
  const funcPattern = /^\s*function\s+([a-zA-Z_$][a-zA-Z0-9_$]*)\s*\(/gm;
  let match;
  while ((match = funcPattern.exec(content)) !== null) {
    definedFunctions.add(match[1]);
  }
  
  // Extract function names from setFunctionName() calls
  // Pattern: setFunctionName('functionName') or setFunctionName("functionName")
  const refPattern = /\.setFunctionName\(['"]([a-zA-Z_$][a-zA-Z0-9_$]*)['"]\)/g;
  while ((match = refPattern.exec(content)) !== null) {
    referencedFunctions.push({
      file: path.basename(file),
      line: content.substring(0, match.index).split('\n').length,
      functionName: match[1]
    });
  }
});

// Check which referenced functions don't exist
const errors = [];
referencedFunctions.forEach(ref => {
  if (!definedFunctions.has(ref.functionName)) {
    errors.push({
      file: ref.file,
      line: ref.line,
      functionName: ref.functionName
    });
  }
});

// Report results
if (errors.length > 0) {
  console.error('❌ Found references to undefined functions:\n');
  errors.forEach(error => {
    console.error(`  ${error.file}:${error.line} - setFunctionName('${error.functionName}')`);
  });
  console.error(`\nTotal: ${errors.length} error(s)`);
  process.exit(1);
} else {
  console.log(`✅ All function references are valid`);
  console.log(`   Checked ${referencedFunctions.length} setFunctionName() calls`);
  console.log(`   All reference functions that exist (${definedFunctions.size} total)`);
  process.exit(0);
}

