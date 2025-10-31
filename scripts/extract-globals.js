#!/usr/bin/env node
/**
 * Extract function and constant definitions from .gs files
 * to generate ESLint globals configuration
 */

const fs = require('fs');
const path = require('path');

const addonDir = path.join(__dirname, '..', 'addon');
const globals = new Set();

// Google Apps Script built-in services (always available)
const builtInServices = [
  'CardService',
  'PropertiesService',
  'GmailApp',
  'ScriptApp',
  'Session',
  'UrlFetchApp',
  'Logger',
  'Utilities',
  'console'
];

builtInServices.forEach(name => globals.add(name));

// Read all .gs files from addon directory
const files = fs.readdirSync(addonDir)
  .filter(file => file.endsWith('.gs'))
  .map(file => path.join(addonDir, file));

files.forEach(file => {
  const content = fs.readFileSync(file, 'utf8');
  
  // Extract all function declarations (Google Apps Script makes all functions global)
  const funcPattern = /^\s*function\s+([a-zA-Z_$][a-zA-Z0-9_$]*)\s*\(/gm;
  let match;
  while ((match = funcPattern.exec(content)) !== null) {
    globals.add(match[1]);
  }
  
  // Extract top-level const declarations (uppercase constants)
  const constPattern = /^const\s+([A-Z_][A-Z0-9_]*)\s*=/gm;
  while ((match = constPattern.exec(content)) !== null) {
    globals.add(match[1]);
  }
});

// Convert to sorted array and format as ESLint globals
const globalNames = Array.from(globals).sort();

const globalsObject = globalNames.reduce((acc, name) => {
  acc[name] = 'readonly';
  return acc;
}, {});

console.log(JSON.stringify(globalsObject, null, 2));
