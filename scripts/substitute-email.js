#!/usr/bin/env node
/**
 * Substitute email address in Constants.js
 * Reads real email from .author-email.txt and replaces dummy email in Constants.js
 * 
 * Usage:
 *   node scripts/substitute-email.js <direction>
 *   direction: 'real' to substitute with real email, 'dummy' to substitute back to dummy
 */

const fs = require('fs');
const path = require('path');

const projectRoot = path.join(__dirname, '..');
const constantsPath = path.join(projectRoot, 'addon', 'Constants.js');
const emailFilePath = path.join(projectRoot, '.author-email.txt');

const DUMMY_EMAIL = 'your-email@example.com';
const EMAIL_PATTERN = /const DEBUG_USER_EMAIL = '([^']+)';/;

function substituteEmail(direction) {
  if (direction !== 'real' && direction !== 'dummy') {
    console.error('Error: Direction must be "real" or "dummy"');
    console.error('Usage: node scripts/substitute-email.js <real|dummy>');
    process.exit(1);
  }

  // Read Constants.js
  if (!fs.existsSync(constantsPath)) {
    console.error(`Error: Constants.js not found at ${constantsPath}`);
    process.exit(1);
  }

  let constantsContent = fs.readFileSync(constantsPath, 'utf8');

  if (direction === 'real') {
    // Substitute dummy email with real email
    if (!fs.existsSync(emailFilePath)) {
      console.error(`Error: .author-email.txt not found at ${emailFilePath}`);
      console.error('Please create this file with your email address.');
      process.exit(1);
    }

    const realEmail = fs.readFileSync(emailFilePath, 'utf8').trim();
    if (!realEmail) {
      console.error('Error: .author-email.txt is empty');
      process.exit(1);
    }

    // Check if already substituted
    const match = constantsContent.match(EMAIL_PATTERN);
    if (match && match[1] === realEmail) {
      console.log('Email already substituted with real address. Skipping.');
      return;
    }

    // Replace with real email
    constantsContent = constantsContent.replace(
      EMAIL_PATTERN,
      `const DEBUG_USER_EMAIL = '${realEmail}';`
    );

    fs.writeFileSync(constantsPath, constantsContent, 'utf8');
    console.log(`✓ Substituted email with real address: ${realEmail}`);
  } else {
    // Substitute back to dummy email
    const match = constantsContent.match(EMAIL_PATTERN);
    if (match && match[1] === DUMMY_EMAIL) {
      console.log('Email already substituted with dummy address. Skipping.');
      return;
    }

    // Replace with dummy email
    constantsContent = constantsContent.replace(
      EMAIL_PATTERN,
      `const DEBUG_USER_EMAIL = '${DUMMY_EMAIL}';`
    );

    fs.writeFileSync(constantsPath, constantsContent, 'utf8');
    console.log(`✓ Substituted email back to dummy address`);
  }
}

// Get direction from command line arguments
const direction = process.argv[2];
substituteEmail(direction);

