# Setup Instructions

## Quick Start

1. **Install dependencies:**
   ```bash
   npm install
   ```

2. **Set up clasp (Google Apps Script CLI):**
   
   Enable the Apps Script API:
   - Visit: https://script.google.com/home/usersettings
   - Enable "Google Apps Script API"
   
   Login to clasp:
   ```bash
   npx clasp login
   ```

3. **Link your Apps Script project:**
   
   If you have an existing project:
   ```bash
   cp .clasp.json.template .clasp.json
   # Edit .clasp.json and add your Script ID
   ```
   
   If creating a new project:
   ```bash
   npx clasp create --title "Gmail Fast Reader" --type standalone --rootDir .
   ```

4. **You're ready!**

## Available Commands

- `npm run lint:update-globals` - Extract global function names for the linter
- `npm run lint:validate-functions` - Validate that function references in `setFunctionName()` calls exist
- `npm run lint` - Check code for linting issues (includes function reference validation)
- `npm run lint:fix` - Auto-fix linting issues
- `npm test` - Run tests
- `npm run test:watch` - Run tests in watch mode
- `npm run test:coverage` - Run tests with coverage report
- `npm run push` - Push code to Apps Script
- `npm run open` - Open Apps Script editor in browser
- `npm run deploy` - Push and deploy new version
- `npm run logs` - View execution logs

## Project Structure

- `addon/*.gs` - Google Apps Script source files
- `addon/appsscript.json` - Apps Script manifest
- `tests/` - Jest test files
- `.claspignore` - Files excluded when pushing to Apps Script
