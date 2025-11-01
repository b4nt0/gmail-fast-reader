/**
 * Gmail Fast Reader - Google Drive storage for accumulated results
 */

const ACCUMULATION_FILE_NAME = 'gmail-fast-read-accumulated-results.json';

/**
 * Get or create the accumulation file on Google Drive
 * @returns {GoogleAppsScript.Drive.File} The file object
 */
function getOrCreateAccumulationFile() {
  try {
    const rootFolder = DriveApp.getRootFolder();
    const files = rootFolder.getFilesByName(ACCUMULATION_FILE_NAME);
    
    if (files.hasNext()) {
      return files.next();
    }
    
    // File doesn't exist, create it with empty structure
    const initialData = {
      mustDo: [],
      mustKnow: [],
      totalProcessed: 0,
      firstDate: null,
      lastDate: null
    };
    
    const file = rootFolder.createFile(
      ACCUMULATION_FILE_NAME,
      JSON.stringify(initialData, null, 2),
      'application/json'
    );
    
    console.log('Created accumulation file on Drive:', file.getId());
    return file;
  } catch (error) {
    console.error('Error getting or creating accumulation file:', error);
    throw new Error('Failed to access Google Drive: ' + error.message);
  }
}

/**
 * Load accumulated results from Google Drive
 * @returns {Object} Accumulated results with structure { mustDo: [], mustKnow: [], totalProcessed: 0, firstDate: ISO, lastDate: ISO }
 */
function loadAccumulatedResults() {
  try {
    const file = getOrCreateAccumulationFile();
    const content = file.getBlob().getDataAsString();
    const data = JSON.parse(content);
    
    // Ensure structure is valid
    return {
      mustDo: data.mustDo || [],
      mustKnow: data.mustKnow || [],
      totalProcessed: data.totalProcessed || 0,
      firstDate: data.firstDate || null,
      lastDate: data.lastDate || null
    };
  } catch (error) {
    console.error('Error loading accumulated results:', error);
    // Return empty structure on error
    return {
      mustDo: [],
      mustKnow: [],
      totalProcessed: 0,
      firstDate: null,
      lastDate: null
    };
  }
}

/**
 * Save accumulated results to Google Drive
 * @param {Object} results - Results object with mustDo, mustKnow, totalProcessed, firstDate, lastDate
 */
function saveAccumulatedResults(results) {
  try {
    const file = getOrCreateAccumulationFile();
    
    // Ensure structure is valid
    const dataToSave = {
      mustDo: results.mustDo || [],
      mustKnow: results.mustKnow || [],
      totalProcessed: results.totalProcessed || 0,
      firstDate: results.firstDate || null,
      lastDate: results.lastDate || null
    };
    
    file.setContent(JSON.stringify(dataToSave, null, 2));
    console.log('Saved accumulated results to Drive:', {
      mustDo: dataToSave.mustDo.length,
      mustKnow: dataToSave.mustKnow.length,
      totalProcessed: dataToSave.totalProcessed
    });
  } catch (error) {
    console.error('Error saving accumulated results:', error);
    throw new Error('Failed to save accumulated results: ' + error.message);
  }
}

/**
 * Clear accumulated results from Google Drive
 */
function clearAccumulatedResults() {
  try {
    const rootFolder = DriveApp.getRootFolder();
    const files = rootFolder.getFilesByName(ACCUMULATION_FILE_NAME);
    
    while (files.hasNext()) {
      const file = files.next();
      file.setTrashed(true);
    }
    
    console.log('Cleared accumulated results from Drive');
  } catch (error) {
    console.error('Error clearing accumulated results:', error);
    throw new Error('Failed to clear accumulated results: ' + error.message);
  }
}

