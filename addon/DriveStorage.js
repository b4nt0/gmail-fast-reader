/**
 * Gmail Fast Reader - Google Drive storage for accumulated results
 */

const ACCUMULATION_FILE_NAME = 'gmail-fast-read-accumulated-results.json';
const FILE_ID_PROPERTY_KEY = 'drive_accumulation_file_id';

/**
 * Get or create the accumulation file on Google Drive
 * Uses drive.file scope - stores file ID in PropertiesService to avoid needing getRootFolder()
 * @returns {GoogleAppsScript.Drive.File} The file object
 */
function getOrCreateAccumulationFile() {
  try {
    const properties = PropertiesService.getUserProperties();
    let fileId = properties.getProperty(FILE_ID_PROPERTY_KEY);
    
    // Try to get file by stored ID first
    if (fileId) {
      try {
        const file = DriveApp.getFileById(fileId);
        // Verify file still exists and is accessible
        file.getName(); // This will throw if file doesn't exist or isn't accessible
        return file;
      } catch (e) {
        // File ID is invalid or file was deleted, clear it and search/create new
        properties.deleteProperty(FILE_ID_PROPERTY_KEY);
        fileId = null;
      }
    }
    
    // File ID not found or invalid, search for file by name
    // This only finds files created by this app (within drive.file scope)
    const files = DriveApp.getFilesByName(ACCUMULATION_FILE_NAME);
    
    if (files.hasNext()) {
      const file = files.next();
      // Store the ID for future reference
      properties.setProperty(FILE_ID_PROPERTY_KEY, file.getId());
      return file;
    }
    
    // File doesn't exist, create it with empty structure
    // DriveApp.createFile() works with drive.file scope and creates in root
    const initialData = {
      mustDo: [],
      mustKnow: [],
      totalProcessed: 0,
      firstDate: null,
      lastDate: null
    };
    
    const file = DriveApp.createFile(
      ACCUMULATION_FILE_NAME,
      JSON.stringify(initialData, null, 2),
      'application/json'
    );
    
    // Store the file ID for future reference
    properties.setProperty(FILE_ID_PROPERTY_KEY, file.getId());
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
    const properties = PropertiesService.getUserProperties();
    const fileId = properties.getProperty(FILE_ID_PROPERTY_KEY);
    
    if (fileId) {
      try {
        const file = DriveApp.getFileById(fileId);
        file.setTrashed(true);
        properties.deleteProperty(FILE_ID_PROPERTY_KEY);
      } catch (e) {
        // File doesn't exist or isn't accessible, just clear the property
        properties.deleteProperty(FILE_ID_PROPERTY_KEY);
      }
    }
    
    // Also try to find and delete by name (in case ID was lost)
    const files = DriveApp.getFilesByName(ACCUMULATION_FILE_NAME);
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

