/**
 * Gmail Fast Reader - Drive API storage for accumulated results
 * 
 * Uses Drive API to store accumulated results. Drive API with alt='media' may return
 * a parsed JSON object directly, which we handle by stringifying it.
 */

const ACCUMULATION_FILE_NAME = 'gmail-fast-read-accumulated-results.json';
const FILE_ID_PROPERTY_KEY = 'accumulation_file_id';
const MAX_RETRIES = 3;
const RETRY_DELAY_MS = 1000; // 1 second base delay

/**
 * Retry a function with exponential backoff for transient errors
 * @param {Function} fn - Function to retry
 * @param {number} maxRetries - Maximum number of retries
 * @param {number} baseDelayMs - Base delay in milliseconds
 * @returns {*} The result of the function
 */
function retryWithBackoff(fn, maxRetries, baseDelayMs) {
  let lastError;
  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      return fn();
    } catch (error) {
      lastError = error;
      // Check if error is likely transient (server error, network issue)
      // Note: Permission errors are NOT transient and should not be retried
      const isTransient = error.message && (
        error.message.includes('сервера') || // Ukrainian: "server"
        error.message.includes('server') ||
        error.message.includes('Exception') ||
        error.toString().includes('Exception')
      ) && !error.message.includes('дозволів') && // Ukrainian: "permissions"
          !error.message.includes('permissions') &&
          !error.message.includes('insufficient');
      
      if (!isTransient || attempt === maxRetries) {
        throw error;
      }
      
      // Exponential backoff: 1s, 2s, 4s
      const delay = baseDelayMs * Math.pow(2, attempt);
      console.log(`Drive operation failed (attempt ${attempt + 1}/${maxRetries + 1}), retrying in ${delay}ms...`, error.message);
      Utilities.sleep(delay);
    }
  }
  throw lastError;
}

/**
 * Get or create the accumulation file ID
 * @returns {string|null} File ID or null if creation fails
 */
function getOrCreateAccumulationFileId() {
  const properties = PropertiesService.getUserProperties();
  let fileId = properties.getProperty(FILE_ID_PROPERTY_KEY);
  
  if (fileId) {
    // Verify file still exists
    try {
      Drive.Files.get(fileId);
      return fileId;
    } catch (e) {
      // File doesn't exist, create new one
      fileId = null;
      properties.deleteProperty(FILE_ID_PROPERTY_KEY);
    }
  }
  
  if (!fileId) {
    // Create new file
    try {
      const fileResource = {
        name: ACCUMULATION_FILE_NAME,
        mimeType: 'application/json'
      };
      const emptyData = JSON.stringify({
        mustDo: [],
        mustKnow: [],
        totalProcessed: 0,
        firstDate: null,
        lastDate: null
      });
      const blob = Utilities.newBlob(emptyData, 'application/json', ACCUMULATION_FILE_NAME);
      const createdFile = Drive.Files.create(fileResource, blob);
      fileId = createdFile.id;
      properties.setProperty(FILE_ID_PROPERTY_KEY, fileId);
      return fileId;
    } catch (e) {
      console.error('Failed to create accumulation file:', e);
      return null;
    }
  }
  
  return fileId;
}

/**
 * Load file content from Drive API
 * Handles the case where Drive API returns a parsed JSON object directly
 * @param {string} fileId - The file ID to load
 * @returns {string|null} File content as string, or null if loading fails
 */
function loadFileContentFromDrive(fileId) {
  try {
    const response = Drive.Files.get(fileId, { alt: 'media' });
    let loadedContent = null;
    
    // Check if response is already a parsed JSON object (this can happen with Drive API)
    if (response && typeof response === 'object' && !Array.isArray(response)) {
      const hasGetDataAsString = typeof response.getDataAsString === 'function';
      const hasGetBytes = typeof response.getBytes === 'function';
      
      if (!hasGetDataAsString && !hasGetBytes) {
        // It's a plain object without Blob methods - stringify it
        try {
          loadedContent = JSON.stringify(response, null, 2);
        } catch (e) {
          console.error('Failed to stringify parsed JSON object:', e);
        }
      }
    }
    
    // Try getDataAsString if we haven't loaded content yet
    if (!loadedContent && response && (typeof response.getDataAsString === 'function' || 'getDataAsString' in response)) {
      try {
        loadedContent = response.getDataAsString('UTF-8');
      } catch (e) {
        // Method exists but might not work, continue to next option
      }
    }
    
    // Fallback to getBytes() if getDataAsString didn't work
    if (!loadedContent && response && (typeof response.getBytes === 'function' || 'getBytes' in response)) {
      try {
        const bytes = response.getBytes();
        loadedContent = Utilities.newBlob(bytes).getDataAsString('UTF-8');
      } catch (e) {
        // Method exists but might not work, continue to next option
      }
    }
    
    // If it's already a string
    if (!loadedContent && typeof response === 'string') {
      loadedContent = response;
    }
    
    // Try object-based approaches
    if (!loadedContent && response && typeof response === 'object') {
      // Try to extract content from object
      if (response.content !== undefined) {
        if (typeof response.content === 'string') {
          loadedContent = response.content;
        } else if (response.content && typeof response.content.getDataAsString === 'function') {
          loadedContent = response.content.getDataAsString('UTF-8');
        } else if (response.content && typeof response.content.getBytes === 'function') {
          const bytes = response.content.getBytes();
          loadedContent = Utilities.newBlob(bytes).getDataAsString('UTF-8');
        }
      }
      
      // Try using UrlFetchApp as fallback
      if (!loadedContent) {
        try {
          const accessToken = ScriptApp.getOAuthToken();
          const url = 'https://www.googleapis.com/drive/v3/files/' + fileId + '?alt=media';
          const options = {
            method: 'get',
            headers: {
              'Authorization': 'Bearer ' + accessToken
            }
          };
          const fetchResponse = UrlFetchApp.fetch(url, options);
          if (fetchResponse.getResponseCode() === 200) {
            loadedContent = fetchResponse.getContentText('UTF-8');
          }
        } catch (e) {
          // UrlFetchApp approach failed
        }
      }
      
      // Final fallback: try to stringify if it's a plain object
      if (!loadedContent && !Array.isArray(response)) {
        const hasGetDataAsString = typeof response.getDataAsString === 'function';
        const hasGetBytes = typeof response.getBytes === 'function';
        if (!hasGetDataAsString && !hasGetBytes) {
          try {
            loadedContent = JSON.stringify(response, null, 2);
          } catch (e) {
            // Still failed
          }
        }
      }
    }
    
    return loadedContent;
  } catch (error) {
    console.error('Error loading file content from Drive:', error);
    return null;
  }
}

/**
 * Load accumulated results from Drive
 * @returns {Object} Accumulated results with structure { mustDo: [], mustKnow: [], totalProcessed: 0, firstDate: ISO, lastDate: ISO }
 */
function loadAccumulatedResults() {
  try {
    return retryWithBackoff(function() {
      const fileId = getOrCreateAccumulationFileId();
      if (!fileId) {
        // File creation failed, return empty structure
        return {
          mustDo: [],
          mustKnow: [],
          totalProcessed: 0,
          firstDate: null,
          lastDate: null
        };
      }
      
      const content = loadFileContentFromDrive(fileId);
      if (!content) {
        // No content found, return empty structure
        return {
          mustDo: [],
          mustKnow: [],
          totalProcessed: 0,
          firstDate: null,
          lastDate: null
        };
      }
      
      const data = JSON.parse(content);
      
      // Ensure structure is valid
      return {
        mustDo: data.mustDo || [],
        mustKnow: data.mustKnow || [],
        totalProcessed: data.totalProcessed || 0,
        firstDate: data.firstDate || null,
        lastDate: data.lastDate || null
      };
    }, MAX_RETRIES, RETRY_DELAY_MS);
  } catch (error) {
    console.error('Error loading accumulated results after retries:', error);
    // Return empty structure on error - better than failing completely
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
 * Save accumulated results to Drive
 * @param {Object} results - Results object with mustDo, mustKnow, totalProcessed, firstDate, lastDate
 */
function saveAccumulatedResults(results) {
  // Ensure structure is valid
  const dataToSave = {
    mustDo: results.mustDo || [],
    mustKnow: results.mustKnow || [],
    totalProcessed: results.totalProcessed || 0,
    firstDate: results.firstDate || null,
    lastDate: results.lastDate || null
  };
  
  const contentToSave = JSON.stringify(dataToSave);
  
  return retryWithBackoff(function() {
    const properties = PropertiesService.getUserProperties();
    let fileId = properties.getProperty(FILE_ID_PROPERTY_KEY);
    
    const blob = Utilities.newBlob(contentToSave, 'application/json', ACCUMULATION_FILE_NAME);
    
    if (fileId) {
      // Update existing file
      try {
        Drive.Files.update({}, fileId, blob);
        console.log('Saved accumulated results to Drive (updated):', {
          mustDo: dataToSave.mustDo.length,
          mustKnow: dataToSave.mustKnow.length,
          totalProcessed: dataToSave.totalProcessed,
          size: contentToSave.length
        });
      } catch (error) {
        // File might not exist, create new one
        fileId = null;
        properties.deleteProperty(FILE_ID_PROPERTY_KEY);
      }
    }
    
    if (!fileId) {
      // Create new file
      const fileResource = {
        name: ACCUMULATION_FILE_NAME,
        mimeType: 'application/json'
      };
      const createdFile = Drive.Files.create(fileResource, blob);
      fileId = createdFile.id;
      properties.setProperty(FILE_ID_PROPERTY_KEY, fileId);
      console.log('Saved accumulated results to Drive (created):', {
        mustDo: dataToSave.mustDo.length,
        mustKnow: dataToSave.mustKnow.length,
        totalProcessed: dataToSave.totalProcessed,
        size: contentToSave.length
      });
    }
  }, MAX_RETRIES, RETRY_DELAY_MS);
}

/**
 * Clear accumulated results from Drive
 */
function clearAccumulatedResults() {
  try {
    return retryWithBackoff(function() {
      const properties = PropertiesService.getUserProperties();
      const fileId = properties.getProperty(FILE_ID_PROPERTY_KEY);
      
      if (fileId) {
        try {
          Drive.Files.remove(fileId);
        } catch (e) {
          // File might not exist, that's okay
        }
        properties.deleteProperty(FILE_ID_PROPERTY_KEY);
      }
      
      console.log('Cleared accumulated results from Drive');
    }, MAX_RETRIES, RETRY_DELAY_MS);
  } catch (error) {
    console.error('Error clearing accumulated results after retries:', error);
    // Don't throw - clearing is best-effort
    const properties = PropertiesService.getUserProperties();
    properties.deleteProperty(FILE_ID_PROPERTY_KEY);
  }
}

/**
 * Debug function: Get storage info from Drive
 * Only available in debug mode
 */
function debugGetCurrentFileId() {
  if (!debugMode()) {
    throw new Error('This function is only available in debug mode');
  }
  
  const properties = PropertiesService.getUserProperties();
  const fileId = properties.getProperty(FILE_ID_PROPERTY_KEY);
  
  if (!fileId) {
    return {
      storageType: 'empty',
      fileId: 'N/A',
      fileName: ACCUMULATION_FILE_NAME
    };
  }
  
  try {
    const file = Drive.Files.get(fileId);
    return {
      storageType: 'drive',
      fileId: fileId,
      fileName: file.name || ACCUMULATION_FILE_NAME,
      fileSize: file.size || 'unknown'
    };
  } catch (e) {
    return {
      storageType: 'error',
      fileId: fileId,
      fileName: ACCUMULATION_FILE_NAME,
      error: e.toString()
    };
  }
}

/**
 * Debug function: Store dummy content to test file storage
 * Only available in debug mode
 */
function debugStoreDummyContent() {
  if (!debugMode()) {
    throw new Error('This function is only available in debug mode');
  }
  
  const dummyResults = {
    mustDo: [
      {
        subject: 'Test Task 1',
        keyAction: 'This is a test actionable item',
        date: new Date().toISOString()
      },
      {
        subject: 'Test Task 2',
        keyAction: 'Another test task',
        date: new Date().toISOString()
      }
    ],
    mustKnow: [
      {
        subject: 'Test Info 1',
        keyKnowledge: 'This is a test informational item',
        date: new Date().toISOString()
      }
    ],
    totalProcessed: 42,
    firstDate: new Date(Date.now() - 24 * 60 * 60 * 1000).toISOString(),
    lastDate: new Date().toISOString()
  };
  
  saveAccumulatedResults(dummyResults);
  console.log('Debug: Stored dummy content to accumulation file');
  
  // Verify it was saved correctly
  const loaded = loadAccumulatedResults();
  console.log('Debug: Loaded content to verify:', {
    mustDoCount: loaded.mustDo.length,
    mustKnowCount: loaded.mustKnow.length,
    totalProcessed: loaded.totalProcessed
  });
  
  return {
    success: true,
    message: 'Dummy content stored successfully',
    stored: {
      mustDo: dummyResults.mustDo.length,
      mustKnow: dummyResults.mustKnow.length,
      totalProcessed: dummyResults.totalProcessed
    },
    verified: {
      mustDo: loaded.mustDo.length,
      mustKnow: loaded.mustKnow.length,
      totalProcessed: loaded.totalProcessed
    }
  };
}

/**
 * Debug function: Test Drive API with getBytes() approach
 * Generates random data, saves to Drive, loads it back, and compares
 * Only available in debug mode
 */
function debugTestDriveAPIWithGetBytes() {
  if (!debugMode()) {
    throw new Error('This function is only available in debug mode');
  }
  
  const TEST_FILE_NAME = 'gmail-fast-read-drive-api-test.json';
  const FILE_ID_PROPERTY_KEY = 'drive_api_test_file_id';
  
  const properties = PropertiesService.getUserProperties();
  let fileId = properties.getProperty(FILE_ID_PROPERTY_KEY);
  
  // Generate random test data
  const randomData = {
    timestamp: new Date().toISOString(),
    randomString: Math.random().toString(36).substring(2, 15) + Math.random().toString(36).substring(2, 15),
    randomNumber: Math.floor(Math.random() * 1000000),
    testArray: Array.from({ length: 10 }, () => Math.random()),
    nested: {
      level1: {
        level2: 'test value ' + Math.random().toString(36).substring(2, 10)
      }
    }
  };
  
  const contentToSave = JSON.stringify(randomData, null, 2);
  const blob = Utilities.newBlob(contentToSave, 'application/json', TEST_FILE_NAME);
  
  let saveResult = {};
  let loadResult = {};
  let comparisonResult = {};
  let responseInfo = {};
  
  try {
    // Step 1: Save to Drive
    if (fileId) {
      // Update existing file
      try {
        Drive.Files.update({}, fileId, blob);
        saveResult = {
          success: true,
          action: 'updated',
          fileId: fileId,
          size: contentToSave.length
        };
      } catch (error) {
        // File might not exist, create new one
        fileId = null;
        properties.deleteProperty(FILE_ID_PROPERTY_KEY);
      }
    }
    
    if (!fileId) {
      // Create new file
      const fileResource = {
        name: TEST_FILE_NAME,
        mimeType: 'application/json'
      };
      const createdFile = Drive.Files.create(fileResource, blob);
      fileId = createdFile.id;
      properties.setProperty(FILE_ID_PROPERTY_KEY, fileId);
      saveResult = {
        success: true,
        action: 'created',
        fileId: fileId,
        size: contentToSave.length
      };
    }
    
    // Step 2: Load from Drive
    const response = Drive.Files.get(fileId, { alt: 'media' });
    
    let loadedContent = null;
    let loadMethod = '';
    
    // Log response type and available methods for debugging
    responseInfo.type = typeof response;
    responseInfo.constructor = response ? (response.constructor ? response.constructor.name : 'no constructor') : 'null';
    responseInfo.hasGetDataAsString = response && (typeof response.getDataAsString === 'function' || 'getDataAsString' in response);
    responseInfo.hasGetBytes = response && (typeof response.getBytes === 'function' || 'getBytes' in response);
    responseInfo.isString = typeof response === 'string';
    responseInfo.keys = response && typeof response === 'object' ? Object.keys(response).join(', ') : 'N/A';
    
    // Use the same logic as loadFileContentFromDrive
    loadedContent = loadFileContentFromDrive(fileId);
    if (loadedContent) {
      loadMethod = 'loadFileContentFromDrive';
    } else {
      // If loadFileContentFromDrive failed, try direct approach
      // Check if response is already a parsed JSON object
      if (response && typeof response === 'object' && !Array.isArray(response)) {
        const hasGetDataAsString = typeof response.getDataAsString === 'function';
        const hasGetBytes = typeof response.getBytes === 'function';
        
        if (!hasGetDataAsString && !hasGetBytes) {
          try {
            loadedContent = JSON.stringify(response, null, 2);
            loadMethod = 'direct_json_object';
          } catch (e) {
            console.error('Failed to stringify response:', e);
          }
        }
      }
      
      // If still no content, throw error
      if (!loadedContent) {
        throw new Error('Unrecognized response type from Drive.Files.get: ' + 
          typeof response + ' (constructor: ' + (response ? response.constructor.name : 'null') + 
          ', keys: ' + (response && typeof response === 'object' ? Object.keys(response).join(', ') : 'N/A') + 
          ', hasGetDataAsString: ' + (response && typeof response.getDataAsString === 'function') +
          ', hasGetBytes: ' + (response && typeof response.getBytes === 'function') + ')');
      }
    }
    
    loadResult = {
      success: true,
      method: loadMethod,
      size: loadedContent.length,
      responseInfo: responseInfo
    };
    
    // Step 3: Compare data
    const loadedData = JSON.parse(loadedContent);
    const dataMatches = JSON.stringify(randomData) === JSON.stringify(loadedData);
    
    comparisonResult = {
      success: dataMatches,
      originalSize: contentToSave.length,
      loadedSize: loadedContent.length,
      sizesMatch: contentToSave.length === loadedContent.length,
      dataMatches: dataMatches
    };
    
    if (!dataMatches) {
      // Find differences
      const originalStr = JSON.stringify(randomData);
      const loadedStr = JSON.stringify(loadedData);
      comparisonResult.differences = {
        original: originalStr.substring(0, 200),
        loaded: loadedStr.substring(0, 200)
      };
    }
    
    return {
      success: true,
      message: 'Drive API test completed',
      save: saveResult,
      load: loadResult,
      comparison: comparisonResult,
      originalData: randomData,
      loadedData: loadedData
    };
    
  } catch (error) {
    return {
      success: false,
      message: 'Drive API test failed: ' + error.message,
      error: error.toString(),
      save: saveResult,
      load: {
        ...loadResult,
        responseInfo: responseInfo || {}
      },
      comparison: comparisonResult
    };
  }
}


