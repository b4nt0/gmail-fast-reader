/**
 * Gmail Fast Reader - Main entry points and add-on lifecycle functions
 */

/**
 * Lock mechanism for preventing concurrent workflow execution
 */

/**
 * Acquire a lock for the specified workflow type
 */
function lock(workflowType) {
  const properties = PropertiesService.getUserProperties();
  const lockData = {
    type: workflowType,
    timestamp: new Date().toISOString()
  };
  properties.setProperty('processingLock', JSON.stringify(lockData));
}

/**
 * Release the current lock
 */
function unlock() {
  const properties = PropertiesService.getUserProperties();
  properties.deleteProperty('processingLock');
}

/**
 * Calculate expected start buffer: 30% of trigger delay + 10 seconds
 */
function calculateExpectedStartBuffer(triggerDelayMs) {
  return Math.floor(triggerDelayMs * 0.3) + 600 * 1000;
}

/**
 * Centralized processing state helpers
 */
function getProcessingState() {
  const properties = PropertiesService.getUserProperties();
  return {
    status: properties.getProperty('processingStatus'),
    message: properties.getProperty('processingMessage'),
    startTime: properties.getProperty('processingStartTime'),
    timeRange: properties.getProperty('processingTimeRange'),
    chunkStartTime: properties.getProperty('chunkStartTime'),
    expectedChunkStartTime: properties.getProperty('expectedChunkStartTime')
  };
}

function startProcessingState(timeRange, expectedStartAtIso) {
  const properties = PropertiesService.getUserProperties();
  properties.setProperties({
    'processingStatus': PROCESSING_STATUS.RUNNING,
    'processingStartTime': new Date().toISOString(),
    'processingTimeRange': timeRange,
    'processingProgress': '0',
    'processingMessage': 'Starting email processing...',
    'processedThreads': '0',
    'totalThreads': '0',
    'processedMessages': '0',
    'totalMessages': '0',
    'expectedChunkStartTime': expectedStartAtIso || ''
  });
}

function markChunkStarting(nowIso) {
  const properties = PropertiesService.getUserProperties();
  properties.setProperty('chunkStartTime', nowIso || new Date().toISOString());
}

function markChunkEnded() {
  const properties = PropertiesService.getUserProperties();
  properties.deleteProperty('chunkStartTime');
}

function setExpectedNextChunkStart(triggerDelayMs) {
  const properties = PropertiesService.getUserProperties();
  const expectedBufferMs = calculateExpectedStartBuffer(triggerDelayMs);
  const expectedAt = new Date(Date.now() + triggerDelayMs + expectedBufferMs).toISOString();
  properties.setProperty('expectedChunkStartTime', expectedAt);
}

function releaseProcessingState() {
  const properties = PropertiesService.getUserProperties();
  cleanupProcessingState(properties);
  cleanupChunkState(properties);
  cleanupChunkTiming(properties);
  unlock();
}

function failProcessing(errorMessage) {
  const properties = PropertiesService.getUserProperties();
  properties.setProperties({
    'processingStatus': PROCESSING_STATUS.ERROR,
    'processingMessage': 'Processing failed: ' + errorMessage
  });
  cleanupChunkTiming(properties);
  unlock();
}

function checkAndHandleTimeout(now = new Date()) {
  const properties = PropertiesService.getUserProperties();
  const state = getProcessingState();
  if (state.status !== PROCESSING_STATUS.RUNNING) {
    return false;
  }
  // If a chunk started and exceeded timeout
  if (state.chunkStartTime) {
    const chunkStart = new Date(state.chunkStartTime);
    if (now.getTime() - chunkStart.getTime() > PROCESSING_TIMEOUT_MS) {
      properties.setProperties({
        'processingStatus': PROCESSING_STATUS.TIMEOUT,
        'processingMessage': 'Processing timed out after 10 minutes. Please try again with a smaller time range.'
      });
      cleanupChunkTiming(properties);
      try { sendProcessingTimeoutEmail(); } catch (e) { console.error('Failed to send timeout email:', e); }
      unlock();
      return true;
    }
  } else {
    // No chunk started yet; if expected start time elapsed, timeout
    const expIso = state.expectedChunkStartTime;
    if (expIso) {
      const expected = new Date(expIso);
      if (now.getTime() > expected.getTime()) {
        properties.setProperties({
          'processingStatus': PROCESSING_STATUS.TIMEOUT,
          'processingMessage': 'Processing did not start in expected time window and was timed out.'
        });
        try { sendProcessingTimeoutEmail(); } catch (e2) { console.error('Failed to send timeout email:', e2); }
        unlock();
        return true;
      }
    }
  }
  return false;
}

/**
 * Check if a lock exists and is still valid
 * Returns: { locked: boolean, type: string|null, expired: boolean }
 */
function checkLock() {
  const properties = PropertiesService.getUserProperties();
  const lockStr = properties.getProperty('processingLock');
  
  if (!lockStr) {
    return { locked: false, type: null, expired: false };
  }
  
  try {
    const lockData = JSON.parse(lockStr);
    const lockTime = new Date(lockData.timestamp);
    const currentTime = new Date();
    // Delegate timeout handling to centralized checker
    checkAndHandleTimeout(currentTime);
    
    return { 
      locked: true, 
      type: lockData.type, 
      expired: false 
    };
  } catch (error) {
    console.error('Error parsing lock data:', error);
    // Clear invalid lock data
    unlock();
    return { locked: false, type: null, expired: false };
  }
}

/**
 * Check if processing is currently running (updated to use new lock mechanism)
 */
function isProcessingRunning() {
  const properties = PropertiesService.getUserProperties();
  const status = properties.getProperty('processingStatus');
  if (status !== PROCESSING_STATUS.RUNNING) return false;
  const timedOut = checkAndHandleTimeout(new Date());
  return !timedOut && properties.getProperty('processingStatus') === PROCESSING_STATUS.RUNNING;
}

/**
 * Cleanup helpers for PropertiesService state
 * Note: Does NOT clean passive workflow properties (passiveLastProcessedTimestamp, passiveLastProcessedMessageId)
 */
function cleanupProcessingState(properties) {
  properties.deleteProperty('processingStatus');
  properties.deleteProperty('processingMessage');
  properties.deleteProperty('processingResults');
  properties.deleteProperty('processingStartTime');
  properties.deleteProperty('processingProgress');
  properties.deleteProperty('processedThreads');
  properties.deleteProperty('totalThreads');
  properties.deleteProperty('processedMessages');
  properties.deleteProperty('totalMessages');
  // Note: Intentionally NOT cleaning passive workflow properties
}

function cleanupChunkState(properties) {
  properties.deleteProperty('chunkCurrentStart');
  properties.deleteProperty('chunkEnd');
  properties.deleteProperty('chunkIndex');
  properties.deleteProperty('chunkTotalChunks');
  properties.deleteProperty('accumulatedResults');
  properties.deleteProperty('chunkStartTime');
}

function cleanupChunkTiming(properties) {
  properties.deleteProperty('chunkStartTime');
}

/**
 * Passive workflow - runs hourly to process new emails
 */
function runPassiveWorkflow() {
  const properties = PropertiesService.getUserProperties();
  
  try {
    // Check if another workflow is already running
    const lockStatus = checkLock();
    if (lockStatus.locked) {
      console.log(`Passive workflow skipped - ${lockStatus.type} workflow is already running`);
      return;
    }
    
    // Acquire lock for passive workflow
    lock('passive');
    
    // Get configuration
    const config = getConfiguration();
    if (!config.openaiApiKey) {
      console.log('Passive workflow skipped - OpenAI API key not configured');
      unlock();
      return;
    }
    
    // Calculate date range for passive workflow
    const dateRange = calculatePassiveWorkflowDateRange();
    if (!dateRange) {
      console.log('Passive workflow skipped - no new emails to process');
      unlock();
      return;
    }
    
    console.log(`Passive workflow processing emails from ${dateRange.start.toISOString()} to ${dateRange.end.toISOString()}`);
    // Mark passive chunk start
    markChunkStarting(new Date().toISOString());
    
    // Fetch and filter emails for passive workflow
    const emailThreads = fetchEmailThreadsForPassiveWorkflow(dateRange);
    
    if (emailThreads.length === 0) {
      console.log('Passive workflow completed - no emails found in range');
      unlock();
      return;
    }
    
    // Process emails using existing batch processing logic
    const results = processEmailsInBatches(emailThreads, config);
    
    // Accumulate results if interesting emails were found
    if (results.mustDo.length > 0 || results.mustKnow.length > 0) {
      // Update tracking properties with first processed message
      if (emailThreads.length > 0 && emailThreads[0].emails.length > 0) {
        const firstMessage = emailThreads[0].emails[0];
        properties.setProperties({
          'passiveLastProcessedTimestamp': firstMessage.date.toISOString(),
          'passiveLastProcessedMessageId': firstMessage.id
        });
      }
      
      // Load existing accumulated results from Drive
      const accumulated = loadAccumulatedResults();
      
      // Merge new results into accumulated results
      const mergedResults = mergeAccumulatedResults(
        accumulated,
        results,
        dateRange.start,
        dateRange.end
      );
      
      // Save updated accumulated results to Drive
      saveAccumulatedResults(mergedResults);
      
      console.log(`Passive workflow completed - accumulated results: ${mergedResults.mustDo.length} actionable and ${mergedResults.mustKnow.length} informational items (total processed: ${mergedResults.totalProcessed})`);
    } else {
      console.log('Passive workflow completed - no interesting emails found');
    }
    
    // Check if we should send daily summary (within time window and haven't sent today)
    sendDailySummaryIfNeeded(config);
    
  } catch (error) {
    console.error('Error in passive workflow:', error);
    
    // Send error notification
    try {
      const config = getConfiguration();
      const subject = `${config.addonName} - Passive Workflow Error - ${new Date().toLocaleDateString()}`;
      const body = `Passive workflow failed with the following error:\n\n${error.message}\n\nPlease check your configuration and try again.`;
      
      GmailApp.sendEmail(
        getUserEmailAddress(),
        subject,
        body,
        {
          name: config.addonName
        }
      );
    } catch (emailError) {
      console.error('Failed to send error email:', emailError);
    }
  } finally {
    // Always release lock
    markChunkEnded();
    unlock();
  }
}

/**
 * Get the last date when a summary was sent (YYYY-MM-DD format in user's timezone)
 * @returns {string|null} Date string or null if never sent
 */
function getLastSummaryDate() {
  const properties = PropertiesService.getUserProperties();
  return properties.getProperty('passiveLastSummaryDate');
}

/**
 * Set the last date when a summary was sent
 * @param {string} dateString - Date string in YYYY-MM-DD format
 */
function setLastSummaryDate(dateString) {
  const properties = PropertiesService.getUserProperties();
  properties.setProperty('passiveLastSummaryDate', dateString);
}

/**
 * Check if current time is within the sending window (21:00-23:59) in user's timezone
 * @param {string} userTimeZone - User's timezone (e.g., 'America/New_York')
 * @returns {boolean} True if within time window
 */
function isWithinTimeWindow(userTimeZone) {
  try {
    const now = new Date();
    // Get current time in user's timezone
    const userTimeString = Utilities.formatDate(now, userTimeZone, 'HH:mm');
    const parts = userTimeString.split(':');
    const hour = parseInt(parts[0], 10);
    const minute = parseInt(parts[1], 10);
    const timeInMinutes = hour * 60 + minute;
    
    // 21:00 = 21 * 60 = 1260 minutes
    // 23:59 = 23 * 60 + 59 = 1439 minutes
    return timeInMinutes >= 1260 && timeInMinutes < 1440; // 21:00 to 23:59 (inclusive of 21:00, exclusive of midnight)
  } catch (error) {
    console.error('Error checking time window:', error);
    return false;
  }
}

/**
 * Get current date string in user's timezone (YYYY-MM-DD format)
 * @param {string} userTimeZone - User's timezone (e.g., 'America/New_York')
 * @returns {string} Date string in YYYY-MM-DD format
 */
function getCurrentDateString(userTimeZone) {
  try {
    const now = new Date();
    return Utilities.formatDate(now, userTimeZone, 'yyyy-MM-dd');
  } catch (error) {
    console.error('Error getting current date string:', error);
    // Fallback to UTC
    const now = new Date();
    return Utilities.formatDate(now, 'UTC', 'yyyy-MM-dd');
  }
}

/**
 * Check if summary has already been sent today
 * @param {string} userTimeZone - User's timezone (e.g., 'America/New_York')
 * @returns {boolean} True if summary was sent today
 */
function hasSentSummaryToday(userTimeZone) {
  const lastSummaryDate = getLastSummaryDate();
  if (!lastSummaryDate) {
    return false;
  }
  
  const todayDateString = getCurrentDateString(userTimeZone);
  return lastSummaryDate === todayDateString;
}

/**
 * Check if we should send the daily summary
 * @param {string} userTimeZone - User's timezone (e.g., 'America/New_York')
 * @returns {boolean} True if should send (within time window and haven't sent today)
 */
function shouldSendDailySummary(userTimeZone) {
  if (!isWithinTimeWindow(userTimeZone)) {
    return false;
  }
  
  if (hasSentSummaryToday(userTimeZone)) {
    return false;
  }
  
  return true;
}

/**
 * Merge new results into accumulated results
 * @param {Object} accumulated - Existing accumulated results
 * @param {Object} newResults - New results from current processing
 * @param {Date} processingStartDate - Start date of current processing batch
 * @param {Date} processingEndDate - End date of current processing batch
 * @returns {Object} Merged accumulated results
 */
function mergeAccumulatedResults(accumulated, newResults, processingStartDate, processingEndDate) {
  const merged = {
    mustDo: (accumulated.mustDo || []).concat(newResults.mustDo || []),
    mustKnow: (accumulated.mustKnow || []).concat(newResults.mustKnow || []),
    totalProcessed: (accumulated.totalProcessed || 0) + (newResults.totalProcessed || 0),
    firstDate: accumulated.firstDate || processingStartDate.toISOString(),
    lastDate: processingEndDate.toISOString()
  };
  
  return merged;
}

/**
 * Send daily summary if conditions are met
 * @param {Object} config - Configuration object
 * @returns {boolean} True if summary was sent successfully, false otherwise
 */
function sendDailySummaryIfNeeded(config) {
  if (!shouldSendDailySummary(config.timeZone)) {
    return false;
  }
  
  try {
    // Load accumulated results from Drive
    const accumulated = loadAccumulatedResults();
    
    // Check if there's anything to send
    if (accumulated.mustDo.length === 0 && accumulated.mustKnow.length === 0) {
      console.log('Daily summary check: no accumulated results to send');
      return false;
    }
    
    // Prepare results for summary generation
    const summaryResults = {
      mustDo: accumulated.mustDo || [],
      mustKnow: accumulated.mustKnow || [],
      totalProcessed: accumulated.totalProcessed || 0,
      message: `Daily summary: Processed ${accumulated.totalProcessed || 0} emails.`,
      timeRange: 'daily',
      actualStartDate: accumulated.firstDate || new Date().toISOString(),
      actualEndDate: accumulated.lastDate || new Date().toISOString()
    };
    
    // Generate and send summary email
    const htmlContent = generateSummaryHTML(summaryResults, config);
    const today = getCurrentDateString(config.timeZone);
    const subject = `${config.addonName} - Daily Summary - ${today}`;
    
    GmailApp.sendEmail(
      getUserEmailAddress(),
      subject,
      '',
      {
        htmlBody: htmlContent,
        name: config.addonName
      }
    );
    
    // Mark email as important or starred
    markEmailAsImportantOrStarred(subject);
    
    // Clear accumulated results from Drive after successful send
    clearAccumulatedResults();
    
    // Update last summary date
    setLastSummaryDate(today);
    
    console.log(`Daily summary sent successfully - ${summaryResults.mustDo.length} actionable and ${summaryResults.mustKnow.length} informational items`);
    return true;
  } catch (error) {
    console.error('Error sending daily summary:', error);
    // Don't clear accumulated results on error - they'll be sent next day
    return false;
  }
}

/**
 * Calculate date range for passive workflow
 */
function calculatePassiveWorkflowDateRange() {
  const properties = PropertiesService.getUserProperties();
  const now = new Date();
  
  // Get last processed timestamp
  const lastProcessedTimestampStr = properties.getProperty('passiveLastProcessedTimestamp');
  
  let startDate;
  
  if (lastProcessedTimestampStr) {
    const lastProcessedTimestamp = new Date(lastProcessedTimestampStr);
    // Add 30 minutes safety buffer
    const safetyBuffer = new Date(lastProcessedTimestamp.getTime() + 30 * 60 * 1000);
    startDate = safetyBuffer;
  } else {
    // First run - start from 24 hours ago
    startDate = new Date(now.getTime() - 24 * 60 * 60 * 1000);
  }
  
  // Ensure we don't go back more than 24 hours
  const maxLookback = new Date(now.getTime() - 24 * 60 * 60 * 1000);
  if (startDate < maxLookback) {
    startDate = maxLookback;
  }
  
  // If start date is not before end date, no new emails to process
  if (startDate >= now) {
    return null;
  }
  
  return {
    start: startDate,
    end: now
  };
}

/**
 * Fetch email threads for passive workflow with message filtering
 * Uses the shared fetchEmailThreadsFromGmail function with a stop condition
 * to avoid reprocessing already-processed emails.
 */
function fetchEmailThreadsForPassiveWorkflow(dateRange) {
  try {
    // Get last processed message ID for stopping condition
    const properties = PropertiesService.getUserProperties();
    const lastProcessedMessageId = properties.getProperty('passiveLastProcessedMessageId');
    
    // Use shared function with stop condition to avoid reprocessing
    return fetchEmailThreadsFromGmail(dateRange, lastProcessedMessageId);
  } catch (error) {
    console.error('Error fetching email threads for passive workflow:', error);
    throw new Error('Failed to fetch email threads for passive workflow');
  }
}

/**
 * Main entry point when add-on is opened
 */
function onHomepageTrigger(e) {
  // Make sure the hourly dispatcher trigger exists whenever user visits home
  try { ensureDispatcherScheduled(); } catch (err) { console.error('Failed to ensure dispatcher:', err); }
  return buildMainCard();
}

/**
 * Entry point when Gmail message is opened (for future use)
 */
function onGmailMessageOpen(e) {
  // Future: Quick scan single email functionality
  return buildQuickScanCard(e);
}

/**
 * Check if dispatcher trigger exists
 * @return {boolean} True if dispatcher trigger is installed
 */
function isDispatcherTriggerInstalled() {
  return ScriptApp.getProjectTriggers()
    .some(function(t) { return t.getHandlerFunction() === 'runDispatcher'; });
}

/**
 * Ensure a single dispatcher trigger exists (runs every minute)
 */
function ensureDispatcherScheduled() {
  var hasDispatcher = ScriptApp.getProjectTriggers()
    .some(function(t) { return t.getHandlerFunction() === 'runDispatcher'; });
  if (!hasDispatcher) {
    // Gmail add-ons require at least hourly cadence for time-based triggers
    ScriptApp.newTrigger('runDispatcher').timeBased().everyHours(1).create();
  }
}

/**
 * Delete all hourly dispatcher triggers (used to temporarily free the slot)
 */
function deleteDispatcherTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(function(t) { return t.getHandlerFunction() === 'runDispatcher'; })
    .forEach(function(t) { ScriptApp.deleteTrigger(t); });
}

/**
   * Dispatcher: advances active processing if present; otherwise runs passive hourly
   */
function runDispatcher() {
  var properties = PropertiesService.getUserProperties();
  try {
    // Enforce centralized timeout at entry
    if (checkAndHandleTimeout(new Date())) {
      ensureDispatcherScheduled();
      return;
    }
    // If active processing is running or prepared, advance chunk
    var status = properties.getProperty('processingStatus');
    var lockStatus = checkLock();
    if (status === PROCESSING_STATUS.RUNNING) {
      // Acquire lock if missing
      if (!lockStatus.locked) {
        lock('active');
      }
      processEmailsChunkedStep();
      return;
    }

    // Otherwise run passive workflow hourly (hard-coded)
    var config = getConfiguration();
    if (config && config.openaiApiKey) {
      var lastRunIso = properties.getProperty('passiveLastRunIso');
      var now = new Date();
      var shouldRun = true;
      if (lastRunIso) {
        var lastRun = new Date(lastRunIso);
        shouldRun = (now.getTime() - lastRun.getTime()) >= (60 * 60 * 1000); // 1 hour
      }
      if (shouldRun) {
        properties.setProperty('passiveLastRunIso', now.toISOString());
        runPassiveWorkflow();
      }
    }
  } catch (error) {
    console.error('Dispatcher error:', error);
  }
}

/**
 * Helper function to safely extract form values
 */
function getFormValue(input, defaultValue = '') {
  if (!input) return defaultValue;
  if (Array.isArray(input)) return input[0] || defaultValue;
  return String(input);
}

/**
 * Helper function to safely extract boolean values
 */
function getFormBoolean(input, defaultValue = false) {
  if (!input) return defaultValue;
  if (Array.isArray(input)) return input.includes('true');
  return String(input) === 'true';
}

/**
 * Merge partial configuration with existing configuration
 * @param {Object} partialConfig - Partial config object with only the fields to update
 * @returns {Object} Complete merged configuration
 */
function mergeConfiguration(partialConfig) {
  const existing = getConfiguration();
  return {
    addonName: partialConfig.addonName !== undefined ? partialConfig.addonName : existing.addonName,
    openaiApiKey: partialConfig.openaiApiKey !== undefined ? partialConfig.openaiApiKey : existing.openaiApiKey,
    timeZone: partialConfig.timeZone !== undefined ? partialConfig.timeZone : existing.timeZone,
    mustDoTopics: partialConfig.mustDoTopics !== undefined ? partialConfig.mustDoTopics : existing.mustDoTopics,
    mustKnowTopics: partialConfig.mustKnowTopics !== undefined ? partialConfig.mustKnowTopics : existing.mustKnowTopics,
    mustDoOther: partialConfig.mustDoOther !== undefined ? partialConfig.mustDoOther : existing.mustDoOther,
    mustKnowOther: partialConfig.mustKnowOther !== undefined ? partialConfig.mustKnowOther : existing.mustKnowOther,
    unreadOnly: partialConfig.unreadOnly !== undefined ? partialConfig.unreadOnly : existing.unreadOnly,
    inboxOnly: partialConfig.inboxOnly !== undefined ? partialConfig.inboxOnly : existing.inboxOnly,
    mustDoLabel: partialConfig.mustDoLabel !== undefined ? partialConfig.mustDoLabel : existing.mustDoLabel,
    mustKnowLabel: partialConfig.mustKnowLabel !== undefined ? partialConfig.mustKnowLabel : existing.mustKnowLabel
  };
}

/**
 * Handle configuration form submission (legacy - kept for compatibility)
 */
function handleConfigSubmit(e) {
  try {
    const formInputs = e.formInputs;
    
    // Save configuration
    saveConfiguration({
      addonName: getFormValue(formInputs.addonName, 'Gmail Fast Reader'),
      openaiApiKey: getFormValue(formInputs.openaiApiKey),
      timeZone: getFormValue(formInputs.timeZone, 'Europe/Paris'),
      mustDoTopics: getFormValue(formInputs.mustDoTopics),
      mustKnowTopics: getFormValue(formInputs.mustKnowTopics),
      mustDoOther: getFormBoolean(formInputs.mustDoOther),
      mustKnowOther: getFormBoolean(formInputs.mustKnowOther),
      unreadOnly: getFormBoolean(formInputs.unreadOnly),
      inboxOnly: getFormBoolean(formInputs.inboxOnly),
      mustDoLabel: getFormValue(formInputs.mustDoLabel),
      mustKnowLabel: getFormValue(formInputs.mustKnowLabel),
      automaticScanning: getFormBoolean(formInputs.automaticScanning)
    });
    
    return buildConfigSuccessCard();
  } catch (error) {
    return buildErrorCard('Failed to save configuration: ' + error.message);
  }
}

/**
 * Handle topics form submission
 */
function handleTopicsSubmit(e) {
  try {
    const formInputs = e.formInputs || {};
    const partialConfig = {
      mustDoTopics: getFormValue(formInputs.mustDoTopics),
      mustKnowTopics: getFormValue(formInputs.mustKnowTopics),
      mustDoOther: getFormBoolean(formInputs.mustDoOther),
      mustKnowOther: getFormBoolean(formInputs.mustKnowOther)
    };
    
    const mergedConfig = mergeConfiguration(partialConfig);
    saveConfiguration(mergedConfig);
    
    return buildConfigSuccessCard();
  } catch (error) {
    return buildErrorCard('Failed to save topics: ' + error.message);
  }
}

/**
 * Handle system settings form submission
 */
function handleSystemSettingsSubmit(e) {
  try {
    const formInputs = e.formInputs || {};
    const partialConfig = {
      addonName: getFormValue(formInputs.addonName, 'Gmail Fast Reader'),
      openaiApiKey: getFormValue(formInputs.openaiApiKey),
      timeZone: getFormValue(formInputs.timeZone, 'Europe/Paris')
    };
    
    const mergedConfig = mergeConfiguration(partialConfig);
    saveConfiguration(mergedConfig);
    
    return buildConfigSuccessCard();
  } catch (error) {
    return buildErrorCard('Failed to save system settings: ' + error.message);
  }
}

/**
 * Handle email settings form submission
 */
function handleEmailSettingsSubmit(e) {
  try {
    const formInputs = e.formInputs || {};
    const partialConfig = {
      unreadOnly: getFormBoolean(formInputs.unreadOnly),
      inboxOnly: getFormBoolean(formInputs.inboxOnly),
      mustDoLabel: getFormValue(formInputs.mustDoLabel),
      mustKnowLabel: getFormValue(formInputs.mustKnowLabel)
    };
    
    const mergedConfig = mergeConfiguration(partialConfig);
    saveConfiguration(mergedConfig);
    
    return buildConfigSuccessCard();
  } catch (error) {
    return buildErrorCard('Failed to save email settings: ' + error.message);
  }
}

/**
 * Handle onboarding navigation (Previous button)
 */
function handleOnboardingNavigation(e) {
  try {
    const step = parseInt(e.parameters.step || '1', 10);
    return buildOnboardingCard(step);
  } catch (error) {
    return buildErrorCard('Failed to navigate: ' + error.message);
  }
}

/**
 * Handle onboarding save and next step
 */
function handleOnboardingSaveAndNext(e) {
  try {
    const formInputs = e.formInputs || {};
    const currentStep = parseInt(e.parameters.step || '1', 10);
    const nextStep = parseInt(e.parameters.nextStep || '2', 10);
    
    let partialConfig = {};
    
    // Step 1: System Settings
    if (currentStep === 1) {
      partialConfig = {
        addonName: getFormValue(formInputs.addonName, 'Gmail Fast Reader'),
        openaiApiKey: getFormValue(formInputs.openaiApiKey),
        timeZone: getFormValue(formInputs.timeZone, 'Europe/Paris')
      };
    }
    // Step 2: Topics
    else if (currentStep === 2) {
      partialConfig = {
        mustDoTopics: getFormValue(formInputs.mustDoTopics),
        mustKnowTopics: getFormValue(formInputs.mustKnowTopics),
        mustDoOther: getFormBoolean(formInputs.mustDoOther),
        mustKnowOther: getFormBoolean(formInputs.mustKnowOther)
      };
    }
    
    // Merge and save
    const mergedConfig = mergeConfiguration(partialConfig);
    saveConfiguration(mergedConfig);
    
    // Navigate to next step
    return buildOnboardingCard(nextStep);
  } catch (error) {
    return buildErrorCard('Failed to save and continue: ' + error.message);
  }
}

/**
 * Handle onboarding finish (last step)
 */
function handleOnboardingFinish(e) {
  try {
    const formInputs = e.formInputs || {};
    const partialConfig = {
      unreadOnly: getFormBoolean(formInputs.unreadOnly),
      inboxOnly: getFormBoolean(formInputs.inboxOnly),
      mustDoLabel: getFormValue(formInputs.mustDoLabel),
      mustKnowLabel: getFormValue(formInputs.mustKnowLabel)
    };
    
    // Merge and save
    const mergedConfig = mergeConfiguration(partialConfig);
    saveConfiguration(mergedConfig);
    
    // Return success card
    return buildConfigSuccessCard();
  } catch (error) {
    return buildErrorCard('Failed to finish setup: ' + error.message);
  }
}

/**
 * Handle nuke settings (debug function)
 */
function handleNukeSettings() {
  try {
    const properties = PropertiesService.getUserProperties();
    
    // Clear all configuration properties
    properties.deleteProperty('addonName');
    properties.deleteProperty('openaiApiKey');
    properties.deleteProperty('timeZone');
    properties.deleteProperty('mustDoTopics');
    properties.deleteProperty('mustKnowTopics');
    properties.deleteProperty('mustDoOther');
    properties.deleteProperty('mustKnowOther');
    properties.deleteProperty('unreadOnly');
    properties.deleteProperty('inboxOnly');
    properties.deleteProperty('mustDoLabel');
    properties.deleteProperty('mustKnowLabel');
    
    return buildOnboardingCard(1);
  } catch (error) {
    return buildErrorCard('Failed to nuke settings: ' + error.message);
  }
}

/**
 * Suggestions handler for label text inputs
 */
function handleLabelSuggestions(e) {
  try {
    const query = (e && e.parameter && e.parameter.query) || '';
    const labels = GmailApp.getUserLabels();
    const normalizedQuery = query.toLowerCase();
    const suggestions = CardService.newSuggestions();
    for (var i = 0; i < labels.length; i++) {
      var name = labels[i].getName();
      if (!normalizedQuery || name.toLowerCase().indexOf(normalizedQuery) !== -1) {
        suggestions.addSuggestion(name);
      }
    }
    return CardService.newSuggestionsResponseBuilder()
      .setSuggestions(suggestions)
      .build();
  } catch (error) {
    // On error, return empty suggestions
    return CardService.newSuggestionsResponseBuilder()
      .setSuggestions(CardService.newSuggestions())
      .build();
  }
}

/**
 * Handle active workflow - scan emails
 */
function handleScanEmails(e) {
  try {
    // Check if already running
    if (isProcessingRunning()) {
      return buildErrorCard('Another workflow is already running. Please wait or check status.');
    }
    
    const timeRange = e.parameters.timeRange || '1day';
    
    // Start background processing
    startBackgroundEmailProcessing(timeRange);
    
    // Return progress card with Check Status button
    return buildProgressCardWithStatusButton('Email processing started in background. You will receive an email when complete.');
    
  } catch (error) {
    console.error('Error in handleScanEmails:', error);
    return buildErrorCard('Failed to start email processing: ' + error.message);
  }
}

/**
 * Start background email processing (always uses chunked approach)
 */
function startBackgroundEmailProcessing(timeRange) {
  // Acquire lock for active workflow
  lock('active');

  try {
  
    // Store processing status
    const properties = PropertiesService.getUserProperties();
    // Schedule parameters: one-off trigger delay and expected start buffer
    const triggerDelayMs = 60 * 1000; // 1 minute kickoff
    const expectedBufferMs = calculateExpectedStartBuffer(triggerDelayMs);
    const expectedAt = new Date(Date.now() + expectedBufferMs).toISOString();
    startProcessingState(timeRange, expectedAt);

    // Initialize chunked processing with 2-day chunks
    const dateRange = calculateDateRange(timeRange);
    const start = new Date(dateRange.start);
    const end = new Date(dateRange.end);

    // Compute chunk count based on chunk size
    const totalChunks = Math.max(1, Math.ceil((end.getTime() - start.getTime()) / CHUNK_SIZE_MS));

    properties.setProperties({
      'chunkCurrentStart': start.toISOString(),
      'chunkEnd': end.toISOString(),
      'chunkIndex': '0',
      'chunkTotalChunks': String(totalChunks),
      'accumulatedResults': JSON.stringify({ mustDo: [], mustKnow: [], totalProcessed: 0, batchesProcessed: 0 })
    });

    properties.setProperty('processingMessage', `Processing emails in ${totalChunks} chunks...`);
    // Temporarily remove hourly dispatcher and schedule a one-off active step in 1 minute
    try {
      deleteDispatcherTriggers();
      ScriptApp.newTrigger('processEmailsChunkedStep')
        .timeBased()
        .after(triggerDelayMs)
        .create();
    } catch (error) {
      console.error('Failed to schedule one-off active trigger:', error);
      failProcessing(error.message);
      try { sendProcessingErrorEmail(error.message); } catch (e2) { console.error('Failed to send error email:', e2); }
      // Always release lock on failure and restore dispatcher
      try { ensureDispatcherScheduled(); } catch (e3) { console.error('Failed to restore dispatcher:', e3); }
    }
  } catch (error) {
    console.error('Failed to start background email processing:', error);
    unlock();
  }
}

/**
 * Unified chunked background processing - processes one chunk per invocation and chains next trigger
 */
function processEmailsChunkedStep() {
  const properties = PropertiesService.getUserProperties();
  try {
    // As soon as the one-off fires, restore the hourly dispatcher
    try { ensureDispatcherScheduled(); } catch (e0) { console.error('Failed to ensure dispatcher at step start:', e0); }
    // Mark chunk starting before heavy work
    markChunkStarting(new Date().toISOString());
    const config = getConfiguration();
    if (!config.openaiApiKey) {
      throw new Error('OpenAI API key not configured');
    }

    // Load chunk state
    const chunkStartIso = properties.getProperty('chunkCurrentStart');
    const chunkEndIso = properties.getProperty('chunkEnd');
    const chunkIndex = parseInt(properties.getProperty('chunkIndex') || '0');
    const totalChunks = parseInt(properties.getProperty('chunkTotalChunks') || '1');
    const accumulatedStr = properties.getProperty('accumulatedResults') || '{}';
    const accumulated = accumulatedStr ? JSON.parse(accumulatedStr) : { mustDo: [], mustKnow: [], totalProcessed: 0, batchesProcessed: 0 };

    if (!chunkStartIso || !chunkEndIso) {
      throw new Error('Chunked processing state is missing');
    }

    const overallStart = new Date(chunkStartIso);
    const overallEnd = new Date(chunkEndIso);

    // Compute this chunk's window
    const currentStart = new Date(overallStart.getTime() + (chunkIndex * CHUNK_SIZE_MS));
    const currentStop = new Date(Math.min(currentStart.getTime() + CHUNK_SIZE_MS, overallEnd.getTime()));

    // If we've reached or passed the end, finalize
    if (currentStart.getTime() >= overallEnd.getTime()) {
      finalizeChunkedProcessing(accumulated);
      return;
    }

    properties.setProperty('processingMessage', `Processing chunk ${chunkIndex + 1} of ${totalChunks} (${currentStart.toDateString()} to ${currentStop.toDateString()})...`);

    // chunkStartTime already set at the beginning via markChunkStarting

    // Fetch and process threads for this chunk
    const dateRange = { start: currentStart, end: currentStop };
    
    // Log chunk information before fetching
    console.log(`=== CHUNK ${chunkIndex + 1}/${totalChunks} START ===`);
    console.log(`Chunk Date Range: ${currentStart.toISOString()} to ${currentStop.toISOString()}`);
    
    const emailThreads = fetchEmailThreadsFromGmail(dateRange);
    
    // Log chunk selection results
    console.log(`=== CHUNK ${chunkIndex + 1}/${totalChunks} SELECTION RESULTS ===`);
    console.log(`Chunk ${chunkIndex + 1} - Threads: ${emailThreads.length}`);
    const chunkEmailCount = emailThreads.reduce((sum, thread) => sum + thread.emails.length, 0);
    console.log(`Chunk ${chunkIndex + 1} - Total Emails: ${chunkEmailCount}`);
    
    let chunkResults = { mustDo: [], mustKnow: [], totalProcessed: 0, batchesProcessed: 0 };
    
    if (emailThreads.length > 0) {
      const totalThreads = emailThreads.length;
      const totalMessages = emailThreads.reduce((total, thread) => total + thread.emails.length, 0);
      
      // Update progress tracking
      const currentProcessedThreads = parseInt(properties.getProperty('processedThreads') || '0');
      const currentProcessedMessages = parseInt(properties.getProperty('processedMessages') || '0');
      
      properties.setProperties({
        'totalThreads': (parseInt(properties.getProperty('totalThreads') || '0') + totalThreads).toString(),
        'totalMessages': (parseInt(properties.getProperty('totalMessages') || '0') + totalMessages).toString(),
        'processedThreads': (currentProcessedThreads + totalThreads).toString(),
        'processedMessages': (currentProcessedMessages + totalMessages).toString()
      });
      
      chunkResults = processEmailsInBatches(emailThreads, config);
    }
    
    // Accumulate results from this chunk
    accumulated.mustDo = (accumulated.mustDo || []).concat(chunkResults.mustDo || []);
    accumulated.mustKnow = (accumulated.mustKnow || []).concat(chunkResults.mustKnow || []);
    accumulated.totalProcessed = (accumulated.totalProcessed || 0) + (chunkResults.totalProcessed || 0);
    accumulated.batchesProcessed = (accumulated.batchesProcessed || 0) + (chunkResults.batchesProcessed || 0);

    // Advance to next chunk
    const nextChunkIndex = chunkIndex + 1;
    properties.setProperties({
      'accumulatedResults': JSON.stringify(accumulated),
      'chunkIndex': String(nextChunkIndex)
    });

    // If finished all chunks, finalize and send summary
    if (overallStart.getTime() + (nextChunkIndex * CHUNK_SIZE_MS) >= overallEnd.getTime() || nextChunkIndex >= totalChunks) {
      finalizeChunkedProcessing(accumulated);
      return;
    }

    // Mark this chunk as ended and set expected start for next chunk (dispatcher runs hourly)
    markChunkEnded();
    const dispatcherIntervalMs = 60 * 60 * 1000; // 1 hour
    setExpectedNextChunkStart(dispatcherIntervalMs);
    ensureDispatcherScheduled();

  } catch (error) {
    console.error('Error in chunked background processing:', error);
    properties.setProperties({
      'processingStatus': PROCESSING_STATUS.ERROR,
      'processingMessage': 'Processing failed: ' + error.message
    });
    // Clear chunk timing on error
    cleanupChunkTiming(properties);
    sendProcessingErrorEmail(error.message);
    // Release lock on error
    unlock();
    // Ensure dispatcher remains scheduled
    ensureDispatcherScheduled();
  }
}

/**
 * Finalize chunked processing: store results and send email
 */
function finalizeChunkedProcessing(accumulated) {
  const properties = PropertiesService.getUserProperties();
  try {
    const mustDo = (accumulated && accumulated.mustDo) ? accumulated.mustDo : [];
    const mustKnow = (accumulated && accumulated.mustKnow) ? accumulated.mustKnow : [];
    const totalProcessed = (accumulated && accumulated.totalProcessed) ? accumulated.totalProcessed : 0;

    // Get time range information before clearing properties
    const timeRange = properties.getProperty('processingTimeRange') || '1day';
    const chunkStartIso = properties.getProperty('chunkCurrentStart');
    const chunkEndIso = properties.getProperty('chunkEnd');
    
    // Calculate actual start and end dates
    let actualStartDate, actualEndDate;
    if (chunkStartIso && chunkEndIso) {
      actualStartDate = new Date(chunkStartIso);
      actualEndDate = new Date(chunkEndIso);
    } else {
      // Fallback to calculating from timeRange
      const dateRange = calculateDateRange(timeRange);
      actualStartDate = dateRange.start;
      actualEndDate = dateRange.end;
    }

    const finalResults = {
      mustDo: mustDo,
      mustKnow: mustKnow,
      totalProcessed: totalProcessed,
      message: `Processed ${totalProcessed} emails across chunked processing.`,
      timeRange: timeRange,
      actualStartDate: actualStartDate.toISOString(),
      actualEndDate: actualEndDate.toISOString()
    };

    properties.setProperties({
      'processingStatus': PROCESSING_STATUS.COMPLETED,
      'processingMessage': `Processing complete. Found ${mustDo.length} actionable items and ${mustKnow.length} informational items.`,
      'processingResults': JSON.stringify(finalResults)
    });

    // Clear chunked-specific state
    cleanupChunkState(properties);

    if (mustDo.length > 0 || mustKnow.length > 0) {
      sendProcessingCompleteEmail();
    }
  } catch (error) {
    console.error('Error finalizing chunked processing:', error);
    properties.setProperties({
      'processingStatus': PROCESSING_STATUS.ERROR,
      'processingMessage': 'Processing failed: ' + error.message
    });
    // Clear chunk timing on error
    cleanupChunkTiming(properties);
    sendProcessingErrorEmail(error.message);
  } finally {
    // Always release lock for active workflow
    unlock();
    // Ensure dispatcher remains scheduled
    ensureDispatcherScheduled();
  }
}

/**
 * Send processing complete email
 */
function sendProcessingCompleteEmail() {
  const properties = PropertiesService.getUserProperties();
  const results = JSON.parse(properties.getProperty('processingResults') || '{}');
  const config = getConfiguration();
  
  const htmlContent = generateSummaryHTML(results, config);
  
  // Format time range for subject
  let timeRangeSubject = 'Processing Complete';
  if (results.actualStartDate && results.actualEndDate) {
    const startDate = new Date(results.actualStartDate);
    const endDate = new Date(results.actualEndDate);
    const startFormatted = startDate.toISOString().slice(0, 16).replace('T', ' ');
    const endFormatted = endDate.toISOString().slice(0, 16).replace('T', ' ');
    timeRangeSubject = `Summary - ${startFormatted} to ${endFormatted}`;
  } else if (results.timeRange) {
    timeRangeSubject = `Summary (${results.timeRange})`;
  }
  
  const subject = `${config.addonName} - ${timeRangeSubject}`;
  
  GmailApp.sendEmail(
    getUserEmailAddress(),
    subject,
    '',
    {
      htmlBody: htmlContent,
      name: config.addonName
    }
  );
  
  // Mark email as important or starred
  markEmailAsImportantOrStarred(subject);
}

/**
 * Send processing error email
 */
function sendProcessingErrorEmail(errorMessage) {
  const config = getConfiguration();
  const subject = `${config.addonName} - Processing Error - ${new Date().toLocaleDateString()}`;
  const body = `Email processing failed with the following error:\n\n${errorMessage}\n\nPlease check your configuration and try again.`;
  
  GmailApp.sendEmail(
    getUserEmailAddress(),
    subject,
    body,
    {
      name: config.addonName
    }
  );
}

/**
 * Send processing timeout email
 */
function sendProcessingTimeoutEmail() {
  const config = getConfiguration();
  const subject = `${config.addonName} - Processing Timeout - ${new Date().toLocaleDateString()}`;
  const body = 'Email processing timed out.\n\nThis usually happens when processing a large number of emails. Please try:\n\n1. Using a shorter time range (e.g., 6 hours instead of 7 days)\n2. Reducing the number of topics to focus on\n3. Checking your internet connection\n\nYou can check the status in the Gmail Fast Reader add-on.';
  
  GmailApp.sendEmail(
    getUserEmailAddress(),
    subject,
    body,
    {
      name: config.addonName
    }
  );
}

/**
 * Emergency reset: clear state, delete triggers, restore hourly dispatcher
 */
function resetAddonState() {
  try {
    // Delete all triggers for this project
    ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });
  } catch (e) {
    console.error('Failed deleting triggers:', e);
  }
  try {
    releaseProcessingState();
  } catch (e2) {
    console.error('Failed releasing processing state:', e2);
  }
  try {
    ensureDispatcherScheduled();
  } catch (e3) {
    console.error('Failed ensuring dispatcher after reset:', e3);
  }
}

function handleEmergencyReset() {
  try {
    resetAddonState();
    return buildConfigSuccessCard();
  } catch (e) {
    return buildErrorCard('Emergency reset failed: ' + e.message);
  }
}

/**
 * Handle reinstall dispatcher trigger action from UI
 */
function handleReinstallDispatcher() {
  try {
    ensureDispatcherScheduled();
    return CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader()
        .setTitle('Timer Job Reinstalled'))
      .addSection(CardService.newCardSection()
        .addWidget(CardService.newTextParagraph()
          .setText('✅ The automatic email scanning timer has been successfully reinstalled. Emails will now be scanned automatically on a regular basis.')))
      .addSection(CardService.newCardSection()
        .addWidget(CardService.newButtonSet()
          .addButton(CardService.newTextButton()
            .setText('🏠 Back to Main')
            .setOnClickAction(CardService.newAction()
              .setFunctionName('buildMainCard')))))
      .build();
  } catch (e) {
    return buildErrorCard('Failed to reinstall timer job: ' + e.message);
  }
}

/**
 * Handle time range selection for active workflow
 */
function handleTimeRangeSelection(e) {
  const timeRange = e.parameters.timeRange;
  return buildActiveWorkflowCard(timeRange);
}

/**
 * Check processing status
 */
function checkProcessingStatus() {
  const properties = PropertiesService.getUserProperties();
  const status = properties.getProperty('processingStatus');
  
  if (!status) {
    // No active processing and no status - show main card
    const latestRunStats = properties.getProperty('latestRunStats');
    if (latestRunStats) {
      return buildLatestRunStatsCard();
    }
    else {
      return buildMainCard();
    }
  }
  
  if (status === PROCESSING_STATUS.RUNNING) {
    // Single timeout check
    if (checkAndHandleTimeout(new Date())) {
      saveLatestRunStats(properties, PROCESSING_STATUS.TIMEOUT, 'Processing timed out.');
      cleanupProcessingState(properties);
      cleanupChunkTiming(properties);
      return buildLatestRunStatsCard();
    }
    const startTimeStr = properties.getProperty('processingStartTime');
    const message = properties.getProperty('processingMessage') || 'Processing...';
    const progress = properties.getProperty('processingProgress') || '0';
    const processedThreads = parseInt(properties.getProperty('processedThreads') || '0');
    const totalThreads = parseInt(properties.getProperty('totalThreads') || '0');
    const processedMessages = parseInt(properties.getProperty('processedMessages') || '0');
    const totalMessages = parseInt(properties.getProperty('totalMessages') || '0');
    const chunkStartTimeStr = properties.getProperty('chunkStartTime');
    const expectedStartStr = properties.getProperty('expectedChunkStartTime');
    
    // Build detailed progress message
    let progressMessage = message;
    // If no chunk is currently running but we have an upcoming expected start, show a waiting line
    if (!chunkStartTimeStr && expectedStartStr) {
      const now = new Date();
      const expectedAt = new Date(expectedStartStr);
      if (expectedAt.getTime() > now.getTime()) {
        const remainingMs = expectedAt.getTime() - now.getTime();
        const remainingMin = Math.floor(remainingMs / 60000);
        const remainingSec = Math.floor((remainingMs % 60000) / 1000);
        progressMessage = `⏳ Waiting for next scheduled trigger: ~${expectedAt.toLocaleString()} (in ${remainingMin}m ${remainingSec}s)`;
      }
    }
    
    // Add progress information if available
    if (totalThreads > 0 || totalMessages > 0) {
      progressMessage += `\n\n📊 Progress: ${progress}%`;
      
      if (totalThreads > 0) {
        progressMessage += `\n📧 Threads: ${processedThreads}/${totalThreads}`;
      }
      
      if (totalMessages > 0) {
        progressMessage += `\n💬 Messages: ${processedMessages}/${totalMessages}`;
      }
    } else {
      // Show initial status when counts aren't available yet
      progressMessage += `\n\n📊 Progress: ${progress}%`;
    }
    
    // Add elapsed time
    if (startTimeStr) {
      const startTime = new Date(startTimeStr);
      const currentTime = new Date();
      const timeDiff = currentTime.getTime() - startTime.getTime();
      const elapsedMinutes = Math.floor(timeDiff / (1000 * 60));
      const elapsedSeconds = Math.floor((timeDiff % (1000 * 60)) / 1000);
      progressMessage += `\n\n⏱️ Elapsed time: ${elapsedMinutes}m ${elapsedSeconds}s`;
    }
    
    return buildProgressCardWithAutoRefresh(progressMessage);
  }
  
  if (status === PROCESSING_STATUS.COMPLETED) {
    const results = JSON.parse(properties.getProperty('processingResults') || '{}');
    const startTimeStr = properties.getProperty('processingStartTime');
    const processedThreads = parseInt(properties.getProperty('processedThreads') || '0');
    const totalThreads = parseInt(properties.getProperty('totalThreads') || '0');
    const processedMessages = parseInt(properties.getProperty('processedMessages') || '0');
    const totalMessages = parseInt(properties.getProperty('totalMessages') || '0');
    
    // Calculate duration
    let duration = 0;
    if (startTimeStr) {
      const startTime = new Date(startTimeStr);
      const endTime = new Date();
      duration = endTime.getTime() - startTime.getTime();
    }
    
    // Save latest run stats before clearing
    saveLatestRunStats(properties, PROCESSING_STATUS.COMPLETED, 'Processing completed successfully', {
      results: results,
      duration: duration,
      processedThreads: processedThreads,
      totalThreads: totalThreads,
      processedMessages: processedMessages,
      totalMessages: totalMessages
    });
    
    // Clear processing status
    cleanupProcessingState(properties);
    cleanupChunkTiming(properties);
    
    // Show latest run statistics
    return buildLatestRunStatsCard();
  }
  
  if (status === PROCESSING_STATUS.ERROR) {
    const message = properties.getProperty('processingMessage') || 'Processing failed';
    const startTimeStr = properties.getProperty('processingStartTime');
    const processedThreads = parseInt(properties.getProperty('processedThreads') || '0');
    const totalThreads = parseInt(properties.getProperty('totalThreads') || '0');
    const processedMessages = parseInt(properties.getProperty('processedMessages') || '0');
    const totalMessages = parseInt(properties.getProperty('totalMessages') || '0');
    
    // Calculate duration
    let duration = 0;
    if (startTimeStr) {
      const startTime = new Date(startTimeStr);
      const endTime = new Date();
      duration = endTime.getTime() - startTime.getTime();
    }
    
    // Save latest run stats before clearing
    saveLatestRunStats(properties, PROCESSING_STATUS.ERROR, message, {
      duration: duration,
      processedThreads: processedThreads,
      totalThreads: totalThreads,
      processedMessages: processedMessages,
      totalMessages: totalMessages
    });
    
    // Clear processing status
    cleanupProcessingState(properties);
    cleanupChunkTiming(properties);
    
    // Show latest run statistics
    return buildLatestRunStatsCard();
  }
  
  if (status === PROCESSING_STATUS.TIMEOUT) {
    const message = properties.getProperty('processingMessage') || 'Processing timed out';
    const startTimeStr = properties.getProperty('processingStartTime');
    const processedThreads = parseInt(properties.getProperty('processedThreads') || '0');
    const totalThreads = parseInt(properties.getProperty('totalThreads') || '0');
    const processedMessages = parseInt(properties.getProperty('processedMessages') || '0');
    const totalMessages = parseInt(properties.getProperty('totalMessages') || '0');
    
    // Calculate duration
    let duration = 0;
    if (startTimeStr) {
      const startTime = new Date(startTimeStr);
      const endTime = new Date();
      duration = endTime.getTime() - startTime.getTime();
    }
    
    // Save latest run stats before clearing
    saveLatestRunStats(properties, PROCESSING_STATUS.TIMEOUT, message, {
      duration: duration,
      processedThreads: processedThreads,
      totalThreads: totalThreads,
      processedMessages: processedMessages,
      totalMessages: totalMessages
    });
    
    // Clear processing status
    cleanupProcessingState(properties);
    cleanupChunkTiming(properties);
    
    // Show latest run statistics
    return buildLatestRunStatsCard();
  }
  
  return buildMainCard();
}

/**
 * Save latest run statistics to properties
 */
function saveLatestRunStats(properties, status, message, additionalData = {}) {
  const endTime = new Date();
  
  const latestStats = {
    status: status,
    message: message,
    endTime: endTime.toISOString(),
    ...additionalData
  };
  
  properties.setProperty('latestRunStats', JSON.stringify(latestStats));
}

/**
 * Build card showing latest run statistics
 */
function buildLatestRunStatsCard() {
  const properties = PropertiesService.getUserProperties();
  const latestStatsStr = properties.getProperty('latestRunStats');
  
  if (!latestStatsStr) {
    return buildMainCard();
  }
  
  try {
    const latestStats = JSON.parse(latestStatsStr);
    const config = getConfiguration();
    
    // Format duration
    let durationText = 'Unknown';
    if (latestStats.duration) {
      const durationMinutes = Math.floor(latestStats.duration / (1000 * 60));
      const durationSeconds = Math.floor((latestStats.duration % (1000 * 60)) / 1000);
      durationText = `${durationMinutes}m ${durationSeconds}s`;
    }
    
    // Format end time
    const endTime = new Date(latestStats.endTime);
    const endTimeText = endTime.toLocaleString();
    
    // Build status-specific content
    let statusIcon = '❓';
    let statusColor = 'blue';
    let statusText = latestStats.status;
    
    switch (latestStats.status) {
    case PROCESSING_STATUS.COMPLETED:
      statusIcon = '✅';
      statusColor = 'green';
      statusText = 'Completed Successfully';
      break;
    case PROCESSING_STATUS.ERROR:
      statusIcon = '❌';
      statusColor = 'red';
      statusText = 'Failed';
      break;
    case PROCESSING_STATUS.TIMEOUT:
      statusIcon = '⏰';
      statusColor = 'orange';
      statusText = 'Timed Out';
      break;
    }
    
    // Build statistics text
    let statsText = `${statusIcon} Status: ${statusText}\n`;
    statsText += `📅 Completed: ${endTimeText}\n`;
    statsText += `⏱️ Duration: ${durationText}\n`;
    
    if (latestStats.totalThreads !== undefined) {
      statsText += `📧 Threads: ${latestStats.processedThreads || 0}/${latestStats.totalThreads}\n`;
    }
    
    if (latestStats.totalMessages !== undefined) {
      statsText += `💬 Messages: ${latestStats.processedMessages || 0}/${latestStats.totalMessages}\n`;
    }
    
    if (latestStats.results) {
      const results = latestStats.results;
      if (results.mustDo && results.mustDo.length > 0) {
        statsText += `🎯 Actionable Items: ${results.mustDo.length}\n`;
      }
      if (results.mustKnow && results.mustKnow.length > 0) {
        statsText += `📚 Informational Items: ${results.mustKnow.length}\n`;
      }
    }
    
    statsText += `\n💬 Message: ${latestStats.message}`;
    
    return CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader()
        .setTitle(`${config.addonName} - Latest Run Statistics`)
        .setSubtitle(`Last run: ${endTimeText}`)
        .setImageUrl('https://www.gstatic.com/images/icons/material/system/1x/assessment_black_24dp.png')
        .setImageStyle(CardService.ImageStyle.CIRCLE))
      .addSection(CardService.newCardSection()
        .addWidget(CardService.newTextParagraph()
          .setText(statsText)))
      .addSection(CardService.newCardSection()
        .addWidget(CardService.newTextButton()
          .setText('🧯 Emergency reset')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('handleEmergencyReset'))))
      .addSection(CardService.newCardSection()
        .addWidget(CardService.newTextButton()
          .setText('🔄 Start New Scan')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('buildActiveWorkflowCard'))))
      .addSection(CardService.newCardSection()
        .addWidget(CardService.newTextButton()
          .setText('⚙️ Configuration')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('buildConfigurationCard'))))
      .addSection(CardService.newCardSection()
        .addWidget(CardService.newTextButton()
          .setText('🏠 Main Menu')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('buildMainCard'))))
      .build();
      
  } catch (error) {
    console.error('Error building latest run stats card:', error);
    return buildMainCard();
  }
}

// Handlers for UI (called from toggle in both config & scan cards):
// Removed automation toggle; passive is hard-coded hourly via dispatcher
