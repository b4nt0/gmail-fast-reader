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
    const timeDiff = currentTime.getTime() - lockTime.getTime();
    
    // Check if lock has expired (10 minutes timeout)
    if (timeDiff > PROCESSING_TIMEOUT_MS) {
      // Auto-release expired lock
      unlock();
      return { locked: false, type: null, expired: true };
    }
    
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
  const lockStatus = checkLock();
  
  if (!lockStatus.locked) {
    return false;
  }
  
  // For backward compatibility, also check the old processingStatus
  const properties = PropertiesService.getUserProperties();
  const status = properties.getProperty('processingStatus');
  
  if (status !== PROCESSING_STATUS.RUNNING) {
    return false;
  }
  
  // Check for chunk timeout (not overall processing timeout)
  const chunkStartTimeStr = properties.getProperty('chunkStartTime');
  if (chunkStartTimeStr) {
    const chunkStartTime = new Date(chunkStartTimeStr);
    const currentTime = new Date();
    const timeDiff = currentTime.getTime() - chunkStartTime.getTime();
    
    if (timeDiff > PROCESSING_TIMEOUT_MS) {
      // Current chunk has timed out
      properties.setProperties({
        'processingStatus': PROCESSING_STATUS.TIMEOUT,
        'processingMessage': 'Processing timed out after 10 minutes. Please try again with a smaller time range.'
      });
      
      // Send timeout notification email
      try {
        sendProcessingTimeoutEmail();
      } catch (emailError) {
        console.error('Failed to send timeout email:', emailError);
      }
      
      return false;
    }
  }
  
  return true;
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
    
    // Fetch and filter emails for passive workflow
    const emailThreads = fetchEmailThreadsForPassiveWorkflow(dateRange);
    
    if (emailThreads.length === 0) {
      console.log('Passive workflow completed - no emails found in range');
      unlock();
      return;
    }
    
    // Process emails using existing batch processing logic
    const results = processEmailsInBatches(emailThreads, config);
    
    // Only send summary email if interesting emails were found
    if (results.mustDo.length > 0 || results.mustKnow.length > 0) {
      // Update tracking properties with first processed message
      if (emailThreads.length > 0 && emailThreads[0].emails.length > 0) {
        const firstMessage = emailThreads[0].emails[0];
        properties.setProperties({
          'passiveLastProcessedTimestamp': firstMessage.date.toISOString(),
          'passiveLastProcessedMessageId': firstMessage.id
        });
      }
      
      // Send summary email
      const finalResults = {
        mustDo: results.mustDo || [],
        mustKnow: results.mustKnow || [],
        totalProcessed: results.totalProcessed || 0,
        message: `Processed ${results.totalProcessed || 0} emails in passive workflow.`,
        timeRange: 'passive',
        actualStartDate: dateRange.start.toISOString(),
        actualEndDate: dateRange.end.toISOString()
      };
      
      const htmlContent = generateSummaryHTML(finalResults, config);
      const subject = `${config.addonName} - Passive Workflow Summary - ${new Date().toLocaleDateString()}`;
      
      GmailApp.sendEmail(
        getUserEmailAddress(),
        subject,
        '',
        {
          htmlBody: htmlContent,
          name: config.addonName
        }
      );
      
      console.log(`Passive workflow completed - sent summary with ${finalResults.mustDo.length} actionable and ${finalResults.mustKnow.length} informational items`);
    } else {
      console.log('Passive workflow completed - no interesting emails found');
    }
    
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
    unlock();
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
 */
function fetchEmailThreadsForPassiveWorkflow(dateRange) {
  try {
    // Get configuration for filtering options
    const config = getConfiguration();
    
    // Build base query with date range
    let query = `after:${Math.floor(dateRange.start.getTime() / 1000)} before:${Math.floor(dateRange.end.getTime() / 1000)}`;
    
    // Add filtering criteria based on configuration
    if (config.unreadOnly) {
      query += ' is:unread';
    }
    
    if (config.inboxOnly) {
      query += ' in:inbox';
    }
    
    const threads = GmailApp.search(query, 0, 100);
    const emailThreads = [];
    
    // Get user email and addon name for filtering
    const userEmail = getUserEmailAddress();
    const addonName = config.addonName;
    
    // Get last processed message ID for stopping condition
    const properties = PropertiesService.getUserProperties();
    const lastProcessedMessageId = properties.getProperty('passiveLastProcessedMessageId');
    
    threads.forEach(thread => {
      const messages = thread.getMessages();
      const threadEmails = [];
      let shouldStop = false;
      
      messages.forEach(message => {
        // Stop if we've reached the last processed message
        if (lastProcessedMessageId && message.getId() === lastProcessedMessageId) {
          shouldStop = true;
          return;
        }
        
        if (message.getDate() >= dateRange.start && message.getDate() <= dateRange.end) {
          // Get RFC822 message ID for permalink generation
          let rfc822MessageId = null;
          try {
            const rawContent = message.getRawContent();
            const messageIdMatch = rawContent.match(/Message-ID:\s*<([^>]+)>/i);
            if (messageIdMatch) {
              rfc822MessageId = messageIdMatch[1];
            }
          } catch (error) {
            console.warn('Could not get RFC822 message ID for message:', message.getId(), error);
          }
          
          const email = {
            id: message.getId(),
            subject: message.getSubject(),
            sender: message.getFrom(),
            date: message.getDate(),
            body: message.getPlainBody(),
            snippet: message.getPlainBody().substring(0, 200),
            rfc822MessageId: rfc822MessageId
          };
          
          // Only include emails that should not be ignored
          if (!shouldIgnoreEmail(email, userEmail, addonName)) {
            threadEmails.push(email);
          }
        }
      });
      
      if (threadEmails.length > 0) {
        // Sort emails in thread by date (oldest first)
        threadEmails.sort((a, b) => a.date - b.date);
        emailThreads.push({
          threadId: thread.getId(),
          subject: thread.getFirstMessageSubject(),
          emails: threadEmails,
          totalEmails: threadEmails.length,
          latestDate: threadEmails[threadEmails.length - 1].date
        });
      }
      
      // Stop processing if we found the last processed message
      if (shouldStop) {
        return false; // Break out of forEach
      }
    });
    
    // Sort threads by latest email date (most recent first)
    emailThreads.sort((a, b) => b.latestDate - a.latestDate);
    
    return emailThreads;
  } catch (error) {
    console.error('Error fetching email threads for passive workflow:', error);
    throw new Error('Failed to fetch email threads for passive workflow');
  }
}

/**
 * Main entry point when add-on is opened
 */
function onHomepageTrigger(e) {
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
 * Setup passive workflow trigger (hourly)
 */
function setupPassiveWorkflowTrigger() {
  try {
    const properties = PropertiesService.getUserProperties();
    
    // Delete existing trigger if it exists
    const existingTriggerId = properties.getProperty('passiveWorkflowTriggerId');
    if (existingTriggerId) {
      try {
        const trigger = ScriptApp.getProjectTriggers().find(t => t.getUniqueId() === existingTriggerId);
        if (trigger) {
          ScriptApp.deleteTrigger(trigger);
        }
      } catch (error) {
        console.warn('Could not delete existing passive workflow trigger:', error);
      }
    }
    
    // Create new hourly trigger
    const trigger = ScriptApp.newTrigger('runPassiveWorkflow')
      .timeBased()
      .everyHours(1)
      .create();
    
    // Store trigger ID for future reference
    properties.setProperty('passiveWorkflowTriggerId', trigger.getUniqueId());
    
    console.log('Passive workflow trigger created successfully');
  } catch (error) {
    console.error('Error setting up passive workflow trigger:', error);
  }
}

/**
 * Handle configuration form submission
 */
function handleConfigSubmit(e) {
  try {
    const formInputs = e.formInputs;
    
    // Helper function to safely extract form values
    function getFormValue(input, defaultValue = '') {
      if (!input) return defaultValue;
      if (Array.isArray(input)) return input[0] || defaultValue;
      return String(input);
    }
    
    // Helper function to safely extract boolean values
    function getFormBoolean(input, defaultValue = false) {
      if (!input) return defaultValue;
      if (Array.isArray(input)) return input.includes('true');
      return String(input) === 'true';
    }
    
    // Save configuration
    saveConfiguration({
      addonName: getFormValue(formInputs.addonName, 'Gmail Fast Reader'),
      openaiApiKey: getFormValue(formInputs.openaiApiKey),
      timeZone: getFormValue(formInputs.timeZone, 'America/New_York'),
      mustDoTopics: getFormValue(formInputs.mustDoTopics),
      mustKnowTopics: getFormValue(formInputs.mustKnowTopics),
      mustDoOther: getFormBoolean(formInputs.mustDoOther),
      mustKnowOther: getFormBoolean(formInputs.mustKnowOther),
      unreadOnly: getFormBoolean(formInputs.unreadOnly),
      inboxOnly: getFormBoolean(formInputs.inboxOnly),
      starInterestingEmails: getFormBoolean(formInputs.starInterestingEmails),
      automaticScanning: getFormBoolean(formInputs.automaticScanning)
    });
    
    return buildConfigSuccessCard();
  } catch (error) {
    return buildErrorCard('Failed to save configuration: ' + error.message);
  }
}

/**
 * Handle active workflow - scan emails
 */
function handleScanEmails(e) {
  try {
    // Check if processing is already running using new lock mechanism
    const lockStatus = checkLock();
    if (lockStatus.locked) {
      return buildErrorCard(`Another ${lockStatus.type} workflow is already running. Please wait for it to complete or check the status.`);
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
  
  // Store processing status
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
    'totalMessages': '0'
  });

  // Initialize chunked processing with 1-day chunks
  const dateRange = calculateDateRange(timeRange);
  const start = new Date(dateRange.start);
  const end = new Date(dateRange.end);

  // Compute 1-day chunks count
  const oneDayMs = 24 * 60 * 60 * 1000;
  const totalChunks = Math.max(1, Math.ceil((end.getTime() - start.getTime()) / oneDayMs));

  properties.setProperties({
    'chunkCurrentStart': start.toISOString(),
    'chunkEnd': end.toISOString(),
    'chunkIndex': '0',
    'chunkTotalChunks': String(totalChunks),
    'accumulatedResults': JSON.stringify({ mustDo: [], mustKnow: [], totalProcessed: 0, batchesProcessed: 0 })
  });

  properties.setProperty('processingMessage', `Processing emails in ${totalChunks} 1-day chunks...`);

  // Kick off first chunk
  ScriptApp.newTrigger('processEmailsChunkedStep')
    .timeBased()
    .after(1000)
    .create();
}

/**
 * Unified chunked background processing - processes one day per invocation and chains next trigger
 */
function processEmailsChunkedStep() {
  const properties = PropertiesService.getUserProperties();
  try {
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

    const oneDayMs = 24 * 60 * 60 * 1000;
    const overallStart = new Date(chunkStartIso);
    const overallEnd = new Date(chunkEndIso);

    // Compute this day's window
    const currentStart = new Date(overallStart.getTime() + (chunkIndex * oneDayMs));
    const currentStop = new Date(Math.min(currentStart.getTime() + oneDayMs, overallEnd.getTime()));

    // If we've reached or passed the end, finalize
    if (currentStart.getTime() >= overallEnd.getTime()) {
      finalizeChunkedProcessing(accumulated);
      return;
    }

    properties.setProperty('processingMessage', `Processing day ${chunkIndex + 1} of ${totalChunks} (${currentStart.toDateString()})...`);

    // Set chunk start time for timeout detection
    properties.setProperty('chunkStartTime', new Date().toISOString());

    // Process the day in smaller sub-chunks (6 hours) to stay within time limits
    const subChunkResults = processDayInSubChunks(currentStart, currentStop, config, properties);
    
    // Accumulate results from all sub-chunks
    accumulated.mustDo = (accumulated.mustDo || []).concat(subChunkResults.mustDo || []);
    accumulated.mustKnow = (accumulated.mustKnow || []).concat(subChunkResults.mustKnow || []);
    accumulated.totalProcessed = (accumulated.totalProcessed || 0) + (subChunkResults.totalProcessed || 0);
    accumulated.batchesProcessed = (accumulated.batchesProcessed || 0) + (subChunkResults.batchesProcessed || 0);

    // Advance to next day
    const nextChunkIndex = chunkIndex + 1;
    properties.setProperties({
      'accumulatedResults': JSON.stringify(accumulated),
      'chunkIndex': String(nextChunkIndex)
    });

    // If finished all chunks, finalize and send summary
    if (overallStart.getTime() + (nextChunkIndex * oneDayMs) >= overallEnd.getTime() || nextChunkIndex >= totalChunks) {
      finalizeChunkedProcessing(accumulated);
      return;
    }

    // Chain next trigger for the following day chunk (1 hour delay for Google Apps Script limitation)
    ScriptApp.newTrigger('processEmailsChunkedStep')
      .timeBased()
      .after(60 * 60 * 1000) // 1 hour delay
      .create();

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
  }
}

/**
 * Process a single day in smaller sub-chunks (6 hours) to stay within time limits
 */
function processDayInSubChunks(dayStart, dayEnd, config, properties) {
  const sixHoursMs = 6 * 60 * 60 * 1000;
  const accumulated = { mustDo: [], mustKnow: [], totalProcessed: 0, batchesProcessed: 0 };
  
  let currentSubStart = new Date(dayStart);
  
  while (currentSubStart.getTime() < dayEnd.getTime()) {
    const currentSubEnd = new Date(Math.min(currentSubStart.getTime() + sixHoursMs, dayEnd.getTime()));
    
    // Update progress message for sub-chunk
    const subChunkStartTime = currentSubStart.toLocaleString();
    const subChunkEndTime = currentSubEnd.toLocaleString();
    properties.setProperty('processingMessage', `Processing ${subChunkStartTime} to ${subChunkEndTime}...`);
    
    // Fetch and process threads for this 6-hour window
    const dateRange = { start: currentSubStart, end: currentSubEnd };
    const emailThreads = fetchEmailThreadsFromGmail(dateRange);
    
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
      
      const results = processEmailsInBatches(emailThreads, config);
      
      // Accumulate results from this sub-chunk
      accumulated.mustDo = accumulated.mustDo.concat(results.mustDo || []);
      accumulated.mustKnow = accumulated.mustKnow.concat(results.mustKnow || []);
      accumulated.totalProcessed += (results.totalProcessed || 0);
      accumulated.batchesProcessed += (results.batchesProcessed || 0);
    }
    
    // Move to next 6-hour sub-chunk
    currentSubStart = new Date(currentSubStart.getTime() + sixHoursMs);
  }
  
  return accumulated;
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
      message: `Processed ${totalProcessed} emails across chunked day-by-day processing.`,
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
  const body = `Email processing timed out after 10 minutes.\n\nThis usually happens when processing a large number of emails. Please try:\n\n1. Using a shorter time range (e.g., 6 hours instead of 7 days)\n2. Reducing the number of topics to focus on\n3. Checking your internet connection\n\nYou can check the status in the Gmail Fast Reader add-on.`;
  
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
    const startTimeStr = properties.getProperty('processingStartTime');
    const message = properties.getProperty('processingMessage') || 'Processing...';
    const progress = properties.getProperty('processingProgress') || '0';
    const processedThreads = parseInt(properties.getProperty('processedThreads') || '0');
    const totalThreads = parseInt(properties.getProperty('totalThreads') || '0');
    const processedMessages = parseInt(properties.getProperty('processedMessages') || '0');
    const totalMessages = parseInt(properties.getProperty('totalMessages') || '0');
    
    // Check for chunk timeout (not overall processing timeout)
    const chunkStartTimeStr = properties.getProperty('chunkStartTime');
    if (chunkStartTimeStr) {
      const chunkStartTime = new Date(chunkStartTimeStr);
      const currentTime = new Date();
      const timeDiff = currentTime.getTime() - chunkStartTime.getTime();
      
      if (timeDiff > PROCESSING_TIMEOUT_MS) {
        // Current chunk has timed out - save latest stats before clearing
        saveLatestRunStats(properties, PROCESSING_STATUS.TIMEOUT, 'Processing timed out after 10 minutes. Please try again with a smaller time range.');
        
        // Clear processing status and chunk timing
        cleanupProcessingState(properties);
        cleanupChunkTiming(properties);
        
        // Send timeout notification email
        try {
          sendProcessingTimeoutEmail();
        } catch (emailError) {
          console.error('Failed to send timeout email:', emailError);
        }
        
        // Show latest run statistics
        return buildLatestRunStatsCard();
      }
    }
    
    // Build detailed progress message
    let progressMessage = message;
    
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
          .setText('🔄 Start New Scan')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('buildActiveWorkflowCard')))
        .addWidget(CardService.newTextButton()
          .setText('⚙️ Configuration')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('buildConfigCard')))
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

function isPassiveWorkflowScheduled() {
  return ScriptApp.getProjectTriggers()
    .some(t => t.getHandlerFunction() === 'runPassiveWorkflow');
}

function enablePassiveWorkflowSchedule() {
  if (!isPassiveWorkflowScheduled()) {
    ScriptApp.newTrigger('runPassiveWorkflow').timeBased().everyHours(1).create();
  }
}

function disablePassiveWorkflowSchedule() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'runPassiveWorkflow')
    .forEach(t => ScriptApp.deleteTrigger(t));
}

// Handlers for UI (called from toggle in both config & scan cards):
function handleAutomationToggle(e) {
  const checked = e.formInputs && e.formInputs.automaticScheduleCheckbox && e.formInputs.automaticScheduleCheckbox.includes('true');
  if (checked) {
    enablePassiveWorkflowSchedule();
  } else {
    disablePassiveWorkflowSchedule();
  }
  return buildConfigSuccessCard();
}
