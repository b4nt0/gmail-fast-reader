/**
 * Gmail Fast Reader - Main entry points and add-on lifecycle functions
 */

/**
 * Check if processing is currently running
 */
function isProcessingRunning() {
  const properties = PropertiesService.getUserProperties();
  const status = properties.getProperty('processingStatus');
  
  if (status !== PROCESSING_STATUS.RUNNING) {
    return false;
  }
  
  // Check for timeout
  const startTimeStr = properties.getProperty('processingStartTime');
  if (startTimeStr) {
    const startTime = new Date(startTimeStr);
    const currentTime = new Date();
    const timeDiff = currentTime.getTime() - startTime.getTime();
    
    if (timeDiff > PROCESSING_TIMEOUT_MS) {
      // Processing has timed out
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
      mustKnowOther: getFormBoolean(formInputs.mustKnowOther)
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
    // Check if processing is already running
    if (isProcessingRunning()) {
      return buildErrorCard('Another email scanning process is already running. Please wait for it to complete or check the status.');
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
 * Start background email processing
 */
function startBackgroundEmailProcessing(timeRange) {
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
  
  // Create a time-based trigger to run the processing
  ScriptApp.newTrigger('processEmailsInBackground')
    .timeBased()
    .after(1000) // Run after 1 second
    .create();
}

/**
 * Background email processing function
 */
function processEmailsInBackground() {
  const properties = PropertiesService.getUserProperties();
  
  try {
    // Update status
    properties.setProperty('processingMessage', 'Fetching emails from Gmail...');
    
    // Get configuration
    const config = getConfiguration();
    if (!config.openaiApiKey) {
      throw new Error('OpenAI API key not configured');
    }
    
    // Calculate date range
    const timeRange = properties.getProperty('processingTimeRange') || '1day';
    const dateRange = calculateDateRange(timeRange);
    
    // Fetch email threads from Gmail
    properties.setProperty('processingMessage', 'Fetching emails from Gmail...');
    const emailThreads = fetchEmailThreadsFromGmail(dateRange);
    
    if (emailThreads.length === 0) {
      properties.setProperties({
        'processingStatus': PROCESSING_STATUS.COMPLETED,
        'processingMessage': 'No emails found in the selected time range.',
        'processingResults': JSON.stringify({
          mustDo: [],
          mustKnow: [],
          totalProcessed: 0,
          message: 'No emails found in the selected time range.'
        })
      });
      sendProcessingCompleteEmail();
      return;
    }
    
    // Initialize progress tracking with actual counts
    const totalThreads = emailThreads.length;
    const totalMessages = emailThreads.reduce((total, thread) => total + thread.emails.length, 0);
    properties.setProperties({
      'totalThreads': totalThreads.toString(),
      'totalMessages': totalMessages.toString(),
      'processedThreads': '0',
      'processedMessages': '0',
      'processingMessage': `Found ${totalThreads} threads with ${totalMessages} messages. Starting analysis...`
    });
    
    // Process emails in batches
    const results = processEmailsInBatches(emailThreads, config);
    
    // Store results
    properties.setProperties({
      'processingStatus': PROCESSING_STATUS.COMPLETED,
      'processingMessage': `Processing complete. Found ${results.mustDo.length} actionable items and ${results.mustKnow.length} informational items.`,
      'processingResults': JSON.stringify(results)
    });
    
    // Send completion email
    sendProcessingCompleteEmail();
    
  } catch (error) {
    console.error('Error in background processing:', error);
    properties.setProperties({
      'processingStatus': PROCESSING_STATUS.ERROR,
      'processingMessage': 'Processing failed: ' + error.message
    });
    sendProcessingErrorEmail(error.message);
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
  const subject = `${config.addonName} - Processing Complete - ${new Date().toLocaleDateString()}`;
  
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
    return buildMainCard();
  }
  
  if (status === PROCESSING_STATUS.RUNNING) {
    const startTimeStr = properties.getProperty('processingStartTime');
    const message = properties.getProperty('processingMessage') || 'Processing...';
    const progress = properties.getProperty('processingProgress') || '0';
    const processedThreads = parseInt(properties.getProperty('processedThreads') || '0');
    const totalThreads = parseInt(properties.getProperty('totalThreads') || '0');
    const processedMessages = parseInt(properties.getProperty('processedMessages') || '0');
    const totalMessages = parseInt(properties.getProperty('totalMessages') || '0');
    
    // Check for timeout
    if (startTimeStr) {
      const startTime = new Date(startTimeStr);
      const currentTime = new Date();
      const timeDiff = currentTime.getTime() - startTime.getTime();
      
      if (timeDiff > PROCESSING_TIMEOUT_MS) {
        // Processing has timed out
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
        
        return buildErrorCard('Processing timed out after 10 minutes. Please try again with a smaller time range.');
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
    // Clear processing status
    properties.deleteProperty('processingStatus');
    properties.deleteProperty('processingMessage');
    properties.deleteProperty('processingResults');
    properties.deleteProperty('processingStartTime');
    results.message = 'Processing complete';
    return buildSummaryCard(results);
  }
  
  if (status === PROCESSING_STATUS.ERROR) {
    const message = properties.getProperty('processingMessage') || 'Processing failed';
    // Clear processing status
    properties.deleteProperty('processingStatus');
    properties.deleteProperty('processingMessage');
    properties.deleteProperty('processingStartTime');
    return buildErrorCard(message);
  }
  
  if (status === PROCESSING_STATUS.TIMEOUT) {
    const message = properties.getProperty('processingMessage') || 'Processing timed out';
    // Clear processing status
    properties.deleteProperty('processingStatus');
    properties.deleteProperty('processingMessage');
    properties.deleteProperty('processingStartTime');
    return buildErrorCard(message);
  }
  
  return buildMainCard();
}
