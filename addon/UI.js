/**
 * Gmail Fast Reader - Card UI builders
 */

/**
 * Build main homepage card
 */
function buildMainCard() {
  const config = getConfiguration();
  const isConfigured = isConfigurationComplete();
  const isRunning = isProcessingRunning();
    
  // Check if there are previous run results available
  const properties = PropertiesService.getUserProperties();
  const hasPreviousResults = properties.getProperty('latestRunStats') !== null;
  const hasProcessStatus = properties.getProperty('processingStatus') !== null;
    
  // Create card header
  const header = CardService.newCardHeader()
    .setTitle('Gmail Fast Reader')
    .setSubtitle(config.addonName);
    
  // Create welcome message section
  const welcomeSection = CardService.newCardSection()
    .addWidget(CardService.newTextParagraph()
      .setText('Welcome to Gmail Fast Reader! Configure your topics and scan your emails for important items.'));
    
  // Create configure button
  const configureButton = CardService.newTextButton()
    .setText('⚙️ Configure')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('buildConfigurationCard'));
    
  // Create configure button section
  const configureSection = CardService.newCardSection()
    .addWidget(CardService.newButtonSet()
      .addButton(configureButton));
    
  // Build the card
  const cardBuilder = CardService.newCardBuilder()
    .setHeader(header)
    .addSection(welcomeSection)
    .addSection(configureSection);
    
  // Add passive workflow confirmation message if configured
  if (isConfigured) {
    const passiveWorkflowSection = CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText('✅ Emails are being scanned automatically on a regular basis. You don\'t need to do anything - just wait for your daily summary!'));
    cardBuilder.addSection(passiveWorkflowSection);
  }
    
  // Always show status check button
  const statusButton = CardService.newTextButton()
    .setText(isRunning ? '🔄 Check Status (Running)' : '🔄 Check Status')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('checkProcessingStatus'));
    
  const statusSection = CardService.newCardSection()
    .addWidget(CardService.newButtonSet()
      .addButton(statusButton));
    
  cardBuilder.addSection(statusSection);
    
  // Add warning section if not configured
  if (!isConfigured) {
    const warningSection = CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText('⚠️ Please configure your topics and OpenAI API key first.'));
    cardBuilder.addSection(warningSection);
  }
    
  // Add running status message if processing is active
  if (isRunning) {
    const properties = PropertiesService.getUserProperties();
    const progress = properties.getProperty('processingProgress') || '0';
    const processedThreads = parseInt(properties.getProperty('processedThreads') || '0');
    const totalThreads = parseInt(properties.getProperty('totalThreads') || '0');
    const processedMessages = parseInt(properties.getProperty('processedMessages') || '0');
    const totalMessages = parseInt(properties.getProperty('totalMessages') || '0');
    const currentActivity = properties.getProperty('processingMessage') || 'Processing...';
      
    let statusMessage = `🔄 Email processing is currently running (${progress}% complete).\n\nCurrent: ${currentActivity}`;
      
    if (totalThreads > 0) {
      statusMessage += `\n📧 Threads: ${processedThreads}/${totalThreads}`;
    }
      
    if (totalMessages > 0) {
      statusMessage += `\n💬 Messages: ${processedMessages}/${totalMessages}`;
    }
      
    statusMessage += '\n\nClick "Check Status" to monitor detailed progress.';
      
    const runningSection = CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText(statusMessage));
    cardBuilder.addSection(runningSection);
  }
    
  return cardBuilder.build();
}
  
/**
   * Build system settings card
   */
function buildSystemSettingsCard() {
  const config = getConfiguration();
  const timezoneOptions = getTimezoneOptions();
    
  const timezoneSelection = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setTitle('Timezone')
    .setFieldName('timeZone');
    
  timezoneOptions.forEach(option => {
    timezoneSelection.addItem(option.label, option.value, option.value === config.timeZone);
  });
    
  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('System Settings'))
    .addSection(CardService.newCardSection()
      .setHeader('System Configuration')
      .addWidget(CardService.newTextInput()
        .setFieldName('addonName')
        .setTitle('Add-on Name')
        .setValue(config.addonName)
        .setHint('A friendly name for your add-on instance'))
      .addWidget(CardService.newTextInput()
        .setFieldName('openaiApiKey')
        .setTitle('OpenAI API Key')
        .setValue(config.openaiApiKey)
        .setHint('Your OpenAI API key for email analysis'))
      .addWidget(timezoneSelection))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText('💾 Save System Settings')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('handleSystemSettingsSubmit')))
        .addButton(CardService.newTextButton()
          .setText('⬅️ Back to Configuration')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('buildConfigurationCard')))));
    
  return card.build();
}
  
/**
   * Build email settings card
   */
function buildEmailSettingsCard() {
  const config = getConfiguration();
    
  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Email Settings'))
    .addSection(CardService.newCardSection()
      .setHeader('Email Filtering')
      .addWidget(CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.CHECK_BOX)
        .setTitle('Filter Options')
        .setFieldName('unreadOnly')
        .addItem('Unread only', 'true', config.unreadOnly))
      .addWidget(CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.CHECK_BOX)
        .setTitle('')
        .setFieldName('inboxOnly')
        .addItem('Inbox only', 'true', config.inboxOnly)))
    .addSection(CardService.newCardSection()
      .setHeader('Processing Options')
      .addWidget(CardService.newTextInput()
        .setTitle('Label for "I must do" emails (optional)')
        .setFieldName('mustDoLabel')
        .setValue(config.mustDoLabel || '')
        .setSuggestionsAction(CardService.newAction().setFunctionName('handleLabelSuggestions')))
      .addWidget(CardService.newTextInput()
        .setTitle('Label for "I must know" emails (optional)')
        .setFieldName('mustKnowLabel')
        .setValue(config.mustKnowLabel || '')
        .setSuggestionsAction(CardService.newAction().setFunctionName('handleLabelSuggestions'))))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText('💾 Save Email Settings')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('handleEmailSettingsSubmit')))
        .addButton(CardService.newTextButton()
          .setText('⬅️ Back to Configuration')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('buildConfigurationCard')))));
    
  return card.build();
}
  
/**
   * Build onboarding card (wizard-style)
   * @param {number} step - Current step (1 = System Settings, 2 = Topics, 3 = Email Settings)
   */
function buildOnboardingCard(step = 1) {
  const config = getConfiguration();
  const timezoneOptions = getTimezoneOptions();
  const totalSteps = 3;
    
  // Step titles
  const stepTitles = {
    1: 'System Settings',
    2: 'Important Settings',
    3: 'Email Settings'
  };
    
  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle(`Onboarding - Step ${step} of ${totalSteps}: ${stepTitles[step]}`));
    
  // Step 1: System Settings
  if (step === 1) {
    const timezoneSelection = CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.DROPDOWN)
      .setTitle('Timezone')
      .setFieldName('timeZone');
      
    timezoneOptions.forEach(option => {
      timezoneSelection.addItem(option.label, option.value, option.value === config.timeZone);
    });
      
    card.addSection(CardService.newCardSection()
      .setHeader('System Configuration')
      .addWidget(CardService.newTextInput()
        .setFieldName('addonName')
        .setTitle('Add-on Name')
        .setValue(config.addonName)
        .setHint('A friendly name for your add-on instance'))
      .addWidget(CardService.newTextInput()
        .setFieldName('openaiApiKey')
        .setTitle('OpenAI API Key')
        .setValue(config.openaiApiKey)
        .setHint('Your OpenAI API key for email analysis'))
      .addWidget(timezoneSelection));
  }
    
  // Step 2: Topics
  if (step === 2) {
    card.addSection(CardService.newCardSection()
      .setHeader('I Must Do Topics')
      .addWidget(CardService.newTextInput()
        .setFieldName('mustDoTopics')
        .setTitle('Topics of Interest')
        .setValue(config.mustDoTopics)
        .setHint('Enter topics separated by new lines (e.g., payments, deadlines, meetings)')
        .setMultiline(true))
      .addWidget(CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.CHECK_BOX)
        .setTitle('Additional Options')
        .setFieldName('mustDoOther')
        .addItem('Let AI decide on other relevant topics', 'true', config.mustDoOther)));
      
    card.addSection(CardService.newCardSection()
      .setHeader('I Must Know Topics')
      .addWidget(CardService.newTextInput()
        .setFieldName('mustKnowTopics')
        .setTitle('Topics of Interest')
        .setValue(config.mustKnowTopics)
        .setHint('Enter topics separated by new lines (e.g., school updates, news, announcements)')
        .setMultiline(true))
      .addWidget(CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.CHECK_BOX)
        .setTitle('Additional Options')
        .setFieldName('mustKnowOther')
        .addItem('Let AI decide on other relevant topics', 'true', config.mustKnowOther)));
  }
    
  // Step 3: Email Settings
  if (step === 3) {
    card.addSection(CardService.newCardSection()
      .setHeader('Email Filtering')
      .addWidget(CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.CHECK_BOX)
        .setTitle('Filter Options')
        .setFieldName('unreadOnly')
        .addItem('Unread only', 'true', config.unreadOnly))
      .addWidget(CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.CHECK_BOX)
        .setTitle('')
        .setFieldName('inboxOnly')
        .addItem('Inbox only', 'true', config.inboxOnly)));
      
    card.addSection(CardService.newCardSection()
      .setHeader('Processing Options')
      .addWidget(CardService.newTextInput()
        .setTitle('Label for "I must do" emails (optional)')
        .setFieldName('mustDoLabel')
        .setValue(config.mustDoLabel || '')
        .setSuggestionsAction(CardService.newAction().setFunctionName('handleLabelSuggestions')))
      .addWidget(CardService.newTextInput()
        .setTitle('Label for "I must know" emails (optional)')
        .setFieldName('mustKnowLabel')
        .setValue(config.mustKnowLabel || '')
        .setSuggestionsAction(CardService.newAction().setFunctionName('handleLabelSuggestions'))));
  }
    
  // Navigation buttons
  const buttonSet = CardService.newButtonSet();
    
  // Previous button (not on step 1)
  if (step > 1) {
    buttonSet.addButton(CardService.newTextButton()
      .setText('⬅️ Previous')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('handleOnboardingNavigation')
        .setParameters({ step: String(step - 1) })));
  }
    
  // Save and Next button
  if (step < totalSteps) {
    buttonSet.addButton(CardService.newTextButton()
      .setText('💾 Save & Next ➡️')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('handleOnboardingSaveAndNext')
        .setParameters({ step: String(step), nextStep: String(step + 1) })));
  } else {
    // Last step - Finish button
    buttonSet.addButton(CardService.newTextButton()
      .setText('✅ Finish Setup')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('handleOnboardingFinish')));
  }
    
  card.addSection(CardService.newCardSection()
    .addWidget(buttonSet));
    
  return card.build();
}
  
/**
   * Build configuration card
   */
function buildConfigurationCard() {
  // Check if onboarding is needed first
  if (needsOnboarding()) {
    return buildOnboardingCard(1);
  }
    
  const config = getConfiguration();
  const isRunning = isProcessingRunning();
    
  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Configuration'))
    .addSection(CardService.newCardSection()
      .setHeader('I Must Do Topics')
      .addWidget(CardService.newTextInput()
        .setFieldName('mustDoTopics')
        .setTitle('Topics of Interest')
        .setValue(config.mustDoTopics)
        .setHint('Enter topics separated by new lines (e.g., payments, deadlines, meetings)')
        .setMultiline(true))
      .addWidget(CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.CHECK_BOX)
        .setTitle('Additional Options')
        .setFieldName('mustDoOther')
        .addItem('Let AI decide on other relevant topics', 'true', config.mustDoOther)))
    .addSection(CardService.newCardSection()
      .setHeader('I Must Know Topics')
      .addWidget(CardService.newTextInput()
        .setFieldName('mustKnowTopics')
        .setTitle('Topics of Interest')
        .setValue(config.mustKnowTopics)
        .setHint('Enter topics separated by new lines (e.g., school updates, news, announcements)')
        .setMultiline(true))
      .addWidget(CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.CHECK_BOX)
        .setTitle('Additional Options')
        .setFieldName('mustKnowOther')
        .addItem('Let AI decide on other relevant topics', 'true', config.mustKnowOther)))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText('💾 Save Topics')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('handleTopicsSubmit')))
        .addButton(CardService.newTextButton()
          .setText('📧 Email Settings')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('buildEmailSettingsCard')))
        .addButton(CardService.newTextButton()
          .setText('⚙️ System Settings')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('buildSystemSettingsCard')))))
    .addSection(CardService.newCardSection()
      .setHeader('Scan Emails')
      .addWidget(CardService.newTextParagraph()
        .setText('Manually trigger an email scan for a specific time range. Your emails are also being scanned automatically in the background.'))
      .addWidget(CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText('📧 Scan Emails')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('buildActiveWorkflowCard'))
          .setDisabled(isRunning))));
    
  // Add debug button only for debug user
  const userEmail = getUserEmailAddress();
  if (userEmail === DEBUG_USER_EMAIL) {
    card.addSection(CardService.newCardSection()
      .setHeader('Debug')
      .addWidget(CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText('🗑️ Nuke All Settings')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('handleNukeSettings')))));
  }
    
  return card.build();
}
  
/**
   * Build active workflow card
   */
function buildActiveWorkflowCard(timeRange = null) {
  const timeRanges = [
    { label: 'Last 6 hours', value: '6hours' },
    { label: 'Last 12 hours', value: '12hours' },
    { label: 'Last day', value: '1day' },
    { label: 'Last 2 days', value: '2days' },
    { label: 'Last week', value: '7days' }
  ];
    
  const timeRangeSelection = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.RADIO_BUTTON)
    .setTitle('Time Range')
    .setFieldName('timeRange');
    
  timeRanges.forEach(range => {
    timeRangeSelection.addItem(range.label, range.value, range.value === timeRange);
  });
    
  const isRunning = isProcessingRunning();
  // Automation scheduling is always on via dispatcher; no UI toggle
    
  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Scan Emails'))
    .addSection(CardService.newCardSection()
      .setHeader('Select Time Range')
      .addWidget(timeRangeSelection));
    
  // Add running status message if processing is active
  if (isRunning) {
    card.addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText('⚠️ Another email scanning process is already running. Please wait for it to complete before starting a new scan.')));
  }
    
  // No automation section; passive runs hourly automatically
    
  card.addSection(CardService.newCardSection()
    .addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText('🔍 Scan Emails')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('handleScanEmails'))
        .setDisabled(isRunning))
      .addButton(CardService.newTextButton()
        .setText('🔄 Check Status')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('checkProcessingStatus')))));
    
  return card.build();
}
  
/**
   * Build progress card
   */
function buildProgressCard(message) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Processing...'))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText(message)))
    .build();
}
  
/**
   * Build progress card with Check Status button
   */
function buildProgressCardWithStatusButton(message) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Processing...'))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText(message)))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText('🔄 Check Status')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('checkProcessingStatus')))))
    .build();
}
  
/**
   * Build progress card with refresh functionality
   */
function buildProgressCardWithAutoRefresh(message) {
  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Processing...'));
    
  // Split message into lines for better formatting
  const messageLines = message.split('\n');
    
  // Add main status message
  card.addSection(CardService.newCardSection()
    .addWidget(CardService.newTextParagraph()
      .setText(messageLines[0] || 'Processing...')));
    
  // Add progress details if available
  if (messageLines.length > 1) {
    const progressDetails = messageLines.slice(1).join('\n');
    card.addSection(CardService.newCardSection()
      .setHeader('Progress Details')
      .addWidget(CardService.newTextParagraph()
        .setText(progressDetails)));
  }
    
  // Add refresh notification
  card.addSection(CardService.newCardSection()
    .addWidget(CardService.newTextParagraph()
      .setText('🔄 Click "Refresh Status" below to update the progress information.')));
    
  // Add refresh button
  card.addSection(CardService.newCardSection()
    .addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText('🔄 Refresh Status')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('checkProcessingStatus')))
      .addButton(CardService.newTextButton()
        .setText('🧯 Emergency reset')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('handleEmergencyReset')))));
    
  return card.build();
}
  
/**
   * Build configuration success card
   */
function buildConfigSuccessCard() {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Configuration Saved'))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText('✅ Configuration saved successfully! You can now scan your emails.')))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText('🏠 Back to Main')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('buildMainCard')))))
    .build();
}
  
/**
   * Build error card
   */
function buildErrorCard(errorMessage) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Error'))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText('❌ ' + errorMessage)))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText('🏠 Back to Main')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('buildMainCard')))))
    .build();
}
  
/**
   * Build quick scan card for single email (future use)
   */
function buildQuickScanCard(e) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Quick Scan'))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText('Quick scan functionality coming soon!')))
    .build();
}