/**
 * Gmail Fast Reader - Card UI builders
 */

/**
 * Build main homepage card
 */
function buildMainCard() {
  const config = getConfiguration();
  const isConfigured = isConfigurationComplete();
  
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
  
  // Create scan emails button
  const scanButton = CardService.newTextButton()
    .setText('📧 Scan Emails')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('buildActiveWorkflowCard'))
    .setDisabled(!isConfigured);
  
  // Create scan button section
  const scanSection = CardService.newCardSection()
    .addWidget(CardService.newButtonSet()
      .addButton(scanButton));
  
  // Create status check button
  const statusButton = CardService.newTextButton()
    .setText('🔄 Check Status')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('checkProcessingStatus'));
  
  const statusSection = CardService.newCardSection()
    .addWidget(CardService.newButtonSet()
      .addButton(statusButton));
  
  // Build the card
  const cardBuilder = CardService.newCardBuilder()
    .setHeader(header)
    .addSection(welcomeSection)
    .addSection(configureSection)
    .addSection(scanSection)
    .addSection(statusSection);
  
  // Add warning section if not configured
  if (!isConfigured) {
    const warningSection = CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText('⚠️ Please configure your topics and OpenAI API key first.'));
    cardBuilder.addSection(warningSection);
  }
  
  return cardBuilder.build();
}

/**
 * Build configuration card
 */
function buildConfigurationCard() {
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
      .setTitle('Configuration'))
    .addSection(CardService.newCardSection()
      .setHeader('Basic Settings')
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
          .setText('💾 Save Configuration')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('handleConfigSubmit')))));
  
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
  
  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Scan Emails'))
    .addSection(CardService.newCardSection()
      .setHeader('Select Time Range')
      .addWidget(timeRangeSelection))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText('🔍 Scan Emails')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('handleScanEmails')))));
  
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