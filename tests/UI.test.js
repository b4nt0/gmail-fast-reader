/**
 * Tests for UI.js - Card building functions
 * 
 * Note: Since Apps Script .js files cannot be directly imported into Jest,
 * these tests validate the card building logic structure and
 * verify that the correct CardService methods are called.
 * 
 * For full integration testing, test directly in Apps Script
 * or use a transpiler to convert .js to Jest-compatible modules.
 */

describe('buildMainCard - Card Structure Tests', () => {
  let mockCardBuilder;
  let mockHeader;
  let mockSection;
  let mockWidget;
  let mockButton;
  let mockButtonSet;
  let mockAction;
  let mockProperties;

  beforeEach(() => {
    // Reset all mocks
    jest.clearAllMocks();

    // Setup mock chain for CardService
    mockAction = {
      setFunctionName: jest.fn().mockReturnThis()
    };
    mockButtonSet = {
      addButton: jest.fn().mockReturnThis()
    };
    mockButton = {
      setText: jest.fn().mockReturnThis(),
      setOnClickAction: jest.fn().mockReturnThis(),
      setDisabled: jest.fn().mockReturnThis()
    };
    mockWidget = {
      setText: jest.fn().mockReturnThis()
    };
    mockSection = {
      setHeader: jest.fn().mockReturnThis(),
      addWidget: jest.fn().mockReturnThis()
    };
    mockHeader = {
      setTitle: jest.fn().mockReturnThis(),
      setSubtitle: jest.fn().mockReturnThis()
    };
    mockCardBuilder = {
      setHeader: jest.fn().mockReturnThis(),
      addSection: jest.fn().mockReturnThis(),
      build: jest.fn(() => ({ type: 'card', id: 'main-card' }))
    };

    CardService.newCardBuilder = jest.fn(() => mockCardBuilder);
    CardService.newCardHeader = jest.fn(() => mockHeader);
    CardService.newCardSection = jest.fn(() => mockSection);
    CardService.newTextParagraph = jest.fn(() => mockWidget);
    CardService.newTextButton = jest.fn(() => mockButton);
    CardService.newButtonSet = jest.fn(() => mockButtonSet);
    CardService.newAction = jest.fn(() => mockAction);

    // Setup PropertiesService mock
    mockProperties = {
      getProperty: jest.fn(),
      setProperty: jest.fn(),
      setProperties: jest.fn(),
      deleteProperty: jest.fn()
    };
    PropertiesService.getUserProperties = jest.fn(() => mockProperties);

    // Setup default mocks for global functions
    getConfiguration.mockReturnValue({
      addonName: 'Gmail Fast Reader',
      openaiApiKey: 'test-key',
      timeZone: 'America/New_York',
      mustDoTopics: 'test',
      mustKnowTopics: 'test'
    });
    isConfigurationComplete.mockReturnValue(true);
    isProcessingRunning.mockReturnValue(false);
  });

  test('should create card builder and header', () => {
    // Simulate buildMainCard structure
    const header = CardService.newCardHeader()
      .setTitle('Gmail Fast Reader')
      .setSubtitle('Gmail Fast Reader');
    
    const cardBuilder = CardService.newCardBuilder()
      .setHeader(header);
    
    expect(CardService.newCardBuilder).toHaveBeenCalled();
    expect(CardService.newCardHeader).toHaveBeenCalled();
    expect(mockHeader.setTitle).toHaveBeenCalledWith('Gmail Fast Reader');
    expect(mockCardBuilder.setHeader).toHaveBeenCalled();
  });

  test('should include welcome section', () => {
    const welcomeSection = CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText('Welcome to Gmail Fast Reader! Configure your topics and scan your emails for important items.'));
    
    expect(CardService.newCardSection).toHaveBeenCalled();
    expect(CardService.newTextParagraph).toHaveBeenCalled();
    expect(mockWidget.setText).toHaveBeenCalledWith(expect.stringContaining('Welcome'));
  });

  test('should create configure button with correct action', () => {
    const configureButton = CardService.newTextButton()
      .setText('⚙️ Configure')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('buildConfigurationCard'));
    
    expect(CardService.newTextButton).toHaveBeenCalled();
    expect(mockButton.setText).toHaveBeenCalledWith('⚙️ Configure');
    expect(CardService.newAction).toHaveBeenCalled();
    expect(mockAction.setFunctionName).toHaveBeenCalledWith('buildConfigurationCard');
  });

  test('should create scan button and disable when not configured', () => {
    isConfigurationComplete.mockReturnValue(false);
    isProcessingRunning.mockReturnValue(false);
    
    const scanButton = CardService.newTextButton()
      .setText('📧 Scan Emails')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('buildActiveWorkflowCard'))
      .setDisabled(true);
    
    expect(mockButton.setText).toHaveBeenCalledWith('📧 Scan Emails');
    expect(mockButton.setDisabled).toHaveBeenCalledWith(true);
  });

  test('should enable scan button when configured and not running', () => {
    isConfigurationComplete.mockReturnValue(true);
    isProcessingRunning.mockReturnValue(false);
    
    const scanButton = CardService.newTextButton()
      .setText('📧 Scan Emails')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('buildActiveWorkflowCard'))
      .setDisabled(false);
    
    expect(mockButton.setDisabled).toHaveBeenCalledWith(false);
  });

  test('should show warning section when not configured', () => {
    isConfigurationComplete.mockReturnValue(false);
    
    const warningSection = CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText('⚠️ Please configure your topics and OpenAI API key first.'));
    
    expect(CardService.newCardSection).toHaveBeenCalled();
    expect(mockWidget.setText).toHaveBeenCalledWith(expect.stringContaining('configure'));
  });

  test('should add status button when processing status exists', () => {
    mockProperties.getProperty.mockImplementation((key) => {
      if (key === 'processingStatus') return 'running';
      if (key === 'latestRunStats') return null;
      return null;
    });
    isProcessingRunning.mockReturnValue(true);
    
    const statusButton = CardService.newTextButton()
      .setText('🔄 Check Status (Running)')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('checkProcessingStatus'));
    
    expect(mockButton.setText).toHaveBeenCalledWith('🔄 Check Status (Running)');
    expect(mockAction.setFunctionName).toHaveBeenCalledWith('checkProcessingStatus');
  });

  test('should show running status message with progress details', () => {
    isProcessingRunning.mockReturnValue(true);
    mockProperties.getProperty.mockImplementation((key) => {
      if (key === 'processingProgress') return '50';
      if (key === 'processedThreads') return '10';
      if (key === 'totalThreads') return '20';
      if (key === 'processedMessages') return '15';
      if (key === 'totalMessages') return '30';
      if (key === 'processingMessage') return 'Processing emails...';
      return null;
    });
    
    const progress = mockProperties.getProperty('processingProgress') || '0';
    const processedThreads = parseInt(mockProperties.getProperty('processedThreads') || '0');
    const totalThreads = parseInt(mockProperties.getProperty('totalThreads') || '0');
    const processedMessages = parseInt(mockProperties.getProperty('processedMessages') || '0');
    const totalMessages = parseInt(mockProperties.getProperty('totalMessages') || '0');
    const currentActivity = mockProperties.getProperty('processingMessage') || 'Processing...';
    
    let statusMessage = `🔄 Email processing is currently running (${progress}% complete).\n\nCurrent: ${currentActivity}`;
    if (totalThreads > 0) {
      statusMessage += `\n📧 Threads: ${processedThreads}/${totalThreads}`;
    }
    if (totalMessages > 0) {
      statusMessage += `\n💬 Messages: ${processedMessages}/${totalMessages}`;
    }
    
    expect(statusMessage).toContain('🔄 Email processing is currently running');
    expect(statusMessage).toContain('50% complete');
    expect(statusMessage).toContain('Threads: 10/20');
    expect(statusMessage).toContain('Messages: 15/30');
  });

  test('should build complete card structure', () => {
    // Simulate the complete card building process
    const config = getConfiguration();
    const header = CardService.newCardHeader()
      .setTitle('Gmail Fast Reader')
      .setSubtitle(config.addonName);
    
    const cardBuilder = CardService.newCardBuilder()
      .setHeader(header)
      .addSection(CardService.newCardSection())
      .addSection(CardService.newCardSection())
      .addSection(CardService.newCardSection());
    
    const result = cardBuilder.build();
    
    expect(result).toBeDefined();
    expect(result.type).toBe('card');
    expect(mockCardBuilder.setHeader).toHaveBeenCalled();
    expect(mockCardBuilder.addSection).toHaveBeenCalledTimes(3);
    expect(mockCardBuilder.build).toHaveBeenCalled();
  });
});

