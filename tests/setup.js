/**
 * Jest setup file for Google Apps Script mocking
 */

// Mock Google Apps Script services
global.CardService = {
  newCardBuilder: jest.fn(() => ({
    setHeader: jest.fn().mockReturnThis(),
    addSection: jest.fn().mockReturnThis(),
    build: jest.fn(() => ({ type: 'card' }))
  })),
  newCardHeader: jest.fn(() => ({
    setTitle: jest.fn().mockReturnThis(),
    setSubtitle: jest.fn().mockReturnThis(),
    setImageUrl: jest.fn().mockReturnThis(),
    setImageStyle: jest.fn().mockReturnThis()
  })),
  newCardSection: jest.fn(() => ({
    setHeader: jest.fn().mockReturnThis(),
    addWidget: jest.fn().mockReturnThis()
  })),
  newTextParagraph: jest.fn(() => ({
    setText: jest.fn().mockReturnThis()
  })),
  newTextButton: jest.fn(() => ({
    setText: jest.fn().mockReturnThis(),
    setOnClickAction: jest.fn().mockReturnThis(),
    setDisabled: jest.fn().mockReturnThis()
  })),
  newButtonSet: jest.fn(() => ({
    addButton: jest.fn().mockReturnThis()
  })),
  newAction: jest.fn(() => ({
    setFunctionName: jest.fn().mockReturnThis(),
    setParameter: jest.fn().mockReturnThis()
  })),
  newTextInput: jest.fn(() => ({
    setFieldName: jest.fn().mockReturnThis(),
    setTitle: jest.fn().mockReturnThis(),
    setValue: jest.fn().mockReturnThis(),
    setHint: jest.fn().mockReturnThis(),
    setMultiline: jest.fn().mockReturnThis(),
    setSuggestionsAction: jest.fn().mockReturnThis()
  })),
  newSelectionInput: jest.fn(() => ({
    setType: jest.fn().mockReturnThis(),
    setTitle: jest.fn().mockReturnThis(),
    setFieldName: jest.fn().mockReturnThis(),
    addItem: jest.fn().mockReturnThis()
  })),
  ImageStyle: {
    CIRCLE: 'CIRCLE'
  },
  SelectionInputType: {
    DROPDOWN: 'DROPDOWN',
    CHECK_BOX: 'CHECK_BOX',
    RADIO_BUTTON: 'RADIO_BUTTON'
  }
};

global.PropertiesService = {
  getUserProperties: jest.fn(() => ({
    getProperty: jest.fn(),
    setProperty: jest.fn(),
    setProperties: jest.fn(),
    deleteProperty: jest.fn()
  })),
  getScriptProperties: jest.fn(() => ({
    getProperty: jest.fn(),
    setProperty: jest.fn()
  }))
};

global.GmailApp = {
  search: jest.fn(() => []),
  sendEmail: jest.fn(),
  getUserLabels: jest.fn(() => [])
};

global.ScriptApp = {
  newTrigger: jest.fn(() => ({
    timeBased: jest.fn(() => ({
      everyHours: jest.fn(() => ({
        create: jest.fn()
      })),
      after: jest.fn(() => ({
        create: jest.fn()
      }))
    })),
    eventBased: jest.fn(() => ({
      create: jest.fn()
    }))
  })),
  getProjectTriggers: jest.fn(() => []),
  deleteTrigger: jest.fn()
};

global.Session = {
  getActiveUser: jest.fn(() => ({
    getEmail: jest.fn(() => 'test@example.com')
  }))
};

global.UrlFetchApp = {
  fetch: jest.fn(() => ({
    getContentText: jest.fn(() => '{}'),
    getResponseCode: jest.fn(() => 200)
  }))
};

global.console = {
  log: jest.fn(),
  error: jest.fn(),
  warn: jest.fn(),
  info: jest.fn()
};

global.Logger = {
  log: jest.fn(),
  logData: jest.fn()
};

// Mock global functions that are defined in other files
// These will be overridden by actual implementations in tests
global.getConfiguration = jest.fn();
global.isConfigurationComplete = jest.fn();
global.isProcessingRunning = jest.fn();
global.getUserEmailAddress = jest.fn(() => 'test@example.com');
global.buildConfigurationCard = jest.fn();
global.buildActiveWorkflowCard = jest.fn();
global.checkProcessingStatus = jest.fn();

