/**
 * Regression tests for inboxOnly configuration setting
 * 
 * Tests verify that the inboxOnly config setting is properly applied
 * in both passive and active workflows.
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');

// Create a context for executing the Apps Script code
const scriptContext = vm.createContext({
  ...global,
  PropertiesService: null,
  ScriptApp: null,
  GmailApp: null,
  Session: null,
  console: global.console,
  Logger: global.Logger,
  Date: Date,
  JSON: JSON,
  Math: Math,
  parseInt: parseInt,
  parseFloat: parseFloat,
  String: String,
  Array: Array,
  Object: Object,
  Error: Error,
  TypeError: TypeError,
  ReferenceError: ReferenceError
});

// Define PROCESSING_STATUS constant (from Constants.js)
const PROCESSING_STATUS = {
  RUNNING: 'running',
  COMPLETED: 'completed',
  ERROR: 'error',
  TIMEOUT: 'timeout'
};

// Set PROCESSING_STATUS in script context
scriptContext.PROCESSING_STATUS = PROCESSING_STATUS;
scriptContext.PROCESSING_TIMEOUT_MS = 10 * 60 * 1000;

// Load Constants.js
const constantsJsPath = path.join(__dirname, '../addon/Constants.js');
const constantsCode = fs.readFileSync(constantsJsPath, 'utf8');
vm.runInContext(constantsCode, scriptContext);

// Load Config.js (needed for getConfiguration and saveConfiguration)
const configJsPath = path.join(__dirname, '../addon/Config.js');
const configCode = fs.readFileSync(configJsPath, 'utf8');
vm.runInContext(configCode, scriptContext);

// Load EmailProcessor.js (needed for fetchEmailThreadsFromGmail)
const emailProcessorJsPath = path.join(__dirname, '../addon/EmailProcessor.js');
const emailProcessorCode = fs.readFileSync(emailProcessorJsPath, 'utf8');
vm.runInContext(emailProcessorCode, scriptContext);

// Load Code.js (needed for fetchEmailThreadsForPassiveWorkflow and runPassiveWorkflow)
const codeJsPath = path.join(__dirname, '../addon/Code.js');
const codeJs = fs.readFileSync(codeJsPath, 'utf8');
vm.runInContext(codeJs, scriptContext);

describe('inboxOnly Configuration Regression Tests', () => {
  let mockPropertiesStore;
  let mockPropertiesService;
  let mockGmailApp;
  let searchCallQueries;

  beforeEach(() => {
    jest.clearAllMocks();
    searchCallQueries = [];

    // Create mock PropertiesService
    mockPropertiesStore = {};
    mockPropertiesService = {
      getProperty: jest.fn((key) => mockPropertiesStore[key] || null),
      setProperty: jest.fn((key, value) => {
        mockPropertiesStore[key] = value;
      }),
      setProperties: jest.fn((properties) => {
        Object.keys(properties).forEach(key => {
          mockPropertiesStore[key] = properties[key];
        });
      }),
      deleteProperty: jest.fn((key) => {
        delete mockPropertiesStore[key];
      }),
      getStore: () => mockPropertiesStore
    };

    // Create mock GmailApp that tracks search queries
    mockGmailApp = {
      search: jest.fn((query) => {
        searchCallQueries.push(query);
        return []; // Return empty array - we only care about the query
      }),
      sendEmail: jest.fn(),
      getUserLabels: jest.fn(() => [])
    };

    // Set up global mocks
    global.PropertiesService = {
      getUserProperties: jest.fn(() => mockPropertiesService)
    };

    global.GmailApp = mockGmailApp;
    global.Session = {
      getActiveUser: jest.fn(() => ({
        getEmail: jest.fn(() => 'test@example.com')
      }))
    };

    // Update script context
    scriptContext.PropertiesService = global.PropertiesService;
    scriptContext.GmailApp = mockGmailApp;
    scriptContext.Session = global.Session;

    // Mock getUserEmailAddress
    global.getUserEmailAddress = jest.fn(() => 'test@example.com');
    scriptContext.getUserEmailAddress = global.getUserEmailAddress;

    // Mock shouldIgnoreEmail (from EmailProcessor.js)
    scriptContext.shouldIgnoreEmail = jest.fn(() => false);

    // Mock processEmailsInBatches to avoid actual processing
    global.processEmailsInBatches = jest.fn(() => ({
      mustDo: [],
      mustKnow: [],
      totalProcessed: 0,
      batchesProcessed: 0
    }));
    scriptContext.processEmailsInBatches = global.processEmailsInBatches;

    // Mock other functions
    global.calculatePassiveWorkflowDateRange = jest.fn(() => {
      const now = new Date();
      const start = new Date(now.getTime() - 24 * 60 * 60 * 1000);
      return { start, end: now };
    });
    scriptContext.calculatePassiveWorkflowDateRange = global.calculatePassiveWorkflowDateRange;

    global.calculateDateRange = jest.fn(() => {
      const now = new Date();
      const start = new Date(now.getTime() - 24 * 60 * 60 * 1000);
      return { start, end: now };
    });
    scriptContext.calculateDateRange = global.calculateDateRange;
  });

  describe('Passive workflow applies inboxOnly setting', () => {
    test('should include "in:inbox" in query when inboxOnly is true', () => {
      // Arrange: Set inboxOnly to true in configuration
      mockPropertiesStore['inboxOnly'] = 'true';
      mockPropertiesStore['addonName'] = 'Gmail Fast Reader';
      mockPropertiesStore['openaiApiKey'] = 'test-key';

      // Act: Call fetchEmailThreadsForPassiveWorkflow
      const dateRange = {
        start: new Date(Date.now() - 24 * 60 * 60 * 1000),
        end: new Date()
      };
      scriptContext.fetchEmailThreadsForPassiveWorkflow(dateRange);

      // Assert: Verify the search query includes "in:inbox"
      expect(mockGmailApp.search).toHaveBeenCalled();
      const lastQuery = searchCallQueries[searchCallQueries.length - 1];
      expect(lastQuery).toContain('in:inbox');
    });

    test('should not include "in:inbox" in query when inboxOnly is false', () => {
      // Arrange: Set inboxOnly to false in configuration
      mockPropertiesStore['inboxOnly'] = 'false';
      mockPropertiesStore['addonName'] = 'Gmail Fast Reader';
      mockPropertiesStore['openaiApiKey'] = 'test-key';

      // Act: Call fetchEmailThreadsForPassiveWorkflow
      const dateRange = {
        start: new Date(Date.now() - 24 * 60 * 60 * 1000),
        end: new Date()
      };
      scriptContext.fetchEmailThreadsForPassiveWorkflow(dateRange);

      // Assert: Verify the search query does not include "in:inbox"
      expect(mockGmailApp.search).toHaveBeenCalled();
      const lastQuery = searchCallQueries[searchCallQueries.length - 1];
      expect(lastQuery).not.toContain('in:inbox');
    });

    test('should include "in:inbox" when inboxOnly is undefined (defaults to false, but test explicit true)', () => {
      // Arrange: Explicitly set inboxOnly to true
      mockPropertiesStore['inboxOnly'] = 'true';
      mockPropertiesStore['addonName'] = 'Gmail Fast Reader';
      mockPropertiesStore['openaiApiKey'] = 'test-key';

      // Act: Call fetchEmailThreadsForPassiveWorkflow
      const dateRange = {
        start: new Date(Date.now() - 24 * 60 * 60 * 1000),
        end: new Date()
      };
      scriptContext.fetchEmailThreadsForPassiveWorkflow(dateRange);

      // Assert: Verify the search query includes "in:inbox"
      expect(mockGmailApp.search).toHaveBeenCalled();
      const lastQuery = searchCallQueries[searchCallQueries.length - 1];
      expect(lastQuery).toContain('in:inbox');
    });
  });

  describe('Active workflow applies inboxOnly setting', () => {
    test('should include "in:inbox" in query when inboxOnly is true', () => {
      // Arrange: Set inboxOnly to true in configuration
      mockPropertiesStore['inboxOnly'] = 'true';
      mockPropertiesStore['addonName'] = 'Gmail Fast Reader';
      mockPropertiesStore['openaiApiKey'] = 'test-key';

      // Act: Call fetchEmailThreadsFromGmail (used by active workflow)
      const dateRange = {
        start: new Date(Date.now() - 24 * 60 * 60 * 1000),
        end: new Date()
      };
      scriptContext.fetchEmailThreadsFromGmail(dateRange);

      // Assert: Verify the search query includes "in:inbox"
      expect(mockGmailApp.search).toHaveBeenCalled();
      const lastQuery = searchCallQueries[searchCallQueries.length - 1];
      expect(lastQuery).toContain('in:inbox');
    });

    test('should not include "in:inbox" in query when inboxOnly is false', () => {
      // Arrange: Set inboxOnly to false in configuration
      mockPropertiesStore['inboxOnly'] = 'false';
      mockPropertiesStore['addonName'] = 'Gmail Fast Reader';
      mockPropertiesStore['openaiApiKey'] = 'test-key';

      // Act: Call fetchEmailThreadsFromGmail (used by active workflow)
      const dateRange = {
        start: new Date(Date.now() - 24 * 60 * 60 * 1000),
        end: new Date()
      };
      scriptContext.fetchEmailThreadsFromGmail(dateRange);

      // Assert: Verify the search query does not include "in:inbox"
      expect(mockGmailApp.search).toHaveBeenCalled();
      const lastQuery = searchCallQueries[searchCallQueries.length - 1];
      expect(lastQuery).not.toContain('in:inbox');
    });

    test('should work with unreadOnly and inboxOnly together', () => {
      // Arrange: Set both unreadOnly and inboxOnly to true
      mockPropertiesStore['inboxOnly'] = 'true';
      mockPropertiesStore['unreadOnly'] = 'true';
      mockPropertiesStore['addonName'] = 'Gmail Fast Reader';
      mockPropertiesStore['openaiApiKey'] = 'test-key';

      // Act: Call fetchEmailThreadsFromGmail
      const dateRange = {
        start: new Date(Date.now() - 24 * 60 * 60 * 1000),
        end: new Date()
      };
      scriptContext.fetchEmailThreadsFromGmail(dateRange);

      // Assert: Verify the search query includes both "in:inbox" and "is:unread"
      expect(mockGmailApp.search).toHaveBeenCalled();
      const lastQuery = searchCallQueries[searchCallQueries.length - 1];
      expect(lastQuery).toContain('in:inbox');
      expect(lastQuery).toContain('is:unread');
    });
  });

  describe('Both workflows consistently apply inboxOnly', () => {
    test('should apply inboxOnly setting consistently in both passive and active workflows', () => {
      // Arrange: Set inboxOnly to true
      mockPropertiesStore['inboxOnly'] = 'true';
      mockPropertiesStore['addonName'] = 'Gmail Fast Reader';
      mockPropertiesStore['openaiApiKey'] = 'test-key';

      const dateRange = {
        start: new Date(Date.now() - 24 * 60 * 60 * 1000),
        end: new Date()
      };

      // Clear previous search calls
      searchCallQueries = [];
      mockGmailApp.search.mockClear();

      // Act: Call both fetch functions
      scriptContext.fetchEmailThreadsForPassiveWorkflow(dateRange);
      const passiveQuery = searchCallQueries[searchCallQueries.length - 1];

      searchCallQueries = [];
      mockGmailApp.search.mockClear();

      scriptContext.fetchEmailThreadsFromGmail(dateRange);
      const activeQuery = searchCallQueries[searchCallQueries.length - 1];

      // Assert: Both queries should include "in:inbox"
      expect(passiveQuery).toContain('in:inbox');
      expect(activeQuery).toContain('in:inbox');
    });
  });
});

