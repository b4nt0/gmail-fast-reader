/**
 * Regression tests for email processing options
 * 
 * Tests verify that:
 * 1. markProcessedAsRead configuration defaults to false and works correctly
 * 2. removeUninterestingFromInbox configuration defaults to false and works correctly
 * 3. Both options work together correctly
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
  ReferenceError: ReferenceError,
  Set: Set
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

// Load Config.js
const configJsPath = path.join(__dirname, '../addon/Config.js');
const configCode = fs.readFileSync(configJsPath, 'utf8');
vm.runInContext(configCode, scriptContext);

// Load EmailProcessor.js
const emailProcessorJsPath = path.join(__dirname, '../addon/EmailProcessor.js');
const emailProcessorCode = fs.readFileSync(emailProcessorJsPath, 'utf8');
vm.runInContext(emailProcessorCode, scriptContext);

describe('Email Processing Options Regression Tests', () => {
  let mockPropertiesStore;
  let mockPropertiesService;
  let mockGmailApp;
  let mockMessages;
  let mockThreads;

  beforeEach(() => {
    jest.clearAllMocks();
    mockMessages = [];
    mockThreads = [];

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

    // Create mock GmailApp
    const mockInboxLabel = {
      getName: () => 'INBOX'
    };

    mockGmailApp = {
      getMessageById: jest.fn((id) => {
        const msg = mockMessages.find(m => m.getId() === id);
        return msg || null;
      }),
      getThreadById: jest.fn((id) => {
        const thread = mockThreads.find(t => t.getId() === id);
        return thread || null;
      }),
      search: jest.fn((query) => {
        // Return mock threads if searching for rfc822msgid
        if (query.includes('rfc822msgid:')) {
          return mockThreads.filter(t => {
            const messages = t.getMessages();
            return messages.some(m => {
              try {
                const rawContent = m.getRawContent();
                const match = rawContent.match(/Message-ID:\s*<([^>]+)>/i);
                return match && match[1] === query.split(':')[1];
              } catch (e) {
                return false;
              }
            });
          });
        }
        return [];
      }),
      getInboxLabel: jest.fn(() => mockInboxLabel),
      getUserLabelByName: jest.fn(() => null),
      createLabel: jest.fn((name) => ({ getName: () => name })),
      sendEmail: jest.fn()
    };

    // Setup mocks in script context
    scriptContext.PropertiesService = {
      getUserProperties: jest.fn(() => mockPropertiesService)
    };
    scriptContext.GmailApp = mockGmailApp;
    scriptContext.getUserEmailAddress = jest.fn(() => 'test@example.com');

    // Reset properties store
    mockPropertiesStore = {
      'addonName': 'Gmail Fast Reader',
      'openaiApiKey': 'test-key'
    };
  });

  describe('Configuration defaults', () => {
    test('markProcessedAsRead defaults to false', () => {
      const config = scriptContext.getConfiguration();
      expect(config.markProcessedAsRead).toBe(false);
    });

    test('removeUninterestingFromInbox defaults to false', () => {
      const config = scriptContext.getConfiguration();
      expect(config.removeUninterestingFromInbox).toBe(false);
    });

    test('markProcessedAsRead can be set to true', () => {
      mockPropertiesStore['markProcessedAsRead'] = 'true';
      const config = scriptContext.getConfiguration();
      expect(config.markProcessedAsRead).toBe(true);
    });

    test('removeUninterestingFromInbox can be set to true', () => {
      mockPropertiesStore['removeUninterestingFromInbox'] = 'true';
      const config = scriptContext.getConfiguration();
      expect(config.removeUninterestingFromInbox).toBe(true);
    });
  });

  describe('markProcessedEmailsAsRead function', () => {
    test('should not mark emails as read when option is disabled', () => {
      const mockMessage = {
        getId: () => 'msg1',
        markRead: jest.fn(),
        getThread: () => ({ addLabel: jest.fn() })
      };
      mockMessages.push(mockMessage);

      const results = {
        mustDo: [{
          emailId: 'msg1',
          subject: 'Test Email',
          sender: 'test@example.com'
        }],
        mustKnow: []
      };

      const config = {
        markProcessedAsRead: false,
        mustDoLabel: '',
        mustKnowLabel: ''
      };

      scriptContext.markProcessedEmailsAsRead(results, config);

      expect(mockMessage.markRead).not.toHaveBeenCalled();
    });

    test('should mark emails as read when option is enabled', () => {
      const mockMessage = {
        getId: () => 'msg1',
        markRead: jest.fn(),
        getThread: () => ({ addLabel: jest.fn() })
      };
      mockMessages.push(mockMessage);

      const results = {
        mustDo: [{
          emailId: 'msg1',
          subject: 'Test Email',
          sender: 'test@example.com'
        }],
        mustKnow: []
      };

      const config = {
        markProcessedAsRead: true,
        mustDoLabel: '',
        mustKnowLabel: ''
      };

      scriptContext.markProcessedEmailsAsRead(results, config);

      expect(mockMessage.markRead).toHaveBeenCalled();
    });

    test('should mark both mustDo and mustKnow emails as read', () => {
      const mockMessage1 = {
        getId: () => 'msg1',
        markRead: jest.fn(),
        getThread: () => ({ addLabel: jest.fn() })
      };
      const mockMessage2 = {
        getId: () => 'msg2',
        markRead: jest.fn(),
        getThread: () => ({ addLabel: jest.fn() })
      };
      mockMessages.push(mockMessage1, mockMessage2);

      const results = {
        mustDo: [{
          emailId: 'msg1',
          subject: 'Test Email 1',
          sender: 'test@example.com'
        }],
        mustKnow: [{
          emailId: 'msg2',
          subject: 'Test Email 2',
          sender: 'test2@example.com'
        }]
      };

      const config = {
        markProcessedAsRead: true,
        mustDoLabel: '',
        mustKnowLabel: ''
      };

      scriptContext.markProcessedEmailsAsRead(results, config);

      expect(mockMessage1.markRead).toHaveBeenCalled();
      expect(mockMessage2.markRead).toHaveBeenCalled();
    });

    test('should handle rfc822MessageId fallback', () => {
      const mockMessage = {
        getId: () => 'msg1',
        markRead: jest.fn(),
        getRawContent: () => 'Message-ID: <test@example.com>',
        getThread: () => ({ addLabel: jest.fn() })
      };
      const mockThread = {
        getId: () => 'thread1',
        getMessages: () => [mockMessage]
      };
      mockThreads.push(mockThread);

      const results = {
        mustDo: [{
          emailId: 'nonexistent',
          rfc822MessageId: 'test@example.com',
          subject: 'Test Email',
          sender: 'test@example.com'
        }],
        mustKnow: []
      };

      const config = {
        markProcessedAsRead: true,
        mustDoLabel: '',
        mustKnowLabel: ''
      };

      scriptContext.markProcessedEmailsAsRead(results, config);

      expect(mockMessage.markRead).toHaveBeenCalled();
    });
  });

  describe('removeUninterestingEmailsFromInbox function', () => {
    test('should not remove threads from inbox when option is disabled', () => {
      const mockThread = {
        getId: () => 'thread1',
        removeLabel: jest.fn(),
        getMessages: () => []
      };
      mockThreads.push(mockThread);

      const emailThreads = [{
        threadId: 'thread1',
        emails: [{
          id: 'msg1',
          subject: 'Uninteresting Email'
        }]
      }];

      const results = {
        mustDo: [],
        mustKnow: []
      };

      const config = {
        removeUninterestingFromInbox: false
      };

      scriptContext.removeUninterestingEmailsFromInbox(emailThreads, results, config);

      expect(mockThread.removeLabel).not.toHaveBeenCalled();
    });

    test('should remove uninteresting threads from inbox when option is enabled', () => {
      const mockInboxLabel = { getName: () => 'INBOX' };
      mockGmailApp.getInboxLabel = jest.fn(() => mockInboxLabel);

      const mockThread = {
        getId: () => 'thread1',
        removeLabel: jest.fn(),
        getMessages: () => []
      };
      mockThreads.push(mockThread);
      mockGmailApp.getThreadById = jest.fn((id) => {
        if (id === 'thread1') return mockThread;
        return null;
      });

      const emailThreads = [{
        threadId: 'thread1',
        emails: [{
          id: 'msg1',
          subject: 'Uninteresting Email'
        }]
      }];

      const results = {
        mustDo: [],
        mustKnow: []
      };

      const config = {
        removeUninterestingFromInbox: true
      };

      scriptContext.removeUninterestingEmailsFromInbox(emailThreads, results, config);

      expect(mockThread.removeLabel).toHaveBeenCalledWith(mockInboxLabel);
    });

    test('should not remove threads with interesting emails', () => {
      const mockInboxLabel = { getName: () => 'INBOX' };
      mockGmailApp.getInboxLabel = jest.fn(() => mockInboxLabel);

      const mockThread = {
        getId: () => 'thread1',
        removeLabel: jest.fn(),
        getMessages: () => []
      };
      mockThreads.push(mockThread);
      mockGmailApp.getThreadById = jest.fn((id) => {
        if (id === 'thread1') return mockThread;
        return null;
      });

      const emailThreads = [{
        threadId: 'thread1',
        emails: [{
          id: 'msg1',
          subject: 'Interesting Email'
        }]
      }];

      const results = {
        mustDo: [{
          emailId: 'msg1',
          subject: 'Interesting Email'
        }],
        mustKnow: []
      };

      const config = {
        removeUninterestingFromInbox: true
      };

      scriptContext.removeUninterestingEmailsFromInbox(emailThreads, results, config);

      expect(mockThread.removeLabel).not.toHaveBeenCalled();
    });

    test('should handle threads with rfc822MessageId matching', () => {
      const mockInboxLabel = { getName: () => 'INBOX' };
      mockGmailApp.getInboxLabel = jest.fn(() => mockInboxLabel);

      const mockThread = {
        getId: () => 'thread1',
        removeLabel: jest.fn(),
        getMessages: () => []
      };
      mockThreads.push(mockThread);
      mockGmailApp.getThreadById = jest.fn((id) => {
        if (id === 'thread1') return mockThread;
        return null;
      });

      const emailThreads = [{
        threadId: 'thread1',
        emails: [{
          id: 'msg1',
          rfc822MessageId: 'test@example.com',
          subject: 'Uninteresting Email'
        }]
      }];

      const results = {
        mustDo: [],
        mustKnow: []
      };

      const config = {
        removeUninterestingFromInbox: true
      };

      scriptContext.removeUninterestingEmailsFromInbox(emailThreads, results, config);

      expect(mockThread.removeLabel).toHaveBeenCalledWith(mockInboxLabel);
    });
  });

  describe('Integration tests', () => {
    test('both options work together correctly', () => {
      const mockInboxLabel = { getName: () => 'INBOX' };
      mockGmailApp.getInboxLabel = jest.fn(() => mockInboxLabel);

      const mockMessage1 = {
        getId: () => 'msg1',
        markRead: jest.fn(),
        getThread: () => ({ addLabel: jest.fn() })
      };
      const mockMessage2 = {
        getId: () => 'msg2',
        markRead: jest.fn(),
        getThread: () => ({ addLabel: jest.fn() })
      };
      mockMessages.push(mockMessage1, mockMessage2);

      const mockThread1 = {
        getId: () => 'thread1',
        removeLabel: jest.fn(),
        getMessages: () => []
      };
      const mockThread2 = {
        getId: () => 'thread2',
        removeLabel: jest.fn(),
        getMessages: () => []
      };
      mockThreads.push(mockThread1, mockThread2);
      mockGmailApp.getThreadById = jest.fn((id) => {
        if (id === 'thread1') return mockThread1;
        if (id === 'thread2') return mockThread2;
        return null;
      });

      const emailThreads = [
        {
          threadId: 'thread1',
          emails: [{
            id: 'msg1',
            subject: 'Interesting Email'
          }]
        },
        {
          threadId: 'thread2',
          emails: [{
            id: 'msg2',
            subject: 'Uninteresting Email'
          }]
        }
      ];

      const results = {
        mustDo: [{
          emailId: 'msg1',
          subject: 'Interesting Email'
        }],
        mustKnow: []
      };

      const config = {
        markProcessedAsRead: true,
        removeUninterestingFromInbox: true,
        mustDoLabel: '',
        mustKnowLabel: ''
      };

      // Test markProcessedEmailsAsRead
      scriptContext.markProcessedEmailsAsRead(results, config);
      expect(mockMessage1.markRead).toHaveBeenCalled();
      expect(mockMessage2.markRead).not.toHaveBeenCalled();

      // Test removeUninterestingEmailsFromInbox
      scriptContext.removeUninterestingEmailsFromInbox(emailThreads, results, config);
      expect(mockThread1.removeLabel).not.toHaveBeenCalled(); // Has interesting email
      expect(mockThread2.removeLabel).toHaveBeenCalledWith(mockInboxLabel); // Uninteresting
    });
  });
});

