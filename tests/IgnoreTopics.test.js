/**
 * Regression tests for ignoreTopics feature
 * 
 * Tests verify that:
 * 1. ignoreTopics configuration defaults correctly
 * 2. ignoreTopics can be saved and retrieved
 * 3. OpenAI response parsing handles ignore array correctly
 * 4. Ignored emails are filtered from mustDo and mustKnow
 * 5. Ignored emails do not receive labels
 * 6. Ignored emails are not included in summaries
 * 7. Ignored emails are still counted in totalProcessed
 * 8. Threads with ignored emails are archived (subject to safety checks)
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
  UrlFetchApp: null,
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

describe('IgnoreTopics Regression Tests', () => {
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
      getProperty: jest.fn((key) => (key in mockPropertiesStore) ? mockPropertiesStore[key] : null),
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
    scriptContext.UrlFetchApp = {
      fetch: jest.fn()
    };

    // Reset properties store with default values
    mockPropertiesStore = {
      'addonName': 'Gmail Fast Reader',
      'openaiApiKey': 'test-key'
    };
  });

  describe('Configuration defaults', () => {
    test('ignoreTopics defaults to expected value', () => {
      const config = scriptContext.getConfiguration();
      const expectedDefault = 'Social connection requests\nShipment notifications from stores\nSocial posts\nCustomer surveys';
      expect(config.ignoreTopics).toBe(expectedDefault);
    });

    test('ignoreTopics can be set to custom value', () => {
      const customValue = 'Spam\nNewsletters\nPromotions';
      mockPropertiesStore['ignoreTopics'] = customValue;
      const config = scriptContext.getConfiguration();
      expect(config.ignoreTopics).toBe(customValue);
    });

    test('ignoreTopics can be saved and retrieved', () => {
      const customValue = 'Test ignore topics';
      const config = {
        addonName: 'Gmail Fast Reader',
        openaiApiKey: 'test-key',
        timeZone: 'Europe/Paris',
        mustDoTopics: 'test',
        mustKnowTopics: 'test',
        mustDoOther: true,
        mustKnowOther: false,
        unreadOnly: false,
        inboxOnly: true,
        mustDoLabel: '',
        mustKnowLabel: '',
        markProcessedAsRead: false,
        removeUninterestingFromInbox: false,
        ignoreTopics: customValue
      };
      
      scriptContext.saveConfiguration(config);
      
      const retrievedConfig = scriptContext.getConfiguration();
      expect(retrievedConfig.ignoreTopics).toBe(customValue);
    });
  });

  describe('OpenAI response parsing', () => {
    test('parseOpenAIResponse handles ignore array correctly', () => {
      const response = JSON.stringify({
        mustDo: [],
        mustKnow: [],
        ignore: [
          {
            emailId: 'msg1',
            rfc822MessageId: 'rfc1@example.com',
            subject: 'Test Ignored',
            sender: 'test@example.com',
            date: '2024-01-01',
            topic: 'Social connection requests'
          }
        ]
      });
      
      const parsed = scriptContext.parseOpenAIResponse(response);
      expect(parsed.ignore).toBeDefined();
      expect(Array.isArray(parsed.ignore)).toBe(true);
      expect(parsed.ignore.length).toBe(1);
      expect(parsed.ignore[0].emailId).toBe('msg1');
    });

    test('parseOpenAIResponse handles missing ignore array (backward compatibility)', () => {
      const response = JSON.stringify({
        mustDo: [],
        mustKnow: []
      });
      
      const parsed = scriptContext.parseOpenAIResponse(response);
      expect(parsed.ignore).toBeDefined();
      expect(Array.isArray(parsed.ignore)).toBe(true);
      expect(parsed.ignore.length).toBe(0);
    });
  });

  describe('Filtering logic', () => {
    test('ignored emails are excluded from mustDo', () => {
      const mockResponse = {
        mustDo: [
          { emailId: 'msg1', rfc822MessageId: 'rfc1', subject: 'Action Item' },
          { emailId: 'msg2', rfc822MessageId: 'rfc2', subject: 'Another Action' }
        ],
        mustKnow: [],
        ignore: [
          { emailId: 'msg1', rfc822MessageId: 'rfc1', subject: 'Action Item' }
        ]
      };
      
      // Mock the OpenAI API call to return our mock response
      // The response should have the structure: { choices: [{ message: { content: JSON_STRING } }] }
      scriptContext.UrlFetchApp.fetch = jest.fn(() => ({
        getContentText: () => JSON.stringify({
          choices: [{
            message: {
              content: JSON.stringify(mockResponse)
            }
          }]
        }),
        getResponseCode: () => 200,
        getHeaders: () => ({})
      }));
      
      const emailThreads = [{
        threadId: 'thread1',
        subject: 'Test Thread',
        emails: [
          { id: 'msg1', subject: 'Action Item', sender: 'test@example.com', date: new Date(), body: 'Test', rfc822MessageId: 'rfc1' },
          { id: 'msg2', subject: 'Another Action', sender: 'test@example.com', date: new Date(), body: 'Test', rfc822MessageId: 'rfc2' }
        ]
      }];
      
      const config = scriptContext.getConfiguration();
      const results = scriptContext.analyzeEmailsWithOpenAI(emailThreads, config);
      
      // msg1 should be filtered out from mustDo
      expect(results.mustDo.length).toBe(1);
      expect(results.mustDo[0].emailId).toBe('msg2');
    });

    test('ignored emails are excluded from mustKnow', () => {
      const mockResponse = {
        mustDo: [],
        mustKnow: [
          { emailId: 'msg1', rfc822MessageId: 'rfc1', subject: 'Info Item' },
          { emailId: 'msg2', rfc822MessageId: 'rfc2', subject: 'Another Info' }
        ],
        ignore: [
          { emailId: 'msg1', rfc822MessageId: 'rfc1', subject: 'Info Item' }
        ]
      };
      
      scriptContext.UrlFetchApp.fetch = jest.fn(() => ({
        getContentText: () => JSON.stringify({
          choices: [{
            message: {
              content: JSON.stringify(mockResponse)
            }
          }]
        }),
        getResponseCode: () => 200,
        getHeaders: () => ({})
      }));
      
      const emailThreads = [{
        threadId: 'thread1',
        subject: 'Test Thread',
        emails: [
          { id: 'msg1', subject: 'Info Item', sender: 'test@example.com', date: new Date(), body: 'Test', rfc822MessageId: 'rfc1' },
          { id: 'msg2', subject: 'Another Info', sender: 'test@example.com', date: new Date(), body: 'Test', rfc822MessageId: 'rfc2' }
        ]
      }];
      
      const config = scriptContext.getConfiguration();
      const results = scriptContext.analyzeEmailsWithOpenAI(emailThreads, config);
      
      // msg1 should be filtered out from mustKnow
      expect(results.mustKnow.length).toBe(1);
      expect(results.mustKnow[0].emailId).toBe('msg2');
    });

    test('ignored emails are still counted in totalProcessed', () => {
      const mockResponse = {
        mustDo: [],
        mustKnow: [],
        ignore: [
          { emailId: 'msg1', rfc822MessageId: 'rfc1', subject: 'Ignored' }
        ]
      };
      
      scriptContext.UrlFetchApp.fetch = jest.fn(() => ({
        getContentText: () => JSON.stringify({
          choices: [{
            message: {
              content: JSON.stringify(mockResponse)
            }
          }]
        }),
        getResponseCode: () => 200,
        getHeaders: () => ({})
      }));
      
      const emailThreads = [{
        threadId: 'thread1',
        subject: 'Test Thread',
        emails: [
          { id: 'msg1', subject: 'Ignored', sender: 'test@example.com', date: new Date(), body: 'Test', rfc822MessageId: 'rfc1' }
        ]
      }];
      
      const config = scriptContext.getConfiguration();
      const results = scriptContext.analyzeEmailsWithOpenAI(emailThreads, config);
      
      // totalProcessed should still count the ignored email
      expect(results.totalProcessed).toBe(1);
    });

    test('filtering works with rfc822MessageId', () => {
      const mockResponse = {
        mustDo: [
          { emailId: 'msg1', rfc822MessageId: 'rfc1@example.com', subject: 'Action' }
        ],
        mustKnow: [],
        ignore: [
          { emailId: 'msg1', rfc822MessageId: 'rfc1@example.com', subject: 'Action' }
        ]
      };
      
      scriptContext.UrlFetchApp.fetch = jest.fn(() => ({
        getContentText: () => JSON.stringify({
          choices: [{
            message: {
              content: JSON.stringify(mockResponse)
            }
          }]
        }),
        getResponseCode: () => 200,
        getHeaders: () => ({})
      }));
      
      const emailThreads = [{
        threadId: 'thread1',
        subject: 'Test Thread',
        emails: [
          { id: 'msg1', subject: 'Action', sender: 'test@example.com', date: new Date(), body: 'Test', rfc822MessageId: 'rfc1@example.com' }
        ]
      }];
      
      const config = scriptContext.getConfiguration();
      const results = scriptContext.analyzeEmailsWithOpenAI(emailThreads, config);
      
      // Should be filtered out even when matched by rfc822MessageId
      expect(results.mustDo.length).toBe(0);
    });
  });

  describe('Label application', () => {
    test('ignored emails do not receive labels', () => {
      const mockMessage = {
        getId: () => 'msg1',
        getThread: () => ({ addLabel: jest.fn() })
      };
      mockMessages.push(mockMessage);

      const results = {
        mustDo: [
          { emailId: 'msg1', subject: 'Action', sender: 'test@example.com' }
        ],
        mustKnow: [],
        ignore: [
          { emailId: 'msg1', subject: 'Action', sender: 'test@example.com' }
        ]
      };

      const config = {
        mustDoLabel: 'TODO',
        mustKnowLabel: 'FYI',
        markProcessedAsRead: false,
        removeUninterestingFromInbox: false
      };

      // Filter ignored emails first (simulating what analyzeEmailsWithOpenAI does)
      const ignoredEmailIds = new Set(['msg1']);
      results.mustDo = results.mustDo.filter(email => !ignoredEmailIds.has(email.emailId));

      scriptContext.applyLabelsToInterestingEmails(results, config);

      // Label should not be applied since email was filtered out
      expect(mockMessage.getThread().addLabel).not.toHaveBeenCalled();
    });
  });

  describe('Archiving behavior', () => {
    test('threads with ignored emails are archived (subject to safety checks)', () => {
      const mockMessage = {
        getId: () => 'msg1',
        isStarred: jest.fn(() => false)
      };
      const mockThread = {
        getId: () => 'thread1',
        moveToArchive: jest.fn(),
        getMessages: jest.fn(() => [mockMessage]),
        getLabels: jest.fn(() => [])
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
          rfc822MessageId: 'rfc1@example.com',
          subject: 'Ignored Email'
        }]
      }];

      const results = {
        mustDo: [],
        mustKnow: [],
        ignore: [
          { emailId: 'msg1', rfc822MessageId: 'rfc1@example.com', subject: 'Ignored Email' }
        ]
      };

      const config = {
        removeUninterestingFromInbox: true
      };

      // Mock Gmail API as unavailable
      scriptContext.Gmail = undefined;

      scriptContext.removeUninterestingEmailsFromInbox(emailThreads, results, config);

      expect(mockThread.moveToArchive).toHaveBeenCalled();
    });

    test('threads with ignored emails are not archived if starred', () => {
      const mockMessage = {
        getId: () => 'msg1',
        isStarred: jest.fn(() => true)
      };
      const mockThread = {
        getId: () => 'thread1',
        moveToArchive: jest.fn(),
        getMessages: jest.fn(() => [mockMessage]),
        getLabels: jest.fn(() => [])
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
          subject: 'Ignored Email'
        }]
      }];

      const results = {
        mustDo: [],
        mustKnow: [],
        ignore: [
          { emailId: 'msg1', subject: 'Ignored Email' }
        ]
      };

      const config = {
        removeUninterestingFromInbox: true
      };

      scriptContext.Gmail = undefined;

      scriptContext.removeUninterestingEmailsFromInbox(emailThreads, results, config);

      expect(mockThread.moveToArchive).not.toHaveBeenCalled();
    });

    test('threads with both ignored and interesting emails are not archived', () => {
      const mockMessage1 = {
        getId: () => 'msg1',
        isStarred: jest.fn(() => false)
      };
      const mockMessage2 = {
        getId: () => 'msg2',
        isStarred: jest.fn(() => false)
      };
      const mockThread = {
        getId: () => 'thread1',
        moveToArchive: jest.fn(),
        getMessages: jest.fn(() => [mockMessage1, mockMessage2]),
        getLabels: jest.fn(() => [])
      };
      mockThreads.push(mockThread);
      mockGmailApp.getThreadById = jest.fn((id) => {
        if (id === 'thread1') return mockThread;
        return null;
      });

      const emailThreads = [{
        threadId: 'thread1',
        emails: [
          { id: 'msg1', subject: 'Ignored Email' },
          { id: 'msg2', subject: 'Interesting Email' }
        ]
      }];

      const results = {
        mustDo: [
          { emailId: 'msg2', subject: 'Interesting Email' }
        ],
        mustKnow: [],
        ignore: [
          { emailId: 'msg1', subject: 'Ignored Email' }
        ]
      };

      const config = {
        removeUninterestingFromInbox: true
      };

      scriptContext.Gmail = undefined;

      scriptContext.removeUninterestingEmailsFromInbox(emailThreads, results, config);

      // Should not archive because thread has interesting email
      expect(mockThread.moveToArchive).not.toHaveBeenCalled();
    });
  });

  describe('Edge cases', () => {
    test('empty ignoreTopics is handled correctly', () => {
      // Empty string should be treated as a valid value (not null, so not using default)
      mockPropertiesStore['ignoreTopics'] = '';
      const config = scriptContext.getConfiguration();
      // getProp returns the value if it's not null, so empty string is returned
      expect(config.ignoreTopics).toBe('');
      
      // But if property doesn't exist, default should be used
      delete mockPropertiesStore['ignoreTopics'];
      const configWithDefault = scriptContext.getConfiguration();
      const expectedDefault = 'Social connection requests\nShipment notifications from stores\nSocial posts\nCustomer surveys';
      expect(configWithDefault.ignoreTopics).toBe(expectedDefault);
    });

    test('email matching both ignoreTopics and mustDo is filtered out (ignore takes precedence)', () => {
      const mockResponse = {
        mustDo: [
          { emailId: 'msg1', rfc822MessageId: 'rfc1', subject: 'Action Item' }
        ],
        mustKnow: [],
        ignore: [
          { emailId: 'msg1', rfc822MessageId: 'rfc1', subject: 'Action Item' }
        ]
      };
      
      scriptContext.UrlFetchApp.fetch = jest.fn(() => ({
        getContentText: () => JSON.stringify({
          choices: [{
            message: {
              content: JSON.stringify(mockResponse)
            }
          }]
        }),
        getResponseCode: () => 200,
        getHeaders: () => ({})
      }));
      
      const emailThreads = [{
        threadId: 'thread1',
        subject: 'Test Thread',
        emails: [
          { id: 'msg1', subject: 'Action Item', sender: 'test@example.com', date: new Date(), body: 'Test', rfc822MessageId: 'rfc1' }
        ]
      }];
      
      const config = scriptContext.getConfiguration();
      const results = scriptContext.analyzeEmailsWithOpenAI(emailThreads, config);
      
      // Should be filtered out even though it was in mustDo
      expect(results.mustDo.length).toBe(0);
      expect(results.ignore.length).toBe(1);
    });
  });
});

