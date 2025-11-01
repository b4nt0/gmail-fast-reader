/**
 * Regression tests for daily summary email functionality
 * 
 * Tests verify that:
 * - Results accumulate across multiple passive workflow runs
 * - Summary is only sent between 21:00-23:59 in user's timezone
 * - Summary is only sent once per day
 * - If window is missed, results accumulate and send next day
 * - Accumulated results are cleared only after successful send
 * - Drive operations work correctly
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
  DriveApp: null,
  Utilities: null,
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

// Define PROCESSING_STATUS constant
const PROCESSING_STATUS = {
  RUNNING: 'running',
  COMPLETED: 'completed',
  ERROR: 'error',
  TIMEOUT: 'timeout'
};

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

// Load DriveStorage.js
const driveStorageJsPath = path.join(__dirname, '../addon/DriveStorage.js');
const driveStorageCode = fs.readFileSync(driveStorageJsPath, 'utf8');
vm.runInContext(driveStorageCode, scriptContext);

// Load SummaryBuilder.js
const summaryBuilderJsPath = path.join(__dirname, '../addon/SummaryBuilder.js');
const summaryBuilderCode = fs.readFileSync(summaryBuilderJsPath, 'utf8');
vm.runInContext(summaryBuilderCode, scriptContext);

// Load Code.js
const codeJsPath = path.join(__dirname, '../addon/Code.js');
const codeJs = fs.readFileSync(codeJsPath, 'utf8');
vm.runInContext(codeJs, scriptContext);

describe('Daily Summary Regression Tests', () => {
  let mockPropertiesStore;
  let mockPropertiesService;
  let mockGmailApp;
  let mockDriveApp;
  let mockUtilities;
  let driveFiles;
  let emailSent;
  let emailSubject;

  beforeEach(() => {
    jest.clearAllMocks();
    driveFiles = {};
    emailSent = [];
    emailSubject = null;

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

    // Create mock DriveApp
    const mockRootFolder = {
      getFilesByName: jest.fn((name) => {
        // Create a new iterator each time - but track consumption per iterator
        const iteratorState = { consumed: false };
        const fileIterator = {
          hasNext: jest.fn(() => {
            // Return true if file exists and hasn't been consumed by this iterator
            return !iteratorState.consumed && Object.keys(driveFiles).includes(name);
          }),
          next: jest.fn(() => {
            if (!driveFiles[name]) {
              // File doesn't exist - create it lazily
              const defaultContent = JSON.stringify({
                mustDo: [],
                mustKnow: [],
                totalProcessed: 0,
                firstDate: null,
                lastDate: null
              });
              driveFiles[name] = {
                getId: jest.fn(() => 'file-' + name),
                getBlob: jest.fn(() => ({
                  getDataAsString: jest.fn(() => {
                    // Always return the current content
                    return driveFiles[name].content || defaultContent;
                  })
                })),
                setContent: jest.fn((content) => {
                  driveFiles[name].content = content;
                }),
                setTrashed: jest.fn(() => {
                  delete driveFiles[name];
                }),
                content: defaultContent
              };
            }
            const file = driveFiles[name];
            iteratorState.consumed = true; // Mark this iterator as consumed
            return file;
          })
        };
        return fileIterator;
      }),
      createFile: jest.fn((name, content, mimeType) => {
        if (!driveFiles[name]) {
          driveFiles[name] = {
            getId: jest.fn(() => 'file-' + name),
            getBlob: jest.fn(() => ({
              getDataAsString: jest.fn(() => {
                // Always return the current content, not the initial content
                return driveFiles[name].content || content;
              })
            })),
            setContent: jest.fn((newContent) => {
              driveFiles[name].content = newContent;
            }),
            setTrashed: jest.fn(() => {
              delete driveFiles[name];
            }),
            content: content
          };
        }
        return driveFiles[name];
      })
    };

    mockDriveApp = {
      getRootFolder: jest.fn(() => mockRootFolder)
    };

    // Create mock GmailApp
    mockGmailApp = {
      search: jest.fn(() => []),
      sendEmail: jest.fn((to, subject, body, options) => {
        emailSent.push({ to, subject, body, options });
        emailSubject = subject;
      }),
      getUserLabels: jest.fn(() => [])
    };

    // Create mock Utilities for timezone operations
    // Default behavior - will be overridden in tests
    mockUtilities = {
      formatDate: jest.fn((date, timezone, format) => {
        // Mock timezone conversion - simplified for testing
        const d = new Date(date);
        if (format === 'HH:mm') {
          // Default to current UTC time
          const hour = d.getUTCHours();
          const minute = d.getUTCMinutes();
          return `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}`;
        } else if (format === 'yyyy-MM-dd') {
          // Return date string
          const year = d.getUTCFullYear();
          const month = String(d.getUTCMonth() + 1).padStart(2, '0');
          const day = String(d.getUTCDate()).padStart(2, '0');
          return `${year}-${month}-${day}`;
        }
        return '';
      })
    };

    // Set up global mocks
    global.PropertiesService = {
      getUserProperties: jest.fn(() => mockPropertiesService)
    };

    global.GmailApp = mockGmailApp;
    global.DriveApp = mockDriveApp;
    global.Utilities = mockUtilities;
    global.Session = {
      getActiveUser: jest.fn(() => ({
        getEmail: jest.fn(() => 'test@example.com')
      }))
    };

    // Update script context
    scriptContext.PropertiesService = global.PropertiesService;
    scriptContext.GmailApp = mockGmailApp;
    scriptContext.DriveApp = mockDriveApp;
    scriptContext.Utilities = mockUtilities;
    scriptContext.Session = global.Session;

    // Mock getUserEmailAddress
    global.getUserEmailAddress = jest.fn(() => 'test@example.com');
    scriptContext.getUserEmailAddress = global.getUserEmailAddress;

    // Mock markEmailAsImportantOrStarred
    global.markEmailAsImportantOrStarred = jest.fn();
    scriptContext.markEmailAsImportantOrStarred = global.markEmailAsImportantOrStarred;

    // Mock generateSummaryHTML
    global.generateSummaryHTML = jest.fn(() => '<html>Summary</html>');
    scriptContext.generateSummaryHTML = global.generateSummaryHTML;

    // Ensure DriveStorage functions are available in scriptContext
    // (they're loaded from DriveStorage.js, but we need to make sure they're accessible)
    scriptContext.loadAccumulatedResults = global.loadAccumulatedResults || scriptContext.loadAccumulatedResults;
    scriptContext.saveAccumulatedResults = global.saveAccumulatedResults || scriptContext.saveAccumulatedResults;
    scriptContext.clearAccumulatedResults = global.clearAccumulatedResults || scriptContext.clearAccumulatedResults;
    scriptContext.getOrCreateAccumulationFile = global.getOrCreateAccumulationFile || scriptContext.getOrCreateAccumulationFile;
    
    // Mock Gmail API (used by markEmailAsImportantOrStarred)
    global.Gmail = {
      Users: {
        Messages: {
          modify: jest.fn()
        }
      }
    };
    scriptContext.Gmail = global.Gmail;

    // Mock processEmailsInBatches
    global.processEmailsInBatches = jest.fn(() => ({
      mustDo: [],
      mustKnow: [],
      totalProcessed: 0,
      batchesProcessed: 0
    }));
    scriptContext.processEmailsInBatches = global.processEmailsInBatches;

    // Mock calculatePassiveWorkflowDateRange
    global.calculatePassiveWorkflowDateRange = jest.fn(() => {
      const now = new Date();
      const start = new Date(now.getTime() - 60 * 60 * 1000); // 1 hour ago
      return { start, end: now };
    });
    scriptContext.calculatePassiveWorkflowDateRange = global.calculatePassiveWorkflowDateRange;

    // Mock fetchEmailThreadsForPassiveWorkflow
    global.fetchEmailThreadsForPassiveWorkflow = jest.fn(() => []);
    scriptContext.fetchEmailThreadsForPassiveWorkflow = global.fetchEmailThreadsForPassiveWorkflow;

    // Mock checkLock
    global.checkLock = jest.fn(() => ({ locked: false, type: null, expired: false }));
    scriptContext.checkLock = global.checkLock;

    // Mock lock/unlock
    global.lock = jest.fn();
    global.unlock = jest.fn();
    scriptContext.lock = global.lock;
    scriptContext.unlock = global.unlock;

    // Mock markChunkStarting/markChunkEnded
    global.markChunkStarting = jest.fn();
    global.markChunkEnded = jest.fn();
    scriptContext.markChunkStarting = global.markChunkStarting;
    scriptContext.markChunkEnded = global.markChunkEnded;

    // Set default config
    global.getConfiguration = jest.fn(() => ({
      addonName: 'Gmail Fast Reader',
      openaiApiKey: 'test-key',
      timeZone: 'America/New_York',
      mustDoTopics: 'test',
      mustKnowTopics: 'test',
      mustDoOther: false,
      mustKnowOther: false,
      unreadOnly: false,
      inboxOnly: false,
      mustDoLabel: '',
      mustKnowLabel: ''
    }));
    scriptContext.getConfiguration = global.getConfiguration;
  });

  describe('Result Accumulation', () => {
    test('should accumulate results across multiple passive workflow runs', () => {
      // First run with some results
      global.processEmailsInBatches = jest.fn(() => ({
        mustDo: [{ subject: 'Task 1', keyAction: 'Do something' }],
        mustKnow: [{ subject: 'Info 1', keyKnowledge: 'Know something' }],
        totalProcessed: 5
      }));
      scriptContext.processEmailsInBatches = global.processEmailsInBatches;

      const dateRange1 = { start: new Date('2024-01-15T10:00:00Z'), end: new Date('2024-01-15T11:00:00Z') };
      global.calculatePassiveWorkflowDateRange = jest.fn(() => dateRange1);
      scriptContext.calculatePassiveWorkflowDateRange = global.calculatePassiveWorkflowDateRange;

      global.fetchEmailThreadsForPassiveWorkflow = jest.fn(() => [{ emails: [{ date: dateRange1.start, id: 'msg1' }] }]);
      scriptContext.fetchEmailThreadsForPassiveWorkflow = global.fetchEmailThreadsForPassiveWorkflow;

      // Mock time outside window (so summary won't be sent)
      mockUtilities.formatDate = jest.fn((date, tz, format) => {
        if (format === 'HH:mm') return '10:00'; // 10 AM - outside window
        if (format === 'yyyy-MM-dd') return '2024-01-15';
        return '';
      });

      scriptContext.runPassiveWorkflow();

      // Verify results were saved to Drive
      const saved1 = JSON.parse(driveFiles['gmail-fast-read-accumulated-results.json'].content);
      expect(saved1.mustDo).toHaveLength(1);
      expect(saved1.mustKnow).toHaveLength(1);
      expect(saved1.totalProcessed).toBe(5);

      // Second run with more results
      global.processEmailsInBatches = jest.fn(() => ({
        mustDo: [{ subject: 'Task 2', keyAction: 'Do something else' }],
        mustKnow: [{ subject: 'Info 2', keyKnowledge: 'Know something else' }],
        totalProcessed: 3
      }));
      scriptContext.processEmailsInBatches = global.processEmailsInBatches;

      const dateRange2 = { start: new Date('2024-01-15T11:00:00Z'), end: new Date('2024-01-15T12:00:00Z') };
      global.calculatePassiveWorkflowDateRange = jest.fn(() => dateRange2);
      scriptContext.calculatePassiveWorkflowDateRange = global.calculatePassiveWorkflowDateRange;

      global.fetchEmailThreadsForPassiveWorkflow = jest.fn(() => [{ emails: [{ date: dateRange2.start, id: 'msg2' }] }]);
      scriptContext.fetchEmailThreadsForPassiveWorkflow = global.fetchEmailThreadsForPassiveWorkflow;

      scriptContext.runPassiveWorkflow();

      // Verify results were accumulated
      const saved2 = JSON.parse(driveFiles['gmail-fast-read-accumulated-results.json'].content);
      expect(saved2.mustDo).toHaveLength(2);
      expect(saved2.mustKnow).toHaveLength(2);
      expect(saved2.totalProcessed).toBe(8); // 5 + 3
    });
  });

  describe('Time Window Check', () => {
    test('should only send summary between 21:00-23:59 in user timezone', () => {
      // Mock time at 20:59 - outside window
      mockUtilities.formatDate = jest.fn((date, tz, format) => {
        if (format === 'HH:mm') return '20:59';
        if (format === 'yyyy-MM-dd') return '2024-01-15';
        return '';
      });

      const result = scriptContext.shouldSendDailySummary('America/New_York');
      expect(result).toBe(false);

      // Mock time at 21:00 - inside window
      mockUtilities.formatDate = jest.fn((date, tz, format) => {
        if (format === 'HH:mm') return '21:00';
        if (format === 'yyyy-MM-dd') return '2024-01-15';
        return '';
      });

      const result2 = scriptContext.shouldSendDailySummary('America/New_York');
      expect(result2).toBe(true); // Assuming no summary sent today

      // Mock time at 23:59 - inside window
      mockUtilities.formatDate = jest.fn((date, tz, format) => {
        if (format === 'HH:mm') return '23:59';
        if (format === 'yyyy-MM-dd') return '2024-01-15';
        return '';
      });

      const result3 = scriptContext.shouldSendDailySummary('America/New_York');
      expect(result3).toBe(true);

      // Mock time at 00:00 - outside window
      mockUtilities.formatDate = jest.fn((date, tz, format) => {
        if (format === 'HH:mm') return '00:00';
        if (format === 'yyyy-MM-dd') return '2024-01-16';
        return '';
      });

      const result4 = scriptContext.shouldSendDailySummary('America/New_York');
      expect(result4).toBe(false);
    });
  });

  describe('Date Tracking', () => {
    test('should only send summary once per day', () => {
      // Set last summary date to today
      mockPropertiesStore['passiveLastSummaryDate'] = '2024-01-15';
      
      // Mock time within window
      mockUtilities.formatDate = jest.fn((date, tz, format) => {
        if (format === 'HH:mm') return '21:00';
        if (format === 'yyyy-MM-dd') return '2024-01-15';
        return '';
      });

      const result = scriptContext.shouldSendDailySummary('America/New_York');
      expect(result).toBe(false); // Already sent today

      // Next day - should be able to send
      mockUtilities.formatDate = jest.fn((date, tz, format) => {
        if (format === 'HH:mm') return '21:00';
        if (format === 'yyyy-MM-dd') return '2024-01-16';
        return '';
      });

      const result2 = scriptContext.shouldSendDailySummary('America/New_York');
      expect(result2).toBe(true); // Different day
    });
  });

  describe('Missed Window Handling', () => {
    test('should accumulate results and send next day if window is missed', () => {
      // First day - outside window (e.g., 10:00)
      mockUtilities.formatDate = jest.fn((date, tz, format) => {
        if (format === 'HH:mm') return '10:00';
        if (format === 'yyyy-MM-dd') return '2024-01-15';
        return '';
      });

      global.processEmailsInBatches = jest.fn(() => ({
        mustDo: [{ subject: 'Task 1', keyAction: 'Do something' }],
        mustKnow: [],
        totalProcessed: 5
      }));
      scriptContext.processEmailsInBatches = global.processEmailsInBatches;

      const dateRange = { start: new Date('2024-01-15T10:00:00Z'), end: new Date('2024-01-15T11:00:00Z') };
      global.calculatePassiveWorkflowDateRange = jest.fn(() => dateRange);
      scriptContext.calculatePassiveWorkflowDateRange = global.calculatePassiveWorkflowDateRange;

      global.fetchEmailThreadsForPassiveWorkflow = jest.fn(() => [{ emails: [{ date: dateRange.start, id: 'msg1' }] }]);
      scriptContext.fetchEmailThreadsForPassiveWorkflow = global.fetchEmailThreadsForPassiveWorkflow;

      scriptContext.runPassiveWorkflow();

      // Verify results accumulated but not sent
      expect(emailSent).toHaveLength(0);
      const saved = JSON.parse(driveFiles['gmail-fast-read-accumulated-results.json'].content);
      expect(saved.mustDo).toHaveLength(1);

      // Next day - within window (21:00)
      mockUtilities.formatDate = jest.fn((date, tz, format) => {
        if (format === 'HH:mm') return '21:00';
        if (format === 'yyyy-MM-dd') return '2024-01-16';
        return '';
      });

      scriptContext.runPassiveWorkflow();

      // Verify summary was sent
      expect(emailSent.length).toBeGreaterThan(0);
      expect(emailSent[0].subject).toContain('Daily Summary');
      expect(emailSent[0].subject).toContain('2024-01-16');

      // Verify accumulated results were cleared
      expect(driveFiles['gmail-fast-read-accumulated-results.json']).toBeUndefined();
    });
  });

  describe('Cleanup After Send', () => {
    test('should clear accumulated results only after successful send', () => {
      // Set up accumulated results
      driveFiles['gmail-fast-read-accumulated-results.json'] = {
        getId: jest.fn(() => 'file-id'),
        getBlob: jest.fn(() => ({
          getDataAsString: jest.fn(() => JSON.stringify({
            mustDo: [{ subject: 'Task 1', keyAction: 'Do something' }],
            mustKnow: [],
            totalProcessed: 5,
            firstDate: '2024-01-15T10:00:00Z',
            lastDate: '2024-01-15T11:00:00Z'
          }))
        })),
        setContent: jest.fn(),
        setTrashed: jest.fn()
      };

      // Mock time within window
      mockUtilities.formatDate = jest.fn((date, tz, format) => {
        if (format === 'HH:mm') return '21:00';
        if (format === 'yyyy-MM-dd') return '2024-01-15';
        return '';
      });

      scriptContext.sendDailySummaryIfNeeded({
        addonName: 'Gmail Fast Reader',
        timeZone: 'America/New_York'
      });

      // Verify email was sent
      expect(emailSent.length).toBeGreaterThan(0);

      // Verify file was trashed (cleared)
      expect(driveFiles['gmail-fast-read-accumulated-results.json'].setTrashed).toHaveBeenCalled();

      // Verify last summary date was set
      expect(mockPropertiesService.setProperty).toHaveBeenCalledWith('passiveLastSummaryDate', '2024-01-15');
    });

    test('should keep accumulated results if send fails', () => {
      // Set up accumulated results
      driveFiles['gmail-fast-read-accumulated-results.json'] = {
        getId: jest.fn(() => 'file-id'),
        getBlob: jest.fn(() => ({
          getDataAsString: jest.fn(() => JSON.stringify({
            mustDo: [{ subject: 'Task 1', keyAction: 'Do something' }],
            mustKnow: [],
            totalProcessed: 5,
            firstDate: '2024-01-15T10:00:00Z',
            lastDate: '2024-01-15T11:00:00Z'
          }))
        })),
        setContent: jest.fn(),
        setTrashed: jest.fn()
      };

      // Mock email send to throw error
      mockGmailApp.sendEmail = jest.fn(() => {
        throw new Error('Send failed');
      });

      // Mock time within window
      mockUtilities.formatDate = jest.fn((date, tz, format) => {
        if (format === 'HH:mm') return '21:00';
        if (format === 'yyyy-MM-dd') return '2024-01-15';
        return '';
      });

      expect(() => {
        scriptContext.sendDailySummaryIfNeeded({
          addonName: 'Gmail Fast Reader',
          timeZone: 'America/New_York'
        });
      }).not.toThrow();

      // Verify file was NOT trashed
      expect(driveFiles['gmail-fast-read-accumulated-results.json'].setTrashed).not.toHaveBeenCalled();
    });
  });

  describe('Drive Operations', () => {
    test('should create file if it does not exist', () => {
      const result = scriptContext.loadAccumulatedResults();
      
      expect(mockDriveApp.getRootFolder).toHaveBeenCalled();
      expect(driveFiles['gmail-fast-read-accumulated-results.json']).toBeDefined();
      expect(result).toEqual({
        mustDo: [],
        mustKnow: [],
        totalProcessed: 0,
        firstDate: null,
        lastDate: null
      });
    });

    test('should read existing file', () => {
      driveFiles['gmail-fast-read-accumulated-results.json'] = {
        getId: jest.fn(() => 'file-id'),
        getBlob: jest.fn(() => ({
          getDataAsString: jest.fn(() => JSON.stringify({
            mustDo: [{ subject: 'Test' }],
            mustKnow: [],
            totalProcessed: 10,
            firstDate: '2024-01-15T10:00:00Z',
            lastDate: '2024-01-15T11:00:00Z'
          }))
        })),
        setContent: jest.fn(),
        setTrashed: jest.fn()
      };

      const result = scriptContext.loadAccumulatedResults();
      
      expect(result.mustDo).toHaveLength(1);
      expect(result.totalProcessed).toBe(10);
    });

    test('should save results to Drive', () => {
      const results = {
        mustDo: [{ subject: 'Task 1' }],
        mustKnow: [{ subject: 'Info 1' }],
        totalProcessed: 5,
        firstDate: '2024-01-15T10:00:00Z',
        lastDate: '2024-01-15T11:00:00Z'
      };

      scriptContext.saveAccumulatedResults(results);

      const saved = JSON.parse(driveFiles['gmail-fast-read-accumulated-results.json'].content);
      expect(saved.mustDo).toHaveLength(1);
      expect(saved.totalProcessed).toBe(5);
    });

    test('should clear file from Drive', () => {
      driveFiles['gmail-fast-read-accumulated-results.json'] = {
        getId: jest.fn(() => 'file-id'),
        getBlob: jest.fn(() => ({
          getDataAsString: jest.fn(() => '{}')
        })),
        setContent: jest.fn(),
        setTrashed: jest.fn()
      };

      scriptContext.clearAccumulatedResults();

      expect(driveFiles['gmail-fast-read-accumulated-results.json'].setTrashed).toHaveBeenCalled();
    });
  });
});

