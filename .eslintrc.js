module.exports = {
  env: {
    es2021: true,
    node: true,
    jest: true
  },
  extends: [
    'eslint:recommended'
  ],
  parserOptions: {
    ecmaVersion: 2021,
    sourceType: 'module'
  },
  rules: {
    'no-unused-vars': ['warn', { argsIgnorePattern: '^_' }],
    'no-console': 'off',
    'no-undef': 'warn', // Changed to warn for .gs files since functions are defined across files
    'indent': ['warn', 2], // Changed to warn to be less strict
    'linebreak-style': 'off', // Allow any line endings
    'quotes': ['warn', 'single'],
    'semi': ['warn', 'always'],
    'no-inner-declarations': 'off' // Allow function declarations in inner scopes
  },
  overrides: [
    {
      files: ['*.gs'],
      globals: {
        // Google Apps Script services
        'CardService': 'readonly',
        'PropertiesService': 'readonly',
        'GmailApp': 'readonly',
        'ScriptApp': 'readonly',
        'Session': 'readonly',
        'UrlFetchApp': 'readonly',
        'Logger': 'readonly',
        'Utilities': 'readonly',
        'console': 'readonly',
        // Project-specific globals (defined in other .gs files)
        'PROCESSING_STATUS': 'readonly',
        'PROCESSING_TIMEOUT_MS': 'readonly',
        'STATUS_REFRESH_INTERVAL_MS': 'readonly',
        'getConfiguration': 'readonly',
        'saveConfiguration': 'readonly',
        'isConfigurationComplete': 'readonly',
        'getTimezoneOptions': 'readonly',
        'getUserEmailAddress': 'readonly',
        'isProcessingRunning': 'readonly',
        'buildMainCard': 'readonly',
        'buildConfigurationCard': 'readonly',
        'buildActiveWorkflowCard': 'readonly',
        'buildProgressCard': 'readonly',
        'buildProgressCardWithStatusButton': 'readonly',
        'buildProgressCardWithAutoRefresh': 'readonly',
        'buildConfigSuccessCard': 'readonly',
        'buildErrorCard': 'readonly',
        'buildQuickScanCard': 'readonly',
        'buildLatestRunStatsCard': 'readonly',
        'processEmailsInBatches': 'readonly',
        'generateSummaryHTML': 'readonly',
        'shouldIgnoreEmail': 'readonly',
        'calculateDateRange': 'readonly',
        'fetchEmailThreadsFromGmail': 'readonly'
      },
      rules: {
        'no-undef': 'off' // Turn off for .gs files since functions are global
      }
    }
  ]
};

