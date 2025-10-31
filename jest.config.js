module.exports = {
  testEnvironment: 'node',
  testMatch: [
    '**/tests/**/*.test.js',
    '**/*.test.js'
  ],
  collectCoverageFrom: [
    '**/*.gs',
    '!**/node_modules/**',
    '!**/tests/**',
    '!**/__mocks__/**'
  ],
  moduleFileExtensions: ['js', 'gs'],
  setupFilesAfterEnv: ['<rootDir>/tests/setup.js'],
  testTimeout: 10000,
  verbose: true
};

