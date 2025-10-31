/**
 * Gmail Fast Reader - Global Constants and Utility Functions
 */

// Processing timeout threshold (10 minutes in milliseconds)
const PROCESSING_TIMEOUT_MS = 10 * 60 * 1000;

// Auto-refresh interval for status checking (5 seconds in milliseconds)
const STATUS_REFRESH_INTERVAL_MS = 5 * 1000;

// Processing status constants
const PROCESSING_STATUS = {
  RUNNING: 'running',
  COMPLETED: 'completed',
  ERROR: 'error',
  TIMEOUT: 'timeout'
};

/**
 * Get the current user's email address
 */
function getUserEmailAddress() {
  try {
    return Session.getActiveUser().getEmail();
  } catch (error) {
    console.error('Error getting user email address:', error);
    return null;
  }
}
