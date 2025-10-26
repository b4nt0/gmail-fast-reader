/**
 * Gmail Fast Reader - Configuration management
 */

/**
 * Get all configuration from User Properties
 */
function getConfiguration() {
  const properties = PropertiesService.getUserProperties();
  
  return {
    addonName: properties.getProperty('addonName') || 'Gmail Fast Reader',
    openaiApiKey: properties.getProperty('openaiApiKey') || '',
    timeZone: properties.getProperty('timeZone') || 'America/New_York',
    mustDoTopics: properties.getProperty('mustDoTopics') || '',
    mustKnowTopics: properties.getProperty('mustKnowTopics') || '',
    mustDoOther: properties.getProperty('mustDoOther') === 'true',
    mustKnowOther: properties.getProperty('mustKnowOther') === 'true'
  };
}

/**
 * Save configuration to User Properties
 */
function saveConfiguration(config) {
  const properties = PropertiesService.getUserProperties();
  
  // Helper function to safely convert to string
  function safeString(value, defaultValue = '') {
    if (value === null || value === undefined) return defaultValue;
    if (Array.isArray(value)) return value.join('\n');
    if (typeof value === 'object') return JSON.stringify(value);
    return String(value);
  }
  
  // Helper function to safely convert to boolean
  function safeBoolean(value) {
    if (value === null || value === undefined) return false;
    if (typeof value === 'boolean') return value;
    if (typeof value === 'string') return value === 'true';
    if (Array.isArray(value)) return value.includes('true');
    return false;
  }
  
  properties.setProperties({
    'addonName': safeString(config.addonName, 'Gmail Fast Reader'),
    'openaiApiKey': safeString(config.openaiApiKey, ''),
    'timeZone': safeString(config.timeZone, 'America/New_York'),
    'mustDoTopics': safeString(config.mustDoTopics, ''),
    'mustKnowTopics': safeString(config.mustKnowTopics, ''),
    'mustDoOther': safeBoolean(config.mustDoOther) ? 'true' : 'false',
    'mustKnowOther': safeBoolean(config.mustKnowOther) ? 'true' : 'false'
  });
}

/**
 * Check if configuration is complete
 */
function isConfigurationComplete() {
  const config = getConfiguration();
  return config.openaiApiKey && 
         (config.mustDoTopics || config.mustDoOther) && 
         (config.mustKnowTopics || config.mustKnowOther);
}

/**
 * Get timezone options for dropdown
 */
function getTimezoneOptions() {
  return [
    { label: 'Eastern Time (ET)', value: 'America/New_York' },
    { label: 'Central Time (CT)', value: 'America/Chicago' },
    { label: 'Mountain Time (MT)', value: 'America/Denver' },
    { label: 'Pacific Time (PT)', value: 'America/Los_Angeles' },
    { label: 'UTC', value: 'UTC' },
    { label: 'London (GMT)', value: 'Europe/London' },
    { label: 'Paris (CET)', value: 'Europe/Paris' },
    { label: 'Tokyo (JST)', value: 'Asia/Tokyo' },
    { label: 'Sydney (AEST)', value: 'Australia/Sydney' }
  ];
}
