/**
 * Gmail Fast Reader - Configuration management
 */

/**
 * Get all configuration from User Properties
 */
function getConfiguration() {
  const properties = PropertiesService.getUserProperties();
  
  // Helper to get property with default, only using default if property doesn't exist
  function getProp(key, defaultValue) {
    const value = properties.getProperty(key);
    return value !== null ? value : defaultValue;
  }
  
  // Helper to get boolean property with default
  function getBoolProp(key, defaultValue) {
    const value = properties.getProperty(key);
    if (value === null) return defaultValue;
    return value === 'true';
  }
  
  return {
    addonName: getProp('addonName', 'Gmail Fast Reader'),
    openaiApiKey: getProp('openaiApiKey', ''),
    timeZone: getProp('timeZone', 'Europe/Paris'), // CET timezone
    mustDoTopics: getProp('mustDoTopics', 'tax forms to file'),
    mustKnowTopics: getProp('mustKnowTopics', 'parent-teacher meetings at school\nschool trips'),
    mustDoOther: getBoolProp('mustDoOther', true),
    mustKnowOther: getBoolProp('mustKnowOther', false),
    unreadOnly: getBoolProp('unreadOnly', false),
    inboxOnly: getBoolProp('inboxOnly', true),
    mustDoLabel: getProp('mustDoLabel', 'TODO'),
    mustKnowLabel: getProp('mustKnowLabel', 'FYI')
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
    'timeZone': safeString(config.timeZone, 'Europe/Paris'), // CET timezone
    'mustDoTopics': safeString(config.mustDoTopics, ''),
    'mustKnowTopics': safeString(config.mustKnowTopics, ''),
    'mustDoOther': safeBoolean(config.mustDoOther) ? 'true' : 'false',
    'mustKnowOther': safeBoolean(config.mustKnowOther) ? 'true' : 'false',
    'unreadOnly': safeBoolean(config.unreadOnly) ? 'true' : 'false',
    'inboxOnly': safeBoolean(config.inboxOnly) ? 'true' : 'false',
    'mustDoLabel': safeString(config.mustDoLabel, ''),
    'mustKnowLabel': safeString(config.mustKnowLabel, '')
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
 * Check if onboarding is needed (first-time setup)
 * Returns true if openaiApiKey is empty, indicating first-time setup
 */
function needsOnboarding() {
  const properties = PropertiesService.getUserProperties();
  const openaiApiKey = properties.getProperty('openaiApiKey');
  return !openaiApiKey || openaiApiKey === '';
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
