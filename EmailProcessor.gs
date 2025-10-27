/**
 * Gmail Fast Reader - Email scanning and OpenAI analysis logic
 */

// Token limit for GPT-5-nano (conservative estimate)
const MAX_TOKENS = 200000;
const TOKENS_PER_CHAR = 0.25; // Rough estimate for English text

/**
 * Process emails for a given time range
 */
function processEmails(timeRange) {
  try {
    // Get configuration
    const config = getConfiguration();
    if (!config.openaiApiKey) {
      throw new Error('OpenAI API key not configured');
    }
    
    // Calculate date range
    const dateRange = calculateDateRange(timeRange);
    
    // Fetch email threads from Gmail (prioritizing threads)
    const emailThreads = fetchEmailThreadsFromGmail(dateRange);
    
    if (emailThreads.length === 0) {
      return {
        mustDo: [],
        mustKnow: [],
        totalProcessed: 0,
        message: 'No emails found in the selected time range.'
      };
    }
    
    // Process emails in batches that respect token limits
    const results = processEmailsInBatches(emailThreads, config);
    
    return {
      mustDo: results.mustDo || [],
      mustKnow: results.mustKnow || [],
      totalProcessed: results.totalProcessed,
      message: `Processed ${results.totalProcessed} emails in ${results.batchesProcessed} batches.`
    };
    
  } catch (error) {
    console.error('Error processing emails:', error);
    throw new Error('Failed to process emails: ' + error.message);
  }
}

/**
 * Calculate date range based on time range string
 */
function calculateDateRange(timeRange) {
  const now = new Date();
  let startDate;
  
  switch (timeRange) {
    case '6hours':
      startDate = new Date(now.getTime() - 6 * 60 * 60 * 1000);
      break;
    case '12hours':
      startDate = new Date(now.getTime() - 12 * 60 * 60 * 1000);
      break;
    case '1day':
      startDate = new Date(now.getTime() - 24 * 60 * 60 * 1000);
      break;
    case '2days':
      startDate = new Date(now.getTime() - 2 * 24 * 60 * 60 * 1000);
      break;
    case '7days':
      startDate = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
      break;
    default:
      startDate = new Date(now.getTime() - 24 * 60 * 60 * 1000);
  }
  
  return {
    start: startDate,
    end: now
  };
}

/**
 * Fetch email threads from Gmail API (prioritizing threads)
 */
function fetchEmailThreadsFromGmail(dateRange) {
  try {
    const query = `after:${Math.floor(dateRange.start.getTime() / 1000)} before:${Math.floor(dateRange.end.getTime() / 1000)}`;
    
    const threads = GmailApp.search(query, 0, 100); // Get more threads initially
    const emailThreads = [];
    
    threads.forEach(thread => {
      const messages = thread.getMessages();
      const threadEmails = [];
      
      messages.forEach(message => {
        if (message.getDate() >= dateRange.start && message.getDate() <= dateRange.end) {
          threadEmails.push({
            id: message.getId(),
            subject: message.getSubject(),
            sender: message.getFrom(),
            date: message.getDate(),
            body: message.getPlainBody(),
            snippet: message.getPlainBody().substring(0, 200)
          });
        }
      });
      
      if (threadEmails.length > 0) {
        // Sort emails in thread by date (oldest first)
        threadEmails.sort((a, b) => a.date - b.date);
        emailThreads.push({
          threadId: thread.getId(),
          subject: thread.getFirstMessageSubject(),
          emails: threadEmails,
          totalEmails: threadEmails.length,
          latestDate: threadEmails[threadEmails.length - 1].date
        });
      }
    });
    
    // Sort threads by latest email date (most recent first)
    emailThreads.sort((a, b) => b.latestDate - a.latestDate);
    
    return emailThreads;
  } catch (error) {
    console.error('Error fetching email threads:', error);
    throw new Error('Failed to fetch email threads from Gmail');
  }
}

/**
 * Process emails in batches that respect token limits
 */
function processEmailsInBatches(emailThreads, config) {
  const allResults = {
    mustDo: [],
    mustKnow: [],
    totalProcessed: 0,
    batchesProcessed: 0
  };
  
  // Calculate total threads and messages for progress tracking
  const totalThreads = emailThreads.length;
  const totalMessages = emailThreads.reduce((total, thread) => total + thread.emails.length, 0);
  
  // Update initial progress
  updateProcessingProgress(0, totalThreads, 0, totalMessages, 'Starting email analysis...');
  
  let currentBatch = [];
  let currentTokenCount = 0;
  let processedThreads = 0;
  let processedMessages = 0;
  
  for (let i = 0; i < emailThreads.length; i++) {
    const thread = emailThreads[i];
    
    // Update current processing status - show which thread we're about to process
    updateProcessingProgress(processedThreads, totalThreads, processedMessages, totalMessages, 
      `Processing thread ${i + 1}/${totalThreads}: ${thread.subject}`);
    
    // Calculate tokens for this thread
    const threadTokens = estimateTokensForThread(thread);
    
    // If adding this thread would exceed the limit, process current batch first
    if (currentTokenCount + threadTokens > MAX_TOKENS && currentBatch.length > 0) {
      // Update status before processing batch
      updateProcessingProgress(processedThreads, totalThreads, processedMessages, totalMessages, 
        `Processing batch ${allResults.batchesProcessed + 1} with ${currentBatch.length} threads...`);
      
      const batchResults = analyzeEmailsWithOpenAI(currentBatch, config);
      mergeResults(allResults, batchResults);
      
      // Update progress after batch processing
      processedThreads += currentBatch.length;
      processedMessages += currentBatch.reduce((total, t) => total + t.emails.length, 0);
      updateProcessingProgress(processedThreads, totalThreads, processedMessages, totalMessages, 
        `Completed batch ${allResults.batchesProcessed + 1} - ${processedThreads}/${totalThreads} threads processed`);
      
      // Reset for next batch
      currentBatch = [];
      currentTokenCount = 0;
      allResults.batchesProcessed++;
    }
    
    // Add thread to current batch
    currentBatch.push(thread);
    currentTokenCount += threadTokens;
    
    // If this single thread exceeds the limit, process it alone
    if (threadTokens > MAX_TOKENS) {
      // Update status before processing single thread
      updateProcessingProgress(processedThreads, totalThreads, processedMessages, totalMessages, 
        `Processing large thread: ${thread.subject}`);
      
      const batchResults = analyzeEmailsWithOpenAI([thread], config);
      mergeResults(allResults, batchResults);
      allResults.batchesProcessed++;
      
      // Update progress after single thread processing
      processedThreads += 1;
      processedMessages += thread.emails.length;
      updateProcessingProgress(processedThreads, totalThreads, processedMessages, totalMessages, 
        `Completed thread: ${thread.subject} - ${processedThreads}/${totalThreads} threads processed`);
      
      // Reset for next batch
      currentBatch = [];
      currentTokenCount = 0;
    }
  }
  
  // Process remaining emails in final batch
  if (currentBatch.length > 0) {
    updateProcessingProgress(processedThreads, totalThreads, processedMessages, totalMessages, 
      `Processing final batch with ${currentBatch.length} threads...`);
    
    const batchResults = analyzeEmailsWithOpenAI(currentBatch, config);
    mergeResults(allResults, batchResults);
    allResults.batchesProcessed++;
    
    // Update final progress
    processedThreads += currentBatch.length;
    processedMessages += currentBatch.reduce((total, t) => total + t.emails.length, 0);
    updateProcessingProgress(processedThreads, totalThreads, processedMessages, totalMessages, 
      `Analysis complete! Processed ${processedThreads}/${totalThreads} threads and ${processedMessages}/${totalMessages} messages.`);
  }
  
  return allResults;
}

/**
 * Update processing progress in user properties
 */
function updateProcessingProgress(processedThreads, totalThreads, processedMessages, totalMessages, currentActivity) {
  const properties = PropertiesService.getUserProperties();
  
  // Calculate percentages
  const threadProgress = totalThreads > 0 ? Math.round((processedThreads / totalThreads) * 100) : 0;
  const messageProgress = totalMessages > 0 ? Math.round((processedMessages / totalMessages) * 100) : 0;
  
  // Update progress properties
  properties.setProperties({
    'processingProgress': threadProgress.toString(),
    'processingMessage': currentActivity,
    'processedThreads': processedThreads.toString(),
    'totalThreads': totalThreads.toString(),
    'processedMessages': processedMessages.toString(),
    'totalMessages': totalMessages.toString()
  });
}

/**
 * Estimate token count for a thread
 */
function estimateTokensForThread(thread) {
  let totalTokens = 0;
  
  // Base prompt tokens (rough estimate)
  totalTokens += 500;
  
  // Add tokens for each email in thread
  thread.emails.forEach(email => {
    const emailText = `${email.subject} ${email.sender} ${email.body}`;
    totalTokens += Math.ceil(emailText.length * TOKENS_PER_CHAR);
  });
  
  return totalTokens;
}

/**
 * Merge results from multiple batches
 */
function mergeResults(allResults, batchResults) {
  if (batchResults.mustDo) {
    allResults.mustDo = allResults.mustDo.concat(batchResults.mustDo);
  }
  if (batchResults.mustKnow) {
    allResults.mustKnow = allResults.mustKnow.concat(batchResults.mustKnow);
  }
  allResults.totalProcessed += batchResults.totalProcessed || 0;
}

/**
 * Analyze emails using OpenAI API
 */
function analyzeEmailsWithOpenAI(emailThreads, config) {
  try {
    const prompt = buildAnalysisPrompt(emailThreads, config);
    
    const response = callOpenAIAPI(prompt, config.openaiApiKey);
    
    const results = parseOpenAIResponse(response);
    
    // Count total emails processed
    results.totalProcessed = emailThreads.reduce((total, thread) => total + thread.emails.length, 0);
    
    return results;
  } catch (error) {
    console.error('Error analyzing emails with OpenAI:', error);
    throw new Error('Failed to analyze emails with OpenAI: ' + error.message);
  }
}

/**
 * Build prompt for OpenAI analysis
 */
function buildAnalysisPrompt(emailThreads, config) {
  const mustDoTopics = config.mustDoTopics.split('\n').filter(t => t.trim());
  const mustKnowTopics = config.mustKnowTopics.split('\n').filter(t => t.trim());
  
  // Format topics as bullet lists
  const mustDoTopicsList = mustDoTopics.length > 0 
    ? mustDoTopics.map(topic => `- ${topic}`).join('\n')
    : 'None specified';
    
  const mustKnowTopicsList = mustKnowTopics.length > 0 
    ? mustKnowTopics.map(topic => `- ${topic}`).join('\n')
    : 'None specified';
  
  // Build messages array with separate user prompt for each thread
  const messages = [
    {
      role: 'developer',
      content: `You are a helpful AI assistant that analyzes, filters, and categorizes emails. 
      Your job is to reduce the reading effort for the user by identifying important emails, key knowledge and actions, and dates.
      Always respond with valid JSON.

TASK: Categorize email threads into "I must do" and "I must know" categories.

For each email thread, determine:
1. If it fits "I must do" category (actionable items, deadlines, tasks)
2. If it fits "I must know" category (informational, updates, news)
3. Extract the earliest key action or knowledge from the entire thread

IMPORTANT: Not every email must make it to the output. Skipping emails is okay and even desired when they don't match any topic. Only include emails that are truly relevant to the user's specified topics.

MANDATE: You MUST respond with valid JSON in exactly this structure:
{
  "mustDo": [
    {
      "emailId": "email_id",
      "subject": "email subject",
      "sender": "sender email",
      "keyAction": "what the user must do",
      "date": "YYYY-MM-DD or null if no date found",
      "topic": "which topic it matches"
    }
  ],
  "mustKnow": [
    {
      "emailId": "email_id", 
      "subject": "email subject",
      "sender": "sender email",
      "keyKnowledge": "what the user must know",
      "date": "YYYY-MM-DD or null if no date found",
      "topic": "which topic it matches"
    }
  ]
}

EXAMPLE: If an email is about a payment due next week, it should go in "mustDo" with keyAction like "Pay invoice #123 by 2024-01-15" and topic matching one of the user's specified topics.`
    },
    {
      role: 'user',
      content: `MY "I MUST DO" TOPICS:
${mustDoTopicsList}

${config.mustDoOther ? 'If an email does not fit any of my topics, but feels important, you may create a new topic. If it is not important enough, do not include it in any topic, skip it instead.' : 'If an email does not fit any of my topics, do not include it in any topic, skip it instead.'}`
    },
    {
      role: 'user',
      content: `MY "I MUST KNOW" TOPICS:
${mustKnowTopicsList}

${config.mustKnowOther ? 'If an email does not fit any of my topics, but feels important, you may create a new topic. If it is not important enough, do not include it in any topic, skip it instead.' : 'If an email does not fit any of my topics, do not include it in any topic, skip it instead.'}`
    }
  ];
  
  // Add each thread as a separate user message
  emailThreads.forEach((thread, threadIndex) => {
    let threadData = `THREAD ${threadIndex + 1} (${thread.emails.length} emails):
Subject: ${thread.subject}
Thread ID: ${thread.threadId}
---`;
    
    thread.emails.forEach((email, emailIndex) => {
      threadData += `\nEmail ${emailIndex + 1}:
From: ${email.sender}
Date: ${email.date.toISOString().split('T')[0]}
Body: ${email.body.substring(0, 800)}${email.body.length > 800 ? '...' : ''}
---`;
    });
    
    messages.push({
      role: 'user',
      content: threadData
    });
  });
  
  return messages;
}

/**
 * Call OpenAI API
 */
function callOpenAIAPI(prompt, apiKey) {
  const url = 'https://api.openai.com/v1/chat/completions';
  
  const payload = {
    model: 'gpt-5-nano',
    messages: prompt,
    max_completion_tokens: 40000
  };
  
  const options = {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${apiKey}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload)
  };
  
  // Log the full request
  console.log('=== OPENAI API REQUEST ===');
  console.log('URL:', url);
  console.log('Headers:', JSON.stringify(options.headers, null, 2));
  console.log('Payload:', JSON.stringify(payload, null, 2));
  
  const response = UrlFetchApp.fetch(url, options);
  const responseText = response.getContentText();
  
  // Log the full response
  console.log('=== OPENAI API RESPONSE ===');
  console.log('Status Code:', response.getResponseCode());
  console.log('Response Headers:', JSON.stringify(response.getHeaders(), null, 2));
  console.log('Response Body:', responseText);
  
  const responseData = JSON.parse(responseText);
  
  if (responseData.error) {
    console.error('OpenAI API Error:', responseData.error);
    throw new Error('OpenAI API error: ' + responseData.error.message);
  }
  
  return responseData.choices[0].message.content;
}

/**
 * Parse OpenAI response
 */
function parseOpenAIResponse(response) {
  try {
    // Extract JSON from response (in case there's extra text)
    const jsonMatch = response.match(/\{[\s\S]*\}/);
    if (!jsonMatch) {
      throw new Error('No JSON found in OpenAI response');
    }
    
    const parsed = JSON.parse(jsonMatch[0]);
    
    // Validate structure
    if (!parsed.mustDo || !parsed.mustKnow) {
      throw new Error('Invalid response structure from OpenAI');
    }
    
    return parsed;
  } catch (error) {
    console.error('Error parsing OpenAI response:', error);
    console.error('Response was:', response);
    throw new Error('Failed to parse OpenAI response: ' + error.message);
  }
}
