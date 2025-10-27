/**
 * Gmail Fast Reader - Generate and send summary emails
 */

/**
 * Build summary card for display
 */
function buildSummaryCard(results) {
  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Email Analysis Complete'))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText(`✅ ${results.message}`)));
  
  if (results.mustDo.length > 0) {
    card.addSection(CardService.newCardSection()
      .setHeader('📋 I Must Do')
      .addWidget(CardService.newTextParagraph()
        .setText(formatItemsForCard(results.mustDo))));
  }
  
  if (results.mustKnow.length > 0) {
    card.addSection(CardService.newCardSection()
      .setHeader('📰 I Must Know')
      .addWidget(CardService.newTextParagraph()
        .setText(formatItemsForCard(results.mustKnow))));
  }
  
  if (results.mustDo.length === 0 && results.mustKnow.length === 0) {
    card.addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText('No relevant emails found in the selected time range.')));
  }
  
  card.addSection(CardService.newCardSection()
    .addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText('📧 Send Summary Email')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('sendSummaryEmail')
          .setParameters({ results: JSON.stringify(results) })))
      .addButton(CardService.newTextButton()
        .setText('🏠 Back to Main')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('buildMainCard')))));
  
  return card.build();
}

/**
 * Send summary email
 */
function sendSummaryEmail(e) {
  try {
    const results = JSON.parse(e.parameters.results);
    const config = getConfiguration();
    
    const htmlContent = generateSummaryHTML(results, config);
    const subject = `${config.addonName} - Daily Summary - ${new Date().toLocaleDateString()}`;
    
    GmailApp.sendEmail(
      Session.getActiveUser().getEmail(),
      subject,
      '', // Plain text version (empty, we're using HTML)
      {
        htmlBody: htmlContent,
        name: config.addonName
      }
    );
    
    return buildConfigSuccessCard(); // Reuse success card for email sent
    
  } catch (error) {
    console.error('Error sending summary email:', error);
    return buildErrorCard('Failed to send summary email: ' + error.message);
  }
}

/**
 * Generate HTML content for summary email
 */
function generateSummaryHTML(results, config) {
  const now = new Date();
  const today = now.toISOString().split('T')[0];
  const tomorrow = new Date(now.getTime() + 24 * 60 * 60 * 1000).toISOString().split('T')[0];
  
  let html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <style>
        body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; max-width: 800px; margin: 0 auto; padding: 20px; }
        .header { background: #f8f9fa; padding: 20px; border-radius: 8px; margin-bottom: 20px; }
        .section { margin-bottom: 30px; }
        .section h2 { color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 10px; }
        .item { background: #fff; border: 1px solid #e1e8ed; border-radius: 6px; padding: 15px; margin-bottom: 10px; }
        .urgent { border-left: 4px solid #e74c3c; background: #fdf2f2; }
        .item-header { font-weight: bold; color: #2c3e50; margin-bottom: 5px; }
        .item-meta { font-size: 0.9em; color: #7f8c8d; margin-bottom: 8px; }
        .item-content { margin-bottom: 5px; }
        .date { color: #e67e22; font-weight: bold; }
        .urgent-date { color: #e74c3c; font-weight: bold; }
        .outdated-date { color: #95a5a6; font-weight: normal; }
        .footer { margin-top: 30px; padding-top: 20px; border-top: 1px solid #e1e8ed; font-size: 0.9em; color: #7f8c8d; }
      </style>
    </head>
    <body>
      <div class="header">
        <h1>${config.addonName} - Daily Summary</h1>
        <p>Generated on ${now.toLocaleDateString()} at ${now.toLocaleTimeString()}</p>
        <p>Total emails processed: ${results.totalProcessed}</p>
      </div>
  `;
  
  // I Must Do section
  if (results.mustDo.length > 0) {
    html += `
      <div class="section">
        <h2>📋 I Must Do (${results.mustDo.length} items)</h2>
    `;
    
    // Sort by date
    const sortedMustDo = results.mustDo.sort((a, b) => {
      if (!a.date && !b.date) return 0;
      if (!a.date) return 1;
      if (!b.date) return -1;
      return new Date(a.date) - new Date(b.date);
    });
    
    sortedMustDo.forEach(item => {
      const isUrgent = item.date && (item.date === today || item.date === tomorrow);
      const isOutdated = item.date && item.date < today;
      const dateClass = isUrgent ? 'urgent-date' : (isOutdated ? 'outdated-date' : 'date');
      const itemClass = isUrgent ? 'item urgent' : 'item';
      
      html += `
        <div class="${itemClass}">
          <div class="item-header">${item.subject}</div>
          <div class="item-meta">From: ${item.sender} | Topic: ${item.topic}</div>
          <div class="item-content"><strong>Action:</strong> ${item.keyAction}</div>
          ${item.date ? `<div class="${dateClass}">📅 ${item.date}${isUrgent ? ' (URGENT!)' : (isOutdated ? ' (OUTDATED)' : '')}</div>` : ''}
        </div>
      `;
    });
    
    html += '</div>';
  }
  
  // I Must Know section
  if (results.mustKnow.length > 0) {
    html += `
      <div class="section">
        <h2>📰 I Must Know (${results.mustKnow.length} items)</h2>
    `;
    
    // Sort by date
    const sortedMustKnow = results.mustKnow.sort((a, b) => {
      if (!a.date && !b.date) return 0;
      if (!a.date) return 1;
      if (!b.date) return -1;
      return new Date(a.date) - new Date(b.date);
    });
    
    sortedMustKnow.forEach(item => {
      html += `
        <div class="item">
          <div class="item-header">${item.subject}</div>
          <div class="item-meta">From: ${item.sender} | Topic: ${item.topic}</div>
          <div class="item-content"><strong>Key Info:</strong> ${item.keyKnowledge}</div>
          ${item.date ? `<div class="date">📅 ${item.date}</div>` : ''}
        </div>
      `;
    });
    
    html += '</div>';
  }
  
  if (results.mustDo.length === 0 && results.mustKnow.length === 0) {
    html += `
      <div class="section">
        <h2>📭 No Relevant Emails Found</h2>
        <p>No emails in the selected time range matched your configured topics of interest.</p>
      </div>
    `;
  }
  
  html += `
      <div class="footer">
        <p>This summary was generated by ${config.addonName}. To configure your topics and preferences, open the Gmail Fast Reader add-on.</p>
      </div>
    </body>
    </html>
  `;
  
  return html;
}

/**
 * Format items for card display
 */
function formatItemsForCard(items) {
  if (items.length === 0) return 'No items found.';
  
  return items.map((item, index) => {
    const dateStr = item.date ? ` (${item.date})` : '';
    
    return `${index + 1}. ${item.subject}${dateStr}\n   From: ${item.sender}\n   ${item.keyAction || item.keyKnowledge}`;
  }).join('\n\n');
}
