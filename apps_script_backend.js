// Google Apps Script Web App for DDM Website Forms
// Deployment: Deploy as Web App, execute as Me, access to Anyone
// Sheet must have tabs: "Submissions" and "Playbook Leads"

const SHEET_ID = '1AeANyW_aGKrYk47qwBfXbu3rWiscBB4Ng3Bw85rOzRU';

// Spam detection patterns
const SPAM_PATTERNS = {
  // Known spam emails
  blockedEmails: [
    'de.x.ajizi.c36@gmail.com',
    'zekisuquc419@gmail.com'
  ],
  // Gibberish detection: names with random uppercase/lowercase mix
  gibberishNamePattern: /[a-z][A-Z]{3,}|[A-Z]{3,}[a-z]{3,}/,
  // Long random strings without spaces (likely gibberish)
  longGibberishPattern: /^[a-zA-Z0-9]{20,}$/,
  // Username-style names (e.g., RobertGak)
  usernameStylePattern: /^[A-Z][a-z]+[A-Z][a-z]+$/,
  // Random alphanumeric emails (suspicious)
  suspiciousEmailPattern: /^[a-z0-9]{15,}@/i
};

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: 'ok',
    message: 'DDM Form Handler is running'
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const params = e.parameter;
    const formType = params.formType || 'contact';

    const ss = SpreadsheetApp.openById(SHEET_ID);

    if (formType === 'playbook') {
      return handlePlaybookSubmission(ss, params);
    } else {
      return handleContactSubmission(ss, params);
    }
  } catch (error) {
    console.error('Form submission error:', error);
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Check if submission is spam based on multiple indicators
 * Returns: { isSpam: boolean, reasons: string[] }
 */
function checkSpamIndicators(params, formType) {
  const reasons = [];

  // 1. Honeypot check - if companyWebsite is filled, it's spam
  if (params.companyWebsite && params.companyWebsite.trim() !== '') {
    reasons.push('Honeypot field filled');
  }

  const email = (params.email || '').toLowerCase().trim();
  const firstName = (params.firstName || '').trim();
  const lastName = (params.lastName || '').trim();
  const organization = (params.organization || '').trim();
  const message = (params.message || '').trim();

  // 2. Blocked email check
  if (SPAM_PATTERNS.blockedEmails.includes(email)) {
    reasons.push('Known spam email address');
  }

  // 3. Gibberish name detection
  const fullName = firstName + ' ' + lastName;
  if (SPAM_PATTERNS.gibberishNamePattern.test(fullName)) {
    reasons.push('Gibberish name pattern detected');
  }
  if (SPAM_PATTERNS.longGibberishPattern.test(firstName) || SPAM_PATTERNS.longGibberishPattern.test(lastName)) {
    reasons.push('Long random character name');
  }
  if (SPAM_PATTERNS.usernameStylePattern.test(firstName) || SPAM_PATTERNS.usernameStylePattern.test(lastName)) {
    reasons.push('Username-style name pattern');
  }

  // 4. Gibberish organization detection
  if (organization.length > 15 && SPAM_PATTERNS.longGibberishPattern.test(organization)) {
    reasons.push('Gibberish organization name');
  }
  if (SPAM_PATTERNS.gibberishNamePattern.test(organization)) {
    reasons.push('Random case organization name');
  }

  // 5. Suspicious email pattern
  if (SPAM_PATTERNS.suspiciousEmailPattern.test(email)) {
    reasons.push('Suspicious email format');
  }

  // 6. Gibberish message detection (for contact form)
  if (formType === 'contact' && message) {
    // Very short message that's just random letters (no real words)
    if (message.length < 50 && /^[a-zA-Z\s]+$/.test(message) && !/[aeiou]{2,}/i.test(message)) {
      reasons.push('Gibberish message content');
    }
  }

  return {
    isSpam: reasons.length > 0,
    reasons: reasons
  };
}

function handleContactSubmission(ss, params) {
  const sheet = ss.getSheetByName('Submissions');
  if (!sheet) {
    throw new Error('Submissions sheet not found');
  }

  // Check for spam
  const spamCheck = checkSpamIndicators(params, 'contact');

  const timestamp = new Date();
  // Match the sheet structure: Timestamp, Name, Email, Phone, Organization, Service Interest, Message, Page Source, Lead Status, Notes
  const row = [
    timestamp,
    (params.firstName || '') + ' ' + (params.lastName || ''), // Full name
    params.email || '',
    params.phone || '',
    params.organization || '',
    params.interest || '',  // Service Interest
    params.message || '',
    'contact.html', // Page Source
    spamCheck.isSpam ? 'Possible Spam' : 'New Lead',
    spamCheck.isSpam ? 'Flagged: ' + spamCheck.reasons.join(', ') : ''
  ];

  sheet.appendRow(row);

  // Only send email notifications for non-spam submissions
  if (!spamCheck.isSpam) {
    sendContactNotification(params);
  } else {
    console.log('Spam submission detected and logged without notification. Reasons: ' + spamCheck.reasons.join(', '));
  }

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    message: 'Contact form submitted successfully'
  })).setMimeType(ContentService.MimeType.JSON);
}

function handlePlaybookSubmission(ss, params) {
  const sheet = ss.getSheetByName('Playbook Leads');
  if (!sheet) {
    throw new Error('Playbook Leads sheet not found');
  }

  // Check for spam
  const spamCheck = checkSpamIndicators(params, 'playbook');

  const timestamp = new Date();
  const row = [
    timestamp,
    params.firstName || '',
    params.lastName || '',
    params.email || '',
    params.organization || '',
    params.role || '',
    'Playbook Page',
    spamCheck.isSpam ? 'Possible Spam' : 'New',
    spamCheck.isSpam ? 'Flagged: ' + spamCheck.reasons.join(', ') : ''
  ];

  sheet.appendRow(row);

  // Only send email notifications for non-spam submissions
  if (!spamCheck.isSpam) {
    sendPlaybookNotification(params);
  } else {
    console.log('Spam submission detected and logged without notification. Reasons: ' + spamCheck.reasons.join(', '));
  }

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    message: 'Playbook request submitted successfully'
  })).setMimeType(ContentService.MimeType.JSON);
}

function sendContactNotification(params) {
  const subject = `New Contact Form Submission: ${params.firstName} ${params.lastName}`;
  const body = `
New contact form submission:

Name: ${params.firstName} ${params.lastName}
Email: ${params.email}
Phone: ${params.phone || 'Not provided'}
Organization: ${params.organization}
Type: ${params.type || 'Not selected'}
Interest: ${params.interest || 'Not selected'}

Message:
${params.message || 'No message provided'}

---
Submitted via dixiedigital.media contact form
  `.trim();

  const recipients = [
    'dixie@dixiedigital.media',
    'dixiedigitalmedia@gmail.com'
  ];

  recipients.forEach(email => {
    try {
      GmailApp.sendEmail(email, subject, body);
    } catch (e) {
      console.error('Failed to send to ' + email + ': ' + e);
    }
  });
}

function sendPlaybookNotification(params) {
  const subject = `New Playbook Download Request: ${params.firstName} ${params.lastName}`;
  const body = `
New Digital Heritage Playbook request:

Name: ${params.firstName} ${params.lastName}
Email: ${params.email}
Organization: ${params.organization}
Role: ${params.role || 'Not provided'}

Action needed: Send the Playbook PDF to this lead.

---
Submitted via dixiedigital.media playbook page
  `.trim();

  const recipients = [
    'dixie@dixiedigital.media',
    'dixiedigitalmedia@gmail.com'
  ];

  recipients.forEach(email => {
    try {
      GmailApp.sendEmail(email, subject, body);
    } catch (e) {
      console.error('Failed to send to ' + email + ': ' + e);
    }
  });
}

/**
 * Test function to trigger Gmail authorization
 * Run this function in the Apps Script editor to authorize Gmail permissions
 */
function testPlaybookNotificationAuthorization() {
  const sampleData = {
    firstName: 'Test',
    lastName: 'Authorization',
    email: 'test@example.com',
    organization: 'Test Organization',
    role: 'board-member'
  };

  console.log('Testing Playbook notification authorization...');
  sendPlaybookNotification(sampleData);
  console.log('Authorization test complete. Check your email.');
}
