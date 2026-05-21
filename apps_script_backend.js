// Google Apps Script Web App for DDM Website Forms
// Deployment: Deploy as Web App, execute as Me, access to Anyone
// Sheet must have tabs: "Submissions" and "Playbook Leads"

const SHEET_ID = '1AeANyW_aGKrYk47qwBfXbu3rWiscBB4Ng3Bw85rOzRU';

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

function handleContactSubmission(ss, params) {
  const sheet = ss.getSheetByName('Submissions');
  if (!sheet) {
    throw new Error('Submissions sheet not found');
  }
  
  const timestamp = new Date();
  const row = [
    timestamp,
    params.firstName || '',
    params.lastName || '',
    params.email || '',
    params.phone || '',
    params.organization || '',
    params.type || '',
    params.interest || '',
    params.message || ''
  ];
  
  sheet.appendRow(row);
  
  // Send notification emails
  sendContactNotification(params);
  
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
  
  const timestamp = new Date();
  const row = [
    timestamp,
    params.firstName || '',
    params.lastName || '',
    params.email || '',
    params.organization || '',
    params.role || '',
    'Playbook Page',
    'New',
    ''
  ];
  
  sheet.appendRow(row);
  
  // Send notification emails
  sendPlaybookNotification(params);
  
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
