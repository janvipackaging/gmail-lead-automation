const {google} = require('googleapis');
const cheerio = require('cheerio');
const {authorize} = require('./authenticate');
const fs = require('fs');

// --- 1. CONFIGURATION ---
const SPREADSHEET_ID = '1Xo46YyTvM8CohicxGR2mVAvi94TLs8ab1eT9aMzWlMs'; 
const SHEET_NAME = 'Sheet1'; 
const LAST_RUN_FILE = 'last_run.txt'; 
const MAX_CELL_CHARS = 49999;

// --- SAFETY FUNCTION ---
function truncateString(str) {
  if (str && str.length > MAX_CELL_CHARS) {
    console.warn(`Warning: Data was truncated.`);
    return str.substring(0, MAX_CELL_CHARS) + "\n... (truncated)";
  }
  return str;
}

/**
 * Reads the timestamp of the last successful run from a file.
 */
function getLastRunTime() {
    try {
        if (fs.existsSync(LAST_RUN_FILE)) {
            const lastRunTime = fs.readFileSync(LAST_RUN_FILE, 'utf8').trim();
            const date = new Date(lastRunTime);
            
            const year = date.getFullYear();
            const month = String(date.getMonth() + 1).padStart(2, '0');
            const day = String(date.getDate()).padStart(2, '0');
            return `${year}/${month}/${day}`;
        }
    } catch (err) {
        console.error("Error reading last_run.txt:", err);
    }
    // Fallback: If no file exists, search for the last 48 hours
    const fallbackDate = new Date(Date.now() - 48 * 60 * 60 * 1000);
    const year = fallbackDate.getFullYear();
    const month = String(fallbackDate.getMonth() + 1).padStart(2, '0');
    const day = String(fallbackDate.getDate()).padStart(2, '0');
    return `${year}/${month}/${day}`;
}

/**
 * Saves the current time to the file for the next run.
 */
function saveCurrentRunTime() {
    try {
        const currentTime = new Date().toISOString();
        fs.writeFileSync(LAST_RUN_FILE, currentTime);
        console.log(`Updated last run time to ${currentTime}`);
    } catch (err) {
        console.error("Error writing last_run.txt:", err);
    }
}

/**
 * Fetches all existing Message IDs from the Unique ID column (J).
 */
async function getExistingMessageIds(sheets, spreadsheetId, sheetName) {
    try {
        const range = `${sheetName}!J:J`; // Check only the Unique Message ID column (J)
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: spreadsheetId,
            range: range,
        });
        
        return new Set((response.data.values || []).flat().filter(id => id && id !== 'Unique Message ID'));
    } catch (error) {
        console.error("Error fetching existing IDs:", error.message);
        return new Set(); 
    }
}


/**
 * The main function to check for leads and process them.
 */
async function checkAndFetchLeads() {
  console.log(`[${new Date().toLocaleString()}] Running lead check...`);
  
  try {
    const lastRunTime = getLastRunTime();
    
    const auth = await authorize();
    const gmail = google.gmail({version: 'v1', auth});
    const sheets = google.sheets({version: 'v4', auth});
    
    const existingIds = await getExistingMessageIds(sheets, SPREADSHEET_ID, SHEET_NAME);
    console.log(`Found ${existingIds.size} existing unique Message IDs in the sheet.`);

    const gmailQuery = `from:indiamart subject:film after:${lastRunTime} in:inbox`;
    console.log(`Using Gmail Query: ${gmailQuery}`);

    const listRes = await gmail.users.messages.list({
      userId: 'me',
      q: gmailQuery,
    });

    const messages = listRes.data.messages;
    if (!messages || messages.length === 0) {
      console.log('No new leads found since last run.');
      saveCurrentRunTime(); 
      return; 
    }

    console.log(`Found ${messages.length} potential lead(s).`);

    for (const message of messages) {
      const msgId = message.id;

      const msgRes = await gmail.users.messages.get({
        userId: 'me',
        id: msgId,
        format: 'FULL' 
      });

      const emailData = msgRes.data;
      const headers = emailData.payload.headers;
      const uniqueMessageIdHeader = headers.find(h => h.name.toLowerCase() === 'message-id');
      const uniqueMessageId = uniqueMessageIdHeader ? uniqueMessageIdHeader.value : null;

      if (uniqueMessageId && existingIds.has(uniqueMessageId)) {
          console.log(`Skipping message ${msgId}: Duplicate entry found for Message-ID: ${uniqueMessageId}`);
          continue; 
      }

      const emailHtmlBody = getEmailBody(emailData);
      if (!emailHtmlBody) {
        console.log(`Skipping message ${msgId}: No HTML body found.`);
        continue;
      }

      // --- RUN THE ULTIMATE PARSER ---
      const leadDetails = parseIndiaMartLead(emailHtmlBody, headers); // [Name, Phone, Email, Product]
      
      // --- !!! N/A PREVENTION CHECK (THE FIX) !!! ---
      // If Name AND Phone are N/A, it's a ghost email.
      const isJunkEntry = (leadDetails[0] === 'N/A' && leadDetails[1] === 'N/A');
      
      if (isJunkEntry) {
        console.warn(`Skipping message ${msgId}: Parser failed, Name and Phone were N/A. This is a ghost email.`);
        // Mark as read to get it out of the inbox, but DO NOT log it to the sheet.
        await gmail.users.messages.modify({
          userId: 'me',
          id: msgId,
          resource: { removeLabelIds: ['UNREAD'] },
        });
        continue; // Stop processing this loop and go to the next email.
      }
      // --- END OF N/A PREVENTION CHECK ---

      // --- If we are here, the lead is valid ---
      leadDetails.push(new Date().toLocaleString()); // Processed Date (Col E)
      leadDetails.push('New');                       // Lead Status (Col F)
      leadDetails.push('Yes');                       // Welcome Sent (Col G)
      leadDetails.push('');                          // Contacted Sent (Col H)
      leadDetails.push('');                          // Order Confirmed Sent (Col I)
      leadDetails.push(uniqueMessageId || msgId);    // Unique Message ID (Col J) 
      
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: `${SHEET_NAME}!A1`,
        valueInputOption: 'USER_ENTERED',
        resource: {
          values: [leadDetails], 
        },
      });

      console.log(`Successfully added unique lead for "${leadDetails[0]}" to Sheet.`);

      await gmail.users.messages.modify({
        userId: 'me',
        id: msgId,
        resource: {
          removeLabelIds: ['UNREAD'],
        },
      });
    }
    
    saveCurrentRunTime();

  } catch (error) {
    console.error('Error processing leads:', error.message);
    if (error.response && error.response.data) {
        console.error('Error details:', JSON.stringify(error.response.data, null, 2));
    }
    process.exit(1); 
  }
}

/**
 * Extracts the HTML body from the email.
 */
function getEmailBody(message) {
  let htmlBody = '';
  const payload = message.payload;
  
  if (payload.parts) {
    const part = payload.parts.find(p => p.mimeType === 'text/html');
    if (part && part.body.data) {
      htmlBody = Buffer.from(part.body.data, 'base64').toString('utf-8');
    }
  } else if (payload.body.data) {
    htmlBody = Buffer.from(payload.body.data, 'base64').toString('utf-8');
  }
  
  return htmlBody;
}

/**
 * --- THE ULTIMATE PARSER (v3.1) ---
 * This version explicitly returns "N/A" for empty fields to ensure the
 * N/A Prevention Check works correctly.
 */
function parseIndiaMartLead(body, headers) {
  const $ = cheerio.load(body);

  let name = '', phone = '', email = '', product = '';

  // --- 1. HUNT FOR PRODUCT ---
  product = $('div[style*="font-size:18px"] strong').text().trim();
  if (!product) {
    product = $('p:contains("I need") b, p:contains("I am looking for") b').text().trim();
  }

  // --- 2. HUNT FOR PHONE (Universal) ---
  phone = $('a[href*="call+91-"]').first().text().trim();
  if (phone.includes('(verified)')) {
      phone = phone.split(' (verified)')[0].trim();
  }

  // --- 3. HUNT FOR NAME ---
  const replyToHeader = headers.find(h => h.name.toLowerCase() === 'reply-to');
  if (replyToHeader) {
    const match = replyToHeader.value.match(/(.+)\s<(.+)>/);
    if (match) {
      name = match[1].trim();
    }
  }
  
  const buyLeadContactDiv = $('div[style*="color:#000000;line-height:1.5em;"]').first();
  if (!name || name === 'N/A' || name.toLowerCase() === 'indiamart') {
    name = buyLeadContactDiv.contents().first().text().trim();
  }

  if (!name || name === 'N/A' || name.toLowerCase() === 'indiamart') {
      const phoneSpan = $('span:contains("Click to call:")');
      name = phoneSpan.closest('tr').prev().prev().find('span').first().text().trim();
  }

  // --- 4. HUNT FOR EMAIL ---
  let htmlEmail = buyLeadContactDiv.find('a[href*="mailto:"]').first().text().trim();

  if (htmlEmail && htmlEmail !== 'buyleads@indiamart.com' && !htmlEmail.includes('@reply.indiamart.com')) {
    email = htmlEmail;
  } else {
    htmlEmail = $('span:contains("Email:")').find('a[href*="mailto:"]').first().text().trim();
    if (htmlEmail && htmlEmail !== 'buyleads@indiamart.com' && !htmlEmail.includes('@reply.indiamart.com')) {
      if (htmlEmail.includes('(verified)')) {
          htmlEmail = htmlEmail.split(' (verified)')[0].trim();
      }
      email = htmlEmail;
    }
  }
  
  if (email.includes('@reply.indiamart.com') || email === 'buyleads@indiamart.com') {
      email = 'N/A';
  }

  // --- 5. FINAL FORMATTING & N/A ASSIGNMENT ---
  
  // This is the critical fix. We ensure empty strings become "N/A".
  name = (name && name.toLowerCase() !== 'indiamart' && name !== 'Dear User') ? name : 'N/A';
  phone = (phone) ? phone : 'N/A';
  email = (email) ? email : 'N/A';
  product = (product) ? product : 'N/A';
  
  // Format phone if it's valid
  if (phone !== 'N/A') {
    phone = phone.replace(/-/g, ''); 
    if (!phone.startsWith("'")) {
      phone = "'" + phone; 
    }
  }

  return [
    truncateString(name),
    truncateString(phone),
    truncateString(email),
    truncateString(product),
  ];
}


// --- 4. SCRIPT EXECUTION ---
console.log('Script started. Running checkAndFetchLeads() one time...');
checkAndFetchLeads();