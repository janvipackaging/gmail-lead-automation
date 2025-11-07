const {google} = require('googleapis');
const cheerio = require('cheerio');
const {authorize} = require('./authenticate');
const fs = require('fs');

// --- 1. CONFIGURATION ---
const SPREADSHEET_ID = '1Xo46YyTvM8CohicxGR2mVAvi94TLs8ab1eT9aMzWlMs'; 
const SHEET_NAME = 'Sheet1'; 
const LAST_RUN_FILE = 'last_run.txt'; 
const MAX_CELL_CHARS = 49999;
const UNIQUE_ID_COLUMN_INDEX = 10; // Column J is the 10th column (1-indexed)

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
            
            // Gmail API search requires YYYY/MM/DD format for 'after:' query
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
        
        // Flatten the array of arrays and filter out empty cells/header
        return new Set((response.data.values || []).flat().filter(id => id && id !== 'Unique Message ID'));
    } catch (error) {
        console.error("Error fetching existing IDs:", error.message);
        // If it fails, assume no duplicates exist to avoid blocking the job
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
    
    // 1. Authorize
    const auth = await authorize();
    const gmail = google.gmail({version: 'v1', auth});
    const sheets = google.sheets({version: 'v4', auth});
    
    // 2. Fetch existing IDs from the sheet
    const existingIds = await getExistingMessageIds(sheets, SPREADSHEET_ID, SHEET_NAME);
    console.log(`Found ${existingIds.size} existing unique Message IDs in the sheet.`);

    // 3. Build the Gmail Query
    const gmailQuery = `from:indiamart subject:film after:${lastRunTime}`;
    console.log(`Using Gmail Query: ${gmailQuery}`);

    // 4. List messages
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

      // 5. Get the full email details (including headers)
      const msgRes = await gmail.users.messages.get({
        userId: 'me',
        id: msgId,
        format: 'FULL' 
      });

      const emailData = msgRes.data;
      
      // Find the message-ID in the headers
      const headers = emailData.payload.headers;
      const uniqueMessageIdHeader = headers.find(h => h.name.toLowerCase() === 'message-id');
      const uniqueMessageId = uniqueMessageIdHeader ? uniqueMessageIdHeader.value : null;

      // --- CRITICAL DUPLICATE CHECK ---
      if (uniqueMessageId && existingIds.has(uniqueMessageId)) {
          console.log(`Skipping message ${msgId}: Duplicate entry found for Message-ID: ${uniqueMessageId}`);
          continue; // Skip to the next message
      }

      // 6. Get the HTML Body
      const emailHtmlBody = getEmailBody(emailData);
      if (!emailHtmlBody) {
        console.log(`Skipping message ${msgId}: No HTML body found.`);
        continue;
      }

      // 7. Parse the HTML body
      const leadDetails = parseIndiaMartLead(emailHtmlBody);
      
      // Prepare the row: [Name, Phone, Email, Product, Date, Unique Message ID]
      leadDetails.push(new Date().toLocaleString()); // Processed Date (Col E)
      leadDetails.push('New');                       // Lead Status (Col F)
      leadDetails.push('Yes');                       // Welcome Sent (Col G)
      leadDetails.push('');                          // Contacted Sent (Col H)
      leadDetails.push('');                          // Order Confirmed Sent (Col I)
      leadDetails.push(uniqueMessageId || msgId);    // Unique Message ID (Col J) 
      
      // 8. Write to Google Sheets
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: `${SHEET_NAME}!A1`,
        valueInputOption: 'USER_ENTERED',
        resource: {
          values: [leadDetails], 
        },
      });

      console.log(`Successfully added unique lead for "${leadDetails[0]}" to Sheet.`);

      // 9. Mark the email as "Read" 
      await gmail.users.messages.modify({
        userId: 'me',
        id: msgId,
        resource: {
          removeLabelIds: ['UNREAD'],
        },
      });
    }
    
    // Save the time *after* successfully processing all leads
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
 * Parses the HTML for lead details.
 */
function parseIndiaMartLead(body) {
  const $ = cheerio.load(body);

  let name, phone, email, product;

  // --- Try Parsing Template 1 (The "Buylead" format) ---
  product = $('div[style*="font-size:18px"] strong').text().trim();
  
  if (product) {
    // IT IS TEMPLATE 1
    const contactDiv = $('div[style*="color:#000000;line-height:1.5em;"]').first();
    name = contactDiv.contents().first().text().trim();
    phone = contactDiv.find('a[href*="call+91-"]').text().trim();
    email = contactDiv.find('a[href*="mailto:"]').text().trim();

  } else {
    // --- Try Parsing Template 2 (The "Enquiry" format) ---
    product = $('p:contains("I need") b, p:contains("I am looking for") b').text().trim();
    
    if (product) {
        // IT IS TEMPLATE 2
        const phoneSpan = $('span:contains("Click to call:")');
        phone = phoneSpan.find('a[href*="call+91-"]').first().text().trim();
        name = phoneSpan.closest('tr').prev().prev().find('span').first().text().trim();
        
        const emailSpans = $('span:contains("Email:")').find('a[href*="mailto:"]');
        const emails = [];
        emailSpans.each((i, el) => {
            emails.push($(el).text().trim());
        });
        email = emails.join(', ');

    } else {
        // --- UNKNOWN TEMPLATE ---
        console.warn('Unknown email template found. Parsing may be incomplete.');
        name = 'N/A';
        phone = $('a[href*="call+91-"]').first().text().trim();
        email = $('a[href*="mailto:"]').first().text().trim(); 
        product = 'N/A';
    }
  }
  
  // --- PHONE FORMATTING ---
  if (phone && phone !== 'N/A') {
    phone = phone.replace(' (verified)', ''); 
    phone = phone.replace(/-/g, ''); 
    if (!phone.startsWith("'")) {
      phone = "'" + phone; 
    }
  }

  
  // --- FINAL SAFETY TRUNCATION & RETURN ---
  // Returns: [Name, Phone, Email, Product]
  return [
    truncateString(name || 'N/A'),
    truncateString(phone || 'N/A'),
    truncateString(email || 'N/A'),
    truncateString(product || 'N/A'),
  ];
}


// --- 4. SCRIPT EXECUTION ---
console.log('Script started. Running checkAndFetchLeads() one time...');
checkAndFetchLeads();