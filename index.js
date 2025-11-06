const {google} = require('googleapis');
const cheerio = require('cheerio');
const {authorize} = require('./authenticate');
const fs = require('fs'); // Import the File System module

// --- 1. CONFIGURATION ---
const SPREADSHEET_ID = '1Xo46YyTvM8CohicxGR2mVAvi94TLs8ab1eT9aMzWlMs'; 
const SHEET_NAME = 'Sheet1'; 
const LAST_RUN_FILE = 'last_run.txt'; // File to store last execution timestamp

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
 * Defaults to 48 hours ago if the file doesn't exist.
 */
function getLastRunTime() {
    try {
        // If the last_run.txt file exists, read the stored ISO timestamp
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
    // Fallback: If no file exists (first run), search for the last 48 hours to be safe
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
        // Save the current timestamp as a full ISO date string
        const currentTime = new Date().toISOString();
        fs.writeFileSync(LAST_RUN_FILE, currentTime);
        console.log(`Updated last run time to ${currentTime}`);
    } catch (err) {
        console.error("Error writing last_run.txt:", err);
    }
}


/**
 * The main function to check for leads and process them.
 */
async function checkAndFetchLeads() {
  console.log(`[${new Date().toLocaleString()}] Running lead check...`);
  
  try {
    // 1. Determine last run time
    const lastRunTime = getLastRunTime();
    console.log(`Searching for emails AFTER ${lastRunTime}...`);
    
    // 2. Authorize
    const auth = await authorize();
    const gmail = google.gmail({version: 'v1', auth});
    const sheets = google.sheets({version: 'v4', auth});

    // 3. Build the Gmail Query: REMOVED 'is:unread' and added 'after:'
    const gmailQuery = `from:indiamart subject:film after:${lastRunTime}`;
    console.log(`Using Gmail Query: ${gmailQuery}`);

    // 4. List messages that match the query
    const listRes = await gmail.users.messages.list({
      userId: 'me',
      q: gmailQuery,
    });

    const messages = listRes.data.messages;
    if (!messages || messages.length === 0) {
      console.log('No new leads found since last run.');
      saveCurrentRunTime(); // Save current time even if no messages were found
      return; 
    }

    console.log(`Found ${messages.length} new lead(s).`);

    for (const message of messages) {
      const msgId = message.id;

      // 5. Get the full email details
      const msgRes = await gmail.users.messages.get({
        userId: 'me',
        id: msgId,
      });

      const emailData = msgRes.data;
      
      // 6. Get the HTML Body
      const emailHtmlBody = getEmailBody(emailData);
      if (!emailHtmlBody) {
        console.log(`Skipping message ${msgId}: No HTML body found.`);
        continue;
      }

      // 7. Parse the HTML body (This function is updated)
      const leadDetails = parseIndiaMartLead(emailHtmlBody);
      
      // Add the current date
      leadDetails.push(new Date().toLocaleString()); 
      
      // 8. Write to Google Sheets
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: `${SHEET_NAME}!A1`,
        valueInputOption: 'USER_ENTERED',
        resource: {
          values: [leadDetails], 
        },
      });

      console.log(`Successfully added lead for "${leadDetails[0]}" to Sheet.`);

      // 9. Mark the email as "Read" (optional, but good for inbox hygiene)
      await gmail.users.messages.modify({
        userId: 'me',
        id: msgId,
        resource: {
          removeLabelIds: ['UNREAD'],
        },
      });

      console.log(`Marked message ${msgId} as read.`);
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