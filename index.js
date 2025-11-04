const {google} = require('googleapis');
const cron = require('node-cron');
const cheerio = require('cheerio'); // Import the HTML parser
const {authorize} = require('./authenticate');

// --- 1. CONFIGURATION ---
const SPREADSHEET_ID = '1Xo46YyTvM8CohicxGR2mVAvi94TLs8ab1eT9aMzWlMs'; 
const SHEET_NAME = 'Sheet1'; // As you requested
// ---------------------

const MAX_CELL_CHARS = 49999; // Google's character limit for a cell

/**
 * --- SAFETY FUNCTION ---
 * Truncates any string to be safely under Google's limit.
 */
function truncateString(str) {
  if (str && str.length > MAX_CELL_CHARS) {
    console.warn(`Warning: Data was truncated.`);
    return str.substring(0, MAX_CELL_CHARS) + "\n... (truncated)";
  }
  return str;
}


/**
 * The main function to check for leads and process them.
 */
async function checkAndFetchLeads() {
  console.log(`[${new Date().toLocaleString()}] Running lead check...`);
  
  try {
    // 1. Authorize
    const auth = await authorize();
    const gmail = google.gmail({version: 'v1', auth});
    const sheets = google.sheets({version: 'v4', auth});

    // 2. Build the Gmail Query
    const gmailQuery = 'from:indiamart subject:film after:2025/11/01 is:unread';

    // 3. List messages that match the query
    const listRes = await gmail.users.messages.list({
      userId: 'me',
      q: gmailQuery,
    });

    const messages = listRes.data.messages;
    if (!messages || messages.length === 0) {
      console.log('No new leads found.');
      return;
    }

    console.log(`Found ${messages.length} new lead(s).`);

    for (const message of messages) {
      const msgId = message.id;

      // 4. Get the full email details
      const msgRes = await gmail.users.messages.get({
        userId: 'me',
        id: msgId,
      });

      const emailData = msgRes.data;
      
      // 5. Get the HTML Body
      const emailHtmlBody = getEmailBody(emailData);
      if (!emailHtmlBody) {
        console.log(`Skipping message ${msgId}: No HTML body found.`);
        continue;
      }

      // 6. Parse the HTML body (This function is updated)
      const leadDetails = parseIndiaMartLead(emailHtmlBody);
      
      // Add the current date
      leadDetails.push(new Date().toLocaleString()); 
      // leadDetails is now: [Name, Phone, Email, Product, ProcessedDate]

      // 7. Write to Google Sheets
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: `${SHEET_NAME}!A1`,
        valueInputOption: 'USER_ENTERED',
        resource: {
          values: [leadDetails], 
        },
      });

      console.log(`Successfully added lead for "${leadDetails[0]}" to Sheet.`);

      // 8. Mark the email as "Read"
      await gmail.users.messages.modify({
        userId: 'me',
        id: msgId,
        resource: {
          removeLabelIds: ['UNREAD'],
        },
      });

      console.log(`Marked message ${msgId} as read.`);
    }

  } catch (error) {
    console.error('Error processing leads:', error.message);
    if (error.response && error.response.data) {
        console.error('Error details:', JSON.stringify(error.response.data, null, 2));
    }
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
 * --- UPDATED FUNCTION ---
 * This function now has all "Requirement Details" logic REMOVED.
 */
function parseIndiaMartLead(body) {
  // Load the HTML into cheerio
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
    // --- Try Parsing Template 2 (The "Enquiry" format, like Bikesh's & Ganesh's) ---
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
  
  // --- !!! UPDATED PHONE FORMATTING !!! ---
  if (phone && phone !== 'N/A') {
    phone = phone.replace(' (verified)', ''); // <-- THIS IS THE FIX
    phone = phone.replace(/-/g, ''); // Remove all dashes
    if (!phone.startsWith("'")) {
      phone = "'" + phone; // Add leading apostrophe to force string format in Sheets
    }
  }

  
  // --- FINAL SAFETY TRUNCATION & RETURN ---
  // Return only the 4 items you want.
  return [
    truncateString(name || 'N/A'),
    truncateString(phone || 'N/A'),
    truncateString(email || 'N/A'),
    truncateString(product || 'N/A'),
  ];
}


// --- 4. Scheduling ---
console.log('Script started. Waiting for the next scheduled run...');
cron.schedule('0 */12 * * *', checkAndFetchLeads);

// Run it once immediately when the script starts
checkAndFetchLeads();