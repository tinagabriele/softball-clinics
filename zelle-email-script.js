/**
 * ZELLE PAYMENT EMAIL SCRIPT (Manual / Option C)
 * 
 * This is a SEPARATE Apps Script project — do NOT paste this into your
 * existing registration handler. Create a new script from your spreadsheet:
 * 
 * 1. Open your spreadsheet: https://docs.google.com/spreadsheets/d/1X1YjOTRvG4YNntntCQe9fJHvhNStW77XowyLX7LyMWA
 * 2. Go to Extensions → Apps Script
 * 3. If it opens your existing script, go to https://script.google.com and click "New Project" instead
 *    OR create a second .gs file in the same project (but the trigger approach below is safest as a new project)
 * 4. Paste this entire script
 * 5. Save and run sendZelleEmails() when you're ready to send
 * 
 * HOW IT WORKS:
 * - YOU run sendZelleEmails() manually when you're ready to process new registrations
 * - It looks for rows that don't have "EMAILED" or "SKIP" in the Email Sent column
 * - Sends a Zelle payment email to the parent/guardian
 * - Marks the row as "EMAILED" so it won't send again
 * 
 * TO EXCLUDE A ROW:
 * - Put "SKIP" (or "DO NOT EMAIL" or "NO") in the "Email Sent" column for that row
 * - The script will ignore it completely
 */

// ============ CONFIGURATION ============

const SHEET_ID = '1X1YjOTRvG4YNntntCQe9fJHvhNStW77XowyLX7LyMWA';
const FROM_NAME = 'Coach Christina Gabriele';
const ZELLE_QR_DRIVE_ID = '1pBB-KmOW-Cd0jrBKB2fEkw1jaxyv24O8';

// Pricing lookup
const PRICING = {
  'Curveball Craft - June 13': 75,
  'Riseball Rundown - August 22': 75,
  'Mini-Clinic': 50,
  "Mitchell College Catcher's Prospect & Development Camp - August 22, 2026 - 1:00 PM to 3:00 PM": 75  // post-July 1 price; change to 65 if before July 1
};

// ============ MAIN FUNCTION ============

function sendZelleEmails() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheets()[0]; // First sheet
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // Find column indices based on what your form submits
  const emailCol = findColumn(headers, ['email', 'parent_email']);
  const firstNameCol = findColumn(headers, ['first name', 'athlete_first_name']);
  const lastNameCol = findColumn(headers, ['last name', 'athlete_last_name']);
  const parentNameCol = findColumn(headers, ['parent/guardian', 'parent_name']);
  const sessionsCol = findColumn(headers, ['clinics interested', 'sessions_selected', 'clinics']);
  const notesCol = findColumn(headers, ['notes']);
  const pricingCol = findColumn(headers, ['pricing tier', 'pricing', 'price']);
  
  // The "Email Sent" column — we'll use the last column + 1 if it doesn't exist yet
  let emailSentCol = findColumn(headers, ['email sent', 'emailed', 'email_sent']);
  if (emailSentCol === -1) {
    // Add the header
    emailSentCol = headers.length;
    sheet.getRange(1, emailSentCol + 1).setValue('Email Sent');
  }
  
  // Process each row (skip header)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // Skip if already emailed or flagged to skip
    const emailStatus = (row[emailSentCol] || '').toString().toUpperCase().trim();
    const skipFlags = ['EMAILED', 'SKIP', 'DO NOT EMAIL', 'NO', 'DNE', 'EXCLUDE'];
    if (skipFlags.some(flag => emailStatus.includes(flag))) {
      continue;
    }
    
    // Skip empty rows
    if (!row[emailCol]) {
      continue;
    }
    
    const recipientEmail = row[emailCol];
    const athleteFirst = row[firstNameCol] || '';
    const athleteLast = row[lastNameCol] || '';
    const parentName = row[parentNameCol] || '';
    const sessions = row[sessionsCol] || '';
    const notes = row[notesCol] || '';
    const pricingTier = (pricingCol !== -1) ? (row[pricingCol] || '').toString().trim() : '';
    
    // Determine if this is a catchers camp registration
    const isCatchersCamp = notes.toString().includes('CATCHERS CAMP');
    
    // Calculate amount — use Pricing Tier column if filled, otherwise auto-calculate
    const amount = parsePricingTier(pricingTier) || calculateAmount(sessions, isCatchersCamp);
    
    // Build and send the email
    const subject = '🥎 Payment Info — ' + athleteFirst + ' ' + athleteLast + ' Registration';
    const htmlBody = buildPaymentEmail(parentName, athleteFirst, athleteLast, sessions, amount, isCatchersCamp);
    
    try {
      // Get the Zelle QR code as a blob for inline attachment
      const qrBlob = DriveApp.getFileById(ZELLE_QR_DRIVE_ID).getBlob().setName('zelle-qr.png');
      
      MailApp.sendEmail({
        to: recipientEmail,
        subject: subject,
        htmlBody: htmlBody,
        name: FROM_NAME,
        inlineImages: { zelleQR: qrBlob }
      });
      
      // Mark as emailed
      sheet.getRange(i + 1, emailSentCol + 1).setValue('EMAILED');
      sheet.getRange(i + 1, emailSentCol + 2).setValue(new Date().toLocaleString());
      
      // Small delay to avoid rate limits
      Utilities.sleep(1000);
      
    } catch (error) {
      Logger.log('Error sending to ' + recipientEmail + ': ' + error.toString());
      sheet.getRange(i + 1, emailSentCol + 1).setValue('ERROR: ' + error.toString());
    }
  }
}

// ============ EMAIL TEMPLATE ============

function buildPaymentEmail(parentName, athleteFirst, athleteLast, sessions, amount, isCatchersCamp) {
  const greeting = parentName ? parentName.split(' ')[0] : 'Hi there';
  const clinicName = isCatchersCamp 
    ? "Mitchell College Catcher's Prospect & Development Camp" 
    : sessions;
  
  return `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; color: #333;">
      <div style="background: ${isCatchersCamp ? '#d90429' : 'linear-gradient(135deg, #FF66C4 0%, #FFD93B 100%)'}; padding: 30px 20px; text-align: center; border-radius: 8px 8px 0 0;">
        <h1 style="color: white; margin: 0; font-size: 1.6em; text-shadow: 1px 1px 3px rgba(0,0,0,0.2);">🥎 Registration Received!</h1>
      </div>
      
      <div style="background: white; padding: 30px 25px; border: 1px solid #eee; border-top: none;">
        <p style="font-size: 1.05em;">Hi ${greeting},</p>
        
        <p>Thank you for registering <strong>${athleteFirst} ${athleteLast}</strong> for:</p>
        
        <div style="background: #f9f9f9; border-left: 4px solid ${isCatchersCamp ? '#d90429' : '#FF66C4'}; padding: 15px 20px; margin: 20px 0; border-radius: 0 6px 6px 0;">
          <strong style="color: ${isCatchersCamp ? '#d90429' : '#FF66C4'};">${clinicName}</strong>
        </div>
        
        <div style="background: #fff3cd; border: 1px solid #ffe082; border-radius: 8px; padding: 20px; margin: 25px 0; text-align: center;">
          <p style="margin: 0 0 5px; font-size: 0.95em; color: #856404;">⚠️ <strong>Important:</strong></p>
          <p style="margin: 0; font-size: 1.15em; color: #333; font-weight: bold;">Your spot is NOT confirmed until payment is received.</p>
          <p style="margin: 8px 0 0; font-size: 0.92em; color: #666;">Space is limited. Spots are filled on a first-paid basis.</p>
        </div>
        
        <h2 style="color: ${isCatchersCamp ? '#d90429' : '#FF66C4'}; font-size: 1.3em; margin-top: 30px;">💰 Payment Details</h2>
        
        <table style="width: 100%; border-collapse: collapse; margin: 15px 0;">
          <tr>
            <td style="padding: 10px 0; border-bottom: 1px solid #eee; font-weight: bold;">Amount Due:</td>
            <td style="padding: 10px 0; border-bottom: 1px solid #eee; text-align: right; font-size: 1.2em; font-weight: bold; color: #333;">$${amount}</td>
          </tr>
          <tr>
            <td style="padding: 10px 0; border-bottom: 1px solid #eee; font-weight: bold;">Payment Method:</td>
            <td style="padding: 10px 0; border-bottom: 1px solid #eee; text-align: right;">Zelle</td>
          </tr>

        </table>
        
        <div style="text-align: center; margin: 30px 0;">
          <p style="font-weight: bold; margin-bottom: 12px;">Scan to pay via Zelle:</p>
          <img src="cid:zelleQR" alt="Zelle QR Code" style="max-width: 250px; width: 100%; border: 2px solid #eee; border-radius: 8px; padding: 10px;">
          <p style="margin-top: 12px; font-size: 0.9em; color: #555;"><strong>When sending payment via Zelle, please select "Paying a friend, partner, or family member."</strong> Do not select any of the other options as this may cause delays or issues with your payment.</p>
          <p style="margin-top: 10px; font-size: 0.85em; color: #888;">${isCatchersCamp || !sessions.toString().toLowerCase().includes('mini-clinic') ? 'Advance payment is required for all clinics.' : 'For mini-clinics/private lessons, cash is also accepted at the session.'}</p>
        </div>
        
        <div style="background: #f0f8ff; border: 1px solid #b8daff; border-radius: 8px; padding: 15px 20px; margin: 20px 0;">
          <p style="margin: 0; font-size: 0.95em; color: #555;"><strong>📋 Cancellation Policy:</strong></p>
          <ul style="margin: 10px 0 0; padding-left: 20px; color: #555; font-size: 0.9em;">
            <li>Within 48 hours of registration: Full refund</li>
            <li>More than 2 weeks before session: 50% refund</li>
            <li>Less than 2 weeks before session: Non-refundable</li>
          </ul>
          <p style="margin: 10px 0 0; font-size: 0.85em; color: #888;">If you need to cancel within 2 weeks of your session, please reach out — we may be able to find someone to fill your spot.</p>
        </div>
        
        <p style="margin-top: 25px;">If you have any questions or need to make changes, just reply to this email or reach out to me directly.</p>
        
        <p style="margin-top: 20px;">See you on the field! 🥎</p>
        <p><strong>— Coach Christina Gabriele</strong><br>
        <span style="color: #888; font-size: 0.9em;">gabriele_c@mitchell.edu</span></p>
      </div>
      
      <div style="background: #333; padding: 20px; text-align: center; border-radius: 0 0 8px 8px;">
        <p style="color: #aaa; margin: 0; font-size: 0.85em;">Christina Gabriele | CT Softball Pitching Coach</p>
        <p style="color: #aaa; margin: 5px 0 0; font-size: 0.8em;">📍 Southeastern CT</p>
      </div>
    </div>
  `;
}

// ============ HELPERS ============

function parsePricingTier(pricingTier) {
  if (!pricingTier) return null;
  
  // If it's just a number like "$75" or "75", use it directly
  const directMatch = pricingTier.match(/^\$?(\d+)$/);
  if (directMatch) return parseInt(directMatch[1]);
  
  // If it contains "early bird" logic, check the date
  if (pricingTier.toLowerCase().includes('early bird') || pricingTier.toLowerCase().includes('before july')) {
    const now = new Date();
    const july1 = new Date('2026-07-01');
    return now < july1 ? 65 : 75;
  }
  
  // Try to find any dollar amount
  const dollarMatch = pricingTier.match(/\$(\d+)/);
  if (dollarMatch) return parseInt(dollarMatch[1]);
  
  return null; // Couldn't parse — fall back to auto-calculate
}

function calculateAmount(sessions, isCatchersCamp) {
  if (isCatchersCamp) {
    // After July 1 = $75, before = $65
    const now = new Date();
    const july1 = new Date('2026-07-01');
    return now < july1 ? 65 : 75;
  }
  
  // Pitching clinics — could have multiple sessions
  const sessionList = sessions.toString().split(',').map(s => s.trim());
  let total = 0;
  
  for (const session of sessionList) {
    if (session.toLowerCase().includes('mini-clinic') || session.toLowerCase().includes('mini clinic')) {
      total += 50;
    } else if (session.toLowerCase().includes('curveball') || session.toLowerCase().includes('riseball')) {
      total += 75;
    }
  }
  
  return total || 75; // Default to $75 if we can't parse
}

function findColumn(headers, possibleNames) {
  for (let i = 0; i < headers.length; i++) {
    const header = headers[i].toString().toLowerCase().replace(/[_\/]/g, ' ').trim();
    for (const name of possibleNames) {
      if (header.includes(name.toLowerCase())) {
        return i;
      }
    }
  }
  return -1;
}

// ============ HOW TO USE ============

/**
 * SENDING EMAILS:
 * 1. Open your spreadsheet and review new registrations
 * 2. For any rows you do NOT want to email, type "SKIP" in the "Email Sent" column
 * 3. Come here and run sendZelleEmails()
 * 4. It will email everyone who doesn't have EMAILED or SKIP in that column
 * 5. Check the sheet — processed rows will say "EMAILED" with a timestamp
 * 
 * SKIP FLAGS (any of these in the "Email Sent" column will exclude a row):
 *   SKIP, DO NOT EMAIL, NO, DNE, EXCLUDE
 */

/**
 * Run this to send Zelle payment emails to all un-processed, un-skipped rows.
 */
function testRun() {
  sendZelleEmails();
  Logger.log('✅ Done! Check your sheet for EMAILED markers and the recipient inbox.');
}
