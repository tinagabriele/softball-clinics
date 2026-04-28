# Google Sheets Registration Backend — Setup Guide

## Step 1: Create the Google Sheet

1. Log into **elenagabriele81@gmail.com**
2. Go to [sheets.google.com](https://sheets.google.com)
3. Create a new spreadsheet — name it **"Pitch Perfect Registrations"**
4. In Row 1, add these column headers (A through N):

```
Timestamp | First Name | Last Name | Age | Grade | Experience Level | Current Team | Parent/Guardian | Relationship | Email | Phone | Clinics Interested | Referral Source | Notes
```

## Step 2: Add the Apps Script

1. In your spreadsheet, go to **Extensions → Apps Script**
2. Delete any existing code in the editor
3. Paste this entire script:

```javascript
function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);

    sheet.appendRow([
      data.timestamp,
      data.athlete_first_name,
      data.athlete_last_name,
      data.athlete_age,
      data.athlete_grade,
      data.experience_level,
      data.current_team,
      data.parent_name,
      data.relationship,
      data.parent_email,
      data.parent_phone,
      data.clinics_interested,
      data.referral_source,
      data.notes
    ]);

    // Send email notification
    var subject = "🥎 New Clinic Registration: " + data.athlete_first_name + " " + data.athlete_last_name;
    var body = "New registration received!\n\n" +
      "Athlete: " + data.athlete_first_name + " " + data.athlete_last_name + "\n" +
      "Age: " + data.athlete_age + " | Grade: " + data.athlete_grade + "\n" +
      "Experience: " + data.experience_level + "\n" +
      "Team: " + (data.current_team || "N/A") + "\n\n" +
      "Parent/Guardian: " + data.parent_name + " (" + data.relationship + ")\n" +
      "Email: " + data.parent_email + "\n" +
      "Phone: " + data.parent_phone + "\n\n" +
      "Clinics: " + data.clinics_interested + "\n" +
      "Referral: " + (data.referral_source || "N/A") + "\n" +
      "Notes: " + (data.notes || "None") + "\n\n" +
      "Submitted: " + data.timestamp;

    MailApp.sendEmail("gabriele_c@mitchell.edu", subject, body);

    return ContentService
      .createTextOutput(JSON.stringify({"result": "success"}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({"result": "error", "error": error.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
```

4. Click **Save** (💾 icon or Ctrl+S)
5. Name the project: **"Clinic Registration Handler"**

## Step 3: Deploy as Web App

1. Click **Deploy → New deployment**
2. Click the gear icon ⚙️ next to "Select type" → choose **Web app**
3. Set these options:
   - **Description:** "Clinic registration form handler"
   - **Execute as:** Me (elenagabriele81@gmail.com)
   - **Who has access:** Anyone
4. Click **Deploy**
5. It will ask you to authorize — click through and allow access
6. **Copy the Web App URL** it gives you (looks like: `https://script.google.com/macros/s/AKfycb.../exec`)

## Step 4: Connect to Your Website

Open `register.html` and find this line near the bottom:

```javascript
const GOOGLE_SCRIPT_URL = 'YOUR_APPS_SCRIPT_URL_HERE';
```

Replace `YOUR_APPS_SCRIPT_URL_HERE` with the URL you copied in Step 3.

## Step 5: Test It!

1. Push the changes to GitHub
2. Open your registration page
3. Fill out the form and submit
4. Check your Google Sheet — the row should appear
5. Check gabriele_c@mitchell.edu — you should get an email notification

## What You Get

Every registration will:
- ✅ Add a row to your Google Sheet (full history)
- ✅ Send an email to gabriele_c@mitchell.edu with the details
- ✅ Show a success message to the person registering

## Troubleshooting

- **Form submits but no data appears:** Make sure the Apps Script URL is correct and the deployment is set to "Anyone" access
- **Authorization error:** Re-deploy and make sure you authorized all the permissions (Sheets + Mail)
- **CORS error in console:** The form uses `mode: 'no-cors'` which should handle this, but if issues persist, try redeploying
