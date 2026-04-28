# OR Scheduling Optimizer

A browser-based internal tool for optimizing operating room scheduling. Reads consultation data from Google Sheets, generates ranked surgery slot recommendations per case, tracks OR utilization in real time, and writes confirmed bookings back to your spreadsheet with email reports to the scheduler.

---

## Features

- **Live Google Sheets integration** — reads from your Consult Raw Data and Procedure Durations tabs
- **Smart scheduling optimizer** — prioritizes cases by revenue per minute ($/min), respects surgeon and OR constraints, and places each case in the earliest feasible slot
- **3 ranked options per case** — Option 1 is always the best fit; PAs can pick any option or change their selection later
- **Real-time re-optimization** — every time a booking is confirmed, all unconfirmed cases are re-slotted around it automatically
- **OR utilization dashboard** — live timeline view per room per day, with utilization % tracking
- **Write-back to Google Sheets** — confirmed bookings are written to a "Confirmed Bookings" tab automatically
- **Email reports** — sends a styled daily summary grouped by OR room to the scheduler, auto-triggered on each confirmation and available on demand

---

## Scheduling Rules

| Constraint | Value |
|---|---|
| OR rooms | 2 |
| OR hours per room per day | 8 hours (9am – 5pm) |
| Surgeon max hours per day | 8 hours |
| Minimum lead time | 2 weeks after consult date |
| Working days | Monday – Friday only |

**Priority order:**
1. Highest revenue per minute ($/min) first
2. Earlier consult date gets earlier slot options (tiebreaker)
3. Each case placed in the earliest available gap that fits its duration

---

## Google Sheets Setup

The spreadsheet needs two tabs with the following columns (names are configurable in the tool):

**Consult Raw Data**
| Column | Description |
|---|---|
| ID | Unique booking identifier (e.g. ORD-001) |
| Procedure | Name of the surgical procedure |
| Surgeon | Surgeon's name |
| Date | Date the consultation was booked (YYYY-MM-DD) |

**Procedure Durations**
| Column | Description |
|---|---|
| Procedure | Procedure name (must match Consult Raw Data) |
| DurationMinutes | Duration of the procedure in minutes |
| Revenue | Revenue value for the procedure (optional but recommended) |

**Sharing:** The spreadsheet must be shared as **"Anyone with the link can view"** for the Google Sheets API to read it.

---

## Google Sheets API Setup

1. Go to [Google Cloud Console](https://console.cloud.google.com)
2. Create a new project (or use an existing one)
3. Enable the **Google Sheets API**
4. Go to **Credentials → Create Credentials → API Key**
5. Restrict the key to the Sheets API only
6. Copy the API key into the tool's Setup tab

---

## Apps Script Setup (Write-back & Email)

To enable writing confirmed bookings back to your sheet and sending email reports, you need to deploy a Google Apps Script Web App from within your spreadsheet.

### Steps

1. Open your Google Sheet → **Extensions → Apps Script**
2. Delete any existing code and paste the script below
3. Replace `YOUR_SECRET_TOKEN_HERE` with a secret string of your choice (e.g. `abcxyz`)
4. Click **Deploy → New deployment**
   - Type: **Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
5. Click **Deploy**, authorise when prompted, and copy the Web App URL
6. Paste the URL and your secret token into the tool's Setup tab

### Apps Script Code

```javascript
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);

    // Security: reject requests without the correct secret token
    var SECRET_TOKEN = 'YOUR_SECRET_TOKEN_HERE'; // change this to your own secret
    if (payload.token !== SECRET_TOKEN) {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'Unauthorized' }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    var action = payload.action;
    var ss     = SpreadsheetApp.getActiveSpreadsheet();

    if (action === 'writeBookings') {
      var tabName = 'Confirmed Bookings';
      var sheet   = ss.getSheetByName(tabName) || ss.insertSheet(tabName);
      sheet.clearContents();
      var headers = ['Booking ID','Procedure','Surgeon','Consult Date',
                     'Surgery Date','OR Room','Start Time','End Time',
                     'Duration (min)','Revenue','Confirmed At'];
      sheet.appendRow(headers);
      payload.bookings.forEach(function(b) { sheet.appendRow(b); });
      sheet.autoResizeColumns(1, headers.length);
    }

    if (action === 'sendEmail') {
      var to      = payload.to;
      var subject = payload.subject;
      var body    = payload.body;
      MailApp.sendEmail({ to: to, subject: subject, htmlBody: body });
    }

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
                         .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.message }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}
```

> **Security note:** The Web App is deployed as "Anyone" so the URL works without a login flow, but the secret token acts as a password — requests without it are rejected. Share the token with your team the same way you would share a password.

---

## Deploying to GitHub Pages

1. Create a new **public** GitHub repository
2. Upload `or_scheduler.html` and rename it to `index.html`
3. Go to **Settings → Pages → Source: Deploy from a branch → main / root → Save**
4. Your tool will be live at `https://yourusername.github.io/your-repo-name` within 1–2 minutes

To update the tool, edit `index.html` in the repository and commit — GitHub Pages redeploys automatically.

---

## Usage

### Setup tab
1. Enter your Google Sheets API key and Spreadsheet ID
2. Confirm your tab and column names match your sheet
3. Enter your Apps Script Web App URL, secret token, and scheduler email
4. Click **Connect & Load Data**
5. Click **Generate Schedule Recommendations**

### Schedule tab
- Each case shows 3 ranked slot options (Option 1 = best recommendation)
- Click **Select this slot** to confirm a booking
- Click **Change** on a confirmed booking to reselect
- Confirming a booking automatically re-optimizes all pending cases and writes to the sheet
- Send Report Now button allows you to send the email anytime you are done with scheduling.
      + While sending — the button disables and changes to "Sending…" so they know it's in progress.
      + On success — the button turns green and shows "✓ Report Sent", then resets back to "Send Report Now" after 3 seconds. A green toast notification also appears next to the button saying "✓ Email report sent to scheduler@hospital.com" and fades out after 5 seconds.
      + On failure — the button resets immediately and a red toast shows the specific error message for 6 seconds.
      + If the email isn't configured yet, an amber warning toast prompts them to check the Setup tab.

### OR Dashboard tab
- Pick a date to view the OR timeline and utilization for both rooms
- All confirmed bookings are listed at the bottom sorted by surgery date

---

## Tech Stack

- Vanilla HTML / CSS / JavaScript — no frameworks, no dependencies, single file
- Google Sheets API v4 (read)
- Google Apps Script Web App (write-back + email)

---

## Limitations & Known Issues

- The optimizer generates options sequentially (highest $/min first), so in edge cases a lower-priority case may block an ideal slot for a higher-priority one if it was placed first in a previous run. Re-running "Generate Schedule Recommendations" resets all unconfirmed options.
- Revenue and duration data must be present in the Procedure Durations tab for accurate prioritization. Cases with missing revenue are treated as $0 and scheduled last.
- The tool does not currently handle public holidays — only weekends are excluded.
