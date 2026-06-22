/**
 * BCOC 2026 Elections — Candidacy form backend (Google Apps Script)
 * Emails each submission to the Elections Committee and (optionally) logs it to a Sheet.
 *
 * DEPLOY:
 *   1. Go to https://script.google.com  →  New project
 *   2. Paste this code (replace the default Code.gs contents)
 *   3. (Optional) set SHEET_ID below to a Google Sheet ID to also log rows
 *   4. Deploy ▸ New deployment ▸ type "Web app"
 *        - Execute as: Me
 *        - Who has access: Anyone
 *   5. Copy the Web app URL and paste it into ENDPOINT in
 *      BCOC_Elections_Candidacy_Form.html
 */

// Where submissions go:
var MAIL_TO  = 'elections@bcoc.us';
var MAIL_CC  = 'hsmith@bcoc.us,glawrence@bcoc.us';

// Optional: paste a Google Sheet ID to log submissions (leave '' to skip).
var SHEET_ID = '';

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    // Build a readable email body from all submitted fields
    var order = ['Full Name','Rank / Title','Department / Agency','Email','Phone',
                 'Office Sought','Years of BCOC Membership','Statement of Interest',
                 'Nominated By','timestamp'];
    var lines = [];
    order.forEach(function(k){ if (data[k]) lines.push(k + ': ' + data[k]); });
    // include any extra fields not in the ordered list
    Object.keys(data).forEach(function(k){
      if (order.indexOf(k) === -1 && k !== 'formType') lines.push(k + ': ' + data[k]);
    });

    var name = data['Full Name'] || 'Candidate';
    var office = data['Office Sought'] || '';
    var subject = 'BCOC 2026 Candidacy — ' + name + (office ? ' (' + office + ')' : '');

    MailApp.sendEmail({
      to: MAIL_TO,
      cc: MAIL_CC,
      replyTo: data['Email'] || MAIL_TO,
      subject: subject,
      body: 'A new candidacy declaration was submitted via bcoc.us:\n\n' + lines.join('\n') + '\n'
    });

    if (SHEET_ID) {
      var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
      sheet.appendRow([
        new Date(), name, data['Rank / Title']||'', data['Department / Agency']||'',
        data['Email']||'', data['Phone']||'', office, data['Years of BCOC Membership']||'',
        data['Statement of Interest']||'', data['Nominated By']||''
      ]);
    }

    return ContentService.createTextOutput(JSON.stringify({ok:true}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
