// Google Sheets contact-form endpoint for aadi.html.
// Deploy this project as a Web app, execute as your account, and allow access
// to anyone. Then set the resulting /exec URL in GOOGLE_SHEETS_FORM_ENDPOINT.

const CONTACT_SHEET_ID = "1CB3tFkAJsKJHVuoCVRfDvdyk3okoFiNJn7mKH6vdMig";
const CONTACT_HEADERS = [
  "Submitted at",
  "Name",
  "Email",
  "Organization",
  "Interested in",
  "Details",
  "WhatsApp"
];

function doGet() {
  return renderMessage(
    "Contact form endpoint",
    "This endpoint accepts contact-form submissions. Return to the portfolio and submit the form there."
  );
}

function doPost(event) {
  const fields = event && event.parameter ? event.parameter : {};

  if (fields.website) {
    return renderMessage("Request not accepted", "Please return to the portfolio and try again.");
  }

  const name = readField(fields.name, 120);
  const email = readField(fields.email, 254);
  const whatsapp = readField(fields.whatsapp, 30);
  const organization = readField(fields.organization, 200);
  const engagement = readField(fields.engagement, 80);
  const details = readField(fields.details, 2000);
  const allowedEngagements = ["mentoring", "training", "consulting", "course"];
  const whatsappDigits = whatsapp.replace(/\D/g, "");

  if (!name || !email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) ||
      !/^[+()\d\s.-]{7,30}$/.test(whatsapp) || whatsappDigits.length < 7 || whatsappDigits.length > 15 ||
      allowedEngagements.indexOf(engagement) === -1) {
    return renderMessage("Request not accepted", "Check your name, email, WhatsApp number, and area of interest, then submit the form again.");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const spreadsheet = SpreadsheetApp.openById(CONTACT_SHEET_ID);
    const sheet = spreadsheet.getSheets()[0];

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(CONTACT_HEADERS);
    } else {
      const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      CONTACT_HEADERS.forEach(header => {
        if (!existingHeaders.some(value => String(value).trim().toLowerCase() === header.toLowerCase())) {
          sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
          existingHeaders.push(header);
        }
      });
    }

    const valuesByHeader = {
      "submitted at": new Date(),
      "name": safeCell(name),
      "email": safeCell(email),
      "organization": safeCell(organization),
      "interested in": safeCell(engagement),
      "details": safeCell(details),
      "whatsapp": safeCell(whatsapp)
    };
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    sheet.appendRow(headers.map(header => valuesByHeader[String(header).trim().toLowerCase()] || ""));
  } finally {
    lock.releaseLock();
  }

  return renderMessage("Request received", "Thanks for reaching out. Your request has been added to the contact sheet.");
}

function readField(value, maxLength) {
  return String(value || "").trim().slice(0, maxLength);
}

function safeCell(value) {
  const text = String(value || "");
  return /^[=+\-@\t\r]/.test(text) ? "'" + text : text;
}

function renderMessage(title, message) {
  const html = [
    "<!doctype html><html lang=\"en\"><head><meta charset=\"utf-8\">",
    "<meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">",
    "<title>", title, "</title>",
    "<style>",
    "body{margin:0;min-height:100vh;display:grid;place-items:center;background:#050505;color:#fff;font:16px Arial,sans-serif}",
    "main{max-width:34rem;margin:1.5rem;padding:2.5rem;border:1px solid #2a2a2a;background:#0a0a0a}",
    "p{color:#aaa;line-height:1.7}a{color:#fff}",
    "</style></head><body><main><p>CONTACT REQUEST</p><h1>", title, "</h1><p>",
    message,
    "</p><p>You can close this tab to return to the portfolio.</p></main></body></html>"
  ].join("");

  return HtmlService.createHtmlOutput(html);
}
