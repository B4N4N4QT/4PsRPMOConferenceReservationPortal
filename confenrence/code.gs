/**
 * CONFERENCE ROOM RESERVATION SYSTEM
 * Google Apps Script Backend
 */

const SHEET_NAME_RESERVATIONS = "Reservations";
const SHEET_NAME_ROOMS = "Rooms";
const SHEET_NAME_ADMINS = "Admins";

// Specific Admin Emails for Notifications and Privileges
const ADMIN_NOTIFICATION_EMAILS = "teamshinji15@gmail.com, joarciaga@dswd.gov.ph";

/**
 * Serves the HTML page or handles email-based approval actions
 */
function doGet(e) {
  // Check if e and e.parameter exist to prevent "Cannot read properties of undefined" errors
  if (e && e.parameter && e.parameter.action && e.parameter.id) {
    try {
      const result = updateReservationStatus(e.parameter.id, e.parameter.action);
      const message = result.success ? 
        `Tagumpay: Ang request ay ${e.parameter.action === 'Approved' ? 'na-approve na' : 'tinanggihan na'}.` : 
        `Error: ${result.message}`;
      
      return HtmlService.createHtmlOutput(`
        <div style="font-family: sans-serif; text-align: center; padding: 50px;">
          <h2 style="color: #1e3a8a;">4Ps RPMO Conference Reservation Portal</h2>
          <p style="font-size: 1.2rem; color: #333;">${message}</p>
          <p style="color: #666;">Maaari mo nang isara ang window na ito.</p>
        </div>
      `);
    } catch (err) {
      return HtmlService.createHtmlOutput(`
        <div style="font-family: sans-serif; text-align: center; padding: 50px;">
          <h2 style="color: #be123c;">System Error</h2>
          <p style="font-size: 1.1rem;">Hindi maproseso ang iyong request sa ngayon.</p>
          <p style="color: #666;">${err.toString()}</p>
        </div>
      `);
    }
  }

  // Fallback to the main application if no parameters are present
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('4Ps RPMO Conference Reservation Portal')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Get current user email for admin verification
 */
function getUserEmail() {
  return Session.getActiveUser().getEmail();
}

/**
 * Fetch data for the frontend
 */
function getAppData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userEmail = Session.getActiveUser().getEmail();
  const tz = Session.getScriptTimeZone();
  
  const roomSheet = ss.getSheetByName(SHEET_NAME_ROOMS);
  let rooms = [];
  if (roomSheet && roomSheet.getLastRow() > 1) {
    rooms = roomSheet.getRange(2, 1, roomSheet.getLastRow() - 1, 2).getValues();
  }
  
  const resSheet = ss.getSheetByName(SHEET_NAME_RESERVATIONS);
  let reservations = [];
  if (resSheet && resSheet.getLastRow() > 1) {
    const data = resSheet.getRange(2, 1, resSheet.getLastRow() - 1, 17).getValues();
    reservations = data.map((r, index) => {
      const formatVal = (val) => {
        if (val instanceof Date) return Utilities.formatDate(val, tz, "yyyy-MM-dd");
        return val ? String(val).split('T')[0] : "";
      };

      let sTime = r[13];
      let eTime = r[14];
      if (sTime instanceof Date) sTime = Utilities.formatDate(sTime, tz, "HH:mm");
      if (eTime instanceof Date) eTime = Utilities.formatDate(eTime, tz, "HH:mm");

      return {
        rowId: index + 2,
        id: r[0],
        status: r[2],
        title: r[10] || "Untitled",
        ln: r[4] || "",
        fn: r[5] || "",
        mi: r[6] || "",
        idNumber: r[7] || "",
        email: r[3] || "",
        office: r[8] || "",
        contact: r[9] || "",
        room: r[12] || "",
        startTime: sTime ? String(sTime) : "",
        endTime: eTime ? String(eTime) : "",
        start: formatVal(r[15]),
        end: formatVal(r[16]) || formatVal(r[15]),
        purpose: r[11] || ""
      };
    });
  }

  return {
    rooms: rooms,
    reservations: reservations,
    currentUser: userEmail
  };
}

/**
 * Handle Form Submission
 */
function submitReservation(form) {
  if (!form) return { success: false, message: "Form data is missing." };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_RESERVATIONS);
  const userEmail = Session.getActiveUser().getEmail();

  const id = "RES-" + Utilities.getUuid().substring(0,8).toUpperCase();
  
  const startDate = form.startDate || "";
  const endDate = form.endDate || startDate; 

  // Column Mapping:
  // A: ID, B: Date, C: Status, D: Requester Email, E: LN, F: FN...
  sheet.appendRow([
    id, 
    new Date(), 
    "Pending", 
    form.email || userEmail, // Save the email provided in the form to Column D
    form.ln || "", 
    form.fn || "", 
    form.mi || "", 
    form.idNumber || "", 
    form.office || "", 
    form.contact || "", 
    form.title || "Untitled", 
    form.purpose || "", 
    form.room || "4Ps Conference Room", 
    form.startTime || "", 
    form.endTime || "",   
    startDate,
    endDate 
  ]);

  // Ensure data is written before notifications
  SpreadsheetApp.flush();

  // Initial notification to Admins
  sendDetailedAdminNotification(form, userEmail, id);
  return { success: true };
}

/**
 * Admin Update/Edit Function
 */
function updateReservation(formInput) {
  let form;
  try {
    if (typeof formInput === 'string') {
      form = JSON.parse(formInput);
    } else {
      form = formInput;
    }
  } catch (e) {
    console.error("JSON Parse Error in updateReservation:", e.toString());
    return { success: false, message: "Server Error: Invalid data format." };
  }

  if (!form || typeof form !== 'object') {
    return { success: false, message: "System Error: No valid data was received." };
  }

  if (!form.rowId) {
    return { success: false, message: "System Error: Missing reference row ID." };
  }

  const userEmail = Session.getActiveUser().getEmail();
  const allowedAdmins = ADMIN_NOTIFICATION_EMAILS.split(",").map(e => e.trim().toLowerCase());
  if (!allowedAdmins.includes(userEmail.toLowerCase())) {
    throw new Error("Unauthorized: Only admins can edit requests.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_RESERVATIONS);
  const rowIndex = parseInt(form.rowId);

  try {
    const val = (v) => (v === undefined || v === null) ? "" : String(v);

    sheet.getRange(rowIndex, 5).setValue(val(form.ln));
    sheet.getRange(rowIndex, 6).setValue(val(form.fn));
    sheet.getRange(rowIndex, 7).setValue(val(form.mi));
    sheet.getRange(rowIndex, 8).setValue(val(form.idNumber));
    sheet.getRange(rowIndex, 9).setValue(val(form.office));
    sheet.getRange(rowIndex, 10).setValue(val(form.contact));
    sheet.getRange(rowIndex, 11).setValue(val(form.title));
    sheet.getRange(rowIndex, 12).setValue(val(form.purpose));
    sheet.getRange(rowIndex, 14).setValue(val(form.startTime));
    sheet.getRange(rowIndex, 15).setValue(val(form.endTime));
    sheet.getRange(rowIndex, 16).setValue(val(form.startDate));
    sheet.getRange(rowIndex, 17).setValue(val(form.endDate || form.startDate));

    return { success: true, message: "Reservation updated successfully." };
  } catch (err) {
    return { success: false, message: "Update failed: " + err.toString() };
  }
}

/**
 * Update status via Email Links
 */
function updateReservationStatus(id, newStatus) {
  if (!id) return { success: false, message: "ID is missing." };
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_RESERVATIONS);
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) return { success: false, message: "No data in sheet." };

  const idColumnData = sheet.getRange(1, 1, lastRow, 1).getValues();
  const searchId = id.toString().trim();
  
  let rowIndex = -1;
  for (let i = 0; i < idColumnData.length; i++) {
    if (idColumnData[i][0] && idColumnData[i][0].toString().trim() === searchId) {
      rowIndex = i + 1;
      break;
    }
  }
  
  if (rowIndex === -1) return { success: false, message: "Request ID not found." };

  // Update sheet status
  sheet.getRange(rowIndex, 3).setValue(newStatus);
  SpreadsheetApp.flush(); // Ensure the status is committed before reading row data
  
  // Fetch detailed data for the confirmation email
  const rowDataArray = sheet.getRange(rowIndex, 1, 1, 17).getValues();
  const rowData = rowDataArray[0];
  
  // Column D (index 3) is where the email is saved during submitReservation
  const requesterEmail = rowData[3]; 
  const tz = Session.getScriptTimeZone();

  const formatVal = (val) => {
    if (val instanceof Date) return Utilities.formatDate(val, tz, "yyyy-MM-dd");
    return val ? String(val).split('T')[0] : "Not set";
  };

  const details = {
    id: searchId,
    title: rowData[10] || "Untitled Activity",
    room: rowData[12] || "4Ps Conference Room",
    schedule: `${formatVal(rowData[15])} (${rowData[13]} - ${rowData[14]})`,
    requesterName: `${rowData[5] || ""} ${rowData[4] || ""}`.trim() || "Valued Requester",
    office: rowData[8] || "N/A"
  };

  // Send notification to requester
  if (requesterEmail && requesterEmail.includes("@")) {
    sendRequesterResponse(requesterEmail, newStatus, details);
  } else {
    console.error("Invalid or missing requester email for ID: " + searchId);
  }
  
  return { success: true };
}

/**
 * Enhanced Notification to Admins with complete details and action links
 */
function sendDetailedAdminNotification(form, userEmail, id) {
  if (!form) return;
  
  const webAppUrl = ScriptApp.getService().getUrl();
  const approveUrl = `${webAppUrl}?action=Approved&id=${id}`;
  const rejectUrl = `${webAppUrl}?action=Rejected&id=${id}`;
  
  const title = form.title || "Untitled Activity";
  const office = form.office || "Unknown Office";
  const name = `${form.fn || ""} ${form.mi ? form.mi + ' ' : ''}${form.ln || ""}`.trim() || "Anonymous";

  const htmlBody = `
    <div style="font-family: 'Segoe UI', Arial, sans-serif; max-width: 600px; border: 1px solid #e2e8f0; border-radius: 12px; overflow: hidden; color: #1e293b; background-color: #ffffff;">
      <div style="background-color: #1e3a8a; padding: 30px; text-align: center;">
        <h2 style="color: #ffffff; margin: 0; font-size: 22px; font-weight: 800; letter-spacing: -0.5px;">New Reservation Request</h2>
        <p style="color: #bfdbfe; margin: 10px 0 0 0; font-size: 14px; font-weight: 600;">Ref ID: ${id || 'N/A'}</p>
      </div>
      
      <div style="padding: 30px;">
        <h3 style="color: #1e3a8a; border-bottom: 2px solid #f1f5f9; padding-bottom: 10px; font-size: 16px; margin-bottom: 15px; text-transform: uppercase; letter-spacing: 1px;">Activity Information</h3>
        <table style="width: 100%; font-size: 14px; border-collapse: collapse;">
          <tr><td style="padding: 8px 0; color: #64748b; width: 140px;">Activity Title:</td><td style="font-weight: 700; color: #0f172a;">${title}</td></tr>
          <tr><td style="padding: 8px 0; color: #64748b;">Schedule:</td><td style="font-weight: 700; color: #0f172a;">${form.startDate || 'N/A'} at ${form.startTime || '--'} - ${form.endTime || '--'}</td></tr>
          <tr><td style="padding: 8px 0; color: #64748b;">Purpose:</td><td style="color: #334155;">${form.purpose || 'None provided'}</td></tr>
        </table>

        <h3 style="color: #1e3a8a; border-bottom: 2px solid #f1f5f9; padding-bottom: 10px; font-size: 16px; margin: 30px 0 15px 0; text-transform: uppercase; letter-spacing: 1px;">Requester Details</h3>
        <table style="width: 100%; font-size: 14px; border-collapse: collapse;">
          <tr><td style="padding: 8px 0; color: #64748b; width: 140px;">Name:</td><td style="font-weight: 700; color: #0f172a;">${name}</td></tr>
          <tr><td style="padding: 8px 0; color: #64748b;">Office / Unit:</td><td style="font-weight: 700; color: #0f172a;">${office}</td></tr>
          <tr><td style="padding: 8px 0; color: #64748b;">ID Number:</td><td style="color: #334155;">${form.idNumber || 'N/A'}</td></tr>
          <tr><td style="padding: 8px 0; color: #64748b;">Contact No.:</td><td style="color: #334155;">${form.contact || 'N/A'}</td></tr>
          <tr><td style="padding: 8px 0; color: #64748b;">Email Address:</td><td style="color: #334155;">${form.email || userEmail || 'N/A'}</td></tr>
        </table>
      </div>

      <div style="padding: 30px; background-color: #f8fafc; text-align: center; border-top: 1px solid #e2e8f0;">
        <p style="margin-bottom: 25px; font-size: 14px; color: #475569; font-weight: 500;">Review this request and take action:</p>
        <div style="display: flex; justify-content: center; gap: 15px;">
          <a href="${approveUrl}" style="background-color: #10b981; color: #ffffff; padding: 14px 28px; text-decoration: none; border-radius: 10px; font-weight: 800; font-size: 14px; display: inline-block; transition: background 0.2s;">APPROVE</a>
          <a href="${rejectUrl}" style="background-color: #ef4444; color: #ffffff; padding: 14px 28px; text-decoration: none; border-radius: 10px; font-weight: 800; font-size: 14px; display: inline-block; transition: background 0.2s;">REJECT</a>
        </div>
        <p style="margin-top: 20px; font-size: 11px; color: #94a3b8;">Clicking these links will immediately update the reservation status.</p>
      </div>
    </div>
  `;

  MailApp.sendEmail({
    to: ADMIN_NOTIFICATION_EMAILS,
    subject: `[RESERVATION REQUEST] ${title} - ${office}`,
    htmlBody: htmlBody
  });
}

/**
 * Professional Notification to the Requester regarding Approval or Rejection
 */
function sendRequesterResponse(email, status, details) {
  if (!details || !email) return;
  
  const statusColor = status === 'Approved' ? '#10b981' : '#ef4444';
  const statusIcon = status === 'Approved' ? '✓' : '✕';
  const statusText = status === 'Approved' ? 'APPROVED' : 'REJECTED';
  const activityTitle = details.title || "Untitled Activity";
  
  const htmlBody = `
    <div style="font-family: 'Segoe UI', Arial, sans-serif; max-width: 600px; border: 1px solid #e2e8f0; border-radius: 12px; overflow: hidden; color: #1e293b; background-color: #ffffff;">
      <div style="background-color: ${statusColor}; padding: 30px; text-align: center;">
        <div style="display: inline-block; width: 40px; height: 40px; border: 3px solid #ffffff; border-radius: 50%; color: #ffffff; line-height: 36px; font-size: 24px; font-weight: bold; margin-bottom: 15px;">${statusIcon}</div>
        <h2 style="color: #ffffff; margin: 0; font-size: 22px; font-weight: 800; letter-spacing: 0.5px;">Reservation ${statusText}</h2>
        <p style="color: #ffffff; margin: 8px 0 0 0; opacity: 0.9; font-size: 14px;">Reference ID: ${details.id || 'N/A'}</p>
      </div>
      
      <div style="padding: 30px; background-color: #ffffff;">
        <p style="font-size: 15px; line-height: 1.6; color: #334155;">Good day, <strong>${details.requesterName || 'Requester'}</strong>,</p>
        <p style="font-size: 15px; line-height: 1.6; color: #334155;">Your reservation request for the 4Ps RPMO Conference Room has been <strong>${status ? status.toLowerCase() : 'processed'}</strong> by the management.</p>
        
        <div style="background-color: #f8fafc; padding: 20px; border-radius: 12px; margin: 25px 0; border-left: 5px solid ${statusColor};">
          <h4 style="margin: 0 0 10px 0; font-size: 14px; text-transform: uppercase; color: #64748b; letter-spacing: 0.5px;">Booking Details</h4>
          <p style="margin: 5px 0; font-size: 15px; color: #0f172a;"><strong>Activity:</strong> ${activityTitle}</p>
          <p style="margin: 5px 0; font-size: 15px; color: #0f172a;"><strong>Schedule:</strong> ${details.schedule || 'N/A'}</p>
          <p style="margin: 5px 0; font-size: 15px; color: #0f172a;"><strong>Location:</strong> ${details.room || '4Ps Conference Room'}</p>
        </div>

        ${status === 'Approved' ? 
          `<div style="background-color: #ecfdf5; padding: 15px; border-radius: 8px; color: #065f46; font-size: 14px;">
            <strong>Next Steps:</strong> Please coordinate to Mr. Jonald or 4Ps RPMO   at least 1 day before the activity for logistics and equipment setup.
            <strong style="color: #FF3131;">Your cooperation in observing CLAYGO at all times is HIGHLY APPRECIATED.</strong>
          </div>` : 
          `<div style="background-color: #fef2f2; padding: 15px; border-radius: 8px; color: #991b1b; font-size: 14px;">
            <strong>Note:</strong> If you believe this was an error or wish to reschedule, please contact Mr. Jonald or 4Ps RPMO  directly.
            
          </div>`}
        
        <p style="margin-top: 35px; font-size: 14px; color: #64748b; line-height: 1.6;">
          Best regards,<br>
          <strong style="color: #1e3a8a;">4Ps RPMO</strong>
        </p>
      </div>
      
      <div style="padding: 20px; background-color: #f1f5f9; text-align: center; color: #94a3b8; font-size: 11px;">
        This is an automated message. Please do not reply to this email.
      </div>
    </div>
  `;

  MailApp.sendEmail({ 
    to: email, 
    subject: `RESERVATION ${statusText}: ${activityTitle}`, 
    htmlBody: htmlBody 
  });
}
