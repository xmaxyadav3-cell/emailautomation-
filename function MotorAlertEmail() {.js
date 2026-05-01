function MotorAlertEmail() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // IMPORTANT: Change "Sheet1" to the exact name of your tab
  var sheet = ss.getSheetByName("JUNE'25"); 
  
  if (!sheet) {
    Logger.log("ERROR: Sheet not found. Check the sheet name in the code.");
    return;
  }

  var data = sheet.getDataRange().getValues();
  var recipient = "xmaxyadav3@gmail.com";

  for (var i = 1; i < data.length; i++) {
    // Column Mapping (0-based index)
    var orderno = data[i][2];    // Col C
    var partyname = data[i][3];  // Col D
    var motor = data[i][4];     // Col E
    var catRef = data[i][5];    // Col F
    var motorType = data[i][7];  // Col H
    var qty = data[i][8];        // Col I
    var qtypending = data[i][9]; // Col J
    var make = data[i][10];      // Col K
    var delhi = data[i][11];     // Col L
    var bilaspur = data[i][12];  // Col M
    
    // Status (Col S) - handle empty cells safely
    var finalstatus = data[i][18] ? data[i][18].toString().toLowerCase().trim() : "";
    
    // Due Days (Col V) - handle formulas/decimals
    var dueDays= data[i][24];

    // 🔍 DEBUG LOGGING: See what the script sees for every row
    Logger.log("Row " + (i+1) + " | Order: " + orderno + " | Status: '" + finalstatus + "' | DueDays: " + dueDays);

    var alertMsg = "";
    if (dueDays === 15) alertMsg = "⏳ 15 Days Remaining";
    else if (dueDays === 10) alertMsg = "⚠️ 10 Days Remaining";
    else if (dueDays === 5) alertMsg = "🚨 5 Days Remaining";
    else if (dueDays === 0) alertMsg = "📅 Due Today";
    else if (dueDays === -5) alertMsg = "❌ 5 Days OVERDUE";
    else if (dueDays === -10) alertMsg = "❌ 10 Days OVERDUE";
    else if (dueDays === -15) alertMsg = "❌ 15 Days OVERDUE";

    // ✅ CHECK CONDITIONS
    if (alertMsg !== "" && finalstatus === "processing") {
      
      var subject = "⚠ " + partyname + " | Order No: " + orderno + " | " + alertMsg;
       var htmlBody = "";

      htmlBody += "<p>Dear Team,</p>";

      htmlBody += "<p style='color:red; font-weight:bold;'>" + alertMsg + "</p>";

      htmlBody += "<table border='1' cellpadding='6' style='border-collapse:collapse; font-family:Arial;'>";

      htmlBody += "<tr><td><b>Party Name</b></td><td>" + partyname + "</td></tr>";
      htmlBody += "<tr><td><b>Order no.</b></td><td>" + orderno + "</td></tr>";
      htmlBody += "<tr><td><b>Motor</b></td><td>" + motor + "</td></tr>";
      htmlBody += "<tr><td><b>Cat Ref</b></td><td>" + catRef + "</td></tr>";
      htmlBody += "<tr><td><b>Motor Type</b></td><td>" + motorType + "</td></tr>";
      htmlBody += "<tr><td><b> Quantity</b></td><td>" + qty + "</td></tr>";
      htmlBody += "<tr><td><b> QtyPending</b></td><td><b>" + qtypending + "</b></td></tr>";
      htmlBody += "<tr><td><b>Make</b></td><td>" + make + "</td></tr>";
      htmlBody += "<tr><td><b>Delhi</b></td><td>" + delhi + "</td></tr>";
      htmlBody += "<tr><td><b>Bilaspur</b></td><td>" + bilaspur + "</td></tr>";
      htmlBody += "<tr><td><b>Due in</b></td><td style='color:red;'><b>" + dueDays + " days</b></td></tr>";

      htmlBody += "</table>";

      htmlBody += "<br><p>Please take necessary action immediately.</p>";

      htmlBody += "<p>Regards,<br><b>Order Tracking System</b></p>";

      try {
        MailApp.sendEmail({
          to: recipient,
          subject: subject,
          htmlBody: htmlBody
        });
        Logger.log("✅ EMAIL SENT for Row " + (i+1));
      } catch (e) {
        Logger.log("❌ ERROR sending email for Row " + (i+1) + ": " + e.message);
      }
    }
  }
}