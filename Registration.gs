/**
 * Copies the latest form response from "Registration" to "Directory".
 * Map columns as specified and format text/date.
 *
 * UPDATES (per request):
 * 1) Write to the NEXT AVAILABLE ROW where Directory Column C AND D are blank
 *    (because Column A/B are usually blank).
 * 2) Email sending: send to ANY valid email available (even if some emails are invalid).
 *    - Collect valid emails from: Registrant email (H), Inviter email (O), Config!G list
 *    - Use first valid as TO, the rest as CC (invalid emails are skipped)
 */
function onFormSubmit(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("New Registrant");
  if (!sheet) return;

  var directorySheet = ss.getSheetByName("Directory");
  if (!directorySheet) return;

  // Get the last submitted row from "New Registrant"
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return; // no data yet

  var lastCol = sheet.getLastColumn();
  var values = sheet.getRange(lastRow, 1, 1, lastCol).getValues()[0];

  // --- Helper: Title Case ---
  function toTitleCase_(text) {
    if (!text) return "";
    text = String(text).toLowerCase();
    return text.replace(/\b\w/g, function (c) {
      return c.toUpperCase();
    });
  }

  // --- Helper: matching utilities for Directory dedup ---
  function cleanString_(s) {
    if (!s) return "";
    return String(s).toLowerCase().replace(/[^a-z0-9]/g, "");
  }

  function areAlmostSame_(a, b) {
    if (!a || !b) return false;
    a = String(a);
    b = String(b);
    var la = a.length;
    var lb = b.length;
    if (Math.abs(la - lb) > 1) return false;
    var i = 0;
    var j = 0;
    var diff = 0;
    while (i < la && j < lb) {
      if (a.charAt(i) === b.charAt(j)) {
        i++;
        j++;
      } else {
        diff++;
        if (diff > 1) return false;
        if (la > lb) {
          i++;
        } else if (lb > la) {
          j++;
        } else {
          i++;
          j++;
        }
      }
    }
    if (i < la || j < lb) diff++;
    return diff <= 1;
  }

  function namesSimilar_(first1, last1, first2, last2) {
    var f1 = cleanString_(first1);
    var f2 = cleanString_(first2);
    var l1 = cleanString_(last1);
    var l2 = cleanString_(last2);
    if (!l1 || !l2) return false;
    if (l1 !== l2) return false; // last name must match after cleaning

    if (!f1 || !f2) return false;
    if (f1 === f2) return true;
    if (f1.indexOf(f2) === 0 || f2.indexOf(f1) === 0) return true; // handles middle names
    return areAlmostSame_(f1, f2); // allow 1-letter difference
  }

  function normalizePhone_(p) {
    if (!p) return "";
    return String(p).replace(/\D/g, "");
  }

  function phonesEqual_(p1, p2) {
    var s1 = normalizePhone_(p1);
    var s2 = normalizePhone_(p2);
    if (!s1 || !s2) return false;
    return s1 === s2;
  }

  function emailsEqual_(e1, e2) {
    if (!e1 || !e2) return false;
    return String(e1).toLowerCase().trim() === String(e2).toLowerCase().trim();
  }

  // --- Helper: find next available row based on blank Column C & D ---
  function getNextAvailableRowByBlankCD_(sh, startRow) {
    startRow = startRow || 2;
    var last = sh.getLastRow();
    if (last < startRow) return startRow;

    // Read C:D
    var cdVals = sh.getRange(startRow, 3, last - startRow + 1, 2).getValues();
    for (var i = 0; i < cdVals.length; i++) {
      var c = cdVals[i][0];
      var d = cdVals[i][1];
      var cBlank = (c === "" || c === null);
      var dBlank = (d === "" || d === null);
      if (cBlank && dBlank) {
        return startRow + i;
      }
    }
    return last + 1;
  }

  // --- Helper: ensure row exists (avoid "outside dimensions") ---
  function ensureRowExists_(sh, row) {
    var max = sh.getMaxRows();
    if (row > max) {
      sh.insertRowsAfter(max, row - max);
    }
  }

  // --- Helper: email validation + unique list ---
  function isValidEmail_(s) {
    if (!s) return false;
    s = String(s).trim();
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s);
  }

  function pushUniqueValid_(arr, emailStr) {
    if (!emailStr) return;
    var e = String(emailStr).trim().toLowerCase();
    if (!isValidEmail_(e)) return;
    if (arr.indexOf(e) === -1) arr.push(e);
  }

  // --- Get values from "Registration" ---
  var timestamp = values[0];
  var firstName = values[2];
  var lastName = values[3];
  var gender = values[4];
  var birthdate = values[5];
  var phone = values[6];
  var email = values[7];          // Column H – main recipient (To)
  var streetAddress = values[8];
  var city = values[9];
  var state = values[10];
  var zipCode = values[11];
  var inviterFirstName = values[12];
  var inviterLastName = values[13];
  var inviterEmail = values[14];  // Column O – to be CC'd

  // Normalize
  firstName = toTitleCase_(firstName);
  lastName = toTitleCase_(lastName);
  gender = toTitleCase_(gender);
  streetAddress = toTitleCase_(streetAddress);
  city = toTitleCase_(city);
  inviterFirstName = toTitleCase_(inviterFirstName);
  inviterLastName = toTitleCase_(inviterLastName);

  //
  //  FIND IF THIS PERSON ALREADY EXISTS IN "Directory"
  //
  var isExistingMatch = false;
  var existingRow = -1;

  var dirLastRow = directorySheet.getLastRow();
  if (dirLastRow >= 2) {
    var dirLastCol = directorySheet.getLastColumn();
    var dirData = directorySheet.getRange(2, 1, dirLastRow - 1, dirLastCol).getValues();

    var tzMatch = ss.getSpreadsheetTimeZone();
    var regBirthKey = "";
    if (birthdate instanceof Date && !isNaN(birthdate.getTime())) {
      regBirthKey = Utilities.formatDate(birthdate, tzMatch, "yyyy-MM-dd");
    } else if (birthdate) {
      var bdTemp = new Date(birthdate);
      if (bdTemp instanceof Date && !isNaN(bdTemp.getTime())) {
        regBirthKey = Utilities.formatDate(bdTemp, tzMatch, "yyyy-MM-dd");
      }
    }

    var bestScore = 0;
    var bestRowIndex = -1;

    for (var i = 0; i < dirData.length; i++) {
      var row = dirData[i];

      var dirLastName = row[2];  // Column C
      var dirFirstName = row[3]; // Column D
      var dirBirth = row[6];     // Column G
      var dirPhone = row[14];    // Column O
      var dirEmail = row[15];    // Column P

      var dirBirthKey = "";
      if (dirBirth instanceof Date && !isNaN(dirBirth.getTime())) {
        dirBirthKey = Utilities.formatDate(dirBirth, tzMatch, "yyyy-MM-dd");
      }

      // RULE: if both have birthdate and they are different, treat as different person
      if (regBirthKey && dirBirthKey && regBirthKey !== dirBirthKey) {
        continue;
      }

      // First, require names to be "similar"
      if (!namesSimilar_(firstName, lastName, dirFirstName, dirLastName)) {
        continue;
      }

      var score = 2; // base score if name is similar

      if (regBirthKey && dirBirthKey && regBirthKey === dirBirthKey) {
        score += 1;
      }
      if (phonesEqual_(phone, dirPhone)) {
        score += 1;
      }
      if (emailsEqual_(email, dirEmail)) {
        score += 1;
      }

      if (score > bestScore) {
        bestScore = score;
        bestRowIndex = i;
      }
    }

    // Threshold: name similar (2 points) + at least one of birth/phone/email (score >= 3)
    if (bestScore >= 3 && bestRowIndex >= 0) {
      isExistingMatch = true;
      existingRow = 2 + bestRowIndex; // convert index in dirData back to sheet row
    }
  }

  // Target row logic:
  // - If existing match: update the existing row (NO moving)
  // - If new person: write to NEXT AVAILABLE ROW where Column C & D are blank
  var targetRow;
  if (isExistingMatch && existingRow > 0) {
    targetRow = existingRow;
  } else {
    targetRow = getNextAvailableRowByBlankCD_(directorySheet, 2);
  }
  ensureRowExists_(directorySheet, targetRow);

  // Convert timestamp to US spreadsheet timezone
  var dateValue = "";
  if (timestamp instanceof Date) {
    var tz = ss.getSpreadsheetTimeZone();
    dateValue = Utilities.formatDate(timestamp, tz, "MM/dd/yyyy");
  } else if (timestamp) {
    var tempDate = new Date(timestamp);
    var tz2 = ss.getSpreadsheetTimeZone();
    dateValue = Utilities.formatDate(tempDate, tz2, "MM/dd/yyyy");
  }

  // U (21) – Date – CENTER
  // Do NOT change col U if this is an existing match
  if (dateValue && !isExistingMatch) {
    directorySheet
      .getRange(targetRow, 21)
      .setValue(dateValue)
      .setNumberFormat("MM/dd/yyyy")
      .setVerticalAlignment("middle")
      .setHorizontalAlignment("center");
  }

  // D (4) – First Name – LEFT
  directorySheet
    .getRange(targetRow, 4)
    .setValue(firstName)
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("left");

  // C (3) – Last Name – LEFT
  directorySheet
    .getRange(targetRow, 3)
    .setValue(lastName)
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("left");

  // E (5) – Gender – CENTER
  directorySheet
    .getRange(targetRow, 5)
    .setValue(gender)
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("center");

  // G (7) – Birthdate – CENTER
  var birthdateObj;
  if (birthdate instanceof Date) birthdateObj = birthdate;
  else if (birthdate) birthdateObj = new Date(birthdate);
  else birthdateObj = "";

  if (birthdateObj instanceof Date && !isNaN(birthdateObj.getTime())) {
    directorySheet
      .getRange(targetRow, 7)
      .setValue(birthdateObj)
      .setNumberFormat("MM/dd/yyyy")
      .setVerticalAlignment("middle")
      .setHorizontalAlignment("center");
  }

  // O (15) – Phone – CENTER
  directorySheet
    .getRange(targetRow, 15)
    .setValue(phone)
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("center");

  // P (16) – Email – LEFT
  directorySheet
    .getRange(targetRow, 16)
    .setValue(email)
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("left");

  // Q (17) – Street Address – LEFT
  directorySheet
    .getRange(targetRow, 17)
    .setValue(streetAddress)
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("left");

  // R (18) – City – LEFT
  directorySheet
    .getRange(targetRow, 18)
    .setValue(city)
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("left");

  // S (19) – State – CENTER
  directorySheet
    .getRange(targetRow, 19)
    .setValue(state)
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("center");

  // T (20) – Zip Code – CENTER
  directorySheet
    .getRange(targetRow, 20)
    .setValue(zipCode)
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("center");

  // V (22) – Inviter Last Name – LEFT
  directorySheet
    .getRange(targetRow, 22)
    .setValue(inviterLastName)
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("left");

  // W (23) – Inviter First Name – LEFT
  directorySheet
    .getRange(targetRow, 23)
    .setValue(inviterFirstName)
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("left");

  // X (24) – Inviter Email – LEFT
  directorySheet
    .getRange(targetRow, 24)
    .setValue(inviterEmail)
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("left");

  // ==============================
  //   SEND CONFIRMATION EMAIL
  // ==============================
  try {
    var configSheet = ss.getSheetByName("Config");
    if (configSheet) {
      var subject = configSheet.getRange("E2").getDisplayValue();
      var htmlTemplate = configSheet.getRange("E3").getValue(); // HTML content

      // Header image: allow raw Drive link or direct URL from E4
      var rawHeaderUrl = configSheet.getRange("E4").getDisplayValue();
      var headerImageUrl = "";
      if (rawHeaderUrl && rawHeaderUrl.indexOf("drive.google.com") !== -1) {
        var match = rawHeaderUrl.match(/\/d\/(.*?)\//);
        if (match && match[1]) {
          var fileId = match[1];
          headerImageUrl = "https://drive.google.com/uc?export=view&id=" + fileId;
        }
      } else {
        headerImageUrl = rawHeaderUrl;
      }

      var senderName = configSheet.getRange("E5").getDisplayValue();

      if (!subject) subject = "Welcome to our community!";

      // Read dynamic fields for placeholders
      var pastorName    = configSheet.getRange("H2").getDisplayValue();
      var pastorTitle   = configSheet.getRange("I2").getDisplayValue();
      var communityName = configSheet.getRange("E5").getDisplayValue();
      var address       = configSheet.getRange("J2").getDisplayValue();
      var officeEmail   = configSheet.getRange("G2").getDisplayValue();

      var support1Name  = configSheet.getRange("H3").getDisplayValue();
      var support1Role  = configSheet.getRange("I3").getDisplayValue();
      var support1Email = configSheet.getRange("G3").getDisplayValue();

      var support2Name  = configSheet.getRange("H4").getDisplayValue();
      var support2Role  = configSheet.getRange("I4").getDisplayValue();
      var support2Email = configSheet.getRange("G4").getDisplayValue();

      var htmlBody = String(htmlTemplate)
        .replace(/{{FirstName}}/g, firstName)
        .replace(/{{LastName}}/g, lastName)
        .replace(/{{PastorName}}/g, pastorName)
        .replace(/{{PastorTitle}}/g, pastorTitle)
        .replace(/{{CommunityName}}/g, communityName)
        .replace(/{{Address}}/g, address)
        .replace(/{{OfficeEmail}}/g, officeEmail)
        .replace(/{{Support1Name}}/g, support1Name)
        .replace(/{{Support1Role}}/g, support1Role)
        .replace(/{{Support1Email}}/g, support1Email)
        .replace(/{{Support2Name}}/g, support2Name)
        .replace(/{{Support2Role}}/g, support2Role)
        .replace(/{{Support2Email}}/g, support2Email);

      if (headerImageUrl) {
        var headerHtml =
          '<div style="text-align:center;margin-bottom:16px;">' +
          '<img src="' + headerImageUrl + '" style="max-width:100%;height:auto;" />' +
          "</div>";
        htmlBody = headerHtml + htmlBody;
      }

      htmlBody =
        htmlBody +
        '<p style="margin-top:24px;font-size:12px;color:#555;">' +
        "Please do not reply directly to this email." +
        "</p>";

      // Build recipient list from ANY valid emails available
      var recipients = [];
      pushUniqueValid_(recipients, email);        // registrant
      pushUniqueValid_(recipients, inviterEmail); // inviter

      var configLastRow = configSheet.getLastRow();
      if (configLastRow > 1) {
        var ccData = configSheet.getRange(2, 7, configLastRow - 1, 1).getValues();
        for (var j = 0; j < ccData.length; j++) {
          pushUniqueValid_(recipients, ccData[j][0]);
        }
      }

      if (recipients.length > 0) {
        var toEmail = recipients[0];
        var ccEmails = recipients.slice(1);

        var options = { htmlBody: htmlBody };
        if (senderName) options.name = senderName;
        if (ccEmails.length > 0) options.cc = ccEmails.join(",");

        GmailApp.sendEmail(toEmail, subject, "", options);
      }
    }
  } catch (err) {
    Logger.log("Error sending confirmation email: " + err);
  }
}
