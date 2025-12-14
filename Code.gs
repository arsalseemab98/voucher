/**
 * ============================================
 * CODE.gs - Main Entry Points & Processing
 * ============================================
 *
 * This file contains:
 * - Web app handlers (doGet, doPost)
 * - Form submission handler
 * - Approval/Rejection processing
 * - Payment processing
 * - PDF generation
 * - Admin portal & backend functions
 * - Test/Debug utilities
 * - Reset/Backup functions
 */

// ============ WEB APP HANDLERS ============

function doGet(e) {
  Logger.log("=== doGet STARTED ===");

  var debugInfo = [];
  debugInfo.push("doGet started at: " + new Date().toISOString());

  try {
    debugInfo.push("Step 1: Checking parameters");
    var action = "";
    if (e && e.parameter && e.parameter.action) {
      action = e.parameter.action;
    }
    debugInfo.push("Action: " + action);
    Logger.log("Action received: " + action);

    // Admin action
    if (action === "admin") {
      debugInfo.push("Step 2: Admin action detected - showing admin portal");
      Logger.log("Showing admin portal");
      return createHtmlResponse(getAdminPortalPage());
    }

    // Validate parameters
    if (!e || !e.parameter) {
      return createHtmlResponse(getHomePage());
    }

    const voucherId = e.parameter.id || "";
    const approver = e.parameter.approver || "";

    // Approve/Reject actions (direct access - no PIN required)
    if ((action === "approve" || action === "reject") && voucherId && approver) {
      debugInfo.push("Step 3: Approve/Reject action - showing page directly");

      if (action === "approve") {
        return createHtmlResponse(getSignaturePage(voucherId, approver));
      } else {
        return createHtmlResponse(getRejectPage(voucherId, approver));
      }
    }

    // Payment page requires email access check
    if (action === "payment" && voucherId) {
      const accessCheck = checkEmailAccess();
      if (!accessCheck.allowed) {
        return createHtmlResponse(getAccessDeniedPage(accessCheck.email));
      }
      return createHtmlResponse(getMuhtamimMaalPaymentPage(voucherId));
    }

    return createHtmlResponse(getHomePage());
  } catch (error) {
    Logger.log("ERROR in doGet: " + error.toString() + " | Stack: " + error.stack);

    function escapeForHtml(str) {
      return String(str || '').replace(/&/g,'&amp;').replace(new RegExp('<','g'),'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
    }
    var errorHtml = '<!DOCTYPE html><html><head><title>ERROR<\/title><\/head><body style="font-family:Arial;padding:40px;background:#ffe6e6;">';
    errorHtml += '<h1 style="color:red;">Server Error<\/h1>';
    errorHtml += '<div style="background:white;padding:20px;border-radius:8px;">';
    errorHtml += '<p><strong>Error:<\/strong> ' + escapeForHtml(error.toString()) + '<\/p>';
    errorHtml += '<p><strong>Stack:<\/strong><\/p>';
    errorHtml += '<pre style="background:#f5f5f5;padding:10px;overflow:auto;">' + escapeForHtml(error.stack || 'No stack trace') + '<\/pre>';
    errorHtml += '<h3>Debug Log:<\/h3>';
    errorHtml += '<pre>' + escapeForHtml(debugInfo.join('\n')) + '<\/pre>';
    errorHtml += '<\/div>';
    errorHtml += '<\/body><\/html>';

    return HtmlService.createHtmlOutput(errorHtml).setTitle('ERROR').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

function createHtmlResponse(html) {
  return HtmlService.createHtmlOutput(html)
    .setTitle('MKA Voucher System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  Logger.log("=== doPost STARTED ===");

  try {
    var body = {};
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }

    Logger.log("POST body: " + JSON.stringify(body));
    var action = body.action || "";

    // Process approval
    if (action === "processApproval") {
      var result = processApprovalFromWeb(body.voucherId, body.approverKey, body.signature);
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Process rejection
    if (action === "processRejection") {
      var result = processRejectionFromWeb(body.voucherId, body.approverKey, body.reason);
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Get voucher data (for Vercel frontend)
    if (action === "getVoucherData") {
      try {
        var voucher = getVoucherById(body.voucherId);
        if (!voucher) {
          return ContentService.createTextOutput(JSON.stringify({
            success: false, message: "Voucher not found"
          })).setMimeType(ContentService.MimeType.JSON);
        }

        var approverInfo = getApproverInfo(body.approverKey);
        if (!approverInfo) {
          return ContentService.createTextOutput(JSON.stringify({
            success: false, message: "Invalid approver"
          })).setMimeType(ContentService.MimeType.JSON);
        }

        var alreadyApproved = isAlreadyApproved(voucher, body.approverKey);
        var isTurn = isApproversTurn(voucher, body.approverKey);

        return ContentService.createTextOutput(JSON.stringify({
          success: true,
          voucher: {
            id: voucher.id,
            claimerName: voucher.claimedBy,
            amount: voucher.amountInNumbers,
            purpose: voucher.category,
            description: voucher.descriptionOfProgram,
            status: voucher.status,
            date: voucher.timestamp
          },
          approver: {
            title: approverInfo.title,
            key: body.approverKey
          },
          alreadyApproved: alreadyApproved,
          isYourTurn: isTurn
        })).setMimeType(ContentService.MimeType.JSON);
      } catch (err) {
        return ContentService.createTextOutput(JSON.stringify({
          success: false, message: "Error: " + err.toString()
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // Unknown action
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: "Unknown action: " + action
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log("ERROR in doPost: " + error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: "Server error: " + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}


// ============ FORM SUBMISSION HANDLER ============

function onFormSubmit(e) {
  if (!e || !e.values) {
    throw new Error("Use testFullWorkflow() to test, or submit the Google Form.");
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);

    const responses = e.values;
    // Form field order:
    // 0: Timestamp
    // 1: Email Address (auto-collected by Google Form)
    // 2: Your Full Name
    // 3: Your Email Address (manually entered)
    // 4: Description of Program
    // 5: Amount in Numbers
    // 6: Amount in Letters
    // 7: Category
    // 8: Majlis
    // 9: Receipt/Attachment
    const voucherData = {
      id: generateVoucherId(),
      timestamp: responses[0],
      claimedBy: responses[2],
      claimerEmail: responses[3],
      descriptionOfProgram: responses[4],
      amountInNumbers: responses[5],
      amountInLetters: responses[6],
      category: responses[7],
      majlis: responses[8],
      attachmentUrl: responses[9] || "",
      status: "pending_qaid"
    };

    saveVoucher(voucherData);

    // Share attachment with anyone who has the link
    if (voucherData.attachmentUrl && voucherData.attachmentUrl.trim() !== '') {
      shareAttachmentWithAnyone(voucherData.attachmentUrl);
      Logger.log("Attachment shared for voucher " + voucherData.id);
    }

    // Send approval request (no PDF link - generated at final step)
    sendApprovalRequest(voucherData, "qaid", "");
    sendClaimerConfirmation(voucherData, "");

    Logger.log("Voucher " + voucherData.id + " created successfully!");
  } finally {
    lock.releaseLock();
  }
}


// ============ ACCESS CONTROL ============

function checkEmailAccess() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      return { allowed: false, email: null, message: "Unable to verify your identity. Please make sure you are logged in to Google." };
    }
    const allowedList = getAllowedEmails();
    const isAllowed = allowedList.some(email => email.toLowerCase() === userEmail.toLowerCase());
    return { allowed: isAllowed, email: userEmail, message: isAllowed ? "Access granted" : "Access denied for " + userEmail };
  } catch (error) {
    Logger.log("Error checking email access: " + error);
    return { allowed: false, email: null, message: "Error verifying access: " + error.toString() };
  }
}


// ============ PROCESS APPROVAL ============

function processApprovalFromWeb(voucherId, approverKey, signature) {
  const lock = LockService.getScriptLock();

  try {
    if (!lock.tryLock(10000)) {
      return { success: false, message: "Another approval is being processed. Please wait a moment and try again." };
    }

    const voucher = getVoucherById(voucherId);

    if (!voucher) {
      return { success: false, message: "Voucher not found. It may have been deleted." };
    }

    if (isAlreadyApproved(voucher, approverKey)) {
      return { success: false, message: "This voucher has already been approved by this role." };
    }

    if (!isApproversTurn(voucher, approverKey)) {
      return { success: false, message: "It's not your turn to approve. Please wait for previous approvers." };
    }

    const approver = getApproverInfo(approverKey);
    if (!approver) {
      return { success: false, message: "Invalid approver." };
    }

    const now = new Date().toISOString();
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName("Vouchers");

    const columns = {
      qaid: { approved: 12, signature: 13, date: 14 },
      nazimMaal: { approved: 15, signature: 16, date: 17 },
      sadarMka: { approved: 18, signature: 19, date: 20 }
    };

    const col = columns[approverKey];
    if (!col) {
      return { success: false, message: "Invalid approver configuration." };
    }

    // Update database
    sheet.getRange(voucher.row, col.approved).setValue(true);
    sheet.getRange(voucher.row, col.signature).setValue(signature ? signature.substring(0, 50000) : "");
    sheet.getRange(voucher.row, col.date).setValue(now);
    sheet.getRange(voucher.row, 11).setValue(approver.statusAfter);
    SpreadsheetApp.flush();

    // Update local voucher object
    voucher.status = approver.statusAfter;
    if (approverKey === "qaid") voucher.qaidApproval = { approved: true, date: now };
    if (approverKey === "nazimMaal") voucher.nazimApproval = { approved: true, date: now };
    if (approverKey === "sadarMka") voucher.sadarApproval = { approved: true, date: now };

    sendStatusUpdateToClaimer(voucher, approver.title);

    // Handle next step
    if (approver.statusAfter === "approved") {
      sendFinalNotification(voucher, "");
      return { success: true, message: "All approvals complete! Sent to Muhtamim Maal for payment." };
    } else {
      const nextApprover = getApprovalFlow().find(a => a.statusBefore === approver.statusAfter);
      if (nextApprover) {
        sendApprovalRequest(voucher, nextApprover.key, "");
        return { success: true, message: "Sent to " + nextApprover.title + " for approval." };
      }
      return { success: true, message: "Approval recorded." };
    }

  } catch (error) {
    Logger.log("Error in processApprovalFromWeb: " + error);
    return { success: false, message: "An error occurred: " + error.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// Alias for doPost
function processRejectionFromWeb(voucherId, approverKey, reason) {
  return rejectVoucher(voucherId, approverKey, reason);
}


// ============ REJECT VOUCHER ============

function rejectVoucher(voucherId, approverKey, reason) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);

    const voucher = getVoucherById(voucherId);
    const approver = getApproverInfo(approverKey);

    if (!voucher) {
      return { success: false, message: "Voucher not found." };
    }

    if (!approver) {
      return { success: false, message: "Invalid approver." };
    }

    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName("Vouchers");
    sheet.getRange(voucher.row, 11).setValue("rejected");
    SpreadsheetApp.flush();

    // Send rejection email to claimer
    const subject = "Voucher REJECTED: " + voucherId;
    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: #dc3545; color: white; padding: 20px; text-align: center; border-radius: 10px 10px 0 0;">
          <h1 style="margin: 0;">Voucher Rejected<\/h1>
        <\/div>
        <div style="padding: 25px; background: #fff; border: 1px solid #e0e0e0; border-radius: 0 0 10px 10px;">
          <p>Assalamu Alaikum <strong>${voucher.claimedBy}<\/strong>,<\/p>
          <p>Your voucher <strong>${voucherId}<\/strong> has been rejected by <strong>${approver.title}<\/strong>.<\/p>
          <div style="background: #f8d7da; padding: 15px; border-radius: 8px; margin: 20px 0;">
            <p style="margin: 0; color: #721c24;"><strong>Reason:<\/strong><\/p>
            <p style="margin: 10px 0 0 0; color: #721c24;">${reason}<\/p>
          <\/div>
          <p>If you have questions or wish to resubmit, please contact ${approver.title}.<\/p>
          <p>JazakAllah,<br>MKA Finance Team<\/p>
        <\/div>
      <\/div>
    `;

    GmailApp.sendEmail(voucher.claimerEmail, subject, "Please view this email in HTML format.", { htmlBody: htmlBody });

    return { success: true };
  } catch (error) {
    Logger.log("Error in rejectVoucher: " + error);
    return { success: false, message: error.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


// ============ PROCESS PAYMENT ============

function processPaymentFromWeb(voucherId, paymentMethod, accountInfo, signature) {
  const lock = LockService.getScriptLock();

  try {
    if (!lock.tryLock(10000)) {
      return { success: false, message: "Another operation is in progress. Please try again." };
    }

    const voucher = getVoucherById(voucherId);

    if (!voucher) {
      return { success: false, message: "Voucher not found." };
    }

    if (voucher.status !== "approved") {
      return { success: false, message: "Voucher is not yet approved." };
    }

    if (voucher.payment && voucher.payment.paid) {
      return { success: false, message: "This voucher has already been paid." };
    }

    const now = new Date().toISOString();
    const paymentDate = new Date().toLocaleDateString('sv-SE');
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName("Vouchers");

    // Update Muhtamim Maal payment columns
    sheet.getRange(voucher.row, 22).setValue(true);  // Muhtamim Maal Reviewed
    sheet.getRange(voucher.row, 23).setValue(signature ? signature.substring(0, 50000) : "");
    sheet.getRange(voucher.row, 24).setValue(now);
    sheet.getRange(voucher.row, 25).setValue(true);  // Paid
    sheet.getRange(voucher.row, 26).setValue("");    // Paid By
    sheet.getRange(voucher.row, 27).setValue(paymentDate);
    sheet.getRange(voucher.row, 28).setValue(paymentMethod);
    sheet.getRange(voucher.row, 29).setValue("");    // Paid To
    sheet.getRange(voucher.row, 30).setValue("");    // Bank Name
    sheet.getRange(voucher.row, 31).setValue(accountInfo || "");
    sheet.getRange(voucher.row, 32).setValue("");    // Clearing Number
    sheet.getRange(voucher.row, 11).setValue("paid");

    SpreadsheetApp.flush();

    const updatedVoucher = getVoucherById(voucherId);

    // Generate FINAL PDF
    var pdfUrl = "";
    try {
      pdfUrl = generateFinalVoucher(updatedVoucher);
      Logger.log("Final PDF generated: " + pdfUrl);

      if (pdfUrl) {
        sheet.getRange(voucher.row, 21).setValue(pdfUrl);
        SpreadsheetApp.flush();
      }
    } catch (pdfError) {
      Logger.log("Error generating final PDF: " + pdfError);
    }

    sendPaymentConfirmationToClaimer(updatedVoucher, paymentMethod, paymentDate, pdfUrl, accountInfo);

    return { success: true, message: "Payment registered and PDF generated." };
  } catch (error) {
    Logger.log("Error in processPaymentFromWeb: " + error);
    return { success: false, message: "Error: " + error.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


// ============ PDF GENERATION ============

function isImageMimeType(mimeType) {
  return mimeType && (
    mimeType.startsWith('image/') ||
    mimeType === 'image/jpeg' ||
    mimeType === 'image/png' ||
    mimeType === 'image/gif' ||
    mimeType === 'image/webp'
  );
}

function generateFinalVoucher(voucher) {
  try {
    const fullVoucher = getVoucherById(voucher.id);
    const doc = DocumentApp.create("Voucher_" + voucher.id);
    const body = doc.getBody();

    body.setMarginTop(40);
    body.setMarginBottom(40);
    body.setMarginLeft(50);
    body.setMarginRight(50);

    // Header table
    const headerTable = body.appendTable();
    headerTable.setBorderWidth(0);
    const headerRow = headerTable.appendTableRow();

    const leftCell = headerRow.appendTableCell();
    leftCell.appendParagraph("Majlis Khuddam-ul-Ahmadiyya").setBold(true).setFontSize(12);
    leftCell.appendParagraph("Tolvskillingsgatan 1").setFontSize(9);
    leftCell.appendParagraph("414 82 Göteborg, Sverige").setFontSize(9);
    leftCell.setWidth(280);

    const rightCell = headerRow.appendTableCell();
    rightCell.appendParagraph("MKA").setAlignment(DocumentApp.HorizontalAlignment.RIGHT).setBold(true).setFontSize(24);
    rightCell.appendParagraph("MAJLIS KHUDDAM-UL-AHMADIYYA SVERIGE").setAlignment(DocumentApp.HorizontalAlignment.RIGHT).setFontSize(7).setBold(true);

    body.appendParagraph("");

    // Reference
    const refTable = body.appendTable();
    refTable.setBorderWidth(1);
    refTable.appendTableRow().appendTableCell().appendParagraph("Reference: " + voucher.id).setBold(true).setFontSize(10);

    body.appendParagraph("");

    // Title
    body.appendParagraph("Voucher for Majlis Khuddam-ul-Ahmadiyya Swedens")
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER).setBold(true).setFontSize(14);

    body.appendParagraph("");

    // Details
    body.appendParagraph("Description: " + (voucher.descriptionOfProgram || "N/A")).setFontSize(10);
    body.appendParagraph("Amount: " + (voucher.amountInNumbers || "N/A") + " SEK (" + (voucher.amountInLetters || "") + ")").setFontSize(10).setBold(true);
    body.appendParagraph("Category: " + (voucher.category || "N/A") + "  |  Majlis: " + (voucher.majlis || "N/A")).setFontSize(10);

    body.appendParagraph("");

    function fmtDate(d) {
      if (!d) return "";
      try { return new Date(d).toLocaleDateString('sv-SE'); } catch(e) { return String(d); }
    }

    // Approval table with signatures
    const appTable = body.appendTable();
    appTable.setBorderWidth(1);

    var r1 = appTable.appendTableRow();
    var c1a = r1.appendTableCell(); c1a.appendParagraph("Claimed by:").setBold(true).setFontSize(9); c1a.appendParagraph(voucher.claimedBy || "").setFontSize(9); c1a.setWidth(230);
    var c1b = r1.appendTableCell(); c1b.appendParagraph("Date of Claim:").setBold(true).setFontSize(9); c1b.appendParagraph(fmtDate(voucher.timestamp)).setFontSize(9);

    // Qaid Majlis row with signature
    var r2 = appTable.appendTableRow();
    var c2a = r2.appendTableCell();
    c2a.appendParagraph("Qaid Majlis:").setBold(true).setFontSize(9);
    if (fullVoucher && fullVoucher.qaidApproval && fullVoucher.qaidApproval.approved) {
      c2a.appendParagraph("Approved - " + fmtDate(fullVoucher.qaidApproval.date)).setFontSize(9);
      if (fullVoucher.qaidApproval.signature) {
        var sigBlob = signatureToBlob(fullVoucher.qaidApproval.signature, "qaid_sig");
        if (sigBlob) {
          var sigImg = c2a.appendImage(sigBlob);
          sigImg.setWidth(100).setHeight(40);
        }
      }
    } else {
      c2a.appendParagraph("Pending").setFontSize(9);
    }
    c2a.setWidth(230);

    // Nazim Maal with signature
    var c2b = r2.appendTableCell();
    c2b.appendParagraph("Nazim Maal:").setBold(true).setFontSize(9);
    if (fullVoucher && fullVoucher.nazimApproval && fullVoucher.nazimApproval.approved) {
      c2b.appendParagraph("Approved - " + fmtDate(fullVoucher.nazimApproval.date)).setFontSize(9);
      if (fullVoucher.nazimApproval.signature) {
        var sigBlob2 = signatureToBlob(fullVoucher.nazimApproval.signature, "nazim_sig");
        if (sigBlob2) {
          var sigImg2 = c2b.appendImage(sigBlob2);
          sigImg2.setWidth(100).setHeight(40);
        }
      }
    } else {
      c2b.appendParagraph("Pending").setFontSize(9);
    }

    // Sadar MKA with signature
    var r3 = appTable.appendTableRow();
    var c3a = r3.appendTableCell();
    c3a.appendParagraph("Sadar MKA:").setBold(true).setFontSize(9);
    if (fullVoucher && fullVoucher.sadarApproval && fullVoucher.sadarApproval.approved) {
      c3a.appendParagraph("Approved - " + fmtDate(fullVoucher.sadarApproval.date)).setFontSize(9);
      if (fullVoucher.sadarApproval.signature) {
        var sigBlob3 = signatureToBlob(fullVoucher.sadarApproval.signature, "sadar_sig");
        if (sigBlob3) {
          var sigImg3 = c3a.appendImage(sigBlob3);
          sigImg3.setWidth(100).setHeight(40);
        }
      }
    } else {
      c3a.appendParagraph("Pending").setFontSize(9);
    }
    c3a.setWidth(230);

    // Muhtamim Maal with signature
    var c3b = r3.appendTableCell();
    c3b.appendParagraph("Muhtamim Maal:").setBold(true).setFontSize(9);
    if (fullVoucher && fullVoucher.muhtamimMaalApproval && fullVoucher.muhtamimMaalApproval.approved) {
      c3b.appendParagraph("Paid - " + fmtDate(fullVoucher.muhtamimMaalApproval.date)).setFontSize(9);
      if (fullVoucher.muhtamimMaalApproval.signature) {
        var sigBlob4 = signatureToBlob(fullVoucher.muhtamimMaalApproval.signature, "muhtamim_sig");
        if (sigBlob4) {
          var sigImg4 = c3b.appendImage(sigBlob4);
          sigImg4.setWidth(100).setHeight(40);
        }
      }
    } else {
      c3b.appendParagraph("Awaiting Payment").setFontSize(9);
    }

    body.appendParagraph("");

    // Payment info section
    if (fullVoucher && fullVoucher.payment && fullVoucher.payment.paid) {
      body.appendParagraph("PAYMENT INFORMATION").setBold(true).setFontSize(11).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      body.appendParagraph("");

      const payTable = body.appendTable();
      payTable.setBorderWidth(1);

      var payRow1 = payTable.appendTableRow();
      var payCell1a = payRow1.appendTableCell();
      payCell1a.appendParagraph("Payment Method:").setBold(true).setFontSize(9);
      payCell1a.setWidth(200);
      var payCell1b = payRow1.appendTableCell();
      payCell1b.appendParagraph(fullVoucher.payment.paymentMethod || "Not specified").setFontSize(9);

      var payRow2 = payTable.appendTableRow();
      var payCell2a = payRow2.appendTableCell();
      payCell2a.appendParagraph("Payment Date:").setBold(true).setFontSize(9);
      var payCell2b = payRow2.appendTableCell();
      payCell2b.appendParagraph(fmtDate(fullVoucher.payment.paymentDate) || "Not specified").setFontSize(9);

      var accountInfo = String(fullVoucher.payment.accountNumber || "");
      if (accountInfo && accountInfo.trim()) {
        var payRow3 = payTable.appendTableRow();
        var payCell3a = payRow3.appendTableCell();
        payCell3a.appendParagraph("Account Number:").setBold(true).setFontSize(9);
        var payCell3b = payRow3.appendTableCell();
        payCell3b.appendParagraph(accountInfo).setFontSize(9);
      }

      body.appendParagraph("");
    }

    // Attachment section (page 2)
    var attachmentUrl = fullVoucher ? fullVoucher.attachmentUrl : (voucher.attachmentUrl || "");
    if (attachmentUrl && attachmentUrl.trim() !== '') {
      body.appendPageBreak();

      body.appendParagraph("ATTACHMENT / RECEIPT").setBold(true).setFontSize(14).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      body.appendParagraph("Reference: " + voucher.id).setFontSize(10).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setForegroundColor('#666666');
      body.appendParagraph("");

      var attachment = getAttachmentFromUrl(attachmentUrl);
      if (attachment) {
        if (isImageMimeType(attachment.mimeType)) {
          try {
            var attachImg = body.appendImage(attachment.blob);
            var imgWidth = attachImg.getWidth();
            var imgHeight = attachImg.getHeight();
            var maxWidth = 500;
            var maxHeight = 650;

            var widthRatio = maxWidth / imgWidth;
            var heightRatio = maxHeight / imgHeight;
            var ratio = Math.min(widthRatio, heightRatio, 1);

            if (ratio < 1) {
              attachImg.setWidth(imgWidth * ratio);
              attachImg.setHeight(imgHeight * ratio);
            }

            body.appendParagraph("");
            body.appendParagraph("File: " + attachment.name).setFontSize(9).setForegroundColor('#666666').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
          } catch (imgError) {
            Logger.log("Error embedding image: " + imgError);
            body.appendParagraph("Attachment: " + attachment.name).setFontSize(10);
            body.appendParagraph("View attachment: " + attachmentUrl).setFontSize(9).setForegroundColor('#0066cc');
          }
        } else {
          body.appendParagraph("Attachment: " + attachment.name).setFontSize(10);
          body.appendParagraph("(File type: " + attachment.mimeType + ")").setFontSize(9).setForegroundColor('#666666');
          body.appendParagraph("View attachment: " + attachmentUrl).setFontSize(9).setForegroundColor('#0066cc');
        }
      } else {
        body.appendParagraph("Attachment link: " + attachmentUrl).setFontSize(9).setForegroundColor('#0066cc');
      }
      body.appendParagraph("");
    }

    body.appendParagraph("Generated: " + new Date().toLocaleString('sv-SE') + " | MKA Voucher System").setFontSize(8).setForegroundColor('#888888').setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    doc.saveAndClose();

    // Get/create voucher folder
    const voucherFolder = getVoucherFolder();

    // Delete existing PDFs with same name
    const pdfName = "Voucher_" + voucher.id + ".pdf";
    const existingFiles = voucherFolder.getFilesByName(pdfName);
    while (existingFiles.hasNext()) {
      existingFiles.next().setTrashed(true);
      Logger.log("Deleted existing PDF: " + pdfName);
    }

    // Convert to PDF
    const docFile = DriveApp.getFileById(doc.getId());
    const pdfBlob = docFile.getAs('application/pdf');
    pdfBlob.setName(pdfName);

    const pdfFile = voucherFolder.createFile(pdfBlob);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // Save PDF URL
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName("Vouchers");
    if (fullVoucher) {
      sheet.getRange(fullVoucher.row, 21).setValue(pdfFile.getUrl());
    }

    // Delete temporary doc
    docFile.setTrashed(true);

    Logger.log("PDF created in folder: " + voucherFolder.getName() + " - URL: " + pdfFile.getUrl());
    return pdfFile.getUrl();
  } catch (error) {
    Logger.log("Error generating PDF: " + error);
    return "";
  }
}


// ============ ADMIN PORTAL BACKEND FUNCTIONS ============

function getAdminData() {
  try {
    var ss = getOrCreateSpreadsheet();
    var sheet = ss.getSheetByName("Vouchers");
    var data = sheet.getDataRange().getValues();

    var vouchers = [];
    var stats = {
      total: 0,
      pending: 0,
      pending_qaid: 0,
      pending_nazim: 0,
      pending_sadar: 0,
      approved: 0,
      paid: 0,
      rejected: 0
    };
    var claimerEmailMap = {};

    for (var i = 1; i < data.length; i++) {
      var status = data[i][10] || "unknown";
      stats.total++;

      if (status.startsWith("pending")) stats.pending++;
      if (status === "pending_qaid") stats.pending_qaid++;
      if (status === "pending_nazim") stats.pending_nazim++;
      if (status === "pending_sadar") stats.pending_sadar++;
      if (status === "approved") stats.approved++;
      if (status === "paid") stats.paid++;
      if (status === "rejected") stats.rejected++;

      var email = data[i][3] || "";
      if (email) {
        claimerEmailMap[email] = (claimerEmailMap[email] || 0) + 1;
      }

      vouchers.push({
        id: data[i][0],
        claimedBy: data[i][2] || "",
        email: email,
        description: data[i][4] || "",
        amount: data[i][5] || 0,
        category: data[i][7] || "",
        majlis: data[i][8] || "",
        status: status
      });
    }

    var claimerEmails = Object.keys(claimerEmailMap).map(function(email) {
      return { email: email, count: claimerEmailMap[email] };
    });

    return {
      vouchers: vouchers,
      stats: stats,
      approvers: getAllApproverEmails(),
      allowedEmails: getAllowedEmails(),
      authorizedClaimers: getAuthorizedClaimers(),
      claimerEmails: claimerEmails,
      settings: {
        spreadsheetName: CONFIG.SPREADSHEET_NAME,
        folderName: CONFIG.VOUCHER_FOLDER_NAME,
        webappUrl: CONFIG.WEBAPP_URL
      }
    };
  } catch (error) {
    Logger.log("Error in getAdminData: " + error);
    return { vouchers: [], stats: {}, approvers: {}, authorizedClaimers: [], claimerEmails: [], settings: {} };
  }
}

function updateVoucherFromAdmin(data) {
  try {
    var voucher = getVoucherById(data.id);
    if (!voucher) {
      return { success: false, message: "Voucher not found" };
    }

    var ss = getOrCreateSpreadsheet();
    var sheet = ss.getSheetByName("Vouchers");

    sheet.getRange(voucher.row, 3).setValue(data.claimedBy);
    sheet.getRange(voucher.row, 4).setValue(data.email);
    sheet.getRange(voucher.row, 5).setValue(data.description);
    sheet.getRange(voucher.row, 6).setValue(data.amount);

    SpreadsheetApp.flush();
    return { success: true };
  } catch (error) {
    Logger.log("Error in updateVoucherFromAdmin: " + error);
    return { success: false, message: error.toString() };
  }
}

function resendApprovalEmailFromAdmin(voucherId) {
  try {
    var voucher = getVoucherById(voucherId);
    if (!voucher) {
      return { success: false, message: "Voucher not found" };
    }

    var approverKey = null;
    if (voucher.status === "pending_qaid") approverKey = "qaid";
    else if (voucher.status === "pending_nazim") approverKey = "nazimMaal";
    else if (voucher.status === "pending_sadar") approverKey = "sadarMka";
    else if (voucher.status === "approved") {
      sendMuhtamimMaalPaymentRequest(voucher, voucher.pdfUrl || "");
      return { success: true, message: "Payment request sent to Muhtamim Maal" };
    }

    if (!approverKey) {
      return { success: false, message: "Voucher status is '" + voucher.status + "' - cannot resend" };
    }

    sendApprovalRequest(voucher, approverKey);
    var approver = getApproverInfo(approverKey);
    return { success: true, message: "Email sent to " + approver.title };
  } catch (error) {
    Logger.log("Error in resendApprovalEmailFromAdmin: " + error);
    return { success: false, message: error.toString() };
  }
}

function changeVoucherStatusFromAdmin(voucherId, newStatus) {
  try {
    var voucher = getVoucherById(voucherId);
    if (!voucher) {
      return { success: false, message: "Voucher not found" };
    }

    var validStatuses = ["pending_qaid", "pending_nazim", "pending_sadar", "approved", "paid", "rejected"];
    if (!validStatuses.includes(newStatus)) {
      return { success: false, message: "Invalid status" };
    }

    var ss = getOrCreateSpreadsheet();
    var sheet = ss.getSheetByName("Vouchers");
    sheet.getRange(voucher.row, 11).setValue(newStatus);
    SpreadsheetApp.flush();

    return { success: true };
  } catch (error) {
    Logger.log("Error in changeVoucherStatusFromAdmin: " + error);
    return { success: false, message: error.toString() };
  }
}

function generatePdfFromAdmin(voucherId) {
  try {
    var voucher = getVoucherById(voucherId);
    if (!voucher) {
      return { success: false, message: "Voucher not found" };
    }

    var pdfUrl = generateFinalVoucher(voucher);
    if (pdfUrl) {
      return { success: true, url: pdfUrl };
    } else {
      return { success: false, message: "Failed to generate PDF" };
    }
  } catch (error) {
    Logger.log("Error in generatePdfFromAdmin: " + error);
    return { success: false, message: error.toString() };
  }
}

function backupFromAdmin() {
  try {
    var url = backupSpreadsheet();
    return { success: true, url: url };
  } catch (error) {
    Logger.log("Error in backupFromAdmin: " + error);
    return { success: false, message: error.toString() };
  }
}

function resetFromAdmin() {
  try {
    var ss = getOrCreateSpreadsheet();
    var sheet = ss.getSheetByName("Vouchers");
    var lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
      return { success: true, message: "Database already empty" };
    }

    var dataRows = lastRow - 1;
    sheet.deleteRows(2, dataRows);
    SpreadsheetApp.flush();

    return { success: true, message: "Deleted " + dataRows + " rows" };
  } catch (error) {
    Logger.log("Error in resetFromAdmin: " + error);
    return { success: false, message: error.toString() };
  }
}

function sendTestEmailFromAdmin(email) {
  try {
    var subject = "MKA Voucher System - Test Email";
    var htmlBody = '<div style="font-family:Arial;padding:20px;"><h2 style="color:#1a5f2a;">Test Email<\/h2><p>This is a test email from the MKA Voucher System.<\/p><p>If you received this, email delivery is working correctly.<\/p><p style="color:#666;font-size:12px;">Sent: ' + new Date().toISOString() + '<\/p><\/div>';

    GmailApp.sendEmail(email, subject, "Test email from MKA Voucher System", { htmlBody: htmlBody });
    return { success: true };
  } catch (error) {
    Logger.log("Error in sendTestEmailFromAdmin: " + error);
    return { success: false, message: error.toString() };
  }
}

function createTestVoucherFromAdmin(data) {
  try {
    var voucherData = {
      id: generateVoucherId(),
      timestamp: new Date().toISOString(),
      claimedBy: data.name || "Test Khadim",
      claimerEmail: data.email || "test@example.com",
      descriptionOfProgram: data.description || "Test Voucher",
      amountInNumbers: data.amount || "1000",
      amountInLetters: data.amountLetters || data.amount + " SEK",
      category: data.category || "Other",
      majlis: data.majlis || "Other",
      attachmentUrl: data.attachmentUrl || "",
      status: "pending_qaid"
    };

    saveVoucher(voucherData);
    sendApprovalRequest(voucherData, "qaid", "");
    sendClaimerConfirmation(voucherData, "");

    return { success: true, id: voucherData.id };
  } catch (error) {
    Logger.log("Error in createTestVoucherFromAdmin: " + error);
    return { success: false, message: error.toString() };
  }
}

function fixStatusesFromAdmin() {
  try {
    var ss = getOrCreateSpreadsheet();
    var sheet = ss.getSheetByName("Vouchers");
    var data = sheet.getDataRange().getValues();
    var validStatuses = ["pending_qaid", "pending_nazim", "pending_sadar", "approved", "paid", "rejected"];
    var fixed = 0;

    for (var i = 1; i < data.length; i++) {
      var status = data[i][10];
      if (!validStatuses.includes(status)) {
        sheet.getRange(i + 1, 11).setValue("pending_qaid");
        fixed++;
      }
    }

    if (fixed > 0) {
      SpreadsheetApp.flush();
    }

    return { success: true, message: "Fixed " + fixed + " vouchers" };
  } catch (error) {
    Logger.log("Error in fixStatusesFromAdmin: " + error);
    return { success: false, message: error.toString() };
  }
}

function updateHeadersFromAdmin() {
  try {
    var ss = getOrCreateSpreadsheet();
    var sheet = ss.getSheetByName("Vouchers");

    var headers = [
      "ID", "Timestamp", "Claimed By", "Email", "Description",
      "Amount (Numbers)", "Amount (Letters)", "Category", "Majlis",
      "Attachment URL", "Status",
      "Qaid Approved", "Qaid Signature", "Qaid Date",
      "Nazim Approved", "Nazim Signature", "Nazim Date",
      "Sadar Approved", "Sadar Signature", "Sadar Date",
      "PDF URL",
      "Muhtamim Maal Reviewed", "Muhtamim Maal Signature", "Muhtamim Maal Date",
      "Paid", "Paid By", "Payment Date", "Payment Method", "Paid To",
      "Bank Name", "Account Number", "Clearing Number"
    ];

    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");

    SpreadsheetApp.flush();
    return { success: true };
  } catch (error) {
    Logger.log("Error in updateHeadersFromAdmin: " + error);
    return { success: false, message: error.toString() };
  }
}

function updateClaimerEmailsFromAdmin(oldEmail, newEmail) {
  try {
    var ss = getOrCreateSpreadsheet();
    var sheet = ss.getSheetByName("Vouchers");
    var data = sheet.getDataRange().getValues();
    var updated = 0;

    for (var i = 1; i < data.length; i++) {
      if (data[i][3] === oldEmail) {
        sheet.getRange(i + 1, 4).setValue(newEmail);
        updated++;
      }
    }

    if (updated > 0) {
      SpreadsheetApp.flush();
    }

    return { success: true, message: "Updated " + updated + " vouchers" };
  } catch (error) {
    Logger.log("Error in updateClaimerEmailsFromAdmin: " + error);
    return { success: false, message: error.toString() };
  }
}

function updateApproverEmailFromAdmin(role, email) {
  try {
    if (!role || !email) {
      return { success: false, message: "Role and email are required" };
    }

    var validRoles = ['qaidMajlis', 'nazimMaal', 'sadarMka', 'muhtamimMaal'];
    if (!validRoles.includes(role)) {
      return { success: false, message: "Invalid role: " + role };
    }

    var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(email)) {
      return { success: false, message: "Invalid email format" };
    }

    var result = setApproverEmail(role, email);
    if (result.success) {
      Logger.log("Admin updated " + role + " email to: " + email);
      return { success: true, message: role + " email updated to " + email };
    }
    return result;
  } catch (error) {
    Logger.log("Error in updateApproverEmailFromAdmin: " + error);
    return { success: false, message: error.toString() };
  }
}

function addAllowedEmailFromAdmin(email) {
  try {
    if (!email) {
      return { success: false, message: "Email is required" };
    }

    var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(email)) {
      return { success: false, message: "Invalid email format" };
    }

    var result = addAllowedEmailToProps(email);
    if (result.success) {
      Logger.log("Admin added allowed email: " + email);
      return { success: true, message: email + " added to allowed list" };
    }
    return result;
  } catch (error) {
    Logger.log("Error in addAllowedEmailFromAdmin: " + error);
    return { success: false, message: error.toString() };
  }
}

function removeAllowedEmailFromAdmin(email) {
  try {
    var result = removeAllowedEmailFromProps(email);
    if (result.success) {
      Logger.log("Admin removed allowed email: " + email);
      return { success: true, message: email + " removed from allowed list" };
    }
    return result;
  } catch (error) {
    Logger.log("Error in removeAllowedEmailFromAdmin: " + error);
    return { success: false, message: error.toString() };
  }
}

function addAuthorizedClaimerFromAdmin(email) {
  try {
    if (!email) {
      return { success: false, message: "Email is required" };
    }

    var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(email)) {
      return { success: false, message: "Invalid email format" };
    }

    var claimers = getAuthorizedClaimers();
    var normalizedEmail = email.trim().toLowerCase();

    if (claimers.some(function(e) { return e.toLowerCase() === normalizedEmail; })) {
      return { success: false, message: email + " is already authorized" };
    }

    claimers.push(normalizedEmail);
    var props = PropertiesService.getScriptProperties();
    props.setProperty('authorized_claimers', claimers.join(','));

    Logger.log("Admin added authorized claimer: " + email);
    return { success: true, message: email + " added to authorized claimers" };
  } catch (error) {
    Logger.log("Error in addAuthorizedClaimerFromAdmin: " + error);
    return { success: false, message: error.toString() };
  }
}

function removeAuthorizedClaimerFromAdmin(email) {
  try {
    var claimers = getAuthorizedClaimers();
    var normalizedEmail = email.trim().toLowerCase();

    claimers = claimers.filter(function(e) { return e.toLowerCase() !== normalizedEmail; });

    var props = PropertiesService.getScriptProperties();
    props.setProperty('authorized_claimers', claimers.join(','));

    Logger.log("Admin removed authorized claimer: " + email);
    return { success: true, message: email + " removed from authorized claimers" };
  } catch (error) {
    Logger.log("Error in removeAuthorizedClaimerFromAdmin: " + error);
    return { success: false, message: error.toString() };
  }
}

function sendTestSignatureEmailFromAdmin(role, customEmail) {
  return sendTestSignatureEmail(role, customEmail);
}


// ============ TEST & DEBUG FUNCTIONS ============

function checkScriptProperties() {
  var props = PropertiesService.getScriptProperties();
  var all = props.getProperties();
  Logger.log("=== ALL SCRIPT PROPERTIES ===");
  for (var key in all) {
    Logger.log(key + ": " + all[key]);
  }
  Logger.log("=== CURRENT APPROVERS ===");
  Logger.log("qaidMajlis: " + getApproverEmail('qaidMajlis'));
  Logger.log("nazimMaal: " + getApproverEmail('nazimMaal'));
  Logger.log("sadarMka: " + getApproverEmail('sadarMka'));
  Logger.log("muhtamimMaal: " + getApproverEmail('muhtamimMaal'));
}

function manuallySetApproverEmail() {
  var role = "qaidMajlis";
  var email = "newemail@example.com";

  var result = setApproverEmail(role, email);
  Logger.log("Result: " + JSON.stringify(result));
  Logger.log("Verification - " + role + " is now: " + getApproverEmail(role));
}

function debugApprovalUrls() {
  Logger.log("============================================");
  Logger.log("DEBUG: URL CHECK FOR ALL APPROVERS");
  Logger.log("============================================");
  Logger.log("CONFIG.WEBAPP_URL: " + CONFIG.WEBAPP_URL);
  Logger.log("");

  var testId = "TEST-VOUCHER-123";
  var approvers = ["qaid", "nazimMaal", "sadarMka"];
  var roles = {
    "qaid": "Qaid Majlis",
    "nazimMaal": "Nazim Maal",
    "sadarMka": "Sadar MKA"
  };

  approvers.forEach(function(key) {
    var approveUrl = CONFIG.WEBAPP_URL + "?action=approve&id=" + testId + "&approver=" + key;
    var rejectUrl = CONFIG.WEBAPP_URL + "?action=reject&id=" + testId + "&approver=" + key;

    Logger.log("=== " + roles[key] + " (" + key + ") ===");
    Logger.log("Approve URL: " + approveUrl);
    Logger.log("Reject URL: " + rejectUrl);
  });

  var paymentUrl = CONFIG.WEBAPP_URL + "?action=payment&id=" + testId;
  Logger.log("=== Muhtamim Maal (Payment) ===");
  Logger.log("Payment URL: " + paymentUrl);
}

function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(getOrCreateSpreadsheet()).onFormSubmit().create();
  Logger.log('Trigger created!');
}

function listTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    Logger.log("Trigger: " + t.getHandlerFunction() + " | Type: " + t.getEventType());
  });
  Logger.log("Total triggers: " + triggers.length);
}

function deleteTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });
  Logger.log('All triggers deleted!');
}

function testFullWorkflow() {
  Logger.log("=== Starting Test ===");

  var testVoucher = {
    id: generateVoucherId(),
    timestamp: new Date().toISOString(),
    claimedBy: "Test Khadim",
    claimerEmail: "fakturaamlcars@gmail.com",
    descriptionOfProgram: "Ijtema 2024 Transport",
    amountInNumbers: "2500",
    amountInLetters: "Two Thousand Five Hundred",
    category: "Travelling",
    majlis: "Lulea",
    attachmentUrl: "",
    status: "pending_qaid"
  };

  Logger.log("1. Saving voucher: " + testVoucher.id);
  saveVoucher(testVoucher);

  Logger.log("2. Sending to Qaid Majlis: " + CONFIG.APPROVERS.qaidMajlis);
  sendApprovalRequest(testVoucher, "qaid");

  Logger.log("3. Sending confirmation to claimer: " + testVoucher.claimerEmail);
  sendClaimerConfirmation(testVoucher);

  Logger.log("=== Test Complete ===");
  Logger.log("Voucher ID: " + testVoucher.id);
}

function testCreateSpreadsheet() {
  var ss = getOrCreateSpreadsheet();
  Logger.log("Spreadsheet: " + ss.getUrl());
}

function debugVoucher() {
  var voucherId = "MKA-20251205-15184770";

  Logger.log("========== DEBUG VOUCHER ==========");
  var voucher = getVoucherById(voucherId);

  if (!voucher) {
    Logger.log("ERROR: Voucher not found: " + voucherId);
    return;
  }

  Logger.log("ID: " + voucher.id);
  Logger.log("Row: " + voucher.row);
  Logger.log("Status: " + voucher.status);
  Logger.log("Claimed By: " + voucher.claimedBy);
  Logger.log("Claimer Email: " + voucher.claimerEmail);
  Logger.log("Amount: " + voucher.amountInNumbers + " SEK");
  Logger.log("");
  Logger.log("--- Approvals ---");
  Logger.log("Qaid: " + (voucher.qaidApproval.approved ? "APPROVED" : "Pending"));
  Logger.log("Nazim: " + (voucher.nazimApproval.approved ? "APPROVED" : "Pending"));
  Logger.log("Sadar: " + (voucher.sadarApproval.approved ? "APPROVED" : "Pending"));
  Logger.log("");
  Logger.log("PDF URL: " + (voucher.pdfUrl || "Not generated"));
}

function listAllVouchers() {
  Logger.log("========== ALL VOUCHERS ==========");
  var ss = getOrCreateSpreadsheet();
  var sheet = ss.getSheetByName("Vouchers");
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var id = data[i][0];
    var status = data[i][10];
    var claimedBy = data[i][2];
    var amount = data[i][5];
    Logger.log((i) + ". " + id + " | " + status + " | " + claimedBy + " | " + amount + " SEK");
  }
  Logger.log("Total: " + (data.length - 1) + " vouchers");
}

function manualSetStatus() {
  var voucherId = "MKA-20251205-15184770";
  var newStatus = "pending_sadar";

  Logger.log("Setting status for " + voucherId + " to: " + newStatus);

  var voucher = getVoucherById(voucherId);
  if (!voucher) {
    Logger.log("ERROR: Voucher not found");
    return;
  }

  var validStatuses = ["pending_qaid", "pending_nazim", "pending_sadar", "approved", "rejected"];
  if (!validStatuses.includes(newStatus)) {
    Logger.log("ERROR: Invalid status");
    return;
  }

  var ss = getOrCreateSpreadsheet();
  var sheet = ss.getSheetByName("Vouchers");
  sheet.getRange(voucher.row, 11).setValue(newStatus);
  SpreadsheetApp.flush();

  Logger.log("SUCCESS: Status updated to " + newStatus);
}

function resendApprovalEmail() {
  var voucherId = "MKA-20251205-15184770";

  Logger.log("Resending approval email for: " + voucherId);

  var voucher = getVoucherById(voucherId);
  if (!voucher) {
    Logger.log("ERROR: Voucher not found");
    return;
  }

  var approverKey = null;
  if (voucher.status === "pending_qaid") approverKey = "qaid";
  else if (voucher.status === "pending_nazim") approverKey = "nazimMaal";
  else if (voucher.status === "pending_sadar") approverKey = "sadarMka";

  if (!approverKey) {
    Logger.log("ERROR: Voucher status is '" + voucher.status + "' - not pending approval");
    return;
  }

  var approver = getApproverInfo(approverKey);
  Logger.log("Sending to: " + approver.title + " (" + approver.email + ")");

  sendApprovalRequest(voucher, approverKey);
  Logger.log("SUCCESS: Email sent!");
}

function manualGeneratePdf() {
  var voucherId = "MKA-20251205-15184770";

  Logger.log("Generating PDF for: " + voucherId);

  var voucher = getVoucherById(voucherId);
  if (!voucher) {
    Logger.log("ERROR: Voucher not found");
    return;
  }

  var pdfUrl = generateFinalVoucher(voucher);
  if (pdfUrl) {
    Logger.log("SUCCESS: PDF created at: " + pdfUrl);
  } else {
    Logger.log("ERROR: Failed to generate PDF");
  }
}

function fixInvalidStatuses() {
  Logger.log("Checking for invalid statuses...");

  var ss = getOrCreateSpreadsheet();
  var sheet = ss.getSheetByName("Vouchers");
  var data = sheet.getDataRange().getValues();
  var validStatuses = ["pending_qaid", "pending_nazim", "pending_sadar", "approved", "rejected"];
  var fixed = 0;

  for (var i = 1; i < data.length; i++) {
    var status = data[i][10];
    if (!validStatuses.includes(status)) {
      var id = data[i][0];
      Logger.log("Fixing row " + (i+1) + " (" + id + "): '" + status + "' -> 'pending_qaid'");
      sheet.getRange(i + 1, 11).setValue("pending_qaid");
      fixed++;
    }
  }

  if (fixed > 0) {
    SpreadsheetApp.flush();
    Logger.log("Fixed " + fixed + " vouchers with invalid status");
  } else {
    Logger.log("All statuses are valid!");
  }
}

function manualSendFinalNotification() {
  var voucherId = "MKA-20251205-15184770";

  Logger.log("Sending final notification for: " + voucherId);

  var voucher = getVoucherById(voucherId);
  if (!voucher) {
    Logger.log("ERROR: Voucher not found");
    return;
  }

  if (voucher.status !== "approved") {
    Logger.log("WARNING: Voucher status is '" + voucher.status + "' not 'approved'");
  }

  sendFinalNotification(voucher, voucher.pdfUrl || "");
  Logger.log("SUCCESS: Final notification sent!");
}

function statusOverview() {
  Logger.log("========== STATUS OVERVIEW ==========");
  var ss = getOrCreateSpreadsheet();
  var sheet = ss.getSheetByName("Vouchers");
  var data = sheet.getDataRange().getValues();

  var counts = {};
  for (var i = 1; i < data.length; i++) {
    var status = data[i][10] || "unknown";
    counts[status] = (counts[status] || 0) + 1;
  }

  for (var status in counts) {
    Logger.log(status + ": " + counts[status]);
  }
  Logger.log("Total: " + (data.length - 1) + " vouchers");
}


// ============ RESET & BACKUP FUNCTIONS ============

function resetSpreadsheet() {
  var confirmReset = true;

  if (!confirmReset) {
    Logger.log("To actually reset, change confirmReset to true in the code");
    return;
  }

  Logger.log("========== RESETTING SPREADSHEET ==========");

  var ss = getOrCreateSpreadsheet();
  var sheet = ss.getSheetByName("Vouchers");
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    Logger.log("Spreadsheet is already empty");
    return;
  }

  var dataRows = lastRow - 1;
  Logger.log("Deleting " + dataRows + " voucher rows...");

  sheet.deleteRows(2, dataRows);
  SpreadsheetApp.flush();

  Logger.log("SUCCESS: Spreadsheet has been reset!");
}

function backupSpreadsheet() {
  Logger.log("========== BACKING UP SPREADSHEET ==========");

  var ss = getOrCreateSpreadsheet();
  var originalName = ss.getName();

  var now = new Date();
  var timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm-ss");
  var backupName = originalName + "_BACKUP_" + timestamp;

  var file = DriveApp.getFileById(ss.getId());
  var backupFile = file.makeCopy(backupName);

  Logger.log("SUCCESS: Backup created!");
  Logger.log("- Original: " + originalName);
  Logger.log("- Backup: " + backupName);
  Logger.log("- Backup URL: " + backupFile.getUrl());

  return backupFile.getUrl();
}

function backupSpreadsheetToFolder() {
  Logger.log("========== BACKING UP TO FOLDER ==========");

  var ss = getOrCreateSpreadsheet();
  var originalName = ss.getName();

  var now = new Date();
  var timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm-ss");
  var backupName = originalName + "_BACKUP_" + timestamp;

  var folderName = "Voucher Backups";
  var folders = DriveApp.getFoldersByName(folderName);
  var backupFolder;

  if (folders.hasNext()) {
    backupFolder = folders.next();
  } else {
    backupFolder = DriveApp.createFolder(folderName);
  }

  var file = DriveApp.getFileById(ss.getId());
  var backupFile = file.makeCopy(backupName, backupFolder);

  Logger.log("SUCCESS: Backup created!");
  Logger.log("- Backup: " + backupName);
  Logger.log("- Folder: " + folderName);
  Logger.log("- URL: " + backupFile.getUrl());

  return backupFile.getUrl();
}

function listBackups() {
  Logger.log("========== VOUCHER BACKUPS ==========");

  var folderName = "Voucher Backups";
  var folders = DriveApp.getFoldersByName(folderName);

  if (!folders.hasNext()) {
    Logger.log("No backup folder found.");
    return;
  }

  var backupFolder = folders.next();
  var files = backupFolder.getFiles();
  var count = 0;

  while (files.hasNext()) {
    var file = files.next();
    count++;
    Logger.log(count + ". " + file.getName());
    Logger.log("   Created: " + file.getDateCreated());
    Logger.log("   URL: " + file.getUrl());
  }

  if (count === 0) {
    Logger.log("No backups found in folder.");
  } else {
    Logger.log("Total backups: " + count);
  }
}

function listAllEmails() {
  Logger.log("========== ALL CONFIGURED EMAILS ==========");
  Logger.log("");
  Logger.log("--- APPROVER EMAILS ---");
  Logger.log("Qaid Majlis:    " + getApproverEmail('qaidMajlis'));
  Logger.log("Nazim Maal:     " + getApproverEmail('nazimMaal'));
  Logger.log("Sadar MKA:      " + getApproverEmail('sadarMka'));
  Logger.log("Muhtamim Maal:  " + getApproverEmail('muhtamimMaal'));
  Logger.log("");
  Logger.log("--- ACCESS CONTROL ---");
  var allowed = getAllowedEmails();
  allowed.forEach(function(email, i) {
    Logger.log("  " + (i + 1) + ". " + email);
  });
}

function listActualApproverEmails() {
  Logger.log("=== ACTUAL APPROVER EMAILS (from Script Properties) ===");
  Logger.log("Qaid Majlis: " + getApproverEmail('qaidMajlis'));
  Logger.log("Nazim Maal: " + getApproverEmail('nazimMaal'));
  Logger.log("Sadar MKA: " + getApproverEmail('sadarMka'));
  Logger.log("Muhtamim Maal: " + getApproverEmail('muhtamimMaal'));

  Logger.log("");
  Logger.log("=== DEFAULT EMAILS (from CONFIG - fallback only) ===");
  Logger.log("Qaid Majlis: " + CONFIG.APPROVERS.qaidMajlis);
  Logger.log("Nazim Maal: " + CONFIG.APPROVERS.nazimMaal);
  Logger.log("Sadar MKA: " + CONFIG.APPROVERS.sadarMka);
  Logger.log("Muhtamim Maal: " + CONFIG.APPROVERS.muhtamimMaal);
}

function fixAllAttachmentPermissions() {
  Logger.log("=== FIXING ALL ATTACHMENT PERMISSIONS ===");
  var ss = getOrCreateSpreadsheet();
  var sheet = ss.getSheetByName("Vouchers");
  var data = sheet.getDataRange().getValues();

  var fixed = 0;
  var errors = 0;
  for (var i = 1; i < data.length; i++) {
    var attachmentUrl = data[i][9];
    if (attachmentUrl && attachmentUrl.trim() !== '') {
      if (shareAttachmentWithAnyone(attachmentUrl)) {
        fixed++;
        Logger.log("Fixed: Row " + (i+1) + " - " + data[i][0]);
      } else {
        errors++;
        Logger.log("ERROR: Row " + (i+1) + " - " + data[i][0]);
      }
    }
  }

  Logger.log("=== COMPLETE ===");
  Logger.log("Fixed: " + fixed + " attachments");
  Logger.log("Errors: " + errors);

  return { fixed: fixed, errors: errors };
}


// ============ ADMIN PORTAL PAGE ============

/**
 * Get the admin portal HTML page
 */
function getAdminPortalPage() {
  var html = [];
  html.push('<!DOCTYPE html>');
  html.push('<html>');
  html.push('<head>');
  html.push('<meta charset="UTF-8">');
  html.push('<meta name="viewport" content="width=device-width, initial-scale=1">');
  html.push('<title>Admin Portal - MKA Voucher System</title>');
  html.push('<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">');
  html.push('<style>');
  html.push(':root { --primary: #00c853; --primary-dark: #009624; --dark: #1a1a1a; --darker: #0d0d0d; --light: #f5f5f5; --gray: #888; }');
  html.push('* { box-sizing: border-box; margin: 0; padding: 0; }');
  html.push('body { font-family: Inter, -apple-system, BlinkMacSystemFont, Segoe UI, Roboto, sans-serif; background: var(--darker); min-height: 100vh; color: var(--light); }');
  html.push('.header { background: var(--dark); border-bottom: 1px solid rgba(0,200,83,0.3); padding: 15px 25px; position: sticky; top: 0; z-index: 100; }');
  html.push('.header-content { max-width: 1400px; margin: 0 auto; display: flex; justify-content: space-between; align-items: center; }');
  html.push('.header h1 { font-size: 24px; font-weight: 700; color: var(--primary); }');
  html.push('.header-right { display: flex; gap: 12px; align-items: center; }');
  html.push('.header-btn { background: var(--primary); color: var(--dark); border: none; padding: 10px 20px; border-radius: 8px; cursor: pointer; font-size: 13px; font-weight: 600; text-decoration: none; transition: all 0.3s ease; }');
  html.push('.header-btn:hover { background: var(--primary-dark); transform: translateY(-2px); }');
  html.push('.container { max-width: 1400px; margin: 0 auto; padding: 25px; }');
  html.push('.stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 20px; margin-bottom: 30px; }');
  html.push('.stat-card { background: var(--dark); padding: 25px; border-radius: 12px; text-align: center; border: 1px solid rgba(255,255,255,0.1); transition: all 0.3s ease; }');
  html.push('.stat-card:hover { border-color: var(--primary); transform: translateY(-3px); }');
  html.push('.stat-number { font-size: 42px; font-weight: 700; color: var(--primary); }');
  html.push('.stat-label { font-size: 11px; color: var(--gray); margin-top: 8px; text-transform: uppercase; letter-spacing: 1px; font-weight: 600; }');
  html.push('.stat-card.pending .stat-number { color: #ffc107; }');
  html.push('.stat-card.approved .stat-number { color: #17a2b8; }');
  html.push('.stat-card.paid .stat-number { color: var(--primary); }');
  html.push('.stat-card.rejected .stat-number { color: #dc3545; }');
  html.push('.tabs { display: flex; gap: 5px; margin-bottom: 25px; flex-wrap: wrap; background: var(--dark); padding: 8px; border-radius: 10px; border: 1px solid rgba(255,255,255,0.1); }');
  html.push('.tab { padding: 12px 24px; background: transparent; border: none; border-radius: 6px; cursor: pointer; font-size: 14px; font-weight: 600; color: var(--gray); transition: all 0.3s ease; }');
  html.push('.tab.active { background: var(--primary); color: var(--dark); }');
  html.push('.tab:hover:not(.active) { background: rgba(0,200,83,0.1); color: var(--primary); }');
  html.push('.panel { display: none; background: var(--dark); border-radius: 12px; padding: 30px; border: 1px solid rgba(255,255,255,0.1); }');
  html.push('.panel.active { display: block; }');
  html.push('.table-container { overflow-x: auto; border-radius: 8px; }');
  html.push('table { width: 100%; border-collapse: collapse; font-size: 14px; }');
  html.push('th, td { padding: 16px 15px; text-align: left; }');
  html.push('th { background: rgba(0,200,83,0.15); color: var(--primary); font-weight: 600; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid var(--primary); }');
  html.push('tr { border-bottom: 1px solid rgba(255,255,255,0.05); transition: all 0.2s ease; }');
  html.push('tr:hover { background: rgba(0,200,83,0.05); }');
  html.push('.status { padding: 6px 14px; border-radius: 6px; font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; }');
  html.push('.status.pending_qaid { background: rgba(255,193,7,0.2); color: #ffc107; border: 1px solid #ffc107; }');
  html.push('.status.pending_nazim { background: rgba(23,162,184,0.2); color: #17a2b8; border: 1px solid #17a2b8; }');
  html.push('.status.pending_sadar { background: rgba(111,66,193,0.2); color: #6f42c1; border: 1px solid #6f42c1; }');
  html.push('.status.approved { background: rgba(0,200,83,0.2); color: var(--primary); border: 1px solid var(--primary); }');
  html.push('.status.paid { background: rgba(0,200,83,0.3); color: var(--primary); border: 1px solid var(--primary); }');
  html.push('.status.rejected { background: rgba(220,53,69,0.2); color: #dc3545; border: 1px solid #dc3545; }');
  html.push('.action-btn { padding: 8px 14px; border: none; border-radius: 6px; cursor: pointer; font-size: 12px; font-weight: 600; margin: 2px; transition: all 0.3s ease; }');
  html.push('.action-btn.edit { background: rgba(0,200,83,0.2); color: var(--primary); border: 1px solid var(--primary); }');
  html.push('.action-btn.email { background: rgba(23,162,184,0.2); color: #17a2b8; border: 1px solid #17a2b8; }');
  html.push('.action-btn.status { background: rgba(255,193,7,0.2); color: #ffc107; border: 1px solid #ffc107; }');
  html.push('.action-btn:hover { transform: translateY(-2px); filter: brightness(1.2); }');
  html.push('.form-group { margin-bottom: 20px; }');
  html.push('.form-group label { display: block; font-weight: 600; margin-bottom: 8px; font-size: 13px; color: var(--light); }');
  html.push('.form-group input, .form-group select, .form-group textarea { width: 100%; padding: 14px 16px; background: var(--darker); border: 1px solid rgba(255,255,255,0.2); border-radius: 8px; font-size: 14px; color: var(--light); transition: all 0.3s ease; }');
  html.push('.form-group input:focus, .form-group select:focus, .form-group textarea:focus { outline: none; border-color: var(--primary); box-shadow: 0 0 0 3px rgba(0,200,83,0.2); }');
  html.push('.form-group textarea { min-height: 80px; resize: vertical; }');
  html.push('.btn { padding: 14px 28px; border: none; border-radius: 8px; cursor: pointer; font-size: 14px; font-weight: 600; transition: all 0.3s ease; }');
  html.push('.btn-primary { background: var(--primary); color: var(--dark); }');
  html.push('.btn-danger { background: #dc3545; color: white; }');
  html.push('.btn-secondary { background: rgba(255,255,255,0.1); color: var(--light); border: 1px solid rgba(255,255,255,0.2); }');
  html.push('.btn:hover { transform: translateY(-2px); filter: brightness(1.1); }');
  html.push('.btn:disabled { opacity: 0.5; cursor: not-allowed; transform: none; }');
  html.push('.modal { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.8); backdrop-filter: blur(5px); z-index: 1000; align-items: center; justify-content: center; }');
  html.push('.modal.show { display: flex; }');
  html.push('.modal-content { background: var(--dark); padding: 35px; border-radius: 16px; max-width: 500px; width: 90%; max-height: 90vh; overflow-y: auto; border: 1px solid rgba(0,200,83,0.3); }');
  html.push('.modal-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 25px; padding-bottom: 15px; border-bottom: 1px solid rgba(255,255,255,0.1); }');
  html.push('.modal-header h3 { font-size: 20px; font-weight: 700; color: var(--primary); }');
  html.push('.modal-close { background: rgba(255,255,255,0.1); border: none; width: 36px; height: 36px; border-radius: 50%; font-size: 20px; cursor: pointer; color: var(--gray); transition: all 0.3s ease; }');
  html.push('.modal-close:hover { background: var(--primary); color: var(--dark); }');
  html.push('.settings-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 25px; }');
  html.push('.settings-card { background: rgba(255,255,255,0.03); padding: 25px; border-radius: 12px; border: 1px solid rgba(255,255,255,0.1); }');
  html.push('.settings-card h4 { margin-bottom: 20px; color: var(--primary); font-size: 16px; font-weight: 700; }');
  html.push('.settings-item { display: flex; justify-content: space-between; align-items: center; padding: 12px 0; border-bottom: 1px solid rgba(255,255,255,0.05); font-size: 14px; }');
  html.push('.settings-item:last-child { border-bottom: none; }');
  html.push('.settings-label { color: var(--gray); font-weight: 500; }');
  html.push('.settings-value { font-weight: 600; color: var(--light); word-break: break-all; }');
  html.push('.quick-actions { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; }');
  html.push('.quick-action-card { background: rgba(255,255,255,0.03); padding: 30px; border-radius: 12px; text-align: center; cursor: pointer; transition: all 0.3s ease; border: 1px solid rgba(255,255,255,0.1); }');
  html.push('.quick-action-card:hover { border-color: var(--primary); transform: translateY(-5px); background: rgba(0,200,83,0.05); }');
  html.push('.quick-action-icon { font-size: 40px; margin-bottom: 15px; }');
  html.push('.quick-action-title { font-weight: 700; margin-bottom: 8px; font-size: 16px; color: var(--light); }');
  html.push('.quick-action-desc { font-size: 13px; color: var(--gray); }');
  html.push('.toolbar { display: flex; gap: 15px; margin-bottom: 25px; flex-wrap: wrap; }');
  html.push('.toolbar input, .toolbar select { padding: 12px 18px; background: var(--darker); border: 1px solid rgba(255,255,255,0.2); border-radius: 8px; font-size: 14px; color: var(--light); transition: all 0.3s ease; }');
  html.push('.toolbar input { flex: 1; min-width: 250px; }');
  html.push('.toolbar input:focus, .toolbar select:focus { border-color: var(--primary); box-shadow: 0 0 0 3px rgba(0,200,83,0.2); outline: none; }');
  html.push('.toast { position: fixed; bottom: 30px; right: 30px; background: var(--dark); color: var(--light); padding: 18px 30px; border-radius: 8px; display: none; z-index: 2000; font-weight: 600; border: 1px solid rgba(255,255,255,0.1); }');
  html.push('.toast.success { background: var(--primary); color: var(--dark); border-color: var(--primary); }');
  html.push('.toast.error { background: #dc3545; color: white; border-color: #dc3545; }');
  html.push('.toast.show { display: block; animation: slideIn 0.4s ease; }');
  html.push('@keyframes slideIn { from { transform: translateX(120%); opacity: 0; } to { transform: translateX(0); opacity: 1; } }');
  html.push('.loading { text-align: center; padding: 50px; color: var(--gray); }');
  html.push('.spinner { border: 4px solid rgba(0,200,83,0.2); border-top: 4px solid var(--primary); border-radius: 50%; width: 40px; height: 40px; animation: spin 0.8s linear infinite; margin: 0 auto 15px; }');
  html.push('@keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }');
  html.push('.progress-bar { display: flex; align-items: center; gap: 0; background: var(--darker); border-radius: 8px; padding: 15px 20px; margin-top: 8px; }');
  html.push('.progress-step { display: flex; align-items: center; flex: 1; }');
  html.push('.progress-step:last-child { flex: 0; }');
  html.push('.step-circle { width: 32px; height: 32px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 14px; font-weight: 700; flex-shrink: 0; transition: all 0.3s ease; }');
  html.push('.step-circle.completed { background: var(--primary); color: var(--dark); }');
  html.push('.step-circle.active { background: #ffc107; color: var(--dark); animation: glow 1.5s infinite; }');
  html.push('.step-circle.pending { background: rgba(255,255,255,0.1); color: var(--gray); }');
  html.push('.step-circle.rejected { background: #dc3545; color: white; }');
  html.push('.step-line { flex: 1; height: 3px; background: rgba(255,255,255,0.1); margin: 0 8px; }');
  html.push('.step-line.completed { background: var(--primary); }');
  html.push('.step-info { display: flex; flex-direction: column; align-items: center; }');
  html.push('.step-label { font-size: 10px; color: var(--gray); margin-top: 6px; font-weight: 600; white-space: nowrap; }');
  html.push('@keyframes glow { 0%, 100% { box-shadow: 0 0 5px #ffc107, 0 0 10px #ffc107; } 50% { box-shadow: 0 0 15px #ffc107, 0 0 25px #ffc107; } }');
  html.push('@media (max-width: 768px) { .header-content { flex-direction: column; gap: 15px; } .stats-grid { grid-template-columns: repeat(2, 1fr); } .tabs { justify-content: center; } .progress-bar { flex-wrap: wrap; gap: 10px; } }');
  html.push('</style>');
  html.push('</head>');
  html.push('<body>');
  html.push('<div class="header"><div class="header-content"><h1>MKA Admin Portal</h1><div class="header-right"><button class="header-btn" onclick="refreshData()">Refresh</button><a href="?action=admin" class="header-btn">Admin</a><a href="?" class="header-btn">Main Site</a></div></div></div>');
  html.push('<div class="container">');
  html.push('<div class="stats-grid" id="statsGrid"><div class="stat-card"><div class="spinner"></div></div></div>');
  html.push('<div class="tabs"><button class="tab active" onclick="showTab(\'vouchers\')">Vouchers</button><button class="tab" onclick="showTab(\'emails\')">Email Management</button><button class="tab" onclick="showTab(\'actions\')">Quick Actions</button><button class="tab" onclick="showTab(\'settings\')">Settings</button></div>');
  html.push('<div class="panel active" id="vouchers-panel"><div class="toolbar"><input type="text" id="searchInput" placeholder="Search..." onkeyup="filterTable()"><select id="statusFilter" onchange="filterTable()"><option value="">All</option><option value="pending_qaid">Pending Qaid</option><option value="pending_nazim">Pending Nazim</option><option value="pending_sadar">Pending Sadar</option><option value="approved">Approved</option><option value="paid">Paid</option><option value="rejected">Rejected</option></select></div><div class="table-container"><table><thead><tr><th>ID</th><th>Claimer</th><th>Amount</th><th>Progress</th><th>Actions</th></tr></thead><tbody id="vouchersBody"><tr><td colspan="5" class="loading"><div class="spinner"></div>Loading...</td></tr></tbody></table></div></div>');
  html.push('<div class="panel" id="emails-panel"><h3 style="margin-bottom:20px;">Approver Email Configuration</h3><div class="settings-grid"><div class="settings-card"><h4>Current Approvers</h4><div id="approverEmails">Loading...</div></div><div class="settings-card"><h4>Update Approver Email</h4><div class="form-group"><label>Select Role:</label><select id="approverRole"><option value="qaidMajlis">Qaid Majlis</option><option value="nazimMaal">Nazim Maal</option><option value="sadarMka">Sadar MKA</option><option value="muhtamimMaal">Muhtamim Maal</option></select></div><div class="form-group"><label>New Email:</label><input type="email" id="newApproverEmail" placeholder="Enter new email"></div><button class="btn btn-primary" onclick="updateApproverEmail()">Update Email</button></div></div><h3 style="margin:30px 0 20px;">Authorized Claimers</h3><p style="color:var(--gray);margin-bottom:15px;font-size:13px;">Emails allowed to submit voucher claims</p><div class="settings-grid"><div class="settings-card"><h4>Current Authorized Claimers</h4><div id="authorizedClaimers">Loading...</div></div><div class="settings-card"><h4>Add New Claimer</h4><div class="form-group"><label>Email Address:</label><input type="email" id="newClaimerEmail" placeholder="email@example.com"></div><button class="btn btn-primary" onclick="addAuthorizedClaimer()">Add Claimer</button></div></div><h3 style="margin:30px 0 20px;">Claimer History</h3><p style="color:var(--gray);margin-bottom:15px;font-size:13px;">Emails that have submitted vouchers</p><div class="table-container"><table><thead><tr><th>Email</th><th>Vouchers</th><th>Actions</th></tr></thead><tbody id="claimerEmailsBody"><tr><td colspan="3" class="loading">Loading...</td></tr></tbody></table></div></div>');
  html.push('<div class="panel" id="actions-panel"><div class="quick-actions"><div class="quick-action-card" onclick="showBackupModal()"><div class="quick-action-icon">💾</div><div class="quick-action-title">Backup</div><div class="quick-action-desc">Create backup</div></div><div class="quick-action-card" onclick="showResetModal()"><div class="quick-action-icon">🔄</div><div class="quick-action-title">Reset</div><div class="quick-action-desc">Clear data</div></div><div class="quick-action-card" onclick="showTestEmailModal()"><div class="quick-action-icon">📧</div><div class="quick-action-title">Test Email</div><div class="quick-action-desc">Send test</div></div><div class="quick-action-card" onclick="showTestSignatureModal()"><div class="quick-action-icon">✍️</div><div class="quick-action-title">Test Signature</div><div class="quick-action-desc">Test approver</div></div><div class="quick-action-card" onclick="showNewVoucherModal()"><div class="quick-action-icon">➕</div><div class="quick-action-title">New Voucher</div><div class="quick-action-desc">Create test</div></div></div></div>');
  html.push('<div class="panel" id="settings-panel"><div class="settings-grid"><div class="settings-card"><h4>System</h4><div class="settings-item"><span class="settings-label">Spreadsheet:</span><span class="settings-value" id="ssName">Loading...</span></div><div class="settings-item"><span class="settings-label">Folder:</span><span class="settings-value" id="folderName">Loading...</span></div></div><div class="settings-card"><h4>Access Control</h4><div id="allowedEmails">Loading...</div><div class="form-group" style="margin-top:15px;"><label>Add Email:</label><input type="email" id="newAllowedEmail" placeholder="email@example.com"></div><button class="btn btn-primary" onclick="addAllowedEmail()">Add</button></div><div class="settings-card"><h4>Statistics</h4><div id="detailedStats">Loading...</div></div></div></div>');
  html.push('</div>');
  html.push('<div class="modal" id="editModal"><div class="modal-content"><div class="modal-header"><h3>Edit Voucher</h3><button class="modal-close" onclick="closeModal(\'editModal\')">&times;</button></div><div id="editModalBody"></div></div></div>');
  html.push('<div class="modal" id="genericModal"><div class="modal-content"><div class="modal-header"><h3 id="genericModalTitle">Modal</h3><button class="modal-close" onclick="closeModal(\'genericModal\')">&times;</button></div><div id="genericModalBody"></div></div></div>');
  html.push('<div class="toast" id="toast"></div>');
  html.push('<script>');
  html.push('var allVouchers = [];');
  html.push('var T = String.fromCharCode(60);');
  html.push('function tag(n,c,a){var s=T+n;if(a){if(typeof a==="string"){s+=" "+a;}else{for(var k in a)s+=" "+k+"=\\""+a[k]+"\\"";}}return s+">"+(c||"")+T+"/"+n+">";}');
  html.push('document.addEventListener("DOMContentLoaded", function() { loadAllData(); });');
  html.push('function escapeHtml(str) { if (str === null || str === undefined) return ""; return String(str).replace(/&/g,"&amp;").replace(new RegExp(T,"g"),"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"); }');
  html.push('function loadAllData() { document.getElementById("statsGrid").innerHTML = tag("div",tag("div","","class=spinner"),"class=stat-card"); google.script.run.withSuccessHandler(function(data) { allVouchers = data.vouchers || []; renderStats(data.stats || {}); renderVouchers(allVouchers); renderApproverEmails(data.approvers || {}, {}); renderClaimerEmails(data.claimerEmails || []); renderSettings(data.settings || {}); renderDetailedStats(data.stats || {}); renderAllowedEmails(data.allowedEmails || []); renderAuthorizedClaimers(data.authorizedClaimers || []); }).withFailureHandler(function(err) { showToast("Error: " + err, "error"); }).getAdminData(); }');
  html.push('function refreshData() { showToast("Refreshing...", ""); loadAllData(); }');
  html.push('function renderStats(stats) { var h = ""; h += tag("div",tag("div",stats.total||0,"class=stat-number")+tag("div","Total","class=stat-label"),"class=stat-card"); h += tag("div",tag("div",stats.pending||0,"class=stat-number")+tag("div","Pending","class=stat-label"),"class=stat-card pending"); h += tag("div",tag("div",stats.approved||0,"class=stat-number")+tag("div","Approved","class=stat-label"),"class=stat-card approved"); h += tag("div",tag("div",stats.paid||0,"class=stat-number")+tag("div","Paid","class=stat-label"),"class=stat-card paid"); h += tag("div",tag("div",stats.rejected||0,"class=stat-number")+tag("div","Rejected","class=stat-label"),"class=stat-card rejected"); document.getElementById("statsGrid").innerHTML = h; }');
  html.push('function btn(cls,txt,fn,arg){var Q=String.fromCharCode(39);return tag("button",txt,"class=\\"action-btn "+cls+"\\" onclick=\\""+fn+"("+Q+arg+Q+")\\"");}');
  html.push('function renderVouchers(vouchers) { var h = ""; if (vouchers.length === 0) { h = tag("tr",tag("td","No vouchers","colspan=5 style=text-align:center")); } else { vouchers.forEach(function(v) { if (!v) return; var vid = escapeHtml(v.id || ""); var btns = btn("edit","Edit","editVoucher",vid)+btn("email","Resend","resendEmail",vid)+btn("status","Status","changeStatus",vid); h += tag("tr",tag("td",tag("strong",vid))+tag("td",escapeHtml(v.claimedBy||"")+tag("br","")+tag("small",escapeHtml(v.email||"")))+tag("td",escapeHtml(v.amount||0)+" SEK")+tag("td",getTimeline(v.status))+tag("td",btns)); }); } document.getElementById("vouchersBody").innerHTML = h; }');
  html.push('function getTimeline(status) { var steps = ["Claimer","Qaid","Nazim","Sadar","Muhtamim"]; var curr = 0; var rejected = status === "rejected"; if (status === "pending_qaid") curr = 1; else if (status === "pending_nazim") curr = 2; else if (status === "pending_sadar") curr = 3; else if (status === "approved") curr = 4; else if (status === "paid") curr = 5; var h = ""; steps.forEach(function(s, i) { var cls = rejected ? "rejected" : (i < curr ? "completed" : (i === curr ? "active" : "pending")); var icon = rejected ? "&#10007;" : (i < curr ? "&#10003;" : (i === curr ? (i+1) : (i+1))); h += tag("div",tag("div",icon,"class=step-circle "+cls)+tag("div",s,"class=step-label"),"class=step-info"); if (i < steps.length - 1) h += tag("div","","class=step-line "+(i < curr ? "completed" : "")); }); return tag("div",h,"class=progress-bar"); }');
  html.push('function filterTable() { var search = document.getElementById("searchInput").value.toLowerCase(); var status = document.getElementById("statusFilter").value; var filtered = allVouchers.filter(function(v) { var matchSearch = v.id.toLowerCase().includes(search) || v.claimedBy.toLowerCase().includes(search) || v.email.toLowerCase().includes(search); var matchStatus = !status || v.status === status; return matchSearch && matchStatus; }); renderVouchers(filtered); }');
  html.push('function renderApproverEmails(approvers, pins) { var roleLabels = { qaidMajlis: "Qaid Majlis", nazimMaal: "Nazim Maal", sadarMka: "Sadar MKA", muhtamimMaal: "Muhtamim Maal" }; var h = ""; for (var key in approvers) { var label = roleLabels[key] || key; h += tag("div",tag("div",tag("strong",escapeHtml(label)),"style=margin-bottom:8px")+tag("div",tag("span","Email: ","style=color:var(--gray)")+tag("span",escapeHtml(approvers[key]||""),"style=word-break:break-all"),"style=font-size:13px"),"style=background:rgba(255,255,255,0.03);padding:15px;border-radius:8px;margin-bottom:12px;border:1px solid rgba(255,255,255,0.1)"); } document.getElementById("approverEmails").innerHTML = h || "No approvers"; }');
  html.push('function renderClaimerEmails(emails) { var h = ""; var Q = String.fromCharCode(39); if (emails.length === 0) { h = tag("tr",tag("td","No emails","colspan=3")); } else { emails.forEach(function(e) { var enc = encodeURIComponent(e.email||""); h += tag("tr",tag("td",escapeHtml(e.email||""))+tag("td",e.count||0)+tag("td",tag("button","Edit","class=\\"action-btn edit\\" onclick=\\"editClaimerEmail("+Q+enc+Q+")\\""))); }); } document.getElementById("claimerEmailsBody").innerHTML = h; }');
  html.push('function renderSettings(s) { document.getElementById("ssName").textContent = s.spreadsheetName || "N/A"; document.getElementById("folderName").textContent = s.folderName || "N/A"; }');
  html.push('function renderDetailedStats(stats) { var h = ""; h += tag("div",tag("span","Total:","class=settings-label")+tag("span",stats.total||0,"class=settings-value"),"class=settings-item"); h += tag("div",tag("span","Pending:","class=settings-label")+tag("span",stats.pending||0,"class=settings-value"),"class=settings-item"); h += tag("div",tag("span","Approved:","class=settings-label")+tag("span",stats.approved||0,"class=settings-value"),"class=settings-item"); h += tag("div",tag("span","Paid:","class=settings-label")+tag("span",stats.paid||0,"class=settings-value"),"class=settings-item"); document.getElementById("detailedStats").innerHTML = h; }');
  html.push('function renderAllowedEmails(emails) { var h = tag("p","Allowed emails:","style=font-size:13px;margin-bottom:10px"); var Q = String.fromCharCode(39); emails.forEach(function(email) { var enc = encodeURIComponent(email); h += tag("div",tag("span",escapeHtml(email),"class=settings-value")+tag("button","Remove","class=\\"action-btn status\\" onclick=\\"removeAllowedEmail("+Q+enc+Q+")\\""),"class=settings-item"); }); document.getElementById("allowedEmails").innerHTML = h; }');
  html.push('function removeAllowedEmail(enc) { var email = decodeURIComponent(enc); if (!confirm("Remove " + email + "?")) return; google.script.run.withSuccessHandler(function(r) { showToast(r.message, r.success ? "success" : "error"); refreshData(); }).removeAllowedEmailFromAdmin(email); }');
  html.push('function showTab(tabName) { document.querySelectorAll(".tab").forEach(function(t) { t.classList.remove("active"); }); document.querySelectorAll(".panel").forEach(function(p) { p.classList.remove("active"); }); event.target.classList.add("active"); document.getElementById(tabName + "-panel").classList.add("active"); }');
  html.push('function closeModal(id) { document.getElementById(id).classList.remove("show"); }');
  html.push('function showToast(msg, type) { var t = document.getElementById("toast"); t.textContent = msg; t.className = "toast show " + (type || ""); setTimeout(function() { t.classList.remove("show"); }, 3000); }');
  html.push('var currentEditId = null; var currentStatusId = null;');
  html.push('function editVoucher(id) { var v = allVouchers.find(function(x) { return x.id === id; }); if (!v) return; currentEditId = id; var h = tag("div",tag("label","ID:")+tag("input","","value=\\""+escapeHtml(v.id)+"\\" disabled"),"class=form-group"); h += tag("div",tag("label","Name:")+tag("input","","id=edit_name value=\\""+escapeHtml(v.claimedBy)+"\\""),"class=form-group"); h += tag("div",tag("label","Email:")+tag("input","","id=edit_email value=\\""+escapeHtml(v.email)+"\\""),"class=form-group"); h += tag("div",tag("label","Amount:")+tag("input","","id=edit_amount value=\\""+escapeHtml(v.amount)+"\\""),"class=form-group"); h += tag("button","Save","class=\\"btn btn-primary\\" onclick=\\"saveVoucherEdit()\\""); document.getElementById("editModalBody").innerHTML = h; document.getElementById("editModal").classList.add("show"); }');
  html.push('function saveVoucherEdit() { if (!currentEditId) return; var data = { id: currentEditId, claimedBy: document.getElementById("edit_name").value, email: document.getElementById("edit_email").value, amount: document.getElementById("edit_amount").value }; google.script.run.withSuccessHandler(function(r) { showToast(r.success ? "Updated!" : r.message, r.success ? "success" : "error"); closeModal("editModal"); currentEditId = null; refreshData(); }).updateVoucherFromAdmin(data); }');
  html.push('function resendEmail(id) { if (!confirm("Send reminder for " + id + "?")) return; google.script.run.withSuccessHandler(function(r) { showToast(r.success ? "Sent!" : r.message, r.success ? "success" : "error"); }).resendApprovalEmailFromAdmin(id); }');
  html.push('function changeStatus(id) { var v = allVouchers.find(function(x) { return x.id === id; }); if (!v) return; currentStatusId = id; var opts = tag("option","Pending Qaid","value=pending_qaid")+tag("option","Pending Nazim","value=pending_nazim")+tag("option","Pending Sadar","value=pending_sadar")+tag("option","Approved","value=approved")+tag("option","Paid","value=paid")+tag("option","Rejected","value=rejected"); var h = tag("div",tag("label","New Status:")+tag("select",opts,"id=new_status"),"class=form-group"); h += tag("button","Update","class=\\"btn btn-primary\\" onclick=\\"saveStatusChange()\\""); document.getElementById("genericModalTitle").textContent = "Change Status"; document.getElementById("genericModalBody").innerHTML = h; document.getElementById("genericModal").classList.add("show"); }');
  html.push('function saveStatusChange() { if (!currentStatusId) return; var ns = document.getElementById("new_status").value; google.script.run.withSuccessHandler(function(r) { showToast(r.success ? "Updated!" : r.message, r.success ? "success" : "error"); closeModal("genericModal"); currentStatusId = null; refreshData(); }).changeVoucherStatusFromAdmin(currentStatusId, ns); }');
  html.push('function showBackupModal() { document.getElementById("genericModalTitle").textContent = "Backup"; document.getElementById("genericModalBody").innerHTML = tag("p","Create backup?")+tag("button","Backup","class=\\"btn btn-primary\\" onclick=\\"doBackup()\\""); document.getElementById("genericModal").classList.add("show"); }');
  html.push('function doBackup() { google.script.run.withSuccessHandler(function(r) { showToast(r.success ? "Done!" : r.message, r.success ? "success" : "error"); closeModal("genericModal"); }).backupFromAdmin(); }');
  html.push('function showResetModal() { document.getElementById("genericModalTitle").textContent = "Reset"; document.getElementById("genericModalBody").innerHTML = tag("p","Delete ALL data?","style=color:red")+tag("p","Type RESET:")+tag("input","","id=resetConfirm")+tag("button","Reset","class=\\"btn btn-danger\\" onclick=\\"doReset()\\""); document.getElementById("genericModal").classList.add("show"); }');
  html.push('function doReset() { if (document.getElementById("resetConfirm").value !== "RESET") { showToast("Type RESET", "error"); return; } google.script.run.withSuccessHandler(function(r) { showToast(r.success ? "Done!" : r.message, r.success ? "success" : "error"); closeModal("genericModal"); refreshData(); }).resetFromAdmin(); }');
  html.push('function showTestEmailModal() { document.getElementById("genericModalTitle").textContent = "Test Email"; document.getElementById("genericModalBody").innerHTML = tag("div",tag("label","Email:")+tag("input","","id=testEmail type=email"),"class=form-group")+tag("button","Send","class=\\"btn btn-primary\\" onclick=\\"sendTestEmail()\\""); document.getElementById("genericModal").classList.add("show"); }');
  html.push('function sendTestEmail() { var email = document.getElementById("testEmail").value; if (!email) { showToast("Enter email", "error"); return; } google.script.run.withSuccessHandler(function(r) { showToast(r.success ? "Sent!" : r.message, r.success ? "success" : "error"); closeModal("genericModal"); }).sendTestEmailFromAdmin(email); }');
  html.push('function showNewVoucherModal() { document.getElementById("genericModalTitle").textContent = "Create New Voucher"; document.getElementById("genericModalBody").innerHTML = tag("div",tag("label","Claimed By (Name):")+tag("input","","id=test_name placeholder=\\"Enter name\\""),"class=form-group")+tag("div",tag("label","Email:")+tag("input","","id=test_email type=email placeholder=\\"email@example.com\\""),"class=form-group")+tag("div",tag("label","Description of Program:")+tag("textarea","","id=test_desc placeholder=\\"Program description\\" rows=2 style=\\"width:100%;padding:8px;border:1px solid #ddd;border-radius:4px;\\""),"class=form-group")+tag("div",tag("label","Amount in Numbers (SEK):")+tag("input","","id=test_amount type=number placeholder=\\"0\\""),"class=form-group")+tag("div",tag("label","Amount in Letters:")+tag("input","","id=test_amount_letters placeholder=\\"e.g. One Thousand\\""),"class=form-group")+tag("div",tag("label","Category:")+tag("select",tag("option","Select category...","value=")+tag("option","Travelling","value=Travelling")+tag("option","Ishaat","value=Ishaat")+tag("option","Ziafat","value=Ziafat")+tag("option","Sehat-e-Jismani","value=Sehat-e-Jismani")+tag("option","Atfal","value=Atfal")+tag("option","Telephone","value=Telephone")+tag("option","Office","value=Office")+tag("option","Contingency","value=Contingency")+tag("option","National","value=National")+tag("option","Ijtema","value=Ijtema")+tag("option","Other","value=Other"),"id=test_category style=\\"width:100%;padding:8px;border:1px solid #ddd;border-radius:4px;\\""),"class=form-group")+tag("div",tag("label","Majlis:")+tag("select",tag("option","Select majlis...","value=")+tag("option","Stockholm","value=Stockholm")+tag("option","Göteborg","value=Göteborg")+tag("option","Malmö","value=Malmö")+tag("option","Luleå","value=Luleå")+tag("option","Kalmar","value=Kalmar")+tag("option","National","value=National")+tag("option","Other","value=Other"),"id=test_majlis style=\\"width:100%;padding:8px;border:1px solid #ddd;border-radius:4px;\\""),"class=form-group")+tag("div",tag("label","Attachment URL (optional):")+tag("input","","id=test_attachment placeholder=\\"Google Drive link (optional)\\""),"class=form-group")+tag("button","Create Voucher","class=\\"btn btn-primary\\" onclick=\\"createTestVoucher()\\""); document.getElementById("genericModal").classList.add("show"); }');
  html.push('function createTestVoucher() { var name = document.getElementById("test_name").value; var email = document.getElementById("test_email").value; var desc = document.getElementById("test_desc").value; var amount = document.getElementById("test_amount").value; var amountLetters = document.getElementById("test_amount_letters").value; var category = document.getElementById("test_category").value; var majlis = document.getElementById("test_majlis").value; var attachment = document.getElementById("test_attachment").value; if (!name || !email || !amount || !category || !majlis) { showToast("Fill all required fields", "error"); return; } var data = { name: name, email: email, description: desc || "Voucher", amount: amount, amountLetters: amountLetters || amount + " SEK", category: category, majlis: majlis, attachmentUrl: attachment || "" }; showToast("Creating...", ""); google.script.run.withSuccessHandler(function(r) { if (r.success) { showToast("Created: " + r.id, "success"); closeModal("genericModal"); refreshData(); } else { showToast(r.message || "Error", "error"); } }).withFailureHandler(function(err) { showToast("Error: " + err, "error"); }).createTestVoucherFromAdmin(data); }');
  html.push('function editClaimerEmail(enc) { var old = decodeURIComponent(enc); var Q = String.fromCharCode(39); document.getElementById("genericModalTitle").textContent = "Update Email"; document.getElementById("genericModalBody").innerHTML = tag("p","Current: " + escapeHtml(old))+tag("div",tag("label","New:")+tag("input","","id=new_claimer_email"),"class=form-group")+tag("button","Update All","class=\\"btn btn-primary\\" onclick=\\"updateClaimerEmails("+Q+enc+Q+")\\""); document.getElementById("genericModal").classList.add("show"); }');
  html.push('function updateClaimerEmails(enc) { var old = decodeURIComponent(enc); var newEmail = document.getElementById("new_claimer_email").value; if (!newEmail) { showToast("Enter email", "error"); return; } google.script.run.withSuccessHandler(function(r) { showToast(r.success ? "Updated!" : r.message, r.success ? "success" : "error"); closeModal("genericModal"); refreshData(); }).updateClaimerEmailsFromAdmin(old, newEmail); }');
  html.push('function updateApproverEmail() { var role = document.getElementById("approverRole").value; var email = document.getElementById("newApproverEmail").value; if (!email) { showToast("Enter email", "error"); return; } google.script.run.withSuccessHandler(function(r) { showToast(r.message, r.success ? "success" : "error"); document.getElementById("newApproverEmail").value = ""; refreshData(); }).updateApproverEmailFromAdmin(role, email); }');
  html.push('function showTestSignatureModal() { var roleOpts = tag("option","Qaid Majlis","value=qaidMajlis")+tag("option","Nazim Maal","value=nazimMaal")+tag("option","Sadar MKA","value=sadarMka")+tag("option","Muhtamim Maal","value=muhtamimMaal"); var h = tag("p","Send a test signature email to verify an approver can receive emails and write signatures.","style=color:var(--gray);margin-bottom:20px;font-size:13px"); h += tag("div",tag("label","Select Approver Role:")+tag("select",roleOpts,"id=testSigRole style=\\"width:100%;padding:12px;background:var(--darker);border:1px solid rgba(255,255,255,0.2);border-radius:8px;color:var(--light);\\""),"class=form-group"); h += tag("div",tag("label","Custom Email (optional):")+tag("input","","id=testSigEmail type=email placeholder=\\"Leave blank to use configured email\\""),"class=form-group"); h += tag("div",tag("p","This will send:","style=font-weight:bold;margin-bottom:8px")+tag("ul",tag("li","Test signature link")+tag("li","Instructions to verify"),"style=color:var(--gray);font-size:13px;margin-left:20px"),"style=background:rgba(0,200,83,0.1);padding:15px;border-radius:8px;margin-bottom:20px"); h += tag("button","Send Test Signature Email","class=\\"btn btn-primary\\" onclick=\\"sendTestSignatureEmailJS()\\""); document.getElementById("genericModalTitle").textContent = "Test Signature Email"; document.getElementById("genericModalBody").innerHTML = h; document.getElementById("genericModal").classList.add("show"); }');
  html.push('function sendTestSignatureEmailJS() { var role = document.getElementById("testSigRole").value; var customEmail = document.getElementById("testSigEmail").value; if (!role) { showToast("Select a role", "error"); return; } showToast("Sending test email...", ""); google.script.run.withSuccessHandler(function(r) { if (r.success) { showToast("Sent to " + r.email + "!", "success"); closeModal("genericModal"); } else { showToast(r.message || "Error", "error"); } }).withFailureHandler(function(err) { showToast("Error: " + err, "error"); }).sendTestSignatureEmailFromAdmin(role, customEmail); }');
  html.push('function addAllowedEmail() { var email = document.getElementById("newAllowedEmail").value; if (!email) { showToast("Enter email", "error"); return; } google.script.run.withSuccessHandler(function(r) { showToast(r.message, r.success ? "success" : "error"); document.getElementById("newAllowedEmail").value = ""; refreshData(); }).addAllowedEmailFromAdmin(email); }');
  html.push('function renderAuthorizedClaimers(claimers) { var h = ""; var Q = String.fromCharCode(39); if (!claimers || claimers.length === 0) { h = tag("p","No authorized claimers yet","style=color:var(--gray)"); } else { claimers.forEach(function(email) { var enc = encodeURIComponent(email); h += tag("div",tag("span",escapeHtml(email),"class=settings-value")+tag("button","Remove","class=\\"action-btn status\\" onclick=\\"removeAuthorizedClaimer("+Q+enc+Q+")\\""),"class=settings-item"); }); } document.getElementById("authorizedClaimers").innerHTML = h; }');
  html.push('function addAuthorizedClaimer() { var email = document.getElementById("newClaimerEmail").value; if (!email) { showToast("Enter email", "error"); return; } google.script.run.withSuccessHandler(function(r) { showToast(r.message, r.success ? "success" : "error"); document.getElementById("newClaimerEmail").value = ""; refreshData(); }).withFailureHandler(function(err) { showToast("Error: " + err, "error"); }).addAuthorizedClaimerFromAdmin(email); }');
  html.push('function removeAuthorizedClaimer(enc) { var email = decodeURIComponent(enc); if (!confirm("Remove " + email + "?")) return; google.script.run.withSuccessHandler(function(r) { showToast(r.message, r.success ? "success" : "error"); refreshData(); }).withFailureHandler(function(err) { showToast("Error: " + err, "error"); }).removeAuthorizedClaimerFromAdmin(email); }');
  html.push('</scr' + 'ipt>');
  html.push('</body>');
  html.push('</html>');
  return html.join('');
}
