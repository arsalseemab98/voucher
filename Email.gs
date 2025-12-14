/**
 * ============================================
 * EMAIL.gs - Email Functions
 * ============================================
 *
 * This file contains:
 * - All email sending functions
 * - Approval request emails
 * - Status update emails
 * - Payment confirmation emails
 * - Test email functions
 */


// ============ TEST SIGNATURE EMAIL ============

/**
 * Send test signature email to approver
 */
function sendTestSignatureEmail(role, customEmail) {
  try {
    var roleLabels = {
      'qaidMajlis': 'Qaid Majlis',
      'nazimMaal': 'Nazim Maal',
      'sadarMka': 'Sadar MKA',
      'muhtamimMaal': 'Muhtamim Maal'
    };

    var roleKeys = {
      'qaidMajlis': 'qaid',
      'nazimMaal': 'nazimMaal',
      'sadarMka': 'sadarMka',
      'muhtamimMaal': 'muhtamimMaal'
    };

    var email = customEmail || getApproverEmail(role);
    var roleLabel = roleLabels[role] || role;
    var approverKey = roleKeys[role] || role;

    if (!email) {
      return { success: false, message: "No email configured for " + role };
    }

    // Create a test voucher ID
    var testVoucherId = "TEST-SIG-" + Date.now();

    // Build the test approval link using Vercel redirect to avoid /u/X/ session issues
    var testLink = createRedirectUrl("approve", testVoucherId, approverKey);

    var subject = "MKA Voucher System - Test Signature Request";

    var htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
</head>
<body style="margin:0;padding:0;font-family:Arial,sans-serif;background:#f5f5f5;">
  <div style="max-width:600px;margin:0 auto;padding:20px;">
    <div style="background:linear-gradient(135deg,#1a5f2a 0%,#2d8f4e 100%);padding:30px;border-radius:12px 12px 0 0;text-align:center;">
      <h1 style="color:white;margin:0;font-size:24px;">Test Signature Request</h1>
      <p style="color:rgba(255,255,255,0.9);margin:10px 0 0 0;">MKA Voucher System</p>
    </div>

    <div style="background:white;padding:30px;border-radius:0 0 12px 12px;box-shadow:0 4px 15px rgba(0,0,0,0.1);">
      <p style="font-size:16px;color:#333;margin-bottom:20px;">
        Assalamu Alaikum <strong>${roleLabel}</strong>,
      </p>

      <div style="background:#fff3cd;border:1px solid #ffc107;border-radius:8px;padding:15px;margin-bottom:20px;">
        <p style="margin:0;color:#856404;font-size:14px;">
          <strong>Warning: This is a TEST email</strong> to verify you can receive emails and write signatures in our voucher system.
        </p>
      </div>

      <p style="color:#333;margin-bottom:20px;">
        Please click the button below to test signing. This will verify that:
      </p>
      <ul style="color:#555;margin-bottom:25px;">
        <li>You can receive approval emails</li>
        <li>You can access the approval page</li>
        <li>You can draw your signature</li>
      </ul>

      <div style="text-align:center;margin:25px 0;">
        <a href="${testLink}" style="display:inline-block;background:#1a5f2a;color:white;text-decoration:none;padding:16px 40px;border-radius:8px;font-weight:bold;font-size:16px;">
          Test Signature
        </a>
      </div>

      <p style="color:#666;font-size:13px;margin-top:20px;">
        If the button doesn't work, copy and paste this link:<br>
        <a href="${testLink}" style="color:#1a5f2a;word-break:break-all;">${testLink}</a>
      </p>

      <hr style="border:none;border-top:1px solid #eee;margin:25px 0;">

      <p style="color:#888;font-size:12px;text-align:center;margin:0;">
        Test sent: ${new Date().toLocaleString('sv-SE')}<br>
        Majlis Khuddam-ul-Ahmadiyya Sverige
      </p>
    </div>
  </div>
</body>
</html>`;

    GmailApp.sendEmail(email, subject, "Test signature request from MKA Voucher System.", {
      htmlBody: htmlBody,
      name: "MKA Voucher System"
    });

    Logger.log("Test signature email sent to " + email + " for role " + role);
    return { success: true, message: "Test email sent to " + email, email: email };

  } catch (error) {
    Logger.log("Error in sendTestSignatureEmail: " + error);
    return { success: false, message: error.toString() };
  }
}

/**
 * Admin portal function to send test signature email
 */
function sendTestSignatureEmailFromAdmin(role, customEmail) {
  return sendTestSignatureEmail(role, customEmail);
}


// ============ APPROVAL REQUEST EMAIL ============

/**
 * Send approval request email to approver
 */
function sendApprovalRequest(voucherData, approverKey, pdfUrl) {
  try {
    const approver = getApproverInfo(approverKey);
    if (!approver) {
      Logger.log("Invalid approver key: " + approverKey);
      return;
    }

    // If no pdfUrl provided, try to get it from the voucher
    if (!pdfUrl) {
      var fullVoucher = getVoucherById(voucherData.id);
      if (fullVoucher && fullVoucher.pdfUrl) {
        pdfUrl = fullVoucher.pdfUrl;
      }
    }

    // Create Vercel redirect URLs to avoid Google session /u/X/ issues
    const approvalUrl = createRedirectUrl("approve", voucherData.id, approverKey);
    const rejectUrl = createRedirectUrl("reject", voucherData.id, approverKey);

    const subject = "MKA Voucher Approval: " + voucherData.id + " - " + voucherData.amountInNumbers + " SEK";

    // Build PDF button HTML if available
    var pdfButtonHtml = '';
    if (pdfUrl) {
      pdfButtonHtml = `
          <div style="background: #e8f4ea; border-radius: 8px; padding: 15px; margin: 15px 0; text-align: center;">
            <p style="margin: 0 0 10px 0; font-size: 13px; color: #155724;"><strong>View Full Voucher Details (with receipt/attachment):</strong></p>
            <a href="${pdfUrl}" style="display: inline-block; background: #fd7e14; color: white; padding: 10px 25px; text-decoration: none; border-radius: 6px; font-weight: bold; font-size: 14px;">Download Voucher PDF</a>
          </div>
      `;
    }

    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: linear-gradient(135deg, #1a5f2a 0%, #2d8f4e 100%); color: white; padding: 25px; text-align: center; border-radius: 10px 10px 0 0;">
          <h1 style="margin: 0; font-size: 24px;">MKA Voucher System</h1>
          <p style="margin: 10px 0 0 0; opacity: 0.9;">Approval Request for ${approver.title}</p>
        </div>

        <div style="padding: 25px; background: #ffffff; border: 1px solid #e0e0e0;">
          <p style="font-size: 16px;">Assalamu Alaikum <strong>${approver.title}</strong>,</p>
          <p>A voucher requires your approval:</p>

          <div style="background: #f8f9fa; border-radius: 8px; padding: 20px; margin: 20px 0;">
            <table style="width: 100%; border-collapse: collapse;">
              <tr><td style="padding: 8px 0; color: #666; width: 140px;">Reference:</td><td style="padding: 8px 0; font-weight: bold;">${voucherData.id}</td></tr>
              <tr><td style="padding: 8px 0; color: #666;">Claimed By:</td><td style="padding: 8px 0;">${voucherData.claimedBy}</td></tr>
              <tr><td style="padding: 8px 0; color: #666;">Description:</td><td style="padding: 8px 0;">${voucherData.descriptionOfProgram}</td></tr>
              <tr><td style="padding: 8px 0; color: #666;">Amount:</td><td style="padding: 8px 0; font-size: 18px; font-weight: bold; color: #1a5f2a;">${voucherData.amountInNumbers} SEK</td></tr>
              <tr><td style="padding: 8px 0; color: #666;">Category:</td><td style="padding: 8px 0;">${voucherData.category}</td></tr>
              <tr><td style="padding: 8px 0; color: #666;">Majlis:</td><td style="padding: 8px 0;">${voucherData.majlis}</td></tr>
            </table>
          </div>

          ${voucherData.attachmentUrl ? `
          <div style="background: #fff3cd; border-radius: 8px; padding: 15px; margin: 15px 0; text-align: center; border: 1px solid #ffc107;">
            <p style="margin: 0 0 10px 0; font-size: 13px; color: #856404;"><strong>Receipt/Attachment:</strong></p>
            <a href="${voucherData.attachmentUrl}" target="_blank" style="display: inline-block; background: #ffc107; color: #000; padding: 10px 25px; text-decoration: none; border-radius: 6px; font-weight: bold; font-size: 14px;">View Attachment</a>
          </div>
          ` : ''}

          ${pdfButtonHtml}

          <div style="text-align: center; margin: 30px 0;">
            <a href="${approvalUrl}" style="display: inline-block; background: #1a5f2a; color: white; padding: 15px 50px; text-decoration: none; border-radius: 8px; font-weight: bold; font-size: 16px;">APPROVE & SIGN</a>
            <br><br>
            <a href="${rejectUrl}" style="display: inline-block; background: #dc3545; color: white; padding: 12px 40px; text-decoration: none; border-radius: 8px; font-size: 14px;">REJECT</a>
          </div>
        </div>

        <div style="background: #333; color: white; padding: 15px; text-align: center; font-size: 12px; border-radius: 0 0 10px 10px;">
          <p style="margin: 0;">Majlis Khuddam-ul-Ahmadiyya Sverige</p>
        </div>
      </div>
    `;

    GmailApp.sendEmail(approver.email, subject, "Please view this email in HTML format.", { htmlBody: htmlBody });
    Logger.log("Approval request sent to " + approver.title + " (" + approver.email + ") with PDF: " + (pdfUrl ? "Yes" : "No"));
  } catch (error) {
    Logger.log("Error sending approval request: " + error);
  }
}


// ============ CLAIMER CONFIRMATION EMAIL ============

/**
 * Send confirmation email to claimer after submission
 */
function sendClaimerConfirmation(voucherData, pdfUrl) {
  try {
    const subject = "Voucher Submitted: " + voucherData.id;

    const progressHtml = APPROVAL_FLOW.map((approver, i) =>
      '<div style="display:inline-block;text-align:center;margin:0 10px;">' +
        '<div style="width:30px;height:30px;border-radius:50%;background:' + (i === 0 ? '#ffc107' : '#e0e0e0') + ';color:' + (i === 0 ? '#000' : '#999') + ';line-height:30px;margin:0 auto;">' + (i + 1) + '</div>' +
        '<div style="font-size:11px;margin-top:5px;color:#666;">' + approver.title + '</div>' +
      '</div>'
    ).join('<div style="display:inline-block;width:30px;height:2px;background:#e0e0e0;vertical-align:middle;"></div>');

    // Build PDF button if available
    var pdfButtonHtml = '';
    if (pdfUrl) {
      pdfButtonHtml = `
          <div style="text-align: center; margin: 20px 0;">
            <a href="${pdfUrl}" style="display: inline-block; background: #fd7e14; color: white; padding: 12px 30px; text-decoration: none; border-radius: 6px; font-weight: bold;">Download Your Voucher PDF</a>
          </div>
      `;
    }

    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: linear-gradient(135deg, #1a5f2a 0%, #2d8f4e 100%); color: white; padding: 25px; text-align: center; border-radius: 10px 10px 0 0;">
          <h1 style="margin: 0;">Voucher Submitted</h1>
        </div>

        <div style="padding: 25px; background: #ffffff; border: 1px solid #e0e0e0;">
          <p>Assalamu Alaikum <strong>${voucherData.claimedBy}</strong>,</p>
          <p>Your voucher has been submitted and is now in the approval process.</p>

          <div style="background: #f8f9fa; border-radius: 8px; padding: 20px; margin: 20px 0; text-align: center;">
            <p style="margin: 0 0 10px 0; font-weight: bold;">Reference: ${voucherData.id}</p>
            <p style="margin: 0; font-size: 24px; color: #1a5f2a; font-weight: bold;">${voucherData.amountInNumbers} SEK</p>
          </div>

          ${pdfButtonHtml}

          <p style="font-weight: bold; margin-bottom: 15px;">Approval Progress:</p>
          <div style="text-align: center; padding: 20px 0;">${progressHtml}</div>

          <p style="color: #666; font-size: 13px;">You will receive email updates as your voucher progresses.</p>
          <p>JazakAllah,<br>MKA Finance Team</p>
        </div>

        <div style="background: #333; color: white; padding: 15px; text-align: center; font-size: 12px; border-radius: 0 0 10px 10px;">
          <p style="margin: 0;">Majlis Khuddam-ul-Ahmadiyya Sverige</p>
        </div>
      </div>
    `;

    GmailApp.sendEmail(voucherData.claimerEmail, subject, "Please view this email in HTML format.", { htmlBody: htmlBody });
  } catch (error) {
    Logger.log("Error sending claimer confirmation: " + error);
  }
}


// ============ STATUS UPDATE EMAIL ============

/**
 * Send status update to claimer when voucher is approved by an approver
 */
function sendStatusUpdateToClaimer(voucher, approverTitle) {
  try {
    const progress = getApprovalProgress(voucher);
    const completedSteps = progress.filter(p => p.approved).length;
    const totalSteps = APPROVAL_FLOW.length;

    const progressHtml = progress.map(p => {
      let bgColor = '#e0e0e0', textColor = '#999', icon = String(p.step);
      if (p.approved) { bgColor = '#28a745'; textColor = '#fff'; icon = 'OK'; }
      else if (p.status === 'current') { bgColor = '#ffc107'; textColor = '#000'; }

      return '<div style="display:inline-block;text-align:center;margin:0 10px;">' +
        '<div style="width:30px;height:30px;border-radius:50%;background:' + bgColor + ';color:' + textColor + ';line-height:30px;margin:0 auto;font-weight:bold;">' + icon + '</div>' +
        '<div style="font-size:11px;margin-top:5px;color:#666;">' + p.title + '</div>' +
      '</div>';
    }).join('<div style="display:inline-block;width:30px;height:2px;background:#e0e0e0;vertical-align:middle;"></div>');

    const subject = "Voucher Update: " + voucher.id + " - Approved by " + approverTitle;

    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: #28a745; color: white; padding: 20px; text-align: center; border-radius: 10px 10px 0 0;">
          <h1 style="margin: 0; font-size: 20px;">Voucher Progress Update</h1>
        </div>

        <div style="padding: 25px; background: #ffffff; border: 1px solid #e0e0e0;">
          <p>Assalamu Alaikum <strong>${voucher.claimedBy}</strong>,</p>
          <p>Your voucher <strong>${voucher.id}</strong> has been approved by <strong>${approverTitle}</strong>.</p>

          <div style="background: #d4edda; border-radius: 8px; padding: 15px; margin: 20px 0; text-align: center;">
            <p style="margin: 0; color: #155724; font-weight: bold;">Step ${completedSteps} of ${totalSteps} completed</p>
          </div>

          <p style="font-weight: bold; margin-bottom: 15px;">Progress:</p>
          <div style="text-align: center; padding: 20px 0;">${progressHtml}</div>

          <p>JazakAllah,<br>MKA Finance Team</p>
        </div>
      </div>
    `;

    GmailApp.sendEmail(voucher.claimerEmail, subject, "Please view this email in HTML format.", { htmlBody: htmlBody });
  } catch (error) {
    Logger.log("Error sending status update: " + error);
  }
}


// ============ FINAL NOTIFICATION EMAIL ============

/**
 * Send final notification when voucher is fully approved
 */
function sendFinalNotification(voucher, pdfUrl) {
  try {
    const subject = "Voucher APPROVED: " + voucher.id + " - Ready for Payment";

    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: linear-gradient(135deg, #28a745 0%, #20c997 100%); color: white; padding: 25px; text-align: center; border-radius: 10px 10px 0 0;">
          <h1 style="margin: 0; font-size: 28px;">FULLY APPROVED</h1>
          <p style="margin: 10px 0 0 0; font-size: 16px;">Voucher ${voucher.id}</p>
        </div>

        <div style="padding: 25px; background: #ffffff; border: 1px solid #e0e0e0;">
          <p style="font-size: 16px;">Assalamu Alaikum,</p>
          <p>This voucher has been approved by all required approvers and is ready for payment.</p>

          <div style="background: #f8f9fa; border-radius: 8px; padding: 20px; margin: 20px 0;">
            <table style="width: 100%; border-collapse: collapse;">
              <tr><td style="padding: 10px; border-bottom: 1px solid #e0e0e0; color: #666;">Reference:</td><td style="padding: 10px; border-bottom: 1px solid #e0e0e0; font-weight: bold;">${voucher.id}</td></tr>
              <tr><td style="padding: 10px; border-bottom: 1px solid #e0e0e0; color: #666;">Pay To:</td><td style="padding: 10px; border-bottom: 1px solid #e0e0e0; font-weight: bold;">${voucher.claimedBy}</td></tr>
              <tr><td style="padding: 10px; border-bottom: 1px solid #e0e0e0; color: #666;">Description:</td><td style="padding: 10px; border-bottom: 1px solid #e0e0e0;">${voucher.descriptionOfProgram}</td></tr>
              <tr><td style="padding: 10px; color: #666;">Amount:</td><td style="padding: 10px; font-size: 24px; font-weight: bold; color: #28a745;">${voucher.amountInNumbers} SEK</td></tr>
            </table>
          </div>

          <div style="background: #d4edda; border-radius: 8px; padding: 15px; margin: 20px 0;">
            <p style="margin: 0; color: #155724;"><strong>Approvals:</strong></p>
            <p style="margin: 5px 0 0 0; color: #155724;">OK Qaid Majlis - Approved</p>
            <p style="margin: 5px 0 0 0; color: #155724;">OK Nazim Maal - Approved</p>
            <p style="margin: 5px 0 0 0; color: #155724;">OK Sadar MKA - Approved</p>
          </div>

          ${pdfUrl ? '<div style="text-align: center; margin: 25px 0;"><a href="' + pdfUrl + '" style="display: inline-block; background: #1a5f2a; color: white; padding: 15px 40px; text-decoration: none; border-radius: 8px; font-weight: bold; font-size: 16px;">Download Voucher PDF</a></div>' : ''}

          <p>JazakAllah,<br>MKA Finance System</p>
        </div>

        <div style="background: #333; color: white; padding: 15px; text-align: center; font-size: 12px; border-radius: 0 0 10px 10px;">
          <p style="margin: 0;">Majlis Khuddam-ul-Ahmadiyya Sverige</p>
        </div>
      </div>
    `;

    // Send payment request to Muhtamim Maal with payment action link
    sendMuhtamimMaalPaymentRequest(voucher, pdfUrl);

    // Send to claimer
    try {
      GmailApp.sendEmail(voucher.claimerEmail, "Your Voucher " + voucher.id + " is APPROVED!", "Please view this email in HTML format.", { htmlBody: htmlBody });
    } catch (claimerError) {
      Logger.log("Failed to send to claimer (" + voucher.claimerEmail + "): " + claimerError);
    }
  } catch (error) {
    Logger.log("Error sending final notification: " + error);
  }
}


// ============ MUHTAMIM MAAL PAYMENT REQUEST ============

/**
 * Send payment request to Muhtamim Maal
 */
function sendMuhtamimMaalPaymentRequest(voucher, pdfUrl) {
  try {
    const paymentUrl = CONFIG.WEBAPP_URL + "?action=payment&id=" + encodeURIComponent(voucher.id);
    const subject = "PAYMENT REQUIRED: Voucher " + voucher.id + " - " + voucher.amountInNumbers + " SEK";

    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: linear-gradient(135deg, #28a745 0%, #1e7e34 100%); color: white; padding: 25px; text-align: center; border-radius: 10px 10px 0 0;">
          <h1 style="margin: 0; font-size: 24px;">Payment Required</h1>
          <p style="margin: 10px 0 0 0; opacity: 0.9;">Muhtamim Maal - Action Required</p>
        </div>

        <div style="padding: 25px; background: #ffffff; border: 1px solid #e0e0e0;">
          <p style="font-size: 16px;">Assalamu Alaikum <strong>Muhtamim Maal</strong>,</p>
          <p>The following voucher has been approved and requires payment:</p>

          <div style="background: #f8f9fa; border-radius: 8px; padding: 20px; margin: 20px 0;">
            <table style="width: 100%; border-collapse: collapse;">
              <tr><td style="padding: 8px 0; color: #666; width: 140px;">Reference:</td><td style="padding: 8px 0; font-weight: bold;">${voucher.id}</td></tr>
              <tr><td style="padding: 8px 0; color: #666;">Pay To:</td><td style="padding: 8px 0; font-weight: bold;">${voucher.claimedBy}</td></tr>
              <tr><td style="padding: 8px 0; color: #666;">Description:</td><td style="padding: 8px 0;">${voucher.descriptionOfProgram}</td></tr>
              <tr><td style="padding: 8px 0; color: #666;">Category:</td><td style="padding: 8px 0;">${voucher.category}</td></tr>
              <tr><td style="padding: 8px 0; color: #666;">Amount:</td><td style="padding: 8px 0; font-size: 20px; font-weight: bold; color: #28a745;">${voucher.amountInNumbers} SEK</td></tr>
            </table>
          </div>

          <div style="text-align: center; margin: 30px 0;">
            <a href="${paymentUrl}" style="display: inline-block; background: #28a745; color: white; padding: 15px 50px; text-decoration: none; border-radius: 8px; font-weight: bold; font-size: 16px;">REGISTER PAYMENT</a>
          </div>

          ${pdfUrl ? '<p style="text-align: center;"><a href="' + pdfUrl + '" style="color: #28a745;">Download Voucher PDF</a></p>' : ''}

          <p>JazakAllah,<br>MKA Finance System</p>
        </div>

        <div style="background: #333; color: white; padding: 15px; text-align: center; font-size: 12px; border-radius: 0 0 10px 10px;">
          <p style="margin: 0;">Majlis Khuddam-ul-Ahmadiyya Sverige</p>
        </div>
      </div>
    `;

    var muhtamimMaalEmail = getApproverEmail('muhtamimMaal');
    GmailApp.sendEmail(muhtamimMaalEmail, subject, "Please view this email in HTML format.", { htmlBody: htmlBody });
    Logger.log("Payment request sent to Muhtamim Maal: " + muhtamimMaalEmail);
  } catch (error) {
    Logger.log("Error sending Muhtamim Maal payment request: " + error);
  }
}


// ============ PAYMENT CONFIRMATION EMAIL ============

/**
 * Send payment confirmation to claimer
 */
function sendPaymentConfirmationToClaimer(voucher, paymentMethod, paymentDate, pdfUrl, accountInfo) {
  try {
    const subject = "PAYMENT COMPLETE: Voucher " + voucher.id + " - Download Receipt";

    // Build PDF button HTML if PDF exists
    var pdfButtonHtml = '';
    if (pdfUrl) {
      pdfButtonHtml = `
          <div style="text-align: center; margin: 25px 0;">
            <a href="${pdfUrl}" style="display: inline-block; background: #1a5f2a; color: white; padding: 15px 40px; text-decoration: none; border-radius: 8px; font-weight: bold; font-size: 16px;">Download Voucher PDF</a>
          </div>
          <p style="text-align: center; font-size: 12px; color: #666;">Keep this PDF for your records</p>
      `;
    }

    // Show account info row only if provided
    var accountInfoRow = '';
    if (accountInfo && accountInfo.trim()) {
      accountInfoRow = `<tr><td style="padding: 8px 0; color: #155724;">Account Number:</td><td style="padding: 8px 0; color: #155724;">${accountInfo}</td></tr>`;
    }

    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: linear-gradient(135deg, #28a745 0%, #20c997 100%); color: white; padding: 25px; text-align: center; border-radius: 10px 10px 0 0;">
          <h1 style="margin: 0; font-size: 28px;">PAYMENT COMPLETE</h1>
          <p style="margin: 10px 0 0 0; font-size: 16px;">Voucher ${voucher.id}</p>
        </div>

        <div style="padding: 25px; background: #ffffff; border: 1px solid #e0e0e0;">
          <p style="font-size: 16px;">Assalamu Alaikum <strong>${voucher.claimedBy}</strong>,</p>
          <p>Great news! Your voucher has been processed and payment has been made.</p>

          <div style="background: #d4edda; border-radius: 8px; padding: 20px; margin: 20px 0;">
            <table style="width: 100%; border-collapse: collapse;">
              <tr><td style="padding: 8px 0; color: #155724;">Reference:</td><td style="padding: 8px 0; font-weight: bold; color: #155724;">${voucher.id}</td></tr>
              <tr><td style="padding: 8px 0; color: #155724;">Amount:</td><td style="padding: 8px 0; font-weight: bold; color: #155724; font-size: 18px;">${voucher.amountInNumbers} SEK</td></tr>
              <tr><td style="padding: 8px 0; color: #155724;">Payment Method:</td><td style="padding: 8px 0; color: #155724;">${paymentMethod}</td></tr>
              ${accountInfoRow}
              <tr><td style="padding: 8px 0; color: #155724;">Payment Date:</td><td style="padding: 8px 0; color: #155724;">${paymentDate}</td></tr>
            </table>
          </div>

          ${pdfButtonHtml}

          <p>If you have any questions about this payment, please contact the finance team.</p>
          <p>JazakAllah,<br>MKA Finance Team</p>
        </div>

        <div style="background: #333; color: white; padding: 15px; text-align: center; font-size: 12px; border-radius: 0 0 10px 10px;">
          <p style="margin: 0;">Majlis Khuddam-ul-Ahmadiyya Sverige</p>
        </div>
      </div>
    `;

    GmailApp.sendEmail(voucher.claimerEmail, subject, "Please view this email in HTML format.", { htmlBody: htmlBody });
    Logger.log("Payment confirmation with PDF sent to: " + voucher.claimerEmail);
  } catch (error) {
    Logger.log("Error sending payment confirmation: " + error);
  }
}


// ============ TEST EMAIL FROM ADMIN ============

/**
 * Send simple test email from admin portal
 */
function sendTestEmailFromAdmin(email) {
  try {
    var subject = "MKA Voucher System - Test Email";
    var htmlBody = '<div style="font-family:Arial;padding:20px;"><h2 style="color:#1a5f2a;">Test Email</h2><p>This is a test email from the MKA Voucher System.</p><p>If you received this, email delivery is working correctly.</p><p style="color:#666;font-size:12px;">Sent: ' + new Date().toISOString() + '</p></div>';

    GmailApp.sendEmail(email, subject, "Test email from MKA Voucher System", { htmlBody: htmlBody });
    return { success: true };
  } catch (error) {
    Logger.log("Error in sendTestEmailFromAdmin: " + error);
    return { success: false, message: error.toString() };
  }
}
