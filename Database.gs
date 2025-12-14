/**
 * ============================================
 * DATABASE.gs - Spreadsheet & Data Functions
 * ============================================
 *
 * This file contains:
 * - Spreadsheet creation and management
 * - Voucher CRUD operations
 * - Data retrieval functions
 * - Drive/File operations
 */

// ============ VOUCHER ID GENERATOR ============

/**
 * Generate unique voucher ID
 * Format: MKA-YYYYMMDD-HHMMSSxx
 */
function generateVoucherId() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  const hour = String(now.getHours()).padStart(2, '0');
  const min = String(now.getMinutes()).padStart(2, '0');
  const sec = String(now.getSeconds()).padStart(2, '0');
  const random = Math.floor(Math.random() * 100).toString().padStart(2, '0');
  return `MKA-${year}${month}${day}-${hour}${min}${sec}${random}`;
}


// ============ SPREADSHEET FUNCTIONS ============

/**
 * Get or create the voucher database spreadsheet
 */
function getOrCreateSpreadsheet() {
  try {
    const files = DriveApp.getFilesByName(CONFIG.SPREADSHEET_NAME);
    if (files.hasNext()) {
      return SpreadsheetApp.open(files.next());
    }

    const ss = SpreadsheetApp.create(CONFIG.SPREADSHEET_NAME);
    const sheet = ss.getActiveSheet();
    sheet.setName("Vouchers");

    const headers = [
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
    sheet.setFrozenRows(1);

    return ss;
  } catch (error) {
    Logger.log("Error in getOrCreateSpreadsheet: " + error);
    throw error;
  }
}


// ============ VOUCHER CRUD ============

/**
 * Save new voucher to spreadsheet
 */
function saveVoucher(data) {
  const ss = getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName("Vouchers");

  const row = [
    data.id, data.timestamp, data.claimedBy, data.claimerEmail,
    data.descriptionOfProgram, data.amountInNumbers, data.amountInLetters,
    data.category, data.majlis, data.attachmentUrl, data.status,
    false, "", "",  // Qaid
    false, "", "",  // Nazim
    false, "", "",  // Sadar
    "",             // PDF URL
    false, "", "",  // Muhtamim Maal
    false, "", "", "", "",  // Payment info
    "", "", ""      // Bank details
  ];

  sheet.appendRow(row);
}

/**
 * Get voucher by ID
 */
function getVoucherById(voucherId) {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName("Vouchers");
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === voucherId) {
        return {
          row: i + 1,
          id: data[i][0],
          timestamp: data[i][1],
          claimedBy: data[i][2],
          claimerEmail: data[i][3],
          descriptionOfProgram: data[i][4],
          amountInNumbers: data[i][5],
          amountInLetters: data[i][6],
          category: data[i][7],
          majlis: data[i][8],
          attachmentUrl: data[i][9],
          status: data[i][10],
          qaidApproval: { approved: data[i][11] === true, signature: data[i][12], date: data[i][13] },
          nazimApproval: { approved: data[i][14] === true, signature: data[i][15], date: data[i][16] },
          sadarApproval: { approved: data[i][17] === true, signature: data[i][18], date: data[i][19] },
          pdfUrl: data[i][20],
          muhtamimMaalApproval: { approved: data[i][21] === true, signature: data[i][22], date: data[i][23] },
          payment: {
            paid: data[i][24] === true,
            paidBy: data[i][25],
            paymentDate: data[i][26],
            paymentMethod: data[i][27],
            paidTo: data[i][28],
            bankName: data[i][29],
            accountNumber: data[i][30],
            clearingNumber: data[i][31]
          }
        };
      }
    }
    return null;
  } catch (error) {
    Logger.log("Error in getVoucherById: " + error);
    return null;
  }
}

/**
 * Get all vouchers for admin display
 */
function getAllVouchers() {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName("Vouchers");
    const data = sheet.getDataRange().getValues();
    const vouchers = [];

    for (let i = 1; i < data.length; i++) {
      vouchers.push({
        id: data[i][0],
        timestamp: data[i][1],
        claimedBy: data[i][2],
        email: data[i][3],
        description: data[i][4],
        amount: data[i][5],
        category: data[i][7],
        majlis: data[i][8],
        status: data[i][10],
        pdfUrl: data[i][20]
      });
    }

    return vouchers;
  } catch (error) {
    Logger.log("Error in getAllVouchers: " + error);
    return [];
  }
}

/**
 * Update voucher field in spreadsheet
 */
function updateVoucherField(voucherId, column, value) {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName("Vouchers");
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === voucherId) {
        sheet.getRange(i + 1, column).setValue(value);
        return true;
      }
    }
    return false;
  } catch (error) {
    Logger.log("Error in updateVoucherField: " + error);
    return false;
  }
}

/**
 * Update voucher status
 */
function updateVoucherStatus(voucherId, newStatus) {
  return updateVoucherField(voucherId, 11, newStatus); // Column K (11) = Status
}


// ============ DRIVE/FILE FUNCTIONS ============

/**
 * Get or create voucher folder in Drive
 */
function getVoucherFolder() {
  try {
    const folders = DriveApp.getFoldersByName(CONFIG.VOUCHER_FOLDER_NAME);
    if (folders.hasNext()) {
      return folders.next();
    }
    return DriveApp.createFolder(CONFIG.VOUCHER_FOLDER_NAME);
  } catch (error) {
    Logger.log("Error getting/creating voucher folder: " + error);
    return DriveApp.getRootFolder();
  }
}

/**
 * Convert base64 signature to blob
 */
function signatureToBlob(base64Data, name) {
  try {
    if (!base64Data || !base64Data.includes(',')) {
      return null;
    }
    const base64 = base64Data.split(',')[1];
    const decoded = Utilities.base64Decode(base64);
    return Utilities.newBlob(decoded, 'image/png', name + '.png');
  } catch (error) {
    Logger.log("Error converting signature to blob: " + error);
    return null;
  }
}

/**
 * Get attachment from Google Drive URL
 */
function getAttachmentFromUrl(url) {
  try {
    if (!url || url.trim() === '') {
      return null;
    }

    var fileId = null;

    // Format: https://drive.google.com/open?id=FILE_ID
    var match = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if (match) fileId = match[1];

    // Format: https://drive.google.com/file/d/FILE_ID/view
    if (!fileId) {
      match = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
      if (match) fileId = match[1];
    }

    // Format: https://drive.google.com/uc?id=FILE_ID
    if (!fileId) {
      match = url.match(/\/uc\?.*id=([a-zA-Z0-9_-]+)/);
      if (match) fileId = match[1];
    }

    if (!fileId) {
      Logger.log("Could not extract file ID from URL: " + url);
      return null;
    }

    var file = DriveApp.getFileById(fileId);
    return {
      blob: file.getBlob(),
      mimeType: file.getMimeType(),
      name: file.getName(),
      url: url
    };
  } catch (error) {
    Logger.log("Error getting attachment from URL: " + error);
    return null;
  }
}

/**
 * Share attachment with anyone who has the link
 */
function shareAttachmentWithAnyone(url) {
  try {
    if (!url || url.trim() === '') return false;

    var fileId = null;
    var match = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if (match) fileId = match[1];

    if (!fileId) {
      match = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
      if (match) fileId = match[1];
    }

    if (!fileId) {
      match = url.match(/\/uc\?.*id=([a-zA-Z0-9_-]+)/);
      if (match) fileId = match[1];
    }

    if (!fileId) {
      Logger.log("shareAttachmentWithAnyone: Could not extract file ID");
      return false;
    }

    var file = DriveApp.getFileById(fileId);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    Logger.log("Shared file: " + fileId);
    return true;
  } catch (error) {
    Logger.log("Error sharing attachment: " + error);
    return false;
  }
}


// ============ STATISTICS ============

/**
 * Get voucher statistics for admin
 */
function getVoucherStats() {
  try {
    const vouchers = getAllVouchers();
    const stats = {
      total: vouchers.length,
      pending: 0,
      approved: 0,
      paid: 0,
      rejected: 0
    };

    vouchers.forEach(function(v) {
      if (v.status === 'rejected') stats.rejected++;
      else if (v.status === 'paid') stats.paid++;
      else if (v.status === 'approved') stats.approved++;
      else stats.pending++;
    });

    return stats;
  } catch (error) {
    Logger.log("Error in getVoucherStats: " + error);
    return { total: 0, pending: 0, approved: 0, paid: 0, rejected: 0 };
  }
}
