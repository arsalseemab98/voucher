/**
 * ============================================
 * CONFIG.gs - Configuration & Settings
 * ============================================
 *
 * This file contains:
 * - All configuration constants
 * - Dynamic configuration functions (Script Properties)
 * - Approval flow definition
 * - Allowed emails management
 */

// ============ MAIN CONFIGURATION ============
const CONFIG = {
  APPROVERS: {
    qaidMajlis: "agnagn1942@gmail.com",
    nazimMaal: "agnagn1942@gmail.com",
    sadarMka: "agnagn1942@gmail.com",
    muhtamimMaal: "agnagn1942@gmail.com"
  },
  ALLOWED_EMAILS: [
    "agnagn1942@gmail.com"
  ],
  VOUCHER_FOLDER_NAME: "Voucher (File responses)",
  SPREADSHEET_NAME: "MKA Voucher Database",
  WEBAPP_URL: "https://script.google.com/macros/s/AKfycbwUB12dwmKrOTEm1_45_FMLIYyOReSjsflqXJtfE-ntGgZvP2VaFCL3dIe--uh_j0lS/exec",
  ORG: {
    name: "Majlis Khuddam-ul-Ahmadiyya",
    address: "Tolvskillingsgatan 1",
    postalCode: "414 82 Göteborg",
    country: "Sverige"
  },
  LOCK_TIMEOUT_MS: 30000
};

// Vercel redirect URL for avoiding Chrome session issues
const VERCEL_REDIRECT_URL = "https://mka-voucher.vercel.app";


// ============ DYNAMIC CONFIGURATION ============

/**
 * Get approver email - checks Script Properties first, falls back to CONFIG
 */
function getApproverEmail(role) {
  try {
    var props = PropertiesService.getScriptProperties();
    var email = props.getProperty('approver_' + role);
    if (email && email.trim() !== '') {
      return email.trim();
    }
  } catch (e) {
    Logger.log("Error reading Script Properties: " + e);
  }
  return CONFIG.APPROVERS[role] || "";
}

/**
 * Set approver email in Script Properties
 */
function setApproverEmail(role, email) {
  try {
    var props = PropertiesService.getScriptProperties();
    props.setProperty('approver_' + role, email.trim());
    Logger.log("Updated " + role + " email to: " + email);
    return { success: true };
  } catch (e) {
    Logger.log("Error saving Script Properties: " + e);
    return { success: false, message: e.toString() };
  }
}

/**
 * Get all approver emails
 */
function getAllApproverEmails() {
  return {
    qaidMajlis: getApproverEmail('qaidMajlis'),
    nazimMaal: getApproverEmail('nazimMaal'),
    sadarMka: getApproverEmail('sadarMka'),
    muhtamimMaal: getApproverEmail('muhtamimMaal')
  };
}

/**
 * Get allowed emails - checks Script Properties first, falls back to CONFIG
 */
function getAllowedEmails() {
  try {
    var props = PropertiesService.getScriptProperties();
    var emailsStr = props.getProperty('allowed_emails');
    if (emailsStr && emailsStr.trim() !== '') {
      return emailsStr.split(',').map(function(e) { return e.trim(); }).filter(function(e) { return e !== ''; });
    }
  } catch (e) {
    Logger.log("Error reading allowed emails: " + e);
  }
  return CONFIG.ALLOWED_EMAILS;
}

/**
 * Add an allowed email
 */
function addAllowedEmailToProps(email) {
  try {
    var emails = getAllowedEmails();
    if (!emails.includes(email.trim().toLowerCase())) {
      emails.push(email.trim().toLowerCase());
    }
    var props = PropertiesService.getScriptProperties();
    props.setProperty('allowed_emails', emails.join(','));
    return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Remove an allowed email
 */
function removeAllowedEmailFromProps(email) {
  try {
    var emails = getAllowedEmails();
    emails = emails.filter(function(e) { return e.toLowerCase() !== email.trim().toLowerCase(); });
    var props = PropertiesService.getScriptProperties();
    props.setProperty('allowed_emails', emails.join(','));
    return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}


// ============ APPROVAL FLOW ============

/**
 * Get approval flow with current emails from Script Properties
 */
function getApprovalFlow() {
  return [
    { key: "qaid", title: "Qaid Majlis", email: getApproverEmail('qaidMajlis'), statusBefore: "pending_qaid", statusAfter: "pending_nazim" },
    { key: "nazimMaal", title: "Nazim Maal", email: getApproverEmail('nazimMaal'), statusBefore: "pending_nazim", statusAfter: "pending_sadar" },
    { key: "sadarMka", title: "Sadar MKA", email: getApproverEmail('sadarMka'), statusBefore: "pending_sadar", statusAfter: "approved" }
  ];
}

// Static approval flow for backward compatibility
const APPROVAL_FLOW = [
  { key: "qaid", title: "Qaid Majlis", email: CONFIG.APPROVERS.qaidMajlis, statusBefore: "pending_qaid", statusAfter: "pending_nazim" },
  { key: "nazimMaal", title: "Nazim Maal", email: CONFIG.APPROVERS.nazimMaal, statusBefore: "pending_nazim", statusAfter: "pending_sadar" },
  { key: "sadarMka", title: "Sadar MKA", email: CONFIG.APPROVERS.sadarMka, statusBefore: "pending_sadar", statusAfter: "approved" }
];


// ============ HELPER FUNCTIONS ============

/**
 * Create a Vercel redirect URL to avoid Chrome session issues
 */
function createRedirectUrl(action, id, approver) {
  var params = [];
  if (action) params.push("action=" + encodeURIComponent(action));
  if (id) params.push("id=" + encodeURIComponent(id));
  if (approver) params.push("approver=" + encodeURIComponent(approver));

  var url = VERCEL_REDIRECT_URL + "/?" + params.join("&");
  Logger.log("Created redirect URL: " + url);
  return url;
}

/**
 * Get approver info by key
 */
function getApproverInfo(approverKey) {
  return getApprovalFlow().find(a => a.key === approverKey) || null;
}

/**
 * Check if voucher is already approved by approver
 */
function isAlreadyApproved(voucher, approverKey) {
  if (!voucher) return false;
  const map = {
    qaid: voucher.qaidApproval?.approved === true,
    nazimMaal: voucher.nazimApproval?.approved === true,
    sadarMka: voucher.sadarApproval?.approved === true
  };
  return map[approverKey] === true;
}

/**
 * Check if it's approver's turn
 */
function isApproversTurn(voucher, approverKey) {
  if (!voucher) return false;
  const approver = getApproverInfo(approverKey);
  return approver && voucher.status === approver.statusBefore;
}

/**
 * Get approval progress for display
 */
function getApprovalProgress(voucher) {
  if (!voucher) return [];
  return APPROVAL_FLOW.map((approver, index) => {
    let status = "pending";
    let approved = false;

    if (approver.key === "qaid" && voucher.qaidApproval?.approved === true) {
      status = "approved"; approved = true;
    } else if (approver.key === "nazimMaal" && voucher.nazimApproval?.approved === true) {
      status = "approved"; approved = true;
    } else if (approver.key === "sadarMka" && voucher.sadarApproval?.approved === true) {
      status = "approved"; approved = true;
    } else if (voucher.status === approver.statusBefore) {
      status = "current";
    }

    return { step: index + 1, key: approver.key, title: approver.title, status, approved };
  });
}


// ============ AUTHORIZED CLAIMERS ============

/**
 * Get authorized claimers
 */
function getAuthorizedClaimers() {
  try {
    var props = PropertiesService.getScriptProperties();
    var claimersStr = props.getProperty('authorized_claimers');
    if (claimersStr && claimersStr.trim() !== '') {
      return claimersStr.split(',').map(function(e) { return e.trim(); }).filter(function(e) { return e !== ''; });
    }
  } catch (e) {
    Logger.log("Error reading authorized claimers: " + e);
  }
  return [];
}

/**
 * Check if email is authorized to claim
 */
function isAuthorizedClaimer(email) {
  var claimers = getAuthorizedClaimers();
  if (claimers.length === 0) {
    return true; // If no claimers are set, allow all
  }
  var normalizedEmail = email.trim().toLowerCase();
  return claimers.some(function(e) { return e.toLowerCase() === normalizedEmail; });
}
