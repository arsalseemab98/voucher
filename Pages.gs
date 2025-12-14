/**
 * ============================================
 * PAGES.gs - HTML Page Functions
 * ============================================
 *
 * This file contains:
 * - Signature approval page
 * - Reject page
 * - Payment page (Muhtamim Maal)
 * - Error pages (Access denied, Already approved, Not your turn, Already paid)
 * - Home page
 *
 * Note: Admin Portal page is in AdminPortal.gs due to its size
 */


// ============ ACCESS DENIED PAGE ============

function getAccessDeniedPage(email) {
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Access Denied - MKA Voucher System</title>
  <style>
    body { font-family: Arial, sans-serif; background: #f0f2f5; min-height: 100vh; display: flex; align-items: center; justify-content: center; margin: 0; padding: 15px; }
    .box { background: white; padding: 40px; border-radius: 12px; text-align: center; max-width: 450px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); }
    .icon { font-size: 50px; margin-bottom: 20px; }
    h1 { color: #dc3545; margin-bottom: 15px; font-size: 22px; }
    p { color: #666; line-height: 1.6; }
    .email-box { background: #f8d7da; color: #721c24; padding: 10px 15px; border-radius: 8px; margin: 15px 0; font-family: monospace; }
    .contact { margin-top: 20px; font-size: 13px; color: #999; }
  </style>
</head>
<body>
  <div class="box">
    <div class="icon">Lock</div>
    <h1>Access Denied</h1>
    <p>You do not have permission to access this system.</p>
    <div class="email-box">${email || "Unknown email"}</div>
    <p>If you believe this is an error, please contact the administrator.</p>
    <div class="contact">MKA Voucher System - Restricted Access</div>
  </div>
</body>
</html>`;
}


// ============ ERROR PAGE ============

function getErrorPage(title, message) {
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Error - MKA Voucher System</title>
  <style>
    body { font-family: Arial, sans-serif; background: #f0f2f5; min-height: 100vh; display: flex; align-items: center; justify-content: center; margin: 0; padding: 15px; }
    .box { background: white; padding: 40px; border-radius: 12px; text-align: center; max-width: 400px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); }
    .icon { font-size: 50px; margin-bottom: 20px; }
    h1 { color: #dc3545; margin-bottom: 15px; font-size: 22px; }
    p { color: #666; line-height: 1.6; }
    .home-link { display: inline-block; margin-top: 20px; color: #1a5f2a; text-decoration: none; }
  </style>
</head>
<body>
  <div class="box">
    <div class="icon">Warning</div>
    <h1>${title}</h1>
    <p>${message}</p>
    <a href="${CONFIG.WEBAPP_URL}" class="home-link">Back to Home</a>
  </div>
</body>
</html>`;
}


// ============ ALREADY APPROVED PAGE ============

function getAlreadyApprovedPage(voucher, approverTitle) {
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Already Approved - MKA</title>
  <style>
    body { font-family: Arial, sans-serif; background: #f0f2f5; min-height: 100vh; display: flex; align-items: center; justify-content: center; margin: 0; padding: 15px; }
    .box { background: white; padding: 30px; border-radius: 12px; text-align: center; max-width: 450px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); }
    .icon { font-size: 50px; margin-bottom: 15px; }
    h1 { color: #28a745; margin-bottom: 15px; font-size: 20px; }
    p { color: #666; margin-bottom: 15px; }
    .info { background: #d4edda; padding: 15px; border-radius: 8px; text-align: left; }
    .info p { margin: 5px 0; font-size: 13px; color: #155724; }
  </style>
</head>
<body>
  <div class="box">
    <div class="icon">OK</div>
    <h1>Already Signed</h1>
    <p>This voucher has already been approved by <strong>${approverTitle}</strong>.</p>
    <div class="info">
      <p><strong>Voucher:</strong> ${voucher.id}</p>
      <p><strong>Amount:</strong> ${voucher.amountInNumbers} SEK</p>
      <p><strong>Status:</strong> ${voucher.status.replace('pending_', 'Awaiting ').replace('_', ' ')}</p>
    </div>
    <p style="margin-top:20px;font-size:12px;color:#999;">Each approver can only sign once.</p>
  </div>
</body>
</html>`;
}


// ============ NOT YOUR TURN PAGE ============

function getNotYourTurnPage(voucher, approverTitle) {
  const progress = getApprovalProgress(voucher);
  const currentStep = progress.find(p => p.status === 'current');

  const progressHtml = progress.map(p => {
    let bgColor = '#e9ecef', textColor = '#6c757d', icon = String(p.step);
    if (p.approved) { bgColor = '#28a745'; textColor = '#fff'; icon = 'OK'; }
    else if (p.status === 'current') { bgColor = '#007bff'; textColor = '#fff'; icon = '...'; }

    return '<div style="display:flex;align-items:center;padding:12px;margin:8px 0;background:' +
      (p.approved ? '#d4edda' : (p.status === 'current' ? '#cce5ff' : '#f8f9fa')) +
      ';border-radius:8px;border-left:4px solid ' + bgColor + '">' +
      '<div style="width:32px;height:32px;border-radius:50%;background:' + bgColor + ';color:' + textColor +
      ';display:flex;align-items:center;justify-content:center;font-weight:bold;margin-right:12px;">' + icon + '</div>' +
      '<div><div style="font-weight:600;font-size:14px;">' + p.title + '</div>' +
      '<div style="font-size:12px;color:#666;">' + (p.approved ? 'Approved' : (p.status === 'current' ? 'Waiting for approval...' : 'Pending')) + '</div></div></div>';
  }).join('');

  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Not Your Turn - MKA</title>
  <style>
    body { font-family: Arial, sans-serif; background: #f0f2f5; min-height: 100vh; margin: 0; padding: 15px; }
    .container { max-width: 500px; margin: 0 auto; }
    .header { background: #6c757d; color: white; padding: 20px; text-align: center; border-radius: 12px 12px 0 0; }
    .header h1 { margin: 0; font-size: 20px; }
    .header p { margin: 10px 0 0 0; opacity: 0.9; }
    .content { background: white; padding: 20px; border-radius: 0 0 12px 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); }
    .info-box { background: #e9ecef; padding: 15px; border-radius: 8px; margin-bottom: 20px; }
    .info-box p { margin: 5px 0; }
    h3 { font-size: 14px; margin-bottom: 10px; color: #333; }
    .waiting { margin-top: 20px; padding: 15px; background: #fff3cd; border-radius: 8px; font-size: 13px; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>Not Your Turn Yet</h1>
      <p>You are: ${approverTitle}</p>
    </div>
    <div class="content">
      <div class="info-box">
        <p><strong>Voucher:</strong> ${voucher.id}</p>
        <p><strong>Amount:</strong> ${voucher.amountInNumbers} SEK</p>
        <p><strong>Claimed by:</strong> ${voucher.claimedBy}</p>
      </div>

      <h3>Approval Progress:</h3>
      ${progressHtml}

      <div class="waiting">
        <strong>Currently waiting for:</strong> ${currentStep ? currentStep.title : 'Unknown'}<br><br>
        You will receive a new email when it's your turn to approve.
      </div>
    </div>
  </div>
</body>
</html>`;
}


// ============ ALREADY PAID PAGE ============

function getAlreadyPaidPage(voucher) {
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Already Paid - MKA</title>
  <style>
    body { font-family: Arial, sans-serif; background: #f0f2f5; min-height: 100vh; display: flex; align-items: center; justify-content: center; margin: 0; padding: 15px; }
    .box { background: white; padding: 30px; border-radius: 12px; text-align: center; max-width: 450px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); }
    .icon { font-size: 50px; margin-bottom: 15px; }
    h1 { color: #28a745; margin-bottom: 15px; font-size: 20px; }
    p { color: #666; margin-bottom: 15px; }
    .info { background: #d4edda; padding: 15px; border-radius: 8px; text-align: left; }
    .info p { margin: 5px 0; font-size: 13px; color: #155724; }
  </style>
</head>
<body>
  <div class="box">
    <div class="icon">Paid</div>
    <h1>Already Paid</h1>
    <p>This voucher has already been marked as paid.</p>
    <div class="info">
      <p><strong>Voucher:</strong> ${voucher.id}</p>
      <p><strong>Amount:</strong> ${voucher.amountInNumbers} SEK</p>
      <p><strong>Paid By:</strong> ${voucher.payment.paidBy || 'N/A'}</p>
      <p><strong>Payment Date:</strong> ${voucher.payment.paymentDate || 'N/A'}</p>
      <p><strong>Method:</strong> ${voucher.payment.paymentMethod || 'N/A'}</p>
    </div>
  </div>
</body>
</html>`;
}


// ============ HOME PAGE ============

function getHomePage() {
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>MKA Voucher System</title>
  <style>
    body { font-family: Arial, sans-serif; background: #f0f2f5; min-height: 100vh; display: flex; align-items: center; justify-content: center; margin: 0; }
    .box { background: white; padding: 40px; border-radius: 12px; text-align: center; max-width: 400px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); }
    .logo { font-size: 48px; font-weight: bold; color: #1a5f2a; margin-bottom: 10px; }
    h1 { color: #333; margin-bottom: 10px; font-size: 22px; }
    p { color: #666; }
    .footer { margin-top: 20px; padding-top: 20px; border-top: 1px solid #eee; font-size: 12px; color: #999; }
    .admin-link { display: inline-block; margin-top: 15px; padding: 10px 20px; background: #1a5f2a; color: white; text-decoration: none; border-radius: 6px; font-size: 14px; }
    .admin-link:hover { background: #145222; }
  </style>
</head>
<body>
  <div class="box">
    <div class="logo">MKA</div>
    <h1>Voucher Approval System</h1>
    <p>Majlis Khuddam-ul-Ahmadiyya Sverige</p>
    <p style="margin-top:20px;font-size:14px;">Use the links in your email to approve or reject vouchers.</p>
    <a href="${CONFIG.WEBAPP_URL}?action=admin" class="admin-link">Admin Portal</a>
    <div class="footer">
      <p>Tolvskillingsgatan 1, 414 82 Goteborg</p>
    </div>
  </div>
</body>
</html>`;
}


// ============ SIGNATURE PAGE ============

function getSignaturePage(voucherId, approverKey) {
  try {
    const voucher = getVoucherById(voucherId);

    if (!voucher) {
      return getErrorPage("Voucher Not Found", "This voucher (ID: " + voucherId + ") does not exist or has been deleted. Please check your email for the correct link.");
    }

    const approver = getApproverInfo(approverKey);
    if (!approver) {
      return getErrorPage("Invalid Link", "This approval link is invalid. Please use the link from your email.");
    }

    if (isAlreadyApproved(voucher, approverKey)) {
      return getAlreadyApprovedPage(voucher, approver.title);
    }

    if (!isApproversTurn(voucher, approverKey)) {
      return getNotYourTurnPage(voucher, approver.title);
    }

    return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Approve Voucher - MKA</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: Arial, sans-serif; background: #f0f2f5; min-height: 100vh; }
    .container { max-width: 500px; margin: 0 auto; padding: 15px; }
    .header { background: linear-gradient(135deg, #1a5f2a 0%, #2d8f4e 100%); color: white; padding: 20px; text-align: center; border-radius: 12px 12px 0 0; }
    .header h1 { font-size: 20px; margin-bottom: 5px; }
    .header p { opacity: 0.9; font-size: 14px; }
    .content { background: white; padding: 20px; border-radius: 0 0 12px 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); }
    .info-box { background: #f8f9fa; border-radius: 8px; padding: 15px; margin-bottom: 20px; }
    .info-row { display: flex; justify-content: space-between; padding: 8px 0; border-bottom: 1px solid #eee; }
    .info-row:last-child { border-bottom: none; }
    .info-label { color: #666; font-size: 13px; }
    .info-value { font-weight: 600; font-size: 13px; text-align: right; max-width: 60%; }
    .amount { color: #1a5f2a; font-size: 18px !important; }
    h3 { font-size: 14px; margin-bottom: 10px; color: #333; }
    .signature-area { border: 2px dashed #ccc; border-radius: 10px; background: #fafafa; margin-bottom: 15px; position: relative; }
    #signatureCanvas { width: 100%; height: 150px; cursor: crosshair; touch-action: none; display: block; }
    .btn { display: block; width: 100%; padding: 14px; border: none; border-radius: 8px; font-size: 16px; font-weight: 600; cursor: pointer; margin-bottom: 10px; transition: opacity 0.2s; }
    .btn-primary { background: #1a5f2a; color: white; }
    .btn-secondary { background: #6c757d; color: white; }
    .btn:disabled { opacity: 0.5; cursor: not-allowed; }
    .success { background: #d4edda; color: #155724; padding: 30px; border-radius: 8px; text-align: center; }
    .success h2 { margin-bottom: 10px; font-size: 24px; }
    .loading { display: none; text-align: center; padding: 20px; }
    .spinner { border: 3px solid #f3f3f3; border-top: 3px solid #1a5f2a; border-radius: 50%; width: 30px; height: 30px; animation: spin 1s linear infinite; margin: 0 auto 10px; }
    @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
    .error-msg { background: #f8d7da; color: #721c24; padding: 10px; border-radius: 5px; margin-bottom: 15px; display: none; }
    .attachment-section { margin-bottom: 20px; }
    .attachment-box { background: #e8f5e9; border: 1px solid #c8e6c9; border-radius: 8px; padding: 12px; }
    .attachment-link { display: flex; align-items: center; gap: 10px; color: #1a5f2a; text-decoration: none; font-weight: 600; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>Voucher Approval</h1>
      <p>Signing as ${approver.title}</p>
    </div>

    <div class="content" id="mainContent">
      <div class="error-msg" id="errorMsg"></div>

      <div class="info-box">
        <div class="info-row">
          <span class="info-label">Reference</span>
          <span class="info-value">${voucher.id}</span>
        </div>
        <div class="info-row">
          <span class="info-label">Claimed By</span>
          <span class="info-value">${voucher.claimedBy}</span>
        </div>
        <div class="info-row">
          <span class="info-label">Description</span>
          <span class="info-value">${voucher.descriptionOfProgram}</span>
        </div>
        <div class="info-row">
          <span class="info-label">Category</span>
          <span class="info-value">${voucher.category || 'N/A'}</span>
        </div>
        <div class="info-row">
          <span class="info-label">Majlis</span>
          <span class="info-value">${voucher.majlis || 'N/A'}</span>
        </div>
        <div class="info-row">
          <span class="info-label">Amount</span>
          <span class="info-value amount">${voucher.amountInNumbers} SEK</span>
        </div>
      </div>

      ${voucher.attachmentUrl ? `
      <div class="attachment-section">
        <h3>Attachment / Receipt:</h3>
        <div class="attachment-box">
          <a href="${voucher.attachmentUrl}" target="_blank" class="attachment-link">
            <span>Attachment - View</span>
          </a>
        </div>
      </div>
      ` : ''}

      <h3>Draw Your Signature:</h3>
      <div class="signature-area">
        <canvas id="signatureCanvas"></canvas>
      </div>

      <button class="btn btn-secondary" id="clearBtn" onclick="clearSignature()">Clear Signature</button>
      <button class="btn btn-primary" id="submitBtn" onclick="submitSignature()">APPROVE & SUBMIT</button>

      <div class="loading" id="loading">
        <div class="spinner"></div>
        <p>Processing your approval...</p>
        <p style="font-size:12px;color:#666;margin-top:10px;">Please wait, do not close this page.</p>
      </div>
    </div>
  </div>

  <` + `script>
    var canvas = document.getElementById('signatureCanvas');
    var ctx = canvas.getContext('2d');
    var isDrawing = false;
    var lastX = 0, lastY = 0;
    var isSubmitting = false;

    function resizeCanvas() {
      var rect = canvas.parentElement.getBoundingClientRect();
      canvas.width = rect.width - 4;
      canvas.height = 150;
      ctx.strokeStyle = '#000';
      ctx.lineWidth = 2;
      ctx.lineCap = 'round';
      ctx.lineJoin = 'round';
    }

    resizeCanvas();
    window.addEventListener('resize', resizeCanvas);

    function getPos(e) {
      var rect = canvas.getBoundingClientRect();
      if (e.touches) {
        return { x: e.touches[0].clientX - rect.left, y: e.touches[0].clientY - rect.top };
      }
      return { x: e.offsetX, y: e.offsetY };
    }

    canvas.addEventListener('mousedown', function(e) { isDrawing = true; var pos = getPos(e); lastX = pos.x; lastY = pos.y; });
    canvas.addEventListener('mousemove', function(e) {
      if (!isDrawing) return;
      var pos = getPos(e);
      ctx.beginPath(); ctx.moveTo(lastX, lastY); ctx.lineTo(pos.x, pos.y); ctx.stroke();
      lastX = pos.x; lastY = pos.y;
    });
    canvas.addEventListener('mouseup', function() { isDrawing = false; });
    canvas.addEventListener('mouseout', function() { isDrawing = false; });

    canvas.addEventListener('touchstart', function(e) {
      e.preventDefault(); isDrawing = true;
      var pos = getPos(e); lastX = pos.x; lastY = pos.y;
    }, { passive: false });

    canvas.addEventListener('touchmove', function(e) {
      e.preventDefault();
      if (!isDrawing) return;
      var pos = getPos(e);
      ctx.beginPath(); ctx.moveTo(lastX, lastY); ctx.lineTo(pos.x, pos.y); ctx.stroke();
      lastX = pos.x; lastY = pos.y;
    }, { passive: false });

    canvas.addEventListener('touchend', function() { isDrawing = false; });

    function clearSignature() {
      ctx.clearRect(0, 0, canvas.width, canvas.height);
      document.getElementById('errorMsg').style.display = 'none';
    }

    function showError(msg) {
      var el = document.getElementById('errorMsg');
      el.textContent = msg;
      el.style.display = 'block';
    }

    function submitSignature() {
      if (isSubmitting) return;

      var imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
      var hasSignature = false;
      for (var i = 3; i < imageData.data.length; i += 4) {
        if (imageData.data[i] > 0) { hasSignature = true; break; }
      }

      if (!hasSignature) {
        showError('Please draw your signature before submitting.');
        return;
      }

      isSubmitting = true;
      document.getElementById('submitBtn').disabled = true;
      document.getElementById('clearBtn').disabled = true;
      document.getElementById('loading').style.display = 'block';
      document.getElementById('errorMsg').style.display = 'none';

      var signatureData = canvas.toDataURL('image/png');

      google.script.run
        .withSuccessHandler(function(result) {
          if (result && result.success) {
            var safeMsg = (result.message || 'The voucher will proceed to the next step.').replace(/&/g,'&amp;').replace(new RegExp('<','g'),'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
            document.getElementById('mainContent').innerHTML =
              '<div class="success">' +
                '<h2>Approved!</h2>' +
                '<p>Your signature has been recorded.</p>' +
                '<p style="margin-top:15px;color:#666;font-size:13px;">' + safeMsg + '</p>' +
              '</div>';
          } else {
            isSubmitting = false;
            document.getElementById('loading').style.display = 'none';
            document.getElementById('submitBtn').disabled = false;
            document.getElementById('clearBtn').disabled = false;
            showError(result ? result.message : 'An error occurred. Please try again.');
          }
        })
        .withFailureHandler(function(error) {
          isSubmitting = false;
          document.getElementById('loading').style.display = 'none';
          document.getElementById('submitBtn').disabled = false;
          document.getElementById('clearBtn').disabled = false;
          showError('Connection error. Please check your internet and try again.');
        })
        .processApprovalFromWeb('${voucherId}', '${approverKey}', signatureData);
    }
  <` + `/script>
</body>
</html>`;
  } catch (error) {
    Logger.log("Error in getSignaturePage: " + error);
    return getErrorPage("System Error", "Unable to load the approval page. Please try again later.");
  }
}


// ============ REJECT PAGE ============

function getRejectPage(voucherId, approverKey) {
  try {
    const voucher = getVoucherById(voucherId);
    const approver = getApproverInfo(approverKey);

    if (!voucher) {
      return getErrorPage("Voucher Not Found", "This voucher does not exist.");
    }

    if (!approver) {
      return getErrorPage("Invalid Link", "This rejection link is invalid.");
    }

    return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Reject Voucher - MKA</title>
  <style>
    body { font-family: Arial, sans-serif; background: #f0f2f5; min-height: 100vh; margin: 0; padding: 15px; }
    .container { max-width: 500px; margin: 0 auto; }
    .header { background: #dc3545; color: white; padding: 20px; text-align: center; border-radius: 12px 12px 0 0; }
    .header h1 { margin: 0; font-size: 20px; }
    .content { background: white; padding: 20px; border-radius: 0 0 12px 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); }
    .info { background: #f8f9fa; padding: 15px; border-radius: 8px; margin-bottom: 20px; }
    .info p { margin: 5px 0; }
    label { font-weight: 600; display: block; margin-bottom: 8px; }
    textarea { width: 100%; height: 100px; padding: 12px; border: 1px solid #ddd; border-radius: 8px; font-family: inherit; resize: vertical; box-sizing: border-box; }
    .btn { display: block; width: 100%; padding: 14px; border: none; border-radius: 8px; font-size: 16px; font-weight: 600; cursor: pointer; background: #dc3545; color: white; margin-top: 15px; }
    .btn:disabled { opacity: 0.5; cursor: not-allowed; }
    .success { text-align: center; padding: 30px; }
    .success h2 { color: #dc3545; margin-bottom: 10px; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>Reject Voucher</h1>
    </div>
    <div class="content" id="mainContent">
      <div class="info">
        <p><strong>Voucher:</strong> ${voucher.id}</p>
        <p><strong>Amount:</strong> ${voucher.amountInNumbers} SEK</p>
        <p><strong>Claimed by:</strong> ${voucher.claimedBy}</p>
      </div>

      <label>Reason for rejection:</label>
      <textarea id="reason" placeholder="Please explain why you are rejecting this voucher..."></textarea>

      <button class="btn" id="rejectBtn" onclick="submitRejection()">Submit Rejection</button>
    </div>
  </div>

  <` + `script>
    var isSubmitting = false;

    function submitRejection() {
      if (isSubmitting) return;

      var reason = document.getElementById('reason').value.trim();
      if (!reason) {
        alert('Please provide a reason for rejection.');
        return;
      }

      isSubmitting = true;
      document.getElementById('rejectBtn').disabled = true;
      document.getElementById('rejectBtn').textContent = 'Processing...';

      google.script.run
        .withSuccessHandler(function(result) {
          if (result && result.success) {
            document.getElementById('mainContent').innerHTML =
              '<div class="success"><h2>Voucher Rejected</h2><p>The claimer has been notified of the rejection.</p></div>';
          } else {
            isSubmitting = false;
            document.getElementById('rejectBtn').disabled = false;
            document.getElementById('rejectBtn').textContent = 'Submit Rejection';
            alert(result ? result.message : 'An error occurred.');
          }
        })
        .withFailureHandler(function(error) {
          isSubmitting = false;
          document.getElementById('rejectBtn').disabled = false;
          document.getElementById('rejectBtn').textContent = 'Submit Rejection';
          alert('Error: ' + error);
        })
        .rejectVoucher('${voucherId}', '${approverKey}', reason);
    }
  <` + `/script>
</body>
</html>`;
  } catch (error) {
    Logger.log("Error in getRejectPage: " + error);
    return getErrorPage("System Error", "Unable to load the rejection page.");
  }
}


// ============ MUHTAMIM MAAL PAYMENT PAGE ============

function getMuhtamimMaalPaymentPage(voucherId) {
  try {
    const voucher = getVoucherById(voucherId);

    if (!voucher) {
      return getErrorPage("Voucher Not Found", "This voucher (ID: " + voucherId + ") does not exist.");
    }

    if (voucher.status !== "approved") {
      return getErrorPage("Not Ready for Payment", "This voucher is not yet fully approved. Current status: " + voucher.status);
    }

    if (voucher.payment && voucher.payment.paid) {
      return getAlreadyPaidPage(voucher);
    }

    return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Process Payment - MKA</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: Arial, sans-serif; background: #f0f2f5; min-height: 100vh; }
    .container { max-width: 500px; margin: 0 auto; padding: 15px; }
    .header { background: linear-gradient(135deg, #1a5f2a 0%, #2d8f4e 100%); color: white; padding: 20px; text-align: center; border-radius: 12px 12px 0 0; }
    .header h1 { font-size: 20px; margin-bottom: 5px; }
    .header p { opacity: 0.9; font-size: 14px; }
    .content { background: white; padding: 20px; border-radius: 0 0 12px 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); }
    .info-box { background: #f8f9fa; border-radius: 8px; padding: 15px; margin-bottom: 20px; }
    .info-row { display: flex; justify-content: space-between; padding: 8px 0; border-bottom: 1px solid #eee; }
    .info-row:last-child { border-bottom: none; }
    .info-label { color: #666; font-size: 13px; }
    .info-value { font-weight: 600; font-size: 13px; text-align: right; max-width: 60%; }
    .amount { color: #1a5f2a; font-size: 18px !important; }
    h3 { font-size: 14px; margin-bottom: 10px; color: #333; }
    .form-group { margin-bottom: 15px; }
    label { display: block; font-weight: 600; margin-bottom: 5px; font-size: 13px; }
    input, select { width: 100%; padding: 12px; border: 1px solid #ddd; border-radius: 8px; font-size: 14px; }
    input:focus, select:focus { outline: none; border-color: #1a5f2a; }
    .signature-area { border: 2px dashed #ccc; border-radius: 10px; background: #fafafa; margin-bottom: 15px; }
    #signatureCanvas { width: 100%; height: 120px; cursor: crosshair; touch-action: none; display: block; }
    .btn { display: block; width: 100%; padding: 14px; border: none; border-radius: 8px; font-size: 16px; font-weight: 600; cursor: pointer; margin-bottom: 10px; }
    .btn-primary { background: #1a5f2a; color: white; }
    .btn-secondary { background: #6c757d; color: white; }
    .btn:disabled { opacity: 0.5; cursor: not-allowed; }
    .success { background: #d4edda; color: #155724; padding: 30px; border-radius: 8px; text-align: center; }
    .success h2 { margin-bottom: 10px; font-size: 24px; }
    .loading { display: none; text-align: center; padding: 20px; }
    .spinner { border: 3px solid #f3f3f3; border-top: 3px solid #1a5f2a; border-radius: 50%; width: 30px; height: 30px; animation: spin 1s linear infinite; margin: 0 auto 10px; }
    @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
    .error-msg { background: #f8d7da; color: #721c24; padding: 10px; border-radius: 5px; margin-bottom: 15px; display: none; }
    .section-title { background: #e9ecef; padding: 10px; border-radius: 6px; margin: 20px 0 15px 0; font-weight: 600; font-size: 14px; }
    .attachment-section { margin-bottom: 20px; }
    .attachment-box { background: #e8f5e9; border: 1px solid #c8e6c9; border-radius: 8px; padding: 12px; }
    .attachment-link { display: flex; align-items: center; gap: 10px; color: #1a5f2a; text-decoration: none; font-weight: 600; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>Process Payment</h1>
      <p>Muhtamim Maal - Register Payment</p>
    </div>

    <div class="content" id="mainContent">
      <div class="error-msg" id="errorMsg"></div>

      <div class="info-box">
        <div class="info-row">
          <span class="info-label">Reference</span>
          <span class="info-value">${voucher.id}</span>
        </div>
        <div class="info-row">
          <span class="info-label">Pay To</span>
          <span class="info-value">${voucher.claimedBy}</span>
        </div>
        <div class="info-row">
          <span class="info-label">Description</span>
          <span class="info-value">${voucher.descriptionOfProgram}</span>
        </div>
        <div class="info-row">
          <span class="info-label">Category</span>
          <span class="info-value">${voucher.category || 'N/A'}</span>
        </div>
        <div class="info-row">
          <span class="info-label">Majlis</span>
          <span class="info-value">${voucher.majlis || 'N/A'}</span>
        </div>
        <div class="info-row">
          <span class="info-label">Amount</span>
          <span class="info-value amount">${voucher.amountInNumbers} SEK</span>
        </div>
      </div>

      ${voucher.attachmentUrl ? `
      <div class="attachment-section">
        <h3>Attachment / Receipt:</h3>
        <div class="attachment-box">
          <a href="${voucher.attachmentUrl}" target="_blank" class="attachment-link">
            <span>View Attachment</span>
          </a>
        </div>
      </div>
      ` : ''}

      <div class="section-title">Payment Information</div>

      <div class="form-group">
        <label>Payment Method:</label>
        <select id="paymentMethod">
          <option value="">Select payment method...</option>
          <option value="Bank Transfer">Bank Transfer</option>
          <option value="Swish">Swish</option>
          <option value="Cash">Cash</option>
        </select>
      </div>

      <div class="form-group">
        <label>Account Number / Clearing Number (optional):</label>
        <input type="text" id="accountInfo" placeholder="e.g. 1234-5678901234">
      </div>

      <div class="section-title">Muhtamim Maal Signature</div>
      <div class="signature-area">
        <canvas id="signatureCanvas"></canvas>
      </div>

      <button class="btn btn-secondary" onclick="clearSignature()">Clear Signature</button>
      <button class="btn btn-primary" id="submitBtn" onclick="submitPayment()">CONFIRM PAYMENT</button>

      <div class="loading" id="loading">
        <div class="spinner"></div>
        <p>Processing payment...</p>
      </div>
    </div>
  </div>

  <` + `script>
    var canvas = document.getElementById('signatureCanvas');
    var ctx = canvas.getContext('2d');
    var isDrawing = false;
    var lastX = 0, lastY = 0;
    var isSubmitting = false;

    function resizeCanvas() {
      var rect = canvas.parentElement.getBoundingClientRect();
      canvas.width = rect.width - 4;
      canvas.height = 120;
      ctx.strokeStyle = '#000';
      ctx.lineWidth = 2;
      ctx.lineCap = 'round';
    }

    resizeCanvas();
    window.addEventListener('resize', resizeCanvas);

    function getPos(e) {
      var rect = canvas.getBoundingClientRect();
      if (e.touches) return { x: e.touches[0].clientX - rect.left, y: e.touches[0].clientY - rect.top };
      return { x: e.offsetX, y: e.offsetY };
    }

    canvas.addEventListener('mousedown', function(e) { isDrawing = true; var pos = getPos(e); lastX = pos.x; lastY = pos.y; });
    canvas.addEventListener('mousemove', function(e) {
      if (!isDrawing) return;
      var pos = getPos(e);
      ctx.beginPath(); ctx.moveTo(lastX, lastY); ctx.lineTo(pos.x, pos.y); ctx.stroke();
      lastX = pos.x; lastY = pos.y;
    });
    canvas.addEventListener('mouseup', function() { isDrawing = false; });
    canvas.addEventListener('mouseout', function() { isDrawing = false; });

    canvas.addEventListener('touchstart', function(e) { e.preventDefault(); isDrawing = true; var pos = getPos(e); lastX = pos.x; lastY = pos.y; }, { passive: false });
    canvas.addEventListener('touchmove', function(e) {
      e.preventDefault();
      if (!isDrawing) return;
      var pos = getPos(e);
      ctx.beginPath(); ctx.moveTo(lastX, lastY); ctx.lineTo(pos.x, pos.y); ctx.stroke();
      lastX = pos.x; lastY = pos.y;
    }, { passive: false });
    canvas.addEventListener('touchend', function() { isDrawing = false; });

    function clearSignature() {
      ctx.clearRect(0, 0, canvas.width, canvas.height);
      document.getElementById('errorMsg').style.display = 'none';
    }

    function showError(msg) {
      var el = document.getElementById('errorMsg');
      el.textContent = msg;
      el.style.display = 'block';
    }

    function submitPayment() {
      if (isSubmitting) return;

      var paymentMethod = document.getElementById('paymentMethod').value;
      var accountInfo = document.getElementById('accountInfo').value.trim();

      if (!paymentMethod) { showError('Please select a payment method.'); return; }

      var imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
      var hasSignature = false;
      for (var i = 3; i < imageData.data.length; i += 4) {
        if (imageData.data[i] > 0) { hasSignature = true; break; }
      }

      if (!hasSignature) { showError('Please draw your signature.'); return; }

      isSubmitting = true;
      document.getElementById('submitBtn').disabled = true;
      document.getElementById('loading').style.display = 'block';
      document.getElementById('errorMsg').style.display = 'none';

      var signatureData = canvas.toDataURL('image/png');

      google.script.run
        .withSuccessHandler(function(result) {
          if (result && result.success) {
            document.getElementById('mainContent').innerHTML =
              '<div class="success">' +
                '<h2>Payment Registered!</h2>' +
                '<p>Payment has been recorded and PDF is being generated.</p>' +
                '<p style="margin-top:15px;color:#666;font-size:13px;">The final voucher is now in the Voucher folder.</p>' +
              '</div>';
          } else {
            isSubmitting = false;
            document.getElementById('loading').style.display = 'none';
            document.getElementById('submitBtn').disabled = false;
            showError(result ? result.message : 'An error occurred.');
          }
        })
        .withFailureHandler(function(error) {
          isSubmitting = false;
          document.getElementById('loading').style.display = 'none';
          document.getElementById('submitBtn').disabled = false;
          showError('Connection error. Please try again.');
        })
        .processPaymentFromWeb('${voucherId}', paymentMethod, accountInfo, signatureData);
    }
  <` + `/script>
</body>
</html>`;
  } catch (error) {
    Logger.log("Error in getMuhtamimMaalPaymentPage: " + error);
    return getErrorPage("System Error", "Unable to load the payment page.");
  }
}
