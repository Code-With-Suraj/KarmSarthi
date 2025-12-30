/**
 * Professional Email Template with Branding
 * Creates modern HTML emails for leave notifications
 */

/**
 * Generate HTML email template with KarmSarthi branding
 */
function getEmailTemplate(title, content, actionButton = null) {
  const buttonHtml = actionButton ? `
    <table role="presentation" cellspacing="0" cellpadding="0" border="0" align="center" style="margin: 30px auto;">
      <tr>
        <td style="border-radius: 8px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);">
          <a href="${actionButton.url}" target="_blank" style="border: none; color: #ffffff; padding: 14px 32px; text-decoration: none; display: inline-block; font-size: 16px; font-weight: 600; border-radius: 8px;">
            ${actionButton.text}
          </a>
        </td>
      </tr>
    </table>
  ` : '';

  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${title}</title>
</head>
<body style="margin: 0; padding: 0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f7fa;">
  <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="background-color: #f4f7fa; padding: 40px 20px;">
    <tr>
      <td align="center">
        <!-- Main Container -->
        <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="600" style="max-width: 600px; background-color: #ffffff; border-radius: 16px; box-shadow: 0 4px 20px rgba(0,0,0,0.1); overflow: hidden;">
          
          <!-- Header with Gradient -->
          <tr>
            <td style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 40px 30px; text-align: center;">
              <h1 style="margin: 0; color: #ffffff; font-size: 28px; font-weight: 700; letter-spacing: -0.5px;">
                KarmSarthi
              </h1>
              <p style="margin: 8px 0 0 0; color: rgba(255,255,255,0.9); font-size: 14px; font-weight: 500;">
                Har din, har chhutti ka bharosa
              </p>
            </td>
          </tr>

          <!-- Title Bar -->
          <tr>
            <td style="background-color: #f8f9fc; padding: 24px 30px; border-bottom: 3px solid #667eea;">
              <h2 style="margin: 0; color: #1a202c; font-size: 22px; font-weight: 600;">
                ${title}
              </h2>
            </td>
          </tr>

          <!-- Content -->
          <tr>
            <td style="padding: 40px 30px; color: #4a5568; font-size: 15px; line-height: 1.8;">
              ${content}
            </td>
          </tr>

          <!-- Action Button (if provided) -->
          ${actionButton ? `
          <tr>
            <td style="padding: 0 30px 40px 30px;">
              ${buttonHtml}
            </td>
          </tr>
          ` : ''}

          <!-- Footer -->
          <tr>
            <td style="background-color: #f8f9fc; padding: 30px; text-align: center; border-top: 1px solid #e2e8f0;">
              <p style="margin: 0 0 12px 0; color: #718096; font-size: 13px;">
                This is an automated notification from KarmSarthi Leave Management System
              </p>
              <p style="margin: 0; color: #a0aec0; font-size: 12px;">
                © ${new Date().getFullYear()} KarmSarthi. All rights reserved.
              </p>
            </td>
          </tr>

        </table>
      </td>
    </tr>
  </table>
</body>
</html>
  `;
}

/**
 * Send email notification with professional template
 */
function sendEmailNotification(to, subject, plainText, htmlContent = null) {
  try {
    if (!htmlContent) {
      // If no HTML provided, wrap plain text in template
      htmlContent = getEmailTemplate(subject, `<p>${plainText.replace(/\n/g, '<br>')}</p>`);
    }
    
    MailApp.sendEmail({
      to: to,
      subject: `[KarmSarthi] ${subject}`,
      body: plainText,
      htmlBody: htmlContent
    });
  } catch (error) {
    Logger.log('Error sending email: ' + error.toString());
  }
}

/**
 * Send leave submission notification to manager
 */
function sendLeaveSubmissionEmail(managerEmail, employeeName, employeeEmail, leaveType, startDate, endDate, units, dayType, reason, requestId) {
  const content = `
    <p style="margin: 0 0 20px 0; font-size: 16px; color: #2d3748;">
      Dear Manager,
    </p>
    <p style="margin: 0 0 24px 0;">
      <strong>${employeeName}</strong> has submitted a new leave request that requires your approval.
    </p>
    
    <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="background-color: #f7fafc; border-radius: 8px; padding: 20px; margin: 24px 0;">
      <tr>
        <td>
          <table role="presentation" cellspacing="0" cellpadding="8" border="0" width="100%">
            <tr>
              <td style="color: #718096; font-size: 14px; width: 40%;">Request ID:</td>
              <td style="color: #2d3748; font-weight: 600; font-size: 14px;">${requestId}</td>
            </tr>
            <tr>
              <td style="color: #718096; font-size: 14px;">Employee:</td>
              <td style="color: #2d3748; font-weight: 600; font-size: 14px;">${employeeName}</td>
            </tr>
            <tr>
              <td style="color: #718096; font-size: 14px;">Email:</td>
              <td style="color: #2d3748; font-size: 14px;">${employeeEmail}</td>
            </tr>
            <tr>
              <td style="color: #718096; font-size: 14px;">Leave Type:</td>
              <td style="color: #2d3748; font-weight: 600; font-size: 14px;">${leaveType}</td>
            </tr>
            <tr>
              <td style="color: #718096; font-size: 14px;">Duration:</td>
              <td style="color: #2d3748; font-weight: 600; font-size: 14px;">${units} ${units === 1 ? 'day' : 'days'} (${dayType})</td>
            </tr>
            <tr>
              <td style="color: #718096; font-size: 14px;">From:</td>
              <td style="color: #2d3748; font-size: 14px;">${formatDate(startDate)}</td>
            </tr>
            <tr>
              <td style="color: #718096; font-size: 14px;">To:</td>
              <td style="color: #2d3748; font-size: 14px;">${formatDate(endDate)}</td>
            </tr>
            ${reason ? `
            <tr>
              <td style="color: #718096; font-size: 14px; vertical-align: top;">Reason:</td>
              <td style="color: #2d3748; font-size: 14px; font-style: italic;">"${reason}"</td>
            </tr>
            ` : ''}
          </table>
        </td>
      </tr>
    </table>

    <p style="margin: 24px 0 0 0; color: #4a5568;">
      Please review and take action on this request at your earliest convenience.
    </p>
  `;

  const htmlContent = getEmailTemplate(
    'New Leave Request - Pending Approval',
    content,
    {
      text: 'Review Request',
      url: ScriptApp.getService().getUrl()
    }
  );

  const plainText = `New Leave Request - Pending Approval\n\n` +
    `${employeeName} has submitted a new leave request.\n\n` +
    `Request ID: ${requestId}\n` +
    `Leave Type: ${leaveType}\n` +
    `Duration: ${units} ${units === 1 ? 'day' : 'days'} (${dayType})\n` +
    `From: ${formatDate(startDate)}\n` +
    `To: ${formatDate(endDate)}\n` +
    (reason ? `Reason: ${reason}\n` : '') +
    `\nPlease review this request.`;

  sendEmailNotification(managerEmail, 'New Leave Request - Pending Approval', plainText, htmlContent);
}

/**
 * Send leave approval notification to employee
 */
function sendLeaveApprovalEmail(employeeEmail, employeeName, leaveType, startDate, endDate, units, dayType, requestId, managerComments) {
  const content = `
    <div style="text-align: center; margin: 0 0 30px 0;">
      <div style="display: inline-block; background: linear-gradient(135deg, #10b981 0%, #059669 100%); border-radius: 50%; width: 80px; height: 80px; line-height: 80px;">
        <span style="font-size: 40px;">✓</span>
      </div>
    </div>

    <p style="margin: 0 0 20px 0; font-size: 16px; color: #2d3748;">
      Dear <strong>${employeeName}</strong>,
    </p>
    <p style="margin: 0 0 24px 0; font-size: 16px; color: #10b981; font-weight: 600;">
      Great news! Your leave request has been approved.
    </p>
    
    <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="background-color: #f0fdf4; border-left: 4px solid #10b981; border-radius: 8px; padding: 20px; margin: 24px 0;">
      <tr>
        <td>
          <table role="presentation" cellspacing="0" cellpadding="8" border="0" width="100%">
            <tr>
              <td style="color: #065f46; font-size: 14px; width: 40%;">Request ID:</td>
              <td style="color: #064e3b; font-weight: 600; font-size: 14px;">${requestId}</td>
            </tr>
            <tr>
              <td style="color: #065f46; font-size: 14px;">Leave Type:</td>
              <td style="color: #064e3b; font-weight: 600; font-size: 14px;">${leaveType}</td>
            </tr>
            <tr>
              <td style="color: #065f46; font-size: 14px;">Duration:</td>
              <td style="color: #064e3b; font-weight: 600; font-size: 14px;">${units} ${units === 1 ? 'day' : 'days'} (${dayType})</td>
            </tr>
            <tr>
              <td style="color: #065f46; font-size: 14px;">From:</td>
              <td style="color: #064e3b; font-size: 14px;">${formatDate(startDate)}</td>
            </tr>
            <tr>
              <td style="color: #065f46; font-size: 14px;">To:</td>
              <td style="color: #064e3b; font-size: 14px;">${formatDate(endDate)}</td>
            </tr>
            ${managerComments && managerComments !== 'Approved' ? `
            <tr>
              <td style="color: #065f46; font-size: 14px; vertical-align: top;">Manager's Note:</td>
              <td style="color: #064e3b; font-size: 14px; font-style: italic;">"${managerComments}"</td>
            </tr>
            ` : ''}
          </table>
        </td>
      </tr>
    </table>

    <p style="margin: 24px 0 0 0; color: #4a5568;">
      Enjoy your time off! We hope you have a wonderful break.
    </p>
  `;

  const htmlContent = getEmailTemplate('Leave Request Approved ✓', content);

  const plainText = `Leave Request Approved\n\n` +
    `Dear ${employeeName},\n\n` +
    `Your leave request has been approved!\n\n` +
    `Request ID: ${requestId}\n` +
    `Leave Type: ${leaveType}\n` +
    `Duration: ${units} ${units === 1 ? 'day' : 'days'} (${dayType})\n` +
    `From: ${formatDate(startDate)}\n` +
    `To: ${formatDate(endDate)}\n` +
    (managerComments && managerComments !== 'Approved' ? `Manager's Note: ${managerComments}\n` : '') +
    `\nEnjoy your time off!`;

  sendEmailNotification(employeeEmail, 'Leave Request Approved ✓', plainText, htmlContent);
}

/**
 * Send leave rejection notification to employee
 */
function sendLeaveRejectionEmail(employeeEmail, employeeName, leaveType, startDate, endDate, units, dayType, requestId, managerComments) {
  const content = `
    <div style="text-align: center; margin: 0 0 30px 0;">
      <div style="display: inline-block; background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%); border-radius: 50%; width: 80px; height: 80px; line-height: 80px;">
        <span style="font-size: 40px; color: #ffffff;">✕</span>
      </div>
    </div>

    <p style="margin: 0 0 20px 0; font-size: 16px; color: #2d3748;">
      Dear <strong>${employeeName}</strong>,
    </p>
    <p style="margin: 0 0 24px 0;">
      We regret to inform you that your leave request has been declined.
    </p>
    
    <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="background-color: #fef2f2; border-left: 4px solid #ef4444; border-radius: 8px; padding: 20px; margin: 24px 0;">
      <tr>
        <td>
          <table role="presentation" cellspacing="0" cellpadding="8" border="0" width="100%">
            <tr>
              <td style="color: #991b1b; font-size: 14px; width: 40%;">Request ID:</td>
              <td style="color: #7f1d1d; font-weight: 600; font-size: 14px;">${requestId}</td>
            </tr>
            <tr>
              <td style="color: #991b1b; font-size: 14px;">Leave Type:</td>
              <td style="color: #7f1d1d; font-weight: 600; font-size: 14px;">${leaveType}</td>
            </tr>
            <tr>
              <td style="color: #991b1b; font-size: 14px;">Duration:</td>
              <td style="color: #7f1d1d; font-weight: 600; font-size: 14px;">${units} ${units === 1 ? 'day' : 'days'} (${dayType})</td>
            </tr>
            <tr>
              <td style="color: #991b1b; font-size: 14px;">From:</td>
              <td style="color: #7f1d1d; font-size: 14px;">${formatDate(startDate)}</td>
            </tr>
            <tr>
              <td style="color: #991b1b; font-size: 14px;">To:</td>
              <td style="color: #7f1d1d; font-size: 14px;">${formatDate(endDate)}</td>
            </tr>
            ${managerComments ? `
            <tr>
              <td style="color: #991b1b; font-size: 14px; vertical-align: top;">Reason for Rejection:</td>
              <td style="color: #7f1d1d; font-size: 14px; font-weight: 600;">"${managerComments}"</td>
            </tr>
            ` : ''}
          </table>
        </td>
      </tr>
    </table>

    <p style="margin: 24px 0 0 0; color: #4a5568;">
      If you have any questions or would like to discuss this decision, please contact your manager.
    </p>
  `;

  const htmlContent = getEmailTemplate('Leave Request Declined', content);

  const plainText = `Leave Request Declined\n\n` +
    `Dear ${employeeName},\n\n` +
    `Your leave request has been declined.\n\n` +
    `Request ID: ${requestId}\n` +
    `Leave Type: ${leaveType}\n` +
    `Duration: ${units} ${units === 1 ? 'day' : 'days'} (${dayType})\n` +
    `From: ${formatDate(startDate)}\n` +
    `To: ${formatDate(endDate)}\n` +
    (managerComments ? `Reason: ${managerComments}\n` : '') +
    `\nPlease contact your manager if you have questions.`;

  sendEmailNotification(employeeEmail, 'Leave Request Declined', plainText, htmlContent);
}

/**
 * Send leave cancellation notification
 */
function sendLeaveCancellationEmail(managerEmail, employeeName, employeeEmail, leaveType, startDate, endDate, units, dayType, requestId, reason) {
  const content = `
    <p style="margin: 0 0 20px 0; font-size: 16px; color: #2d3748;">
      Dear Manager,
    </p>
    <p style="margin: 0 0 24px 0;">
      <strong>${employeeName}</strong> has cancelled their leave request.
    </p>
    
    <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="background-color: #fffbeb; border-left: 4px solid #f59e0b; border-radius: 8px; padding: 20px; margin: 24px 0;">
      <tr>
        <td>
          <table role="presentation" cellspacing="0" cellpadding="8" border="0" width="100%">
            <tr>
              <td style="color: #92400e; font-size: 14px; width: 40%;">Request ID:</td>
              <td style="color: #78350f; font-weight: 600; font-size: 14px;">${requestId}</td>
            </tr>
            <tr>
              <td style="color: #92400e; font-size: 14px;">Employee:</td>
              <td style="color: #78350f; font-weight: 600; font-size: 14px;">${employeeName}</td>
            </tr>
            <tr>
              <td style="color: #92400e; font-size: 14px;">Leave Type:</td>
              <td style="color: #78350f; font-weight: 600; font-size: 14px;">${leaveType}</td>
            </tr>
            <tr>
              <td style="color: #92400e; font-size: 14px;">Duration:</td>
              <td style="color: #78350f; font-weight: 600; font-size: 14px;">${units} ${units === 1 ? 'day' : 'days'} (${dayType})</td>
            </tr>
            <tr>
              <td style="color: #92400e; font-size: 14px;">From:</td>
              <td style="color: #78350f; font-size: 14px;">${formatDate(startDate)}</td>
            </tr>
            <tr>
              <td style="color: #92400e; font-size: 14px;">To:</td>
              <td style="color: #78350f; font-size: 14px;">${formatDate(endDate)}</td>
            </tr>
            ${reason ? `
            <tr>
              <td style="color: #92400e; font-size: 14px; vertical-align: top;">Cancellation Reason:</td>
              <td style="color: #78350f; font-size: 14px; font-style: italic;">"${reason}"</td>
            </tr>
            ` : ''}
          </table>
        </td>
      </tr>
    </table>

    <p style="margin: 24px 0 0 0; color: #4a5568;">
      This is for your information. No action is required.
    </p>
  `;

  const htmlContent = getEmailTemplate('Leave Request Cancelled', content);

  const plainText = `Leave Request Cancelled\n\n` +
    `${employeeName} has cancelled their leave request.\n\n` +
    `Request ID: ${requestId}\n` +
    `Leave Type: ${leaveType}\n` +
    `Duration: ${units} ${units === 1 ? 'day' : 'days'} (${dayType})\n` +
    `From: ${formatDate(startDate)}\n` +
    `To: ${formatDate(endDate)}\n` +
    (reason ? `Reason: ${reason}\n` : '');

  sendEmailNotification(managerEmail, 'Leave Request Cancelled', plainText, htmlContent);
}

// ============================================================================
// REGULARIZATION REQUEST EMAIL TEMPLATES
// ============================================================================

/**
 * Send regularization request notification to manager
 */
function sendRegularizationRequestEmail(managerEmail, employeeName, attendanceDate, reason) {
  const content = `
    <p style="margin: 0 0 20px 0; font-size: 16px; color: #2d3748;">
      Dear Manager,
    </p>
    <p style="margin: 0 0 24px 0;">
      <strong>${employeeName}</strong> has submitted an attendance regularization request that requires your approval.
    </p>
    
    <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="background-color: #fef3c7; border-left: 4px solid #f59e0b; border-radius: 8px; padding: 20px; margin: 24px 0;">
      <tr>
        <td>
          <table role="presentation" cellspacing="0" cellpadding="8" border="0" width="100%">
            <tr>
              <td style="color: #92400e; font-size: 14px; width: 40%;">Employee:</td>
              <td style="color: #78350f; font-weight: 600; font-size: 14px;">${employeeName}</td>
            </tr>
            <tr>
              <td style="color: #92400e; font-size: 14px;">Attendance Date:</td>
              <td style="color: #78350f; font-weight: 600; font-size: 14px;">${attendanceDate}</td>
            </tr>
            <tr>
              <td style="color: #92400e; font-size: 14px; vertical-align: top;">Reason:</td>
              <td style="color: #78350f; font-size: 14px; font-style: italic;">"${reason}"</td>
            </tr>
          </table>
        </td>
      </tr>
    </table>

    <p style="margin: 24px 0 0 0; color: #4a5568;">
      Please review and take action on this regularization request at your earliest convenience.
    </p>
  `;

  const htmlContent = getEmailTemplate(
    'Attendance Regularization Request - Pending Approval',
    content,
    {
      text: 'Review Request',
      url: ScriptApp.getService().getUrl()
    }
  );

  const plainText = `Attendance Regularization Request - Pending Approval\n\n` +
    `${employeeName} has submitted an attendance regularization request.\n\n` +
    `Attendance Date: ${attendanceDate}\n` +
    `Reason: ${reason}\n` +
    `\nPlease review this request.`;

  sendEmailNotification(managerEmail, 'Attendance Regularization Request - Pending Approval', plainText, htmlContent);
}

/**
 * Send regularization approval notification to employee
 */
function sendRegularizationApprovedEmail(employeeEmail, employeeName, attendanceDate, managerComments) {
  const content = `
    <div style="text-align: center; margin: 0 0 30px 0;">
      <div style="display: inline-block; background: linear-gradient(135deg, #10b981 0%, #059669 100%); border-radius: 50%; width: 80px; height: 80px; line-height: 80px;">
        <span style="font-size: 40px;">✓</span>
      </div>
    </div>

    <p style="margin: 0 0 20px 0; font-size: 16px; color: #2d3748;">
      Dear <strong>${employeeName}</strong>,
    </p>
    <p style="margin: 0 0 24px 0; font-size: 16px; color: #10b981; font-weight: 600;">
      Great news! Your attendance regularization request has been approved.
    </p>
    
    <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="background-color: #f0fdf4; border-left: 4px solid #10b981; border-radius: 8px; padding: 20px; margin: 24px 0;">
      <tr>
        <td>
          <table role="presentation" cellspacing="0" cellpadding="8" border="0" width="100%">
            <tr>
              <td style="color: #065f46; font-size: 14px; width: 40%;">Attendance Date:</td>
              <td style="color: #064e3b; font-weight: 600; font-size: 14px;">${attendanceDate}</td>
            </tr>
            <tr>
              <td style="color: #065f46; font-size: 14px;">Status:</td>
              <td style="color: #064e3b; font-weight: 600; font-size: 14px;">Regularized</td>
            </tr>
            ${managerComments && managerComments !== 'Approved' ? `
            <tr>
              <td style="color: #065f46; font-size: 14px; vertical-align: top;">Manager's Note:</td>
              <td style="color: #064e3b; font-size: 14px; font-style: italic;">"${managerComments}"</td>
            </tr>
            ` : ''}
          </table>
        </td>
      </tr>
    </table>

    <p style="margin: 24px 0 0 0; color: #4a5568;">
      Your attendance for ${attendanceDate} has been marked as "Regularized" in the system.
    </p>
  `;

  const htmlContent = getEmailTemplate('Regularization Request Approved ✓', content);

  const plainText = `Regularization Request Approved\n\n` +
    `Dear ${employeeName},\n\n` +
    `Your attendance regularization request has been approved!\n\n` +
    `Attendance Date: ${attendanceDate}\n` +
    `Status: Regularized\n` +
    (managerComments && managerComments !== 'Approved' ? `Manager's Note: ${managerComments}\n` : '') +
    `\nYour attendance has been updated in the system.`;

  sendEmailNotification(employeeEmail, 'Regularization Request Approved ✓', plainText, htmlContent);
}

/**
 * Send regularization rejection notification to employee
 */
function sendRegularizationRejectedEmail(employeeEmail, employeeName, attendanceDate, managerComments) {
  const content = `
    <div style="text-align: center; margin: 0 0 30px 0;">
      <div style="display: inline-block; background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%); border-radius: 50%; width: 80px; height: 80px; line-height: 80px;">
        <span style="font-size: 40px; color: #ffffff;">✕</span>
      </div>
    </div>

    <p style="margin: 0 0 20px 0; font-size: 16px; color: #2d3748;">
      Dear <strong>${employeeName}</strong>,
    </p>
    <p style="margin: 0 0 24px 0;">
      We regret to inform you that your attendance regularization request has been declined.
    </p>
    
    <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="background-color: #fef2f2; border-left: 4px solid #ef4444; border-radius: 8px; padding: 20px; margin: 24px 0;">
      <tr>
        <td>
          <table role="presentation" cellspacing="0" cellpadding="8" border="0" width="100%">
            <tr>
              <td style="color: #991b1b; font-size: 14px; width: 40%;">Attendance Date:</td>
              <td style="color: #7f1d1d; font-weight: 600; font-size: 14px;">${attendanceDate}</td>
            </tr>
            ${managerComments ? `
            <tr>
              <td style="color: #991b1b; font-size: 14px; vertical-align: top;">Reason for Rejection:</td>
              <td style="color: #7f1d1d; font-size: 14px; font-weight: 600;">"${managerComments}"</td>
            </tr>
            ` : ''}
          </table>
        </td>
      </tr>
    </table>

    <p style="margin: 24px 0 0 0; color: #4a5568;">
      If you have any questions or would like to discuss this decision, please contact your manager.
    </p>
  `;

  const htmlContent = getEmailTemplate('Regularization Request Declined', content);

  const plainText = `Regularization Request Declined\n\n` +
    `Dear ${employeeName},\n\n` +
    `Your attendance regularization request has been declined.\n\n` +
    `Attendance Date: ${attendanceDate}\n` +
    (managerComments ? `Reason: ${managerComments}\n` : '') +
    `\nPlease contact your manager if you have questions.`;

  sendEmailNotification(employeeEmail, 'Regularization Request Declined', plainText, htmlContent);
}
