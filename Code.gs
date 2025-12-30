/**
 * Leave Management System - Backend Code
 * Handles all server-side operations for leave management
 */

// Configuration
const CONFIG = {
  MONTHLY_LEAVE_ACCRUAL: 2,
  MAX_REGULARIZATION_DAYS: 7, // Maximum days back for regularization requests
  SHEETS: {
    EMPLOYEES: 'Employees',
    EMPLOYEE_DETAILS: 'Employee Details',
    LEAVE_REQUESTS: 'Leave Requests',
    LEAVE_HISTORY: 'Leave History',
    COMPANY_HOLIDAYS: 'Company Holidays',
    ATTENDANCE_CONFIG: 'Attendance Config',
    ATTENDANCE_RECORDS: 'Attendance Records',
    REGULARIZATION_REQUESTS: 'Regularization Requests',
    NOTIFICATIONS: 'Notifications',
    SPECIAL_EVENTS: 'Special Events',
    SALARY_DETAILS: 'Salary Details',
    SHIFT_CONFIG: 'Shift Configuration',
    EMPLOYEE_DOCUMENTS: 'Employee Documents'
  },
  // Document management configuration
  DOCUMENTS: {
    FOLDER_NAME: 'KarmSarthi Employee Documents',
    MAX_FILE_SIZE: 10 * 1024 * 1024, // 10MB in bytes
    ALLOWED_TYPES: ['application/pdf', 'image/jpeg', 'image/jpg', 'image/png', 
                    'application/msword', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'],
    ALLOWED_EXTENSIONS: ['pdf', 'jpg', 'jpeg', 'png', 'doc', 'docx'],
    DOCUMENT_TYPES: [
      'Aadhaar Card',
      'PAN Card',
      'Offer Letter',
      'Appointment Letter',
      'Educational Certificates',
      'Experience Letters',
      'Bank Passbook/Cancelled Cheque',
      'Passport',
      'Driving License',
      'Other'
    ],
    STATUS: {
      PENDING: 'Pending',
      VERIFIED: 'Verified',
      REJECTED: 'Rejected'
    }
  },
  // Default office location (New Delhi - placeholder)
  DEFAULT_OFFICE: {
    NAME: 'Head Office',
    LATITUDE: 28.6139,
    LONGITUDE: 77.2090,
    RADIUS: 100 // meters
  },
  // Default shift configurations
  DEFAULT_SHIFTS: [
    { name: 'General Shift', startTime: '09:00', endTime: '18:00', gracePeriod: 15, description: 'Standard office hours' },
    { name: 'Morning Shift', startTime: '09:00', endTime: '18:00', gracePeriod: 15, description: 'Morning shift - 9 AM to 6 PM' },
    { name: 'Evening Shift', startTime: '15:00', endTime: '00:00', gracePeriod: 15, description: 'Evening shift - 3 PM to 12 AM' },
    { name: 'Night Shift', startTime: '22:00', endTime: '07:00', gracePeriod: 15, description: 'Night shift - 10 PM to 7 AM' }
  ],
  DEFAULT_SHIFT_NAME: 'General Shift'
};

// ============================================================================
// LOCK SERVICE UTILITIES - For Concurrent Operations
// ============================================================================

/**
 * Acquire a lock with retry logic
 * @param {string} lockName - Name of the lock (for logging)
 * @param {number} timeoutSeconds - Maximum time to wait for lock (default: 30)
 * @param {number} maxRetries - Maximum number of retry attempts (default: 3)
 * @returns {Lock|null} Lock object if acquired, null if failed
 */
function acquireLock(lockName, timeoutSeconds, maxRetries) {
  timeoutSeconds = timeoutSeconds || 30;
  maxRetries = maxRetries || 3;
  
  const lock = LockService.getScriptLock();
  let attempt = 0;
  
  while (attempt < maxRetries) {
    try {
      const acquired = lock.tryLock(timeoutSeconds * 1000);
      if (acquired) {
        Logger.log(`Lock acquired: ${lockName} (attempt ${attempt + 1})`);
        return lock;
      }
    } catch (e) {
      Logger.log(`Lock acquisition error for ${lockName}: ${e.toString()}`);
    }
    
    attempt++;
    if (attempt < maxRetries) {
      // Exponential backoff: 1s, 2s, 4s
      const backoffMs = Math.pow(2, attempt - 1) * 1000;
      Utilities.sleep(backoffMs);
      Logger.log(`Retrying lock acquisition for ${lockName} (attempt ${attempt + 1})`);
    }
  }
  
  Logger.log(`Failed to acquire lock: ${lockName} after ${maxRetries} attempts`);
  return null;
}

/**
 * Safely release a lock
 * @param {Lock} lock - Lock object to release
 */
function releaseLock(lock) {
  if (lock) {
    try {
      lock.releaseLock();
      Logger.log('Lock released successfully');
    } catch (e) {
      Logger.log(`Error releasing lock: ${e.toString()}`);
    }
  }
}


/**
 * Serves the web application
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('KarmSarthi - Har din, har chhutti ka bharosa')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/**
 * Include HTML files (for CSS and JS)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSpreadsheetId() {

  return '1luy6oakqKxiMKZZB-J8KlWfA5ckcvh-TY9NGjTnegYs'; 
}


/**
 * Get or create a sheet by name
 */
function getOrCreateSheet(sheetName, headers) {
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (headers && headers.length > 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  }

  return sheet;
}

/**
 * Initialize all required sheets
 */
function initializeSheets() {
  // Employees Sheet
  getOrCreateSheet(CONFIG.SHEETS.EMPLOYEES, [
    'Employee Email', 'Employee Name', 'Department', 'Manager Email',
    'Date of Birth', 'Total Leaves', 'Leaves Used', 'Leaves Available', 'Shift Name'
  ]);

  // Leave Requests Sheet
  getOrCreateSheet(CONFIG.SHEETS.LEAVE_REQUESTS, [
    'Request ID', 'Employee Email', 'Employee Name', 'Leave Type',
    'Start Date', 'End Date', 'Days', 'Day Type', 'Units', 'Reason', 'Status',
    'Applied Date', 'Manager Email', 'Action Date', 'Manager Comments'
  ]);

  // Leave History Sheet
  getOrCreateSheet(CONFIG.SHEETS.LEAVE_HISTORY, [
    'Transaction ID', 'Employee Email', 'Date', 'Type',
    'Amount', 'Balance After', 'Description'
  ]);

  // Company Holidays Sheet
  getOrCreateSheet(CONFIG.SHEETS.COMPANY_HOLIDAYS, [
    'Holiday Date', 'Holiday Name', 'Holiday Type'
  ]);

  // Attendance Config Sheet
  const attendanceConfigSheet = getOrCreateSheet(CONFIG.SHEETS.ATTENDANCE_CONFIG, [
    'Office Name', 'Latitude', 'Longitude', 'Allowed Radius (meters)', 'Last Updated'
  ]);
  
  // Add default office location if sheet is empty
  const configData = attendanceConfigSheet.getDataRange().getValues();
  if (configData.length === 1) { // Only header row
    attendanceConfigSheet.appendRow([
      CONFIG.DEFAULT_OFFICE.NAME,
      CONFIG.DEFAULT_OFFICE.LATITUDE,
      CONFIG.DEFAULT_OFFICE.LONGITUDE,
      CONFIG.DEFAULT_OFFICE.RADIUS,
      new Date()
    ]);
  }

  // Attendance Records Sheet
  getOrCreateSheet(CONFIG.SHEETS.ATTENDANCE_RECORDS, [
    'Attendance ID', 'Work Mode', 'Employee Email', 'Employee Name', 'Date',
    'Check-In Time', 'Check-In Lat', 'Check-In Long', 'Check-In Distance',
    'Check-Out Time', 'Check-Out Lat', 'Check-Out Long', 'Check-Out Distance',
    'Total Hours', 'Status', 'Timestamp'
  ]);

  // Regularization Requests Sheet
  getOrCreateSheet(CONFIG.SHEETS.REGULARIZATION_REQUESTS, [
    'Request ID', 'Employee Email', 'Employee Name', 'Request Date',
    'Attendance Date', 'Reason', 'Status', 'Manager Email',
    'Manager Comments', 'Action Date', 'Action By'
  ]);

  // Notifications Sheet
  getOrCreateSheet(CONFIG.SHEETS.NOTIFICATIONS, [
    'Notification ID', 'Recipient Email', 'Sender Email', 'Sender Name',
    'Event Type', 'Request ID', 'Message', 'Timestamp', 'Read Status'
  ]);

  // Special Events Sheet
  getOrCreateSheet(CONFIG.SHEETS.SPECIAL_EVENTS, [
    'Event Title', 'Event Date', 'Description', 'Type'
  ]);

  // Employee Details Sheet
  getOrCreateSheet(CONFIG.SHEETS.EMPLOYEE_DETAILS, [
    'Employee Email', 'Employee ID', 'Father Name', 'Date of Joining',
    'Phone Number', 'Address', 'Designation', 'PAN Number', 'Aadhaar Number',
    'UAN Number', 'ESI Number', 'CTC', 'Bank Name', 'Account Number',
    'IFSC Code', 'Branch Name', 'Account Type'
  ]);

  // Salary Details Sheet
  getOrCreateSheet(CONFIG.SHEETS.SALARY_DETAILS, [
    'Employee Email', 'Month', 'Year', 'Basic Salary', 'HRA',
    'Conveyance Allowance', 'Medical Allowance', 'Special Allowance',
    'Other Earnings', 'Gross Salary', 'PF Deduction', 'ESI Deduction',
    'Professional Tax', 'TDS', 'Other Deductions', 'Total Deductions',
    'Net Salary', 'Entry Date', 'Updated By'
  ]);

  // Shift Configuration Sheet
  const shiftConfigSheet = getOrCreateSheet(CONFIG.SHEETS.SHIFT_CONFIG, [
    'Shift Name', 'Start Time', 'End Time', 'Grace Period (minutes)', 'Description'
  ]);
  
  // Add default shifts if sheet is empty
  const shiftData = shiftConfigSheet.getDataRange().getValues();
  if (shiftData.length === 1) { // Only header row
    CONFIG.DEFAULT_SHIFTS.forEach(shift => {
      shiftConfigSheet.appendRow([
        shift.name,
        shift.startTime,
        shift.endTime,
        shift.gracePeriod,
        shift.description
      ]);
    });
  }

  // Employee Documents Sheet
  getOrCreateSheet(CONFIG.SHEETS.EMPLOYEE_DOCUMENTS, [
    'Document ID', 'Employee Email', 'Employee Name', 'Document Type',
    'File Name', 'Drive File ID', 'Upload Date', 'File Size (bytes)',
    'Verification Status', 'Verified By', 'Verification Date', 'HR Comments'
  ]);
}

/**
 * Get current user's email
 */
function getCurrentUserEmail() {
  return Session.getActiveUser().getEmail();
}

/**
 * Get employee data for current user
 */
function getEmployeeData() {
  const email = getCurrentUserEmail();
  const sheet = getOrCreateSheet(CONFIG.SHEETS.EMPLOYEES, [
    'Employee Email', 'Employee Name', 'Department', 'Manager Email',
    'Date of Birth', 'Total Leaves', 'Leaves Used', 'Leaves Available', 'Shift Name'
  ]);

  const data = sheet.getDataRange().getValues();

  // Find employee record
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      return {
        email: data[i][0],
        name: data[i][1],
        department: data[i][2],
        managerEmail: data[i][3],
        dateOfBirth: data[i][4] ? formatDate(data[i][4]) : '',
        totalLeaves: data[i][5],
        leavesUsed: data[i][6],
        leavesAvailable: data[i][7],
        shiftName: data[i][8] || CONFIG.DEFAULT_SHIFT_NAME // Default to General Shift for backward compatibility
      };
    }
  }

  // If employee not found, return null
  return null;
}

/**
 * Get comprehensive employee details for profile display
 */
function getEmployeeDetails() {
  const email = getCurrentUserEmail();
  
  // Get basic employee data from Employees sheet
  const employeeSheet = getOrCreateSheet(CONFIG.SHEETS.EMPLOYEES, [
    'Employee Email', 'Employee Name', 'Department', 'Manager Email',
    'Date of Birth', 'Total Leaves', 'Leaves Used', 'Leaves Available', 'Shift Name'
  ]);
  
  const employeeData = employeeSheet.getDataRange().getValues();
  let employeeInfo = null;
  
  // Find employee record
  for (let i = 1; i < employeeData.length; i++) {
    if (employeeData[i][0] === email) {
      employeeInfo = {
        email: employeeData[i][0],
        name: employeeData[i][1],
        department: employeeData[i][2],
        managerEmail: employeeData[i][3],
        dateOfBirth: employeeData[i][4] ? formatDate(employeeData[i][4]) : '',
        shiftName: employeeData[i][8] || CONFIG.DEFAULT_SHIFT_NAME
      };
      break;
    }
  }
  
  if (!employeeInfo) {
    return null;
  }
  
  // Get additional details from Employee Details sheet
  const detailsSheet = getOrCreateSheet(CONFIG.SHEETS.EMPLOYEE_DETAILS, [
    'Employee Email', 'Employee ID', 'Father Name', 'Date of Joining',
    'Phone Number', 'Address', 'Designation', 'PAN Number', 'Aadhaar Number',
    'UAN Number', 'ESI Number', 'CTC', 'Bank Name', 'Account Number',
    'IFSC Code', 'Branch Name', 'Account Type'
  ]);
  
  const detailsData = detailsSheet.getDataRange().getValues();
  
  // Find employee details record
  for (let i = 1; i < detailsData.length; i++) {
    if (detailsData[i][0] === email) {
      employeeInfo.employeeId = detailsData[i][1] || '';
      employeeInfo.fatherName = detailsData[i][2] || '';
      employeeInfo.dateOfJoining = detailsData[i][3] ? formatDate(detailsData[i][3]) : '';
      employeeInfo.phoneNumber = detailsData[i][4] || '';
      employeeInfo.address = detailsData[i][5] || '';
      employeeInfo.designation = detailsData[i][6] || '';
      employeeInfo.panNumber = detailsData[i][7] || '';
      employeeInfo.aadhaarNumber = detailsData[i][8] || '';
      employeeInfo.uanNumber = detailsData[i][9] || '';
      employeeInfo.esiNumber = detailsData[i][10] || '';
      employeeInfo.ctc = detailsData[i][11] || '';
      employeeInfo.bankName = detailsData[i][12] || '';
      employeeInfo.accountNumber = detailsData[i][13] || '';
      employeeInfo.ifscCode = detailsData[i][14] || '';
      employeeInfo.branchName = detailsData[i][15] || '';
      employeeInfo.accountType = detailsData[i][16] || '';
      break;
    }
  }
  
  // If no details found, set default empty values
  if (!employeeInfo.employeeId) {
    employeeInfo.employeeId = '';
    employeeInfo.fatherName = '';
    employeeInfo.dateOfJoining = '';
    employeeInfo.phoneNumber = '';
    employeeInfo.address = '';
    employeeInfo.designation = '';
    employeeInfo.panNumber = '';
    employeeInfo.aadhaarNumber = '';
    employeeInfo.uanNumber = '';
    employeeInfo.esiNumber = '';
    employeeInfo.ctc = '';
    employeeInfo.bankName = '';
    employeeInfo.accountNumber = '';
    employeeInfo.ifscCode = '';
    employeeInfo.branchName = '';
    employeeInfo.accountType = '';
  }
  
  return employeeInfo;
}


/**
 * Initialize a new employee
 */
function initializeEmployee(email, name, department, managerEmail, dateOfBirth, shiftName) {
  const sheet = getOrCreateSheet(CONFIG.SHEETS.EMPLOYEES, [
    'Employee Email', 'Employee Name', 'Department', 'Manager Email',
    'Date of Birth', 'Total Leaves', 'Leaves Used', 'Leaves Available', 'Shift Name'
  ]);

  // Check if employee already exists
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      return { success: false, message: 'Employee already exists' };
    }
  }

  // Add new employee with initial leave balance and shift
  const initialLeaves = CONFIG.MONTHLY_LEAVE_ACCRUAL;
  const assignedShift = shiftName || CONFIG.DEFAULT_SHIFT_NAME;
  sheet.appendRow([email, name, department, managerEmail, dateOfBirth || '', initialLeaves, 0, initialLeaves, assignedShift]);

  // Log in history
  logLeaveHistory(email, 'Accrual', initialLeaves, initialLeaves, 'Initial leave allocation');

  return { success: true, message: 'Employee initialized successfully' };
}

/**
 * Get leave requests for current user
 */
function getLeaveRequests() {
  const email = getCurrentUserEmail();
  const sheet = getOrCreateSheet(CONFIG.SHEETS.LEAVE_REQUESTS, [
    'Request ID', 'Employee Email', 'Employee Name', 'Leave Type',
    'Start Date', 'End Date', 'Days', 'Day Type', 'Units', 'Reason', 'Status',
    'Applied Date', 'Manager Email', 'Action Date', 'Manager Comments'
  ]);

  const data = sheet.getDataRange().getValues();
  const requests = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === email) {
      requests.push({
        requestId: data[i][0],
        employeeEmail: data[i][1],
        employeeName: data[i][2],
        leaveType: data[i][3],
        startDate: formatDate(data[i][4]),
        endDate: formatDate(data[i][5]),
        days: data[i][6],
        dayType: data[i][7] || 'Full Day', // Backward compatibility
        units: data[i][8] || data[i][6], // Backward compatibility
        reason: data[i][9],
        status: data[i][10],
        appliedDate: formatDate(data[i][11]),
        managerEmail: data[i][12],
        actionDate: data[i][13] ? formatDate(data[i][13]) : '',
        managerComments: data[i][14] || ''
      });
    }
  }

  return requests.reverse(); // Most recent first
}

/**
 * Get pending leave requests for manager
 */
function getManagerData() {
  const email = getCurrentUserEmail();
  const sheet = getOrCreateSheet(CONFIG.SHEETS.LEAVE_REQUESTS, [
    'Request ID', 'Employee Email', 'Employee Name', 'Leave Type',
    'Start Date', 'End Date', 'Days', 'Day Type', 'Units', 'Reason', 'Status',
    'Applied Date', 'Manager Email', 'Action Date', 'Manager Comments'
  ]);

  const data = sheet.getDataRange().getValues();
  const pendingRequests = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][12] === email && data[i][10] === 'Pending') {
      pendingRequests.push({
        requestId: data[i][0],
        employeeEmail: data[i][1],
        employeeName: data[i][2],
        leaveType: data[i][3],
        startDate: formatDate(data[i][4]),
        endDate: formatDate(data[i][5]),
        days: data[i][6],
        dayType: data[i][7] || 'Full Day',
        units: data[i][8] || data[i][6],
        reason: data[i][9],
        status: data[i][10],
        appliedDate: formatDate(data[i][11]),
        managerEmail: data[i][12]
      });
    }
  }

  return pendingRequests;
}

/**
 * Get team leave calendar data for a specific month/year
 * Returns all approved and pending leaves for the manager's team
 */
function getTeamLeaveCalendar(month, year) {
  const managerEmail = getCurrentUserEmail();
  
  // Get all employees under this manager
  const employeeSheet = getOrCreateSheet(CONFIG.SHEETS.EMPLOYEES, [
    'Employee Email', 'Employee Name', 'Department', 'Manager Email',
    'Date of Birth', 'Total Leaves', 'Leaves Used', 'Leaves Available'
  ]);
  
  const employeeData = employeeSheet.getDataRange().getValues();
  const teamEmails = [];
  const teamMembers = {};
  
  // Collect team members
  for (let i = 1; i < employeeData.length; i++) {
    if (employeeData[i][3] === managerEmail) {
      const email = employeeData[i][0];
      teamEmails.push(email);
      teamMembers[email] = {
        name: employeeData[i][1],
        department: employeeData[i][2]
      };
    }
  }
  
  // Get leave requests for the team
  const leaveSheet = getOrCreateSheet(CONFIG.SHEETS.LEAVE_REQUESTS, [
    'Request ID', 'Employee Email', 'Employee Name', 'Leave Type',
    'Start Date', 'End Date', 'Days', 'Day Type', 'Units', 'Reason', 'Status',
    'Applied Date', 'Manager Email', 'Action Date', 'Manager Comments'
  ]);
  
  const leaveData = leaveSheet.getDataRange().getValues();
  const calendarData = {};
  
  // Filter leaves for the specified month/year and team members
  for (let i = 1; i < leaveData.length; i++) {
    const employeeEmail = leaveData[i][1];
    const status = leaveData[i][10];
    
    // Only include approved and pending leaves for team members
    if (teamEmails.includes(employeeEmail) && (status === 'Approved' || status === 'Pending')) {
      const startDate = new Date(leaveData[i][4]);
      const endDate = new Date(leaveData[i][5]);
      
      // Iterate through each day of the leave
      const currentDate = new Date(startDate);
      while (currentDate <= endDate) {
        const dateMonth = currentDate.getMonth();
        const dateYear = currentDate.getFullYear();
        
        // Check if this date falls in the requested month/year
        if (dateMonth === month && dateYear === year) {
          const dateKey = formatDate(currentDate);
          
          if (!calendarData[dateKey]) {
            calendarData[dateKey] = [];
          }
          
          // Check if this employee already has an entry for this date
          const existingEntry = calendarData[dateKey].find(entry => entry.employeeEmail === employeeEmail);
          
          if (!existingEntry) {
            calendarData[dateKey].push({
              employeeEmail: employeeEmail,
              employeeName: leaveData[i][2],
              leaveType: leaveData[i][3],
              dayType: leaveData[i][7] || 'Full Day',
              status: status,
              requestId: leaveData[i][0]
            });
          }
        }
        
        currentDate.setDate(currentDate.getDate() + 1);
      }
    }
  }
  
  return {
    calendarData: calendarData,
    teamSize: teamEmails.length
  };
}

/**
 * Get team availability statistics for a specific date
 */
function getTeamAvailabilityStats(date) {
  const managerEmail = getCurrentUserEmail();
  
  // Get all employees under this manager
  const employeeSheet = getOrCreateSheet(CONFIG.SHEETS.EMPLOYEES, [
    'Employee Email', 'Employee Name', 'Department', 'Manager Email',
    'Date of Birth', 'Total Leaves', 'Leaves Used', 'Leaves Available'
  ]);
  
  const employeeData = employeeSheet.getDataRange().getValues();
  const teamEmails = [];
  
  for (let i = 1; i < employeeData.length; i++) {
    if (employeeData[i][3] === managerEmail) {
      teamEmails.push(employeeData[i][0]);
    }
  }
  
  const totalTeam = teamEmails.length;
  
  // Get leave requests for this date
  const leaveSheet = getOrCreateSheet(CONFIG.SHEETS.LEAVE_REQUESTS, [
    'Request ID', 'Employee Email', 'Employee Name', 'Leave Type',
    'Start Date', 'End Date', 'Days', 'Day Type', 'Units', 'Reason', 'Status',
    'Applied Date', 'Manager Email', 'Action Date', 'Manager Comments'
  ]);
  
  const leaveData = leaveSheet.getDataRange().getValues();
  const onLeave = new Set();
  const leaveBreakdown = {};
  
  const targetDate = new Date(date);
  
  for (let i = 1; i < leaveData.length; i++) {
    const employeeEmail = leaveData[i][1];
    const status = leaveData[i][10];
    
    if (teamEmails.includes(employeeEmail) && status === 'Approved') {
      const startDate = new Date(leaveData[i][4]);
      const endDate = new Date(leaveData[i][5]);
      
      // Check if target date falls within leave period
      if (targetDate >= startDate && targetDate <= endDate) {
        onLeave.add(employeeEmail);
        
        const leaveType = leaveData[i][3];
        leaveBreakdown[leaveType] = (leaveBreakdown[leaveType] || 0) + 1;
      }
    }
  }
  
  return {
    totalTeam: totalTeam,
    onLeave: onLeave.size,
    available: totalTeam - onLeave.size,
    leaveBreakdown: leaveBreakdown
  };
}


/**
 * Validate half day leave request
 */
function validateHalfDayRequest(startDate, endDate, dayType) {
  // Half day only allowed for single date
  if ((dayType === 'First Half' || dayType === 'Second Half')) {
    if (startDate !== endDate) {
      return { valid: false, message: 'Half day leave is only allowed for single date' };
    }

    // Check if date is weekend
    const date = new Date(startDate);
    if (isWeekend(date)) {
      return { valid: false, message: 'Half day leave cannot be taken on weekends (Sunday)' };
    }

    // Check if date is company holiday
    const holidays = getCompanyHolidays();
    if (isCompanyHoliday(date, holidays)) {
      return { valid: false, message: 'Half day leave cannot be taken on company holidays' };
    }
  }

  return { valid: true, message: '' };
}

/**
 * Submit a new leave request
 */
function submitLeaveRequest(leaveType, startDate, endDate, reason, dayType) {
  const email = getCurrentUserEmail();
  const employeeData = getEmployeeData();

  if (!employeeData) {
    return { success: false, message: 'Employee not found. Please contact HR.' };
  }

  // Default to Full Day if not specified
  if (!dayType) {
    dayType = 'Full Day';
  }

  // Validate half day request
  const validation = validateHalfDayRequest(startDate, endDate, dayType);
  if (!validation.valid) {
    return { success: false, message: validation.message };
  }

  // Calculate leave days and units
  let days, units;
  if (dayType === 'First Half' || dayType === 'Second Half') {
    days = 1;
    units = 0.5;
  } else {
    days = calculateLeaveDays(new Date(startDate), new Date(endDate));
    units = days * 1.0;
  }

  if (days <= 0 || units <= 0) {
    return { success: false, message: 'Invalid date range' };
  }

  // Note: We allow submission even with insufficient balance
  // The frontend will show a warning and require confirmation
  // Excess days will be deducted from salary upon approval

  // Acquire lock for concurrent operations
  const lock = acquireLock('leave_submit', 30, 3);
  
  if (!lock) {
    return {
      success: false,
      message: 'System is busy processing multiple requests. Please try again in a moment.'
    };
  }

  try {
    const sheet = getOrCreateSheet(CONFIG.SHEETS.LEAVE_REQUESTS, [
      'Request ID', 'Employee Email', 'Employee Name', 'Leave Type',
      'Start Date', 'End Date', 'Days', 'Day Type', 'Units', 'Reason', 'Status',
      'Applied Date', 'Manager Email', 'Action Date', 'Manager Comments'
    ]);

    const requestId = generateRequestId();
    const appliedDate = new Date();

    sheet.appendRow([
      requestId,
      email,
      employeeData.name,
      leaveType,
      startDate,
      endDate,
      days,
      dayType,
      units,
      reason,
      'Pending',
      appliedDate,
      employeeData.managerEmail,
      '',
      ''
    ]);

    // Send professional notification to manager
    try {
      sendLeaveSubmissionEmail(
        employeeData.managerEmail,
        employeeData.name,
        email,
        leaveType,
        startDate,
        endDate,
        units,
        dayType,
        reason,
        requestId
      );
    } catch (e) {
      Logger.log('Failed to send submission email: ' + e.toString());
    }

    // Create notification for manager
    try {
      const dateRange = startDate === endDate ? startDate : `${startDate} to ${endDate}`;
      createNotification(
        employeeData.managerEmail,
        email,
        employeeData.name,
        'new_request',
        requestId,
        `${employeeData.name} has submitted a new ${leaveType} request for ${dateRange} (${dayType})`
      );
    } catch (e) {
      Logger.log('Failed to create notification: ' + e.toString());
    }

    return { success: true, message: 'Leave request submitted successfully', requestId: requestId };
    
  } finally {
    releaseLock(lock);
  }
}

/**
 * Approve a leave request
 */
function approveLeave(requestId, comments) {
  const managerEmail = getCurrentUserEmail();
  const sheet = getOrCreateSheet(CONFIG.SHEETS.LEAVE_REQUESTS, [
    'Request ID', 'Employee Email', 'Employee Name', 'Leave Type',
    'Start Date', 'End Date', 'Days', 'Day Type', 'Units', 'Reason', 'Status',
    'Applied Date', 'Manager Email', 'Action Date', 'Manager Comments'
  ]);

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === requestId && data[i][12] === managerEmail) {
      const employeeEmail = data[i][1];
      const employeeName = data[i][2];
      const leaveType = data[i][3];
      const dayType = data[i][7] || 'Full Day';
      const units = data[i][8] || data[i][6]; // Use units for deduction

      // Update request status
      sheet.getRange(i + 1, 11).setValue('Approved');
      sheet.getRange(i + 1, 14).setValue(new Date());
      sheet.getRange(i + 1, 15).setValue(comments || 'Approved');

      // Deduct leaves from employee balance (using units)
      deductLeaves(employeeEmail, units, `Leave approved - ${leaveType} (${dayType}) (${requestId})`);

      // Send professional notification to employee
      try {
        sendLeaveApprovalEmail(
          employeeEmail,
          employeeName,
          leaveType,
          data[i][4], // startDate
          data[i][5], // endDate
          units,
          dayType,
          requestId,
          comments
        );
      } catch (e) {
        Logger.log('Failed to send approval email: ' + e.toString());
      }

      // Create notification for employee
      try {
        const dateRange = formatDate(data[i][4]) === formatDate(data[i][5]) ? formatDate(data[i][4]) : `${formatDate(data[i][4])} to ${formatDate(data[i][5])}`;
        const managerData = getEmployeeData();
        createNotification(
          employeeEmail,
          managerEmail,
          managerData ? managerData.name : 'Manager',
          'approved',
          requestId,
          `Your ${leaveType} request for ${dateRange} has been approved`
        );
      } catch (e) {
        Logger.log('Failed to create notification: ' + e.toString());
      }

      return { success: true, message: 'Leave approved successfully' };
    }
  }

  return { success: false, message: 'Request not found or unauthorized' };
}

/**
 * Reject a leave request
 */
function rejectLeave(requestId, comments) {
  const managerEmail = getCurrentUserEmail();
  const sheet = getOrCreateSheet(CONFIG.SHEETS.LEAVE_REQUESTS, [
    'Request ID', 'Employee Email', 'Employee Name', 'Leave Type',
    'Start Date', 'End Date', 'Days', 'Day Type', 'Units', 'Reason', 'Status',
    'Applied Date', 'Manager Email', 'Action Date', 'Manager Comments'
  ]);

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === requestId && data[i][12] === managerEmail) {
      const employeeEmail = data[i][1];
      const units = data[i][8] || data[i][6];

      // Update request status
      sheet.getRange(i + 1, 11).setValue('Rejected');
      sheet.getRange(i + 1, 14).setValue(new Date());
      sheet.getRange(i + 1, 15).setValue(comments || 'Rejected');

      // Send professional notification to employee
      try {
        sendLeaveRejectionEmail(
          employeeEmail,
          data[i][2], // employeeName
          data[i][3], // leaveType
          data[i][4], // startDate
          data[i][5], // endDate
          units,
          data[i][7] || 'Full Day', // dayType
          requestId,
          comments
        );
      } catch (e) {
        Logger.log('Failed to send rejection email: ' + e.toString());
      }

      // Create notification for employee
      try {
        const dateRange = formatDate(data[i][4]) === formatDate(data[i][5]) ? formatDate(data[i][4]) : `${formatDate(data[i][4])} to ${formatDate(data[i][5])}`;
        const managerData = getEmployeeData();
        createNotification(
          employeeEmail,
          managerEmail,
          managerData ? managerData.name : 'Manager',
          'rejected',
          requestId,
          `Your ${data[i][3]} request for ${dateRange} has been rejected`
        );
      } catch (e) {
        Logger.log('Failed to create notification: ' + e.toString());
      }

      return { success: true, message: 'Leave rejected successfully' };
    }
  }

  return { success: false, message: 'Request not found or unauthorized' };
}

/**
 * Cancel a leave request (employee can cancel their own pending or approved leaves)
 */
function cancelLeaveRequest(requestId, cancellationReason) {
  const employeeEmail = getCurrentUserEmail();
  const sheet = getOrCreateSheet(CONFIG.SHEETS.LEAVE_REQUESTS, [
    'Request ID', 'Employee Email', 'Employee Name', 'Leave Type',
    'Start Date', 'End Date', 'Days', 'Day Type', 'Units', 'Reason', 'Status',
    'Applied Date', 'Manager Email', 'Action Date', 'Manager Comments'
  ]);

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === requestId && data[i][1] === employeeEmail) {
      const status = data[i][10];
      const units = data[i][8] || data[i][6];
      const dayType = data[i][7] || 'Full Day';
      const leaveType = data[i][3];
      const managerEmail = data[i][12];
      const employeeName = data[i][2];

      // Check if leave can be cancelled
      if (status === 'Cancelled') {
        return { success: false, message: 'Leave is already cancelled' };
      }
      if (status === 'Rejected') {
        return { success: false, message: 'Cannot cancel a rejected leave' };
      }

      // For approved leaves, require cancellation reason
      if (status === 'Approved' && (!cancellationReason || cancellationReason.trim() === '')) {
        return { success: false, message: 'Cancellation reason is required for approved leaves' };
      }

      // Update request status
      sheet.getRange(i + 1, 11).setValue('Cancelled');
      sheet.getRange(i + 1, 14).setValue(new Date());
      const comments = status === 'Approved' 
        ? `Cancelled by employee. Reason: ${cancellationReason}`
        : 'Cancelled by employee';
      sheet.getRange(i + 1, 15).setValue(comments);

      // If leave was approved, restore the leave balance (using units)
      if (status === 'Approved') {
        restoreLeaves(employeeEmail, units, `Leave cancelled - ${leaveType} (${dayType}) (${requestId})`);
      }

      // Send professional notification to manager
      try {
        sendLeaveCancellationEmail(
          managerEmail,
          employeeName,
          employeeEmail,
          leaveType,
          data[i][4], // startDate
          data[i][5], // endDate
          units,
          dayType,
          requestId,
          cancellationReason
        );
      } catch (e) {
        Logger.log('Failed to send cancellation email: ' + e.toString());
      }

      // Create notification for manager
      try {
        const dateRange = formatDate(data[i][4]) === formatDate(data[i][5]) ? formatDate(data[i][4]) : `${formatDate(data[i][4])} to ${formatDate(data[i][5])}`;
        createNotification(
          managerEmail,
          employeeEmail,
          employeeName,
          'cancelled',
          requestId,
          `${employeeName} has cancelled their ${leaveType} request for ${dateRange}`
        );
      } catch (e) {
        Logger.log('Failed to create notification: ' + e.toString());
      }

      return { success: true, message: 'Leave request cancelled successfully' };
    }
  }

  return { success: false, message: 'Request not found or unauthorized' };
}

/**
 * Edit a pending leave request
 */
function editLeaveRequest(requestId, leaveType, startDate, endDate, reason, dayType) {
  const employeeEmail = getCurrentUserEmail();
  const sheet = getOrCreateSheet(CONFIG.SHEETS.LEAVE_REQUESTS, [
    'Request ID', 'Employee Email', 'Employee Name', 'Leave Type',
    'Start Date', 'End Date', 'Days', 'Day Type', 'Units', 'Reason', 'Status',
    'Applied Date', 'Manager Email', 'Action Date', 'Manager Comments'
  ]);

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === requestId && data[i][1] === employeeEmail) {
      const status = data[i][10];
      const employeeName = data[i][2];
      const managerEmail = data[i][12];

      // Only pending leaves can be edited
      if (status !== 'Pending') {
        return { success: false, message: 'Only pending leave requests can be edited' };
      }

      // Default to Full Day if not specified
      if (!dayType) {
        dayType = 'Full Day';
      }

      // Validate half day request
      const validation = validateHalfDayRequest(startDate, endDate, dayType);
      if (!validation.valid) {
        return { success: false, message: validation.message };
      }

      // Calculate new leave days and units
      let days, units;
      if (dayType === 'First Half' || dayType === 'Second Half') {
        days = 1;
        units = 0.5;
      } else {
        days = calculateLeaveDays(new Date(startDate), new Date(endDate));
        units = days * 1.0;
      }

      if (days <= 0 || units <= 0) {
        return { success: false, message: 'Invalid date range' };
      }

      // Update leave request
      sheet.getRange(i + 1, 4).setValue(leaveType);
      sheet.getRange(i + 1, 5).setValue(startDate);
      sheet.getRange(i + 1, 6).setValue(endDate);
      sheet.getRange(i + 1, 7).setValue(days);
      sheet.getRange(i + 1, 8).setValue(dayType);
      sheet.getRange(i + 1, 9).setValue(units);
      sheet.getRange(i + 1, 10).setValue(reason);

      // Send updated notification to manager
      const dateRange = startDate === endDate ? startDate : `${startDate} to ${endDate}`;
      sendEmailNotification(
        managerEmail,
        'Leave Request Updated',
        `${employeeName} has updated their leave request (${requestId}).\n\nUpdated Details:\nLeave Type: ${leaveType}\nDay Type: ${dayType}\nDate: ${dateRange}\nUnits: ${units}\nReason: ${reason}\n\nPlease review and approve/reject this updated request.`
      );

      // Create notification for manager
      try {
        createNotification(
          managerEmail,
          employeeEmail,
          employeeName,
          'updated',
          requestId,
          `${employeeName} has updated their ${leaveType} request for ${dateRange} (${dayType})`
        );
      } catch (e) {
        Logger.log('Failed to create notification: ' + e.toString());
      }

      return { success: true, message: 'Leave request updated successfully' };
    }
  }

  return { success: false, message: 'Request not found or unauthorized' };
}

/**
 * Restore leaves to employee balance (when leave is cancelled)
 */
function restoreLeaves(employeeEmail, days, description) {
  const sheet = getOrCreateSheet(CONFIG.SHEETS.EMPLOYEES, [
    'Employee Email', 'Employee Name', 'Department', 'Manager Email',
    'Date of Birth', 'Total Leaves', 'Leaves Used', 'Leaves Available'
  ]);

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === employeeEmail) {
      const leavesUsed = data[i][6] - days;
      const leavesAvailable = data[i][7] + days;

      sheet.getRange(i + 1, 7).setValue(leavesUsed);
      sheet.getRange(i + 1, 8).setValue(leavesAvailable);

      // Log in history
      logLeaveHistory(employeeEmail, 'Restoration', days, leavesAvailable, description);

      return true;
    }
  }

  return false;
}


/**
 * Deduct leaves from employee balance
 */
function deductLeaves(employeeEmail, days, description) {
  const sheet = getOrCreateSheet(CONFIG.SHEETS.EMPLOYEES, [
    'Employee Email', 'Employee Name', 'Department', 'Manager Email',
    'Date of Birth', 'Total Leaves', 'Leaves Used', 'Leaves Available'
  ]);

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === employeeEmail) {
      const leavesUsed = data[i][6] + days;
      const leavesAvailable = data[i][5] - leavesUsed;

      sheet.getRange(i + 1, 7).setValue(leavesUsed);
      sheet.getRange(i + 1, 8).setValue(leavesAvailable);

      // Log in history
      logLeaveHistory(employeeEmail, 'Deduction', days, leavesAvailable, description);

      return true;
    }
  }

  return false;
}

/**
 * Accrue monthly leaves for all employees
 * This should be triggered on the 1st of every month
 */
function accrueMonthlyLeaves() {
  const sheet = getOrCreateSheet(CONFIG.SHEETS.EMPLOYEES, [
    'Employee Email', 'Employee Name', 'Department', 'Manager Email',
    'Date of Birth', 'Total Leaves', 'Leaves Used', 'Leaves Available'
  ]);

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const employeeEmail = data[i][0];
    const totalLeaves = data[i][5] + CONFIG.MONTHLY_LEAVE_ACCRUAL;
    const leavesAvailable = data[i][7] + CONFIG.MONTHLY_LEAVE_ACCRUAL;

    sheet.getRange(i + 1, 6).setValue(totalLeaves);
    sheet.getRange(i + 1, 8).setValue(leavesAvailable);

    // Log in history
    logLeaveHistory(
      employeeEmail,
      'Accrual',
      CONFIG.MONTHLY_LEAVE_ACCRUAL,
      leavesAvailable,
      'Monthly leave accrual'
    );
  }

  Logger.log(`Accrued ${CONFIG.MONTHLY_LEAVE_ACCRUAL} leaves for ${data.length - 1} employees`);
}

/**
 * Log leave history
 */
function logLeaveHistory(employeeEmail, type, amount, balanceAfter, description) {
  const sheet = getOrCreateSheet(CONFIG.SHEETS.LEAVE_HISTORY, [
    'Transaction ID', 'Employee Email', 'Date', 'Type',
    'Amount', 'Balance After', 'Description'
  ]);

  const transactionId = generateTransactionId();

  sheet.appendRow([
    transactionId,
    employeeEmail,
    new Date(),
    type,
    amount,
    balanceAfter,
    description
  ]);
}

/**
 * Generate unique request ID
 */
function generateRequestId() {
  return 'LR' + new Date().getTime();
}

/**
 * Generate unique transaction ID
 */
function generateTransactionId() {
  return 'TXN' + new Date().getTime();
}

/**
 * Generate unique notification ID
 */
function generateNotificationId() {
  return 'NOTIF' + new Date().getTime();
}

/**
 * Generate unique regularization request ID
 */
function generateRegularizationId() {
  return 'REG' + new Date().getTime() + Math.random().toString(36).substr(2, 5);
}

/**
 * Create a new notification
 */
function createNotification(recipientEmail, senderEmail, senderName, eventType, requestId, message) {
  const sheet = getOrCreateSheet(CONFIG.SHEETS.NOTIFICATIONS, [
    'Notification ID', 'Recipient Email', 'Sender Email', 'Sender Name',
    'Event Type', 'Request ID', 'Message', 'Timestamp', 'Read Status'
  ]);

  const notificationId = generateNotificationId();
  const timestamp = new Date();

  sheet.appendRow([
    notificationId,
    recipientEmail,
    senderEmail,
    senderName,
    eventType,
    requestId,
    message,
    timestamp,
    false // Read Status
  ]);

  return notificationId;
}

/**
 * Get all notifications for a user
 */
function getNotifications(userEmail) {
  const sheet = getOrCreateSheet(CONFIG.SHEETS.NOTIFICATIONS, [
    'Notification ID', 'Recipient Email', 'Sender Email', 'Sender Name',
    'Event Type', 'Request ID', 'Message', 'Timestamp', 'Read Status'
  ]);

  const data = sheet.getDataRange().getValues();
  const notifications = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === userEmail) {
      notifications.push({
        notificationId: data[i][0],
        recipientEmail: data[i][1],
        senderEmail: data[i][2],
        senderName: data[i][3],
        eventType: data[i][4],
        requestId: data[i][5],
        message: data[i][6],
        timestamp: formatDateTime(data[i][7]),
        readStatus: data[i][8]
      });
    }
  }

  // Return most recent first
  return notifications.reverse();
}

/**
 * Get current user's notifications with unread count (wrapper for client-side)
 */
function getCurrentUserNotifications() {
  const userEmail = Session.getActiveUser().getEmail();
  const notifications = getNotifications(userEmail);
  const unreadCount = getUnreadNotificationCount(userEmail);
  
  return {
    notifications: notifications,
    unreadCount: unreadCount
  };
}

/**
 * Get unread notification count for a user
 */
function getUnreadNotificationCount(userEmail) {
  const sheet = getOrCreateSheet(CONFIG.SHEETS.NOTIFICATIONS, [
    'Notification ID', 'Recipient Email', 'Sender Email', 'Sender Name',
    'Event Type', 'Request ID', 'Message', 'Timestamp', 'Read Status'
  ]);

  const data = sheet.getDataRange().getValues();
  let count = 0;

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === userEmail && data[i][8] === false) {
      count++;
    }
  }

  return count;
}

/**
 * Mark a notification as read
 */
function markNotificationAsRead(notificationId) {
  const sheet = getOrCreateSheet(CONFIG.SHEETS.NOTIFICATIONS, [
    'Notification ID', 'Recipient Email', 'Sender Email', 'Sender Name',
    'Event Type', 'Request ID', 'Message', 'Timestamp', 'Read Status'
  ]);

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === notificationId) {
      sheet.getRange(i + 1, 9).setValue(true);
      return { success: true, message: 'Notification marked as read' };
    }
  }

  return { success: false, message: 'Notification not found' };
}

/**
 * Mark all notifications as read for a user
 */
function markAllNotificationsAsRead(userEmail) {
  const sheet = getOrCreateSheet(CONFIG.SHEETS.NOTIFICATIONS, [
    'Notification ID', 'Recipient Email', 'Sender Email', 'Sender Name',
    'Event Type', 'Request ID', 'Message', 'Timestamp', 'Read Status'
  ]);

  const data = sheet.getDataRange().getValues();
  let count = 0;

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === userEmail && data[i][8] === false) {
      sheet.getRange(i + 1, 9).setValue(true);
      count++;
    }
  }

  return { success: true, message: `Marked ${count} notifications as read`, count: count };
}

/**
 * Mark all notifications as read for current user (wrapper for client-side)
 */
function markAllNotificationsAsReadForCurrentUser() {
  const userEmail = Session.getActiveUser().getEmail();
  return markAllNotificationsAsRead(userEmail);
}

/**
 * Clear all notifications for current user
 */
function clearAllNotifications() {
  const userEmail = Session.getActiveUser().getEmail();
  const sheet = getOrCreateSheet(CONFIG.SHEETS.NOTIFICATIONS, [
    'Notification ID', 'Recipient Email', 'Sender Email', 'Sender Name',
    'Event Type', 'Request ID', 'Message', 'Timestamp', 'Read Status'
  ]);
  
  const data = sheet.getDataRange().getValues();
  
  // Keep header row
  if (data.length <= 1) return { success: true, message: 'No notifications to clear' };
  
  const header = data[0];
  const recipientIdx = header.indexOf('Recipient Email');
  
  // Filter out user's notifications (keep ones that DON'T match)
  const newData = [header];
  let clearedCount = 0;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][recipientIdx] !== userEmail) {
      newData.push(data[i]);
    } else {
      clearedCount++;
    }
  }
  
  // Write back if changes were made
  if (clearedCount > 0) {
    sheet.clearContents();
    if (newData.length > 0) {
      sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
    }
    return { success: true, message: 'Notifications cleared successfully' };
  } else {
    return { success: true, message: 'No notifications to clear' };
  }
}

// ============================================================================
// REGULARIZATION REQUEST FUNCTIONS
// ============================================================================

/**
 * Submit a regularization request for a past date
 * @param {string} attendanceDate - Date to regularize (YYYY-MM-DD format)
 * @param {string} reason - Reason for regularization
 */
function submitRegularizationRequest(attendanceDate, reason) {
  const email = getCurrentUserEmail();
  const employeeData = getEmployeeData();
  
  if (!employeeData) {
    return { success: false, message: 'Employee not found. Please contact HR.' };
  }
  
  // Validate inputs
  if (!attendanceDate || !reason) {
    return { success: false, message: 'Please provide both date and reason.' };
  }
  
  if (reason.trim().length < 10) {
    return { success: false, message: 'Please provide a detailed reason (minimum 10 characters).' };
  }
  
  // Parse and validate date
  const requestedDate = new Date(attendanceDate);
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  requestedDate.setHours(0, 0, 0, 0);
  
  // Check if date is in the past (not today or future)
  if (requestedDate >= today) {
    return { success: false, message: 'You can only request regularization for past dates.' };
  }
  
  // Check if date is within allowed range (last 7 days)
  const maxPastDate = new Date(today);
  maxPastDate.setDate(maxPastDate.getDate() - CONFIG.MAX_REGULARIZATION_DAYS);
  
  if (requestedDate < maxPastDate) {
    return { 
      success: false, 
      message: `You can only request regularization for the last ${CONFIG.MAX_REGULARIZATION_DAYS} days.` 
    };
  }
  
  // Check if attendance already exists for this date
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const monthlySheetName = getMonthlySheetName(requestedDate);
  const monthlySheet = ss.getSheetByName(monthlySheetName);
  
  if (monthlySheet) {
    const attendanceData = monthlySheet.getDataRange().getValues();
    const requestedDateStr = formatDate(requestedDate);
    
    for (let i = 1; i < attendanceData.length; i++) {
      if (attendanceData[i][1] === email && formatDate(attendanceData[i][3]) === requestedDateStr) {
        return { 
          success: false, 
          message: 'Attendance already marked for this date. Regularization not needed.' 
        };
      }
    }
  }
  
  // Check for existing pending request for same date
  const regSheet = getOrCreateSheet(CONFIG.SHEETS.REGULARIZATION_REQUESTS, [
    'Request ID', 'Employee Email', 'Employee Name', 'Request Date',
    'Attendance Date', 'Reason', 'Status', 'Manager Email',
    'Manager Comments', 'Action Date', 'Action By'
  ]);
  
  const regData = regSheet.getDataRange().getValues();
  const requestedDateStr = formatDate(requestedDate);
  
  for (let i = 1; i < regData.length; i++) {
    if (regData[i][1] === email && 
        formatDate(regData[i][4]) === requestedDateStr && 
        regData[i][6] === 'Pending') {
      return { 
        success: false, 
        message: 'You already have a pending regularization request for this date.' 
      };
    }
  }
  
  // Create regularization request
  const requestId = generateRegularizationId();
  const now = new Date();
  
  regSheet.appendRow([
    requestId,
    email,
    employeeData.name,
    now,
    requestedDate,
    reason.trim(),
    'Pending',
    employeeData.managerEmail,
    '',
    '',
    ''
  ]);
  
  // Send notification to manager
  const managerData = getEmployeeDataByEmail(employeeData.managerEmail);
  if (managerData) {
    createNotification(
      employeeData.managerEmail,
      email,
      employeeData.name,
      'regularization_request',
      requestId,
      `${employeeData.name} has requested attendance regularization for ${formatDate(requestedDate)}`
    );
    
    // Send email to manager
    try {
      sendRegularizationRequestEmail(
        employeeData.managerEmail,
        employeeData.name,
        formatDate(requestedDate),
        reason.trim()
      );
    } catch (error) {
      Logger.log('Error sending regularization request email: ' + error);
    }
  }
  
  return {
    success: true,
    message: 'Regularization request submitted successfully. Your manager will review it.',
    requestId: requestId
  };
}

/**
 * Get regularization requests for current employee
 */
function getEmployeeRegularizationRequests() {
  const email = getCurrentUserEmail();
  const sheet = getOrCreateSheet(CONFIG.SHEETS.REGULARIZATION_REQUESTS, [
    'Request ID', 'Employee Email', 'Employee Name', 'Request Date',
    'Attendance Date', 'Reason', 'Status', 'Manager Email',
    'Manager Comments', 'Action Date', 'Action By'
  ]);
  
  const data = sheet.getDataRange().getValues();
  const requests = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === email) {
      requests.push({
        requestId: data[i][0],
        requestDate: formatDate(data[i][3]),
        attendanceDate: formatDate(data[i][4]),
        reason: data[i][5],
        status: data[i][6],
        managerComments: data[i][8] || '',
        actionDate: data[i][9] ? formatDate(data[i][9]) : ''
      });
    }
  }
  
  // Sort by request date (most recent first)
  requests.sort((a, b) => new Date(b.requestDate) - new Date(a.requestDate));
  
  return requests;
}

/**
 * Get pending regularization requests for manager's team
 */
function getManagerRegularizationRequests() {
  const email = getCurrentUserEmail();
  const sheet = getOrCreateSheet(CONFIG.SHEETS.REGULARIZATION_REQUESTS, [
    'Request ID', 'Employee Email', 'Employee Name', 'Request Date',
    'Attendance Date', 'Reason', 'Status', 'Manager Email',
    'Manager Comments', 'Action Date', 'Action By'
  ]);
  
  const data = sheet.getDataRange().getValues();
  const requests = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][7] === email && data[i][6] === 'Pending') {
      requests.push({
        requestId: data[i][0],
        employeeEmail: data[i][1],
        employeeName: data[i][2],
        requestDate: formatDate(data[i][3]),
        attendanceDate: formatDate(data[i][4]),
        reason: data[i][5],
        status: data[i][6]
      });
    }
  }
  
  // Sort by request date (oldest first for priority)
  requests.sort((a, b) => new Date(a.requestDate) - new Date(b.requestDate));
  
  return requests;
}

/**
 * Approve a regularization request
 * @param {string} requestId - Request ID to approve
 * @param {string} managerComments - Optional comments from manager
 */
function approveRegularizationRequest(requestId, managerComments) {
  const managerEmail = getCurrentUserEmail();
  const sheet = getOrCreateSheet(CONFIG.SHEETS.REGULARIZATION_REQUESTS, [
    'Request ID', 'Employee Email', 'Employee Name', 'Request Date',
    'Attendance Date', 'Reason', 'Status', 'Manager Email',
    'Manager Comments', 'Action Date', 'Action By'
  ]);
  
  const data = sheet.getDataRange().getValues();
  let requestRow = -1;
  let requestData = null;
  
  // Find the request
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === requestId) {
      requestRow = i + 1;
      requestData = data[i];
      break;
    }
  }
  
  if (!requestData) {
    return { success: false, message: 'Regularization request not found.' };
  }
  
  // Verify manager
  if (requestData[7] !== managerEmail) {
    return { success: false, message: 'You are not authorized to approve this request.' };
  }
  
  // Check if already processed
  if (requestData[6] !== 'Pending') {
    return { success: false, message: 'This request has already been processed.' };
  }
  
  const lock = acquireLock('regularization_approve_' + requestId, 30, 3);
  
  if (!lock) {
    return { success: false, message: 'System is busy. Please try again.' };
  }
  
  try {
    const now = new Date();
    
    // Update request status
    sheet.getRange(requestRow, 7).setValue('Approved');
    sheet.getRange(requestRow, 9).setValue(managerComments || 'Approved');
    sheet.getRange(requestRow, 10).setValue(now);
    sheet.getRange(requestRow, 11).setValue(managerEmail);
    
    // Create attendance record for the regularized date
    const attendanceDate = new Date(requestData[4]);
    const attendanceSheet = getOrCreateMonthlyAttendanceSheet(attendanceDate);
    const attendanceId = generateAttendanceId();
    
    attendanceSheet.appendRow([
      attendanceId,
      requestData[1], // Employee Email
      requestData[2], // Employee Name
      attendanceDate,
      'Regularized', // Check-In Time
      '', // Check-In Lat
      '', // Check-In Long
      0, // Check-In Distance
      'Regularized', // Check-Out Time
      '', // Check-Out Lat
      '', // Check-Out Long
      0, // Check-Out Distance
      '0', // Total Hours
      'Regularized', // Status
      now // Timestamp
    ]);
    
    // Send notification to employee
    const managerData = getEmployeeData();
    createNotification(
      requestData[1],
      managerEmail,
      managerData.name,
      'regularization_approved',
      requestId,
      `Your regularization request for ${formatDate(attendanceDate)} has been approved`
    );
    
    // Send email to employee
    try {
      sendRegularizationApprovedEmail(
        requestData[1],
        requestData[2],
        formatDate(attendanceDate),
        managerComments || 'Approved'
      );
    } catch (error) {
      Logger.log('Error sending regularization approved email: ' + error);
    }
    
    return {
      success: true,
      message: 'Regularization request approved successfully.'
    };
    
  } finally {
    releaseLock(lock);
  }
}

/**
 * Reject a regularization request
 * @param {string} requestId - Request ID to reject
 * @param {string} managerComments - Reason for rejection (required)
 */
function rejectRegularizationRequest(requestId, managerComments) {
  const managerEmail = getCurrentUserEmail();
  
  if (!managerComments || managerComments.trim().length < 5) {
    return { success: false, message: 'Please provide a reason for rejection.' };
  }
  
  const sheet = getOrCreateSheet(CONFIG.SHEETS.REGULARIZATION_REQUESTS, [
    'Request ID', 'Employee Email', 'Employee Name', 'Request Date',
    'Attendance Date', 'Reason', 'Status', 'Manager Email',
    'Manager Comments', 'Action Date', 'Action By'
  ]);
  
  const data = sheet.getDataRange().getValues();
  let requestRow = -1;
  let requestData = null;
  
  // Find the request
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === requestId) {
      requestRow = i + 1;
      requestData = data[i];
      break;
    }
  }
  
  if (!requestData) {
    return { success: false, message: 'Regularization request not found.' };
  }
  
  // Verify manager
  if (requestData[7] !== managerEmail) {
    return { success: false, message: 'You are not authorized to reject this request.' };
  }
  
  // Check if already processed
  if (requestData[6] !== 'Pending') {
    return { success: false, message: 'This request has already been processed.' };
  }
  
  const lock = acquireLock('regularization_reject_' + requestId, 30, 3);
  
  if (!lock) {
    return { success: false, message: 'System is busy. Please try again.' };
  }
  
  try {
    const now = new Date();
    
    // Update request status
    sheet.getRange(requestRow, 7).setValue('Rejected');
    sheet.getRange(requestRow, 9).setValue(managerComments.trim());
    sheet.getRange(requestRow, 10).setValue(now);
    sheet.getRange(requestRow, 11).setValue(managerEmail);
    
    // Send notification to employee
    const managerData = getEmployeeData();
    const attendanceDate = new Date(requestData[4]);
    
    createNotification(
      requestData[1],
      managerEmail,
      managerData.name,
      'regularization_rejected',
      requestId,
      `Your regularization request for ${formatDate(attendanceDate)} has been rejected`
    );
    
    // Send email to employee
    try {
      sendRegularizationRejectedEmail(
        requestData[1],
        requestData[2],
        formatDate(attendanceDate),
        managerComments.trim()
      );
    } catch (error) {
      Logger.log('Error sending regularization rejected email: ' + error);
    }
    
    return {
      success: true,
      message: 'Regularization request rejected.'
    };
    
  } finally {
    releaseLock(lock);
  }
}

/**
 * Helper function to get employee data by email
 */
function getEmployeeDataByEmail(email) {
  const sheet = getOrCreateSheet(CONFIG.SHEETS.EMPLOYEES, [
    'Employee Email', 'Employee Name', 'Department', 'Manager Email',
    'Date of Birth', 'Total Leaves', 'Leaves Used', 'Leaves Available'
  ]);
  
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      return {
        email: data[i][0],
        name: data[i][1],
        department: data[i][2],
        managerEmail: data[i][3],
        dateOfBirth: data[i][4],
        totalLeaves: data[i][5],
        leavesUsed: data[i][6],
        leavesAvailable: data[i][7]
      };
    }
  }
  
  return null;
}

/**
 * Get leave request details by request ID
 */
function getLeaveRequestDetails(requestId) {
  const sheet = getOrCreateSheet(CONFIG.SHEETS.LEAVE_REQUESTS, [
    'Request ID', 'Employee Email', 'Employee Name', 'Leave Type',
    'Start Date', 'End Date', 'Days', 'Day Type', 'Units', 'Reason', 'Status',
    'Applied Date', 'Manager Email', 'Action Date', 'Manager Comments'
  ]);

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === requestId) {
      return {
        requestId: data[i][0],
        employeeEmail: data[i][1],
        employeeName: data[i][2],
        leaveType: data[i][3],
        startDate: formatDate(data[i][4]),
        endDate: formatDate(data[i][5]),
        days: data[i][6],
        dayType: data[i][7] || 'Full Day',
        units: data[i][8] || data[i][6],
        reason: data[i][9],
        status: data[i][10],
        appliedDate: formatDate(data[i][11]),
        managerEmail: data[i][12],
        actionDate: data[i][13] ? formatDate(data[i][13]) : '',
        managerComments: data[i][14] || ''
      };
    }
  }

  return null;
}


/**
 * Format date as YYYY-MM-DD
 */
function formatDate(date) {
  if (!date) return '';
  const d = new Date(date);
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

/**
 * Format date and time as readable string
 */
function formatDateTime(date) {
  if (!date) return '';
  const d = new Date(date);
  return d.toLocaleString('en-US', {
    year: 'numeric',
    month: 'short',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit'
  });
}

/**
 * Get all company holidays from the sheet
 */
function getCompanyHolidays() {
  const sheet = getOrCreateSheet(CONFIG.SHEETS.COMPANY_HOLIDAYS, [
    'Holiday Date', 'Holiday Name', 'Holiday Type'
  ]);

  const data = sheet.getDataRange().getValues();
  const holidays = [];

  // Skip header row
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      // Store holiday dates as formatted strings for comparison
      const holidayDate = new Date(data[i][0]);
      holidays.push(formatDate(holidayDate));
    }
  }

  return holidays;
}

/**
 * Check if a date is Sunday (weekly off)
 */
function isWeekend(date) {
  const day = date.getDay();
  return day === 0; // Sunday only
}

/**
 * Check if a date is a company holiday
 */
function isCompanyHoliday(date, holidays) {
  const dateStr = formatDate(date);
  return holidays.includes(dateStr);
}

/**
 * Calculate business days between two dates (excluding Sundays and company holidays)
 */
function calculateBusinessDays(startDate, endDate) {
  const holidays = getCompanyHolidays();
  let businessDays = 0;
  
  const currentDate = new Date(startDate);
  const end = new Date(endDate);
  
  // Iterate through each day
  while (currentDate <= end) {
    // Check if it's not a Sunday and not a company holiday
    if (!isWeekend(currentDate) && !isCompanyHoliday(currentDate, holidays)) {
      businessDays++;
    }
    // Move to next day
    currentDate.setDate(currentDate.getDate() + 1);
  }
  
  return businessDays;
}

/**
 * Calculate leave days between two dates
 * This now uses calculateBusinessDays to exclude Sundays and holidays
 */
function calculateLeaveDays(startDate, endDate) {
  return calculateBusinessDays(startDate, endDate);
}

/**
 * Format date for display
 */
function formatDate(date) {
  if (!date) return '';
  const d = new Date(date);
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}



/**
 * Add a company holiday to the master sheet
 */
function addCompanyHoliday(date, name, type) {
  const sheet = getOrCreateSheet(CONFIG.SHEETS.COMPANY_HOLIDAYS, [
    'Holiday Date', 'Holiday Name', 'Holiday Type'
  ]);

  sheet.appendRow([new Date(date), name, type || 'Company']);
  
  return { success: true, message: 'Holiday added successfully' };
}

/**
 * Remove a company holiday from the master sheet
 */
function removeCompanyHoliday(date) {
  const sheet = getOrCreateSheet(CONFIG.SHEETS.COMPANY_HOLIDAYS, [
    'Holiday Date', 'Holiday Name', 'Holiday Type'
  ]);

  const data = sheet.getDataRange().getValues();
  const dateToRemove = formatDate(new Date(date));

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && formatDate(new Date(data[i][0])) === dateToRemove) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'Holiday removed successfully' };
    }
  }

  return { success: false, message: 'Holiday not found' };
}

/**
 * Get all company holidays with full details for calendar display
 */
function getAllCompanyHolidays() {
  const sheet = getOrCreateSheet(CONFIG.SHEETS.COMPANY_HOLIDAYS, [
    'Holiday Date', 'Holiday Name', 'Holiday Type'
  ]);

  const data = sheet.getDataRange().getValues();
  const holidays = [];

  // Skip header row
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      holidays.push({
        date: formatDate(data[i][0]),
        name: data[i][1] || 'Holiday',
        type: data[i][2] || 'Company'
      });
    }
  }

  // Sort by date
  holidays.sort((a, b) => new Date(a.date) - new Date(b.date));

  return holidays;
}

/**
 * Get upcoming holidays in the next N days
 */
function getUpcomingHolidays(days) {
  if (!days) days = 30;
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  const futureDate = new Date(today);
  futureDate.setDate(futureDate.getDate() + days);
  
  const allHolidays = getAllCompanyHolidays();
  const upcomingHolidays = [];
  
  for (let i = 0; i < allHolidays.length; i++) {
    const holidayDate = new Date(allHolidays[i].date);
    holidayDate.setHours(0, 0, 0, 0);
    
    if (holidayDate >= today && holidayDate <= futureDate) {
      upcomingHolidays.push(allHolidays[i]);
    }
  }
  
  return upcomingHolidays;
}

/**
 * Get upcoming special events in the next N days
 */
function getUpcomingSpecialEvents(days) {
  if (!days) days = 30;
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  const futureDate = new Date(today);
  futureDate.setDate(futureDate.getDate() + days);
  
  const sheet = getOrCreateSheet(CONFIG.SHEETS.SPECIAL_EVENTS, [
    'Event Title', 'Event Date', 'Description', 'Type'
  ]);
  
  const data = sheet.getDataRange().getValues();
  const upcomingEvents = [];
  
  // Skip header row
  for (let i = 1; i < data.length; i++) {
    const rawDate = new Date(data[i][1]);
    const eventDate = new Date(rawDate);
    eventDate.setHours(0, 0, 0, 0);
    
    if (eventDate >= today && eventDate <= futureDate) {
      upcomingEvents.push({
        title: data[i][0],
        date: rawDate.toISOString(), // Send full ISO string to frontend
        description: data[i][2],
        type: data[i][3]
      });
    }
  }
  
  // Sort by date
  upcomingEvents.sort((a, b) => new Date(a.date) - new Date(b.date));
  
  return upcomingEvents;
}

/**
 * Check if today is the user's birthday
 */
function isTodayBirthday(dob) {
  if (!dob) return false;
  
  const today = new Date();
  const birthDate = new Date(dob);
  
  return today.getMonth() === birthDate.getMonth() && 
         today.getDate() === birthDate.getDate();
}

/**
 * Get birthday wish data for current user
 */
function getBirthdayWishData() {
  const employeeData = getEmployeeData();
  
  if (!employeeData || !employeeData.dateOfBirth) {
    return { isBirthday: false };
  }
  
  const isBirthday = isTodayBirthday(employeeData.dateOfBirth);
  
  return {
    isBirthday: isBirthday,
    name: employeeData.name,
    email: employeeData.email
  };
}

/**
 * Get upcoming birthdays of all employees in the next N days
 */
function getUpcomingBirthdays(days) {
  if (!days) days = 30;
  
  const sheet = getOrCreateSheet(CONFIG.SHEETS.EMPLOYEES, [
    'Employee Email', 'Employee Name', 'Department', 'Manager Email',
    'Date of Birth', 'Total Leaves', 'Leaves Used', 'Leaves Available'
  ]);

  const data = sheet.getDataRange().getValues();
  const upcomingBirthdays = [];
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  // Skip header row
  for (let i = 1; i < data.length; i++) {
    const dob = data[i][4];
    if (!dob) continue;
    
    const birthDate = new Date(dob);
    const name = data[i][1];
    const department = data[i][2] || 'Not Specified';
    
    // Calculate next birthday
    const nextBirthday = new Date(today.getFullYear(), birthDate.getMonth(), birthDate.getDate());
    nextBirthday.setHours(0, 0, 0, 0);
    
    // If birthday already passed this year, check next year
    if (nextBirthday < today) {
      nextBirthday.setFullYear(today.getFullYear() + 1);
    }
    
    // Calculate days until birthday
    const daysUntil = Math.floor((nextBirthday - today) / (1000 * 60 * 60 * 24));
    
    if (daysUntil >= 0 && daysUntil <= days) {
      upcomingBirthdays.push({
        name: name,
        department: department,
        date: formatDate(nextBirthday),
        daysUntil: daysUntil
      });
    }
  }
  
  // Sort by days until birthday
  upcomingBirthdays.sort((a, b) => a.daysUntil - b.daysUntil);
  
  return upcomingBirthdays;
}


/**
 * Setup monthly trigger for leave accrual
 * Run this once to set up automatic monthly leave accrual
 */
function setupMonthlyTrigger() {
  // Delete existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'accrueMonthlyLeaves') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create new trigger for 1st of every month at 1 AM
  ScriptApp.newTrigger('accrueMonthlyLeaves')
    .timeBased()
    .onMonthDay(1)
    .atHour(1)
    .create();

  Logger.log('Monthly trigger set up successfully');
}


// ============================================================================
// ATTENDANCE SYSTEM FUNCTIONS
// ============================================================================

// ============================================================================
// MONTHLY ATTENDANCE SHEET MANAGEMENT
// ============================================================================

/**
 * Get sheet name for a given date
 * @param {Date} date - Date to format
 * @returns {string} Sheet name (e.g., "Attendance Dec-25")
 */
function getMonthlySheetName(date) {
  return 'Attendance ' + Utilities.formatDate(date, Session.getScriptTimeZone(), 'MMM-yy');
}

/**
 * Get or create monthly attendance sheet with lock protection
 * @param {Date} date - Date to determine the month
 * @returns {Sheet} Monthly attendance sheet
 */
function getOrCreateMonthlyAttendanceSheet(date) {
  const sheetName = getMonthlySheetName(date);
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  let sheet = ss.getSheetByName(sheetName);
  
  // If sheet exists, return it immediately
  if (sheet) {
    return sheet;
  }
  
  // Need to create sheet - acquire lock to prevent duplicate creation
  const lock = acquireLock('sheet_create_' + sheetName, 10, 2);
  
  try {
    // Double-check if sheet was created while waiting for lock
    sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      return sheet;
    }
    
    // Create new sheet with headers (including Work Mode, Shift, and Late Status columns)
    const headers = [
      'Attendance ID', 'Work Mode', 'Employee Email', 'Employee Name', 'Date',
      'Check-In Time', 'Check-In Lat', 'Check-In Long', 'Check-In Distance',
      'Check-Out Time', 'Check-Out Lat', 'Check-Out Long', 'Check-Out Distance',
      'Total Hours', 'Status', 'Shift Name', 'Late Status', 'Late Minutes', 'Timestamp'
    ];
    
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    
    Logger.log(`Created new monthly attendance sheet: ${sheetName}`);
    
    return sheet;
  } finally {
    releaseLock(lock);
  }
}

/**
 * Get all monthly attendance sheet names
 * @returns {Array<string>} List of all monthly attendance sheet names
 */
function getAllMonthlyAttendanceSheets() {
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const sheets = ss.getSheets();
  const monthlySheets = [];
  
  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (name.startsWith('Attendance ') && name !== CONFIG.SHEETS.ATTENDANCE_RECORDS) {
      monthlySheets.push(name);
    }
  });
  
  return monthlySheets;
}


/**
 * Get attendance configuration (office location and radius)
 */
function getAttendanceConfig() {
  const sheet = getOrCreateSheet(CONFIG.SHEETS.ATTENDANCE_CONFIG, [
    'Office Name', 'Latitude', 'Longitude', 'Allowed Radius (meters)', 'Last Updated'
  ]);

  const data = sheet.getDataRange().getValues();
  
  // Return first office configuration (row 2, index 1)
  if (data.length > 1) {
    return {
      officeName: data[1][0],
      latitude: parseFloat(data[1][1]),
      longitude: parseFloat(data[1][2]),
      radius: parseFloat(data[1][3])
    };
  }
  
  // Return default if no config found
  return {
    officeName: CONFIG.DEFAULT_OFFICE.NAME,
    latitude: CONFIG.DEFAULT_OFFICE.LATITUDE,
    longitude: CONFIG.DEFAULT_OFFICE.LONGITUDE,
    radius: CONFIG.DEFAULT_OFFICE.RADIUS
  };
}

// ============================================================================
// SHIFT MANAGEMENT FUNCTIONS
// ============================================================================

/**
 * Get shift configuration by shift name
 * @param {string} shiftName - Name of the shift
 * @returns {object} Shift configuration with startTime, endTime, gracePeriod
 */
function getShiftConfiguration(shiftName) {
  const sheet = getOrCreateSheet(CONFIG.SHEETS.SHIFT_CONFIG, [
    'Shift Name', 'Start Time', 'End Time', 'Grace Period (minutes)', 'Description'
  ]);
  
  const data = sheet.getDataRange().getValues();
  
  // Find shift configuration
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === shiftName) {
      return {
        shiftName: data[i][0],
        startTime: data[i][1], // Format: "HH:mm"
        endTime: data[i][2],   // Format: "HH:mm"
        gracePeriod: parseInt(data[i][3]) || 15, // Default 15 minutes
        description: data[i][4] || ''
      };
    }
  }
  
  // Return default shift if not found
  return {
    shiftName: CONFIG.DEFAULT_SHIFT_NAME,
    startTime: '09:00',
    endTime: '18:00',
    gracePeriod: 15,
    description: 'Default shift'
  };
}

/**
 * Check if check-in time is late based on shift timing
 * @param {string} checkInTime - Check-in time in "HH:mm:ss" format
 * @param {string} shiftName - Name of the shift
 * @returns {object} { isLate: boolean, lateMinutes: number, shiftStartTime: string }
 */
function isLateCheckIn(checkInTime, shiftName) {
  const shiftConfig = getShiftConfiguration(shiftName);
  
  // Parse check-in time (format: "HH:mm:ss")
  const checkInParts = checkInTime.split(':');
  const checkInHour = parseInt(checkInParts[0]);
  const checkInMinute = parseInt(checkInParts[1]);
  const checkInTotalMinutes = checkInHour * 60 + checkInMinute;
  
  // Parse shift start time - handle both string and Date object
  let shiftStartTime = shiftConfig.startTime;
  
  // If it's a Date object, convert to HH:mm format
  if (shiftStartTime instanceof Date) {
    shiftStartTime = Utilities.formatDate(shiftStartTime, Session.getScriptTimeZone(), 'HH:mm');
  }
  
  // Now parse the time string
  const shiftParts = shiftStartTime.toString().split(':');
  const shiftHour = parseInt(shiftParts[0]);
  const shiftMinute = parseInt(shiftParts[1]);
  const shiftTotalMinutes = shiftHour * 60 + shiftMinute;
  
  // Calculate allowed check-in time (shift start + grace period)
  const allowedTotalMinutes = shiftTotalMinutes + shiftConfig.gracePeriod;
  
  // Calculate late minutes
  const lateMinutes = checkInTotalMinutes - allowedTotalMinutes;
  
  return {
    isLate: lateMinutes > 0,
    lateMinutes: Math.max(0, lateMinutes),
    shiftStartTime: shiftStartTime,
    gracePeriod: shiftConfig.gracePeriod
  };
}

/**
 * Update employee shift assignment
 * @param {string} employeeEmail - Employee email
 * @param {string} newShiftName - New shift name
 * @returns {object} Success/failure response
 */
function updateEmployeeShift(employeeEmail, newShiftName) {
  // Validate shift exists
  const shiftConfig = getShiftConfiguration(newShiftName);
  if (!shiftConfig || shiftConfig.shiftName !== newShiftName) {
    return { success: false, message: 'Invalid shift name' };
  }
  
  const sheet = getOrCreateSheet(CONFIG.SHEETS.EMPLOYEES, [
    'Employee Email', 'Employee Name', 'Department', 'Manager Email',
    'Date of Birth', 'Total Leaves', 'Leaves Used', 'Leaves Available', 'Shift Name'
  ]);
  
  const data = sheet.getDataRange().getValues();
  
  // Find and update employee record
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === employeeEmail) {
      sheet.getRange(i + 1, 9).setValue(newShiftName); // Column 9 is Shift Name
      return { 
        success: true, 
        message: `Shift updated to ${newShiftName} for ${data[i][1]}` 
      };
    }
  }
  
  return { success: false, message: 'Employee not found' };
}

/**
 * Get all available shifts
 * @returns {array} List of all shift configurations
 */
function getAllShifts() {
  const sheet = getOrCreateSheet(CONFIG.SHEETS.SHIFT_CONFIG, [
    'Shift Name', 'Start Time', 'End Time', 'Grace Period (minutes)', 'Description'
  ]);
  
  const data = sheet.getDataRange().getValues();
  const shifts = [];
  
  for (let i = 1; i < data.length; i++) {
    shifts.push({
      shiftName: data[i][0],
      startTime: data[i][1],
      endTime: data[i][2],
      gracePeriod: data[i][3],
      description: data[i][4]
    });
  }
  
  return shifts;
}

/**
 * Update attendance configuration (Admin function)
 */
function updateAttendanceConfig(latitude, longitude, radius, officeName) {
  const sheet = getOrCreateSheet(CONFIG.SHEETS.ATTENDANCE_CONFIG, [
    'Office Name', 'Latitude', 'Longitude', 'Allowed Radius (meters)', 'Last Updated'
  ]);

  const data = sheet.getDataRange().getValues();
  
  if (data.length > 1) {
    // Update existing config
    sheet.getRange(2, 1).setValue(officeName || 'Head Office');
    sheet.getRange(2, 2).setValue(parseFloat(latitude));
    sheet.getRange(2, 3).setValue(parseFloat(longitude));
    sheet.getRange(2, 4).setValue(parseFloat(radius));
    sheet.getRange(2, 5).setValue(new Date());
  } else {
    // Add new config
    sheet.appendRow([
      officeName || 'Head Office',
      parseFloat(latitude),
      parseFloat(longitude),
      parseFloat(radius),
      new Date()
    ]);
  }
  
  return { success: true, message: 'Attendance configuration updated successfully' };
}

/**
 * Calculate distance between two coordinates using Haversine formula
 * Returns distance in meters
 */
function calculateHaversineDistance(lat1, lon1, lat2, lon2) {
  const R = 6371000; // Earth's radius in meters
  
  // Convert degrees to radians
  const φ1 = lat1 * Math.PI / 180;
  const φ2 = lat2 * Math.PI / 180;
  const Δφ = (lat2 - lat1) * Math.PI / 180;
  const Δλ = (lon2 - lon1) * Math.PI / 180;
  
  // Haversine formula
  const a = Math.sin(Δφ / 2) * Math.sin(Δφ / 2) +
            Math.cos(φ1) * Math.cos(φ2) *
            Math.sin(Δλ / 2) * Math.sin(Δλ / 2);
  
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  
  const distance = R * c; // Distance in meters
  
  return Math.round(distance); // Round to nearest meter
}

/**
 * Generate unique attendance ID
 */
function generateAttendanceId() {
  return 'ATT' + new Date().getTime();
}

/**
 * Mark attendance for current user (Check-In or Check-Out)
 * @param {number} latitude - User's current latitude
 * @param {number} longitude - User's current longitude
 * @param {string} action - 'checkin' or 'checkout'
 * @param {string} workMode - 'Office', 'Work From Home', or 'On-Duty' (default: 'Office')
 */
function markAttendance(latitude, longitude, action, workMode) {
  // Default to Office mode if not specified (backward compatibility)
  if (!workMode) {
    workMode = 'Office';
  }
  const email = getCurrentUserEmail();
  const employeeData = getEmployeeData();
  
  if (!employeeData) {
    return { success: false, message: 'Employee not found. Please contact HR.' };
  }
  
  // Get office configuration
  const officeConfig = getAttendanceConfig();
  
  // Calculate distance from office
  const distance = calculateHaversineDistance(
    latitude,
    longitude,
    officeConfig.latitude,
    officeConfig.longitude
  );
  
  // Check if within allowed radius (only for Office mode)
  const isWithinRadius = distance <= officeConfig.radius;
  
  // GPS radius validation only applies to Office mode
  // WFH and On-Duty modes skip radius check but still capture location
  if (workMode === 'Office' && !isWithinRadius) {
    return {
      success: false,
      message: `You are ${distance} meters away from office. Attendance can only be marked within ${officeConfig.radius} meters.`,
      distance: distance,
      allowedRadius: officeConfig.radius,
      status: 'Outside Location'
    };
  }
  
  // Acquire lock for concurrent operations
  const lock = acquireLock('attendance_mark_' + email, 30, 3);
  
  if (!lock) {
    return {
      success: false,
      message: 'System is busy processing multiple requests. Please try again in a moment.'
    };
  }
  
  try {
    // Get today's date and current time
    const today = formatDate(new Date());
    const now = new Date();
    const time = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');
    
    // Use monthly attendance sheet
    const attendanceSheet = getOrCreateMonthlyAttendanceSheet(now);
    
    const existingRecords = attendanceSheet.getDataRange().getValues();
    let todayRecordRow = -1;
    let todayRecord = null;
    
    // Find today's record (column indices shifted by 1 due to Work Mode column)
    for (let i = 1; i < existingRecords.length; i++) {
      if (existingRecords[i][2] === email && formatDate(existingRecords[i][4]) === today) {
        todayRecordRow = i + 1; // 1-indexed for sheet
        todayRecord = existingRecords[i];
        break;
      }
    }
    
    // Handle CHECK-IN
    if (action === 'checkin') {
      if (todayRecord && todayRecord[5]) { // Check-In Time exists (column index shifted)
        return {
          success: false,
          message: 'You have already checked in today',
          alreadyMarked: true,
          time: todayRecord[5],
          action: 'checkin'
        };
      }
      
      // Get employee shift and check for late arrival
      const shiftName = employeeData.shiftName || CONFIG.DEFAULT_SHIFT_NAME;
      const lateCheck = isLateCheckIn(time, shiftName);
      const lateStatus = lateCheck.isLate ? 'Late' : 'On Time';
      const lateMinutes = lateCheck.lateMinutes;
      
      if (todayRecord) {
        // Update existing record with check-in (column indices: Work Mode=2, Check-In Time=6, etc.)
        attendanceSheet.getRange(todayRecordRow, 2).setValue(workMode); // Work Mode
        attendanceSheet.getRange(todayRecordRow, 6).setValue(time); // Check-In Time
        attendanceSheet.getRange(todayRecordRow, 7).setValue(latitude); // Check-In Lat
        attendanceSheet.getRange(todayRecordRow, 8).setValue(longitude); // Check-In Long
        attendanceSheet.getRange(todayRecordRow, 9).setValue(distance); // Check-In Distance
        attendanceSheet.getRange(todayRecordRow, 15).setValue('Checked In'); // Status
        attendanceSheet.getRange(todayRecordRow, 16).setValue(shiftName); // Shift Name
        attendanceSheet.getRange(todayRecordRow, 17).setValue(lateStatus); // Late Status
        attendanceSheet.getRange(todayRecordRow, 18).setValue(lateMinutes); // Late Minutes
      } else {
        // Create new record (with Work Mode, Shift, and Late Status columns)
        const attendanceId = generateAttendanceId();
        attendanceSheet.appendRow([
          attendanceId,
          workMode, // Work Mode
          email,
          employeeData.name,
          now,
          time, // Check-In Time
          latitude, // Check-In Lat
          longitude, // Check-In Long
          distance, // Check-In Distance
          '', // Check-Out Time
          '', // Check-Out Lat
          '', // Check-Out Long
          '', // Check-Out Distance
          '', // Total Hours
          'Checked In', // Status
          shiftName, // Shift Name
          lateStatus, // Late Status
          lateMinutes, // Late Minutes
          now // Timestamp
        ]);
      }
      
      return {
        success: true,
        message: `Checked in successfully at ${time} (${workMode})`,
        time: time,
        distance: distance,
        workMode: workMode,
        status: 'Checked In',
        action: 'checkin',
        shiftName: shiftName,
        lateStatus: lateStatus,
        lateMinutes: lateMinutes
      };
    }
    
    // Handle CHECK-OUT
    if (action === 'checkout') {
      if (!todayRecord || !todayRecord[5]) { // No check-in time (column index shifted)
        return {
          success: false,
          message: 'You must check in first before checking out',
          action: 'checkout'
        };
      }
      
      if (todayRecord[9]) { // Check-Out Time exists (column index shifted)
        return {
          success: false,
          message: 'You have already checked out today',
          alreadyMarked: true,
          time: todayRecord[9],
          action: 'checkout'
        };
      }
      
      // Calculate working hours (column indices shifted)
      const checkInTime = new Date(todayRecord[4].toDateString() + ' ' + todayRecord[5]);
      const checkOutTime = new Date(now.toDateString() + ' ' + time);
      const diffMs = checkOutTime - checkInTime;
      const totalHours = (diffMs / (1000 * 60 * 60)).toFixed(2);
      
      // Update record with check-out (column indices shifted by 1)
      attendanceSheet.getRange(todayRecordRow, 10).setValue(time); // Check-Out Time
      attendanceSheet.getRange(todayRecordRow, 11).setValue(latitude); // Check-Out Lat
      attendanceSheet.getRange(todayRecordRow, 12).setValue(longitude); // Check-Out Long
      attendanceSheet.getRange(todayRecordRow, 13).setValue(distance); // Check-Out Distance
      attendanceSheet.getRange(todayRecordRow, 14).setValue(totalHours); // Total Hours
      attendanceSheet.getRange(todayRecordRow, 15).setValue('Checked Out'); // Status
      
      return {
        success: true,
        message: `Checked out successfully at ${time} (${workMode})`,
        time: time,
        distance: distance,
        totalHours: totalHours,
        workMode: workMode,
        status: 'Checked Out',
        action: 'checkout'
      };
    }
    
    return { success: false, message: 'Invalid action' };
    
  } finally {
    releaseLock(lock);
  }
}

/**
 * Get attendance records for current user
 * @param {number} days - Number of days to retrieve (optional)
 * @param {number} month - Month to filter (0-11, optional)
 * @param {number} year - Year to filter (optional)
 */
function getEmployeeAttendance(days, month, year) {
  const email = getCurrentUserEmail();
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  
  Logger.log('=== getEmployeeAttendance START ===');
  Logger.log('Input params: days=' + days + ', month=' + month + ', year=' + year);
  Logger.log('Current user email: ' + email);
  
  let sheetsToRead = [];
  
  // Determine which sheets to read based on filter
  if (month !== null && month !== undefined && year !== null && year !== undefined) {
    // Specific month/year requested - read that monthly sheet
    const filterDate = new Date(year, month, 1);
    const monthlySheetName = getMonthlySheetName(filterDate);
    const monthlySheet = ss.getSheetByName(monthlySheetName);
    
    if (monthlySheet) {
      sheetsToRead.push(monthlySheet);
      Logger.log('Reading from monthly sheet: ' + monthlySheetName);
    }
    
    // Also check legacy sheet for historical data
    const legacySheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE_RECORDS);
    if (legacySheet) {
      sheetsToRead.push(legacySheet);
      Logger.log('Also reading from legacy sheet: ' + CONFIG.SHEETS.ATTENDANCE_RECORDS);
    }
  } else {
    // No specific month - read current month and previous month (for last N days)
    const today = new Date();
    const currentMonthSheet = ss.getSheetByName(getMonthlySheetName(today));
    
    if (currentMonthSheet) {
      sheetsToRead.push(currentMonthSheet);
      Logger.log('Reading from current month sheet: ' + getMonthlySheetName(today));
    }
    
    // Also read previous month if needed
    const previousMonth = new Date(today);
    previousMonth.setMonth(previousMonth.getMonth() - 1);
    const previousMonthSheet = ss.getSheetByName(getMonthlySheetName(previousMonth));
    
    if (previousMonthSheet) {
      sheetsToRead.push(previousMonthSheet);
      Logger.log('Reading from previous month sheet: ' + getMonthlySheetName(previousMonth));
    }
    
    // Also check legacy sheet for historical data
    const legacySheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE_RECORDS);
    if (legacySheet) {
      sheetsToRead.push(legacySheet);
      Logger.log('Also reading from legacy sheet: ' + CONFIG.SHEETS.ATTENDANCE_RECORDS);
    }
  }
  
  const records = [];
  
  // Determine filtering mode
  let filterMonth = null;
  let filterYear = null;
  
  if (month !== undefined && month !== null && year !== undefined && year !== null) {
    filterMonth = Number(month);
    filterYear = Number(year);
    Logger.log('Filter mode: MONTH/YEAR - filterMonth=' + filterMonth + ', filterYear=' + filterYear);
  }
  
  const useMonthYearFilter = (filterMonth !== null && filterYear !== null);
  
  let cutoffDate;
  if (!useMonthYearFilter) {
    if (!days) days = 30;
    cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - days);
    Logger.log('Filter mode: DAYS - days=' + days + ', cutoffDate=' + cutoffDate);
  }
  
  // Read from all identified sheets
  sheetsToRead.forEach(sheet => {
    const data = sheet.getDataRange().getValues();
    Logger.log('Reading sheet: ' + sheet.getName() + ', rows: ' + data.length);
    
    for (let i = 1; i < data.length; i++) {
      const rowEmail = data[i][2]; // Email is now at index 2 (after Work Mode)
      const rowDate = data[i][4];  // Date is now at index 4

      // Check email (case-insensitive) and date existence
      if (rowEmail && String(rowEmail).toLowerCase() === email.toLowerCase() && rowDate) {
        const recordDate = new Date(rowDate);
        
        // Skip invalid dates
        if (isNaN(recordDate.getTime())) {
          continue;
        }
        
        let includeRecord = false;
        
        if (useMonthYearFilter) {
          // Filter by specific month and year
          const rMonth = recordDate.getMonth();
          const rYear = recordDate.getFullYear();
          
          if (rMonth === filterMonth && rYear === filterYear) {
            includeRecord = true;
          }
        } else {
          // Filter by days
          if (recordDate >= cutoffDate) {
            includeRecord = true;
          }
        }
        
        if (includeRecord) {
          // Format times properly
          let checkInTime = '';
          let checkOutTime = '';
          let totalHours = '';
          
          // Convert check-in time (now at index 5)
          if (data[i][5]) {
            if (typeof data[i][5] === 'string') {
              checkInTime = data[i][5];
            } else {
              const timeDate = new Date(data[i][5]);
              checkInTime = Utilities.formatDate(timeDate, Session.getScriptTimeZone(), 'HH:mm:ss');
            }
          }
          
          // Convert check-out time (now at index 9)
          if (data[i][9]) {
            if (typeof data[i][9] === 'string') {
              checkOutTime = data[i][9];
            } else {
              const timeDate = new Date(data[i][9]);
              checkOutTime = Utilities.formatDate(timeDate, Session.getScriptTimeZone(), 'HH:mm:ss');
            }
          }
          
          // Handle totalHours (now at index 13)
          if (data[i][13]) {
            const hours = parseFloat(data[i][13]);
            totalHours = isNaN(hours) ? '' : hours.toString();
          }
          
          records.push({
            attendanceId: data[i][0],
            workMode: data[i][1] || 'Office', // Work Mode (default to Office for old records)
            date: formatDate(data[i][4]),
            checkInTime: checkInTime,
            checkOutTime: checkOutTime,
            totalHours: totalHours,
            status: data[i][14], // Status is now at index 14
            shiftName: data[i][15] || CONFIG.DEFAULT_SHIFT_NAME, // Shift Name at index 15
            lateStatus: data[i][16] || 'On Time', // Late Status at index 16
            lateMinutes: data[i][17] || 0 // Late Minutes at index 17
          });
        }
      }
    }
  });
  
  Logger.log('Total records found: ' + records.length);
  
  const result = records.reverse(); // Most recent first
  
  Logger.log('=== getEmployeeAttendance END ===');
  
  return result;
}

/**
 * Get attendance statistics for current user
 * @param {number} days - Number of days (optional)
 * @param {number} month - Month to filter (0-11, optional)
 * @param {number} year - Year to filter (optional)
 */
function getAttendanceStats(days, month, year) {
  const records = getEmployeeAttendance(days, month, year);
  
  // Count present days (records with check-in)
  const presentDays = records.filter(r => r.checkInTime !== '').length;
  const totalDays = records.length;
  const percentage = totalDays > 0 ? Math.round((presentDays / totalDays) * 100) : 0;
  
  return {
    presentDays: presentDays,
    totalDays: totalDays,
    percentage: percentage
  };
}

/**
 * Get all attendance records for a specific date (Admin/HR function)
 */
function getAllAttendanceRecords(date) {
  const targetDate = new Date(date);
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  
  // Get the monthly sheet for the target date
  const monthlySheetName = getMonthlySheetName(targetDate);
  let sheet = ss.getSheetByName(monthlySheetName);
  
  // If monthly sheet doesn't exist, try legacy sheet
  if (!sheet) {
    sheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE_RECORDS);
  }
  
  if (!sheet) {
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  const records = [];
  const targetDateStr = formatDate(targetDate);
  
  for (let i = 1; i < data.length; i++) {
    if (formatDate(data[i][3]) === targetDateStr) {
      records.push({
        attendanceId: data[i][0],
        employeeEmail: data[i][1],
        employeeName: data[i][2],
        date: formatDate(data[i][3]),
        checkInTime: data[i][4] || '',
        checkOutTime: data[i][8] || '',
        totalHours: data[i][12] || '',
        status: data[i][13]
      });
    }
  }
  
  return records;
}

/**
 * Check if attendance is marked for today
 */
function isTodayAttendanceMarked() {
  const email = getCurrentUserEmail();
  const today = formatDate(new Date());
  const now = new Date();
  
  // Use current month's sheet
  const sheet = getOrCreateMonthlyAttendanceSheet(now);
  
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    // Email is now at index 2, Date at index 4 (after Work Mode column)
    if (data[i][2] === email && formatDate(data[i][4]) === today) {
      return {
        marked: true,
        checkedIn: data[i][5] ? true : false,      // Check-In Time at index 5
        checkedOut: data[i][9] ? true : false,     // Check-Out Time at index 9
        checkInTime: data[i][5] || '',
        checkOutTime: data[i][9] || '',
        totalHours: data[i][13] || '',             // Total Hours at index 13
        status: data[i][14],                       // Status at index 14
        workMode: data[i][1] || 'Office',          // Work Mode at index 1
        checkInDistance: data[i][8] || 0,          // Check-In Distance at index 8
        checkOutDistance: data[i][12] || 0         // Check-Out Distance at index 12
      };
    }
  }
  
  return { marked: false, checkedIn: false, checkedOut: false };
}

/**
 * Get current month attendance statistics for current user
 */
function getCurrentMonthAttendanceStats() {
  const email = getCurrentUserEmail();
  const now = new Date();
  const currentMonth = now.getMonth();
  const currentYear = now.getFullYear();
  
  // Use current month's sheet
  const sheet = getOrCreateMonthlyAttendanceSheet(now);
  
  const data = sheet.getDataRange().getValues();
  
  // Count present days (days with check-in) in current month
  let presentDays = 0;
  
  for (let i = 1; i < data.length; i++) {
    // Email at index 2, Date at index 4, Check-in at index 5
    if (data[i][2] === email && data[i][4]) {
      const recordDate = new Date(data[i][4]);
      if (recordDate.getMonth() === currentMonth && 
          recordDate.getFullYear() === currentYear &&
          data[i][5]) { // Has check-in time (or 'Regularized')
        presentDays++;
      }
    }
  }
  
  // Calculate total working days in current month (excluding Sundays and holidays)
  const firstDay = new Date(currentYear, currentMonth, 1);
  const lastDay = new Date(currentYear, currentMonth + 1, 0);
  const totalWorkingDays = calculateBusinessDays(firstDay, lastDay);
  
  // Calculate absent days
  const absentDays = Math.max(0, totalWorkingDays - presentDays);
  
  return {
    presentDays: presentDays,
    absentDays: absentDays,
    totalWorkingDays: totalWorkingDays,
    month: now.toLocaleDateString('en-US', { month: 'long' }),
    year: currentYear
  };
}

/**
0 * Get birthday wish data for the current user
 -+* Returns user info if today is their birthday
 */
function getBirthdayWishData() {
  const email = getCurrentUserEmail();
  const sheet = getOrCreateSheet(CONFIG.SHEETS.EMPLOYEES, [
    'Employee Email', 'Employee Name', 'Department', 'Manager Email',
    'Date of Birth', 'Total Leaves', 'Leaves Used', 'Leaves Available'
  ]);
  
  const data = sheet.getDataRange().getValues();
  
  // Find the employee
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      const dob = data[i][4];
      
      if (!dob) {
        return { isBirthday: false };
      }
      
      // Check if today is the birthday
      const today = new Date();
      const birthDate = new Date(dob);
      
      if (today.getMonth() === birthDate.getMonth() && today.getDate() === birthDate.getDate()) {
        return {
          isBirthday: true,
          name: data[i][1]
        };
      }
      
      return { isBirthday: false };
    }
  }
  
  return { isBirthday: false };
}

// ============================================================================
// SALARY DETAILS FUNCTIONS
// ============================================================================

/**
 * Get company details for PDF header
 */
function getCompanyDetails() {
  return {
    name: 'KarmSarthi',
    tagline: 'Har din, har chhutti ka bharosa',
    logoUrl: 'https://media.licdn.com/dms/image/v2/D5622AQF9EXLdfJsDAg/feedshare-shrink_1280/B56ZtVIUq9HkAs-/0/1766659804554?e=1768435200&v=beta&t=O62do53dbIFIIGlNq4Ov5n6VAPnmkwDfeJX-Ylg0QXw'
  };
}

/**
 * Get salary details for current user for a specific month and year
 * @param {number} month - Month (1-12)
 * @param {number} year - Year (e.g., 2025)
 */
function getSalaryDetails(month, year) {
  const email = getCurrentUserEmail().toLowerCase();
  
  // Get employee basic info
  const employeeData = getEmployeeData();
  if (!employeeData) {
    return { success: false, message: 'Employee not found' };
  }
  
  // Get employee details
  const employeeDetails = getEmployeeDetails();
  
  // Get salary details from Salary Details sheet
  const salarySheet = getOrCreateSheet(CONFIG.SHEETS.SALARY_DETAILS, [
    'Employee Email', 'Month', 'Year', 'Basic Salary', 'HRA',
    'Conveyance Allowance', 'Medical Allowance', 'Special Allowance',
    'Other Earnings', 'Gross Salary', 'PF Deduction', 'ESI Deduction',
    'Professional Tax', 'TDS', 'Other Deductions', 'Total Deductions',
    'Net Salary', 'Paid Days', 'Entry Date', 'Updated By'
  ]);
  
  const data = salarySheet.getDataRange().getValues();
  const headers = data[0];
  
  // Map headers to indices for robust access
  const col = {};
  headers.forEach((h, i) => col[h] = i);
  
  // Helper to safely get value by header name
  const getValue = (row, header) => {
    const idx = col[header];
    return idx !== undefined && idx < row.length ? row[idx] : 0;
  };

  // Find salary record for this employee, month, and year
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowEmail = String(row[col['Employee Email']]).toLowerCase();
    const rowMonth = row[col['Month']];
    const rowYear = row[col['Year']];
    
    // Loose comparison for month/year
    if (rowEmail === email && 
        rowMonth == month && 
        rowYear == year) {
      
      const paidDaysIdx = col['Paid Days'];
      const paidDays = paidDaysIdx !== undefined ? row[paidDaysIdx] : getDaysInMonth(month, year);
      
      return {
        success: true,
        employeeInfo: {
          name: employeeData.name,
          email: email,
          employeeId: employeeDetails.employeeId || 'N/A',
          designation: employeeDetails.designation || 'N/A',
          department: employeeData.department || 'N/A',
          dateOfJoining: employeeDetails.dateOfJoining || 'N/A'
        },
        salaryBreakdown: {
          earnings: {
            basicSalary: getValue(row, 'Basic Salary'),
            hra: getValue(row, 'HRA'),
            conveyanceAllowance: getValue(row, 'Conveyance Allowance'),
            medicalAllowance: getValue(row, 'Medical Allowance'),
            specialAllowance: getValue(row, 'Special Allowance'),
            otherEarnings: getValue(row, 'Other Earnings'),
            grossSalary: getValue(row, 'Gross Salary')
          },
          deductions: {
            pfDeduction: getValue(row, 'PF Deduction'),
            esiDeduction: getValue(row, 'ESI Deduction'),
            professionalTax: getValue(row, 'Professional Tax'),
            tds: getValue(row, 'TDS'),
            otherDeductions: getValue(row, 'Other Deductions'),
            totalDeductions: getValue(row, 'Total Deductions')
          },
          netSalary: getValue(row, 'Net Salary'),
          paidDays: paidDays || getDaysInMonth(month, year) // Fallback if column exists but value empty
        },
        month: month,
        year: year,
        monthName: getMonthName(month)
      };
    }
  }
  
  // No salary record found for this month
  return {
    success: false,
    message: 'No salary record found for the selected month'
  };
}

/**
 * Get list of months for which salary data exists for current user
 */
function getAvailableMonthsForSalary() {
  const email = getCurrentUserEmail().toLowerCase();
  
  const salarySheet = getOrCreateSheet(CONFIG.SHEETS.SALARY_DETAILS, [
    'Employee Email', 'Month', 'Year', 'Basic Salary', 'HRA',
    'Conveyance Allowance', 'Medical Allowance', 'Special Allowance',
    'Other Earnings', 'Gross Salary', 'PF Deduction', 'ESI Deduction',
    'Professional Tax', 'TDS', 'Other Deductions', 'Total Deductions',
    'Net Salary', 'Entry Date', 'Updated By'
  ]);
  
  const salaryData = salarySheet.getDataRange().getValues();
  const availableMonths = [];
  
  // Find all salary records for this employee
  for (let i = 1; i < salaryData.length; i++) {
    const rowEmail = String(salaryData[i][0]).toLowerCase();
    
    if (rowEmail === email) {
      const monthVal = salaryData[i][1];
      const yearVal = salaryData[i][2];
      
      availableMonths.push({
        month: monthVal,
        year: yearVal,
        monthName: getMonthName(monthVal)
      });
    }
  }
  
  // Sort by year and month (most recent first)
  availableMonths.sort((a, b) => {
    if (b.year !== a.year) {
      return b.year - a.year;
    }
    // Handle month sorting (assuming months might be strings or numbers)
    const monthA = getMonthNumber(a.month);
    const monthB = getMonthNumber(b.month);
    return monthB - monthA;
  });
  
  console.log('Available months for ' + email + ': ' + JSON.stringify(availableMonths));
  return availableMonths;
}

/**
 * Helper function to get month name from month number or name
 * @param {number|string} month - Month (1-12 or "January")
 */
function getMonthName(month) {
  const monthNames = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];
  
  // If it's a number (1-12)
  if (!isNaN(month) && month >= 1 && month <= 12) {
    return monthNames[month - 1];
  }
  
  // If it's already a string, return it capitalized
  if (typeof month === 'string') {
    // Check if it's a number string "1"
    const num = parseInt(month);
    if (!isNaN(num) && num >= 1 && num <= 12) {
      return monthNames[num - 1];
    }
    return month;
  }
  
  return '';
}

/**
 * Helper to get numeric month for sorting
 */
function getMonthNumber(month) {
  if (!isNaN(month)) return parseInt(month);
  
  const monthNames = [
    'january', 'february', 'march', 'april', 'may', 'june',
    'july', 'august', 'september', 'october', 'november', 'december'
  ];
  
  const index = monthNames.indexOf(String(month).toLowerCase());
  return index >= 0 ? index + 1 : 0;
}

/**
 * Helper to get number of days in a month
 */
function getDaysInMonth(month, year) {
  return new Date(year, month, 0).getDate();
}



