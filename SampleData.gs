/**
 * SAMPLE DATA INITIALIZATION SCRIPT
 * Run this once to populate the system with sample employee data
 */

function initializeSampleData() {
  // First, initialize the sheets
  initializeSheets();
  
  // Sample employees data
  const sampleEmployees = [
    {
      email: 'john.doe@company.com',
      name: 'John Doe',
      department: 'Engineering',
      managerEmail: 'sarah.manager@company.com'
    },
    {
      email: 'jane.smith@company.com',
      name: 'Jane Smith',
      department: 'Engineering',
      managerEmail: 'sarah.manager@company.com'
    },
    {
      email: 'mike.johnson@company.com',
      name: 'Mike Johnson',
      department: 'Marketing',
      managerEmail: 'tom.manager@company.com'
    },
    {
      email: 'emily.davis@company.com',
      name: 'Emily Davis',
      department: 'Marketing',
      managerEmail: 'tom.manager@company.com'
    },
    {
      email: 'sarah.manager@company.com',
      name: 'Sarah Manager',
      department: 'Engineering',
      managerEmail: 'ceo@company.com'
    },
    {
      email: 'tom.manager@company.com',
      name: 'Tom Manager',
      department: 'Marketing',
      managerEmail: 'ceo@company.com'
    }
  ];
  
  // Initialize each employee
  sampleEmployees.forEach(emp => {
    initializeEmployee(emp.email, emp.name, emp.department, emp.managerEmail);
  });
  
  Logger.log('Sample data initialized successfully!');
  Logger.log(`Added ${sampleEmployees.length} employees to the system.`);
  
  // Create a few sample leave requests
  createSampleLeaveRequests();
  
  // Add sample company holidays
  createSampleHolidays();
}

/**
 * Create sample leave requests for testing
 */
function createSampleLeaveRequests() {
  const sheet = getOrCreateSheet(CONFIG.SHEETS.LEAVE_REQUESTS, [
    'Request ID', 'Employee Email', 'Employee Name', 'Leave Type', 
    'Start Date', 'End Date', 'Days', 'Reason', 'Status', 
    'Applied Date', 'Manager Email', 'Action Date', 'Manager Comments'
  ]);
  
  const today = new Date();
  const tomorrow = new Date(today);
  tomorrow.setDate(tomorrow.getDate() + 1);
  
  const nextWeek = new Date(today);
  nextWeek.setDate(nextWeek.getDate() + 7);
  
  const sampleRequests = [
    {
      requestId: 'LR' + Date.now() + '001',
      employeeEmail: 'john.doe@company.com',
      employeeName: 'John Doe',
      leaveType: 'Casual Leave',
      startDate: formatDate(tomorrow),
      endDate: formatDate(tomorrow),
      days: 1,
      reason: 'Personal work',
      status: 'Pending',
      appliedDate: formatDate(today),
      managerEmail: 'sarah.manager@company.com',
      actionDate: '',
      managerComments: ''
    },
    {
      requestId: 'LR' + Date.now() + '002',
      employeeEmail: 'jane.smith@company.com',
      employeeName: 'Jane Smith',
      leaveType: 'Sick Leave',
      startDate: formatDate(nextWeek),
      endDate: formatDate(new Date(nextWeek.getTime() + 2 * 24 * 60 * 60 * 1000)),
      days: 3,
      reason: 'Medical appointment',
      status: 'Pending',
      appliedDate: formatDate(today),
      managerEmail: 'sarah.manager@company.com',
      actionDate: '',
      managerComments: ''
    }
  ];
  
  sampleRequests.forEach(req => {
    sheet.appendRow([
      req.requestId,
      req.employeeEmail,
      req.employeeName,
      req.leaveType,
      req.startDate,
      req.endDate,
      req.days,
      req.reason,
      req.status,
      req.appliedDate,
      req.managerEmail,
      req.actionDate,
      req.managerComments
    ]);
  });
  
  Logger.log('Sample leave requests created successfully!');
}

/**
 * Create sample company holidays for 2025
 */
function createSampleHolidays() {
  const sheet = getOrCreateSheet(CONFIG.SHEETS.COMPANY_HOLIDAYS, [
    'Holiday Date', 'Holiday Name', 'Holiday Type'
  ]);
  
  // Sample holidays for 2025 (Indian holidays)
  const sampleHolidays = [
    { date: '2025-01-26', name: 'Republic Day', type: 'National' },
    { date: '2025-03-14', name: 'Holi', type: 'Festival' },
    { date: '2025-03-31', name: 'Eid ul-Fitr', type: 'Festival' },
    { date: '2025-04-10', name: 'Mahavir Jayanti', type: 'Festival' },
    { date: '2025-04-14', name: 'Ambedkar Jayanti', type: 'National' },
    { date: '2025-04-18', name: 'Good Friday', type: 'Festival' },
    { date: '2025-05-01', name: 'May Day', type: 'National' },
    { date: '2025-06-07', name: 'Eid ul-Adha', type: 'Festival' },
    { date: '2025-08-15', name: 'Independence Day', type: 'National' },
    { date: '2025-08-27', name: 'Janmashtami', type: 'Festival' },
    { date: '2025-10-02', name: 'Gandhi Jayanti', type: 'National' },
    { date: '2025-10-22', name: 'Dussehra', type: 'Festival' },
    { date: '2025-11-01', name: 'Diwali', type: 'Festival' },
    { date: '2025-11-05', name: 'Guru Nanak Jayanti', type: 'Festival' },
    { date: '2025-12-25', name: 'Christmas', type: 'Festival' }
  ];
  
  sampleHolidays.forEach(holiday => {
    sheet.appendRow([
      new Date(holiday.date),
      holiday.name,
      holiday.type
    ]);
  });
  
  Logger.log(`Sample holidays created successfully! Added ${sampleHolidays.length} holidays.`);
}

/**
 * Reset all data (use with caution!)
 */
function resetAllData() {
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  
  // Delete all sheets
  const sheets = ss.getSheets();
  sheets.forEach(sheet => {
    if (sheet.getName() !== 'Sheet1') {
      ss.deleteSheet(sheet);
    }
  });
  
  Logger.log('All data has been reset. Run initializeSampleData() to start fresh.');
}

/**
 * View current system statistics
 */
function viewSystemStats() {
  const employeeSheet = SpreadsheetApp.openById(getSpreadsheetId()).getSheetByName(CONFIG.SHEETS.EMPLOYEES);
  const requestSheet = SpreadsheetApp.openById(getSpreadsheetId()).getSheetByName(CONFIG.SHEETS.LEAVE_REQUESTS);
  const historySheet = SpreadsheetApp.openById(getSpreadsheetId()).getSheetByName(CONFIG.SHEETS.LEAVE_HISTORY);
  
  const stats = {
    totalEmployees: employeeSheet ? employeeSheet.getLastRow() - 1 : 0,
    totalRequests: requestSheet ? requestSheet.getLastRow() - 1 : 0,
    totalTransactions: historySheet ? historySheet.getLastRow() - 1 : 0
  };
  
  Logger.log('=== SYSTEM STATISTICS ===');
  Logger.log('Total Employees: ' + stats.totalEmployees);
  Logger.log('Total Leave Requests: ' + stats.totalRequests);
  Logger.log('Total Transactions: ' + stats.totalTransactions);
  
  if (requestSheet && requestSheet.getLastRow() > 1) {
    const requests = requestSheet.getDataRange().getValues();
    let pending = 0, approved = 0, rejected = 0;
    
    for (let i = 1; i < requests.length; i++) {
      const status = requests[i][8];
      if (status === 'Pending') pending++;
      else if (status === 'Approved') approved++;
      else if (status === 'Rejected') rejected++;
    }
    
    Logger.log('Pending Requests: ' + pending);
    Logger.log('Approved Requests: ' + approved);
    Logger.log('Rejected Requests: ' + rejected);
  }
  
  return stats;
}
