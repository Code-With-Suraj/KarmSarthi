// ============================================================================
// DOCUMENT MANAGEMENT FUNCTIONS
// ============================================================================

/**
 * Get or create the Google Drive folder for employee documents
 * @returns {Folder} Google Drive folder object
 */
function getOrCreateDocumentFolder() {
  const folderName = CONFIG.DOCUMENTS.FOLDER_NAME;
  const folders = DriveApp.getFoldersByName(folderName);
  
  if (folders.hasNext()) {
    return folders.next();
  } else {
    // Create new folder
    const folder = DriveApp.createFolder(folderName);
    Logger.log(`Created new folder: ${folderName}`);
    return folder;
  }
}

/**
 * Generate unique document ID
 * @returns {string} Unique document ID
 */
function generateDocumentId() {
  return 'DOC-' + new Date().getTime() + '-' + Math.random().toString(36).substr(2, 9).toUpperCase();
}

/**
 * Upload employee document to Google Drive
 * @param {string} documentType - Type of document
 * @param {string} fileName - Original file name
 * @param {string} fileContent - Base64 encoded file content
 * @param {string} mimeType - MIME type of the file
 * @returns {Object} Result object with success status and message
 */
function uploadEmployeeDocument(documentType, fileName, fileContent, mimeType) {
  try {
    const email = getCurrentUserEmail();
    const employeeData = getEmployeeData();
    
    if (!employeeData) {
      return { success: false, message: 'Employee not found. Please contact HR.' };
    }
    
    // Validate document type
    if (!CONFIG.DOCUMENTS.DOCUMENT_TYPES.includes(documentType)) {
      return { success: false, message: 'Invalid document type.' };
    }
    
    // Decode base64 content
    const blob = Utilities.newBlob(
      Utilities.base64Decode(fileContent),
      mimeType,
      fileName
    );
    
    // Check file size
    const fileSize = blob.getBytes().length;
    if (fileSize > CONFIG.DOCUMENTS.MAX_FILE_SIZE) {
      return { 
        success: false, 
        message: `File size exceeds maximum limit of ${CONFIG.DOCUMENTS.MAX_FILE_SIZE / (1024 * 1024)}MB.` 
      };
    }
    
    // Get or create document folder
    const folder = getOrCreateDocumentFolder();
    
    // Create employee-specific subfolder
    const employeeFolderName = `${employeeData.name} (${email})`;
    let employeeFolder;
    const employeeFolders = folder.getFoldersByName(employeeFolderName);
    
    if (employeeFolders.hasNext()) {
      employeeFolder = employeeFolders.next();
    } else {
      employeeFolder = folder.createFolder(employeeFolderName);
    }
    
    // Upload file to Drive
    const file = employeeFolder.createFile(blob);
    const fileId = file.getId();
    
    // Save metadata to sheet
    const sheet = getOrCreateSheet(CONFIG.SHEETS.EMPLOYEE_DOCUMENTS, [
      'Document ID', 'Employee Email', 'Employee Name', 'Document Type',
      'File Name', 'Drive File ID', 'Upload Date', 'File Size (bytes)',
      'Verification Status', 'Verified By', 'Verification Date', 'HR Comments'
    ]);
    
    const documentId = generateDocumentId();
    const uploadDate = new Date();
    
    sheet.appendRow([
      documentId,
      email,
      employeeData.name,
      documentType,
      fileName,
      fileId,
      uploadDate,
      fileSize,
      CONFIG.DOCUMENTS.STATUS.PENDING,
      '',
      '',
      ''
    ]);
    
    Logger.log(`Document uploaded successfully: ${documentId}`);
    
    return { 
      success: true, 
      message: 'Document uploaded successfully and pending verification.',
      documentId: documentId
    };
    
  } catch (e) {
    Logger.log(`Error uploading document: ${e.toString()}`);
    return { 
      success: false, 
      message: 'Failed to upload document. Please try again.' 
    };
  }
}

/**
 * Get all documents for current employee
 * @returns {Array} Array of document objects
 */
function getEmployeeDocuments() {
  try {
    const email = getCurrentUserEmail();
    const sheet = getOrCreateSheet(CONFIG.SHEETS.EMPLOYEE_DOCUMENTS, [
      'Document ID', 'Employee Email', 'Employee Name', 'Document Type',
      'File Name', 'Drive File ID', 'Upload Date', 'File Size (bytes)',
      'Verification Status', 'Verified By', 'Verification Date', 'HR Comments'
    ]);
    
    const data = sheet.getDataRange().getValues();
    const documents = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === email) {
        documents.push({
          documentId: data[i][0],
          employeeEmail: data[i][1],
          employeeName: data[i][2],
          documentType: data[i][3],
          fileName: data[i][4],
          driveFileId: data[i][5],
          uploadDate: formatDate(data[i][6]),
          fileSize: data[i][7],
          verificationStatus: data[i][8],
          verifiedBy: data[i][9] || '',
          verificationDate: data[i][10] ? formatDate(data[i][10]) : '',
          hrComments: data[i][11] || ''
        });
      }
    }
    
    return documents.reverse(); // Most recent first
    
  } catch (e) {
    Logger.log(`Error getting employee documents: ${e.toString()}`);
    return [];
  }
}

/**
 * Get all employee documents (HR function)
 * @returns {Array} Array of all document objects
 */
function getAllEmployeeDocuments() {
  try {
    const sheet = getOrCreateSheet(CONFIG.SHEETS.EMPLOYEE_DOCUMENTS, [
      'Document ID', 'Employee Email', 'Employee Name', 'Document Type',
      'File Name', 'Drive File ID', 'Upload Date', 'File Size (bytes)',
      'Verification Status', 'Verified By', 'Verification Date', 'HR Comments'
    ]);
    
    const data = sheet.getDataRange().getValues();
    const documents = [];
    
    for (let i = 1; i < data.length; i++) {
      documents.push({
        documentId: data[i][0],
        employeeEmail: data[i][1],
        employeeName: data[i][2],
        documentType: data[i][3],
        fileName: data[i][4],
        driveFileId: data[i][5],
        uploadDate: formatDate(data[i][6]),
        fileSize: data[i][7],
        verificationStatus: data[i][8],
        verifiedBy: data[i][9] || '',
        verificationDate: data[i][10] ? formatDate(data[i][10]) : '',
        hrComments: data[i][11] || ''
      });
    }
    
    return documents.reverse(); // Most recent first
    
  } catch (e) {
    Logger.log(`Error getting all employee documents: ${e.toString()}`);
    return [];
  }
}

/**
 * Verify or reject a document (HR function)
 * @param {string} documentId - Document ID to verify
 * @param {string} status - Verification status (Verified/Rejected)
 * @param {string} comments - HR comments
 * @returns {Object} Result object with success status and message
 */
function verifyDocument(documentId, status, comments) {
  try {
    const hrEmail = getCurrentUserEmail();
    const sheet = getOrCreateSheet(CONFIG.SHEETS.EMPLOYEE_DOCUMENTS, [
      'Document ID', 'Employee Email', 'Employee Name', 'Document Type',
      'File Name', 'Drive File ID', 'Upload Date', 'File Size (bytes)',
      'Verification Status', 'Verified By', 'Verification Date', 'HR Comments'
    ]);
    
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === documentId) {
        const verificationDate = new Date();
        
        // Update verification status
        sheet.getRange(i + 1, 9).setValue(status); // Verification Status
        sheet.getRange(i + 1, 10).setValue(hrEmail); // Verified By
        sheet.getRange(i + 1, 11).setValue(verificationDate); // Verification Date
        sheet.getRange(i + 1, 12).setValue(comments || ''); // HR Comments
        
        // Create notification for employee
        const employeeEmail = data[i][1];
        const employeeName = data[i][2];
        const documentType = data[i][3];
        
        const notificationMessage = status === CONFIG.DOCUMENTS.STATUS.VERIFIED
          ? `Your ${documentType} has been verified by HR.`
          : `Your ${documentType} has been rejected by HR. ${comments ? 'Reason: ' + comments : ''}`;
        
        createNotification(
          employeeEmail,
          hrEmail,
          'HR Team',
          'document_verification',
          documentId,
          notificationMessage
        );
        
        Logger.log(`Document ${documentId} ${status.toLowerCase()} by ${hrEmail}`);
        
        return { 
          success: true, 
          message: `Document ${status.toLowerCase()} successfully.` 
        };
      }
    }
    
    return { success: false, message: 'Document not found.' };
    
  } catch (e) {
    Logger.log(`Error verifying document: ${e.toString()}`);
    return { 
      success: false, 
      message: 'Failed to verify document. Please try again.' 
    };
  }
}

/**
 * Delete a document (only pending documents can be deleted by employee)
 * @param {string} documentId - Document ID to delete
 * @returns {Object} Result object with success status and message
 */
function deleteDocument(documentId) {
  try {
    const email = getCurrentUserEmail();
    const sheet = getOrCreateSheet(CONFIG.SHEETS.EMPLOYEE_DOCUMENTS, [
      'Document ID', 'Employee Email', 'Employee Name', 'Document Type',
      'File Name', 'Drive File ID', 'Upload Date', 'File Size (bytes)',
      'Verification Status', 'Verified By', 'Verification Date', 'HR Comments'
    ]);
    
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === documentId && data[i][1] === email) {
        const status = data[i][8];
        
        // Only allow deletion of pending documents
        if (status !== CONFIG.DOCUMENTS.STATUS.PENDING) {
          return { 
            success: false, 
            message: 'Only pending documents can be deleted.' 
          };
        }
        
        // Delete file from Drive
        try {
          const fileId = data[i][5];
          const file = DriveApp.getFileById(fileId);
          file.setTrashed(true);
        } catch (e) {
          Logger.log(`Error deleting file from Drive: ${e.toString()}`);
        }
        
        // Delete row from sheet
        sheet.deleteRow(i + 1);
        
        Logger.log(`Document ${documentId} deleted by ${email}`);
        
        return { 
          success: true, 
          message: 'Document deleted successfully.' 
        };
      }
    }
    
    return { success: false, message: 'Document not found or access denied.' };
    
  } catch (e) {
    Logger.log(`Error deleting document: ${e.toString()}`);
    return { 
      success: false, 
      message: 'Failed to delete document. Please try again.' 
    };
  }
}

/**
 * Get document download URL
 * @param {string} documentId - Document ID
 * @returns {Object} Result object with download URL or error
 */
function getDocumentDownloadUrl(documentId) {
  try {
    const email = getCurrentUserEmail();
    const sheet = getOrCreateSheet(CONFIG.SHEETS.EMPLOYEE_DOCUMENTS, [
      'Document ID', 'Employee Email', 'Employee Name', 'Document Type',
      'File Name', 'Drive File ID', 'Upload Date', 'File Size (bytes)',
      'Verification Status', 'Verified By', 'Verification Date', 'HR Comments'
    ]);
    
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === documentId && data[i][1] === email) {
        const fileId = data[i][5];
        const file = DriveApp.getFileById(fileId);
        
        // Make file accessible to anyone with the link (temporary)
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        
        const downloadUrl = file.getDownloadUrl();
        
        return { 
          success: true, 
          url: downloadUrl,
          fileName: data[i][4]
        };
      }
    }
    
    return { success: false, message: 'Document not found or access denied.' };
    
  } catch (e) {
    Logger.log(`Error getting download URL: ${e.toString()}`);
    return { 
      success: false, 
      message: 'Failed to get download URL. Please try again.' 
    };
  }
}

/**
 * Get document types configuration
 * @returns {Array} Array of document types
 */
function getDocumentTypes() {
  return CONFIG.DOCUMENTS.DOCUMENT_TYPES;
}
