/**
 * Basic Certificate Generator
 * Generates a single type of certificate based on a Google Docs template.
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🎓 Certificates')
      .addItem('Generate New Certificates', 'generateCertificates')
      .addToUi();
}

function generateCertificates() {
  // --- CONFIGURATION ---
  // Replace these with your actual IDs
  const templateId = 'YOUR_TEMPLATE_DOC_ID_HERE';
  const baseFolderId = 'YOUR_BASE_DRIVE_FOLDER_ID_HERE';

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getDisplayValues();

  const template = DriveApp.getFileById(templateId);
  const baseFolder = DriveApp.getFolderById(baseFolderId);

  const todayDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy");
  
  // Create or get today's folder
  let destinationFolder;
  const folders = baseFolder.getFoldersByName(todayDate);
  if (folders.hasNext()) {
    destinationFolder = folders.next();
  } else {
    destinationFolder = baseFolder.createFolder(todayDate); 
  }

  let processed = 0;

  for (let i = 2; i < data.length; i++) {
    const row = data[i];

    let originalName = row[0]; 
    const name = originalName ? originalName.toString().replace(/\n/g, ' ').toUpperCase() : ''; 
    
    const idNumber = row[1];
    const majorOriginal = row[2];
    const majorName = majorOriginal ? majorOriginal.toString().toUpperCase() : ''; 
    const gender = row[3] ? row[3].toString().toUpperCase().trim() : ''; 
    
    const term = row[5];
    const currentYear = row[7];
    
    const startDate = row[9];
    const endDate = row[10];
    const dateInWords = row[11];
    const certCode = row[12];
    const status = row[13]; // Status column

    // Process if name exists and status is not 'Generated'
    if (name && name.trim() !== "" && status !== 'Generated') {

      // Language specific logic (e.g., gendered text in Spanish)
      let enrolledText = "está inscrito"; 
      if (gender === 'F' || gender === 'FEMENINO' || gender === 'FEMALE') {
        enrolledText = "está inscrita";
      }

      const fileName = certCode + " - " + name;
      const copy = template.makeCopy(fileName, destinationFolder);
      const doc = DocumentApp.openById(copy.getId());
      
      const body = doc.getBody();

      // Replace placeholders in the document body
      body.replaceText('<<CERT_CODE>>', certCode || '');
      body.replaceText('<<STUDENT_NAME>>', name);
      body.replaceText('<<ID_NUMBER>>', idNumber || '');
      body.replaceText('<<MAJOR_NAME>>', majorName);
      body.replaceText('<<TERM>>', term || '');
      body.replaceText('<<ENROLLED_STATUS>>', enrolledText); 
      body.replaceText('<<CURRENT_YEAR>>', currentYear || ''); 
      body.replaceText('<<YEAR>>', currentYear || ''); 
      body.replaceText('<<START_DATE>>', startDate || '');
      body.replaceText('<<END_DATE>>', endDate || '');
      body.replaceText('<<DATE_IN_WORDS>>', dateInWords || '');

      // Replace placeholders in Footer
      const footer = doc.getFooter();
      if (footer) {
        footer.replaceText('<<CERT_CODE>>', certCode || '');
      }
      
      // Replace placeholders in Header
      const header = doc.getHeader();
      if (header) {
        header.replaceText('<<CERT_CODE>>', certCode || '');
      }

      doc.saveAndClose();

      // Mark as 'Generated' in column N (Index 13 in arrays is 14 in the sheet)
      sheet.getRange(i + 1, 14).setValue('Generated');
      processed++;
    }
  }

  // Final alerts
  if (processed > 0) {
    SpreadsheetApp.getUi().alert('Success! ' + processed + ' new certificates were generated and saved in the folder: ' + todayDate);
  } else {
    SpreadsheetApp.getUi().alert('No new students found. Make sure to delete the word "Generated" in the status column to regenerate.');
  }
}