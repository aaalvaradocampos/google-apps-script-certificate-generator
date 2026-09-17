/**
 * Dynamic Certificate Generator
 * Generates different types of certificates based on a specific column value in Google Sheets.
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🎓 Certificates')
      .addItem('Generate New Certificates', 'generateCertificates')
      .addToUi();
}

// Helper function to create or get today's folder
function getDestinationFolder(baseFolder, folderName) {
  const folders = baseFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return baseFolder.createFolder(folderName); 
  }
}

function generateCertificates() {
  // ---------------- CONFIGURATION: TYPE A (E.g., Active Student) ----------------
  const templateIdTypeA = 'YOUR_TYPE_A_TEMPLATE_ID_HERE';
  const baseFolderIdTypeA = 'YOUR_TYPE_A_FOLDER_ID_HERE';

  // ---------------- CONFIGURATION: TYPE B (E.g., Social Service) ----------------
  const templateIdTypeB = 'YOUR_TYPE_B_TEMPLATE_ID_HERE'; 
  const baseFolderIdTypeB = 'YOUR_TYPE_B_FOLDER_ID_HERE';

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getDisplayValues();

  const todayDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy");

  // Instantiate base folders
  const baseFolderTypeA = DriveApp.getFolderById(baseFolderIdTypeA);
  const baseFolderTypeB = DriveApp.getFolderById(baseFolderIdTypeB);

  // Get subfolders with today's date
  const destinationFolderTypeA = getDestinationFolder(baseFolderTypeA, todayDate);
  const destinationFolderTypeB = getDestinationFolder(baseFolderTypeB, todayDate);

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
    
    // Read column O (index 14) for the certificate type
    const certType = row[14] ? row[14].toString().toUpperCase().trim() : 'TYPE_A'; 

    if (name && name.trim() !== "" && status !== 'Generated') {

      // Language specific logic (e.g., gendered text)
      let enrolledText = "está inscrito"; 
      if (gender === 'F' || gender === 'FEMENINO' || gender === 'FEMALE') {
        enrolledText = "está inscrita";
      }

      // Define which template and folder to use based on the type
      let templateIdToUse;
      let destinationFolderToUse;

      if (certType === 'SOCIAL SERVICE' || certType === 'TYPE_B') {
        templateIdToUse = templateIdTypeB;
        destinationFolderToUse = destinationFolderTypeB;
      } else {
        templateIdToUse = templateIdTypeA;
        destinationFolderToUse = destinationFolderTypeA;
      }

      const template = DriveApp.getFileById(templateIdToUse);
      const fileName = certCode + " - " + name;
      const copy = template.makeCopy(fileName, destinationFolderToUse);
      const doc = DocumentApp.openById(copy.getId());
      
      const body = doc.getBody();

      // Replace placeholders
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
    SpreadsheetApp.getUi().alert('Success! ' + processed + ' new certificates were generated and saved in their respective folders under the date: ' + todayDate);
  } else {
    SpreadsheetApp.getUi().alert('No new students found. Make sure to delete the word "Generated" in the status column to regenerate.');
  }
}