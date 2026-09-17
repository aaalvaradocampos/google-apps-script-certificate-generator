# Google Apps Script Certificate Generator 🎓

This repository contains Google Apps Script codes designed to automate the generation of customized certificates (or official documents) in Google Docs using data from Google Sheets. 

It streamlines the process of mass-creating documents, replacing placeholder tags with student data, and organizing the generated files into specific Google Drive folders categorized by date.

## 📂 Included Files

1. **`basic_generator.js`**: A straightforward script that generates a single type of certificate. It reads a Google Sheet, duplicates a specified Google Doc template, replaces the placeholders, and saves it in a daily folder.
2. **`dynamic_generator.js`**: An advanced version of the script. It reads a specific column in the spreadsheet to determine the *type* of certificate needed (e.g., Active Student vs. Social Service). It dynamically routes the document generation to use different templates and different destination folders based on that status.

## ✨ Features

- **Custom Menu:** Adds a custom UI menu in Google Sheets (`🎓 Certificates`) to run the generation with a single click.
- **Dynamic Placeholders:** Replaces tags in the Google Doc body, header, and footer (e.g., `<<STUDENT_NAME>>`, `<<ID_NUMBER>>`).
- **Auto-Formatting:** Automatically converts names and majors to uppercase to maintain formal document standards.
- **Language Logic:** Includes conditional logic to adjust text based on gender (e.g., changing text to female/male grammatical rules in Spanish).
- **Smart Organization:** Automatically creates a new folder in Google Drive named with the current date (`dd-MM-yyyy`) to store the generated documents.
- **Status Tracking:** Marks processed rows as "Generated" in the spreadsheet to prevent duplicate document creation.

## 🚀 How to Use

1. Open your Google Sheet containing the student data.
2. Go to **Extensions > Apps Script**.
3. Clear any existing code and paste either `basic_generator.js` or `dynamic_generator.js`.
4. **Important Configuration:** Replace the placeholder IDs at the top of the script with your actual Google Drive IDs:
   - `YOUR_TEMPLATE_DOC_ID_HERE`: The ID of your Google Doc template.
   - `YOUR_BASE_DRIVE_FOLDER_ID_HERE`: The ID of the Drive folder where daily folders will be created.
5. Save the project and refresh your Google Sheet.
6. A new menu called **🎓 Certificates** will appear. Click it and select **Generate New Certificates**.
7. Grant the necessary permissions to Google Apps Script the first time you run it.

## 📝 Setting up the Template

In your Google Doc template, use the following format for placeholders. The script will automatically find and replace them:
- `<<CERT_CODE>>`
- `<<STUDENT_NAME>>`
- `<<ID_NUMBER>>`
- `<<MAJOR_NAME>>`
- `<<TERM>>`
- `<<ENROLLED_STATUS>>`

*(Make sure these tags match exactly with the ones defined in the `body.replaceText()` functions inside the script).*

## ⚠️ Privacy Note
These scripts are sanitized. Be sure **not** to commit your actual Google Drive IDs or personal student data to public repositories.
