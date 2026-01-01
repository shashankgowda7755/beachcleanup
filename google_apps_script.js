// -------------------------------------------------------------------------------------------------
// GOOGLE APPS SCRIPT CODE
// 1. Go to https://script.google.com/home
// 2. Create a 'New Project'
// 3. Delete any code in 'Code.gs' and paste this entire script.
// 4. (Optional) Create a Folder in Google Drive to store images, and copy its ID into the 'DRIVE_FOLDER_ID' variable below.
// 5. Click 'Deploy' -> 'New Deployment' -> Select type 'Web app'.
// 6. Set 'Execute as': 'Me'
// 7. Set 'Who has access': 'Anyone' (IMPORTANT)
// 8. Click 'Deploy' and copy the 'Web App URL'.
// 9. Paste that URL into your HTML file where it says 'GOOGLE_SCRIPT_URL'.
// -------------------------------------------------------------------------------------------------

const DRIVE_FOLDER_ID = ""; // Paste your Drive Folder ID here if you want to save photos (e.g., "1a2b3c..."). Leave empty to skip photo saving.
const SHEET_NAME = "Sheet1"; // Name of the tab to store data

function doPost(e) {
    try {
        const data = JSON.parse(e.postData.contents);
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        let sheet = ss.getSheetByName(SHEET_NAME);

        // Create sheet if not exists
        if (!sheet) {
            sheet = ss.insertSheet(SHEET_NAME);
            // Add Headers
            sheet.appendRow(["Date", "Full Name", "Phone", "Email", "Photo URL/Base64"]);
        }

        let photoValue = "No Photo";

        // Handle Photo (Upload to Drive or Store Raw)
        if (data.photo) {
            if (DRIVE_FOLDER_ID && DRIVE_FOLDER_ID.length > 5) {
                try {
                    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
                    const type = data.photo.split(';')[0].split('/')[1]; // e.g., 'png'
                    const decoded = Utilities.base64Decode(data.photo.split(',')[1]);
                    const blob = Utilities.newBlob(decoded, 'image/' + type, data.fullName + "_" + Date.now() + "." + type);
                    const file = folder.createFile(blob);
                    photoValue = file.getUrl(); // Store the Drive Link
                } catch (err) {
                    photoValue = "Error saving to Drive: " + err.toString();
                }
            } else {
                // If no folder ID, we can't store huge base64 strings in a cell reliably (50k limit). 
                // We'll store a truncated version or flag it.
                photoValue = "Base64 Image (Drive Folder ID not set)";
            }
        }

        // Append Data
        sheet.appendRow([
            new Date(),
            data.fullName,
            data.phone,
            data.email,
            photoValue
        ]);

        return ContentService.createTextOutput(JSON.stringify({ "status": "success", "message": "Data saved" }))
            .setMimeType(ContentService.MimeType.JSON);

    } catch (error) {
        return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": error.toString() }))
            .setMimeType(ContentService.MimeType.JSON);
    }
}
