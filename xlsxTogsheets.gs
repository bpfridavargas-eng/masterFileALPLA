//ALPLA ARCHIVE FOLDER ID
ALPLA_FOLDER_ID = "1SAWlwCUBOVdU5lm_FAjBKTscpQOFt9iJ";

function convertXLSXtoGoogleSheets(xlsx_file) {
  // --- PREREQUISITE: Ensure the 'Drive' Advanced Service is enabled (via + Service in the editor) ---

    const excelBlob = xlsx_file.getBlob();

    // Configuration for the new Google Sheets file
    Logger.log(xlsx_file.getName().split(".")[0]);
    const resource = {
      'name': xlsx_file.getName().split(".")[0]+' CONVERTED',
      'mimeType': MimeType.GOOGLE_SHEETS,
      'parents': [ALPLA_FOLDER_ID] 
      //parents: [{ id: excelFile.getParents().next().getId() }]
    };

    // Use the *ADVANCED SERVICE* 'Drive.Files.create()' method
    // This is where the conversion magic happens.
    const newSheetFile = Drive.Files.create(resource, excelBlob);
    Logger.log("Converted file ID: " + newSheetFile.id);

    
    return newSheetFile.id;

  //} else {
  //  Logger.log("Error: Excel file not found.");
 // }
}
