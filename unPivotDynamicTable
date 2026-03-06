const fileIDW44 = '1ykSSyY7NH3QtqCXEDdMYhsfSdTwoZ1XK';
const fileIDW45 = '1r_uPMhQODnhWOe_fgNP6XGwOuu2ZlVOU';
const fileIDW46 = '1jBfMssgil1Dp0sHoiAV0-NTsNF6y7sRy';
//const ss = SpreadsheetApp.openById(fileIDW44);


function unpivotDynamicTable() {
  // Opens the spreadsheet file by its ID. If you created your script from a
// Google Sheets file, use SpreadsheetApp.getActiveSpreadsheet().
// TODO(developer): Replace the ID with your own.
const ss = SpreadsheetApp.openById(fileIDW44);

// Gets Sheet1 by its name.
const sheet = ss.getSheetByName('Sheet1');
||

  // 1. Find the starting row of the lowest header level
  const baseHeaderRowIndex = findHeaderRowIndex(data); 
  const headerRowsCount = 3; // Still assume 3 complex header rows (Year, Quarter, Metric)
  const idColumnsCount = 2;  // Still assume 2 ID columns (Product, Region)

  // Extract relevant sections of data
  const complexHeaders = data.slice(baseHeaderRowIndex, baseHeaderRowIndex + headerRowsCount);
  const dataRows = data.slice(baseHeaderRowIndex + headerRowsCount);

  const output = [];
  
  // 2. Map the column indices to the standardized header names
  // This map will look like: { 0: 'Product_Key', 1: 'Region', 2: 'Year_L1', ... }
  const columnHeaderMap = {};
  
  // A. Map the ID Columns (using the lowest header row for the name)
  for (let j = 0; j < idColumnsCount; j++) {
    columnHeaderMap[j] = normalizeHeader(complexHeaders[headerRowsCount - 1][j]);
  }
  
  // B. Map the Complex Headers
  // The standardized column name for the unpivoted output will be the combination
  const newHeaderNames = [];
  for (let j = 0; j < idColumnsCount; j++) {
      newHeaderNames.push(columnHeaderMap[j]);
  }
  newHeaderNames.push("Header_Level_1", "Header_Level_2", "Header_Level_3", "Value");
  output.push(newHeaderNames);


  // 3. Core Unpivot Logic with dynamic column index handling
  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i];
    
    // Extract ID values dynamically
    const idValues = row.slice(0, idColumnsCount);

    // Iterate through the columns containing the measure/value
    for (let j = idColumnsCount; j < complexHeaders[0].length; j++) {
      
      const headerL1 = normalizeHeader(complexHeaders[0][j]); 
      const headerL2 = normalizeHeader(complexHeaders[1][j]);
      const headerL3 = normalizeHeader(complexHeaders[2][j]);
      
      const value = row[j];

      if (value === "" || value === null) {
        continue; 
      }
      
      // Construct the new unpivoted row
      const newRow = [
        ...idValues, 
        headerL1, 
        headerL2, 
        headerL3, 
        value
      ];
      
      output.push(newRow);
    }
  }

  // 4. Write Output (same as before)
  const outputSheet = ss.getSheetByName('Unpivoted Output') || ss.insertSheet('Unpivoted Output');
  outputSheet.clearContents();
  outputSheet.getRange(1, 1, output.length, output[0].length).setValues(output);
}
