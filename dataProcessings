function processPlanningFile() {
  // --- Configuración ---
  const ARCHIVO_SEMANAL_ID = '1J3CFMUSESLzfcc3uvtLllOXXD3m5QbgndLaJ7lyVDwk';
  const NOMBRE_HOJA_ORIGEN = 'Sheet1'; 

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const externalWorkbook = SpreadsheetApp.openById(ARCHIVO_SEMANAL_ID);
  const sourceSheet = externalWorkbook.getSheetByName(NOMBRE_HOJA_ORIGEN);
  const masterSheet = ss.getSheetByName("Consolidado");

  // 1. EXTRACT DATE FROM FILENAME (e.g., "29 Oct 2025")
  const fileName = externalWorkbook.getName();
  const dateMatch = fileName.match(/(\d{1,2}\s+[a-zA-Z]+\s+\d{4})/);
  const reportDate = dateMatch ? new Date(dateMatch[0]) : new Date();

  // 2. DISCOVER HEADERS IN SOURCE
  const lastColSource = sourceSheet.getLastColumn() || 1;
  const searchRange = sourceSheet.getRange(1, 1, 10, lastColSource).getValues();
  let headerRowIndex = -1;
  for (let i = 0; i < 5; i++) {
    if (searchRange[i].indexOf("Símbolo") > -1) { headerRowIndex = i; break; }
  }

  const headers = searchRange[headerRowIndex];
  const sourceData = sourceSheet.getDataRange().getValues().slice(headerRowIndex + 1);

  // 3. IDENTIFY LATEST WEEK
  let targetWeekNum = -1;
  let targetWeekIdx = -1;
  const weekRegex = /(?:SEM|Sem|Week)\s*(\d+)|^(\d+)$/i;

  headers.forEach((header, idx) => {
    const match = String(header).trim().match(weekRegex);
    if (match && !/^w\d+$/i.test(header)) {
      const weekVal = parseInt(match[1] || match[2]);
      if (weekVal > targetWeekNum) { targetWeekNum = weekVal; targetWeekIdx = idx; }
    }
  });

  // 4. MASTER FILE MAPPING (Assuming Column A is Símbolo)
  const masterData = masterSheet.getDataRange().getValues();
  const masterSimbolos = masterData.map(r => r[0].toString());
  
  // Create Composite Header: Row 1 = Week, Row 2 = Date
  let targetMasterCol = masterSheet.getLastColumn() + 1;
  masterSheet.getRange(1, targetMasterCol).setValue(`SEM ${targetWeekNum}`);
  masterSheet.getRange(2, targetMasterCol).setValue(reportDate).setNumberFormat('dd/mm/yyyy');

  // 5. UPDATE DATA
  const colMap = {
    simbolo: headers.indexOf("Símbolo"),
    desc: headers.indexOf("Descripción"),
    cat: headers.indexOf("Categoría"),
    size: headers.indexOf("Size"),
    prov: headers.indexOf("Proveedor")
  };

  sourceData.forEach(row => {
    const simbolo = row[colMap.simbolo];
    if (!simbolo) return;

    const rowIndex = masterSimbolos.indexOf(simbolo.toString());
    const weekValue = row[targetWeekIdx];

    if (rowIndex > -1) {
      // RowIndex + 1 because array is 0-indexed
      masterSheet.getRange(rowIndex + 1, targetMasterCol).setValue(weekValue);
    } else {
      const lastRow = masterSheet.getLastRow() + 1;
      const staticInfo = [row[colMap.simbolo], row[colMap.desc], row[colMap.cat], row[colMap.size], row[colMap.prov]];
      masterSheet.getRange(lastRow, 1, 1, 5).setValues([staticInfo]);
      masterSheet.getRange(lastRow, targetMasterCol).setValue(weekValue);
      masterSimbolos.push(simbolo.toString());
    }
  });

  SpreadsheetApp.getUi().alert(`Imported Week ${targetWeekNum} with date ${reportDate.toLocaleDateString()}`);
}
