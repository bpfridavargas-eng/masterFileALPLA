// =================================================================================
// 1. CONSTANTES DE CONFIGURACIÓN
// DEBES MODIFICAR ESTOS VALORES
// =================================================================================

// El ID de la carpeta de Google Drive donde deseas guardar los archivos.
// Si lo dejas vacío (""), el archivo se guardará en la carpeta raíz (Mi Unidad).
// Ejemplo: "1A2b3C4d5E6f7G8h9I0j"
//ALPLA_attachments_xlsx = 1p41PUas0sKDBkQLKfdW9V3KYtAZeI1vV
const XLSXS_FOLDER_ID = "1p41PUas0sKDBkQLKfdW9V3KYtAZeI1vV"; 

// Criterio de búsqueda en Gmail. Busca en los últimos 7 días correos con adjuntos 
// que contengan el asunto "Repro Alpla WK". Ajusta esto si el asunto cambia.
// Nota: Añadí "is:unread" para evitar procesar el mismo correo dos veces.
const GMAIL_SEARCH_QUERY = 'subject:"ALPLA weekly file" has:attachment is:unread newer_than:7d';

// Nombre exacto de la etiqueta (label) que ya existe en Gmail para marcar estos correos.
const ALPLA_LABEL_NAME = "ALPLA weekly file label";

// =================================================================================
// 2. FUNCIÓN PRINCIPAL
// =================================================================================

/**
 * Busca el correo más reciente con el archivo y lo guarda en la carpeta de Drive.
 */
function guardarAdjuntoEnDrive() {
  Logger.log("Iniciando la búsqueda de archivos adjuntos...");
  
  try {
    // 1. Buscar el correo más reciente
    // Limita la búsqueda a 1 hilo (el más reciente)
    const threads = GmailApp.search(GMAIL_SEARCH_QUERY, 0, 1);

    if (threads.length === 0) {
      Logger.log("No se encontró ningún correo reciente con el archivo.");
      return;
    }

    // Obtenemos el hilo y el mensaje (necesitamos el hilo para la etiqueta)
    const thread = threads[0];
    const message = threads[0].getMessages().pop(); 
    const attachments = message.getAttachments();
    
    // 2. Identificar el archivo adjunto correcto
    // Filtra los adjuntos por su extensión (xlsx o csv)
    const targetAttachment = attachments.find(att => 
      att.getName().toLowerCase().includes('.xlsx') || att.getName().toLowerCase().includes('.csv')
    );

    if (!targetAttachment) {
      Logger.log("Se encontró el correo, pero no el adjunto .xlsx o .csv esperado.");
      return;
    }
    
    // 3. Determinar la carpeta de destino
    let targetFolder;
    if (XLSXS_FOLDER_ID) {
      targetFolder = DriveApp.getFolderById(XLSXS_FOLDER_ID);
    } else {
      // Si no se especifica ID, se guarda en la carpeta raíz (Mi Unidad)
      targetFolder = DriveApp.getRootFolder();
    }
    
    // 4. Guardar el archivo en Drive
    const newFile = targetFolder.createFile(targetAttachment);
    
    Logger.log(`✅ Archivo guardado con éxito: ${newFile.getName()}`);
    Logger.log(`Ubicación: ${newFile.getUrl()}`);

    // 5. Marcar como leído
    message.markRead();
    Logger.log("Correo marcado como leído.");

    // 6. Aplicar la etiqueta
    const label = GmailApp.getUserLabelByName(ALPLA_LABEL_NAME);

    if (label) {
      thread.addLabel(label);
      Logger.log(`Etiqueta "${ALPLA_LABEL_NAME}" aplicada con éxito.`);
    } else {
      Logger.log(`⚠️ Advertencia: No se encontró la etiqueta "${ALPLA_LABEL_NAME}". Verifique que el nombre sea exacto.`);
    }

    //7. Convertir de xlsx a GSheets
    let newSheetFileID = convertXLSXtoGoogleSheets(newFile, targetFolder);

    //8. ETL
    // Nota: Asegúrate de tener estas funciones definidas si las vas a usar
    // let weekNum = detectHeaders(newSheetFileID);
    // convertWeektoDate(weekNum);

    //9. Add new information to the master file
    addNewData(newSheetFileID);
    
  } catch (e) {
    Logger.log("❌ Ocurrió un error: " + e.toString());
  }
}

// =================================================================================
// 3. FUNCIONES DE APOYO (CONEXIÓN Y CONVERSIÓN)
// =================================================================================

/**
 * Convierte el archivo XLSX guardado a un formato Google Sheets para poder leerlo.
 * Requiere el servicio "Drive API" activado.
 */
function convertXLSXtoGoogleSheets(file, folder) {
  const blob = file.getBlob();
  const resource = {
    title: file.getName().replace(/\.xlsx$/i, ""),
    mimeType: MimeType.GOOGLE_SHEETS,
    parents: [{id: folder.getId()}]
  };
  
  // Realiza la conversión usando la API de Drive
  const spreadsheet = Drive.Files.insert(resource, blob);
  return spreadsheet.id;
}
/**
 * Abre el archivo recién convertido y transforma los datos para el Master.
 * Inserta los datos específicamente en la hoja "Sheet4".
 */
function addNewData(newSheetFileID) {
  // 1. Abre el archivo convertido (origen)
  const sourceSs = SpreadsheetApp.openById(newSheetFileID);
  const sourceSheet = sourceSs.getSheets()[0];
  const sourceData = sourceSheet.getDataRange().getValues();

  if (sourceData.length < 4) {
    Logger.log("⚠️ El archivo adjunto no tiene suficientes filas (mínimo 4 para detectar MPS).");
    return;
  }

  // 2. Identificar hasta qué columna debemos copiar
  // Buscamos en la fila 4 (índice 3) la palabra "MPS"
  const row4 = sourceData[3]; 
  let lastColIndex = -1;

  // Recorremos la fila de derecha a izquierda para encontrar el ÚLTIMO "MPS"
  // O de izquierda a derecha si prefieres el PRIMERO. Aquí busco el último:
  for (let i = row4.length - 1; i >= 0; i--) {
    if (String(row4[i]).toUpperCase().trim() === "MPS") {
      lastColIndex = i;
      break; 
    }
  }

  if (lastColIndex === -1) {
    Logger.log("❌ No se encontró el texto 'MPS' en la fila 4. Se copiarán todas las columnas por defecto.");
    lastColIndex = sourceData[0].length - 1;
  }

  // 3. Recortar los datos (solo hasta la columna donde está MPS)
  // El número de columnas a copiar es lastColIndex + 1
  const numColsToCopy = lastColIndex + 1;
  const filteredData = sourceData.map(row => row.slice(0, numColsToCopy));

  // 4. Obtiene la hoja de destino "Sheet4"
  const targetSs = SpreadsheetApp.getActiveSpreadsheet();
  let masterSheet = targetSs.getSheetByName("Recibe de correo");

  if (!masterSheet) {
    Logger.log("❌ Error: No se encontró la hoja 'Sheet4'.");
    return;
  }

  // 5. Limpiar y Pegar
  masterSheet.clear();
  masterSheet.getRange(1, 1, filteredData.length, filteredData[0].length).setValues(filteredData);

  Logger.log(`✅ Datos integrados en Sheet4. Se recortó hasta la columna ${numColsToCopy} (MPS).`);
}
