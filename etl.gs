const COLUMNAS_FIJAS = 5;

function detectHeaders() {
  let ARCHIVO_SEMANAL_ID = '1qecj8wYfQqhLjfp0m0S154AQsFaVZmvpftTwCKdg1Vk';
  //let ARCHIVO_SEMANAL_ID = '1ynhRZvldJeVTLm633DS8w_bvMIozF0p7BRC6xyoh2mY';


  try {
    // 1. OBTENER DATOS     
    // Archivo Semanal (Abre el archivo por su ID)
    const ssSemanal = SpreadsheetApp.openById(ARCHIVO_SEMANAL_ID);
    // Asumimos que los datos están en la primera hoja del archivo semanal
    const hojaSemanal = ssSemanal.getSheets()[0];

    // Obtener todos los datos del archivo semanal (hasta el límite de datos)
    const rangoCompletoSemanal = hojaSemanal.getDataRange();
    const valoresSemanal = rangoCompletoSemanal.getValues();

    // 2. ENCONTRAR LA FILA DEL ENCABEZADO
    const { headerRowIndex, headerValues } = encontrarFilaEncabezado(valoresSemanal);

    // 3. ENCONTRAR COLUMNA
    Logger.log("header X" + headerRowIndex + ' Header Y:' + headerValues);

    const first1erIndex = headerValues.indexOf('1er');
    //Primer columna que encuentre '1er' - 2 se encuentra la información de la última semana
    const col_current_week = first1erIndex - 2;

    let header_text = valoresSemanal[col_current_week];

    // The 'w' format character returns the week number of the year (1-52 or 53)
    // based on the ISO 8601 standard (week 1 contains the first Thursday of the year).
    const weekNumberString = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'w');
    return parseInt(weekNumberString);

    Logger.log(valoresSemanal[headerRowIndex][col_current_week]);

  } catch (e) {
    Logger.log('❌ ERROR al encontrar indices de las columnas: ' + e.toString());
    SpreadsheetApp.getUi().alert(`❌ Error al consolidar los datos de las columnas:\n${e.message}`);
  }
}


function encontrarFilaEncabezado(valores) {
  // Buscamos solo en las primeras 6 filas (índice 0 a 5)
  const maxSearchRows = Math.min(valores.length, 6);

  for (let i = 0; i < maxSearchRows; i++) {
    const row = valores[i];
    // Usamos .some() para buscar si algún valor en la fila coincide con el patrón
    const found = row.some(cellValue =>
      cellValue == 'W4'
    );

    if (found) {
      // Retornamos el índice de la fila y los valores de esa fila
      return {
        headerRowIndex: i,
        headerValues: row
      };
    }
  }

  throw new Error("No se encontró una fila de encabezado con el patrón 'semX' en las primeras 5 filas del archivo semanal.");
}

//Obtener la fecha del primer día de la semana recibida
function convertWeektoDate(weekNum) {
  let year = new Date().getFullYear;
  Logger.log(year);
  // 1. Start with January 4th of the target year.
  // This date is always in Week 1 according to the ISO 8601 standard.
  const jan4 = new Date(year, 0, 4);

  // 2. Determine the day of the week for Jan 4th (0=Sun, 1=Mon, ..., 6=Sat).
  let dayOfWeek = jan4.getDay();

  // Convert to ISO 8601 day index: 1=Mon, 2=Tue, ..., 7=Sun.
  // If it's Sunday (0), set it to 7. Otherwise, use the existing value.
  if (dayOfWeek === 0) {
    dayOfWeek = 7;
  }

  // Define milliseconds in a day for calculations
  const oneDayInMs = 24 * 60 * 60 * 1000;

  // 3. Calculate the date of the Monday of Week 1.
  // We subtract (dayOfWeek - 1) days from Jan 4th to land on the preceding Monday.
  const mondayOfWeek1_ms = jan4.getTime() - (dayOfWeek - 1) * oneDayInMs;

  // 4. Calculate the Monday of the target week.
  // We add (weekNum - 1) * 7 days (in milliseconds) to the Monday of Week 1.
  const targetMonday_ms = mondayOfWeek1_ms + (weekNum - 1) * 7 * oneDayInMs;

  // 5. Create the final Date object.
  const targetDate = new Date(targetMonday_ms);

  // To verify the result in the script's timezone (optional, but helpful):
  Logger.log(
    `The Monday of Week ${weekNum}, ${year} is: ${Utilities.formatDate(
      targetDate,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    )}`
  );

  return targetDate;
}



function extractColumn(matrix, columnIndex) {
  // Verificación básica para evitar errores si la matriz está vacía.
  if (!matrix || matrix.length === 0) {
    return [];
  }
  
  // Usamos .map() para iterar sobre cada 'fila' (row) de la matriz.
  return matrix.map(function(row) {
    // Para cada fila, devolvemos el valor que se encuentra en el columnIndex especificado.
    return row[columnIndex];
  });
}



