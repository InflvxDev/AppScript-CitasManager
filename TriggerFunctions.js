function eliminarRegistrosResiduales(){

  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hojas = libro.getSheets();

  hojas.forEach(hoja => {
    const nombreHoja = hoja.getName();
    
    // Omitir las hojas "CONFIGURACION" y "GENERAL"
    if (nombreHoja === "CONFIGURACION" || nombreHoja === "GENERAL") {
      return;
    }

    const protecciones = hoja.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    const filasProtegidas = new Set();

    Logger.log(hoja.getName());
    
    // Recopilar las filas protegidas
    protecciones.forEach(function(proteccion) {
      if (proteccion.getDescription() === "Bloqueado automáticamente") {
        const rango = proteccion.getRange();
        const filaInicio = rango.getRow();
        const numFilas = rango.getNumRows();
        
        for (let i = 0; i < numFilas; i++) {
          filasProtegidas.add(filaInicio + i);
        }
      }
    });

    const ultimaFila = hoja.getLastRow();
    let filasMargen = 0;
    
    // Recorrer desde la segunda fila hacia abajo (ignorando encabezado)
    for (let fila = 2; fila <= ultimaFila; fila++) {
      const valoresFila = hoja.getRange(fila, 1, 1, hoja.getLastColumn()).getValues()[0];
      const tieneDatos = valoresFila.some((valor, index) => index !== 10 && valor !== "" && valor !== null);
      
      if (!tieneDatos) {
        filasMargen++;
        if(filasMargen > 10){
          break;
        }
        continue;
      } else{
        filasMargen = 0;
      }

      if (!filasProtegidas.has(fila)) {
        const rangoAntes = hoja.getRange(fila, 1, 1, 10); // Columnas A-J
        const rangoDespues = hoja.getRange(fila, 12, 1, hoja.getLastColumn() - 10); // Columnas L en adelante
        
        rangoAntes.clearContent(); // Borra contenido de A-I
        rangoDespues.clearContent(); // Borra contenido de K en adelante
      }
    }

  });
}

function removerEditores() {
  const hojaGeneral = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("GENERAL");
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if(!hoja){
      SpreadsheetApp.getUi().alert("No se encontro una hoja activa");
  }

  const archivo = SpreadsheetApp.getActiveSpreadsheet();
  const file = DriveApp.getFileById(archivo.getId());
  const propietario = file.getOwner().getEmail();
  const protecciones = hoja.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  const proteccionesGeneral = hojaGeneral.getProtections(SpreadsheetApp.ProtectionType.RANGE);

  removerEditoresNoPropietarios(protecciones, propietario);
  removerEditoresNoPropietarios(proteccionesGeneral, propietario);
  
}

function validarCamposUsuario(){
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if(!hoja){
    SpreadsheetApp.getUi().alert("No se encontro una hoja activa");
  }

  if(hoja.getName() === "GENERAL") return;

  const fila =  hoja.getActiveCell().getRow();
  const rangoValores = hoja.getRange(fila, 2, 1, 13).getValues()[0];

  if (fila === 1) {
    return true;
  }
  
  let errores = [];

  const longitudesDocumento = {
    "CC": [10, 8, 7, 6], "CE": [6], "CD": [16], "PA": [16], "SC": [16],
    "PE": [15], "RC": [11], "TI": [11], "CN": [9], "AS": [10], "MS": [12], "PT": [7]
  };

  const gestorMunicipal = rangoValores[0];
  const fechaSolicitud = rangoValores[1];
  const tipoDocumento = rangoValores[2];
  const documento = rangoValores[3];
  const nombrePaciente = rangoValores[5];
  const numeroTelefono = rangoValores[6];

  if(!esCampoVacio(gestorMunicipal)){
    if(!esTextoValido(gestorMunicipal)) errores.push("- El dato (GESTOR MUNICIPAL) no puede contener números o caracteres especiales");
  }
   
  if(!esCampoVacio(fechaSolicitud)) {
    if(!esFechaValida(fechaSolicitud)){
      errores.push("- El dato (FECHA DE SOLICITUD) no se encuentra digitadao correctamente");
    } else {

      const spreadsheetTimeZone = hoja.getParent().getSpreadsheetTimeZone();
      const hoy = new Date();
        
      // Convertir ambas fechas a cadenas en la zona horaria de la hoja
      const fechaActualStr = Utilities.formatDate(hoy, spreadsheetTimeZone, "yyyy-MM-dd");
      const fechaSolicitudStr = Utilities.formatDate(fechaSolicitud, spreadsheetTimeZone, "yyyy-MM-dd");
      
      Logger.log(fechaSolicitud);
      if (fechaActualStr > fechaSolicitudStr) {
          errores.push(`- El dato (FECHA DE SOLICITUD) no puede ser menor que la fecha actual: ${fechaActualStr}`);
      }
    }
  }
  
  if(!esNumeroValido(documento, true)){
    errores.push("- El dato (DOCUMENTO) no se encuentra digitado correctamente")
  } else if (tipoDocumento in longitudesDocumento){
    const logintudEsperada = longitudesDocumento[tipoDocumento];
    if(!esCampoVacio(documento)){
      if(!logintudEsperada.includes(documento.toString().length)) errores.push(`- El dato (DOCUMENTO) debe tener ${logintudEsperada} dígitos según el tipo ${tipoDocumento}.`);
    }
  }

  if(!esCampoVacio(nombrePaciente)){
    if(!esTextoValido(nombrePaciente)) errores.push("- El dato (NOMBRE PACIENTE) no puede contener números o caracteres especiales");
  }

  if(!esNumeroValido(numeroTelefono, true)){
    errores.push("- El dato (NUMERO DE TELEFONO) no se encuentra digitado correctamente")
  } else {
    if (!esCampoVacio(numeroTelefono)){
      if(numeroTelefono.toString().length != 10) errores.push("- El dato (NUMERO DE TELEFONO) debe tener 10 digitos");
    }
  }

  if(errores.length > 0) {
    SpreadsheetApp.getUi().alert("Errores encontrados:\n" + errores.join("\n"));
    return true;
  }

  return false;
}

function validarCampoCita(){
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if(!hoja){
    SpreadsheetApp.getUi().alert("No se encontro una hoja activa");
  }

  const fila =  hoja.getActiveCell().getRow();
  const rangoValores = hoja.getRange(fila, 14, 1, 1).getValues()[0];

  const fechaSolicitud = hoja.getRange(fila, 3, 1, 1).getValue();
  const fechaCita = rangoValores[0];
  
  let errores = []

  if(!esCampoVacio(fechaCita)){
    if(!esFechaValida(fechaCita)) {
      errores.push("- El dato (FECHA DE CITA) no se encuentra digitado correctamente");
    } else if (fechaCita < fechaSolicitud){
        const fechaSolicitudStr = Utilities.formatDate(fechaSolicitud, Session.getScriptTimeZone(), "dd/MM/yyyy");
        errores.push(`- El dato (FECHA DE CITA) no puede ser menor que la (FECHA SOLICITUD): ${fechaSolicitudStr}`);
    }
  }

  if(errores.length > 0) {
    SpreadsheetApp.getUi().alert("Errores encontrados:\n" + errores.join("\n"))
    return true;
  }

}


