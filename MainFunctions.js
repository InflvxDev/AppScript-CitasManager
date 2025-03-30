function bloquearRegistroUsuarios() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if(!hoja){
      SpreadsheetApp.getUi().alert("No se encontro una hoja activa");
  }
  
  const celdaActiva = hoja.getActiveCell();
  const fila = celdaActiva.getRow();
  const archivo = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  const propietario = archivo.getOwner().getEmail();
  const rangoValores = hoja.getRange(fila, 2, 1, 11).getValues()[0];

  if(validaciones(fila, hoja, rangoValores, true, 9) || validarCamposUsuario()){
    return;
  }

  const observacion = hoja.getRange(fila, 13, 1, 1).getValue();
  if (esCampoVacio(observacion)){
    hoja.getRange(fila, 13, 1, 1).setValue("NO APLICA");
  } else if (!esTextoValido(observacion)){
    SpreadsheetApp.getUi().alert("Errores encontrados:\n" + "- El campo (OBSERVACIONES) no puede contener números o caracteres especiales")
  }

  const codigoCita = generarCodigoCita();
  hoja.getRange(fila,1,1,1).setValue(codigoCita);

  const rangoAI = hoja.getRange(fila, 1, 1, 10);
  const rangoKL = hoja.getRange(fila, 12, 1, 2);  
  const proteccionAI = rangoAI.protect().setDescription("Bloqueado automáticamente");
  const proteccionKl = rangoKL.protect().setDescription("Bloqueado automáticamente");

  const valoresRegistro = hoja.getRange(fila, 1,1,13).getValues()[0];
  anexarRegistroUsuario(valoresRegistro);

  SpreadsheetApp.getUi().alert("El usuario N° " + fila + " ha sido registrado."); 
}

function bloquearRegistroAsistencia() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
  if(!hoja){
      SpreadsheetApp.getUi().alert("No se encontro una hoja activa");
  }
  
  const celdaActiva = hoja.getActiveCell();
  const fila = celdaActiva.getRow();
  const archivo = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  const propietario = archivo.getOwner().getEmail();
  const rangoValores = hoja.getRange(fila, 14, 1, 2).getValues()[0];

  if(validaciones(fila, hoja, rangoValores, false, 2) || validarCampoCita()){
    return;
  }

  const rango = hoja.getRange(fila, 14, 1, 2); 
  const proteccion = rango.protect().setDescription("Bloqueado automáticamente");

  const valoresRegistro = hoja.getRange(fila, 14,1,2).getValues()[0];
  const valorPK = hoja.getRange(fila, 1,1,1).getValue();

  anexarRegistroCitas(valoresRegistro , valorPK);

  SpreadsheetApp.getUi().alert("la Asistencia N° " + fila + " ha sido registrada.");  
 
}

function anexarRegistroUsuario(valores){
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("GENERAL");
  if (!hoja) {
    SpreadsheetApp.getUi().alert("No se encontró la hoja 'GENERAL'.");
    return;
  }

  let fila = 2; 
  const ultimaFila = hoja.getLastRow();

  while (fila <= ultimaFila) {
    const valoresRegistro = hoja.getRange(fila, 1, 1, 13).getValues()[0];

    if (valoresRegistro.every(celda => celda === "" || celda === null)) {
      break;
    }
    fila++;
  }
  
  hoja.getRange(fila, 1, 1, valores.length).setValues([valores]);
  const rango = hoja.getRange(fila, 1, 1, valores.length);
  const proteccion = rango.protect().setDescription("Bloqueado AGeneral");

  const propietario = Session.getEffectiveUser(); // Obtener propietario

  proteccion.removeEditors(proteccion.getEditors()); // Remover todos los editores
  proteccion.addEditor(propietario); // Mantener solo al propietario

}

function anexarRegistroCitas(valores,primaryKey){
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("GENERAL");
  if (!hoja) {
    SpreadsheetApp.getUi().alert("No se encontró la hoja 'GENERAL'.");
    return;
  }

  let fila = 2; 
  const ultimaFila = hoja.getLastRow();

  while (fila <= ultimaFila) {
    const primaryKeyGeneral = hoja.getRange(fila, 1, 1, 1).getValue();

    if (primaryKey === primaryKeyGeneral) {
      break;
    }
    fila++;
  }

  hoja.getRange(fila, 14, 1, valores.length).setValues([valores]);
  const rango = hoja.getRange(fila, 14, 1, valores.length);
  const proteccion = rango.protect().setDescription("Bloqueado AGeneral");

  const propietario = Session.getEffectiveUser(); // Obtener propietario

  proteccion.removeEditors(proteccion.getEditors()); // Remover todos los editores
  proteccion.addEditor(propietario); // Mantener solo al propietario

}
