function validaciones (fila, hoja, rangoValores, bloqueandoAK, numColumnas){

  if (fila === 1) {
    SpreadsheetApp.getUi().alert("No puedes registar el encabezado.");
    return true;
  }

  const protecciones = hoja.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  let protegidoPrimeraParte = false;
  let protegidoSegundaParte = false;

  for (let i = 0; i < protecciones.length; i++) {
    const rango = protecciones[i].getRange();
    const columnaInicio = rango.getColumn();
    const numColumnasProtegidas = rango.getNumColumns();

    if(columnaInicio !== 11){
      if (columnaInicio >= 2 && (columnaInicio + numColumnasProtegidas - 1) <= 13 && rango.getRow() === fila) {
        protegidoPrimeraParte = true;
      }
    }
  
    if (columnaInicio >= 14 && (columnaInicio + numColumnasProtegidas - 1) <= 15 && rango.getRow() === fila) {
      protegidoSegundaParte = true;
    }
  }

  // Validaciones según el caso
  if (bloqueandoAK) {
    if (protegidoPrimeraParte) {
      SpreadsheetApp.getUi().alert("El usuario N° " + fila + " ya está registrado.");
      return true;
    }
  } else { // bloqueando LM
    if (!protegidoPrimeraParte) {
      SpreadsheetApp.getUi().alert("No puedes registrar la asistencia N°"+ fila +" sin antes registrar el usuario.");
      return true;
    }
    if (protegidoSegundaParte) {
      SpreadsheetApp.getUi().alert("la asistencia N° " + fila + " ya está registrada.");
      return true;
    }
  }

  const camposVacios = [];
  const columnaInicial = numColumnas === 9 ? 2 : 12;
  for (var i = 0; i < rangoValores.length; i++) {
    if (rangoValores[i] === "" || rangoValores[i] === null) {
      const columnaReal = String.fromCharCode(65 + (columnaInicial - 1) + i);
      camposVacios.push(columnaReal);
    }
  }

  if (camposVacios.length > 0) {
    SpreadsheetApp.getUi().alert("Asegúrese de que se encuentre en el registro correspondiente, ya que faltan campos por llenar: " + camposVacios.join(", "));
    return true;
  }

}

function esCampoVacio (valor){
  return valor == ""? true : false;
}

function esNumeroValido (valor, esEntero = false){
  let numero = Number(valor);
  if(isNaN(numero)) return false;
  return esEntero ? Number.isInteger(numero): true;
}

function esFechaValida (valor){
  return valor instanceof Date && !isNaN(valor);
}

function esTextoValido(valor) {
  var regex = /^[A-Za-zÑñ\s]+$/; // Solo permite letras y espacios
  return typeof valor === "string" && regex.test(valor)
}