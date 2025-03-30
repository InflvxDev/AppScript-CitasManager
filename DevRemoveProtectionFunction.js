function eliminarProteccionesBloqueadas() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet();

  // Obtener todas las protecciones de rangos
  var proteccionesRangos = hoja.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  proteccionesRangos.forEach(function(proteccion) {
    if (proteccion.getDescription() === "Bloqueado automáticamente" || proteccion.getDescription() === "Bloqueado AGeneral") {
      proteccion.remove();
    }
  });

}

