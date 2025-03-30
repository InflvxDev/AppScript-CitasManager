function generarCodigoCita() {
  const caracteres = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
  let codigo = "";

  for(let i=0; i<16; i++){
    const indice = Math.floor(Math.random() * caracteres.length);
    codigo += caracteres.charAt(indice);
  }

  return codigo;
}

function removerEditoresNoPropietarios(protecciones, propietario) {
  protecciones.forEach(function(proteccion) {
    var editores = proteccion.getEditors();

    editores.forEach(function(editor) {
      if (editor.getEmail() !== propietario) {
        proteccion.removeEditor(editor);
      }
    });

    if (proteccion.canDomainEdit()) {
      proteccion.setDomainEdit(false);
    }
  });
}

