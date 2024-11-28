function crearEventos(fechaInicio, fechaFin) {

  //Obtener sesion de usuario
  var emailUsuario = Session.getActiveUser().getEmail();
  console.log("EmailUsuario:: " + emailUsuario);

  //Verificar usuarios con permisos de ejecución
  var usuarioPermiso = verificarUsuarioConPermiso(emailUsuario);

  if (usuarioPermiso) {
    SpreadsheetApp.getUi().alert('Vamos a generar los menús de comedores');
    var calendario = CalendarApp.getCalendarById(CALENDARIO_ID);
    var eventosGenerados = 0;
    var data = obtenerCalendariosMenus();
    if (data != null && data.length > 0) {
      procesarSemanas(data, calendario, eventosGenerados,fechaInicio,fechaFin);
    }
  } else {
    console.log("Persona sin permiso para agregar nuevos eventos al calendar")
    return
  }
}

function test(){
  borrarEventos("02/12/2024","02/12/2024");
}

function borrarEventos(fechaInicio, fechaFin) {
  console.log("en borrarEventos: "+fechaInicio+" - "+fechaFinEvento);
  // Obtener sesión de usuario
  var emailUsuario = Session.getActiveUser().getEmail();
  console.log("EmailUsuario:: " + emailUsuario);

  // Verificar usuarios con permisos de ejecución
  var usuarioPermiso = verificarUsuarioConPermiso(emailUsuario);
  if (usuarioPermiso) {
    //Obtener Calendario
    SpreadsheetApp.getUi().alert('Vamos a eliminar los eventos del calendar.');
    var calendario = CalendarApp.getCalendarById(CALENDARIO_ID);

    //Convertir texto a fecha ("2024-10-16" -> formato fecha)
    var fechaInicioEvento = convertirFechaStringConGuionAFecha(fechaInicio);
    var fechaFinEvento = convertirFechaStringConGuionAFecha(fechaFin);

    // Ajustar fechaFinEvento para incluir todo el día
    fechaFinEvento.setHours(23, 59, 59);

    // Obtener eventos dentro del rango de fechas
    var eventos = calendario.getEvents(fechaInicioEvento, fechaFinEvento);
    console.log("Eventos encontrados: " + eventos.length);

    // Eliminar eventos
    for (var itemEvento = 0; itemEvento < eventos.length; itemEvento++) {
      //console.log("Eliminando evento: " + eventos[itemEvento].getTitle() + " en fecha: " + eventos[itemEvento].getStartTime());
      eventos[itemEvento].deleteEvent();
    }
    //console.log("Se eliminaron " + eventos.length + " eventos.");
  } else {
    console.log("Persona sin permiso para eliminar eventos del calendario.");
    return;
  }
}

function borrarTextoBaseMenu(){
  SHEET_CALENDARIOS_MENUS.getRange(HOJA_FILA_SEMANA_1_DATOS_IDX + 1,HOJA_COLUMNA_SEMANAS_DATOS_IDX+1,CANTIDAD_FILAS_DATOS_X_SEMANA,CANTIDAD_COLUMNAS_DATOS_X_SEMANA).clearContent();

  SHEET_CALENDARIOS_MENUS.getRange(HOJA_FILA_SEMANA_2_DATOS_IDX + 1,HOJA_COLUMNA_SEMANAS_DATOS_IDX+1,CANTIDAD_FILAS_DATOS_X_SEMANA,CANTIDAD_COLUMNAS_DATOS_X_SEMANA).clearContent();

  SHEET_CALENDARIOS_MENUS.getRange(HOJA_FILA_SEMANA_3_DATOS_IDX + 1,HOJA_COLUMNA_SEMANAS_DATOS_IDX+1,CANTIDAD_FILAS_DATOS_X_SEMANA,CANTIDAD_COLUMNAS_DATOS_X_SEMANA).clearContent();

  SHEET_CALENDARIOS_MENUS.getRange(HOJA_FILA_SEMANA_4_DATOS_IDX + 1,HOJA_COLUMNA_SEMANAS_DATOS_IDX+1,CANTIDAD_FILAS_DATOS_X_SEMANA,CANTIDAD_COLUMNAS_DATOS_X_SEMANA).clearContent();

  SHEET_CALENDARIOS_MENUS.getRange(HOJA_FILA_SEMANA_5_DATOS_IDX + 1,HOJA_COLUMNA_SEMANAS_DATOS_IDX+1,CANTIDAD_FILAS_DATOS_X_SEMANA,CANTIDAD_COLUMNAS_DATOS_X_SEMANA).clearContent();
}

