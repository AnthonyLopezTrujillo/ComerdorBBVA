const CALENDARIO_ID = "";

//INDICES DE LA DATA EXTRAIDA CON executeQuery
const FILA_FECHAS_SEMANA_1_IDX = 0;
const FILA_FIN_SEMANA_1_IDX = 31;
const FILA_FECHAS_SEMANA_2_IDX = 32;
const FILA_FIN_SEMANA_2_IDX = 63;
const FILA_FECHAS_SEMANA_3_IDX = 64;
const FILA_FIN_SEMANA_3_IDX = 95;
const FILA_FECHAS_SEMANA_4_IDX = 96;
const FILA_FIN_SEMANA_4_IDX = 127;
const FILA_FECHAS_SEMANA_5_IDX = 128;
const FILA_FIN_SEMANA_5_IDX = 159;


//INDICES DE LA HOJA
const HOJA_COLUMNA_SEMANAS_DATOS_IDX = 1;
const HOJA_FILA_SEMANA_1_DATOS_IDX = 2;
const HOJA_FILA_SEMANA_2_DATOS_IDX = 35;
const HOJA_FILA_SEMANA_3_DATOS_IDX = 68;
const HOJA_FILA_SEMANA_4_DATOS_IDX = 101;
const HOJA_FILA_SEMANA_5_DATOS_IDX = 134;

const CANTIDAD_FILAS_DATOS_X_SEMANA = 31;
const CANTIDAD_COLUMNAS_DATOS_X_SEMANA = 10;




function verificarUsuarioConPermiso(emailUsuario) {
  var permiso = false;

  var emailLowerCase = emailUsuario.toLowerCase();

  if (emailLowerCase == "") {
    permiso = true;

  } else {
    SpreadsheetApp.getUi().alert('No tiene permisos para ejecutar esta acción');
  }

  return permiso;
}

function capitalizarPrimeraPalabra(texto) {
  // Convertimos todo el texto a minúsculas
  var textoMinuscula = texto.toLowerCase();

  // Dividimos el texto en palabras usando el espacio como separador
  var palabras = textoMinuscula.split(" ");

  // Capitalizamos la primera letra de la primera palabra
  palabras[0] = palabras[0].charAt(0).toUpperCase() + palabras[0].slice(1);

  // Unimos las palabras nuevamente en una sola cadena
  return palabras.join(" ");
}


function convertirStringAFecha(fechaTexto) {
  // Extraer la parte del día, mes y año
  var partes = fechaTexto.split(",")[1].trim(); // Obtiene "14 de octubre de 2024"
  var partesFecha = partes.split(" "); // Separa el día, el "de" y el mes

  var dia = parseInt(partesFecha[0]); // "14"
  var mesTexto = partesFecha[2].toLowerCase(); // "octubre"
  var anio = parseInt(partesFecha[4]); // "2024"

  // Convertir el nombre del mes a su valor numérico (0-indexado)
  var meses = {
    "enero": 0, "febrero": 1, "marzo": 2, "abril": 3, "mayo": 4, "junio": 5,
    "julio": 6, "agosto": 7, "septiembre": 8, "octubre": 9, "noviembre": 10, "diciembre": 11
  };

  var mes = meses[mesTexto]; // Convertir "octubre" a 9

  // Crear el objeto Date
  var fecha = new Date(anio, mes, dia);

  return fecha; // Devuelve la fecha en formato Date
}

function convertirFechaStringConGuionAFecha(fechaString) {
  var partesFecha = fechaString.split('-');
  var anio = parseInt(partesFecha[0]);
  var mes = parseInt(partesFecha[1]) - 1;
  var dia = parseInt(partesFecha[2]);

  var fechaObjeto = new Date(anio, mes, dia);
  return fechaObjeto;
}

function procesarSemanas(data, calendario, eventosGenerados, fechaInicio, fechaFin) {
  // Convierte las fechas de inicio y fin a objetos de fecha
  const fechaInicioEventos = convertirFechaStringConGuionAFecha(fechaInicio);
  const fechaFinEventos = convertirFechaStringConGuionAFecha(fechaFin);

  // Constantes de las posiciones de las semanas
  const semanas =
    [{ inicio: FILA_FECHAS_SEMANA_1_IDX, fin: FILA_FIN_SEMANA_1_IDX },
    { inicio: FILA_FECHAS_SEMANA_2_IDX, fin: FILA_FIN_SEMANA_2_IDX },
    { inicio: FILA_FECHAS_SEMANA_3_IDX, fin: FILA_FIN_SEMANA_3_IDX },
    { inicio: FILA_FECHAS_SEMANA_4_IDX, fin: FILA_FIN_SEMANA_4_IDX },
    { inicio: FILA_FECHAS_SEMANA_5_IDX, fin: FILA_FIN_SEMANA_5_IDX }
    ];

  // Iterar sobre todas las semanas
  for (var itemSemana = 0; itemSemana < semanas.length; itemSemana++) {
    var filaInicio = semanas[itemSemana].inicio;
    var filaFin = semanas[itemSemana].fin;
    console.log("Procesando semana: " + (itemSemana + 1));

    // Recorrer columnas (días de la semana: Lunes a Viernes)
    for (var i = 1; i < data[filaInicio].length; i += 2) {
      var fechaEvento = data[filaInicio][i];
      console.log("fechaEvento:: " + fechaEvento);

      fechaEvento = convertirStringAFecha(fechaEvento);

      // Verificar si la fecha del evento está dentro del rango
      if (fechaEvento < fechaInicioEventos || fechaEvento > fechaFinEventos) {
        Logger.log("Evento fuera de rango: " + fechaEvento);
        continue; // Saltar eventos fuera del rango
      }



      if (calendario.getEventsForDay(fechaEvento).length > 0) continue;

      var title = 'Almuerzos Comedor';
      var desc = '';

      // Detectar si es miércoles (columna F)
      var esMiercoles = i === 5;

      if(esMiercoles){
        title = "Arma tu menú";
      }

      // Recorremos las filas dentro del rango de la semana
      for (var j = filaInicio + 1; j <= filaFin; j++) {
        var platillo = data[j][i] || ''; // Obtener platillo

        console.log("platillo:: " + platillo);//ELIMINAR
        if (platillo != '') {
          platillo = capitalizarPrimeraPalabra(platillo);
          console.log("platillo:: " + platillo); //ELIMINAR
        }

        var calorias = data[j][i + 1] || ''; // Obtener calorías del platillo

        // Solo añade a la descripción si hay datos disponibles en platillo o calorías
        if (platillo) {
          if (esMiercoles) {
            switch (j - filaInicio) {
              case 1:
                desc += `🥘 <b>Entradas:</b>\n-${platillo} (${calorias} kcal)`;
                break;
              case 2:
                if (!desc.includes("Entradas:")) {
                  desc += `🥘 <b>Entradas:</b>`;
                }
                desc += `\n-${platillo} (${calorias} kcal)`;
                break;
              case 3:
                if (!desc.includes("Entradas:")) {
                  desc += `🥘 <b>Entradas:</b>`;
                }
                desc += `\n-${platillo} (${calorias} kcal)`;
                break;
              case 6:
                desc += `\n\n🍗 <b>Proteínas:</b>\n-${platillo} (${calorias} kcal)`;
                break;
              case 7:
                if (!desc.includes("Proteínas:")) {
                  desc += `\n\n🍗 <b>Proteínas:</b>`;
                }
                desc += `\n-${platillo} (${calorias} kcal)`;
                break;
              case 8:
                desc += `\n\n🍚 <b>Complementos:</b>\n-${platillo} (${calorias} kcal)`;
                break;
              case 9:
                if (!desc.includes("Complementos:")) {
                  desc += `\n\n🍚 <b>Complementos:</b>`;
                }
                desc += `\n-${platillo} (${calorias} kcal)`;
                break;
              case 10:
                if (!desc.includes("Complementos:")) {
                  desc += `\n\n🍚 <b>Complementos:</b>`;
                }
                desc += `\n-${platillo} (${calorias} kcal)`;
                break;
              case 11:
                if (!desc.includes("Complementos:")) {
                  desc += `\n\n🍚 <b>Complementos:</b>`;
                }
                desc += `\n-${platillo} (${calorias} kcal)`;
                break;
              case 12:
                desc += `\n\n🍨 <b>Postre:</b>\n-${platillo} (${calorias} kcal)`;
                break;
              case 13:
                if (!desc.includes("Postre:")) {
                  desc += `\n\n🍨 <b>Postre:</b>`;
                }
                desc += `\n-${platillo} (${calorias} kcal)`;
                break;
              case 14:
                if (!desc.includes("Postre:")) {
                  desc += `\n\n🍨 <b>Postre:</b>`;
                }
                desc += `\n-${platillo} (${calorias} kcal)`;
                break;
              case 15:
                desc += `\n\n🥤 <b>Refresco:</b>\n-${platillo} (${calorias} kcal)`;
                break;
              case 16:
                desc += `\n\n🥗 <b>Menú Hipocalórico:</b>\n${platillo} (${calorias} kcal)`;
                break;
              case 17:
                desc += `\n\n🥙 <b>Festival de Ensaladas:</b>\n${platillo} (${calorias} kcal)`;
                break;
              case 18:
                desc += `\n\n🍱 <b>Barra:</b>\n<i>Proteínas:</i>\n -${platillo} (${calorias} kcal)`;
                break;
              case 19:
                if (!desc.includes("Barra:")) {
                  desc += `\n\n🍱 <b>Barra:</b>\n<i>Proteínas:</i>`;
                }
                desc += `\n -${platillo} (${calorias} kcal)`;
                break;
              case 20:
                desc += `\n<i>Bases:</i>\n -${platillo} (${calorias} kcal)`;
                break;
              case 21:
                if (!desc.includes("Bases:")) {
                  desc += `\n<i>Bases:</i>`;
                }
                desc += ` \n -${platillo} (${calorias} kcal)`;
                break;
              case 22:
                if (!desc.includes("Bases:")) {
                  desc += `\n<i>Bases:</i>`;
                }
                desc += ` \n -${platillo} (${calorias} kcal)`;
                break;
              case 23:
                desc += `\n<i>Toppins:</i>\n -${platillo} (${calorias} kcal)`;
                break;
              case 24:
                if (!desc.includes("Toppins:")) {
                  desc += `\n<i>Toppins:</i>`;
                }
                desc += `\n -${platillo} (${calorias} kcal)`;
                break;
              case 25:
                if (!desc.includes("Toppins:")) {
                  desc += `\n<i>Toppins:</i>`;
                }
                desc += `\n -${platillo} (${calorias} kcal)`;
                break;
              case 26:
                if (!desc.includes("Toppins:")) {
                  desc += `\n<i>Toppins:</i>`;
                }
                desc += `\n -${platillo} (${calorias} kcal)`;
                break;
              case 27:
                if (!desc.includes("Toppins:")) {
                  desc += `\n<i>Toppins:</i>`;
                }
                desc += `\n -${platillo} (${calorias} kcal)`;
                break;
              case 28:
                if (!desc.includes("Toppins:")) {
                  desc += `\n<i>Toppins:</i>`;
                }
                desc += `\n -${platillo} (${calorias} kcal)`;
                break;
              case 29:
                desc += `\n<i>Salsas:</i>\n -${platillo}`;
                break;
              case 30:
                if (!desc.includes("Salsas:")) {
                  desc += `\n<i>Salsas:</i>`;
                }
                desc += `\n -${platillo}`;
                break;
              case 31:
                if (!desc.includes("Salsas:")) {
                  desc += `\n<i>Salsas:</i>`;
                }
                desc += `\n -${platillo}`;
                break;
              default:
                break;
            }
          } else {
            // Plantilla para los demás días (lunes, martes, jueves, viernes)
            switch (j - filaInicio) {
              case 1:
                desc += `<b>Entradas:</b>\n🥙 ${platillo} (${calorias} kcal)`;
                break;
              case 3:
                if (!desc.includes("Entradas:")) {
                  desc += `🥙 <b>Entradas:</b>`;
                }
                desc += `\n🥘 ${platillo} (${calorias} kcal)`;
                break;
              case 4:
                desc += `\n\n<b>Fondos:</b>\n<b>🍛 Menú ejecutivo: </b>${platillo} (${calorias} kcal)`;
                break;
              case 5:
                if (!desc.includes("Fondos:")) {
                  desc += `\n\n<b>Fondos:</b>`;
                }
                desc += `\n🍲 <b>Menú económico: </b>${platillo} (${calorias} kcal)`;
                break;
              case 12:
                desc += `\n\n🍨 <b>Postres:</b> ${platillo} (${calorias} kcal)`;
                break;
              case 15:
                desc += `\n🥤 <b>Refresco:</b> ${platillo} (${calorias} kcal)`;
                break;
              case 16:
                desc += `\n\n🥗 <b>Menú Hipocalórico:</b>\n${platillo} (${calorias} kcal)`;
                break;
              case 17:
                desc += `\n\n🥙 <b>Festival de Ensaladas:</b>\n${platillo} (${calorias} kcal)`;
                break;
              case 18:
                desc += `\n\n🍱 <b>Barra:</b>\n<i>Proteínas:</i>\n -${platillo} (${calorias} kcal)`;
                break;
              case 19:
                if (!desc.includes("Barra:")) {
                  desc += `\n\n🍱 <b>Barra:</b>\n<i>Proteínas:</i>`;
                }
                desc += `\n -${platillo} (${calorias} kcal)`;
                break;
              case 20:
                desc += `\n<i>Bases:</i>\n -${platillo} (${calorias} kcal)`;
                break;
              case 21:
                if (!desc.includes("Bases:")) {
                  desc += `\n<i>Bases:</i>`;
                }
                desc += ` \n -${platillo} (${calorias} kcal)`;
                break;
              case 22:
                if (!desc.includes("Bases:")) {
                  desc += `\n<i>Bases:</i>`;
                }
                desc += ` \n -${platillo} (${calorias} kcal)`;
                break;
              case 23:
                desc += `\n<i>Toppins:</i>\n -${platillo} (${calorias} kcal)`;
                break;
              case 24:
                if (!desc.includes("Toppins:")) {
                  desc += `\n<i>Toppins:</i>`;
                }
                desc += `\n -${platillo} (${calorias} kcal)`;
                break;
              case 25:
                if (!desc.includes("Toppins:")) {
                  desc += `\n<i>Toppins:</i>`;
                }
                desc += `\n -${platillo} (${calorias} kcal)`;
                break;
              case 26:
                if (!desc.includes("Toppins:")) {
                  desc += `\n<i>Toppins:</i>`;
                }
                desc += `\n -${platillo} (${calorias} kcal)`;
                break;
              case 27:
                if (!desc.includes("Toppins:")) {
                  desc += `\n<i>Toppins:</i>`;
                }
                desc += `\n -${platillo} (${calorias} kcal)`;
                break;
              case 28:
                if (!desc.includes("Toppins:")) {
                  desc += `\n<i>Toppins:</i>`;
                }
                desc += `\n -${platillo} (${calorias} kcal)`;
                break;
              case 29:
                desc += `\n<i>Salsas:</i>\n -${platillo}`;
                break;
              case 30:
                if (!desc.includes("Salsas:")) {
                  desc += `\n<i>Salsas:</i>`;
                }
                desc += `\n -${platillo}`;
                break;
              case 31:
                if (!desc.includes("Salsas:")) {
                  desc += `\n<i>Salsas:</i>`;
                }
                desc += `\n -${platillo}`;
                break;
              default:
                break;
            }
          }
        }
      }

      console.log("Descripcion:: " + desc);

      // Añadir el enlace
      if(desc){
        desc += `\n\n🍝 <b>Platos a la carta:</b> <a href="https://drive.">Link</a>👆🏻`;
      }

      // Crear el evento si la descripción está completa
      if (desc) {
        Logger.log("Título del evento: " + title);
        Logger.log("Fecha del evento: " + fechaEvento);
        Logger.log("Descripción del evento:\n" + desc);

        calendario.createAllDayEvent(title, fechaEvento, { description: desc });
        eventosGenerados++;
      } else {
        Logger.log("Evento omitido: No hay datos suficientes para " + fechaEvento);
      }
    }
  }

  Logger.log("Total eventos generados: " + eventosGenerados);
}




