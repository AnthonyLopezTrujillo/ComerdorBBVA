<!DOCTYPE html>
<html>

<head>
  <base target="_top">
  <style>
    body {
      font-family: 'Lato', sans-serif;
      margin: 0;
      padding: 10px;
      background-color: #f2f2f2;
    }

    .contenedor {
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      min-height: 100px;
      border: 1px solid #cccccc;
      border-radius: 5px;
      padding: 15px;
      background-color: #fff;
    }

    .titulo {
      font-size: 20px;
      color: #004481;
      text-align: center;
      margin-bottom: 10px;
    }

    .etiquetas {
      margin-bottom: 10px;
    }

    input[type="date"] {
      padding: 3px;
      border: 1px solid #cccccc;
      border-radius: 3px;
      margin-bottom: 28px;
    }

    .boton {
      background-color: #ff0000;
      color: #fff;
      border: none;
      padding: 10px 20px;
      border-radius: 5px;
      cursor: pointer;
    }
  </style>
</head>

<body>
  <div class="contenedor">
    <h2 class="titulo">Ingresa las fechas</h2>
    <div class="etiquetas">
      <label for="fechaInicio">Fecha Inicioo:</label>
    </div>
    <input type="date" id="fechaInicio">
    <div class="etiquetas">
      <label for="fechaFin">Fecha Fin:</label>
    </div>
    <input type="date" id="fechaFin">
    <button class="boton" onclick="eliminarEventos()">Eliminar Eventos</button>
  </div>

  <script>
    function eliminarEventos() {
      var fechaInicio = document.getElementById('fechaInicio').value;
      var fechaFin = document.getElementById('fechaFin').value;
      google.script.run.withSuccessHandler(function () {
        google.script.host.close();
      }).borrarEventos(fechaInicio,fechaFin);
    }
  </script>
</body>

</html>
