
/************************************************************
 * Sistema de Evaluación Online
 * Backend principal - Google Apps Script
 *
 * BASE DE DATOS: Google Sheets
 * BACKEND: Google Apps Script
 * FRONTEND: Index.html
 *
 * SEGURIDAD:
 * - Las respuestas correctas nunca se envían al navegador.
 * - La corrección se realiza en el servidor.
 * - Los intentos se controlan en el servidor.
 * - Se validan las respuestas recibidas.
 * - Se bloquean los envíos simultáneos.
 ************************************************************/


/* =========================================================
   1. CONFIGURACIÓN GENERAL
========================================================= */

const CONFIG = {

  HOJAS: {
    CONFIGURACION: "Configuracion",
    EVALUACIONES: "Evaluaciones",
    PREGUNTAS: "Preguntas",
    OPCIONES: "Opciones",
    ESTUDIANTES: "Estudiantes",
    INTENTOS: "Intentos",
    RESPUESTAS: "Respuestas",
    RESULTADOS: "Resultados"
  },

  ESTADOS: {
    PUBLICADA: "PUBLICADA",
    ACTIVO: "ACTIVO"
  },

  NIVELES: {
    INICIAL: "Inicial",
    DESARROLLO: "Desarrollo",
    LOGRADO: "Logrado"
  }

};


/* =========================================================
   2. UTILIDADES GENERALES
========================================================= */

function obtenerHoja_(nombre) {

  const hoja = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(nombre);

  if (!hoja) {
    throw new Error('No existe la hoja "' + nombre + '".');
  }

  return hoja;

}


function obtenerEncabezados_(hoja) {

  if (hoja.getLastColumn() === 0) {
    return [];
  }

  return hoja
    .getRange(1, 1, 1, hoja.getLastColumn())
    .getDisplayValues()[0]
    .map(valor => String(valor).trim());

}


function hojaAObjetos_(nombreHoja) {

  const hoja = obtenerHoja_(nombreHoja);

  const ultimaFila = hoja.getLastRow();
  const ultimaColumna = hoja.getLastColumn();

  if (ultimaFila < 2 || ultimaColumna === 0) {
    return [];
  }

  const encabezados = obtenerEncabezados_(hoja);

  const datos = hoja
    .getRange(2, 1, ultimaFila - 1, ultimaColumna)
    .getValues();

  return datos
    .filter(fila => fila.some(valor => valor !== ""))
    .map(fila => {

      const objeto = {};

      encabezados.forEach((encabezado, indice) => {

        if (encabezado) {
          objeto[encabezado] = fila[indice];
        }

      });

      return objeto;

    });

}


/**
 * Guarda un objeto utilizando los nombres de las columnas.
 * No depende del orden de los encabezados.
 */
function insertarObjeto_(nombreHoja, objeto) {

  const hoja = obtenerHoja_(nombreHoja);
  const encabezados = obtenerEncabezados_(hoja);

  if (encabezados.length === 0) {
    throw new Error(
      "La hoja " + nombreHoja + " no tiene encabezados."
    );
  }

  const faltantes = Object.keys(objeto).filter(
    campo => !encabezados.includes(campo)
  );

  if (faltantes.length > 0) {
    throw new Error(
      "Faltan columnas en " + nombreHoja + ": " +
      faltantes.join(", ")
    );
  }

  const fila = encabezados.map(encabezado => {

    return Object.prototype.hasOwnProperty.call(
      objeto,
      encabezado
    )
      ? objeto[encabezado]
      : "";

  });

  hoja.appendRow(fila);

}


/** Agrega columnas nuevas al final sin alterar las existentes. */
function asegurarColumnas_(nombreHoja, columnas) {
  const hoja = obtenerHoja_(nombreHoja);
  const encabezados = obtenerEncabezados_(hoja);
  const faltantes = columnas.filter(c => !encabezados.includes(c));
  if (faltantes.length) {
    hoja.getRange(1, encabezados.length + 1, 1, faltantes.length)
      .setValues([faltantes]);
  }
}

function buscarColumna_(encabezados, nombre) {

  const indice = encabezados.indexOf(nombre);

  if (indice === -1) {
    throw new Error(
      'No se encontró la columna "' + nombre + '".'
    );
  }

  return indice;

}


function normalizarTexto_(valor) {

  if (valor === null || valor === undefined) {
    return "";
  }

  return String(valor)
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");

}


function generarId_(prefijo) {

  return prefijo + "_" +
    Utilities.getUuid()
      .replace(/-/g, "")
      .substring(0, 12)
      .toUpperCase();

}


function fechaActual_() {

  return new Date();

}


function esActivo_(valor) {

  return normalizarTexto_(valor) === "activo";

}


function esPublicada_(valor) {

  return normalizarTexto_(valor) === "publicada";

}


/**
 * Compara cursos permitiendo que una evaluación general del nivel
 * (por ejemplo, "3° básico") corresponda a una sección ("3° básico A").
 * Si ambos cursos especifican sección, estas deben coincidir.
 */
function appCursosCompatibles_(cursoEvaluacion, cursoEstudiante) {

  const separar = valor => {
    const texto = normalizarTexto_(valor).replace(/º/g, "°");
    const coincidencia = texto.match(/^(.*?)(?:\s+([a-z]))?$/i);
    return {
      base: (coincidencia ? coincidencia[1] : texto).trim(),
      seccion: coincidencia && coincidencia[2] ? coincidencia[2].toLowerCase() : ""
    };
  };

  const evaluacion = separar(cursoEvaluacion);
  const estudiante = separar(cursoEstudiante);

  if (!evaluacion.base || evaluacion.base !== estudiante.base) {
    return false;
  }

  return !(evaluacion.seccion && estudiante.seccion) ||
    evaluacion.seccion === estudiante.seccion;

}


/* =========================================================
   3. CONFIGURACIÓN DEL SISTEMA
========================================================= */

function obtenerConfiguracion_(parametro) {

  const datos = hojaAObjetos_(
    CONFIG.HOJAS.CONFIGURACION
  );

  const fila = datos.find(item =>
    normalizarTexto_(item.parametro) ===
    normalizarTexto_(parametro)
  );

  return fila ? fila.valor : null;

}


function obtenerConfiguracionCompleta_() {

  const datos = hojaAObjetos_(
    CONFIG.HOJAS.CONFIGURACION
  );

  const resultado = {};

  datos.forEach(item => {

    if (item.parametro) {
      resultado[item.parametro] = item.valor;
    }

  });

  return resultado;

}


/* =========================================================
 
/* =========================================================
   4. CURSOS
========================================================= */

function obtenerCursos() {

  const estudiantes = hojaAObjetos_(
    CONFIG.HOJAS.ESTUDIANTES
  );

  const cursos = {};

  // Mostrar todos los cursos registrados, aunque actualmente no tengan
  // estudiantes con estado ACTIVO. La validación individual del alumno
  // sigue comprobando su estado antes de permitir el acceso.
  estudiantes.forEach(estudiante => {

    const curso = String(estudiante.curso || "").trim();

    if (curso) {
      cursos[curso] = true;
    }

  });

  return Object.keys(cursos).sort();

}


function obtenerCursosParaIngreso() {

  return obtenerCursos();

}


/* =========================================================
   5. EVALUACIONES
========================================================= */

function obtenerEvaluacionesPublicadas() {

  const evaluaciones = hojaAObjetos_(
    CONFIG.HOJAS.EVALUACIONES
  );

  return evaluaciones

    .filter(evaluacion =>
      esPublicada_(evaluacion.estado)
    )

    .map(evaluacion => ({

      evaluacionId: String(
        evaluacion.evaluacionId || ""
      ),

      titulo: String(
        evaluacion.titulo || ""
      ),

      curso: String(
        evaluacion.curso || ""
      ),

      nivel: String(
        evaluacion.nivel || ""
      ),

      estado: String(
        evaluacion.estado || ""
      )

    }));

}


function obtenerEvaluacionPorId_(evaluacionId) {

  const evaluaciones = hojaAObjetos_(
    CONFIG.HOJAS.EVALUACIONES
  );

  return evaluaciones.find(evaluacion =>
    String(evaluacion.evaluacionId).trim() ===
    String(evaluacionId).trim()
  );

}


/**
 * Obtiene el límite configurado por el docente.
 *
 * 0 = ilimitado.
 * Vacío = 1 intento.
 * Mayor que 0 = límite establecido.
 */
function obtenerLimiteIntentos_(evaluacion) {

  const valor = evaluacion.intentosPermitidos;

  if (
    valor === "" ||
    valor === null ||
    valor === undefined
  ) {
    return 1;
  }

  const limite = Number(valor);

  if (
    !Number.isInteger(limite) ||
    limite < 0
  ) {
    throw new Error(
      "La cantidad de intentos permitidos no es válida."
    );
  }

  return limite;

}


/* =========================================================
   6. ESTUDIANTES
========================================================= */

function buscarEstudiante_(apellidos, nombres, curso) {

  const estudiantes = hojaAObjetos_(
    CONFIG.HOJAS.ESTUDIANTES
  );

  return estudiantes.find(estudiante => {

    return (

      normalizarTexto_(estudiante.apellidos) ===
      normalizarTexto_(apellidos) &&

      normalizarTexto_(estudiante.nombres) ===
      normalizarTexto_(nombres) &&

      normalizarTexto_(estudiante.curso) ===
      normalizarTexto_(curso) &&

      esActivo_(estudiante.estado)

    );

  });

}


function validarEstudiante(apellidos, nombres, curso) {

  if (!apellidos || !nombres || !curso) {

    return {
      correcto: false,
      mensaje: "Debes completar apellidos, nombres y curso."
    };

  }

  const estudiante = buscarEstudiante_(
    apellidos,
    nombres,
    curso
  );

  if (!estudiante) {

    return {
      correcto: false,
      mensaje:
        "No encontramos al estudiante. Verifica los datos ingresados."
    };

  }

  return {

    correcto: true,

    mensaje: "Estudiante validado correctamente.",

    estudiante: {

      estudianteId: String(
        estudiante.estudianteId
      ),

      apellidos: String(
        estudiante.apellidos
      ),

      nombres: String(
        estudiante.nombres
      ),

      curso: String(
        estudiante.curso
      )

    }

  };

}


/* =========================================================
   7. PREGUNTAS Y OPCIONES
========================================================= */


/**
 * Obtiene preguntas sin enviar respuestas correctas.
 */
function obtenerPreguntasPublicas_(evaluacionId) {

  const preguntas = hojaAObjetos_(
    CONFIG.HOJAS.PREGUNTAS
  );

  const opciones = hojaAObjetos_(
    CONFIG.HOJAS.OPCIONES
  );

  const preguntasEvaluacion = preguntas

    .filter(pregunta =>

      String(pregunta.evaluacionId) ===
      String(evaluacionId) &&

      (
        !pregunta.estado ||
        esActivo_(pregunta.estado)
      )

    )

    .sort((a, b) =>
      Number(a.orden || 0) -
      Number(b.orden || 0)
    );

  return preguntasEvaluacion.map(pregunta => {

    const preguntaId = String(
      pregunta.preguntaId
    );

    const opcionesPregunta = opciones

      .filter(opcion =>
        String(opcion.preguntaId) === preguntaId
      )

      .sort((a, b) =>
        Number(a.orden || 0) -
        Number(b.orden || 0)
      );

    return {

      preguntaId: preguntaId,

      parte: String(
        pregunta.parte || ""
      ),

      asignatura: String(
        pregunta.asignatura || ""
      ),

      tipo: String(
        pregunta.tipo || ""
      ),

      orden: Number(
        pregunta.orden || 0
      ),

      enunciado: String(
        pregunta.enunciado || ""
      ),

      puntaje: Number(
        pregunta.puntaje || 0
      ),

      opciones: opcionesPregunta.map(opcion => ({

        opcionId: String(
          opcion.opcionId || ""
        ),

        grupo: String(
          opcion.grupo || ""
        ),

        codigo: String(
          opcion.codigo || ""
        ),

        texto: String(
          opcion.texto || ""
        )

      }))

    };

  });

}


/**
 * Entrega una evaluación al estudiante.
 */
function obtenerEvaluacionPublica(
  apellidos,
  nombres,
  curso,
  evaluacionId
) {

  if (
    !apellidos ||
    !nombres ||
    !curso ||
    !evaluacionId
  ) {

    return {
      correcto: false,
      mensaje: "Faltan datos para cargar la evaluación."
    };

  }

  const estudiante = buscarEstudiante_(
    apellidos,
    nombres,
    curso
  );

  if (!estudiante) {

    return {
      correcto: false,
      mensaje: "Los datos del estudiante no son válidos."
    };

  }

  const evaluacion = obtenerEvaluacionPorId_(
    evaluacionId
  );

  if (!evaluacion || !esPublicada_(evaluacion.estado)) {

    return {
      correcto: false,
      mensaje: "La evaluación no está disponible."
    };

  }

  if (
    !appCursosCompatibles_(evaluacion.curso, curso)
  ) {

    return {
      correcto: false,
      mensaje:
        "La evaluación no corresponde al curso seleccionado."
    };

  }

  const preguntas = obtenerPreguntasPublicas_(
    evaluacionId
  );

  if (preguntas.length === 0) {

    return {
      correcto: false,
      mensaje: "La evaluación no tiene preguntas configuradas."
    };

  }

  /*
   * Consultamos los intentos disponibles.
   */
  const limite = obtenerLimiteIntentos_(evaluacion);

  const realizados = contarIntentosEstudiante_(
    evaluacionId,
    estudiante.estudianteId
  );

  if (limite > 0 && realizados >= limite) {

    return {
      correcto: false,
      mensaje: "Has alcanzado el máximo de intentos permitidos.",
      intentosPermitidos: limite,
      intentosRealizados: realizados
    };

  }

  return {

    correcto: true,

    estudiante: {

      estudianteId: String(
        estudiante.estudianteId
      ),

      apellidos: String(
        estudiante.apellidos
      ),

      nombres: String(
        estudiante.nombres
      ),

      curso: String(
        estudiante.curso
      )

    },

    evaluacion: {

      evaluacionId: String(
        evaluacion.evaluacionId
      ),

      titulo: String(
        evaluacion.titulo
      ),

      curso: String(
        evaluacion.curso
      ),

      nivel: String(
        evaluacion.nivel
      ),

      estado: String(
        evaluacion.estado
      ),

      intentosPermitidos: limite,

      intentosRealizados: realizados,

      intentosRestantes: limite === 0
        ? null
        : limite - realizados

    },

    preguntas: preguntas

  };

}


/* =========================================================
   8. CONTROL DE INTENTOS
========================================================= */

function contarIntentosEstudiante_(
  evaluacionId,
  estudianteId
) {

  const intentos = hojaAObjetos_(
    CONFIG.HOJAS.INTENTOS
  );

  return intentos.filter(intento => {

    return (

      String(intento.evaluacionId).trim() ===
      String(evaluacionId).trim() &&

      String(intento.estudianteId).trim() ===
      String(estudianteId).trim()

    );

  }).length;

}


/* =========================================================
   9. CORRECCIÓN Y CALIFICACIÓN
========================================================= */

function calcularNivel_(porcentaje) {

  const desarrollo = Number(
    obtenerConfiguracion_("porcentajeDesarrollo") || 60
  );

  const logrado = Number(
    obtenerConfiguracion_("porcentajeLogrado") || 85
  );

  if (
    !Number.isFinite(desarrollo) ||
    !Number.isFinite(logrado) ||
    desarrollo < 0 ||
    logrado > 100 ||
    desarrollo >= logrado
  ) {
    throw new Error(
      "Los porcentajes de configuración no son válidos."
    );
  }

  if (porcentaje >= logrado) {
    return CONFIG.NIVELES.LOGRADO;
  }

  if (porcentaje >= desarrollo) {
    return CONFIG.NIVELES.DESARROLLO;
  }

  return CONFIG.NIVELES.INICIAL;

}


function calcularNota_(porcentaje) {

  let nota;

  if (porcentaje <= 60) {

    nota = 1 + (porcentaje / 60) * 3;

  } else {

    nota = 4 + ((porcentaje - 60) / 40) * 3;

  }

  nota = Math.round(nota * 10) / 10;

  return Math.max(1, Math.min(7, nota));

}


/**
 * Recupera la respuesta correcta exclusivamente
 * desde la base de datos del servidor.
 */
function obtenerRespuestaCorrecta_(pregunta) {

  return normalizarTexto_(
    pregunta.respuestaCorrecta
  );

}


function evaluarRespuesta_(pregunta, respuestaEstudiante) {

  const correcta = obtenerRespuestaCorrecta_(
    pregunta
  );

  const respuesta = normalizarTexto_(
    respuestaEstudiante
  );

  if (!respuesta) {
    return false;
  }

  return respuesta === correcta;

}


/* =========================================================
   10. VALIDACIÓN DE ENVÍOS
========================================================= */

function prepararRespuestas_(submission, preguntas, opcionesDisponibles) {

  const recibidas = Array.isArray(submission.respuestas)
    ? submission.respuestas
    : [];

  const preguntasValidas = {};
  preguntas.forEach(pregunta => {
    preguntasValidas[String(pregunta.preguntaId)] = pregunta;
  });

  // Las opciones se cargan una sola vez desde la hoja, no una vez por pregunta.
  const opcionesPorPregunta = {};
  (opcionesDisponibles || []).forEach(opcion => {
    const id = String(opcion.preguntaId || '');
    if (!opcionesPorPregunta[id]) opcionesPorPregunta[id] = [];
    opcionesPorPregunta[id].push(opcion);
  });

  const mapa = {};
  recibidas.forEach(item => {
    if (!item || typeof item !== 'object') {
      throw new Error('Se recibió una respuesta inválida.');
    }

    const preguntaId = String(item.preguntaId || '').trim();
    if (!preguntaId || !preguntasValidas[preguntaId]) {
      throw new Error('Se recibió una pregunta que no pertenece a esta evaluación.');
    }
    if (Object.prototype.hasOwnProperty.call(mapa, preguntaId)) {
      throw new Error('Se recibió más de una respuesta para la misma pregunta.');
    }

    const respuesta = item.respuesta === null || item.respuesta === undefined
      ? '' : String(item.respuesta).trim();
    const pregunta = preguntasValidas[preguntaId];

    if (respuesta !== '') {
      const tipo = normalizarTexto_(pregunta.tipo);
      if (['seleccion', 'multiple', 'pareados', 'cloze', 'completar'].includes(tipo)) {
        const opciones = opcionesPorPregunta[preguntaId] || [];
        const codigosValidos = opciones.map(opcion => normalizarTexto_(opcion.codigo));
        if (codigosValidos.length > 0 && !codigosValidos.includes(normalizarTexto_(respuesta))) {
          throw new Error('La respuesta seleccionada no es válida.');
        }
      }
    }
    mapa[preguntaId] = respuesta;
  });
  return mapa;
}

/** Escribe varios objetos con una sola operación setValues. */
function appInsertarObjetosEnvioEnLote_(nombreHoja, objetos) {
  if (!objetos || objetos.length === 0) return;
  const hoja = obtenerHoja_(nombreHoja);
  const encabezados = obtenerEncabezados_(hoja);
  if (!encabezados.length) throw new Error('La hoja ' + nombreHoja + ' no tiene encabezados.');

  const faltantes = [...new Set(objetos.flatMap(obj => Object.keys(obj)))].filter(c => !encabezados.includes(c));
  if (faltantes.length) throw new Error('Faltan columnas en ' + nombreHoja + ': ' + faltantes.join(', '));

  const filas = objetos.map(objeto => encabezados.map(campo =>
    Object.prototype.hasOwnProperty.call(objeto, campo) ? objeto[campo] : ''
  ));
  const filaInicio = hoja.getLastRow() + 1;
  const filaFinal = filaInicio + filas.length - 1;
  if (filaFinal > hoja.getMaxRows()) {
    hoja.insertRowsAfter(hoja.getMaxRows(), filaFinal - hoja.getMaxRows());
  }
  hoja.getRange(filaInicio, 1, filas.length, encabezados.length).setValues(filas);
}

/* =========================================================
   11. ENVÍO Y REGISTRO DE EVALUACIÓN
========================================================= */

function enviarEvaluacion(submission) {
  if (!submission || typeof submission !== 'object') {
    return { correcto: false, mensaje: 'No se recibió información de la evaluación.' };
  }

  const evaluacionId = String(submission.evaluacionId || '').trim();
  const estudianteId = String(submission.estudianteId || '').trim();
  const apellidos = String(submission.apellidos || '').trim();
  const nombres = String(submission.nombres || '').trim();
  const curso = String(submission.curso || '').trim();

  if (!evaluacionId || !estudianteId || !apellidos || !nombres || !curso) {
    return { correcto: false, mensaje: 'Faltan datos obligatorios para enviar la evaluación.' };
  }

  try {
    /*
     * PREPARACIÓN FUERA DEL BLOQUEO:
     * lecturas de preguntas/opciones, validaciones y corrección.
     * Así los demás computadores no quedan esperando mientras se corrige.
     */
    const evaluacion = obtenerEvaluacionPorId_(evaluacionId);
    if (!evaluacion || !esPublicada_(evaluacion.estado)) {
      return { correcto: false, mensaje: 'La evaluación no está disponible.' };
    }
    if (!appCursosCompatibles_(evaluacion.curso, curso)) {
      return { correcto: false, mensaje: 'La evaluación no corresponde al curso.' };
    }

    const estudiante = buscarEstudiante_(apellidos, nombres, curso);
    if (!estudiante) {
      return { correcto: false, mensaje: 'El estudiante no pudo ser validado.' };
    }
    if (String(estudiante.estudianteId).trim() !== estudianteId) {
      return { correcto: false, mensaje: 'Los datos del estudiante no coinciden.' };
    }

    const preguntas = hojaAObjetos_(CONFIG.HOJAS.PREGUNTAS)
      .filter(p => String(p.evaluacionId) === evaluacionId && (!p.estado || esActivo_(p.estado)))
      .sort((a, b) => Number(a.orden || 0) - Number(b.orden || 0));

    if (!preguntas.length) {
      return { correcto: false, mensaje: 'La evaluación no tiene preguntas.' };
    }

    preguntas.forEach(pregunta => {
      const puntaje = Number(pregunta.puntaje);
      if (!Number.isFinite(puntaje) || puntaje < 0) {
        throw new Error('Existe un puntaje inválido en la pregunta ' + pregunta.preguntaId);
      }
      if (!normalizarTexto_(pregunta.respuestaCorrecta)) {
        throw new Error('Falta configurar la respuesta correcta de la pregunta ' + pregunta.preguntaId);
      }
    });

    const opciones = hojaAObjetos_(CONFIG.HOJAS.OPCIONES);
    const mapaRespuestas = prepararRespuestas_(submission, preguntas, opciones);

    let puntajeObtenido = 0;
    let puntajeMaximo = 0;
    const resultadosPreguntas = preguntas.map(pregunta => {
      const preguntaId = String(pregunta.preguntaId);
      const puntaje = Number(pregunta.puntaje);
      puntajeMaximo += puntaje;
      const respuestaEstudiante = mapaRespuestas[preguntaId] || '';
      const correcta = evaluarRespuesta_(pregunta, respuestaEstudiante);
      const puntajePregunta = correcta ? puntaje : 0;
      if (correcta) puntajeObtenido += puntaje;
      return {
        preguntaId: preguntaId,
        respuestaEstudiante: respuestaEstudiante,
        correcta: correcta,
        puntaje: puntajePregunta
      };
    });

    if (puntajeMaximo <= 0) {
      throw new Error('La evaluación no tiene un puntaje máximo válido.');
    }

    const porcentaje = Math.round((puntajeObtenido / puntajeMaximo) * 1000) / 10;
    const nivel = calcularNivel_(porcentaje);
    const nota = calcularNota_(porcentaje);
    const fechaEnvio = fechaActual_();
    const fechaInicio = fechaEnvio;
    const intentoId = generarId_('INT');
    const resultadoId = generarId_('RESU');

    /*
     * BLOQUEO CORTO:
     * solo vuelve a comprobar el límite de intentos y registra tres bloques.
     * El control de intentos permanece protegido contra envíos simultáneos.
     */
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(30000);
    } catch (errorBloqueo) {
      return {
        correcto: false,
        mensaje: 'El sistema está procesando otros envíos. Tus respuestas siguen en pantalla; espera unos segundos e inténtalo una sola vez más.'
      };
    }

    try {
      // Revalidar estado y límite dentro del bloqueo para no aceptar intentos concurrentes por encima del máximo.
      const evaluacionActual = obtenerEvaluacionPorId_(evaluacionId);
      if (!evaluacionActual || !esPublicada_(evaluacionActual.estado)) {
        return { correcto: false, mensaje: 'La evaluación ya no está disponible.' };
      }
      const limite = obtenerLimiteIntentos_(evaluacionActual);
      const realizados = contarIntentosEstudiante_(evaluacionId, estudianteId);
      if (limite > 0 && realizados >= limite) {
        return {
          correcto: false,
          mensaje: 'Has alcanzado el máximo de intentos permitidos.',
          intentosPermitidos: limite,
          intentosRealizados: realizados
        };
      }

      // Tres escrituras por lote, en lugar de una escritura por cada respuesta.
      appInsertarObjetosEnvioEnLote_(CONFIG.HOJAS.INTENTOS, [{
        intentoId: intentoId,
        evaluacionId: evaluacionId,
        estudianteId: estudianteId,
        apellidos: apellidos,
        nombres: nombres,
        curso: curso,
        fechaInicio: fechaInicio,
        fechaEnvio: fechaEnvio,
        puntajeObtenido: puntajeObtenido,
        puntajeMaximo: puntajeMaximo,
        porcentaje: porcentaje,
        nivel: nivel
      }]);

      const filasRespuestas = resultadosPreguntas.map(resultado => ({
        respuestaId: generarId_('RES'),
        intentoId: intentoId,
        preguntaId: resultado.preguntaId,
        respuestaEstudiante: resultado.respuestaEstudiante,
        correcta: resultado.correcta,
        puntaje: resultado.puntaje
      }));
      appInsertarObjetosEnvioEnLote_(CONFIG.HOJAS.RESPUESTAS, filasRespuestas);

      asegurarColumnas_(CONFIG.HOJAS.RESULTADOS, ['apellidos', 'nombres']);
      appInsertarObjetosEnvioEnLote_(CONFIG.HOJAS.RESULTADOS, [{
        resultadoId: resultadoId,
        intentoId: intentoId,
        evaluacionId: evaluacionId,
        estudianteId: estudianteId,
        apellidos: apellidos,
        nombres: nombres,
        curso: curso,
        puntaje: puntajeObtenido,
        puntajeMaximo: puntajeMaximo,
        porcentaje: porcentaje,
        nivel: nivel,
        fecha: fechaEnvio
      }]);

      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }

    return {
      correcto: true,
      mensaje: 'Evaluación enviada y corregida correctamente.',
      resultado: {
        intentoId: intentoId,
        puntaje: puntajeObtenido,
        puntajeMaximo: puntajeMaximo,
        porcentaje: porcentaje,
        nota: nota,
        nivel: nivel,
        totalPreguntas: preguntas.length,
        correctas: resultadosPreguntas.filter(item => item.correcta).length,
        incorrectas: resultadosPreguntas.filter(item => !item.correcta).length
      }
    };

  } catch (error) {
    console.error(error);
    return {
      correcto: false,
      mensaje: error.message || 'Ocurrió un error al procesar la evaluación.'
    };
  }
}

/* =========================================================
   12. PANEL DOCENTE
========================================================= */

/**
 * Configuración inicial: ejecutar UNA SOLA VEZ desde el editor de Apps Script.
 * La contraseña se solicita mediante un cuadro privado y se guarda como hash
 * en Propiedades del proyecto; nunca se escribe en el HTML ni en la hoja.
 */
function configurarClaveDocenteInicial() {
  const ui = SpreadsheetApp.getUi();
  const propiedades = PropertiesService.getScriptProperties();
  if (propiedades.getProperty("APP_DOCENTE_PASSWORD_HASH")) {
    ui.alert("La contraseña docente ya está configurada. Para cambiarla, solicita asistencia al administrador del sistema.");
    return;
  }
  const respuesta = ui.prompt(
    "Configurar acceso docente",
    "Ingresa una contraseña exclusiva para docentes y rectoría (mínimo 10 caracteres).",
    ui.ButtonSet.OK_CANCEL
  );
  if (respuesta.getSelectedButton() !== ui.Button.OK) return;
  const clave = String(respuesta.getResponseText() || "");
  if (clave.length < 10) {
    ui.alert("La contraseña debe tener al menos 10 caracteres. Vuelve a ejecutar la configuración.");
    return;
  }
  propiedades.setProperty("APP_DOCENTE_PASSWORD_HASH", appHashClave_(clave));
  ui.alert("Contraseña docente configurada correctamente. Ya puedes publicar la aplicación.");
}

function appHashClave_(clave) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(clave),
    Utilities.Charset.UTF_8
  );
  return bytes.map(b => ("0" + ((b + 256) % 256).toString(16)).slice(-2)).join("");
}

function appValidarSesionDocente_(token) {
  const valor = String(token || "").trim();
  if (!valor || valor.length < 30) return false;
  return CacheService.getScriptCache().get("APP_DOCENTE_SESION_" + valor) === "OK";
}

/** Inicio de sesión docente. La sesión dura 30 minutos y no se guarda en localStorage. */
function validarAccesoDocente(clave) {
  const propiedades = PropertiesService.getScriptProperties();
  const hashGuardado = propiedades.getProperty("APP_DOCENTE_PASSWORD_HASH");
  if (!hashGuardado) {
    return { correcto: false, mensaje: "El acceso docente aún no ha sido configurado por el administrador." };
  }

  const cache = CacheService.getScriptCache();
  const intentos = Number(cache.get("APP_DOCENTE_INTENTOS") || 0);
  if (intentos >= 5) {
    return { correcto: false, mensaje: "Se bloquearon temporalmente los intentos. Espera 10 minutos y vuelve a intentarlo." };
  }

  const recibido = appHashClave_(String(clave || ""));
  if (recibido !== hashGuardado) {
    cache.put("APP_DOCENTE_INTENTOS", String(intentos + 1), 600);
    return { correcto: false, mensaje: "Contraseña incorrecta." };
  }

  cache.remove("APP_DOCENTE_INTENTOS");
  const token = Utilities.getUuid() + Utilities.getUuid();
  cache.put("APP_DOCENTE_SESION_" + token, "OK", 1800);
  return { correcto: true, token: token, duracionMinutos: 30 };
}

/** Normaliza los nombres históricos y limita el panel docente a los cuatro cursos autorizados. */
function appCursoDocenteCanonico_(curso) {
  const texto = normalizarTexto_(curso).replace(/[^a-z0-9]/g, "");
  if (/^(?:3|tercer[op]?)(?:basico|basica)(?:a)?$/.test(texto)) return "3° Básico";
  if (/^(?:4|cuart[oa]?)(?:basico|basica)(?:a)?$/.test(texto)) return "4° Básico";
  if (/^(?:5|quint[oa]?)(?:basico|basica)(?:a)?$/.test(texto)) return "5° Básico";
  if (/^(?:6|sext[oa]?)(?:basico|basica)(?:a)?$/.test(texto)) return "6° Básico";
  return "";
}

/** Devuelve resultados únicamente cuando se presenta una sesión docente válida. */
function obtenerResultadosDocente(token) {
  if (!appValidarSesionDocente_(token)) {
    throw new Error("Tu sesión docente expiró. Ingresa nuevamente la contraseña.");
  }

  const resultados = hojaAObjetos_(CONFIG.HOJAS.RESULTADOS);
  const intentos = hojaAObjetos_(CONFIG.HOJAS.INTENTOS);
  const estudiantes = hojaAObjetos_(CONFIG.HOJAS.ESTUDIANTES);
  const evaluaciones = hojaAObjetos_(CONFIG.HOJAS.EVALUACIONES);

  const mapaIntentos = {};
  intentos.forEach(item => {
    mapaIntentos[String(item.intentoId || "")] = item;
  });

  const mapaEstudiantes = {};
  estudiantes.forEach(item => {
    mapaEstudiantes[String(item.estudianteId || "")] = item;
  });

  const mapaEvaluaciones = {};
  evaluaciones.forEach(item => {
    mapaEvaluaciones[String(item.evaluacionId || "")] = item;
  });

  return resultados.slice().reverse().map(resultado => {
    const intento = mapaIntentos[String(resultado.intentoId || "")] || {};
    const estudiante = mapaEstudiantes[String(resultado.estudianteId || "")] || {};

    const apellidos = String(
      resultado.apellidos ||
      intento.apellidos ||
      estudiante.apellidos ||
      ""
    ).trim();

    const nombres = String(
      resultado.nombres ||
      intento.nombres ||
      estudiante.nombres ||
      ""
    ).trim();

    const evaluacion =
      mapaEvaluaciones[String(resultado.evaluacionId || "")] || {};

    const cursoCanonico = appCursoDocenteCanonico_(
      resultado.curso ||
      intento.curso ||
      estudiante.curso ||
      ""
    );

    if (!cursoCanonico) return null;

    const porcentaje = Number(resultado.porcentaje || 0);

    // La nota se calcula con la escala chilena configurada en EvalúaPro.
    // Se obtiene del porcentaje guardado, por lo que también funciona
    // con resultados históricos que no tienen una columna "nota".
    const nota = calcularNota_(porcentaje);

    const fecha = resultado.fecha;

    return {
      resultadoId: String(resultado.resultadoId || ""),
      intentoId: String(resultado.intentoId || ""),
      evaluacionId: String(resultado.evaluacionId || ""),
      evaluacion: String(
        evaluacion.titulo ||
        resultado.evaluacionId ||
        "Evaluación"
      ),
      estudianteId: String(resultado.estudianteId || ""),
      apellidos: apellidos || "No identificado",
      nombres: nombres || "",
      estudiante: (apellidos + " " + nombres).trim() ||
        "No identificado (" +
        String(resultado.estudianteId || "sin ID") +
        ")",
      curso: cursoCanonico,
      puntaje: Number(resultado.puntaje || 0),
      puntajeMaximo: Number(resultado.puntajeMaximo || 0),
      porcentaje: porcentaje,
      nota: nota,
      nivel: String(resultado.nivel || ""),
      fecha: fecha instanceof Date
        ? Utilities.formatDate(
            fecha,
            Session.getScriptTimeZone(),
            "dd/MM/yyyy HH:mm"
          )
        : String(fecha || "")
    };
  }).filter(Boolean).sort(
    (a, b) => b.resultadoId.localeCompare(a.resultadoId)
  );
}
function cerrarSesionDocente(token) {
  const valor = String(token || "").trim();
  if (valor) CacheService.getScriptCache().remove("APP_DOCENTE_SESION_" + valor);
  return { correcto: true };
}



function obtenerAnalisisPreguntas(token) {

  if (!appValidarSesionDocente_(token)) {
    throw new Error("Acceso docente no autorizado.");
  }

  const respuestas = hojaAObjetos_(
    CONFIG.HOJAS.RESPUESTAS
  );

  const preguntas = hojaAObjetos_(
    CONFIG.HOJAS.PREGUNTAS
  );

  const mapaPreguntas = {};

  preguntas.forEach(pregunta => {

    mapaPreguntas[
      String(pregunta.preguntaId)
    ] = pregunta;

  });

  const resumen = {};

  respuestas.forEach(respuesta => {

    const preguntaId = String(
      respuesta.preguntaId
    );

    if (!resumen[preguntaId]) {

      resumen[preguntaId] = {

        preguntaId: preguntaId,

        enunciado: mapaPreguntas[preguntaId]
          ? String(
              mapaPreguntas[preguntaId].enunciado || ""
            )
          : "",

        total: 0,

        correctas: 0,

        incorrectas: 0,

        omitidas: 0

      };

    }

    resumen[preguntaId].total++;

    if (
      normalizarTexto_(
        respuesta.respuestaEstudiante
      ) === ""
    ) {

      resumen[preguntaId].omitidas++;

    } else if (
      respuesta.correcta === true ||
      normalizarTexto_(respuesta.correcta) === "true"
    ) {

      resumen[preguntaId].correctas++;

    } else {

      resumen[preguntaId].incorrectas++;

    }

  });

  return Object.values(resumen);

}

/* =========================================================
   APLICACIÓN WEB
   Los bancos de preguntas, claves de respuesta, datos de estudiantes,
   herramientas internas de diagnóstico y rutinas de mantenimiento del
   sistema institucional se excluyeron de esta copia de portafolio.
========================================================= */

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile("Index")
    .setTitle("Sistema de Evaluación Online")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
