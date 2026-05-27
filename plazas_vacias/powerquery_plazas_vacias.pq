// ════════════════════════════════════════════════════════════════
// POWER QUERY — PLAZAS VACÍAS DOCENTES (Antioquia)
// ════════════════════════════════════════════════════════════════
// 
// Este código se pega en el Editor Avanzado de Power Query en Excel.
// Conecta al archivo fuente descargado del MEN y filtra las plazas
// vacías de las 9 subregiones oficiales de Antioquia.
//
// INSTRUCCIONES:
//   1. En Excel, ir a Datos → Obtener datos → Desde otras fuentes
//      → Consulta en blanco
//   2. Clic en "Editor avanzado"
//   3. Pegar TODO este código (reemplazando el contenido existente)
//   4. Cambiar la variable RutaArchivo por la ruta real del archivo
//      en OneDrive/SharePoint
//   5. Clic en "Listo" y luego "Cerrar y cargar"
//
// ════════════════════════════════════════════════════════════════

let
    // ══════════════════════════════════════════════════════════
    // CONFIGURACIÓN — Cambiar esta ruta por la del archivo fuente
    // ══════════════════════════════════════════════════════════
    RutaArchivo = "C:\Users\TU_USUARIO\OneDrive - EMPRESA\Planta Docente\22-4-26.xlsx",

    // ══════════════════════════════════════════════════════════
    // PASO 1: Leer la hoja "T" del archivo fuente
    // ══════════════════════════════════════════════════════════
    Origen = Excel.Workbook(File.Contents(RutaArchivo), null, true),
    HojaT = Origen{[Item="T",Kind="Sheet"]}[Data],
    Encabezados = Table.PromoteHeaders(HojaT, [PromoteAllScalars=true]),

    // ══════════════════════════════════════════════════════════
    // PASO 2: Filtrar por las 9 subregiones oficiales de Antioquia
    // ══════════════════════════════════════════════════════════
    SubregionesOficiales = {"BAJO CAUCA", "MAGDALENA MEDIO", "NORDESTE", "NORTE", "OCCIDENTE", "ORIENTE", "SUROESTE", "URABA", "VALLE DE ABURRA"},
    FiltroSubregion = Table.SelectRows(Encabezados, each List.Contains(SubregionesOficiales, [SUBREGION])),

    // ══════════════════════════════════════════════════════════
    // PASO 3: Clasificar tipo de vacante
    // ══════════════════════════════════════════════════════════
    
    // Estados de Vacante Definitiva
    EstadosVacanteDefinitiva = {"Encargo Vacante Definitiva"},
    
    // Estados de Vacante Temporal (por situación laboral)
    EstadosVacanteTemporal = {"Encargo Vacante Temporal", "Reemplazo Vacante Temporal"},
    
    // Estados de Titular Ausente (su plaza necesita cobertura temporal)
    EstadosTitularAusente = {"Comision", "Comisión No Remunerada", "Comision Libre Nom. y Remocion", "Licencia No Remunerada", "Permiso Sindical", "Permiso Remunerado", "Periodo Prueba en otra entidad", "Periodo Prueba misma entidad", "Sancion", "Accion disciplinaria- sus", "Incapacidad Sup 180 Días", "Recuperacion Capacidad Laboral", "Situacion Laboral Remunerada"},

    // Agregar columna TIPO_VACANTE
    ConTipo = Table.AddColumn(FiltroSubregion, "TIPO_VACANTE", each
        if List.Contains(EstadosVacanteDefinitiva, [SITUACIONLABORAL]) then "Vacante Definitiva"
        else if List.Contains(EstadosVacanteTemporal, [SITUACIONLABORAL]) then "Vacante Temporal"
        else if [NOVEDAD_PLANTA] = "Vacancia Temporal" then "Vacante Temporal"
        else if List.Contains(EstadosTitularAusente, [SITUACIONLABORAL]) then "Vacante Temporal"
        else if [RETIRO_FECHA] <> null and [RETIRO_FECHA] < DateTime.LocalNow() 
                and not Text.Contains(Text.From([SITUACIONLABORAL]), "Reemplazo")
                and [NUM_PLAZA] <> 0 then "Vacante Definitiva"
        else null
    , type text),

    // ══════════════════════════════════════════════════════════
    // PASO 4: Filtrar solo plazas vacías (quitar nulls)
    // ══════════════════════════════════════════════════════════
    SoloVacantes = Table.SelectRows(ConTipo, each [TIPO_VACANTE] <> null),

    // ══════════════════════════════════════════════════════════
    // PASO 5: Seleccionar columnas del reporte
    // ══════════════════════════════════════════════════════════
    ColumnasFinales = Table.SelectColumns(SoloVacantes, {
        "TIPO_VACANTE",
        "NUM_PLAZA",
        "MUNICIPIO",
        "SUBREGION",
        "I.E",
        "SEDE",
        "AREA",
        "CARGO",
        "GRADO",
        "CODSEDE_AREA",
        "DANE_ESTABLECIMIENTO",
        "DANE_SEDE",
        "SITUACIONLABORAL",
        "NOVEDAD_PLANTA",
        "NIVELDICTA",
        "AREAEDUCATIVA",
        "NIVELCONTRATACION",
        "ESCALAFON",
        "CODEMPLEADO",
        "EMPLEADO",
        "RETIRO_FECHA",
        "ENCARGO",
        "ENCARGO_GRADO",
        "REEMPLAZO_EMPLEADO",
        "IE_SEDE_AREA"
    }),

    // ══════════════════════════════════════════════════════════
    // PASO 6: Ordenar por subregión y municipio
    // ══════════════════════════════════════════════════════════
    Ordenado = Table.Sort(ColumnasFinales, {
        {"SUBREGION", Order.Ascending},
        {"MUNICIPIO", Order.Ascending},
        {"TIPO_VACANTE", Order.Ascending}
    })
in
    Ordenado
