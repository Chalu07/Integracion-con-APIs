// ════════════════════════════════════════════════════════════════
// POWER QUERY — RESUMEN POR SUBREGIÓN
// ════════════════════════════════════════════════════════════════
//
// Esta consulta depende de la consulta principal "PlazasVacias".
// Crear DESPUÉS de cargar la consulta principal.
//
// INSTRUCCIONES:
//   1. Asegurarse de que la consulta "PlazasVacias" ya esté creada
//   2. Datos → Obtener datos → Desde otras fuentes → Consulta en blanco
//   3. Editor avanzado → Pegar este código
//   4. Clic en "Listo" y luego "Cerrar y cargar en..."
//      → seleccionar "Solo crear conexión" o cargar a una hoja
//
// ════════════════════════════════════════════════════════════════

let
    Fuente = PlazasVacias,
    Agrupado = Table.Group(Fuente, {"SUBREGION", "TIPO_VACANTE"}, {
        {"Cantidad", each Table.RowCount(_), Int64.Type}
    }),
    Pivotado = Table.Pivot(Agrupado, List.Distinct(Agrupado[TIPO_VACANTE]), "TIPO_VACANTE", "Cantidad", List.Sum),
    ConTotal = Table.AddColumn(Pivotado, "TOTAL", each
        (try [Vacante Definitiva] otherwise 0) + (try [Vacante Temporal] otherwise 0)
    , Int64.Type),
    Ordenado = Table.Sort(ConTotal, {{"TOTAL", Order.Descending}})
in
    Ordenado
