Attribute VB_Name = "ModuloVacantes"
' ================================================================
' MACRO: Generar Reporte de Plazas Vacantes Disponibles
' ================================================================
' Este modulo se agrega al archivo "Registro de vacantes DEFINITIVAS"
' en SharePoint. Al ejecutar la macro (mediante un boton), se genera
' un NUEVO archivo Excel con solo las plazas que aun estan vacantes.
'
' Criterio: una plaza esta vacante si el "Estado del nombramiento"
' NO es "FIRMADO" (es decir: "Seleccionar", "EN PROCESO",
' "EN FIRMAS", o vacio).
'
' Solo se incluyen las 9 subregiones oficiales de Antioquia.
' ================================================================

Option Explicit

' ----------------------------------------------------------------
' Subregiones oficiales de Antioquia (9)
' ----------------------------------------------------------------
Private Function EsSubregionOficial(ByVal valor As String) As Boolean
    Dim sub9 As Variant
    Dim s As Variant
    Dim valorUpper As String
    
    valorUpper = UCase(Trim(valor))
    
    sub9 = Array("BAJO CAUCA", "MAGDALENA MEDIO", "NORDESTE", "NORTE", _
                 "OCCIDENTE", "ORIENTE", "SUROESTE", "URABA", _
                 "VALLE DE ABURRA")
    
    EsSubregionOficial = False
    For Each s In sub9
        If valorUpper = CStr(s) Then
            EsSubregionOficial = True
            Exit Function
        End If
    Next s
    
    ' Manejar variaciones con tildes
    If InStr(1, valorUpper, "URABA", vbTextCompare) > 0 Then EsSubregionOficial = True
    If InStr(1, valorUpper, "ABURRA", vbTextCompare) > 0 Then EsSubregionOficial = True
End Function

' ----------------------------------------------------------------
' Buscar columna por nombre de encabezado (fila 1)
' ----------------------------------------------------------------
Private Function BuscarColumna(ws As Worksheet, ByVal nombreCol As String) As Long
    Dim col As Long
    Dim ultimaCol As Long
    
    ultimaCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    For col = 1 To ultimaCol
        If LCase(Trim(CStr(ws.Cells(1, col).Value))) = LCase(Trim(nombreCol)) Then
            BuscarColumna = col
            Exit Function
        End If
    Next col
    
    BuscarColumna = 0
End Function

' ----------------------------------------------------------------
' Leer valor de celda como texto limpio
' ----------------------------------------------------------------
Private Function LeerCelda(ws As Worksheet, fila As Long, col As Long) As String
    If col = 0 Then
        LeerCelda = ""
        Exit Function
    End If
    
    Dim v As Variant
    v = ws.Cells(fila, col).Value
    
    If IsEmpty(v) Or IsNull(v) Then
        LeerCelda = ""
    Else
        LeerCelda = Trim(CStr(v))
    End If
End Function

' ================================================================
' MACRO PRINCIPAL — Ejecutar con el boton
' ================================================================
Public Sub GenerarReportePlazasVacantes()
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' ----------------------------------------------------------
    ' 1. Identificar la hoja de datos
    ' ----------------------------------------------------------
    Dim wsOrigen As Worksheet
    Set wsOrigen = ThisWorkbook.Sheets(1)  ' Primera hoja del libro
    
    Dim ultimaFila As Long
    Dim ultimaCol As Long
    ultimaFila = wsOrigen.Cells(wsOrigen.Rows.Count, 1).End(xlUp).Row
    ultimaCol = wsOrigen.Cells(1, wsOrigen.Columns.Count).End(xlToLeft).Column
    
    If ultimaFila < 2 Then
        MsgBox "No se encontraron datos en la hoja.", vbExclamation, "Sin datos"
        GoTo Limpiar
    End If
    
    ' ----------------------------------------------------------
    ' 2. Localizar columnas clave por nombre de encabezado
    ' ----------------------------------------------------------
    Dim colPlaza As Long
    Dim colSubregion As Long
    Dim colMunicipio As Long
    Dim colEstablecimiento As Long
    Dim colSede As Long
    Dim colZona As Long
    Dim colCargo As Long
    Dim colTipoPlaza As Long
    Dim colNivelAcad As Long
    Dim colAreaEduc As Long
    Dim colMotivo As Long
    Dim colFechaRegistro As Long
    Dim colCedulaGenera As Long
    Dim colNombreGenera As Long
    Dim colActoAdmin As Long
    Dim colFechaActo As Long
    Dim colObservacion As Long
    Dim colElegibles As Long
    Dim colObsPermanencia As Long
    Dim colTomadaPor As Long
    Dim colTomadaPara As Long
    Dim colEstado As Long
    Dim colCedulaSel As Long
    Dim colNombreSel As Long
    Dim colRegistradaPor As Long
    
    colPlaza = BuscarColumna(wsOrigen, "PLAZA")
    colSubregion = BuscarColumna(wsOrigen, "Subregi" & ChrW(243) & "n")
    If colSubregion = 0 Then colSubregion = BuscarColumna(wsOrigen, "Subregion")
    If colSubregion = 0 Then colSubregion = BuscarColumna(wsOrigen, "Subregi" & Chr(243) & "n")
    colMunicipio = BuscarColumna(wsOrigen, "Municipio")
    colEstablecimiento = BuscarColumna(wsOrigen, "Establecimiento")
    colSede = BuscarColumna(wsOrigen, "Sede")
    colZona = BuscarColumna(wsOrigen, "Zona")
    colCargo = BuscarColumna(wsOrigen, "Cargo de la vacante")
    colTipoPlaza = BuscarColumna(wsOrigen, "Tipo de plaza")
    colNivelAcad = BuscarColumna(wsOrigen, "Nivel acad" & ChrW(233) & "mico")
    If colNivelAcad = 0 Then colNivelAcad = BuscarColumna(wsOrigen, "Nivel academico")
    colMotivo = BuscarColumna(wsOrigen, "Motivo de la vacante")
    colFechaRegistro = BuscarColumna(wsOrigen, "Fecha de registro de la vacante definitiva")
    colCedulaGenera = BuscarColumna(wsOrigen, "C" & ChrW(233) & "dula de qui" & ChrW(233) & "n genera la vacante")
    If colCedulaGenera = 0 Then colCedulaGenera = BuscarColumna(wsOrigen, "Cedula de quien genera la vacante")
    colNombreGenera = BuscarColumna(wsOrigen, "Nombre de qui" & ChrW(233) & "n genera la vacante")
    If colNombreGenera = 0 Then colNombreGenera = BuscarColumna(wsOrigen, "Nombre de quien genera la vacante")
    colActoAdmin = BuscarColumna(wsOrigen, "Acto administrativo de la vacante")
    colFechaActo = BuscarColumna(wsOrigen, "Fecha del acto administrativo")
    colObservacion = BuscarColumna(wsOrigen, "Observaci" & ChrW(243) & "n")
    If colObservacion = 0 Then colObservacion = BuscarColumna(wsOrigen, "Observacion")
    colElegibles = BuscarColumna(wsOrigen, "?Tiene lista de Elegibles?")
    If colElegibles = 0 Then colElegibles = BuscarColumna(wsOrigen, ChrW(191) & "Tiene lista de Elegibles?")
    colObsPermanencia = BuscarColumna(wsOrigen, "Observaci" & ChrW(243) & "n Direcci" & ChrW(243) & "n de Permanencia")
    If colObsPermanencia = 0 Then colObsPermanencia = BuscarColumna(wsOrigen, "Observacion Direccion de Permanencia")
    colTomadaPor = BuscarColumna(wsOrigen, "Vacante tomada por")
    colTomadaPara = BuscarColumna(wsOrigen, "Vacante tomada para")
    colEstado = BuscarColumna(wsOrigen, "Estado del nombramiento")
    colCedulaSel = BuscarColumna(wsOrigen, "C" & ChrW(233) & "dula")
    If colCedulaSel = 0 Then colCedulaSel = BuscarColumna(wsOrigen, "Cedula")
    colNombreSel = BuscarColumna(wsOrigen, "Nombre del seleccionado")
    colRegistradaPor = BuscarColumna(wsOrigen, "Vacante temporal registrada por")
    
    ' Columna de area educativa (columna 16, puede tener nombre variable)
    colAreaEduc = 16  ' Posicion fija como respaldo
    
    ' Validar columnas criticas
    If colPlaza = 0 Then
        MsgBox "No se encontro la columna 'PLAZA' en los encabezados." & vbCrLf & _
               "Verifique que el archivo tenga la estructura correcta.", _
               vbCritical, "Error de estructura"
        GoTo Limpiar
    End If
    
    If colEstado = 0 Then
        MsgBox "No se encontro la columna 'Estado del nombramiento' en los encabezados." & vbCrLf & _
               "Verifique que el archivo tenga la estructura correcta.", _
               vbCritical, "Error de estructura"
        GoTo Limpiar
    End If
    
    ' ----------------------------------------------------------
    ' 3. Filtrar filas: plazas aun vacantes
    ' ----------------------------------------------------------
    Dim filasVacantes() As Long
    ReDim filasVacantes(1 To ultimaFila)
    Dim contVacantes As Long
    contVacantes = 0
    
    Dim fila As Long
    Dim valorPlaza As String
    Dim valorEstado As String
    Dim valorSubregion As String
    
    For fila = 2 To ultimaFila
        ' Saltar filas sin numero de plaza
        valorPlaza = LeerCelda(wsOrigen, fila, colPlaza)
        If valorPlaza = "" Or valorPlaza = "0" Then GoTo SiguienteFila
        
        ' Filtrar por estado: excluir las que ya tienen nombramiento FIRMADO
        valorEstado = UCase(Trim(LeerCelda(wsOrigen, fila, colEstado)))
        If valorEstado = "FIRMADO" Then GoTo SiguienteFila
        
        ' Filtrar por subregion oficial (si la columna existe)
        If colSubregion > 0 Then
            valorSubregion = LeerCelda(wsOrigen, fila, colSubregion)
            If valorSubregion <> "" Then
                If Not EsSubregionOficial(valorSubregion) Then GoTo SiguienteFila
            End If
        End If
        
        ' Esta fila es una plaza vacante
        contVacantes = contVacantes + 1
        filasVacantes(contVacantes) = fila
        
SiguienteFila:
    Next fila
    
    If contVacantes = 0 Then
        MsgBox "No se encontraron plazas vacantes con los criterios establecidos.", _
               vbInformation, "Sin resultados"
        GoTo Limpiar
    End If
    
    ' ----------------------------------------------------------
    ' 4. Crear nuevo libro de Excel
    ' ----------------------------------------------------------
    Dim wbNuevo As Workbook
    Dim wsNuevo As Worksheet
    Set wbNuevo = Workbooks.Add(xlWBATWorksheet)
    Set wsNuevo = wbNuevo.Sheets(1)
    wsNuevo.Name = "Plazas Vacantes"
    
    ' ----------------------------------------------------------
    ' 5. Escribir encabezados en el nuevo archivo
    ' ----------------------------------------------------------
    Dim encabezados As Variant
    encabezados = Array( _
        "PLAZA", _
        "Subregi" & ChrW(243) & "n", _
        "Municipio", _
        "Establecimiento", _
        "Sede", _
        "Zona", _
        "Cargo de la vacante", _
        "Tipo de plaza", _
        "Nivel acad" & ChrW(233) & "mico", _
        ChrW(193) & "rea educativa", _
        "Motivo de la vacante", _
        "Fecha de registro", _
        "C" & ChrW(233) & "dula de qui" & ChrW(233) & "n genera la vacante", _
        "Nombre de qui" & ChrW(233) & "n genera la vacante", _
        "Acto administrativo de la vacante", _
        "Fecha del acto administrativo", _
        "Observaci" & ChrW(243) & "n", _
        ChrW(191) & "Tiene lista de Elegibles?", _
        "Obs. Direcci" & ChrW(243) & "n de Permanencia", _
        "Vacante tomada por", _
        "Vacante tomada para", _
        "Estado del nombramiento", _
        "C" & ChrW(233) & "dula seleccionado", _
        "Nombre del seleccionado", _
        "Registrada por" _
    )
    
    ' Columnas origen correspondientes
    Dim colsOrigen As Variant
    colsOrigen = Array( _
        colPlaza, _
        colSubregion, _
        colMunicipio, _
        colEstablecimiento, _
        colSede, _
        colZona, _
        colCargo, _
        colTipoPlaza, _
        colNivelAcad, _
        colAreaEduc, _
        colMotivo, _
        colFechaRegistro, _
        colCedulaGenera, _
        colNombreGenera, _
        colActoAdmin, _
        colFechaActo, _
        colObservacion, _
        colElegibles, _
        colObsPermanencia, _
        colTomadaPor, _
        colTomadaPara, _
        colEstado, _
        colCedulaSel, _
        colNombreSel, _
        colRegistradaPor _
    )
    
    Dim numCols As Long
    numCols = UBound(encabezados) - LBound(encabezados) + 1
    
    Dim col As Long
    For col = 0 To numCols - 1
        wsNuevo.Cells(1, col + 1).Value = encabezados(col)
    Next col
    
    ' Formato de encabezados
    With wsNuevo.Range(wsNuevo.Cells(1, 1), wsNuevo.Cells(1, numCols))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(68, 114, 196)  ' Azul
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .RowHeight = 30
    End With
    
    ' ----------------------------------------------------------
    ' 6. Copiar datos filtrados al nuevo archivo
    ' ----------------------------------------------------------
    Dim filaDestino As Long
    filaDestino = 2
    
    Dim i As Long
    Dim colIdx As Long
    
    For i = 1 To contVacantes
        fila = filasVacantes(i)
        For col = 0 To numCols - 1
            colIdx = CLng(colsOrigen(col))
            If colIdx > 0 Then
                wsNuevo.Cells(filaDestino, col + 1).Value = wsOrigen.Cells(fila, colIdx).Value
            End If
        Next col
        filaDestino = filaDestino + 1
    Next i
    
    ' ----------------------------------------------------------
    ' 7. Formato del nuevo archivo
    ' ----------------------------------------------------------
    ' Ancho de columnas
    wsNuevo.Columns("A:Y").AutoFit
    
    ' Limitar ancho maximo
    Dim c As Long
    For c = 1 To numCols
        If wsNuevo.Columns(c).ColumnWidth > 40 Then
            wsNuevo.Columns(c).ColumnWidth = 40
        End If
    Next c
    
    ' Filtros automaticos
    wsNuevo.Range(wsNuevo.Cells(1, 1), wsNuevo.Cells(filaDestino - 1, numCols)).AutoFilter
    
    ' Inmovilizar primera fila
    wsNuevo.Range("A2").Select
    ActiveWindow.FreezePanes = True
    
    ' Bordes
    With wsNuevo.Range(wsNuevo.Cells(1, 1), wsNuevo.Cells(filaDestino - 1, numCols))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlInsideVertical).Weight = xlThin
        .Borders(xlInsideHorizontal).Weight = xlThin
    End With
    
    ' ----------------------------------------------------------
    ' 8. Agregar hoja de resumen
    ' ----------------------------------------------------------
    Dim wsResumen As Worksheet
    Set wsResumen = wbNuevo.Sheets.Add(After:=wsNuevo)
    wsResumen.Name = "Resumen"
    
    ' Titulo
    wsResumen.Cells(1, 1).Value = "REPORTE DE PLAZAS VACANTES DISPONIBLES"
    wsResumen.Cells(1, 1).Font.Bold = True
    wsResumen.Cells(1, 1).Font.Size = 14
    wsResumen.Cells(1, 1).Font.Color = RGB(68, 114, 196)
    
    wsResumen.Cells(2, 1).Value = "Fecha de generaci" & ChrW(243) & "n: " & Format(Now, "dd/mm/yyyy hh:mm")
    wsResumen.Cells(3, 1).Value = "Archivo fuente: " & ThisWorkbook.Name
    
    wsResumen.Cells(5, 1).Value = "M" & ChrW(233) & "tricas"
    wsResumen.Cells(5, 1).Font.Bold = True
    wsResumen.Cells(5, 1).Font.Size = 12
    
    wsResumen.Cells(6, 1).Value = "Total registros en archivo fuente:"
    wsResumen.Cells(6, 2).Value = ultimaFila - 1
    
    wsResumen.Cells(7, 1).Value = "Plazas a" & ChrW(250) & "n vacantes (reporte):"
    wsResumen.Cells(7, 2).Value = contVacantes
    wsResumen.Cells(7, 2).Font.Bold = True
    wsResumen.Cells(7, 2).Font.Size = 14
    wsResumen.Cells(7, 2).Font.Color = RGB(192, 0, 0)
    
    ' Contar por subregion
    wsResumen.Cells(9, 1).Value = "Por Subregi" & ChrW(243) & "n"
    wsResumen.Cells(9, 1).Font.Bold = True
    wsResumen.Cells(9, 1).Font.Size = 12
    
    wsResumen.Cells(10, 1).Value = "Subregi" & ChrW(243) & "n"
    wsResumen.Cells(10, 2).Value = "Cantidad"
    wsResumen.Range("A10:B10").Font.Bold = True
    wsResumen.Range("A10:B10").Interior.Color = RGB(68, 114, 196)
    wsResumen.Range("A10:B10").Font.Color = RGB(255, 255, 255)
    
    ' Contar vacantes por subregion desde la hoja de datos
    Dim subregiones As Variant
    subregiones = Array("Bajo Cauca", "Magdalena Medio", "Nordeste", "Norte", _
                        "Occidente", "Oriente", "Suroeste", "Urab" & ChrW(225), _
                        "Valle de Aburr" & ChrW(225))
    
    Dim filaRes As Long
    filaRes = 11
    Dim contSub As Long
    Dim subNombre As Variant
    
    For Each subNombre In subregiones
        contSub = 0
        For i = 1 To contVacantes
            fila = filasVacantes(i)
            If colSubregion > 0 Then
                If UCase(Trim(CStr(wsOrigen.Cells(fila, colSubregion).Value))) = UCase(CStr(subNombre)) Then
                    contSub = contSub + 1
                End If
            End If
        Next i
        wsResumen.Cells(filaRes, 1).Value = subNombre
        wsResumen.Cells(filaRes, 2).Value = contSub
        filaRes = filaRes + 1
    Next subNombre
    
    ' Criterio utilizado
    filaRes = filaRes + 1
    wsResumen.Cells(filaRes, 1).Value = "Criterio utilizado"
    wsResumen.Cells(filaRes, 1).Font.Bold = True
    wsResumen.Cells(filaRes, 1).Font.Size = 12
    filaRes = filaRes + 1
    wsResumen.Cells(filaRes, 1).Value = "Se incluyen plazas donde 'Estado del nombramiento' NO es 'FIRMADO'."
    filaRes = filaRes + 1
    wsResumen.Cells(filaRes, 1).Value = "Se excluyen filas sin n" & ChrW(250) & "mero de PLAZA."
    filaRes = filaRes + 1
    wsResumen.Cells(filaRes, 1).Value = "Solo se incluyen las 9 subregiones oficiales de Antioquia."
    
    wsResumen.Columns("A:B").AutoFit
    If wsResumen.Columns(1).ColumnWidth < 45 Then wsResumen.Columns(1).ColumnWidth = 45
    
    ' ----------------------------------------------------------
    ' 9. Mostrar mensaje de exito
    ' ----------------------------------------------------------
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "Reporte generado exitosamente." & vbCrLf & vbCrLf & _
           "Plazas vacantes encontradas: " & contVacantes & vbCrLf & vbCrLf & _
           "Se ha creado un nuevo libro de Excel con los resultados." & vbCrLf & _
           "Recuerde guardarlo con el nombre y ubicaci" & ChrW(243) & "n deseados.", _
           vbInformation, ChrW(201) & "xito"
    
    ' Activar el nuevo libro
    wbNuevo.Activate
    wsNuevo.Select
    wsNuevo.Range("A1").Select
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Ocurri" & ChrW(243) & " un error al generar el reporte:" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Por favor verifique que el archivo tiene la estructura correcta.", _
           vbCritical, "Error"
    Exit Sub
    
Limpiar:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
