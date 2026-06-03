Attribute VB_Name = "ModuloReporteExterno"
' ================================================================
' MACRO: Reporte de Plazas Vacantes (Archivo Externo)
' ================================================================
' Este modulo va en un archivo SEPARADO del archivo de Vacantes
' Definitivas. Al ejecutar la macro (mediante un boton), se conecta
' al archivo fuente en SharePoint (solo lectura, sin modificarlo),
' lee los datos, filtra las plazas vacantes y muestra los resultados
' en este mismo archivo.
'
' CONFIGURACION: La ruta del archivo fuente se configura en la
' hoja "Configuracion" celda B2.
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
    If InStr(1, valorUpper, "URAB", vbTextCompare) > 0 Then EsSubregionOficial = True
    If InStr(1, valorUpper, "ABURRA", vbTextCompare) > 0 Then EsSubregionOficial = True
    If InStr(1, valorUpper, "ABURR", vbTextCompare) > 0 Then EsSubregionOficial = True
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

' ----------------------------------------------------------------
' Obtener la ruta del archivo fuente desde hoja Configuracion
' ----------------------------------------------------------------
Private Function ObtenerRutaFuente() As String
    Dim wsConfig As Worksheet
    
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets("Configuracion")
    On Error GoTo 0
    
    If wsConfig Is Nothing Then
        ObtenerRutaFuente = ""
        Exit Function
    End If
    
    ObtenerRutaFuente = Trim(CStr(wsConfig.Range("B2").Value))
End Function

' ================================================================
' MACRO PRINCIPAL — Ejecutar con el boton
' ================================================================
Public Sub GenerarReportePlazasVacantes()
    
    On Error GoTo ErrorHandler
    
    ' ----------------------------------------------------------
    ' 0. Obtener ruta del archivo fuente
    ' ----------------------------------------------------------
    Dim rutaFuente As String
    rutaFuente = ObtenerRutaFuente()
    
    If rutaFuente = "" Then
        MsgBox "No se ha configurado la ruta del archivo fuente." & vbCrLf & vbCrLf & _
               "Por favor, vaya a la hoja 'Configuracion' y escriba la ruta " & _
               "del archivo 'Vacantes Definitivas' en la celda B2.", _
               vbExclamation, "Ruta no configurada"
        Exit Sub
    End If
    
    ' Verificar que el archivo existe
    If Dir(rutaFuente) = "" Then
        ' Intentar si es ruta de SharePoint/OneDrive
        ' Las rutas de SharePoint a veces no funcionan con Dir()
        ' Intentaremos abrirlo directamente
    End If
    
    ' ----------------------------------------------------------
    ' 1. Abrir archivo fuente en modo solo lectura
    ' ----------------------------------------------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim wbFuente As Workbook
    Dim wsOrigen As Worksheet
    
    Application.StatusBar = "Abriendo archivo fuente (solo lectura)..."
    
    On Error Resume Next
    Set wbFuente = Workbooks.Open(Filename:=rutaFuente, ReadOnly:=True, UpdateLinks:=0)
    On Error GoTo ErrorHandler
    
    If wbFuente Is Nothing Then
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.StatusBar = False
        MsgBox "No se pudo abrir el archivo fuente." & vbCrLf & vbCrLf & _
               "Ruta configurada:" & vbCrLf & rutaFuente & vbCrLf & vbCrLf & _
               "Verifique que:" & vbCrLf & _
               "- La ruta es correcta" & vbCrLf & _
               "- El archivo existe en esa ubicacion" & vbCrLf & _
               "- Tiene acceso al archivo", _
               vbCritical, "Error al abrir archivo"
        Exit Sub
    End If
    
    ' Usar la primera hoja del archivo fuente
    Set wsOrigen = wbFuente.Sheets(1)
    
    Dim ultimaFila As Long
    Dim ultimaCol As Long
    ultimaFila = wsOrigen.Cells(wsOrigen.Rows.Count, 1).End(xlUp).Row
    ultimaCol = wsOrigen.Cells(1, wsOrigen.Columns.Count).End(xlToLeft).Column
    
    If ultimaFila < 2 Then
        wbFuente.Close SaveChanges:=False
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.StatusBar = False
        MsgBox "El archivo fuente no contiene datos (solo tiene encabezados o esta vacio).", _
               vbExclamation, "Sin datos"
        Exit Sub
    End If
    
    Application.StatusBar = "Localizando columnas..."
    
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
    Dim colDANE As Long
    Dim colDANESede As Long
    
    colPlaza = BuscarColumna(wsOrigen, "PLAZA")
    colSubregion = BuscarColumna(wsOrigen, "Subregi" & ChrW(243) & "n")
    If colSubregion = 0 Then colSubregion = BuscarColumna(wsOrigen, "Subregion")
    If colSubregion = 0 Then colSubregion = BuscarColumna(wsOrigen, "SUBREGION")
    
    colMunicipio = BuscarColumna(wsOrigen, "Municipio")
    If colMunicipio = 0 Then colMunicipio = BuscarColumna(wsOrigen, "MUNICIPIO")
    
    colEstablecimiento = BuscarColumna(wsOrigen, "Establecimiento educativo")
    If colEstablecimiento = 0 Then colEstablecimiento = BuscarColumna(wsOrigen, "Establecimiento Educativo")
    If colEstablecimiento = 0 Then colEstablecimiento = BuscarColumna(wsOrigen, "ESTABLECIMIENTO EDUCATIVO")
    
    colSede = BuscarColumna(wsOrigen, "Sede")
    If colSede = 0 Then colSede = BuscarColumna(wsOrigen, "SEDE")
    
    colZona = BuscarColumna(wsOrigen, "Zona")
    If colZona = 0 Then colZona = BuscarColumna(wsOrigen, "ZONA")
    
    colCargo = BuscarColumna(wsOrigen, "Cargo")
    If colCargo = 0 Then colCargo = BuscarColumna(wsOrigen, "CARGO")
    
    colTipoPlaza = BuscarColumna(wsOrigen, "Tipo de plaza")
    If colTipoPlaza = 0 Then colTipoPlaza = BuscarColumna(wsOrigen, "TIPO DE PLAZA")
    
    colNivelAcad = BuscarColumna(wsOrigen, "Nivel acad" & ChrW(233) & "mico")
    If colNivelAcad = 0 Then colNivelAcad = BuscarColumna(wsOrigen, "Nivel academico")
    If colNivelAcad = 0 Then colNivelAcad = BuscarColumna(wsOrigen, "NIVEL ACADEMICO")
    
    colAreaEduc = BuscarColumna(wsOrigen, ChrW(193) & "rea de conocimiento")
    If colAreaEduc = 0 Then colAreaEduc = BuscarColumna(wsOrigen, "Area de conocimiento")
    If colAreaEduc = 0 Then colAreaEduc = BuscarColumna(wsOrigen, "AREA DE CONOCIMIENTO")
    
    colMotivo = BuscarColumna(wsOrigen, "Motivo de la vacante")
    If colMotivo = 0 Then colMotivo = BuscarColumna(wsOrigen, "MOTIVO DE LA VACANTE")
    
    colFechaRegistro = BuscarColumna(wsOrigen, "Fecha de registro de la vacante")
    If colFechaRegistro = 0 Then colFechaRegistro = BuscarColumna(wsOrigen, "FECHA DE REGISTRO")
    
    colCedulaGenera = BuscarColumna(wsOrigen, "C" & ChrW(233) & "dula de quien genera la vacante")
    If colCedulaGenera = 0 Then colCedulaGenera = BuscarColumna(wsOrigen, "Cedula de quien genera la vacante")
    
    colNombreGenera = BuscarColumna(wsOrigen, "Nombre de quien genera la vacante")
    
    colActoAdmin = BuscarColumna(wsOrigen, "Acto administrativo")
    If colActoAdmin = 0 Then colActoAdmin = BuscarColumna(wsOrigen, "ACTO ADMINISTRATIVO")
    
    colFechaActo = BuscarColumna(wsOrigen, "Fecha del acto administrativo")
    
    colObservacion = BuscarColumna(wsOrigen, "Observaci" & ChrW(243) & "n")
    If colObservacion = 0 Then colObservacion = BuscarColumna(wsOrigen, "Observacion")
    If colObservacion = 0 Then colObservacion = BuscarColumna(wsOrigen, "OBSERVACION")
    
    colElegibles = BuscarColumna(wsOrigen, "Elegibles")
    If colElegibles = 0 Then colElegibles = BuscarColumna(wsOrigen, "ELEGIBLES")
    
    colObsPermanencia = BuscarColumna(wsOrigen, "Observaci" & ChrW(243) & "n permanencia")
    If colObsPermanencia = 0 Then colObsPermanencia = BuscarColumna(wsOrigen, "Observacion permanencia")
    
    colTomadaPor = BuscarColumna(wsOrigen, "Tomada por")
    If colTomadaPor = 0 Then colTomadaPor = BuscarColumna(wsOrigen, "TOMADA POR")
    
    colTomadaPara = BuscarColumna(wsOrigen, "Tomada para")
    If colTomadaPara = 0 Then colTomadaPara = BuscarColumna(wsOrigen, "TOMADA PARA")
    
    colEstado = BuscarColumna(wsOrigen, "Estado del nombramiento")
    If colEstado = 0 Then colEstado = BuscarColumna(wsOrigen, "ESTADO DEL NOMBRAMIENTO")
    
    colCedulaSel = BuscarColumna(wsOrigen, "C" & ChrW(233) & "dula del seleccionado")
    If colCedulaSel = 0 Then colCedulaSel = BuscarColumna(wsOrigen, "Cedula del seleccionado")
    
    colNombreSel = BuscarColumna(wsOrigen, "Nombre del seleccionado")
    
    colRegistradaPor = BuscarColumna(wsOrigen, "Registrada por")
    If colRegistradaPor = 0 Then colRegistradaPor = BuscarColumna(wsOrigen, "REGISTRADA POR")
    
    colDANE = BuscarColumna(wsOrigen, "C" & ChrW(243) & "digo DANE")
    If colDANE = 0 Then colDANE = BuscarColumna(wsOrigen, "Codigo DANE")
    If colDANE = 0 Then colDANE = BuscarColumna(wsOrigen, "CODIGO DANE")
    
    colDANESede = BuscarColumna(wsOrigen, "C" & ChrW(243) & "digo DANE sede")
    If colDANESede = 0 Then colDANESede = BuscarColumna(wsOrigen, "Codigo DANE sede")
    If colDANESede = 0 Then colDANESede = BuscarColumna(wsOrigen, "CODIGO DANE SEDE")
    
    ' Validar columnas criticas
    If colPlaza = 0 Then
        wbFuente.Close SaveChanges:=False
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.StatusBar = False
        MsgBox "No se encontro la columna 'PLAZA' en los encabezados del archivo fuente." & vbCrLf & _
               "Verifique que el archivo tiene la estructura correcta." & vbCrLf & vbCrLf & _
               "Columnas encontradas en fila 1: " & ultimaCol, _
               vbCritical, "Columna no encontrada"
        Exit Sub
    End If
    
    If colEstado = 0 Then
        wbFuente.Close SaveChanges:=False
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.StatusBar = False
        MsgBox "No se encontro la columna 'Estado del nombramiento' en el archivo fuente." & vbCrLf & _
               "Verifique que el archivo tiene la estructura correcta.", _
               vbCritical, "Columna no encontrada"
        Exit Sub
    End If
    
    ' ----------------------------------------------------------
    ' 3. Filtrar filas: plazas aun vacantes
    ' ----------------------------------------------------------
    Application.StatusBar = "Filtrando plazas vacantes..."
    
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
        
        ' Filtrar por subregion oficial
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
        wbFuente.Close SaveChanges:=False
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.StatusBar = False
        MsgBox "No se encontraron plazas vacantes con los criterios aplicados." & vbCrLf & vbCrLf & _
               "Criterios:" & vbCrLf & _
               "- Estado del nombramiento diferente de 'FIRMADO'" & vbCrLf & _
               "- Solo subregiones oficiales de Antioquia" & vbCrLf & _
               "- Plaza con numero valido", _
               vbInformation, "Sin resultados"
        Exit Sub
    End If
    
    ' ----------------------------------------------------------
    ' 4. Preparar hojas de resultados en ESTE archivo
    ' ----------------------------------------------------------
    Application.StatusBar = "Generando reporte (" & contVacantes & " plazas vacantes)..."
    
    ' Limpiar hojas existentes o crearlas
    Dim wsResultados As Worksheet
    Dim wsResumen As Worksheet
    
    ' Eliminar hojas previas si existen
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("Plazas Vacantes").Delete
    ThisWorkbook.Sheets("Resumen").Delete
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = True
    
    ' Crear hoja de resultados
    Set wsResultados = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsResultados.Name = "Plazas Vacantes"
    
    ' Crear hoja de resumen
    Set wsResumen = ThisWorkbook.Sheets.Add(After:=wsResultados)
    wsResumen.Name = "Resumen"
    
    ' ----------------------------------------------------------
    ' 5. Escribir encabezados en hoja de resultados
    ' ----------------------------------------------------------
    Dim encabezados As Variant
    encabezados = Array("PLAZA", "Subregi" & ChrW(243) & "n", "Municipio", _
                       "Establecimiento educativo", "C" & ChrW(243) & "digo DANE", _
                       "Sede", "C" & ChrW(243) & "digo DANE sede", "Zona", _
                       "Cargo", "Tipo de plaza", "Nivel acad" & ChrW(233) & "mico", _
                       ChrW(193) & "rea de conocimiento", _
                       "Motivo de la vacante", "Fecha de registro", _
                       "Acto administrativo", "Fecha acto admin.", _
                       "Observaci" & ChrW(243) & "n", "Elegibles", _
                       "Observaci" & ChrW(243) & "n permanencia", _
                       "Tomada por", "Tomada para", _
                       "Estado del nombramiento", _
                       "C" & ChrW(233) & "dula del seleccionado", _
                       "Nombre del seleccionado", "Registrada por")
    
    Dim colsOrigen As Variant
    colsOrigen = Array(colPlaza, colSubregion, colMunicipio, _
                       colEstablecimiento, colDANE, _
                       colSede, colDANESede, colZona, _
                       colCargo, colTipoPlaza, colNivelAcad, _
                       colAreaEduc, _
                       colMotivo, colFechaRegistro, _
                       colActoAdmin, colFechaActo, _
                       colObservacion, colElegibles, _
                       colObsPermanencia, _
                       colTomadaPor, colTomadaPara, _
                       colEstado, _
                       colCedulaSel, _
                       colNombreSel, colRegistradaPor)
    
    Dim numCols As Long
    numCols = UBound(encabezados) + 1
    
    Dim c As Long
    For c = 0 To UBound(encabezados)
        wsResultados.Cells(1, c + 1).Value = encabezados(c)
    Next c
    
    ' Formato encabezados
    With wsResultados.Range(wsResultados.Cells(1, 1), wsResultados.Cells(1, numCols))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 112, 60)  ' Verde oscuro
        .HorizontalAlignment = xlCenter
    End With
    
    ' ----------------------------------------------------------
    ' 6. Escribir datos de plazas vacantes
    ' ----------------------------------------------------------
    Dim filaDestino As Long
    Dim i As Long
    Dim colOrigen As Long
    
    For i = 1 To contVacantes
        filaDestino = i + 1
        For c = 0 To UBound(colsOrigen)
            colOrigen = CLng(colsOrigen(c))
            If colOrigen > 0 Then
                wsResultados.Cells(filaDestino, c + 1).Value = LeerCelda(wsOrigen, filasVacantes(i), colOrigen)
            End If
        Next c
        
        ' Actualizar barra de estado cada 50 filas
        If i Mod 50 = 0 Then
            Application.StatusBar = "Procesando fila " & i & " de " & contVacantes & "..."
        End If
    Next i
    
    ' ----------------------------------------------------------
    ' 7. Formato de la hoja de resultados
    ' ----------------------------------------------------------
    ' Autoajustar columnas
    wsResultados.Columns.AutoFit
    
    ' Bordes
    With wsResultados.Range(wsResultados.Cells(1, 1), wsResultados.Cells(contVacantes + 1, numCols)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(200, 200, 200)
    End With
    
    ' Filtros automaticos
    wsResultados.Range(wsResultados.Cells(1, 1), wsResultados.Cells(contVacantes + 1, numCols)).AutoFilter
    
    ' Inmovilizar primera fila
    wsResultados.Rows("2:2").Select
    ActiveWindow.FreezePanes = True
    
    ' Colores alternos para filas
    Dim r As Long
    For r = 2 To contVacantes + 1
        If r Mod 2 = 0 Then
            wsResultados.Range(wsResultados.Cells(r, 1), wsResultados.Cells(r, numCols)).Interior.Color = RGB(242, 248, 244)
        End If
    Next r
    
    ' ----------------------------------------------------------
    ' 8. Crear hoja de resumen
    ' ----------------------------------------------------------
    Application.StatusBar = "Generando resumen..."
    
    ' Titulo
    wsResumen.Range("A1").Value = "REPORTE DE PLAZAS VACANTES"
    wsResumen.Range("A1").Font.Bold = True
    wsResumen.Range("A1").Font.Size = 16
    wsResumen.Range("A1").Font.Color = RGB(0, 112, 60)
    
    wsResumen.Range("A2").Value = "Departamento de Antioquia"
    wsResumen.Range("A2").Font.Size = 12
    
    wsResumen.Range("A3").Value = "Generado: " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    wsResumen.Range("A3").Font.Italic = True
    wsResumen.Range("A3").Font.Color = RGB(100, 100, 100)
    
    wsResumen.Range("A4").Value = "Fuente: " & rutaFuente
    wsResumen.Range("A4").Font.Italic = True
    wsResumen.Range("A4").Font.Size = 9
    wsResumen.Range("A4").Font.Color = RGB(100, 100, 100)
    
    ' Metricas principales
    wsResumen.Range("A6").Value = "RESUMEN GENERAL"
    wsResumen.Range("A6").Font.Bold = True
    wsResumen.Range("A6").Font.Size = 13
    
    wsResumen.Range("A7").Value = "Total de registros en archivo fuente:"
    wsResumen.Range("B7").Value = ultimaFila - 1
    wsResumen.Range("B7").Font.Bold = True
    
    wsResumen.Range("A8").Value = "Total de plazas vacantes encontradas:"
    wsResumen.Range("B8").Value = contVacantes
    wsResumen.Range("B8").Font.Bold = True
    wsResumen.Range("B8").Font.Color = RGB(192, 0, 0)
    wsResumen.Range("B8").Font.Size = 14
    
    ' Contar por subregion
    wsResumen.Range("A10").Value = "VACANTES POR SUBREGION"
    wsResumen.Range("A10").Font.Bold = True
    wsResumen.Range("A10").Font.Size = 13
    
    wsResumen.Range("A11").Value = "Subregi" & ChrW(243) & "n"
    wsResumen.Range("B11").Value = "Cantidad"
    wsResumen.Range("A11").Font.Bold = True
    wsResumen.Range("B11").Font.Bold = True
    With wsResumen.Range("A11:B11")
        .Interior.Color = RGB(0, 112, 60)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    ' Contar vacantes por cada subregion
    Dim subregiones9 As Variant
    subregiones9 = Array("Bajo Cauca", "Magdalena Medio", "Nordeste", "Norte", _
                         "Occidente", "Oriente", "Suroeste", "Urab" & ChrW(225), _
                         "Valle de Aburr" & ChrW(225))
    
    Dim filaResumen As Long
    filaResumen = 12
    Dim contSub As Long
    Dim totalContado As Long
    totalContado = 0
    
    Dim j As Long
    For j = 0 To UBound(subregiones9)
        wsResumen.Cells(filaResumen, 1).Value = subregiones9(j)
        
        ' Contar plazas de esta subregion
        contSub = 0
        If colSubregion > 0 Then
            For i = 1 To contVacantes
                valorSubregion = UCase(Trim(LeerCelda(wsOrigen, filasVacantes(i), colSubregion)))
                If InStr(1, valorSubregion, UCase(CStr(subregiones9(j))), vbTextCompare) > 0 Then
                    contSub = contSub + 1
                End If
            Next i
            ' Manejo especial para Uraba y Valle de Aburra (tildes)
            If UCase(CStr(subregiones9(j))) = "URAB" & UCase(ChrW(225)) Then
                For i = 1 To contVacantes
                    valorSubregion = UCase(Trim(LeerCelda(wsOrigen, filasVacantes(i), colSubregion)))
                    If InStr(1, valorSubregion, "URABA", vbTextCompare) > 0 And contSub = 0 Then
                        contSub = contSub + 1
                    End If
                Next i
            End If
        End If
        
        wsResumen.Cells(filaResumen, 2).Value = contSub
        totalContado = totalContado + contSub
        filaResumen = filaResumen + 1
    Next j
    
    ' Fila de total
    wsResumen.Cells(filaResumen, 1).Value = "TOTAL"
    wsResumen.Cells(filaResumen, 1).Font.Bold = True
    wsResumen.Cells(filaResumen, 2).Value = contVacantes
    wsResumen.Cells(filaResumen, 2).Font.Bold = True
    With wsResumen.Range(wsResumen.Cells(filaResumen, 1), wsResumen.Cells(filaResumen, 2))
        .Interior.Color = RGB(220, 230, 220)
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlMedium
    End With
    
    ' Bordes tabla subregiones
    With wsResumen.Range("A11:B" & filaResumen).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    ' Criterios utilizados
    filaResumen = filaResumen + 2
    wsResumen.Cells(filaResumen, 1).Value = "CRITERIOS DE FILTRADO"
    wsResumen.Cells(filaResumen, 1).Font.Bold = True
    wsResumen.Cells(filaResumen, 1).Font.Size = 13
    
    filaResumen = filaResumen + 1
    wsResumen.Cells(filaResumen, 1).Value = "1. Estado del nombramiento diferente de 'FIRMADO'"
    filaResumen = filaResumen + 1
    wsResumen.Cells(filaResumen, 1).Value = "   (incluye: Seleccionar, EN PROCESO, EN FIRMAS, vac" & ChrW(237) & "o)"
    filaResumen = filaResumen + 1
    wsResumen.Cells(filaResumen, 1).Value = "2. Solo las 9 subregiones oficiales de Antioquia"
    filaResumen = filaResumen + 1
    wsResumen.Cells(filaResumen, 1).Value = "3. Registros con n" & ChrW(250) & "mero de PLAZA v" & ChrW(225) & "lido (no vac" & ChrW(237) & "o ni 0)"
    
    ' Autoajustar
    wsResumen.Columns("A:B").AutoFit
    
    ' ----------------------------------------------------------
    ' 9. Cerrar archivo fuente SIN guardar cambios
    ' ----------------------------------------------------------
    wbFuente.Close SaveChanges:=False
    
    ' ----------------------------------------------------------
    ' 10. Activar hoja de resultados
    ' ----------------------------------------------------------
    wsResultados.Activate
    wsResultados.Range("A1").Select
    
    ' Restaurar configuracion
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    
    MsgBox "Reporte generado exitosamente." & vbCrLf & vbCrLf & _
           "Plazas vacantes encontradas: " & contVacantes & vbCrLf & vbCrLf & _
           "Los resultados estan en las hojas:" & vbCrLf & _
           "- 'Plazas Vacantes': listado completo" & vbCrLf & _
           "- 'Resumen': totales por subregion" & vbCrLf & vbCrLf & _
           "Recuerde guardar este archivo si desea conservar los resultados.", _
           vbInformation, "Reporte completado"
    
    Exit Sub
    
ErrorHandler:
    ' Cerrar archivo fuente si quedo abierto
    On Error Resume Next
    If Not wbFuente Is Nothing Then
        wbFuente.Close SaveChanges:=False
    End If
    On Error GoTo 0
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    Application.DisplayAlerts = True
    
    MsgBox "Ocurrio un error inesperado:" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Si el error persiste, verifique:" & vbCrLf & _
           "- Que la ruta del archivo fuente es correcta" & vbCrLf & _
           "- Que el archivo fuente no esta danado" & vbCrLf & _
           "- Que tiene permisos de lectura", _
           vbCritical, "Error"

Limpiar:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
End Sub
