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
' CRITERIOS DE FILTRADO (tabla verde, columna Y en adelante):
' Una plaza se considera VACANTE si cumple CUALQUIERA de estos:
'   1. "Estado del nombramiento" esta VACIO o dice "Seleccionar"
'   2. "Vacante tomada por" esta VACIO o contiene "SISTEMA MAESTRO"
'
' Adicionalmente:
'   - Solo se incluyen las 9 subregiones oficiales de Antioquia
'   - El registro debe tener un numero de PLAZA valido
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
' Busca coincidencia exacta (case-insensitive)
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
' Buscar columna por coincidencia parcial (contiene el texto)
' ----------------------------------------------------------------
Private Function BuscarColumnaContiene(ws As Worksheet, ByVal texto As String) As Long
    Dim col As Long
    Dim ultimaCol As Long
    
    ultimaCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    For col = 1 To ultimaCol
        If InStr(1, LCase(Trim(CStr(ws.Cells(1, col).Value))), LCase(Trim(texto)), vbTextCompare) > 0 Then
            BuscarColumnaContiene = col
            Exit Function
        End If
    Next col
    
    BuscarColumnaContiene = 0
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

' ----------------------------------------------------------------
' Extraer solo el nombre del archivo de una ruta completa
' ----------------------------------------------------------------
Private Function ExtraerNombreArchivo(ByVal ruta As String) As String
    Dim pos As Long
    pos = InStrRev(ruta, "\")
    If pos = 0 Then pos = InStrRev(ruta, "/")
    If pos > 0 Then
        ExtraerNombreArchivo = Mid(ruta, pos + 1)
    Else
        ExtraerNombreArchivo = ruta
    End If
End Function

' ----------------------------------------------------------------
' Buscar si el archivo ya esta abierto en Excel
' ----------------------------------------------------------------
Private Function BuscarLibroAbierto(ByVal nombreArchivo As String) As Workbook
    Dim wb As Workbook
    
    On Error Resume Next
    For Each wb In Workbooks
        If LCase(wb.Name) = LCase(nombreArchivo) Then
            Set BuscarLibroAbierto = wb
            Exit Function
        End If
    Next wb
    On Error GoTo 0
    
    Set BuscarLibroAbierto = Nothing
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
    
    ' Verificar que el archivo existe (solo para rutas locales)
    Dim archivoExiste As Boolean
    archivoExiste = False
    
    On Error Resume Next
    If Dir(rutaFuente) <> "" Then archivoExiste = True
    On Error GoTo ErrorHandler
    
    If Not archivoExiste Then
        If InStr(rutaFuente, ".") = 0 Then
            MsgBox "La ruta en B2 no parece incluir el nombre del archivo con su extension (.xlsx)." & vbCrLf & vbCrLf & _
                   "Ruta actual: " & rutaFuente & vbCrLf & vbCrLf & _
                   "La ruta debe terminar en .xlsx (ejemplo: C:\Carpeta\archivo.xlsx)", _
                   vbExclamation, "Ruta incompleta"
            Exit Sub
        End If
    End If
    
    ' ----------------------------------------------------------
    ' 1. Obtener referencia al archivo fuente
    ' ----------------------------------------------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim wbFuente As Workbook
    Dim wsOrigen As Worksheet
    Dim yaEstabAbierto As Boolean
    Dim nombreArchivo As String
    Dim errNum As Long
    Dim errDesc As String
    
    yaEstabAbierto = False
    nombreArchivo = ExtraerNombreArchivo(rutaFuente)
    
    ' Primero: verificar si el archivo ya esta abierto en Excel
    Application.StatusBar = "Buscando archivo fuente..."
    Set wbFuente = BuscarLibroAbierto(nombreArchivo)
    
    If Not wbFuente Is Nothing Then
        yaEstabAbierto = True
        Application.StatusBar = "Archivo fuente encontrado (ya abierto)..."
    Else
        ' Intentar abrir el archivo
        Application.StatusBar = "Abriendo archivo fuente (solo lectura)..."
        Application.DisplayAlerts = False
        
        On Error Resume Next
        Set wbFuente = Workbooks.Open( _
            Filename:=rutaFuente, _
            ReadOnly:=True, _
            UpdateLinks:=0, _
            CorruptLoad:=xlNormalLoad)
        errNum = Err.Number
        errDesc = Err.Description
        On Error GoTo ErrorHandler
        
        Application.DisplayAlerts = True
        
        If wbFuente Is Nothing Then
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Application.StatusBar = False
            MsgBox "No se pudo abrir el archivo fuente." & vbCrLf & vbCrLf & _
                   "Ruta configurada:" & vbCrLf & rutaFuente & vbCrLf & vbCrLf & _
                   "Error: " & errNum & " - " & errDesc & vbCrLf & vbCrLf & _
                   "SOLUCION: Abra el archivo '" & nombreArchivo & "' manualmente " & _
                   "en Excel (doble clic), habilite edicion si lo pide, " & _
                   "dejelo abierto, y luego presione este boton de nuevo.", _
                   vbCritical, "Error al abrir archivo"
            Exit Sub
        End If
    End If
    
    ' Usar la primera hoja del archivo fuente
    Set wsOrigen = wbFuente.Sheets(1)
    
    Dim ultimaFila As Long
    Dim ultimaCol As Long
    ultimaFila = wsOrigen.Cells(wsOrigen.Rows.Count, 1).End(xlUp).Row
    ultimaCol = wsOrigen.Cells(1, wsOrigen.Columns.Count).End(xlToLeft).Column
    
    If ultimaFila < 2 Then
        If Not yaEstabAbierto Then wbFuente.Close SaveChanges:=False
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
    ' --- Tabla azul (columnas A-X): informacion basica de la plaza ---
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
    Dim colActoAdminVacante As Long
    Dim colFechaActoVacante As Long
    Dim colObservacionAzul As Long
    Dim colRegistradaPor As Long
    
    ' --- Tabla verde (columnas Y en adelante): gestion de la vacante ---
    Dim colElegibles As Long
    Dim colObsPermanencia As Long
    Dim colVacanteTomadaPor As Long
    Dim colVacanteTomadaPara As Long
    Dim colObservacionVerde As Long
    Dim colOPEC As Long
    Dim colPosElegible As Long
    Dim colCedulaSel As Long
    Dim colNombreSel As Long
    Dim colCorreoSel As Long
    Dim colCelularSel As Long
    Dim colEstadoNombramiento As Long
    Dim colActoAdminNombram As Long
    Dim colFechaActoNombram As Long
    Dim colNovedadNombram As Long
    
    ' Buscar columnas tabla azul
    colPlaza = BuscarColumna(wsOrigen, "PLAZA")
    
    colSubregion = BuscarColumnaContiene(wsOrigen, "Subregi")
    If colSubregion = 0 Then colSubregion = BuscarColumna(wsOrigen, "Subregion")
    
    colMunicipio = BuscarColumna(wsOrigen, "Municipio")
    If colMunicipio = 0 Then colMunicipio = BuscarColumnaContiene(wsOrigen, "Municipio")
    
    colEstablecimiento = BuscarColumnaContiene(wsOrigen, "Establecimiento")
    
    colSede = BuscarColumna(wsOrigen, "Sede")
    If colSede = 0 Then colSede = BuscarColumna(wsOrigen, "SEDE")
    
    colZona = BuscarColumna(wsOrigen, "Zona")
    If colZona = 0 Then colZona = BuscarColumna(wsOrigen, "ZONA")
    
    colCargo = BuscarColumnaContiene(wsOrigen, "Cargo de la vacante")
    If colCargo = 0 Then colCargo = BuscarColumnaContiene(wsOrigen, "Cargo")
    
    colTipoPlaza = BuscarColumnaContiene(wsOrigen, "Tipo de plaza")
    
    colNivelAcad = BuscarColumnaContiene(wsOrigen, "Nivel acad")
    
    colAreaEduc = BuscarColumnaContiene(wsOrigen, "rea de conocimiento")
    If colAreaEduc = 0 Then colAreaEduc = BuscarColumnaContiene(wsOrigen, "Area de conocimiento")
    
    colMotivo = BuscarColumnaContiene(wsOrigen, "Motivo de la vacante")
    
    colFechaRegistro = BuscarColumnaContiene(wsOrigen, "Fecha de registro de la vacante")
    If colFechaRegistro = 0 Then colFechaRegistro = BuscarColumnaContiene(wsOrigen, "Fecha de registro")
    
    colCedulaGenera = BuscarColumnaContiene(wsOrigen, "dula de qui")
    If colCedulaGenera = 0 Then colCedulaGenera = BuscarColumnaContiene(wsOrigen, "Cedula de quien genera")
    
    colNombreGenera = BuscarColumnaContiene(wsOrigen, "Nombre de qui")
    If colNombreGenera = 0 Then colNombreGenera = BuscarColumnaContiene(wsOrigen, "Nombre de quien genera")
    
    colActoAdminVacante = BuscarColumna(wsOrigen, "Acto administrativo de la vacante")
    If colActoAdminVacante = 0 Then colActoAdminVacante = BuscarColumna(wsOrigen, "Acto administrativo")
    
    colFechaActoVacante = BuscarColumna(wsOrigen, "Fecha del acto administrativo")
    
    colObservacionAzul = BuscarColumnaContiene(wsOrigen, "Observaci")
    If colObservacionAzul = 0 Then colObservacionAzul = BuscarColumna(wsOrigen, "Observacion")
    
    colRegistradaPor = BuscarColumnaContiene(wsOrigen, "Vacante temporal registrada por")
    If colRegistradaPor = 0 Then colRegistradaPor = BuscarColumnaContiene(wsOrigen, "registrada por")
    
    ' Buscar columnas tabla verde (columna Y en adelante)
    colElegibles = BuscarColumnaContiene(wsOrigen, "Tiene lista de Elegibles")
    If colElegibles = 0 Then colElegibles = BuscarColumnaContiene(wsOrigen, "Elegibles")
    
    colObsPermanencia = BuscarColumnaContiene(wsOrigen, "Permanencia")
    
    colVacanteTomadaPor = BuscarColumna(wsOrigen, "Vacante tomada por")
    If colVacanteTomadaPor = 0 Then colVacanteTomadaPor = BuscarColumnaContiene(wsOrigen, "Vacante tomada por")
    
    colVacanteTomadaPara = BuscarColumna(wsOrigen, "Vacante tomada para")
    If colVacanteTomadaPara = 0 Then colVacanteTomadaPara = BuscarColumnaContiene(wsOrigen, "Vacante tomada para")
    
    ' Observacion verde (col AC) - es diferente a la observacion azul (col W)
    ' Buscar especificamente en columnas > 24 (despues de X)
    Dim colTemp As Long
    colObservacionVerde = 0
    For colTemp = 25 To ultimaCol
        Dim headerTemp As String
        headerTemp = LCase(Trim(CStr(wsOrigen.Cells(1, colTemp).Value)))
        If Left(headerTemp, 10) = "observaci" & ChrW(243) & "n" Or headerTemp = "observacion" Or Left(headerTemp, 11) = "observaci" & ChrW(243) & "n " Then
            ' Verificar que no es "Observacion Direccion de Permanencia"
            If InStr(1, headerTemp, "permanencia", vbTextCompare) = 0 Then
                colObservacionVerde = colTemp
                Exit For
            End If
        End If
    Next colTemp
    
    colOPEC = BuscarColumna(wsOrigen, "OPEC")
    
    colPosElegible = BuscarColumnaContiene(wsOrigen, "Posici")
    If colPosElegible = 0 Then colPosElegible = BuscarColumnaContiene(wsOrigen, "elegible")
    
    ' Cedula del seleccionado (col AF) - es "C?dula" en la tabla verde
    colCedulaSel = BuscarColumnaContiene(wsOrigen, "dula")
    ' Puede confundirse con col S "Cedula de quien genera" - buscar en zona verde
    colCedulaSel = 0
    For colTemp = 25 To ultimaCol
        headerTemp = LCase(Trim(CStr(wsOrigen.Cells(1, colTemp).Value)))
        If InStr(1, headerTemp, "dula", vbTextCompare) > 0 Then
            If InStr(1, headerTemp, "quien", vbTextCompare) = 0 Then
                colCedulaSel = colTemp
                Exit For
            End If
        End If
    Next colTemp
    
    colNombreSel = BuscarColumna(wsOrigen, "Nombre del seleccionado")
    If colNombreSel = 0 Then colNombreSel = BuscarColumnaContiene(wsOrigen, "Nombre del seleccionado")
    
    colCorreoSel = 0
    For colTemp = 25 To ultimaCol
        headerTemp = LCase(Trim(CStr(wsOrigen.Cells(1, colTemp).Value)))
        If InStr(1, headerTemp, "correo", vbTextCompare) > 0 Then
            colCorreoSel = colTemp
            Exit For
        End If
    Next colTemp
    
    colCelularSel = BuscarColumna(wsOrigen, "Celular")
    If colCelularSel = 0 Then colCelularSel = BuscarColumnaContiene(wsOrigen, "Celular")
    
    colEstadoNombramiento = BuscarColumna(wsOrigen, "Estado del nombramiento")
    If colEstadoNombramiento = 0 Then colEstadoNombramiento = BuscarColumnaContiene(wsOrigen, "Estado del nombramiento")
    
    colActoAdminNombram = 0
    For colTemp = 25 To ultimaCol
        headerTemp = LCase(Trim(CStr(wsOrigen.Cells(1, colTemp).Value)))
        If InStr(1, headerTemp, "acto administrativo", vbTextCompare) > 0 Then
            colActoAdminNombram = colTemp
            Exit For
        End If
    Next colTemp
    
    colFechaActoNombram = 0
    For colTemp = 25 To ultimaCol
        headerTemp = LCase(Trim(CStr(wsOrigen.Cells(1, colTemp).Value)))
        If InStr(1, headerTemp, "fecha", vbTextCompare) > 0 And InStr(1, headerTemp, "acto", vbTextCompare) > 0 Then
            colFechaActoNombram = colTemp
            Exit For
        End If
    Next colTemp
    
    colNovedadNombram = BuscarColumnaContiene(wsOrigen, "Novedad del nombramiento")
    If colNovedadNombram = 0 Then colNovedadNombram = BuscarColumnaContiene(wsOrigen, "Novedad")
    
    ' Validar columnas criticas
    If colPlaza = 0 Then
        If Not yaEstabAbierto Then wbFuente.Close SaveChanges:=False
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.StatusBar = False
        MsgBox "No se encontro la columna 'PLAZA' en los encabezados del archivo fuente." & vbCrLf & _
               "Verifique que el archivo tiene la estructura correcta." & vbCrLf & vbCrLf & _
               "Columnas encontradas en fila 1: " & ultimaCol, _
               vbCritical, "Columna no encontrada"
        Exit Sub
    End If
    
    If colEstadoNombramiento = 0 And colVacanteTomadaPor = 0 Then
        If Not yaEstabAbierto Then wbFuente.Close SaveChanges:=False
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.StatusBar = False
        MsgBox "No se encontraron las columnas de la tabla verde:" & vbCrLf & _
               "- 'Estado del nombramiento'" & vbCrLf & _
               "- 'Vacante tomada por'" & vbCrLf & vbCrLf & _
               "Verifique que el archivo tiene la estructura correcta.", _
               vbCritical, "Columnas no encontradas"
        Exit Sub
    End If
    
    ' ----------------------------------------------------------
    ' 3. Filtrar filas: plazas aun vacantes (NUEVOS CRITERIOS)
    ' ----------------------------------------------------------
    ' Criterio: una plaza se considera VACANTE si:
    '   - "Estado del nombramiento" esta VACIO o dice "Seleccionar"
    '   - O "Vacante tomada por" esta VACIO o contiene "SISTEMA MAESTRO"
    ' ----------------------------------------------------------
    Application.StatusBar = "Filtrando plazas vacantes..."
    
    Dim filasVacantes() As Long
    ReDim filasVacantes(1 To ultimaFila)
    Dim contVacantes As Long
    contVacantes = 0
    
    Dim fila As Long
    Dim valorPlaza As String
    Dim valorEstado As String
    Dim valorTomadaPor As String
    Dim valorSubregion As String
    Dim esVacante As Boolean
    
    For fila = 2 To ultimaFila
        ' Saltar filas sin numero de plaza
        valorPlaza = LeerCelda(wsOrigen, fila, colPlaza)
        If valorPlaza = "" Or valorPlaza = "0" Then GoTo SiguienteFila
        
        ' Filtrar por subregion oficial
        If colSubregion > 0 Then
            valorSubregion = LeerCelda(wsOrigen, fila, colSubregion)
            If valorSubregion <> "" Then
                If Not EsSubregionOficial(valorSubregion) Then GoTo SiguienteFila
            End If
        End If
        
        ' NUEVOS CRITERIOS DE VACANTE (tabla verde)
        esVacante = False
        
        ' Criterio 1: Estado del nombramiento vacio o "Seleccionar"
        If colEstadoNombramiento > 0 Then
            valorEstado = UCase(Trim(LeerCelda(wsOrigen, fila, colEstadoNombramiento)))
            If valorEstado = "" Or valorEstado = "SELECCIONAR" Then
                esVacante = True
            End If
        End If
        
        ' Criterio 2: Vacante tomada por vacio o contiene "SISTEMA MAESTRO"
        If Not esVacante Then
            If colVacanteTomadaPor > 0 Then
                valorTomadaPor = UCase(Trim(LeerCelda(wsOrigen, fila, colVacanteTomadaPor)))
                If valorTomadaPor = "" Or InStr(1, valorTomadaPor, "SISTEMA MAESTRO", vbTextCompare) > 0 Then
                    esVacante = True
                End If
            End If
        End If
        
        ' Si cumple alguno de los criterios, es vacante
        If esVacante Then
            contVacantes = contVacantes + 1
            filasVacantes(contVacantes) = fila
        End If
        
SiguienteFila:
    Next fila
    
    If contVacantes = 0 Then
        If Not yaEstabAbierto Then wbFuente.Close SaveChanges:=False
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.StatusBar = False
        MsgBox "No se encontraron plazas vacantes con los criterios aplicados." & vbCrLf & vbCrLf & _
               "Criterios:" & vbCrLf & _
               "- Estado del nombramiento: vac" & ChrW(237) & "o o 'Seleccionar'" & vbCrLf & _
               "- Vacante tomada por: vac" & ChrW(237) & "o o 'SISTEMA MAESTRO'" & vbCrLf & _
               "- Solo subregiones oficiales de Antioquia" & vbCrLf & _
               "- Plaza con n" & ChrW(250) & "mero v" & ChrW(225) & "lido", _
               vbInformation, "Sin resultados"
        Exit Sub
    End If
    
    ' ----------------------------------------------------------
    ' 4. Preparar hojas de resultados en ESTE archivo
    ' ----------------------------------------------------------
    Application.StatusBar = "Generando reporte (" & contVacantes & " plazas vacantes)..."
    
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
        "Motivo de la vacante", _
        "Fecha de registro", _
        "Acto admin. vacante", _
        "Fecha acto admin.", _
        "Observaci" & ChrW(243) & "n (azul)", _
        "Registrada por", _
        "Tiene lista Elegibles", _
        "Obs. Permanencia", _
        "Vacante tomada por", _
        "Vacante tomada para", _
        "Observaci" & ChrW(243) & "n (verde)", _
        "OPEC", _
        "Posici" & ChrW(243) & "n elegible", _
        "C" & ChrW(233) & "dula seleccionado", _
        "Nombre del seleccionado", _
        "Correo seleccionado", _
        "Celular", _
        "Estado del nombramiento", _
        "Acto admin. nombramiento", _
        "Fecha acto nombramiento", _
        "Novedad del nombramiento")
    
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
        colMotivo, _
        colFechaRegistro, _
        colActoAdminVacante, _
        colFechaActoVacante, _
        colObservacionAzul, _
        colRegistradaPor, _
        colElegibles, _
        colObsPermanencia, _
        colVacanteTomadaPor, _
        colVacanteTomadaPara, _
        colObservacionVerde, _
        colOPEC, _
        colPosElegible, _
        colCedulaSel, _
        colNombreSel, _
        colCorreoSel, _
        colCelularSel, _
        colEstadoNombramiento, _
        colActoAdminNombram, _
        colFechaActoNombram, _
        colNovedadNombram)
    
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
        .Interior.Color = RGB(0, 112, 60)
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
        
        If i Mod 50 = 0 Then
            Application.StatusBar = "Procesando fila " & i & " de " & contVacantes & "..."
        End If
    Next i
    
    ' ----------------------------------------------------------
    ' 7. Formato de la hoja de resultados
    ' ----------------------------------------------------------
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
    wsResultados.Activate
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
    
    wsResumen.Activate
    
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
                If contSub = 0 Then
                    For i = 1 To contVacantes
                        valorSubregion = UCase(Trim(LeerCelda(wsOrigen, filasVacantes(i), colSubregion)))
                        If InStr(1, valorSubregion, "URABA", vbTextCompare) > 0 Then
                            contSub = contSub + 1
                        End If
                    Next i
                End If
            End If
            If UCase(CStr(subregiones9(j))) = "VALLE DE ABURR" & UCase(ChrW(225)) Then
                If contSub = 0 Then
                    For i = 1 To contVacantes
                        valorSubregion = UCase(Trim(LeerCelda(wsOrigen, filasVacantes(i), colSubregion)))
                        If InStr(1, valorSubregion, "ABURRA", vbTextCompare) > 0 Or _
                           InStr(1, valorSubregion, "ABURR", vbTextCompare) > 0 Then
                            contSub = contSub + 1
                        End If
                    Next i
                End If
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
    wsResumen.Cells(filaResumen, 1).Value = "CRITERIOS DE FILTRADO (tabla verde)"
    wsResumen.Cells(filaResumen, 1).Font.Bold = True
    wsResumen.Cells(filaResumen, 1).Font.Size = 13
    
    filaResumen = filaResumen + 1
    wsResumen.Cells(filaResumen, 1).Value = "Se considera VACANTE si cumple CUALQUIERA de estos criterios:"
    wsResumen.Cells(filaResumen, 1).Font.Italic = True
    
    filaResumen = filaResumen + 1
    wsResumen.Cells(filaResumen, 1).Value = "1. 'Estado del nombramiento' est" & ChrW(225) & " vac" & ChrW(237) & "o o dice 'Seleccionar'"
    filaResumen = filaResumen + 1
    wsResumen.Cells(filaResumen, 1).Value = "2. 'Vacante tomada por' est" & ChrW(225) & " vac" & ChrW(237) & "o o contiene 'SISTEMA MAESTRO'"
    filaResumen = filaResumen + 1
    wsResumen.Cells(filaResumen, 1).Value = ""
    filaResumen = filaResumen + 1
    wsResumen.Cells(filaResumen, 1).Value = "Filtros adicionales:"
    wsResumen.Cells(filaResumen, 1).Font.Bold = True
    filaResumen = filaResumen + 1
    wsResumen.Cells(filaResumen, 1).Value = "3. Solo las 9 subregiones oficiales de Antioquia"
    filaResumen = filaResumen + 1
    wsResumen.Cells(filaResumen, 1).Value = "4. Registros con n" & ChrW(250) & "mero de PLAZA v" & ChrW(225) & "lido (no vac" & ChrW(237) & "o ni 0)"
    
    ' Autoajustar
    wsResumen.Columns("A:B").AutoFit
    
    ' ----------------------------------------------------------
    ' 9. Cerrar archivo fuente SIN guardar cambios
    ' ----------------------------------------------------------
    If Not yaEstabAbierto Then
        wbFuente.Close SaveChanges:=False
    End If
    
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
           "- 'Plazas Vacantes': listado completo con 30 columnas" & vbCrLf & _
           "- 'Resumen': totales por subregion" & vbCrLf & vbCrLf & _
           "Criterios usados:" & vbCrLf & _
           "- Estado del nombramiento vac" & ChrW(237) & "o o 'Seleccionar'" & vbCrLf & _
           "- Vacante tomada por vac" & ChrW(237) & "o o 'SISTEMA MAESTRO'" & vbCrLf & vbCrLf & _
           "Recuerde guardar este archivo si desea conservar los resultados.", _
           vbInformation, "Reporte completado"
    
    Exit Sub
    
ErrorHandler:
    Dim finalErrNum As Long
    Dim finalErrDesc As String
    finalErrNum = Err.Number
    finalErrDesc = Err.Description
    
    On Error Resume Next
    If Not wbFuente Is Nothing Then
        If Not yaEstabAbierto Then
            wbFuente.Close SaveChanges:=False
        End If
    End If
    On Error GoTo 0
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    Application.DisplayAlerts = True
    
    If finalErrNum = 0 And finalErrDesc = "" Then
        MsgBox "El archivo se abrio en Vista Protegida y no se pueden leer los datos." & vbCrLf & vbCrLf & _
               "SOLUCION:" & vbCrLf & _
               "1. Abra manualmente el archivo '" & nombreArchivo & "'" & vbCrLf & _
               "2. Haga clic en 'Habilitar edicion' (barra amarilla arriba)" & vbCrLf & _
               "3. Deje el archivo abierto" & vbCrLf & _
               "4. Vuelva a este archivo y presione el boton de nuevo", _
               vbExclamation, "Vista Protegida"
    Else
        MsgBox "Ocurrio un error:" & vbCrLf & vbCrLf & _
               "Error " & finalErrNum & ": " & finalErrDesc & vbCrLf & vbCrLf & _
               "SOLUCION:" & vbCrLf & _
               "1. Abra manualmente el archivo '" & nombreArchivo & "'" & vbCrLf & _
               "2. Haga clic en 'Habilitar edicion' si aparece" & vbCrLf & _
               "3. Deje el archivo abierto" & vbCrLf & _
               "4. Vuelva a este archivo y presione el boton de nuevo", _
               vbCritical, "Error"
    End If

Limpiar:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
End Sub
