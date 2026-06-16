# Guía: Configurar archivo de reporte EXTERNO (sin tocar el archivo original)

## Concepto

Se crea un archivo Excel **aparte** llamado "Reporte Plazas Vacantes.xlsm". Este archivo:
- Tiene un **botón** que al presionarlo se conecta al archivo "Vacantes Definitivas" en SharePoint
- Abre el archivo fuente en **modo solo lectura** (NO lo modifica)
- Lee los datos, filtra las plazas vacantes, y muestra los resultados **dentro de sí mismo**
- El archivo original de vacantes definitivas **nunca se toca**

---

## Paso 1: Crear el archivo de reporte

1. Abrir **Excel** (aplicación de escritorio)
2. Crear un **nuevo libro en blanco**: Archivo → Nuevo → Libro en blanco
3. Guardar inmediatamente como:
   - Nombre: **Reporte Plazas Vacantes**
   - Tipo: **Libro de Excel habilitado para macros (.xlsm)**
   - Ubicación: Tu escritorio (luego lo subes a SharePoint)

---

## Paso 2: Crear la hoja "Configuracion"

Esta hoja es donde pones la ruta del archivo de Vacantes Definitivas.

1. La primera hoja del libro (Hoja1) la vas a **renombrar**:
   - Clic derecho sobre la pestaña "Hoja1" abajo → "Cambiar nombre"
   - Escribir: **Configuracion** (sin tilde, exactamente así)
   - Presionar Enter

2. En esa hoja, escribir lo siguiente:

| Celda | Contenido |
|-------|-----------|
| **A1** | Ruta del archivo fuente: |
| **B1** | (dejar vacío — es solo el título) |
| **A2** | Archivo Vacantes Definitivas: |
| **B2** | *(aquí va la ruta — ver Paso 3)* |
| **A4** | **Instrucciones:** |
| **A5** | En la celda B2, escriba la ruta completa del archivo |
| **A6** | "Registro de vacantes DEFINITIVAS" en SharePoint. |
| **A7** | |
| **A8** | Para obtener la ruta correcta, vea las instrucciones abajo. |

3. Darle formato a la celda **A1** en negrita
4. Darle un color de fondo a la celda **B2** (por ejemplo amarillo claro) para que sea fácil de encontrar

---

## Paso 3: Obtener la ruta del archivo en SharePoint

### Opción A: Ruta sincronizada con OneDrive (RECOMENDADA — más fácil)

Si tienes el SharePoint **sincronizado con OneDrive** en tu computador:

1. Abre el **Explorador de archivos** de Windows
2. Navega hasta la carpeta sincronizada de SharePoint donde está el archivo "Vacantes Definitivas"
3. Busca el archivo "Registro de vacantes DEFINITIVAS 2026.xlsx"
4. **Clic derecho** sobre el archivo → "Copiar como ruta de acceso" (o simplemente mira la barra de direcciones)
5. La ruta se verá algo así:
   ```
   C:\Users\TuNombre\Tu Organizacion\NombreSitio - Documentos\Registro de vacantes DEFINITIVAS 2026.xlsx
   ```
6. **Pega esa ruta en la celda B2** de la hoja "Configuracion"

### Opción B: Ruta directa de SharePoint (si NO tienes sincronizado)

1. Ve a SharePoint en el navegador
2. Busca el archivo "Registro de vacantes DEFINITIVAS 2026"
3. Haz clic en los **tres puntos (...)** junto al archivo → **"Detalles"** o **"Copiar ruta"**
4. La ruta de SharePoint se ve algo así:
   ```
   https://tuorganizacion.sharepoint.com/sites/NombreSitio/Documentos compartidos/Registro de vacantes DEFINITIVAS 2026.xlsx
   ```
5. **Pega esa ruta en la celda B2**

### Opción C: Archivo local (para pruebas)

Si tienes una copia del archivo en tu computador:
```
C:\Users\TuNombre\Desktop\Registro de vacantes DEFINITIVAS 2026.xlsx
```

> **IMPORTANTE**: La ruta debe incluir el nombre del archivo con su extensión (.xlsx)

---

## Paso 4: Abrir el Editor de Visual Basic

1. Presiona **Alt + F11**
2. Se abre el Editor de Visual Basic

---

## Paso 5: Importar la macro

**Opción A: Importar archivo .bas**
1. En el editor: **Archivo → Importar archivo...**
2. Selecciona el archivo **`MacroReporteExterno.bas`**
3. Se crea el módulo "ModuloReporteExterno"

**Opción B: Copiar y pegar**
1. Clic derecho sobre "VBAProject (Reporte Plazas Vacantes.xlsm)" → **Insertar → Módulo**
2. Abre `MacroReporteExterno.bas` con Bloc de notas
3. Seleccionar todo (Ctrl+A) y copiar (Ctrl+C)
4. Pegar en la ventana del módulo (Ctrl+V)

---

## Paso 6: Cerrar el Editor de VBA

1. Presiona **Alt + F11** (o cierra la ventana)

---

## Paso 7: Habilitar la pestaña "Programador"

Si no la tienes visible:
1. **Archivo → Opciones → Personalizar cinta de opciones**
2. Marca la casilla **"Programador"**
3. Clic en **Aceptar**

---

## Paso 8: Crear el botón

1. Selecciona la hoja **"Configuracion"** (o crea una hoja nueva para el botón si prefieres)
2. Ve a la pestaña **Programador**
3. Clic en **Insertar** → selecciona **Botón (control de formulario)** (primer icono)
4. Dibuja el botón en un área amplia y visible (por ejemplo, en las celdas D2:F4)
5. En la ventana "Asignar macro", selecciona **`GenerarReportePlazasVacantes`**
6. Clic en **Aceptar**
7. Clic derecho sobre el botón → **Modificar texto**
8. Escribe: **Generar Reporte Plazas Vacantes**
9. Clic fuera del botón

---

## Paso 9: Guardar

1. **Ctrl + S**
2. Confirmar formato **.xlsm**

---

## Paso 10: Probar

1. Verifica que la celda **B2** de la hoja "Configuracion" tiene la ruta correcta del archivo fuente
2. Haz clic en el botón **"Generar Reporte Plazas Vacantes"**
3. La macro:
   - Abre el archivo de Vacantes Definitivas en modo solo lectura
   - Lee y filtra los datos
   - Crea dos hojas nuevas en tu archivo: "Plazas Vacantes" y "Resumen"
   - Cierra el archivo fuente sin modificarlo
   - Te muestra un mensaje con el resultado
4. Verifica las hojas creadas

---

## Paso 11: Subir a SharePoint

1. Sube el archivo **"Reporte Plazas Vacantes.xlsm"** a SharePoint
2. Asegúrate de que la ruta en B2 siga siendo válida desde SharePoint
   - Si ambos archivos están en el mismo SharePoint sincronizado, la ruta local funciona
   - Si no, usa la ruta de SharePoint (Opción B del Paso 3)

---

## Uso diario

1. Abrir **"Reporte Plazas Vacantes.xlsm"** en Excel de escritorio
2. Si aparece aviso de macros → **"Habilitar contenido"**
3. Clic en el botón **"Generar Reporte Plazas Vacantes"**
4. Esperar unos segundos
5. Los resultados aparecen en las hojas "Plazas Vacantes" y "Resumen"
6. Cada vez que presiones el botón, los datos se actualizan (reemplaza los anteriores)

---

## Ventajas de este enfoque

| Ventaja | Descripción |
|---------|-------------|
| **No toca el archivo original** | El archivo de Vacantes Definitivas nunca se modifica |
| **Siempre actualizado** | Cada vez que presionas el botón, lee los datos más recientes |
| **Independiente** | Es un archivo aparte, puede estar en otra carpeta/ubicación |
| **Seguro** | Modo solo lectura, no puede dañar la fuente |
| **Reutilizable** | Funciona mientras el archivo fuente exista en la ruta configurada |

---

## Solución de problemas

| Problema | Solución |
|----------|----------|
| "No se ha configurado la ruta del archivo fuente" | Ve a la hoja "Configuracion" y escribe la ruta en la celda B2 |
| "No se pudo abrir el archivo fuente" | Verifica que la ruta en B2 es correcta y que tienes acceso al archivo |
| No aparece la pestaña "Programador" | Archivo → Opciones → Personalizar cinta → marca "Programador" |
| "Las macros han sido deshabilitadas" | Clic en "Habilitar contenido" |
| El botón no hace nada | Clic derecho → Asignar macro → selecciona "GenerarReportePlazasVacantes" |
| "No se encontró la columna PLAZA" | Verifica que el archivo fuente tiene encabezados en la fila 1 |
| Tarda mucho | Es normal con muchos datos. Espera sin tocar nada |
| Los datos están desactualizados | Presiona el botón de nuevo para actualizar con los datos más recientes |

---

## Estructura final del archivo

Tu archivo "Reporte Plazas Vacantes.xlsm" tendrá:

| Hoja | Contenido |
|------|-----------|
| **Configuracion** | Ruta del archivo fuente + botón |
| **Plazas Vacantes** | (se crea/actualiza al presionar el botón) Listado de plazas vacantes |
| **Resumen** | (se crea/actualiza al presionar el botón) Totales y desglose por subregión |
