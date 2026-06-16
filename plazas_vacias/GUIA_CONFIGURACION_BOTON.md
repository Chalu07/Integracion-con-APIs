# Guía: Configurar el botón en "Vacantes Definitivas"

## Paso a paso para agregar la macro y el botón al archivo Excel en SharePoint

---

### Paso 1: Descargar el archivo para editarlo

1. Ir al archivo **"Registro de vacantes DEFINITIVAS 2026"** en SharePoint
2. Clic en **"Abrir en aplicación de escritorio"** (NO en la versión web)
   - La versión web de Excel no soporta macros VBA

---

### Paso 2: Guardar como .xlsm (habilitar macros)

1. Ir a **Archivo → Guardar como**
2. Cambiar el tipo a: **Libro de Excel habilitado para macros (.xlsm)**
3. Guardar en la misma ubicación de SharePoint (o en OneDrive)

> **Importante:** El archivo debe ser `.xlsm` para que las macros funcionen.
> El archivo `.xlsx` original puede mantenerse como respaldo.

---

### Paso 3: Abrir el Editor de Visual Basic

1. Presionar **Alt + F11** (esto abre el editor de VBA)
2. En el panel izquierdo, buscar el nombre del libro (ej: "VBAProject (Registro de vacantes...)")

---

### Paso 4: Importar el módulo de la macro

**Opción A: Importar archivo .bas**
1. En el editor de VBA: **Archivo → Importar archivo...**
2. Seleccionar el archivo `MacroVacantesDefinitivas.bas`
3. Se creará un módulo llamado "ModuloVacantes"

**Opción B: Copiar y pegar el código**
1. En el editor de VBA, clic derecho sobre el nombre del proyecto → **Insertar → Módulo**
2. Se abre una ventana en blanco a la derecha
3. Abrir el archivo `MacroVacantesDefinitivas.bas` con el Bloc de notas
4. Copiar TODO el contenido (Ctrl+A, Ctrl+C)
5. Pegar en la ventana del módulo en VBA (Ctrl+V)

---

### Paso 5: Cerrar el editor de VBA

1. Presionar **Alt + F11** de nuevo (o cerrar la ventana del editor)
2. Regresa a la hoja de Excel

---

### Paso 6: Agregar el botón a la hoja

1. Ir a la pestaña **Programador** (Developer)
   - Si no aparece esta pestaña: **Archivo → Opciones → Personalizar cinta de opciones → marcar "Programador"**
2. Clic en **Insertar** (en la sección Controles)
3. Seleccionar **Botón (control de formulario)** — es el primer icono, un rectángulo
4. Dibujar el botón en la hoja arrastrando con el mouse (en un área libre, por ejemplo arriba de los datos o en una esquina)
5. Aparece una ventana "Asignar macro"
6. Seleccionar **GenerarReportePlazasVacantes**
7. Clic en **Aceptar**
8. El botón se crea con el texto "Botón 1" — clic derecho sobre él → **Modificar texto**
9. Escribir: **"Generar Reporte Plazas Vacantes"**
10. Clic fuera del botón para deseleccionarlo

---

### Paso 7: Guardar el archivo

1. Presionar **Ctrl + S**
2. Si pregunta por el formato, confirmar que quiere guardar como **.xlsm**

---

### Paso 8: Probar el botón

1. Clic en el botón **"Generar Reporte Plazas Vacantes"**
2. Esperar unos segundos (depende de la cantidad de datos)
3. Se abrirá un **nuevo libro de Excel** con dos hojas:
   - **"Plazas Vacantes"**: todas las plazas que aún están vacías
   - **"Resumen"**: métricas y totales por subregión
4. Aparecerá un mensaje indicando cuántas plazas vacantes se encontraron
5. **Guardar el nuevo archivo** con el nombre y ubicación que desee

---

### Paso 9: Subir a SharePoint (si editó localmente)

1. Si descargó el archivo para editarlo, subirlo de vuelta a SharePoint
2. El archivo `.xlsm` quedará listo para que cualquier persona lo use

---

## Uso diario

Cada vez que necesite el reporte actualizado:

1. Abrir el archivo **"Vacantes Definitivas"** en Excel de escritorio
2. Clic en el botón **"Generar Reporte Plazas Vacantes"**
3. Se genera automáticamente un nuevo Excel con las plazas aún vacantes
4. Guardar el reporte donde lo necesite

---

## Notas importantes

- **El archivo debe abrirse en Excel de escritorio**, no en la versión web
  (Excel Online no soporta macros VBA).
- Si al abrir aparece una barra amarilla diciendo "Las macros han sido
  deshabilitadas", clic en **"Habilitar contenido"**.
- El botón funciona con los datos que estén en el archivo al momento de
  hacer clic. Si el formulario agregó nuevos datos, el botón los incluirá
  automáticamente.
- **Criterio de filtrado**: se consideran vacantes las plazas donde
  "Estado del nombramiento" NO es "FIRMADO". Esto incluye plazas con
  estado "Seleccionar", "EN PROCESO", "EN FIRMAS" o sin estado asignado.
- Solo se incluyen las 9 subregiones oficiales de Antioquia.

## Solución de problemas

| Problema | Solución |
|----------|----------|
| No aparece la pestaña "Programador" | Archivo → Opciones → Personalizar cinta → marcar "Programador" |
| "Las macros han sido deshabilitadas" | Clic en "Habilitar contenido" en la barra amarilla |
| "No se encontró la columna PLAZA" | Verificar que la fila 1 tenga los encabezados correctos |
| El botón no hace nada | Clic derecho → Asignar macro → seleccionar "GenerarReportePlazasVacantes" |
| Error al guardar como .xlsm | Verificar que no haya otro archivo con el mismo nombre abierto |
