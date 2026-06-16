# Guía de Configuración — Microsoft 365 (OneDrive / SharePoint)

## Estructura recomendada en la nube

```
📁 OneDrive o SharePoint
 └── 📁 Planta Docente
      ├── 📄 archivo_fuente_MEN.xlsx     ← Se reemplaza cada vez que se descarga
      └── 📄 Excel_Maestro_Plazas_Vacias.xlsx   ← Reporte que se actualiza
```

> **Importante:** El archivo fuente debe tener **siempre el mismo nombre** en
> la misma carpeta para que Power Query lo encuentre automáticamente.

---

## Opción A: Actualización con script Python (más confiable)

### Primera vez

```bash
# Instalar dependencias (solo una vez)
pip install -r requirements.txt

# Generar el Excel maestro
python generar_excel_maestro.py archivo_descargado.xlsx Excel_Maestro.xlsx
```

### Cada vez que se descargue un nuevo archivo

```bash
python generar_excel_maestro.py nuevo_archivo.xlsx Excel_Maestro.xlsx
```

Subir `Excel_Maestro.xlsx` a OneDrive/SharePoint. Listo.

---

## Opción B: Actualización automática con Power Query (sin Python)

Esta opción permite que el Excel maestro se actualice con solo presionar
**Ctrl+Alt+F5** (Actualizar todo), sin necesidad de ejecutar scripts.

### Paso 1: Preparar los archivos

1. Subir el archivo fuente del MEN a OneDrive/SharePoint
2. Abrir el Excel maestro **en Excel de escritorio** (no en la versión web)

### Paso 2: Crear la consulta principal

1. Ir a **Datos → Obtener datos → Desde otras fuentes → Consulta en blanco**
2. En la barra de fórmulas o en **Editor avanzado**, pegar el contenido del
   archivo `powerquery_plazas_vacias.pq`
3. **MUY IMPORTANTE:** Cambiar la línea `RutaArchivo` por la ruta real:
   - Si el archivo está en OneDrive personal:
     ```
     "C:\Users\TU_USUARIO\OneDrive\Planta Docente\archivo_fuente.xlsx"
     ```
   - Si está en OneDrive corporativo:
     ```
     "C:\Users\TU_USUARIO\OneDrive - NOMBRE_EMPRESA\Planta Docente\archivo_fuente.xlsx"
     ```
   - Si está en SharePoint (sincronizado):
     ```
     "C:\Users\TU_USUARIO\NOMBRE_EMPRESA\Nombre del sitio - Documentos\Planta Docente\archivo_fuente.xlsx"
     ```

> **Tip:** Para encontrar la ruta exacta, navegar al archivo en el Explorador
> de Windows, clic derecho → **Copiar como ruta de acceso**.

4. Renombrar la consulta como **PlazasVacias**
5. Clic en **Cerrar y cargar**

### Paso 3: Crear la consulta de resumen (opcional)

1. Repetir el proceso con el archivo `powerquery_resumen.pq`
2. Renombrar como **ResumenSubregion**
3. Clic en **Cerrar y cargar**

### Paso 4: Configurar actualización automática (opcional)

1. Ir a **Datos → Propiedades de consulta** (o clic derecho en la consulta)
2. Marcar **Actualizar cada X minutos** si se desea actualización periódica
3. Marcar **Actualizar datos al abrir el archivo** para que se actualice
   cada vez que se abra el Excel

### Actualización diaria

1. Descargar el nuevo archivo del MEN
2. **Reemplazar** el archivo anterior en OneDrive/SharePoint
   (usar exactamente el **mismo nombre**)
3. Abrir el Excel maestro → **Datos → Actualizar todo** (o Ctrl+Alt+F5)
4. Los datos se actualizan automáticamente

---

## Solución de problemas

| Problema | Solución |
|----------|----------|
| "No se puede encontrar el archivo" | Verificar que la ruta en `RutaArchivo` sea correcta y que el archivo exista |
| Los datos no se actualizan | Asegurarse de que el archivo fuente tenga el mismo nombre que el configurado |
| Error de permisos | Abrir en Excel de escritorio (no web) y autorizar el acceso a OneDrive |
| Power Query no aparece | Verificar que se usa Excel 2016+ o Microsoft 365 |
| Datos incorrectos después de actualizar | El archivo fuente puede tener una estructura diferente; verificar que la hoja "T" exista |

---

## Notas adicionales

- El Excel maestro incluye hojas separadas para **Vacantes Definitivas** y
  **Vacantes Temporales** para facilitar el filtrado.
- La hoja **Dashboard** muestra un resumen ejecutivo con métricas por
  subregión y tipo de vacante.
- Los **filtros automáticos** están habilitados en todas las hojas de datos
  para facilitar la búsqueda.
