# Procesador de Planta Docente — Plazas Vacías

Herramienta para analizar el archivo Excel descargado de la plataforma del
Ministerio de Educación e identificar automáticamente las plazas docentes
vacías en el departamento de Antioquia.

## Requisitos

- Python 3.10+
- Dependencias: `pip install -r requirements.txt`

## Uso

### Reporte rápido
```bash
python procesar_planta.py archivo_descargado.xlsx
```

### Excel Maestro (con Dashboard, hojas separadas por tipo, análisis por I.E.)
```bash
python generar_excel_maestro.py archivo_descargado.xlsx Excel_Maestro.xlsx
```

### Con Power Query en Microsoft 365 (actualización automática)
Ver la guía completa en [`GUIA_MICROSOFT_365.md`](GUIA_MICROSOFT_365.md).

## ¿Qué hace?

1. **Lee** la hoja **"T"** del archivo Excel de planta docente.
2. **Filtra** únicamente las 9 subregiones oficiales de Antioquia:
   - Bajo Cauca, Magdalena Medio, Nordeste, Norte, Occidente,
     Oriente, Suroeste, Urabá, Valle de Aburrá.
3. **Clasifica** las plazas en dos tipos de vacante:
   - **Vacante Definitiva**: la plaza no tiene titular en propiedad.
   - **Vacante Temporal**: la plaza tiene titular pero está temporalmente
     sin ocupar.
4. **Genera** un archivo Excel con las siguientes hojas:
   - `Plazas Vacías`: listado completo con filtros automáticos.
   - `Resumen`: totales, desglose y criterios de identificación.
   - `Por Subregión`: vacantes agrupadas por subregión.
   - `Por Municipio`: vacantes agrupadas por municipio.
   - `Duplicados NUM_PLAZA`: registros con número de plaza duplicado.

## Criterios de identificación

### Vacante Definitiva
- `SITUACIONLABORAL = "Encargo Vacante Definitiva"`: la plaza está
  definitivamente vacía; un docente la cubre temporalmente por encargo.
- Empleado con `RETIRO_FECHA` anterior a hoy, estado no es "Reemplazo",
  `NUM_PLAZA ≠ 0`, y no existe otro empleado activo en la misma plaza.

### Vacante Temporal
- `SITUACIONLABORAL` = "Encargo Vacante Temporal" o "Reemplazo Vacante
  Temporal".
- `NOVEDAD_PLANTA = "Vacancia Temporal"`.
- Titular ausente por: Comisión, Licencia No Remunerada, Permiso Sindical,
  Periodo de Prueba en otra entidad, Sanción, Situación Laboral
  Remunerada (Tutor PTA), entre otros.

## Archivos del proyecto

| Archivo | Descripción |
|---------|-------------|
| `procesar_planta.py` | Script de reporte rápido |
| `generar_excel_maestro.py` | Generador del Excel maestro completo |
| `powerquery_plazas_vacias.pq` | Código Power Query para Excel 365 |
| `powerquery_resumen.pq` | Código Power Query para resumen por subregión |
| `GUIA_MICROSOFT_365.md` | Guía paso a paso para configurar en Microsoft 365 |
| `requirements.txt` | Dependencias Python |

## Flujo de actualización

```
Plataforma MEN → Descargar Excel → Ejecutar script → Excel Maestro → OneDrive/SharePoint
```

Cada vez que se descargue un nuevo archivo de la plataforma, ejecutar
`generar_excel_maestro.py` y subir el resultado a la nube corporativa.

Alternativamente, configurar Power Query en el Excel maestro para que se
actualice automáticamente al reemplazar el archivo fuente en OneDrive
(ver `GUIA_MICROSOFT_365.md`).
