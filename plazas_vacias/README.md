# Procesador de Planta Docente — Plazas Vacías

Herramienta para analizar el archivo Excel descargado de la plataforma del
Ministerio de Educación e identificar automáticamente las plazas docentes
vacías en el departamento de Antioquia.

## Requisitos

- Python 3.10+
- Dependencias: `pip install -r requirements.txt`

## Uso

```bash
# Generar reporte con nombre automático (Reporte_Plazas_Vacias_YYYY-MM-DD.xlsx)
python procesar_planta.py archivo_descargado.xlsx

# Especificar nombre del archivo de salida
python procesar_planta.py archivo_descargado.xlsx reporte_salida.xlsx
```

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

## Flujo de actualización

```
Plataforma MEN → Descargar Excel → Ejecutar script → Reporte actualizado
```

Cada vez que se descargue un nuevo archivo de la plataforma, ejecutar el
script con el archivo como argumento para obtener el reporte actualizado.
