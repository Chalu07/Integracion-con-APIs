#!/usr/bin/env python3
"""
Generador del Excel Maestro — Plazas Vacías Docentes
=====================================================
Genera un archivo Excel "maestro" con los datos procesados del archivo
fuente del MEN, más hojas de resumen, tablas dinámicas por subregión
y municipio, y una hoja de instrucciones para configurar Power Query
en Microsoft 365.

Uso:
    python generar_excel_maestro.py <archivo_fuente.xlsx> [excel_maestro.xlsx]
"""

import sys
import os
from datetime import datetime
from pathlib import Path

import pandas as pd

# Importar el procesador principal
sys.path.insert(0, os.path.dirname(__file__))
from procesar_planta import (
    cargar_datos,
    filtrar_subregiones,
    clasificar_vacantes,
    detectar_duplicados,
    COLUMNAS_REPORTE,
    SUBREGIONES_OFICIALES,
)


def crear_hoja_instrucciones(writer: pd.ExcelWriter) -> None:
    """Crea la hoja de instrucciones para configurar Power Query."""
    instrucciones = [
        ("", ""),
        ("EXCEL MAESTRO — PLAZAS VACÍAS DOCENTES", ""),
        ("Departamento de Antioquia", ""),
        ("", ""),
        ("═══ CÓMO ACTUALIZAR ESTE ARCHIVO ═══", ""),
        ("", ""),
        ("OPCIÓN 1: Usar el script Python (recomendado)", ""),
        ("", ""),
        ("Paso 1:", "Descargar el nuevo archivo del MEN"),
        ("Paso 2:", "Ejecutar: python generar_excel_maestro.py <archivo.xlsx>"),
        ("Paso 3:", "Subir el Excel maestro generado a OneDrive/SharePoint"),
        ("", ""),
        ("OPCIÓN 2: Configurar Power Query en Excel (actualización automática)", ""),
        ("", ""),
        ("Paso 1:", "Abrir este archivo en Excel (escritorio, no web)"),
        ("Paso 2:", "Ir a Datos → Obtener datos → Desde otras fuentes → Consulta en blanco"),
        ("Paso 3:", "Clic en 'Editor avanzado'"),
        ("Paso 4:", "Pegar el contenido del archivo 'powerquery_plazas_vacias.pq'"),
        ("Paso 5:", "Cambiar la variable RutaArchivo por la ruta del archivo fuente en OneDrive"),
        ("Paso 6:", "Clic en 'Listo' → 'Cerrar y cargar'"),
        ("Paso 7:", "Repetir pasos 2-6 con 'powerquery_resumen.pq' para el resumen"),
        ("", ""),
        ("Una vez configurado, para actualizar:", ""),
        ("", "1. Descargar el nuevo archivo del MEN con el MISMO nombre"),
        ("", "2. Reemplazar el archivo anterior en OneDrive/SharePoint"),
        ("", "3. Abrir el Excel maestro → Datos → Actualizar todo (Ctrl+Alt+F5)"),
        ("", ""),
        ("═══ SUBREGIONES OFICIALES INCLUIDAS ═══", ""),
        ("", ""),
    ]
    for sub in sorted(SUBREGIONES_OFICIALES):
        instrucciones.append(("", sub))

    instrucciones.extend([
        ("", ""),
        ("═══ CRITERIOS DE CLASIFICACIÓN ═══", ""),
        ("", ""),
        ("Vacante Definitiva:", "La plaza NO tiene titular en propiedad."),
        ("  Indicadores:", "SITUACIONLABORAL = 'Encargo Vacante Definitiva'"),
        ("", "O empleado retirado cuya plaza no fue reasignada."),
        ("", ""),
        ("Vacante Temporal:", "La plaza TIENE titular pero está temporalmente sin ocupar."),
        ("  Indicadores:", "SITUACIONLABORAL = 'Encargo Vacante Temporal' o 'Reemplazo Vacante Temporal'"),
        ("", "NOVEDAD_PLANTA = 'Vacancia Temporal'"),
        ("", "Titular ausente: Comisión, Licencia, Permiso Sindical, Tutor PTA, etc."),
    ])

    df_instr = pd.DataFrame(instrucciones, columns=["", " "])
    df_instr.to_excel(writer, sheet_name="Instrucciones", index=False)

    ws = writer.sheets["Instrucciones"]
    ws.set_column("A:A", 50)
    ws.set_column("B:B", 70)


def crear_hoja_dashboard(writer: pd.ExcelWriter,
                         vacantes: pd.DataFrame,
                         df_filtrado: pd.DataFrame) -> None:
    """Crea la hoja de dashboard con métricas clave."""
    ahora = datetime.now().strftime("%Y-%m-%d %H:%M")
    n_def = (vacantes["TIPO_VACANTE"] == "Vacante Definitiva").sum()
    n_temp = (vacantes["TIPO_VACANTE"] == "Vacante Temporal").sum()

    metricas = [
        ("DASHBOARD — PLAZAS VACÍAS", ""),
        (f"Última actualización: {ahora}", ""),
        ("", ""),
        ("MÉTRICAS GENERALES", ""),
        ("Total registros en archivo fuente (subregiones oficiales)",
         len(df_filtrado)),
        ("Total plazas vacías identificadas", len(vacantes)),
        ("", ""),
        ("POR TIPO DE VACANTE", ""),
        ("Vacante Definitiva", n_def),
        ("Vacante Temporal", n_temp),
        ("", ""),
        ("POR SUBREGIÓN", ""),
    ]

    for sub in sorted(SUBREGIONES_OFICIALES):
        sub_data = vacantes[vacantes["SUBREGION"] == sub]
        n_sub = len(sub_data)
        n_sub_def = (sub_data["TIPO_VACANTE"] == "Vacante Definitiva").sum()
        n_sub_temp = (sub_data["TIPO_VACANTE"] == "Vacante Temporal").sum()
        metricas.append((sub, f"{n_sub} (Def: {n_sub_def} | Temp: {n_sub_temp})"))

    metricas.extend([
        ("", ""),
        ("POR CARGO (Top 10)", ""),
    ])
    top_cargos = vacantes["CARGO"].value_counts().head(10)
    for cargo, cnt in top_cargos.items():
        metricas.append((cargo, cnt))

    df_met = pd.DataFrame(metricas, columns=["Concepto", "Valor"])
    df_met.to_excel(writer, sheet_name="Dashboard", index=False)

    ws = writer.sheets["Dashboard"]
    ws.set_column("A:A", 55)
    ws.set_column("B:B", 40)


def main() -> None:
    if len(sys.argv) < 2:
        print("Uso: python generar_excel_maestro.py <archivo_fuente.xlsx> "
              "[excel_maestro.xlsx]")
        sys.exit(1)

    ruta_entrada = sys.argv[1]
    if not os.path.isfile(ruta_entrada):
        print(f"Error: No se encontró el archivo '{ruta_entrada}'")
        sys.exit(1)

    if len(sys.argv) >= 3:
        ruta_salida = sys.argv[2]
    else:
        fecha = datetime.now().strftime("%Y-%m-%d")
        ruta_salida = f"Excel_Maestro_Plazas_Vacias_{fecha}.xlsx"

    print("=" * 60)
    print("  GENERADOR DE EXCEL MAESTRO — PLAZAS VACÍAS")
    print("=" * 60)

    print("\n[1/6] Cargando datos...")
    df = cargar_datos(ruta_entrada)

    print("\n[2/6] Filtrando subregiones oficiales...")
    df_filtrado = filtrar_subregiones(df)

    print("\n[3/6] Clasificando plazas vacías...")
    vacantes = clasificar_vacantes(df_filtrado)

    print("\n[4/6] Detectando duplicados...")
    duplicados = detectar_duplicados(df_filtrado)
    print(f"  NUM_PLAZA duplicados: {duplicados['NUM_PLAZA'].nunique():,}")

    print("\n[5/6] Construyendo tablas de análisis...")
    # Tabla cruzada por subregión
    por_subregion = (
        vacantes.groupby(["SUBREGION", "TIPO_VACANTE"])
        .size()
        .unstack(fill_value=0)
    )
    por_subregion["TOTAL"] = por_subregion.sum(axis=1)
    por_subregion = por_subregion.sort_values("TOTAL", ascending=False)

    # Tabla cruzada por municipio
    por_municipio = (
        vacantes.groupby(["MUNICIPIO", "TIPO_VACANTE"])
        .size()
        .unstack(fill_value=0)
    )
    por_municipio["TOTAL"] = por_municipio.sum(axis=1)
    por_municipio = por_municipio.sort_values("TOTAL", ascending=False)

    # Tabla cruzada por I.E.
    por_ie = (
        vacantes.groupby(["I.E", "MUNICIPIO", "SUBREGION", "TIPO_VACANTE"])
        .size()
        .unstack(fill_value=0)
    )
    por_ie["TOTAL"] = por_ie.sum(axis=1)
    por_ie = por_ie.sort_values("TOTAL", ascending=False)

    # Vacantes Definitivas
    v_definitivas = vacantes[vacantes["TIPO_VACANTE"] == "Vacante Definitiva"]
    # Vacantes Temporales
    v_temporales = vacantes[vacantes["TIPO_VACANTE"] == "Vacante Temporal"]

    print("\n[6/6] Generando Excel maestro...")
    columnas = [c for c in COLUMNAS_REPORTE if c in vacantes.columns]

    with pd.ExcelWriter(ruta_salida, engine="xlsxwriter") as writer:
        wb = writer.book

        # Formatos
        header_fmt = wb.add_format({
            "bold": True,
            "text_wrap": True,
            "valign": "top",
            "fg_color": "#4472C4",
            "font_color": "white",
            "border": 1,
        })
        title_fmt = wb.add_format({
            "bold": True,
            "font_size": 14,
            "fg_color": "#2F5496",
            "font_color": "white",
        })
        subtitle_fmt = wb.add_format({
            "bold": True,
            "font_size": 11,
            "fg_color": "#D6E4F0",
        })

        # Hoja 1: Instrucciones
        crear_hoja_instrucciones(writer)

        # Hoja 2: Dashboard
        crear_hoja_dashboard(writer, vacantes, df_filtrado)

        # Hoja 3: Todas las plazas vacías
        vacantes[columnas].to_excel(
            writer, sheet_name="Todas las Vacantes", index=False
        )
        ws = writer.sheets["Todas las Vacantes"]
        for col_num, col_name in enumerate(columnas):
            ws.write(0, col_num, col_name, header_fmt)
        ws.autofilter(0, 0, len(vacantes), len(columnas) - 1)
        ws.set_column("A:Z", 18)

        # Hoja 4: Solo Vacantes Definitivas
        v_definitivas[columnas].to_excel(
            writer, sheet_name="Vacantes Definitivas", index=False
        )
        ws = writer.sheets["Vacantes Definitivas"]
        for col_num, col_name in enumerate(columnas):
            ws.write(0, col_num, col_name, header_fmt)
        ws.autofilter(0, 0, len(v_definitivas), len(columnas) - 1)
        ws.set_column("A:Z", 18)

        # Hoja 5: Solo Vacantes Temporales
        v_temporales[columnas].to_excel(
            writer, sheet_name="Vacantes Temporales", index=False
        )
        ws = writer.sheets["Vacantes Temporales"]
        for col_num, col_name in enumerate(columnas):
            ws.write(0, col_num, col_name, header_fmt)
        ws.autofilter(0, 0, len(v_temporales), len(columnas) - 1)
        ws.set_column("A:Z", 18)

        # Hoja 6: Por Subregión
        por_subregion.to_excel(writer, sheet_name="Por Subregión")
        writer.sheets["Por Subregión"].set_column("A:E", 22)

        # Hoja 7: Por Municipio
        por_municipio.to_excel(writer, sheet_name="Por Municipio")
        writer.sheets["Por Municipio"].set_column("A:E", 22)

        # Hoja 8: Por I.E.
        por_ie.to_excel(writer, sheet_name="Por I.E.")
        writer.sheets["Por I.E."].set_column("A:F", 22)

        # Hoja 9: Duplicados
        if len(duplicados) > 0:
            duplicados.to_excel(
                writer, sheet_name="Duplicados NUM_PLAZA", index=False
            )
            writer.sheets["Duplicados NUM_PLAZA"].set_column("A:I", 18)

    print(f"\n  Archivo generado: {ruta_salida}")
    print(f"\n  Hojas incluidas:")
    print(f"    1. Instrucciones — Guía de uso y configuración")
    print(f"    2. Dashboard — Métricas clave")
    print(f"    3. Todas las Vacantes — {len(vacantes):,} registros")
    print(f"    4. Vacantes Definitivas — {len(v_definitivas):,} registros")
    print(f"    5. Vacantes Temporales — {len(v_temporales):,} registros")
    print(f"    6. Por Subregión")
    print(f"    7. Por Municipio")
    print(f"    8. Por I.E.")
    print(f"    9. Duplicados NUM_PLAZA — {len(duplicados):,} registros")

    print("\n" + "=" * 60)
    print(f"  EXCEL MAESTRO GENERADO EXITOSAMENTE")
    print("=" * 60)


if __name__ == "__main__":
    main()
