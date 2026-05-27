#!/usr/bin/env python3
"""
Procesador de Planta Docente — Plazas Vacías
=============================================
Lee el archivo Excel descargado del Ministerio de Educación (hoja "T"),
identifica las plazas vacías (definitivas y temporales) filtrando
únicamente por las 9 subregiones oficiales del departamento de Antioquia,
y genera un reporte Excel limpio.

Uso:
    python procesar_planta.py <archivo_entrada.xlsx> [archivo_salida.xlsx]

Si no se indica archivo de salida, se genera automáticamente con el
nombre: Reporte_Plazas_Vacias_<fecha>.xlsx
"""

import sys
import os
from datetime import datetime
from pathlib import Path

import pandas as pd

# ════════════════════════════════════════════════════════════════
# CONFIGURACIÓN
# ════════════════════════════════════════════════════════════════

HOJA_ORIGEN = "T"

SUBREGIONES_OFICIALES = [
    "BAJO CAUCA",
    "MAGDALENA MEDIO",
    "NORDESTE",
    "NORTE",
    "OCCIDENTE",
    "ORIENTE",
    "SUROESTE",
    "URABA",
    "VALLE DE ABURRA",
]

# Estados laborales que indican vacante definitiva:
# La plaza NO tiene titular en propiedad; se cubre por encargo temporal.
VACANTE_DEFINITIVA_ESTADOS = [
    "Encargo Vacante Definitiva",
]

# Estados laborales que indican vacante temporal:
# La plaza tiene titular, pero está temporalmente sin ocupar
# (el titular está en comisión, licencia, permiso, etc.).
VACANTE_TEMPORAL_ESTADOS = [
    "Encargo Vacante Temporal",
    "Reemplazo Vacante Temporal",
]

# Estados de ausencia del titular (la plaza necesita cobertura temporal).
TITULAR_AUSENTE_ESTADOS = [
    "Comision",
    "Comisión No Remunerada",
    "Comision Libre Nom. y Remocion",
    "Licencia No Remunerada",
    "Permiso Sindical",
    "Permiso Remunerado",
    "Periodo Prueba en otra entidad",
    "Periodo Prueba misma entidad",
    "Sancion",
    "Accion disciplinaria- sus",
    "Incapacidad Sup 180 Días",
    "Recuperacion Capacidad Laboral",
    "Situacion Laboral Remunerada",
]

# Columnas del reporte de salida
COLUMNAS_REPORTE = [
    "TIPO_VACANTE",
    "NUM_PLAZA",
    "MUNICIPIO",
    "SUBREGION",
    "I.E",
    "SEDE",
    "AREA",
    "CARGO",
    "GRADO",
    "CODSEDE_AREA",
    "DANE_ESTABLECIMIENTO",
    "DANE_SEDE",
    "SITUACIONLABORAL",
    "NOVEDAD_PLANTA",
    "NIVELDICTA",
    "AREAEDUCATIVA",
    "NIVELCONTRATACION",
    "ESCALAFON",
    "CODEMPLEADO",
    "EMPLEADO",
    "RETIRO_FECHA",
    "ENCARGO",
    "ENCARGO_GRADO",
    "REEMPLAZO_EMPLEADO",
    "IE_SEDE_AREA",
    "AREA.1",
]


# ════════════════════════════════════════════════════════════════
# FUNCIONES
# ════════════════════════════════════════════════════════════════


def cargar_datos(ruta: str) -> pd.DataFrame:
    """Lee la hoja T del archivo Excel de planta docente."""
    print(f"  Leyendo archivo: {ruta}")
    df = pd.read_excel(ruta, sheet_name=HOJA_ORIGEN)
    print(f"  Registros cargados: {len(df):,}")
    return df


def filtrar_subregiones(df: pd.DataFrame) -> pd.DataFrame:
    """Conserva solo las filas de las 9 subregiones oficiales de Antioquia."""
    mascara = df["SUBREGION"].isin(SUBREGIONES_OFICIALES)
    resultado = df[mascara].copy()
    descartados = len(df) - len(resultado)
    print(f"  Registros en subregiones oficiales: {len(resultado):,}  "
          f"(descartados: {descartados:,})")
    return resultado


def _plaza_sin_otro_activo(row: pd.Series, df: pd.DataFrame,
                            ahora: pd.Timestamp) -> bool:
    """Verifica que no exista otro empleado activo en la misma NUM_PLAZA."""
    if row["NUM_PLAZA"] == 0:
        return False
    otros = df[
        (df["NUM_PLAZA"] == row["NUM_PLAZA"])
        & (df.index != row.name)
        & ((df["RETIRO_FECHA"].isna()) | (df["RETIRO_FECHA"] >= ahora))
    ]
    return len(otros) == 0


def clasificar_vacantes(df: pd.DataFrame) -> pd.DataFrame:
    """
    Clasifica cada registro en un tipo de vacante o lo descarta.

    Categorías:
        - Vacante Definitiva: plaza sin titular en propiedad.
        - Vacante Temporal: plaza con titular pero temporalmente vacía.
        - Vacante por Retiro: titular se retiró y la plaza no fue reasignada.
        - Titular Ausente: el titular existe pero está en comisión/licencia/etc.

    Retorna solo las filas clasificadas como vacantes.
    """
    ahora = pd.Timestamp.now()
    df = df.copy()
    df["TIPO_VACANTE"] = ""

    # 1. Vacante Definitiva
    m1 = df["SITUACIONLABORAL"].isin(VACANTE_DEFINITIVA_ESTADOS)
    df.loc[m1, "TIPO_VACANTE"] = "Vacante Definitiva"

    # 2. Vacante Temporal
    m2 = df["SITUACIONLABORAL"].isin(VACANTE_TEMPORAL_ESTADOS)
    df.loc[m2 & (df["TIPO_VACANTE"] == ""), "TIPO_VACANTE"] = "Vacante Temporal"

    # 3. Vacancia Temporal por novedad de planta
    m3 = df["NOVEDAD_PLANTA"] == "Vacancia Temporal"
    df.loc[m3 & (df["TIPO_VACANTE"] == ""), "TIPO_VACANTE"] = "Vacante Temporal"

    # 4. Vacante por retiro (empleado se fue, plaza sin reasignar)
    es_separado = (df["RETIRO_FECHA"].notna()) & (df["RETIRO_FECHA"] < ahora)
    no_es_reemplazo = ~df["SITUACIONLABORAL"].str.contains(
        "Reemplazo", case=False, na=False
    )
    tiene_plaza = df["NUM_PLAZA"] != 0
    candidatos = df[es_separado & no_es_reemplazo & tiene_plaza]
    plaza_vacia = candidatos.apply(
        _plaza_sin_otro_activo, axis=1, df=df, ahora=ahora
    )
    indices_retiro = candidatos[plaza_vacia].index
    m4 = df.index.isin(indices_retiro) & (df["TIPO_VACANTE"] == "")
    df.loc[m4, "TIPO_VACANTE"] = "Vacante Definitiva"

    # 5. Titular ausente (comisión, licencia, permiso, etc.)
    activo = (df["RETIRO_FECHA"].isna()) | (df["RETIRO_FECHA"] >= ahora)
    m5 = df["SITUACIONLABORAL"].isin(TITULAR_AUSENTE_ESTADOS) & activo
    df.loc[m5 & (df["TIPO_VACANTE"] == ""), "TIPO_VACANTE"] = "Vacante Temporal"

    vacantes = df[df["TIPO_VACANTE"] != ""]
    print(f"  Plazas vacías identificadas: {len(vacantes):,}")
    for tipo, cnt in vacantes["TIPO_VACANTE"].value_counts().items():
        print(f"    {tipo}: {cnt:,}")
    return vacantes


def detectar_duplicados(df_todo: pd.DataFrame) -> pd.DataFrame:
    """Identifica registros con NUM_PLAZA duplicado (excluyendo 0)."""
    dup = df_todo[
        (df_todo.duplicated(subset="NUM_PLAZA", keep=False))
        & (df_todo["NUM_PLAZA"] != 0)
    ]
    cols = [
        "NUM_PLAZA", "EMPLEADO", "SITUACIONLABORAL", "NOVEDAD_PLANTA",
        "RETIRO_FECHA", "MUNICIPIO", "I.E", "SEDE", "CARGO",
    ]
    return dup[cols].sort_values("NUM_PLAZA")


def generar_resumen(df_todo: pd.DataFrame,
                    vacantes: pd.DataFrame) -> pd.DataFrame:
    """Construye la hoja de resumen del reporte."""
    n_def = (vacantes["TIPO_VACANTE"] == "Vacante Definitiva").sum()
    n_temp = (vacantes["TIPO_VACANTE"] == "Vacante Temporal").sum()

    filas = [
        ("Total de plazas en hoja T (subregiones oficiales)", len(df_todo)),
        ("Total de plazas vacías identificadas", len(vacantes)),
        ("", ""),
        ("--- DESGLOSE POR TIPO ---", ""),
        ("Vacante Definitiva", n_def),
        ("Vacante Temporal", n_temp),
        ("", ""),
        ("--- CRITERIOS DE IDENTIFICACIÓN ---", ""),
        ("Vacante Definitiva",
         'SITUACIONLABORAL = "Encargo Vacante Definitiva", '
         "o empleado retirado cuya plaza no fue reasignada."),
        ("Vacante Temporal",
         'SITUACIONLABORAL = "Encargo Vacante Temporal" / '
         '"Reemplazo Vacante Temporal", NOVEDAD_PLANTA = "Vacancia Temporal", '
         "o titular ausente por comisión/licencia/permiso/sanción/tutor PTA."),
        ("", ""),
        ("--- SUBREGIONES INCLUIDAS ---", ""),
    ]
    for sub in sorted(SUBREGIONES_OFICIALES):
        n = len(vacantes[vacantes["SUBREGION"] == sub])
        filas.append((sub, n))

    return pd.DataFrame(filas, columns=["Concepto", "Valor / Descripción"])


def tabla_por_dimension(vacantes: pd.DataFrame,
                        columna: str) -> pd.DataFrame:
    """Tabla cruzada de vacantes por una dimensión y tipo de vacante."""
    tabla = (
        vacantes.groupby([columna, "TIPO_VACANTE"])
        .size()
        .unstack(fill_value=0)
    )
    tabla["TOTAL"] = tabla.sum(axis=1)
    return tabla.sort_values("TOTAL", ascending=False)


def escribir_excel(vacantes: pd.DataFrame, resumen: pd.DataFrame,
                   por_municipio: pd.DataFrame, por_subregion: pd.DataFrame,
                   duplicados: pd.DataFrame, ruta_salida: str) -> None:
    """Escribe el reporte final en un archivo Excel con múltiples hojas."""
    columnas = [c for c in COLUMNAS_REPORTE if c in vacantes.columns]

    with pd.ExcelWriter(ruta_salida, engine="xlsxwriter") as writer:
        vacantes[columnas].to_excel(
            writer, sheet_name="Plazas Vacías", index=False
        )
        resumen.to_excel(writer, sheet_name="Resumen", index=False)
        por_subregion.to_excel(writer, sheet_name="Por Subregión")
        por_municipio.to_excel(writer, sheet_name="Por Municipio")
        if len(duplicados) > 0:
            duplicados.to_excel(
                writer, sheet_name="Duplicados NUM_PLAZA", index=False
            )

        wb = writer.book
        header_fmt = wb.add_format({
            "bold": True,
            "text_wrap": True,
            "valign": "top",
            "fg_color": "#4472C4",
            "font_color": "white",
            "border": 1,
        })

        ws = writer.sheets["Plazas Vacías"]
        for col_num, col_name in enumerate(columnas):
            ws.write(0, col_num, col_name, header_fmt)
        ws.autofilter(0, 0, len(vacantes), len(columnas) - 1)
        ws.set_column("A:Z", 18)

    print(f"  Archivo generado: {ruta_salida}")


# ════════════════════════════════════════════════════════════════
# MAIN
# ════════════════════════════════════════════════════════════════


def main() -> None:
    if len(sys.argv) < 2:
        print("Uso: python procesar_planta.py <archivo_entrada.xlsx> "
              "[archivo_salida.xlsx]")
        sys.exit(1)

    ruta_entrada = sys.argv[1]
    if not os.path.isfile(ruta_entrada):
        print(f"Error: No se encontró el archivo '{ruta_entrada}'")
        sys.exit(1)

    if len(sys.argv) >= 3:
        ruta_salida = sys.argv[2]
    else:
        fecha = datetime.now().strftime("%Y-%m-%d")
        ruta_salida = f"Reporte_Plazas_Vacias_{fecha}.xlsx"

    print("=" * 60)
    print("  PROCESADOR DE PLANTA DOCENTE — PLAZAS VACÍAS")
    print("=" * 60)

    print("\n[1/5] Cargando datos...")
    df = cargar_datos(ruta_entrada)

    print("\n[2/5] Filtrando subregiones oficiales de Antioquia...")
    df_filtrado = filtrar_subregiones(df)

    print("\n[3/5] Clasificando plazas vacías...")
    vacantes = clasificar_vacantes(df_filtrado)

    print("\n[4/5] Detectando duplicados e inconsistencias...")
    duplicados = detectar_duplicados(df_filtrado)
    print(f"  NUM_PLAZA duplicados (excl. 0): "
          f"{duplicados['NUM_PLAZA'].nunique():,} "
          f"({len(duplicados):,} filas)")

    print("\n[5/5] Generando reporte Excel...")
    resumen = generar_resumen(df_filtrado, vacantes)
    por_municipio = tabla_por_dimension(vacantes, "MUNICIPIO")
    por_subregion = tabla_por_dimension(vacantes, "SUBREGION")
    escribir_excel(
        vacantes, resumen, por_municipio, por_subregion,
        duplicados, ruta_salida,
    )

    print("\n" + "=" * 60)
    print(f"  COMPLETADO — {len(vacantes):,} plazas vacías exportadas")
    print("=" * 60)


if __name__ == "__main__":
    main()
