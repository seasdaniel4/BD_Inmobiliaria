# -*- coding: utf-8 -*-
"""
motor_sql.py · Ejecuta consultas T-SQL (SQL Server) sobre el Excel de la BD.

Carga todas las hojas del Excel como tablas de una base SQLite en memoria y traduce
lo propio de SQL Server (DECLARE, N'', DATEADD, TOP, STRING_SPLIT, procedimientos
almacenados con EXEC...) a SQLite. Solo usa librerías ya instaladas (sqlite3 + pandas).
"""
import re
import sqlite3
import threading

import pandas as pd


# ----------------------------------------------------------------------------
# Base de datos en memoria
# ----------------------------------------------------------------------------
def crear_conexion(origen):
    """Convierte cada hoja del Excel en una tabla. `origen`: ruta o BytesIO."""
    con = sqlite3.connect(":memory:", check_same_thread=False)
    xls = pd.ExcelFile(origen)
    for hoja in xls.sheet_names:
        df = xls.parse(hoja)
        for col in df.columns:                       # fechas como 'YYYY-MM-DD'
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                df[col] = df[col].dt.strftime("%Y-%m-%d").where(df[col].notna(), None)
        df.to_sql(hoja, con, index=False, if_exists="replace")
    con.execute("PRAGMA query_only = ON")            # base de solo lectura (app pública)
    return con


# ----------------------------------------------------------------------------
# Traducción T-SQL -> SQLite
# ----------------------------------------------------------------------------
def _quitar_comentarios(sql):
    sql = re.sub(r"/\*.*?\*/", " ", sql, flags=re.S)
    return re.sub(r"--[^\n]*", "", sql)


def _valor_literal(v):
    """Convierte un valor Python (o texto) a literal SQL."""
    if v is None:
        return "NULL"
    if isinstance(v, (list, tuple)):
        v = ",".join(map(str, v))
    if isinstance(v, (int, float)):
        return str(v)
    v = str(v).strip()
    if v[:2].upper() == "N'":                         # N'texto' -> 'texto'
        v = v[1:]
    if v.upper() == "NULL" or re.fullmatch(r"-?\d+(\.\d+)?", v) or (v.startswith("'") and v.endswith("'")):
        return v
    return "'" + v.replace("'", "''") + "'"


def _sustituir_variables(sql, variables):
    # Primero las de nombre más largo, para no pisar @Zona con @ZonaX
    for nombre in sorted(variables, key=len, reverse=True):
        sql = re.sub(rf"@{nombre}\b", lambda m: _valor_literal(variables[nombre]), sql, flags=re.I)
    return sql


def _traducir_funciones(sql):
    sql = re.sub(r"\bN'", "'", sql)                                   # N'texto' -> 'texto'
    sql = re.sub(r"\bdbo\.", "", sql, flags=re.I)
    sql = re.sub(r"DATEADD\(\s*(MONTH|DAY|YEAR)\s*,\s*(-?\d+)\s*,\s*([^()]+?)\s*\)",
                 lambda m: f"date({m.group(3)}, '{int(m.group(2)):+d} {m.group(1).lower()}s')",
                 sql, flags=re.I)
    sql = re.sub(r"CAST\(([^()]+?)\s+AS\s+(?:DECIMAL|NUMERIC|MONEY|FLOAT)(?:\(\d+\s*,?\s*\d*\))?\)",
                 r"CAST(\1 AS REAL)", sql, flags=re.I)
    sql = re.sub(r"CAST\(([^()]+?)\s+AS\s+(?:NVARCHAR|VARCHAR)(?:\(\w+\))?\)",
                 r"CAST(\1 AS TEXT)", sql, flags=re.I)
    sql = re.sub(r"\bISNULL\s*\(", "IFNULL(", sql, flags=re.I)
    sql = re.sub(r"\bLEN\s*\(", "LENGTH(", sql, flags=re.I)
    sql = re.sub(r"\bGETDATE\s*\(\s*\)", "date('now')", sql, flags=re.I)
    sql = re.sub(r"\bYEAR\s*\(([^()]+)\)", r"CAST(strftime('%Y', \1) AS INTEGER)", sql, flags=re.I)
    sql = re.sub(r"\bMONTH\s*\(([^()]+)\)", r"CAST(strftime('%m', \1) AS INTEGER)", sql, flags=re.I)
    return sql


def _top_a_limit(sql):
    m = re.search(r"\bSELECT\s+TOP\s+(\d+)\s+", sql, flags=re.I)
    if m and not re.search(r"\bLIMIT\b", sql, flags=re.I):
        sql = re.sub(r"\bSELECT\s+TOP\s+\d+\s+", "SELECT ", sql, count=1, flags=re.I)
        sql = sql.rstrip().rstrip(";") + f" LIMIT {m.group(1)}"
    return sql


def _expandir_string_split(sql):
    """IN (SELECT LTRIM(value) FROM STRING_SPLIT('a, b', ',')) -> IN ('a','b')."""
    def reemplazo(m):
        items = [x.strip() for x in m.group(1).split(",") if x.strip()]
        return "(" + ", ".join("'" + x.replace("'", "''") + "'" for x in items) + ")"
    return re.sub(
        r"\(\s*SELECT\s+(?:LTRIM|RTRIM|TRIM)?\(?\s*value\s*\)?\s+FROM\s+STRING_SPLIT\(\s*'([^']*)'\s*,\s*','\s*\)\s*\)",
        reemplazo, sql, flags=re.I)


# ----------------------------------------------------------------------------
# Procedimientos almacenados
# ----------------------------------------------------------------------------
_RE_PROC = re.compile(
    r"CREATE\s+(?:OR\s+ALTER\s+)?PROC(?:EDURE)?\s+(\w+)(.*?)\bAS\b\s*BEGIN(.*)\bEND\b\s*;?\s*$",
    re.I | re.S)


def _parsear_procedimiento(sql):
    m = _RE_PROC.search(sql.strip())
    if not m:
        return None
    nombre, cabecera, cuerpo = m.group(1), m.group(2), m.group(3)
    params = {}
    for p in re.finditer(r"@(\w+)\s+\w+(?:\(\s*\w+(?:\s*,\s*\d+)?\s*\))?\s*(?:=\s*('[^']*'|[^,\s]+))?", cabecera):
        params[p.group(1)] = p.group(2) if p.group(2) is not None else "NULL"
    cuerpo = re.sub(r"\bSET\s+NOCOUNT\s+ON\s*;", "", cuerpo, flags=re.I)
    return nombre, params, cuerpo.strip()


def _parsear_exec(sql):
    m = re.match(r"\s*EXEC(?:UTE)?\s+(\w+)\s*(.*?);?\s*$", sql, flags=re.I | re.S)
    if not m:
        return None
    args = {}
    for a in re.finditer(r"@(\w+)\s*=\s*('(?:[^']|'')*'|N'(?:[^']|'')*'|[^,\s]+)", m.group(2)):
        args[a.group(1)] = a.group(2)
    return m.group(1), args


# ----------------------------------------------------------------------------
# Ejecución
# ----------------------------------------------------------------------------
def traducir(sql, variables=None):
    """T-SQL (una sola sentencia SELECT/CTE, con DECLARE opcionales) -> SQL de SQLite."""
    sql = _quitar_comentarios(sql)
    variables = dict(variables or {})
    for m in re.finditer(r"DECLARE\s+@(\w+)\s+\w+(?:\(\s*\w+(?:\s*,\s*\d+)?\s*\))?\s*=\s*('[^']*'|N'[^']*'|[^;\s]+)\s*;?",
                         sql, flags=re.I):
        variables.setdefault(m.group(1), m.group(2))
    sql = re.sub(r"DECLARE\s+@\w+[^;]*;", "", sql, flags=re.I)
    sql = _traducir_funciones(sql)
    sql = _expandir_string_split(_sustituir_variables(sql, variables))
    sql = _top_a_limit(re.sub(r"^\s*GO\s*$", "", sql, flags=re.I | re.M))
    return sql.strip().rstrip(";")


_CANDADO = threading.Lock()                          # la conexión se comparte entre sesiones
_RE_PROHIBIDO = re.compile(r"\b(DROP|DELETE|UPDATE|INSERT|ALTER|ATTACH|DETACH|PRAGMA|REPLACE|VACUUM|TRUNCATE)\b", re.I)


def _validar_solo_lectura(sql):
    if not re.match(r"\s*(SELECT|WITH)\b", sql, flags=re.I) or _RE_PROHIBIDO.search(re.sub(r"'[^']*'", "''", sql)):
        raise ValueError("Solo se permiten consultas de lectura (SELECT / WITH) y EXEC de procedimientos de consulta.")


def ejecutar(con, script, procedimientos=None):
    """
    Ejecuta un script T-SQL y devuelve (DataFrame, sql_traducido).
    Admite: DECLARE + SELECT/CTE, CREATE PROCEDURE (se registra) y EXEC proc @p = valor.
    `procedimientos` es un dict que se rellena con los procedimientos creados.
    """
    procedimientos = procedimientos if procedimientos is not None else {}
    limpio = _quitar_comentarios(script).strip()

    # El script puede traer CREATE PROCEDURE ... seguido de EXEC ...
    partes = re.split(r"(?=^\s*EXEC(?:UTE)?\s)", limpio, flags=re.I | re.M)
    resultado, ultimo_sql = None, ""
    for parte in [p for p in partes if p.strip()]:
        proc = _parsear_procedimiento(parte)
        if proc:
            procedimientos[proc[0].lower()] = proc[1:]
            continue
        ex = _parsear_exec(parte)
        if ex:
            nombre, args = ex
            if nombre.lower() not in procedimientos:
                raise ValueError(f"El procedimiento «{nombre}» no existe: ejecuta primero su CREATE PROCEDURE.")
            defaults, cuerpo = procedimientos[nombre.lower()]
            vars_ = {**defaults, **args}
            sql = _top_a_limit(_expandir_string_split(
                _sustituir_variables(_traducir_funciones(cuerpo), vars_))).strip().rstrip(";")
        else:
            sql = traducir(parte)
        ultimo_sql = sql
        _validar_solo_lectura(sql)
        with _CANDADO:
            resultado = pd.read_sql_query(sql, con)

    if resultado is None:
        return None, ultimo_sql                      # solo se registró un procedimiento
    return resultado, ultimo_sql
