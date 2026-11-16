"""
App Streamlit · BD Inmobiliaria
-------------------------------
Muestra las consultas Q8, Q13 y Q25 (Buscador) sobre el Excel de la base de datos.
Cada consulta incluye: parámetros, resultado, gráfico, interpretación y el SQL original.

Local:      pip install -r requirements.txt  ->  streamlit run app.py
Streamlit Community Cloud: sube app.py, motor_sql.py, requirements.txt y el Excel a GitHub
y elige app.py como archivo principal (ver LEEME_DESPLIEGUE.md).
"""
import io
import os

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

import motor_sql

# ----------------------------------------------------------------------------
# Configuración general
# ----------------------------------------------------------------------------
st.set_page_config(page_title="BD Inmobiliaria", page_icon="🏠", layout="wide")

CARPETA_APP = os.path.dirname(os.path.abspath(__file__))   # rutas relativas al repositorio, no al cwd
FECHA_CORTE_DEFECTO = pd.Timestamp("2023-12-31")
ORDEN_DISP = ["Inmediata", "Corto plazo", "Largo plazo", "Atrasada", "Sin fecha"]
COLORES_DISP = {
    "Inmediata": "#2E9E5B",
    "Corto plazo": "#3B82C4",
    "Largo plazo": "#8E6BBF",
    "Atrasada": "#D9534F",
    "Sin fecha": "#9AA0A6",
}


def mostrar_df(df, **kw):
    """st.dataframe a todo el ancho (compatible con versiones nuevas y antiguas de Streamlit)."""
    try:
        st.dataframe(df, width="stretch", **kw)
    except TypeError:
        st.dataframe(df, use_container_width=True, **kw)


def mostrar_fig(fig, contenedor=None):
    destino = contenedor or st
    try:
        destino.plotly_chart(fig, width="stretch")
    except TypeError:
        destino.plotly_chart(fig, use_container_width=True)


def buscar_excel(nombre_pedido=""):
    """Devuelve la ruta del Excel de la BD. Busca (sin distinguir mayúsculas) en la carpeta de la
    app y en la actual: primero el nombre pedido, luego cualquier .xlsx con 'inmobiliaria', luego
    cualquier .xlsx. Devuelve (ruta o None, lista de .xlsx encontrados)."""
    carpetas = []
    for c in (CARPETA_APP, os.getcwd()):
        if c not in carpetas and os.path.isdir(c):
            carpetas.append(c)
    # Ruta escrita por el usuario (absoluta o relativa)
    if nombre_pedido:
        for cand in [nombre_pedido] + [os.path.join(c, nombre_pedido) for c in carpetas]:
            if os.path.isfile(cand):
                return cand, [cand]
    encontrados = []
    for c in carpetas:
        for raiz, _, archivos in os.walk(c):
            if any(p in raiz for p in (".git", "__pycache__", "site-packages")):
                continue
            for f in archivos:
                if f.lower().endswith(".xlsx") and not f.startswith("~$"):
                    encontrados.append(os.path.join(raiz, f))
        if encontrados:
            break
    if not encontrados:
        return None, []
    if nombre_pedido:                                   # mismo nombre pero distinta capitalización
        for f in encontrados:
            if os.path.basename(f).lower() == os.path.basename(nombre_pedido).lower():
                return f, encontrados
    preferidos = [f for f in encontrados if "inmobiliaria" in os.path.basename(f).lower()]
    return (preferidos or encontrados)[0], encontrados


# ----------------------------------------------------------------------------
# Carga de datos
# ----------------------------------------------------------------------------
@st.cache_data(show_spinner="Cargando base de datos...")
def cargar_datos(origen):
    """Lee las hojas necesarias del Excel. `origen` es una ruta (str) o bytes."""
    fuente = io.BytesIO(origen) if isinstance(origen, (bytes, bytearray)) else origen
    xls = pd.ExcelFile(fuente)
    ciudad = xls.parse("CIUDAD")
    area = xls.parse("AREA")
    tipo = xls.parse("TIPO_PROPIEDAD")
    prop = xls.parse("PROPIEDAD")
    return ciudad, area, tipo, prop


@st.cache_resource(show_spinner="Creando base SQL en memoria...")
def cargar_conexion(origen):
    """Carga TODAS las hojas del Excel como tablas SQL (para ejecutar consultas reales)."""
    fuente = io.BytesIO(origen) if isinstance(origen, (bytes, bytearray)) else origen
    return motor_sql.crear_conexion(fuente)


def clasificar_disponibilidad(estado, fecha, corte):
    """Réplica exacta del CASE de las consultas Q13 y Q25."""
    corte = pd.Timestamp(corte)
    limite = corte + pd.DateOffset(months=12)
    condiciones = [
        estado == "Listo para entregar",
        fecha.isna(),
        fecha <= corte,
        fecha <= limite,
    ]
    resultados = ["Inmediata", "Sin fecha", "Atrasada", "Corto plazo"]
    return pd.Series(np.select(condiciones, resultados, default="Largo plazo"),
                     index=estado.index)


def construir_base(ciudad, area, tipo, prop):
    """Une PROPIEDAD con AREA, CIUDAD y TIPO_PROPIEDAD (los JOIN de las consultas)."""
    base = (prop
            .merge(area, on="Area_ID", how="inner")
            .merge(ciudad, on="Ciudad_ID", how="inner")
            .merge(tipo, on="Tipo_Propiedad_ID", how="inner"))
    return base


# ----------------------------------------------------------------------------
# Lógica de cada consulta
# ----------------------------------------------------------------------------
def consulta_q8(base, min_propiedades):
    """Q8: zonas con precio por pie² mayor al promedio del portal."""
    d = base[base["Area_Construida"].notna()].copy()
    d["PrecioPie2"] = d["Precio"].astype(float) / d["Area_Construida"]
    promedio_portal = d["PrecioPie2"].mean()

    zonas = (d.groupby(["Area_ID", "Nombre_Area", "Nombre_Ciudad"], as_index=False)
              .agg(Propiedades=("PrecioPie2", "size"), PrecioPie2Zona=("PrecioPie2", "mean")))
    evaluadas = zonas[zonas["Propiedades"] >= min_propiedades].copy()
    res = evaluadas[evaluadas["PrecioPie2Zona"] > promedio_portal].copy()
    res["PrecioPie2Portal"] = promedio_portal
    res["PctSobrePortal"] = (res["PrecioPie2Zona"] / promedio_portal - 1) * 100
    res = res.rename(columns={"Nombre_Area": "Zona", "Nombre_Ciudad": "Ciudad"})
    res = res.sort_values("PctSobrePortal", ascending=False).reset_index(drop=True)
    res = res[["Zona", "Ciudad", "Propiedades", "PrecioPie2Zona",
               "PrecioPie2Portal", "PctSobrePortal"]]
    return res, promedio_portal, len(evaluadas)


def consulta_q13(base, corte):
    """Q13: clasificación de disponibilidad (detalle y resumen)."""
    d = base.copy()
    d["Disponibilidad"] = clasificar_disponibilidad(d["Estado_Posesion"], d["Fecha_Entrega"], corte)
    corte = pd.Timestamp(corte)
    f = d["Fecha_Entrega"]
    d["MesesParaEntrega"] = ((f.dt.year - corte.year) * 12 + (f.dt.month - corte.month))

    detalle = d[["Propiedad_ID", "Nombre_Area", "Estado_Posesion", "Fecha_Entrega",
                 "MesesParaEntrega", "Disponibilidad"]].rename(columns={"Nombre_Area": "Zona"})
    detalle["Fecha_Entrega"] = detalle["Fecha_Entrega"].dt.strftime("%d/%m/%Y")
    detalle["orden"] = detalle["Disponibilidad"].map({k: i for i, k in enumerate(ORDEN_DISP)})
    detalle = detalle.sort_values(["orden", "MesesParaEntrega"]).drop(columns="orden")

    resumen = (d["Disponibilidad"].value_counts()
               .reindex(ORDEN_DISP, fill_value=0)
               .rename_axis("Disponibilidad").reset_index(name="Propiedades"))
    resumen["Porcentaje"] = (100.0 * resumen["Propiedades"] / resumen["Propiedades"].sum()).round(1)
    return detalle.reset_index(drop=True), resumen


def consulta_q25(base, corte, ciudad=None, zona=None, precio_min=None, precio_max=None,
                 dormitorios=None, tipo=None, disponibilidad=None):
    """Q25: equivalente a sp_BuscarPropiedades. Parámetro None = filtro no aplicado."""
    d = base.copy()
    d["Disponibilidad"] = clasificar_disponibilidad(d["Estado_Posesion"], d["Fecha_Entrega"], corte)
    mask = pd.Series(True, index=d.index)
    if ciudad is not None:
        mask &= d["Nombre_Ciudad"] == ciudad
    if zona is not None:
        mask &= d["Nombre_Area"] == zona
    if precio_min is not None:
        mask &= d["Precio"] >= precio_min
    if precio_max is not None:
        mask &= d["Precio"] <= precio_max
    if dormitorios is not None:
        mask &= (d["Dormitorios"] == dormitorios) if dormitorios < 4 else (d["Dormitorios"] >= 4)
    if tipo is not None:
        mask &= d["Nombre_Tipo"] == tipo
    if disponibilidad:
        mask &= d["Disponibilidad"].isin(disponibilidad)

    res = d.loc[mask, ["Propiedad_ID", "Nombre_Ciudad", "Nombre_Area", "Dormitorios",
                       "Nombre_Tipo", "Precio", "Disponibilidad"]]
    res = res.rename(columns={"Propiedad_ID": "Propiedad", "Nombre_Ciudad": "Ciudad",
                              "Nombre_Area": "Zona", "Nombre_Tipo": "Tipo"})
    return res.sort_values("Precio").reset_index(drop=True)


# ----------------------------------------------------------------------------
# Utilidades de formato
# ----------------------------------------------------------------------------
def fmt_num(x, dec=0):
    return f"{x:,.{dec}f}"


def fmt_millones(x):
    return f"{x / 1e6:,.2f} M"


def bloque_interpretacion(texto):
    st.markdown("#### 💡 Interpretación")
    st.info(texto)


def mostrar_sql(titulo, codigo):
    with st.expander(f"Ver SQL original · {titulo}"):
        st.code(codigo, language="sql")


# ----------------------------------------------------------------------------
# SQL originales (solo para mostrarlos)
# ----------------------------------------------------------------------------
SQL_Q8 = """DECLARE @MinPropiedades INT = 5;   -- zonas con al menos 5 avisos (1 = todas las zonas)

WITH PrecioPie2 AS (
    SELECT P.Area_ID,
           CAST(P.Precio AS DECIMAL(18,2)) / P.Area_Construida AS PrecioPie2
    FROM PROPIEDAD AS P
    WHERE P.Area_Construida IS NOT NULL
),
PromedioZona AS (
    SELECT PP.Area_ID, COUNT(*) AS Propiedades, AVG(PP.PrecioPie2) AS PromZona
    FROM PrecioPie2 AS PP
    GROUP BY PP.Area_ID
    HAVING COUNT(*) >= @MinPropiedades
)
SELECT A.Nombre_Area AS Zona, C.Nombre_Ciudad AS Ciudad, PZ.Propiedades,
       ROUND(PZ.PromZona, 0) AS PrecioPie2Zona,
       ROUND((SELECT AVG(PrecioPie2) FROM PrecioPie2), 0) AS PrecioPie2Portal,
       ROUND((PZ.PromZona / (SELECT AVG(PrecioPie2) FROM PrecioPie2) - 1) * 100, 1) AS PctSobrePortal
FROM PromedioZona AS PZ
JOIN AREA   AS A ON A.Area_ID   = PZ.Area_ID
JOIN CIUDAD AS C ON C.Ciudad_ID = A.Ciudad_ID
WHERE PZ.PromZona > (SELECT AVG(PrecioPie2) FROM PrecioPie2)
ORDER BY PctSobrePortal DESC;"""

SQL_Q13 = """DECLARE @FechaCorte DATE = '2023-12-31';

WITH Clasificacion AS (
    SELECT P.Propiedad_ID,
           CASE
               WHEN P.Estado_Posesion = N'Listo para entregar'         THEN 'Inmediata'
               WHEN P.Fecha_Entrega IS NULL                            THEN 'Sin fecha'
               WHEN P.Fecha_Entrega <= @FechaCorte                     THEN 'Atrasada'
               WHEN P.Fecha_Entrega <= DATEADD(MONTH, 12, @FechaCorte) THEN 'Corto plazo'
               ELSE 'Largo plazo'
           END AS Disponibilidad
    FROM PROPIEDAD AS P
)
SELECT CL.Disponibilidad,
       COUNT(*) AS Propiedades,
       ROUND(100.0 * COUNT(*) / SUM(COUNT(*)) OVER (), 1) AS Porcentaje
FROM Clasificacion AS CL
GROUP BY CL.Disponibilidad;"""

SQL_Q25 = """CREATE PROCEDURE sp_BuscarPropiedades
    @Ciudad NVARCHAR(50) = NULL, @Zona NVARCHAR(60) = NULL,
    @PrecioMin MONEY = NULL, @PrecioMax MONEY = NULL,
    @Dormitorios INT = NULL, @Tipo NVARCHAR(30) = NULL,
    @Disponibilidad NVARCHAR(200) = NULL, @FechaCorte DATE = '2023-12-31'
AS
BEGIN
    SET NOCOUNT ON;
    WITH Base AS (
        SELECT P.Propiedad_ID, C.Nombre_Ciudad AS Ciudad, A.Nombre_Area AS Zona,
               P.Dormitorios, T.Nombre_Tipo AS Tipo, P.Precio,
               CASE
                   WHEN P.Estado_Posesion = N'Listo para entregar'         THEN 'Inmediata'
                   WHEN P.Fecha_Entrega IS NULL                            THEN 'Sin fecha'
                   WHEN P.Fecha_Entrega <= @FechaCorte                     THEN 'Atrasada'
                   WHEN P.Fecha_Entrega <= DATEADD(MONTH, 12, @FechaCorte) THEN 'Corto plazo'
                   ELSE 'Largo plazo'
               END AS Disponibilidad
        FROM PROPIEDAD AS P
        JOIN AREA AS A ON A.Area_ID = P.Area_ID
        JOIN CIUDAD AS C ON C.Ciudad_ID = A.Ciudad_ID
        JOIN TIPO_PROPIEDAD AS T ON T.Tipo_Propiedad_ID = P.Tipo_Propiedad_ID
    )
    SELECT B.Propiedad_ID AS Propiedad, B.Ciudad, B.Zona, B.Dormitorios,
           B.Tipo, B.Precio, B.Disponibilidad
    FROM Base AS B
    WHERE (@Ciudad IS NULL OR B.Ciudad = @Ciudad)
      AND (@Zona IS NULL OR B.Zona = @Zona)
      AND (@PrecioMin IS NULL OR B.Precio >= @PrecioMin)
      AND (@PrecioMax IS NULL OR B.Precio <= @PrecioMax)
      AND (@Dormitorios IS NULL
           OR (@Dormitorios < 4  AND B.Dormitorios = @Dormitorios)
           OR (@Dormitorios >= 4 AND B.Dormitorios >= 4))
      AND (@Tipo IS NULL OR B.Tipo = @Tipo)
      AND (@Disponibilidad IS NULL
           OR B.Disponibilidad IN (SELECT LTRIM(value) FROM STRING_SPLIT(@Disponibilidad, ',')))
    ORDER BY B.Precio;
END;"""


# ----------------------------------------------------------------------------
# Vistas (una por consulta)
# ----------------------------------------------------------------------------
def vista_q8(base):
    st.subheader("Q8 · Zonas con precio por pie² superior al promedio del portal")
    st.caption("¿Qué zonas tienen un precio por pie² mayor al promedio del portal, y en qué % lo superan?")

    c1, c2 = st.columns([1, 1])
    min_prop = c1.slider("Mínimo de avisos por zona (@MinPropiedades)", 1, 50, 5,
                         help="1 = se consideran todas las zonas", key="q8_min")
    top_n = c2.slider("Zonas a mostrar en el gráfico", 5, 40, 15, key="q8_top")

    res, prom_portal, n_evaluadas = consulta_q8(base, min_prop)

    m1, m2, m3 = st.columns(3)
    m1.metric("Promedio del portal (por pie²)", fmt_num(prom_portal))
    m2.metric("Zonas evaluadas", fmt_num(n_evaluadas))
    m3.metric("Zonas sobre el promedio", fmt_num(len(res)),
              f"{100 * len(res) / max(n_evaluadas, 1):.0f}% de las evaluadas", delta_color="off")

    if res.empty:
        st.warning("Ninguna zona cumple las condiciones con el mínimo de avisos elegido.")
        return

    graf = res.head(top_n).copy()
    graf["Etiqueta"] = graf["Zona"] + " (" + graf["Ciudad"] + ")"
    fig = px.bar(graf.sort_values("PctSobrePortal"), x="PctSobrePortal", y="Etiqueta",
                 color="Ciudad", orientation="h", text="PctSobrePortal",
                 hover_data={"Propiedades": True, "PrecioPie2Zona": ":,.0f", "Etiqueta": False},
                 labels={"PctSobrePortal": "% sobre el promedio del portal", "Etiqueta": ""},
                 title=f"Top {len(graf)} zonas más caras por pie² respecto al promedio del portal")
    fig.update_traces(texttemplate="%{text:.0f}%", textposition="outside")
    fig.update_layout(height=max(400, 28 * len(graf) + 150), legend_title_text="Ciudad")
    mostrar_fig(fig)

    st.markdown("#### Resultado de la consulta")
    mostrar_df(res.assign(PrecioPie2Zona=res["PrecioPie2Zona"].round(0),
                          PrecioPie2Portal=res["PrecioPie2Portal"].round(0),
                          PctSobrePortal=res["PctSobrePortal"].round(1)), hide_index=True)

    top = res.iloc[0]
    por_ciudad = res.groupby("Ciudad").size().sort_values(ascending=False)
    veces = top["PrecioPie2Zona"] / prom_portal
    bloque_interpretacion(
        f"El promedio del portal es de **{fmt_num(prom_portal)} por pie²**. De las "
        f"**{n_evaluadas}** zonas con al menos {min_prop} aviso(s), **{len(res)}** "
        f"({100 * len(res) / n_evaluadas:.0f}%) lo superan. "
        f"La zona más cara es **{top['Zona']} ({top['Ciudad']})**, con "
        f"{fmt_num(top['PrecioPie2Zona'])} por pie², es decir **{top['PctSobrePortal']:.1f}% por encima** "
        f"del promedio (≈ {veces:.1f} veces el valor del portal). "
        f"La ciudad que concentra más zonas premium es **{por_ciudad.index[0]}** "
        f"({por_ciudad.iloc[0]} zonas). La mediana de sobreprecio entre las zonas listadas es "
        f"{res['PctSobrePortal'].median():.1f}%.\n\n"
        f"*Ojo:* un mínimo de avisos bajo puede incluir zonas con pocos datos, donde un solo aviso "
        f"muy caro distorsiona el promedio; subir el mínimo da una lectura más robusta."
    )
    mostrar_sql("Q8", SQL_Q8)


def vista_q13(base):
    st.subheader("Q13 · Disponibilidad de las propiedades para mudarse")
    st.caption("Clasifica cada propiedad según cuándo estará disponible: Inmediata, Corto plazo, "
               "Largo plazo, Atrasada o Sin fecha.")

    corte = st.date_input("Fecha de corte (\"hoy\")", FECHA_CORTE_DEFECTO.date(), key="q13_corte")
    detalle, resumen = consulta_q13(base, corte)

    cols = st.columns(len(ORDEN_DISP))
    for col, (_, fila) in zip(cols, resumen.iterrows()):
        col.metric(fila["Disponibilidad"], fmt_num(fila["Propiedades"]), f"{fila['Porcentaje']}%",
                   delta_color="off")

    g1, g2 = st.columns(2)
    fig_bar = px.bar(resumen, x="Disponibilidad", y="Propiedades", color="Disponibilidad",
                     color_discrete_map=COLORES_DISP, text="Porcentaje",
                     category_orders={"Disponibilidad": ORDEN_DISP},
                     title="Propiedades por categoría de disponibilidad")
    fig_bar.update_traces(texttemplate="%{text}%", textposition="outside")
    fig_bar.update_layout(showlegend=False)
    mostrar_fig(fig_bar, g1)

    # Solo propiedades en construcción con fecha: distribución de meses para la entrega
    construccion = detalle[detalle["Disponibilidad"].isin(["Corto plazo", "Largo plazo", "Atrasada"])]
    if not construccion.empty:
        fig_hist = px.histogram(construccion, x="MesesParaEntrega", color="Disponibilidad",
                                color_discrete_map=COLORES_DISP, nbins=40,
                                labels={"MesesParaEntrega": "Meses para la entrega (negativo = ya pasó)"},
                                title="Obras en construcción: meses hasta la entrega")
        mostrar_fig(fig_hist, g2)
    else:
        g2.info("No hay propiedades en construcción con fecha para graficar.")

    r = resumen.set_index("Disponibilidad")
    total = int(r["Propiedades"].sum())
    en_obra = int(r.loc[["Corto plazo", "Largo plazo", "Atrasada", "Sin fecha"], "Propiedades"].sum())
    dominante = r["Propiedades"].idxmax()
    bloque_interpretacion(
        f"Con corte al **{pd.Timestamp(corte):%d/%m/%Y}**, de **{fmt_num(total)}** propiedades, "
        f"**{fmt_num(r.loc['Inmediata', 'Propiedades'])} ({r.loc['Inmediata', 'Porcentaje']}%)** están "
        f"listas para mudarse de inmediato; la categoría predominante es **{dominante}**. "
        f"Las {fmt_num(en_obra)} restantes ({100 * en_obra / total:.1f}%) no están listas: "
        f"{fmt_num(r.loc['Corto plazo', 'Propiedades'])} se entregan en 12 meses o menos y "
        f"{fmt_num(r.loc['Largo plazo', 'Propiedades'])} tardarán más. "
        f"Hay **{fmt_num(r.loc['Atrasada', 'Propiedades'])}** propiedades **atrasadas** "
        f"(fecha de entrega vencida) y **{fmt_num(r.loc['Sin fecha', 'Propiedades'])}** sin fecha "
        f"registrada, que son riesgos para el comprador y datos a depurar en el portal."
    )

    with st.expander("Ver detalle por propiedad (13.a)"):
        filtro = st.multiselect("Filtrar por categoría", ORDEN_DISP, default=ORDEN_DISP, key="q13_f")
        mostrar_df(detalle[detalle["Disponibilidad"].isin(filtro)], hide_index=True)
    st.markdown("#### Resumen (13.b)")
    mostrar_df(resumen, hide_index=True)
    mostrar_sql("Q13", SQL_Q13)


def vista_q25(base, con):
    st.subheader("Q25 · Buscador de propiedades (interactivo)")
    st.caption("Cambia los filtros y la consulta se recalcula al instante. Un filtro en \"Todos\" "
               "equivale a enviar NULL al procedimiento sp_BuscarPropiedades.")

    ciudades = sorted(base["Nombre_Ciudad"].unique())
    tipos = sorted(base["Nombre_Tipo"].unique())
    pmin_dato, pmax_dato = int(base["Precio"].min()), int(base["Precio"].max())

    f1, f2, f3 = st.columns(3)
    ciudad_sel = f1.selectbox("Ciudad", ["Todas"] + ciudades, key="q25_ciudad")
    zonas = base if ciudad_sel == "Todas" else base[base["Nombre_Ciudad"] == ciudad_sel]
    zona_sel = f2.selectbox("Zona", ["Todas"] + sorted(zonas["Nombre_Area"].unique()),
                            key=f"q25_zona_{ciudad_sel}")
    tipo_sel = f3.selectbox("Tipo de propiedad", ["Todos"] + tipos, key="q25_tipo")

    f4, f5 = st.columns([2, 1])
    rango = f4.slider("Rango de precio", pmin_dato, pmax_dato, (pmin_dato, pmax_dato),
                      step=100_000, format="%d", key="q25_precio")
    dorm_sel = f5.selectbox("Dormitorios", ["Todos", "1", "2", "3", "4 o más"], key="q25_dorm")

    disp_sel = st.multiselect("Disponibilidad (vacío = todas)", ORDEN_DISP, key="q25_disp")
    corte = st.date_input("Fecha de corte (\"hoy\")", FECHA_CORTE_DEFECTO.date(), key="q25_corte")

    dorm_param = None if dorm_sel == "Todos" else (4 if dorm_sel == "4 o más" else int(dorm_sel))
    params = dict(
        ciudad=None if ciudad_sel == "Todas" else ciudad_sel,
        zona=None if zona_sel == "Todas" else zona_sel,
        precio_min=rango[0] if rango[0] > pmin_dato else None,
        precio_max=rango[1] if rango[1] < pmax_dato else None,
        dormitorios=dorm_param,
        tipo=None if tipo_sel == "Todos" else tipo_sel,
        disponibilidad=disp_sel or None,
    )

    # Llamada equivalente en SQL Server
    def sql_val(v):
        if v is None:
            return "NULL"
        if isinstance(v, (int, float, np.integer)):
            return str(v)
        if isinstance(v, list):
            return "N'" + ",".join(v) + "'"
        return f"N'{v}'"
    exec_sql = ("EXEC sp_BuscarPropiedades\n    "
                f"@Ciudad = {sql_val(params['ciudad'])}, @Zona = {sql_val(params['zona'])},\n    "
                f"@PrecioMin = {sql_val(params['precio_min'])}, @PrecioMax = {sql_val(params['precio_max'])},\n    "
                f"@Dormitorios = {sql_val(params['dormitorios'])}, @Tipo = {sql_val(params['tipo'])},\n    "
                f"@Disponibilidad = {sql_val(params['disponibilidad'])}, "
                f"@FechaCorte = '{pd.Timestamp(corte):%Y-%m-%d}';")
    st.markdown("**Llamada equivalente en SQL Server:**")
    st.code(exec_sql, language="sql")

    usar_sql = st.checkbox("Calcular con el SQL real (ejecuta el EXEC de arriba sobre la base)", key="q25_sql")
    if usar_sql:
        try:
            res, _ = motor_sql.ejecutar(con, SQL_Q25 + "\n" + exec_sql)
        except Exception as e:
            st.error(f"Error ejecutando el SQL: {e}")
            return
    else:
        res = consulta_q25(base, corte, **params)

    st.markdown("---")
    if res.empty:
        st.warning("Ninguna propiedad cumple con los filtros elegidos. Prueba ampliando el rango de precio "
                   "o quitando algún filtro.")
        return

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Propiedades encontradas", fmt_num(len(res)))
    m2.metric("Precio mínimo", fmt_millones(res["Precio"].min()))
    m3.metric("Precio mediano", fmt_millones(res["Precio"].median()))
    m4.metric("Precio máximo", fmt_millones(res["Precio"].max()))

    g1, g2 = st.columns(2)
    graf = res.assign(PrecioM=res["Precio"] / 1e6)
    fig_h = px.histogram(graf, x="PrecioM", nbins=30, color="Disponibilidad",
                         color_discrete_map=COLORES_DISP,
                         category_orders={"Disponibilidad": ORDEN_DISP},
                         labels={"PrecioM": "Precio (millones)", "count": "Propiedades"},
                         title="Distribución de precios de los resultados")
    fig_h.update_layout(yaxis_title="Propiedades", legend_title_text="Disponibilidad")
    mostrar_fig(fig_h, g1)

    # Segundo gráfico: zonas (si hay una ciudad/zona concreta) o ciudades
    if params["zona"] is not None:
        conteo = res.groupby("Tipo").size().reset_index(name="Propiedades")
        fig_b = px.bar(conteo.sort_values("Propiedades"), x="Propiedades", y="Tipo", orientation="h",
                       title="Resultados por tipo de propiedad", text="Propiedades")
    elif params["ciudad"] is not None:
        conteo = (res.groupby("Zona").size().reset_index(name="Propiedades")
                  .sort_values("Propiedades", ascending=False).head(10))
        fig_b = px.bar(conteo.sort_values("Propiedades"), x="Propiedades", y="Zona", orientation="h",
                       title=f"Top 10 zonas con más resultados en {params['ciudad']}", text="Propiedades")
    else:
        conteo = res.groupby("Ciudad").size().reset_index(name="Propiedades")
        fig_b = px.bar(conteo.sort_values("Propiedades"), x="Propiedades", y="Ciudad", orientation="h",
                       title="Resultados por ciudad", text="Propiedades")
    mostrar_fig(fig_b, g2)

    st.markdown("#### Propiedades que cumplen la búsqueda")
    mostrar_df(res, hide_index=True,
               column_config={"Precio": st.column_config.NumberColumn(format="%d")})
    st.download_button("⬇️ Descargar resultados (CSV)", res.to_csv(index=False).encode("utf-8-sig"),
                       "resultados_busqueda.csv", "text/csv")

    disp_top = res["Disponibilidad"].value_counts()
    zona_top = res["Zona"].value_counts()
    inm = (res["Disponibilidad"] == "Inmediata").mean() * 100
    filtros_activos = [k for k, v in params.items() if v is not None]
    bloque_interpretacion(
        f"La búsqueda con {len(filtros_activos)} filtro(s) activo(s) devuelve **{fmt_num(len(res))}** "
        f"propiedades (de {fmt_num(len(base))} en total, {100 * len(res) / len(base):.1f}%). "
        f"Sus precios van de **{fmt_millones(res['Precio'].min())}** a **{fmt_millones(res['Precio'].max())}**, "
        f"con una mediana de **{fmt_millones(res['Precio'].median())}**"
        + f". El {inm:.0f}% está disponible de inmediato y la categoría más frecuente es "
        f"**{disp_top.index[0]}** ({disp_top.iloc[0]} propiedades). "
        f"La zona con más resultados es **{zona_top.index[0]}** ({zona_top.iloc[0]}). "
        + ("Si el histograma está muy cargado a la izquierda, la mayoría de la oferta se concentra en "
           "precios bajos y unas pocas propiedades de lujo estiran la cola derecha."
           if res["Precio"].mean() > res["Precio"].median() * 1.2 else
           "La distribución de precios es relativamente pareja, sin una cola de lujo marcada.")
    )
    mostrar_sql("Q25 · sp_BuscarPropiedades", SQL_Q25)


def vista_sql(con):
    st.subheader("SQL en vivo · escribe o pega una consulta")
    st.caption("La consulta se ejecuta de verdad sobre las hojas del Excel (tablas: CIUDAD, AREA, "
               "DESARROLLADOR, PROYECTO, FUENTE, TIPO_PROPIEDAD, PROPIEDAD, PROPIEDAD_AMENIDAD, "
               "PROPIEDAD_ESTACIONAMIENTO, PROPIEDAD_AUTORIDAD). Acepta sintaxis de SQL Server: "
               "DECLARE, N'texto', DATEADD, TOP, CTE, EXEC de procedimientos, etc.")
    ejemplos = {
        "Q8 · Zonas premium": SQL_Q8,
        "Q13 · Disponibilidad": SQL_Q13,
        "Q25 · Buscador (CREATE + EXEC)": SQL_Q25 + "\n\nEXEC sp_BuscarPropiedades @Ciudad = N'Mumbai', "
                                                  "@Dormitorios = 2, @Disponibilidad = N'Inmediata,Corto plazo';",
        "Mi propia consulta": "SELECT TOP 10 * FROM PROPIEDAD ORDER BY Precio DESC;",
    }
    elegido = st.selectbox("Cargar ejemplo", list(ejemplos), key="sql_ej")
    texto = st.text_area("Consulta SQL", ejemplos[elegido], height=320, key=f"sql_txt_{elegido}")
    if st.button("▶ Ejecutar consulta", type="primary", key="sql_run"):
        try:
            df, traducido = motor_sql.ejecutar(con, texto)
        except Exception as e:                       # errores de sintaxis, tablas inexistentes, etc.
            st.error(f"No se pudo ejecutar la consulta: {e}")
            return
        if df is None:
            st.success("Procedimiento registrado. Añade una línea EXEC para ejecutarlo.")
            return
        st.success(f"{len(df):,} fila(s) devueltas")
        mostrar_df(df, hide_index=True)
        st.download_button("⬇️ Descargar resultado (CSV)", df.to_csv(index=False).encode("utf-8-sig"),
                           "resultado_sql.csv", "text/csv")
        with st.expander("Ver SQL ejecutado (traducido a SQLite)"):
            st.code(traducido, language="sql")


# ----------------------------------------------------------------------------
# Programa principal
# ----------------------------------------------------------------------------
def main():
    st.title("🏠 BD Inmobiliaria · Consultas y buscador")

    with st.sidebar:
        st.header("Origen de datos")
        subido = st.file_uploader("Subir Excel de la BD (opcional)", type=["xlsx"])
        ruta = st.text_input("o ruta del archivo (vacío = detectar automáticamente)", "")
        st.caption("Si subes un archivo, se usa ese; si no, se busca el Excel en el repositorio.")

    if subido is not None:
        origen = subido.getvalue()
    else:
        origen, encontrados = buscar_excel(ruta.strip())
        if origen is None:
            st.error("No se encontró ningún archivo **.xlsx** junto a la app.")
            st.write("Súbelo desde la barra lateral o agrégalo al repositorio de GitHub (misma carpeta que app.py).")
            with st.expander("Diagnóstico"):
                st.write("Carpeta de la app:", CARPETA_APP)
                try:
                    st.write("Archivos en esa carpeta:", sorted(os.listdir(CARPETA_APP)))
                except OSError as e:
                    st.write(e)
            st.stop()
        with st.sidebar:
            st.success(f"Usando: {os.path.basename(origen)}")
            if len(encontrados) > 1:
                st.caption("Otros .xlsx detectados: " + ", ".join(os.path.basename(f) for f in encontrados if f != origen))

    ciudad, area, tipo, prop = cargar_datos(origen)
    base = construir_base(ciudad, area, tipo, prop)
    con = cargar_conexion(origen)

    with st.sidebar:
        st.markdown("---")
        st.metric("Propiedades en la BD", fmt_num(len(base)))
        st.metric("Ciudades / Zonas", f"{len(ciudad)} / {len(area)}")

    tab8, tab13, tab25, tabsql = st.tabs(["Q8 · Zonas premium", "Q13 · Disponibilidad", "Q25 · Buscador", "SQL en vivo"])
    with tab8:
        vista_q8(base)
    with tab13:
        vista_q13(base)
    with tab25:
        vista_q25(base, con)
    with tabsql:
        vista_sql(con)


if __name__ == "__main__":
    main()
