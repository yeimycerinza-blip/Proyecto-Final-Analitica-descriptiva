"""
Proyecto Final - Analítica Descriptiva con IA
Dashboard de Ventas Globales 2024-2026 (Streamlit + Plotly)

Ejecutar:   streamlit run app.py
Requisitos: pip install streamlit pandas plotly openpyxl

Notas de método
- Todos los montos están en USD (Venta_Neta_USD, Utilidad_USD) para que las regiones sean comparables.
- Margen = Utilidad / Venta neta, calculado de forma ponderada (suma de utilidad / suma de ventas),
  nunca como promedio simple de Margen_Porc.
- El archivo tiene una fila por línea de pedido; "pedidos" cuenta ID_Pedido únicos.
- 2026 es un año parcial (datos hasta la última fecha del archivo): las comparaciones anuales
  se hacen también sobre el mismo período de cada año.
"""
import calendar
import re

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from pathlib import Path
from plotly.subplots import make_subplots

# ----------------------------------------------------------------------
# 1. CONFIGURACIÓN
# ----------------------------------------------------------------------
st.set_page_config(page_title="Ventas Globales 2024-2026", page_icon="📊", layout="wide")

ARCHIVO = "ventas_globales_2024_2026.xlsx"
HOJA = "Ventas"

REQUERIDAS = [
    "ID_Venta", "ID_Pedido", "ID_Cliente", "Fecha_Venta", "Año", "Región", "País", "Ciudad",
    "Tipo_Cliente", "Segmento_Cliente", "Industria", "Producto", "Categoría", "Subcategoría",
    "Cantidad", "Descuento_Porc", "Descuento_Local", "Venta_Bruta_Local", "Tasa_Cambio_USD",
    "Venta_Neta_USD", "Costo_Total_USD", "Utilidad_USD", "Margen_Porc", "Canal_Venta",
    "Vendedor", "Equipo_Comercial", "Estado_Pedido", "Dias_Entrega",
]

PREGUNTAS = [
    "¿Cómo ha evolucionado el desempeño comercial de la empresa en términos de ventas, utilidad y margen durante el periodo analizado?",
    "¿Qué regiones, países y ciudades generan mayores ventas y cuáles presentan mayor rentabilidad?",
    "¿Qué categorías, subcategorías y productos explican el desempeño global del negocio?",
    "¿Existe relación entre mayores niveles de ventas y una mayor rentabilidad? ¿Dónde se presentan altos niveles de ventas con márgenes bajos?",
    "¿Qué tipos de clientes, segmentos e industrias representan mayor valor para la empresa?",
    "¿Qué canales de venta presentan el mejor comportamiento comercial y financiero?",
    "¿Qué efecto tienen los descuentos sobre las ventas, la utilidad y el margen de rentabilidad?",
    "¿Cómo se compara el desempeño de los vendedores y de los equipos comerciales?",
    "¿Qué tan concentradas están las ventas en determinados clientes y cuáles presentan mayor recurrencia y valor económico?",
    "¿Dónde existen oportunidades de mejora o problemas en la operación de entrega de los pedidos?",
]
NOMBRES_TABS = [
    "1 · Evolución", "2 · Geografía", "3 · Portafolio", "4 · Ventas vs margen", "5 · Clientes",
    "6 · Canales", "7 · Descuentos", "8 · Vendedores", "9 · Concentración", "10 · Entregas",
]

MESES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
C1, C2, C3, C4 = "#6c7bff", "#2ed3a3", "#ffb454", "#ff6b81"
PALETA = px.colors.qualitative.Vivid
PLANTILLA = "plotly_dark"
ESCALA_MARGEN = "RdYlGn"
DESC_LABELS = ["0–5%", "5–10%", "10–15%", "15–20%", "20% o más"]

# ----------------------------------------------------------------------
# 2. ESTILO
# ----------------------------------------------------------------------
st.markdown(
    """
    <style>
        .block-container {padding-top: 1.4rem; padding-bottom: 3rem; max-width: 1500px;}
        div[data-testid="stMetric"] {
            background: linear-gradient(135deg, #161d2f 0%, #1f2a44 100%);
            border: 1px solid #2c3a5e; border-radius: 12px; padding: 12px 16px;
        }
        div[data-testid="stMetricLabel"] p {color: #9fb0d6; font-size: 0.85rem;}
        .hallazgo {
            background: #141b2b; border-left: 4px solid #6c7bff; border-radius: 8px;
            padding: 12px 16px; margin: 10px 0 4px 0; color: #dfe6f7;
            line-height: 1.65; font-size: 0.95rem;
        }
        .hz-titulo {
            font-weight: 700; color: #8fa2ff; margin-bottom: 4px;
            letter-spacing: .5px; text-transform: uppercase; font-size: .78rem;
        }
        .pregunta {font-size: 1.15rem; font-weight: 600; color: #e8edff; margin: .2rem 0 .9rem 0;}
        button[data-baseweb="tab"] {font-weight: 600;}
    </style>
    """,
    unsafe_allow_html=True,
)


# ----------------------------------------------------------------------
# 3. UTILIDADES
# ----------------------------------------------------------------------
def _version(v: str) -> tuple:
    return tuple(int(x) for x in re.findall(r"\d+", v)[:2])


# Streamlit >= 1.50 usa width="stretch"; las anteriores usan use_container_width=True
ANCHO = {"width": "stretch"} if _version(st.__version__) >= (1, 50) else {"use_container_width": True}


def usd(v: float) -> str:
    a = abs(v)
    if a >= 1e6:
        return f"${v / 1e6:,.1f}M"
    if a >= 1e3:
        return f"${v / 1e3:,.1f}K"
    return f"${v:,.0f}"


def pct(v: float, d: int = 1) -> str:
    return f"{v * 100:.{d}f}%"


def pp(v: float) -> str:
    return f"{v * 100:+.1f} pp"


def mes_txt(ts) -> str:
    return f"{MESES[ts.month - 1]} {ts.year}"


def fuerza(r: float) -> str:
    a = abs(r)
    if a < 0.1:
        return "prácticamente nula"
    adj = "débil" if a < 0.3 else "moderada" if a < 0.5 else "fuerte"
    return f"{'positiva' if r > 0 else 'negativa'} {adj}"


def mostrar(fig, alto: int = 420):
    fig.update_layout(
        template=PLANTILLA, height=alto,
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=10, r=10, t=55, b=10), font=dict(size=13),
        legend=dict(orientation="h", yanchor="top", y=-0.18, x=0),
    )
    st.plotly_chart(fig, **ANCHO)


def tabla(d: pd.DataFrame, alto=None):
    cfg = {}
    for c in d.columns:
        if pd.api.types.is_numeric_dtype(d[c]):
            if "(M USD)" in c:
                f = "%.2f"
            elif "(%)" in c or "(pp)" in c or "días" in c.lower():
                f = "%.1f"
            elif "USD" in c:
                f = "%.0f"
            elif pd.api.types.is_integer_dtype(d[c]):
                f = "%d"
            else:
                f = "%.2f"
            cfg[c] = st.column_config.NumberColumn(c, format=f)
    kw = dict(hide_index=True, column_config=cfg)
    if alto:
        kw["height"] = alto
    st.dataframe(d, **ANCHO, **kw)


def hallazgo(*lineas):
    cuerpo = "<br>".join(
        "• " + re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", l).replace("$", "&#36;") for l in lineas
    )
    st.markdown(
        f'<div class="hallazgo"><div class="hz-titulo">Hallazgos</div>{cuerpo}</div>',
        unsafe_allow_html=True,
    )


def resumen(d: pd.DataFrame, por) -> pd.DataFrame:
    """Agregación estándar con margen ponderado."""
    g = (
        d.groupby(por, observed=True)
        .agg(
            ventas=("Venta_Neta_USD", "sum"),
            utilidad=("Utilidad_USD", "sum"),
            lineas=("ID_Venta", "count"),
            pedidos=("ID_Pedido", "nunique"),
            clientes=("ID_Cliente", "nunique"),
            unidades=("Cantidad", "sum"),
            descuento=("Descuento_USD", "sum"),
            bruta=("Venta_Bruta_USD", "sum"),
        )
        .reset_index()
    )
    g["margen"] = g["utilidad"] / g["ventas"]
    g["ticket"] = g["ventas"] / g["pedidos"]
    g["desc_pct"] = g["descuento"] / g["bruta"]
    return g


def construir_pedidos(d: pd.DataFrame) -> pd.DataFrame:
    """Una fila por pedido (Dias_Entrega, estado, canal, etc. son constantes dentro del pedido)."""
    return (
        d.groupby("ID_Pedido")
        .agg(
            fecha=("Fecha_Venta", "first"), region=("Región", "first"), pais=("País", "first"),
            ciudad=("Ciudad", "first"), canal=("Canal_Venta", "first"),
            estado=("Estado_Pedido", "first"), dias=("Dias_Entrega", "first"),
            ventas=("Venta_Neta_USD", "sum"), utilidad=("Utilidad_USD", "sum"),
        )
        .reset_index()
    )


# ----------------------------------------------------------------------
# 4. CARGA DE DATOS
# ----------------------------------------------------------------------
@st.cache_data(show_spinner="Leyendo el archivo de Excel...")
def cargar_datos(origen, hoja) -> pd.DataFrame:
    df = pd.read_excel(origen, sheet_name=hoja, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    faltan = [c for c in REQUERIDAS if c not in df.columns]
    if faltan:
        raise ValueError("Faltan columnas en el Excel: " + ", ".join(faltan))
    df["Fecha_Venta"] = pd.to_datetime(df["Fecha_Venta"])
    df["Mes_Fecha"] = df["Fecha_Venta"].dt.to_period("M").dt.to_timestamp()
    # Valores en USD del descuento y de la venta bruta (Tasa_Cambio_USD = moneda local por 1 USD)
    df["Descuento_USD"] = df["Descuento_Local"] / df["Tasa_Cambio_USD"]
    df["Venta_Bruta_USD"] = df["Venta_Bruta_Local"] / df["Tasa_Cambio_USD"]
    return df


st.title("📊 Dashboard de Ventas Globales 2024–2026")
st.caption("Proyecto Final · Analítica Descriptiva con IA · Montos en USD")

origen = ARCHIVO if Path(ARCHIVO).exists() else None
if origen is None:
    st.sidebar.warning(f"No encontré `{ARCHIVO}` junto a app.py.")
    origen = st.sidebar.file_uploader("Sube el archivo Excel", type=["xlsx"])
if origen is None:
    st.info("Coloca el Excel en la misma carpeta que `app.py` o súbelo desde la barra lateral.")
    st.stop()

try:
    df = cargar_datos(origen, HOJA)
except Exception as e:  # noqa: BLE001
    st.error(str(e))
    st.stop()

corte = df["Fecha_Venta"].max()  # última fecha disponible en todo el archivo
dias_mes_corte = calendar.monthrange(corte.year, corte.month)[1]
mes_parcial = pd.Timestamp(corte.year, corte.month, 1) if corte.day < dias_mes_corte else None
anio_parcial = corte.year if (corte.month, corte.day) != (12, 31) else None

# ----------------------------------------------------------------------
# 5. FILTROS (BARRA LATERAL)
# ----------------------------------------------------------------------
st.sidebar.header("🎛️ Filtros")


def filtro(titulo, columna):
    opciones = sorted(df[columna].dropna().unique().tolist())
    return st.sidebar.multiselect(titulo, opciones, default=opciones, key=f"f_{columna}")


f_anio = filtro("Año", "Año")
f_region = filtro("Región", "Región")
f_cat = filtro("Categoría", "Categoría")
f_canal = filtro("Canal de venta", "Canal_Venta")

dff = df[
    df["Año"].isin(f_anio)
    & df["Región"].isin(f_region)
    & df["Categoría"].isin(f_cat)
    & df["Canal_Venta"].isin(f_canal)
].copy()

st.sidebar.markdown("---")
st.sidebar.caption(f"{len(dff):,} de {len(df):,} líneas de pedido")
st.sidebar.caption(f"Datos hasta: {corte:%d/%m/%Y}")

if dff.empty:
    st.warning("No hay datos con los filtros seleccionados. Amplía la selección en la barra lateral.")
    st.stop()

# ----------------------------------------------------------------------
# 6. KPIs GLOBALES
# ----------------------------------------------------------------------
ped = construir_pedidos(dff)
tot_v = dff["Venta_Neta_USD"].sum()
tot_u = dff["Utilidad_USD"].sum()
margen_global = tot_u / tot_v

k = st.columns(6)
k[0].metric("Ventas netas", usd(tot_v))
k[1].metric("Utilidad", usd(tot_u))
k[2].metric("Margen", pct(margen_global))
k[3].metric("Pedidos", f"{len(ped):,}")
k[4].metric("Clientes activos", f"{dff['ID_Cliente'].nunique():,}")
k[5].metric("Entrega promedio", f"{ped['dias'].mean():.1f} días")
st.markdown("")


# ----------------------------------------------------------------------
# 7. PREGUNTAS
# ----------------------------------------------------------------------
def tab1():
    m = resumen(dff, "Mes_Fecha").sort_values("Mes_Fecha")

    fig = go.Figure()
    fig.add_bar(x=m["Mes_Fecha"], y=m["ventas"], name="Ventas", marker_color=C1, opacity=0.85)
    fig.add_scatter(x=m["Mes_Fecha"], y=m["utilidad"], name="Utilidad", mode="lines+markers",
                    line=dict(color=C2, width=3))
    fig.update_layout(title="Ventas y utilidad mensuales (USD)")
    fig.update_yaxes(tickprefix="$", tickformat=".2s")
    mostrar(fig, 400)
    if mes_parcial is not None:
        st.caption(f"{mes_txt(mes_parcial)} es un mes parcial (datos hasta el {corte:%d/%m/%Y}).")

    a, b = st.columns(2)
    with a:
        f2 = px.line(m, x="Mes_Fecha", y="margen", markers=True, title="Margen mensual",
                     color_discrete_sequence=[C3])
        f2.add_hline(y=margen_global, line_dash="dash", line_color="#9aa4c7",
                     annotation_text=f"Global {pct(margen_global)}", annotation_position="top left")
        f2.update_yaxes(tickformat=".0%")
        mostrar(f2, 380)

    # Comparación sobre el mismo período de cada año (evita comparar un año parcial con uno completo)
    filas = []
    for y in sorted(dff["Año"].unique()):
        y = int(y)
        fin = pd.Timestamp(y, corte.month, min(corte.day, calendar.monthrange(y, corte.month)[1]))
        sub = dff[(dff["Fecha_Venta"] >= pd.Timestamp(y, 1, 1)) & (dff["Fecha_Venta"] <= fin)]
        filas.append(dict(anio=str(y), ventas=sub["Venta_Neta_USD"].sum(), utilidad=sub["Utilidad_USD"].sum()))
    ytd = pd.DataFrame(filas)
    ytd["margen"] = ytd["utilidad"] / ytd["ventas"]
    ytd["yoy"] = ytd["ventas"].pct_change()
    with b:
        f3 = go.Figure()
        f3.add_bar(x=ytd["anio"], y=ytd["ventas"], name="Ventas", marker_color=C1,
                   text=["" if pd.isna(v) else f"{v:+.1%}" for v in ytd["yoy"]], textposition="outside")
        f3.add_bar(x=ytd["anio"], y=ytd["utilidad"], name="Utilidad", marker_color=C2)
        f3.update_layout(barmode="group",
                         title=f"Mismo período en cada año (1 ene – {corte.day} {MESES[corte.month - 1]}) · % = crec. de ventas")
        f3.update_yaxes(tickprefix="$", tickformat=".2s")
        mostrar(f3, 380)

    an = resumen(dff, "Año").sort_values("Año")
    anios = an["Año"].astype(int)
    etiqueta = anios.astype(str) + np.where(anios == (anio_parcial or -1), " (parcial)", "")
    tabla(pd.DataFrame({
        "Año": etiqueta.values,
        "Ventas (M USD)": an["ventas"].values / 1e6,
        "Utilidad (M USD)": an["utilidad"].values / 1e6,
        "Margen (%)": an["margen"].values * 100,
        "Pedidos": an["pedidos"].values,
        "Clientes": an["clientes"].values,
    }))

    pico = m.loc[m["ventas"].idxmax()]
    mc = m[m["Mes_Fecha"] != mes_parcial] if mes_parcial is not None else m
    mc = mc if len(mc) else m
    valle = mc.loc[mc["ventas"].idxmin()]
    mmax, mmin = m.loc[m["margen"].idxmax()], m.loc[m["margen"].idxmin()]
    hz = [f"En el período filtrado se vendieron **{usd(tot_v)}** con utilidad de **{usd(tot_u)}** y margen global de **{pct(margen_global)}**."]
    if len(ytd) >= 2:
        u, p = ytd.iloc[-1], ytd.iloc[-2]
        hz.append(
            f"En el mismo período, las ventas de **{u['anio']}** cambian **{u['yoy']:+.1%}** frente a {p['anio']}, "
            f"y el margen pasa de {pct(p['margen'])} a {pct(u['margen'])} ({pp(u['margen'] - p['margen'])})."
        )
    hz.append(f"El mes de mayores ventas fue **{mes_txt(pico['Mes_Fecha'])}** ({usd(pico['ventas'])}); el más bajo entre meses completos, **{mes_txt(valle['Mes_Fecha'])}** ({usd(valle['ventas'])}).")
    hz.append(f"El margen mensual osciló entre **{pct(mmin['margen'])}** ({mes_txt(mmin['Mes_Fecha'])}) y **{pct(mmax['margen'])}** ({mes_txt(mmax['Mes_Fecha'])}).")
    hallazgo(*hz)


def tab2():
    c0, c1 = st.columns([2, 1])
    nivel = c0.radio("Nivel geográfico", ["Región", "País", "Ciudad"], horizontal=True, key="geo_nivel")
    n = c1.slider("Mostrar top", 3, 20, 10, key="geo_n")
    g = resumen(dff, nivel)
    rango = [g["margen"].min(), g["margen"].max()]

    izq, der = st.columns(2)
    with izq:
        a = g.nlargest(n, "ventas").sort_values("ventas")
        f1 = px.bar(a, x="ventas", y=nivel, orientation="h", color="margen",
                    color_continuous_scale=ESCALA_MARGEN, range_color=rango,
                    text=a["ventas"].map(usd), title=f"Top {len(a)} por ventas (color = margen)",
                    hover_data={"margen": ":.1%", "utilidad": ":,.0f"})
        f1.update_traces(textposition="outside", cliponaxis=False)
        f1.update_xaxes(tickprefix="$", tickformat=".2s")
        f1.update_coloraxes(colorbar_tickformat=".0%", colorbar_title="Margen")
        mostrar(f1, 460)
    with der:
        b = g.nlargest(n, "margen").sort_values("margen")
        f2 = px.bar(b, x="margen", y=nivel, orientation="h", color="ventas",
                    color_continuous_scale="Viridis", text=b["margen"].map(pct),
                    title=f"Top {len(b)} por rentabilidad (margen %)",
                    hover_data={"ventas": ":,.0f", "utilidad": ":,.0f"})
        f2.update_traces(textposition="outside", cliponaxis=False)
        f2.update_xaxes(tickformat=".0%")
        f2.update_coloraxes(colorbar_tickformat=".2s", colorbar_title="Ventas")
        mostrar(f2, 460)

    t = resumen(dff, ["Región", "País", "Ciudad"])
    f3 = px.treemap(t, path=[px.Constant("Mundo"), "Región", "País", "Ciudad"], values="ventas",
                    color="margen", color_continuous_scale=ESCALA_MARGEN,
                    range_color=[t["margen"].min(), t["margen"].max()],
                    title="Mapa de ventas (tamaño) y margen (color): región → país → ciudad")
    f3.update_coloraxes(colorbar_tickformat=".0%", colorbar_title="Margen")
    mostrar(f3, 520)

    gs = g.sort_values("ventas", ascending=False)
    lider = gs.iloc[0]
    mejor = g.sort_values("margen", ascending=False).iloc[0]
    hz = [
        f"**{lider[nivel]}** lidera en ventas con {usd(lider['ventas'])} ({pct(lider['ventas'] / tot_v)} del total) y margen de {pct(lider['margen'])}.",
        f"La mayor rentabilidad corresponde a **{mejor[nivel]}**, con margen de {pct(mejor['margen'])} sobre ventas de {usd(mejor['ventas'])}.",
    ]
    top5 = gs.head(5)
    bajos = top5[top5["margen"] < margen_global]
    if len(bajos):
        hz.append("Entre los 5 mayores en ventas, rentan por debajo del margen global (" + pct(margen_global) + "): "
                  + ", ".join(f"**{x}** ({pct(m_)})" for x, m_ in zip(bajos[nivel], bajos["margen"])) + ".")
    hallazgo(*hz)


def tab3():
    c = resumen(dff, "Categoría").sort_values("ventas", ascending=False)
    c["pct_v"] = c["ventas"] / c["ventas"].sum()
    c["pct_u"] = c["utilidad"] / c["utilidad"].sum()

    izq, der = st.columns(2)
    with izq:
        f1 = go.Figure()
        f1.add_bar(x=c["Categoría"], y=c["pct_v"], name="% de ventas", marker_color=C1)
        f1.add_bar(x=c["Categoría"], y=c["pct_u"], name="% de utilidad", marker_color=C2)
        f1.update_layout(barmode="group", title="Aporte de cada categoría a ventas vs. utilidad")
        f1.update_yaxes(tickformat=".0%")
        mostrar(f1, 420)
    with der:
        s = resumen(dff, ["Categoría", "Subcategoría"])
        f2 = px.sunburst(s, path=["Categoría", "Subcategoría"], values="utilidad", color="margen",
                         color_continuous_scale=ESCALA_MARGEN, range_color=[s["margen"].min(), s["margen"].max()],
                         title="Utilidad (tamaño) y margen (color): categoría → subcategoría")
        f2.update_coloraxes(colorbar_tickformat=".0%", colorbar_title="Margen")
        mostrar(f2, 420)

    t = resumen(dff, ["Categoría", "Subcategoría", "Producto"])
    f3 = px.treemap(t, path=[px.Constant("Portafolio"), "Categoría", "Subcategoría", "Producto"], values="ventas",
                    color="margen", color_continuous_scale=ESCALA_MARGEN,
                    range_color=[t["margen"].min(), t["margen"].max()],
                    title="Portafolio: ventas (tamaño) y margen (color)")
    f3.update_coloraxes(colorbar_tickformat=".0%", colorbar_title="Margen")
    mostrar(f3, 520)

    base = st.radio("Pareto de productos según", ["Ventas", "Utilidad"], horizontal=True, key="par_base")
    col = "ventas" if base == "Ventas" else "utilidad"
    p = resumen(dff, "Producto").sort_values(col, ascending=False).reset_index(drop=True)
    p["acum"] = p[col].cumsum() / p[col].sum()
    f4 = make_subplots(specs=[[{"secondary_y": True}]])
    f4.add_bar(x=p["Producto"], y=p[col], name=base, marker_color=C1, secondary_y=False)
    f4.add_scatter(x=p["Producto"], y=p["acum"], name="% acumulado", mode="lines+markers",
                   line=dict(color=C3, width=3), secondary_y=True)
    f4.add_scatter(x=p["Producto"], y=[0.8] * len(p), name="80%", mode="lines",
                   line=dict(color="#9aa4c7", dash="dot"), secondary_y=True)
    f4.update_yaxes(tickprefix="$", tickformat=".2s", secondary_y=False)
    f4.update_yaxes(tickformat=".0%", range=[0, 1.05], secondary_y=True)
    f4.update_layout(title=f"Pareto de productos por {base.lower()}")
    mostrar(f4, 440)

    n80 = int((p["acum"] < 0.8).sum() + 1)
    top_c = c.iloc[0]
    sobre = c.assign(dif=c["pct_u"] - c["pct_v"]).sort_values("dif", ascending=False)
    a_, b_ = sobre.iloc[0], sobre.iloc[-1]
    sc = resumen(dff, "Subcategoría").sort_values("ventas", ascending=False).iloc[0]
    hallazgo(
        f"**{top_c['Categoría']}** es la categoría con más ventas ({pct(top_c['pct_v'])} del total) y aporta {pct(top_c['pct_u'])} de la utilidad.",
        f"**{a_['Categoría']}** aporta {pct(a_['pct_u'])} de la utilidad con solo {pct(a_['pct_v'])} de las ventas (la que más sobre-contribuye); "
        f"**{b_['Categoría']}** es la que más sub-contribuye ({pct(b_['pct_u'])} de utilidad vs. {pct(b_['pct_v'])} de ventas).",
        f"La subcategoría líder en ventas es **{sc['Subcategoría']}** ({usd(sc['ventas'])}).",
        f"**{n80} de {len(p)} productos** concentran el 80% de {base.lower()}.",
    )


def tab4():
    c0, c1 = st.columns([3, 1])
    nivel = c0.radio("Nivel de análisis", ["Producto", "Subcategoría", "Categoría", "País", "Ciudad"],
                     horizontal=True, key="rel_nivel")
    logx = c1.checkbox("Escala logarítmica en ventas", value=True, key="rel_log")
    color = {"Producto": "Categoría", "Subcategoría": "Categoría", "Categoría": None,
             "País": "Región", "Ciudad": "Región"}[nivel]
    g = resumen(dff, [nivel] + ([color] if color else []))
    med = g["ventas"].median()
    g["zona"] = np.where((g["ventas"] >= med) & (g["margen"] < margen_global), "Altas ventas, margen bajo", "Resto")

    fig = px.scatter(g, x="ventas", y="margen", size="utilidad", color=color, hover_name=nivel,
                     log_x=logx and len(g) > 1, size_max=45, color_discrete_sequence=PALETA,
                     hover_data={"ventas": ":,.0f", "margen": ":.1%", "utilidad": ":,.0f"},
                     title=f"Ventas vs. margen por {nivel.lower()} (tamaño = utilidad)")
    fig.add_hline(y=margen_global, line_dash="dash", line_color="#9aa4c7",
                  annotation_text=f"Margen global {pct(margen_global)}", annotation_position="top left")
    fig.add_vline(x=med, line_dash="dot", line_color="#9aa4c7",
                  annotation_text="Mediana de ventas", annotation_position="top right")
    fig.update_xaxes(tickprefix="$", tickformat=".2s")
    fig.update_yaxes(tickformat=".0%")
    mostrar(fig, 520)

    izq, der = st.columns(2)
    dd = None
    with izq:
        q = dff[["Venta_Neta_USD", "Utilidad_USD"]].copy()
        if len(q) >= 100:
            q["decil"] = pd.qcut(q["Venta_Neta_USD"].rank(method="first"), 10,
                                 labels=[f"D{i}" for i in range(1, 11)])
            dd = q.groupby("decil", observed=True).agg(v=("Venta_Neta_USD", "sum"), u=("Utilidad_USD", "sum")).reset_index()
            dd["margen"] = dd["u"] / dd["v"]
            f2 = px.bar(dd, x="decil", y="margen", text=dd["margen"].map(pct), color="margen",
                        color_continuous_scale=ESCALA_MARGEN,
                        title="Margen según el tamaño de la línea de venta (D1 = más pequeñas · D10 = más grandes)")
            f2.update_traces(textposition="outside", cliponaxis=False)
            f2.update_yaxes(tickformat=".0%")
            f2.update_coloraxes(showscale=False)
            mostrar(f2, 420)
        else:
            st.info("Muy pocas líneas para calcular deciles con los filtros actuales.")
    prob = g[g["zona"] != "Resto"].sort_values("ventas", ascending=False)
    with der:
        st.markdown(f"**{nivel}s con ventas altas (≥ mediana) y margen bajo (< global)**")
        if len(prob):
            cols = [nivel] + ([color] if color else [])
            t = prob[cols].copy()
            t["Ventas (M USD)"] = prob["ventas"] / 1e6
            t["Margen (%)"] = prob["margen"] * 100
            t["Brecha vs. global (pp)"] = (prob["margen"] - margen_global) * 100
            t["Utilidad (M USD)"] = prob["utilidad"] / 1e6
            tabla(t, alto=380)
        else:
            st.success("Ningún elemento combina ventas altas con margen bajo en este nivel.")

    r_lin = dff["Venta_Neta_USD"].corr(dff["Margen_Porc"], method="spearman")
    r_ent = g["ventas"].corr(g["margen"], method="spearman") if len(g) > 2 else np.nan
    hz = [
        f"A nivel de línea de venta, la correlación de Spearman entre tamaño de la venta y margen es **{r_lin:.2f}** ({fuerza(r_lin)}).",
    ]
    if not np.isnan(r_ent):
        hz.append(f"Por {nivel.lower()}, la relación entre ventas y margen es **{r_ent:.2f}** ({fuerza(r_ent)}).")
    if dd is not None:
        hz.append(f"Las líneas más grandes (D10) rinden un margen de **{pct(dd.iloc[-1]['margen'])}** frente a **{pct(dd.iloc[0]['margen'])}** en las más pequeñas (D1) ({pp(dd.iloc[-1]['margen'] - dd.iloc[0]['margen'])}).")
    if len(prob):
        top = prob.iloc[0]
        hz.append(f"Hay **{len(prob)}** {nivel.lower()}(s) con ventas altas y margen bajo; el mayor es **{top[nivel]}** con {usd(top['ventas'])} y margen de {pct(top['margen'])}.")
    hallazgo(*hz)


def tab5():
    dims = {"Tipo de cliente": "Tipo_Cliente", "Segmento": "Segmento_Cliente", "Industria": "Industria"}
    et = st.radio("Dimensión de cliente", list(dims), horizontal=True, key="cli_dim")
    col = dims[et]
    g = resumen(dff, col).sort_values("ventas")
    g["valor_cliente"] = g["ventas"] / g["clientes"]

    izq, der = st.columns(2)
    with izq:
        f1 = go.Figure()
        f1.add_bar(y=g[col], x=g["ventas"], name="Ventas", orientation="h", marker_color=C1)
        f1.add_bar(y=g[col], x=g["utilidad"], name="Utilidad", orientation="h", marker_color=C2)
        f1.update_layout(barmode="group", title=f"Ventas y utilidad por {et.lower()}")
        f1.update_xaxes(tickprefix="$", tickformat=".2s")
        mostrar(f1, max(380, 26 * len(g) * 2 + 120))
    with der:
        f2 = px.bar(g, x="valor_cliente", y=col, orientation="h", color="margen",
                    color_continuous_scale=ESCALA_MARGEN, text=g["valor_cliente"].map(usd),
                    title="Ventas por cliente activo (color = margen)",
                    hover_data={"clientes": ":,", "margen": ":.1%"})
        f2.update_traces(textposition="outside", cliponaxis=False)
        f2.update_xaxes(tickprefix="$", tickformat=".2s")
        f2.update_coloraxes(colorbar_tickformat=".0%", colorbar_title="Margen")
        mostrar(f2, max(380, 26 * len(g) * 2 + 120))

    pv = dff.pivot_table(index="Industria", columns="Segmento_Cliente", values="Utilidad_USD",
                         aggfunc="sum", fill_value=0) / 1e6
    f3 = px.imshow(pv, text_auto=".1f", aspect="auto", color_continuous_scale="Blues",
                   title="Utilidad (M USD): industria × segmento")
    f3.update_coloraxes(colorbar_title="M USD")
    mostrar(f3, 560)

    gs = g.sort_values("ventas", ascending=False)
    top = gs.iloc[0]
    mm = gs.sort_values("margen", ascending=False).iloc[0]
    vc = gs.sort_values("valor_cliente", ascending=False).iloc[0]
    hallazgo(
        f"Por {et.lower()}, **{top[col]}** aporta más ventas: {usd(top['ventas'])} ({pct(top['ventas'] / tot_v)} del total).",
        f"**{mm[col]}** tiene el mayor margen ({pct(mm['margen'])}).",
        f"**{vc[col]}** deja más ventas por cliente activo ({usd(vc['valor_cliente'])}).",
    )


def tab6():
    c = resumen(dff, "Canal_Venta")
    c["part"] = c["ventas"] / c["ventas"].sum()
    c["dias"] = c["Canal_Venta"].map(ped.groupby("canal")["dias"].mean())
    cs = c.sort_values("ventas", ascending=False)

    tabla(pd.DataFrame({
        "Canal": cs["Canal_Venta"].values,
        "Ventas (M USD)": cs["ventas"].values / 1e6,
        "% de ventas": cs["part"].values * 100,
        "Utilidad (M USD)": cs["utilidad"].values / 1e6,
        "Margen (%)": cs["margen"].values * 100,
        "Pedidos": cs["pedidos"].values,
        "Ticket por pedido (USD)": cs["ticket"].values,
        "Descuento / venta bruta (%)": cs["desc_pct"].values * 100,
        "Días de entrega (prom.)": cs["dias"].values,
    }))

    izq, der = st.columns(2)
    with izq:
        f1 = go.Figure()
        f1.add_bar(x=cs["Canal_Venta"], y=cs["ventas"], name="Ventas", marker_color=C1)
        f1.add_bar(x=cs["Canal_Venta"], y=cs["utilidad"], name="Utilidad", marker_color=C2)
        f1.update_layout(barmode="group", title="Ventas y utilidad por canal")
        f1.update_yaxes(tickprefix="$", tickformat=".2s")
        mostrar(f1, 400)
    with der:
        f2 = px.scatter(c, x="desc_pct", y="margen", size="ventas", color="Canal_Venta", text="Canal_Venta",
                        size_max=55, color_discrete_sequence=PALETA,
                        title="Descuento vs. margen por canal (tamaño = ventas)",
                        hover_data={"ventas": ":,.0f", "desc_pct": ":.1%", "margen": ":.1%"})
        f2.update_traces(textposition="top center")
        f2.update_xaxes(tickformat=".0%", title="Descuento / venta bruta")
        f2.update_yaxes(tickformat=".0%", title="Margen")
        f2.update_layout(showlegend=False)
        mostrar(f2, 400)

    m = resumen(dff, ["Mes_Fecha", "Canal_Venta"]).sort_values("Mes_Fecha")
    f3 = px.line(m, x="Mes_Fecha", y="ventas", color="Canal_Venta", markers=True,
                 color_discrete_sequence=PALETA, title="Evolución mensual de ventas por canal")
    f3.update_yaxes(tickprefix="$", tickformat=".2s")
    mostrar(f3, 400)

    mix = resumen(dff, ["Región", "Canal_Venta"])
    f4 = px.bar(mix, x="Región", y="ventas", color="Canal_Venta",
                color_discrete_sequence=PALETA, title="Mezcla de canales por región (% de ventas)")
    f4.update_layout(barmode="stack", barnorm="percent")
    f4.update_yaxes(ticksuffix="%")
    mostrar(f4, 400)

    lider = cs.iloc[0]
    mm = c.sort_values("margen", ascending=False).iloc[0]
    tk = c.sort_values("ticket", ascending=False).iloc[0]
    rap = c.sort_values("dias").iloc[0]
    hallazgo(
        f"**{lider['Canal_Venta']}** es el canal principal: {usd(lider['ventas'])} ({pct(lider['part'])} de las ventas) con margen de {pct(lider['margen'])}.",
        f"El mayor margen lo tiene **{mm['Canal_Venta']}** ({pct(mm['margen'])}) y el mayor ticket por pedido **{tk['Canal_Venta']}** ({usd(tk['ticket'])}).",
        f"Los canales que más descuentan cedieron hasta {pct(c['desc_pct'].max())} de la venta bruta; los de menor descuento llegan a {pct(c['desc_pct'].min())}.",
        f"**{rap['Canal_Venta']}** es el canal que entrega más rápido ({rap['dias']:.1f} días en promedio).",
    )


def tab7():
    d = dff.copy()
    d["rango"] = pd.cut(d["Descuento_Porc"] * 100, bins=[0, 5, 10, 15, 20, 100], right=False,
                        include_lowest=True, labels=DESC_LABELS)
    g = resumen(d, "rango")
    g["ticket_linea"] = g["ventas"] / g["lineas"]
    g["unid_linea"] = g["unidades"] / g["lineas"]

    r = d["Descuento_Porc"].corr(d["Margen_Porc"], method="spearman")
    k = st.columns(3)
    k[0].metric("Descuento total cedido", usd(d["Descuento_USD"].sum()))
    k[1].metric("Descuento / venta bruta", pct(d["Descuento_USD"].sum() / d["Venta_Bruta_USD"].sum()))
    k[2].metric("Correlación descuento–margen (Spearman)", f"{r:.2f}")

    izq, der = st.columns(2)
    with izq:
        f1 = make_subplots(specs=[[{"secondary_y": True}]])
        f1.add_bar(x=g["rango"].astype(str), y=g["ventas"], name="Ventas", marker_color=C1, secondary_y=False)
        f1.add_scatter(x=g["rango"].astype(str), y=g["margen"], name="Margen", mode="lines+markers+text",
                       text=g["margen"].map(pct), textposition="top center",
                       line=dict(color=C3, width=3), secondary_y=True)
        f1.update_yaxes(tickprefix="$", tickformat=".2s", secondary_y=False)
        f1.update_yaxes(tickformat=".0%", secondary_y=True)
        f1.update_layout(title="Ventas y margen según el rango de descuento")
        mostrar(f1, 420)
    with der:
        f2 = make_subplots(specs=[[{"secondary_y": True}]])
        f2.add_bar(x=g["rango"].astype(str), y=g["ticket_linea"], name="Ticket por línea (USD)",
                   marker_color=C2, secondary_y=False)
        f2.add_scatter(x=g["rango"].astype(str), y=g["unid_linea"], name="Unidades por línea",
                       mode="lines+markers", line=dict(color=C4, width=3), secondary_y=True)
        f2.update_yaxes(tickprefix="$", tickformat=".2s", secondary_y=False)
        f2.update_layout(title="Tamaño de la venta según el rango de descuento")
        mostrar(f2, 420)

    hm = d.groupby(["Categoría", "rango"], observed=True).agg(
        v=("Venta_Neta_USD", "sum"), u=("Utilidad_USD", "sum")).reset_index()
    hm["margen"] = hm["u"] / hm["v"]
    pv = hm.pivot(index="Categoría", columns="rango", values="margen")
    pv.columns = pv.columns.astype(str)
    f3 = px.imshow(pv, text_auto=".0%", aspect="auto", color_continuous_scale=ESCALA_MARGEN,
                   title="Margen por categoría y rango de descuento (controla el efecto del mix de productos)")
    f3.update_coloraxes(colorbar_tickformat=".0%", colorbar_title="Margen")
    mostrar(f3, 440)

    primero, ultimo = g.iloc[0], g.iloc[-1]
    peso = g.loc[g["rango"].isin(DESC_LABELS[2:]), "ventas"].sum() / g["ventas"].sum()
    hz = [
        f"El margen {'baja' if ultimo['margen'] < primero['margen'] else 'sube'} de **{pct(primero['margen'])}** (descuentos {primero['rango']}) a **{pct(ultimo['margen'])}** (descuentos {ultimo['rango']}): {pp(ultimo['margen'] - primero['margen'])}.",
        f"El **{pct(peso)}** de las ventas se cierra con descuentos de 10% o más; en ese tramo el ticket por línea es {usd(g.iloc[2:]['ventas'].sum() / g.iloc[2:]['lineas'].sum())}.",
        f"La relación entre descuento y margen a nivel de línea es {fuerza(r)} (Spearman {r:.2f}). Ojo: los descuentos altos suelen venir con volúmenes y ventas mayores, así que parte del efecto refleja el tamaño del negocio y el mix.",
    ]
    cae = (pv.iloc[:, 0] - pv.iloc[:, -1]).dropna()
    if len(cae):
        if cae.max() > 0:
            hz.append(f"Controlando por categoría, la mayor caída de margen entre el rango más bajo y el más alto de descuento ocurre en **{cae.idxmax()}** ({pp(-cae.max())}).")
    hallazgo(*hz)


def tab8():
    v = resumen(dff, ["Vendedor", "Equipo_Comercial"]).sort_values("ventas")
    eq = resumen(dff, "Equipo_Comercial")
    eq["vendedores"] = eq["Equipo_Comercial"].map(dff.groupby("Equipo_Comercial")["Vendedor"].nunique())
    eq["ventas_vend"] = eq["ventas"] / eq["vendedores"]
    eqs = eq.sort_values("ventas", ascending=False)

    tabla(pd.DataFrame({
        "Equipo": eqs["Equipo_Comercial"].values,
        "Vendedores": eqs["vendedores"].values,
        "Ventas (M USD)": eqs["ventas"].values / 1e6,
        "Utilidad (M USD)": eqs["utilidad"].values / 1e6,
        "Margen (%)": eqs["margen"].values * 100,
        "Ventas por vendedor (M USD)": eqs["ventas_vend"].values / 1e6,
        "Pedidos": eqs["pedidos"].values,
        "Descuento / venta bruta (%)": eqs["desc_pct"].values * 100,
    }))
    st.caption("Cada equipo comercial corresponde a una región, por lo que la comparación entre equipos también refleja el tamaño de cada mercado.")

    izq, der = st.columns(2)
    with izq:
        f1 = px.bar(v, x="ventas", y="Vendedor", orientation="h", color="Equipo_Comercial",
                    color_discrete_sequence=PALETA, title="Ventas por vendedor",
                    hover_data={"margen": ":.1%", "utilidad": ":,.0f"})
        f1.update_xaxes(tickprefix="$", tickformat=".2s")
        mostrar(f1, max(420, 22 * len(v) + 120))
    with der:
        f2 = px.scatter(v, x="ventas", y="margen", size="utilidad", color="Equipo_Comercial",
                        hover_name="Vendedor", size_max=35, color_discrete_sequence=PALETA,
                        title="Ventas vs. margen por vendedor (tamaño = utilidad)",
                        hover_data={"ventas": ":,.0f", "margen": ":.1%", "utilidad": ":,.0f"})
        f2.add_hline(y=margen_global, line_dash="dash", line_color="#9aa4c7",
                     annotation_text=f"Global {pct(margen_global)}", annotation_position="top left")
        f2.update_xaxes(tickprefix="$", tickformat=".2s")
        f2.update_yaxes(tickformat=".0%")
        mostrar(f2, max(420, 22 * len(v) + 120))

    f3 = px.bar(eqs, x="Equipo_Comercial", y="ventas_vend", color="margen",
                color_continuous_scale=ESCALA_MARGEN, text=eqs["ventas_vend"].map(usd),
                title="Productividad: ventas por vendedor en cada equipo (color = margen)")
    f3.update_traces(textposition="outside", cliponaxis=False)
    f3.update_yaxes(tickprefix="$", tickformat=".2s")
    f3.update_coloraxes(colorbar_tickformat=".0%", colorbar_title="Margen")
    mostrar(f3, 400)

    vs = v.sort_values("ventas", ascending=False)
    top = vs.iloc[0]
    mm = vs.sort_values("margen", ascending=False).iloc[0]
    pe = eq.sort_values("ventas_vend", ascending=False).iloc[0]
    hallazgo(
        f"**{top['Vendedor']}** ({top['Equipo_Comercial']}) es el vendedor con más ventas: {usd(top['ventas'])}.",
        f"**{mm['Vendedor']}** logra el mejor margen ({pct(mm['margen'])}); el rango de margen entre vendedores va de {pct(vs['margen'].min())} a {pct(vs['margen'].max())}.",
        f"El equipo más productivo es **{pe['Equipo_Comercial']}** con {usd(pe['ventas_vend'])} por vendedor.",
        f"Los 5 mejores vendedores concentran {pct(vs.head(5)['ventas'].sum() / tot_v)} de las ventas.",
    )


def tab9():
    cl = (
        dff.groupby("ID_Cliente")
        .agg(ventas=("Venta_Neta_USD", "sum"), utilidad=("Utilidad_USD", "sum"),
             pedidos=("ID_Pedido", "nunique"), Segmento=("Segmento_Cliente", "first"),
             Industria=("Industria", "first"), País=("País", "first"))
        .reset_index()
        .sort_values("ventas", ascending=False)
        .reset_index(drop=True)
    )
    cl["margen"] = cl["utilidad"] / cl["ventas"]
    n = len(cl)
    cl["pct_clientes"] = (cl.index + 1) / n
    cl["acum"] = cl["ventas"].cumsum() / cl["ventas"].sum()
    top10 = cl["ventas"].head(10).sum() / tot_v
    i20 = max(int(np.ceil(0.2 * n)) - 1, 0)
    top20 = cl["acum"].iloc[i20]
    n80 = int((cl["acum"] < 0.8).sum() + 1)

    k = st.columns(4)
    k[0].metric("Clientes activos", f"{n:,}")
    k[1].metric("Top 10 clientes", pct(top10), help="Participación en las ventas")
    k[2].metric("20% de clientes", pct(top20), help="Participación en las ventas del 20% de clientes con más ventas")
    k[3].metric("Clientes para el 80%", f"{n80:,}", help=f"{pct(n80 / n)} de la base")

    izq, der = st.columns(2)
    with izq:
        f1 = go.Figure()
        f1.add_scatter(x=cl["pct_clientes"], y=cl["acum"], mode="lines", name="Ventas acumuladas",
                       line=dict(color=C1, width=3), fill="tozeroy", fillcolor="rgba(108,123,255,0.15)")
        f1.add_scatter(x=[0, 1], y=[0, 1], mode="lines", name="Distribución uniforme",
                       line=dict(color="#9aa4c7", dash="dot"))
        f1.add_scatter(x=[cl["pct_clientes"].iloc[i20]], y=[top20], mode="markers+text", name="20% de clientes",
                       marker=dict(color=C3, size=11), text=[f"20% → {pct(top20, 0)}"], textposition="bottom right")
        f1.update_layout(title="Curva de concentración de ventas por cliente")
        f1.update_xaxes(tickformat=".0%", title="% acumulado de clientes (de mayor a menor)")
        f1.update_yaxes(tickformat=".0%", title="% acumulado de ventas")
        mostrar(f1, 430)
    with der:
        cl["frecuencia"] = pd.cut(cl["pedidos"], bins=[0, 1, 3, 6, 10_000],
                                  labels=["1 pedido", "2–3 pedidos", "4–6 pedidos", "7 o más"])
        fr = cl.groupby("frecuencia", observed=True).agg(
            clientes=("ID_Cliente", "count"), ventas=("ventas", "sum"), utilidad=("utilidad", "sum")).reset_index()
        fr["pc"] = fr["clientes"] / fr["clientes"].sum()
        fr["pv"] = fr["ventas"] / fr["ventas"].sum()
        f2 = go.Figure()
        f2.add_bar(x=fr["frecuencia"].astype(str), y=fr["pc"], name="% de clientes", marker_color=C1)
        f2.add_bar(x=fr["frecuencia"].astype(str), y=fr["pv"], name="% de ventas", marker_color=C2)
        f2.update_layout(barmode="group", title="Recurrencia: pedidos por cliente")
        f2.update_yaxes(tickformat=".0%")
        mostrar(f2, 430)

    rng = np.random.default_rng(7)
    cl["pedidos_j"] = cl["pedidos"] + rng.uniform(-0.3, 0.3, size=n)
    f3 = px.scatter(cl, x="pedidos_j", y="ventas", color="Segmento", size="utilidad", size_max=16,
                    log_y=True, opacity=0.6, color_discrete_sequence=PALETA, render_mode="webgl",
                    title="Recurrencia vs. valor: pedidos por cliente y ventas (tamaño = utilidad)",
                    hover_data={"pedidos_j": False, "pedidos": True, "ventas": ":,.0f",
                                "margen": ":.1%", "Industria": True, "País": True, "ID_Cliente": True})
    f3.update_xaxes(title="Pedidos por cliente")
    f3.update_yaxes(tickprefix="$", tickformat=".2s", title="Ventas del cliente")
    mostrar(f3, 480)

    top = cl.head(15)
    st.markdown("**Top 15 clientes por valor económico**")
    tabla(pd.DataFrame({
        "Cliente": top["ID_Cliente"].values, "Segmento": top["Segmento"].values,
        "Industria": top["Industria"].values, "País": top["País"].values,
        "Pedidos": top["pedidos"].values,
        "Ventas (M USD)": top["ventas"].values / 1e6,
        "Utilidad (M USD)": top["utilidad"].values / 1e6,
        "Margen (%)": top["margen"].values * 100,
    }))

    rec = fr.loc[fr["frecuencia"].astype(str).isin(["4–6 pedidos", "7 o más"])]
    lectura = ("concentración moderada, sin dependencia de un cliente puntual" if top10 < 0.10
               else "concentración alta: pocos clientes pesan mucho en las ventas")
    hallazgo(
        f"El top 10 de clientes pesa {pct(top10)} de las ventas y el 20% de la base genera **{pct(top20)}**; se necesitan {n80:,} clientes ({pct(n80 / n)}) para llegar al 80%: {lectura}.",
        f"Los clientes con 4 o más pedidos son el {pct(rec['pc'].sum())} de la base y aportan **{pct(rec['pv'].sum())}** de las ventas.",
        f"El cliente de mayor valor es **{cl.iloc[0]['ID_Cliente']}** ({cl.iloc[0]['Segmento']}, {cl.iloc[0]['Industria']}) con {usd(cl.iloc[0]['ventas'])} en ventas.",
    )


def tab10():
    c1, c2 = st.columns(2)
    umbral = c1.slider("Considerar entrega lenta si supera (días)", 3, 18, 10, key="umb_dias")
    rezago = c2.slider("Pedido sin entregar con rezago crítico desde (días)", 7, 180, 30, key="umb_rez")

    p = ped.copy()
    p["lenta"] = p["dias"] > umbral
    nd = p[p["estado"] != "Entregado"].copy()
    nd["antig"] = (corte - nd["fecha"]).dt.days
    crit = nd[nd["antig"] >= rezago]

    k = st.columns(4)
    k[0].metric("Entrega promedio", f"{p['dias'].mean():.1f} días")
    k[1].metric(f"Entregas de más de {umbral} días", pct(p["lenta"].mean()))
    k[2].metric("Pedidos sin entregar", f"{len(nd):,}", delta=f"{usd(nd['ventas'].sum())} en ventas", delta_color="off")
    k[3].metric(f"Sin entregar con {rezago}+ días", f"{len(crit):,}", delta=f"{usd(crit['ventas'].sum())} en ventas", delta_color="off")

    izq, der = st.columns(2)
    with izq:
        f1 = px.histogram(p, x="dias", color="canal", nbins=int(p["dias"].max()), barmode="stack",
                          color_discrete_sequence=PALETA, title="Distribución de días de entrega por canal")
        f1.add_vline(x=umbral + 0.5, line_dash="dash", line_color=C4, annotation_text=f"Umbral {umbral} días")
        f1.update_xaxes(title="Días de entrega")
        f1.update_yaxes(title="Pedidos")
        mostrar(f1, 400)
    with der:
        gc = p.groupby("canal").agg(dias=("dias", "mean"), lenta=("lenta", "mean"),
                                    pedidos=("dias", "size")).reset_index().sort_values("dias")
        f2 = px.bar(gc, x="dias", y="canal", orientation="h", color="lenta", color_continuous_scale="Reds",
                    text=gc["dias"].map(lambda v: f"{v:.1f}"), title="Días promedio por canal (color = % entregas lentas)",
                    hover_data={"pedidos": ":,", "lenta": ":.1%"})
        f2.update_traces(textposition="outside", cliponaxis=False)
        f2.update_coloraxes(colorbar_tickformat=".0%", colorbar_title="% lentas")
        mostrar(f2, 400)

    izq, der = st.columns(2)
    with izq:
        gp = p.groupby("pais").agg(dias=("dias", "mean"), lenta=("lenta", "mean"),
                                   pedidos=("dias", "size")).reset_index()
        gp = gp.nlargest(12, "dias").sort_values("dias")
        f3 = px.bar(gp, x="dias", y="pais", orientation="h", color="lenta", color_continuous_scale="Reds",
                    text=gp["dias"].map(lambda v: f"{v:.2f}"), title="Países con entrega más lenta (top 12)",
                    hover_data={"pedidos": ":,", "lenta": ":.1%"})
        f3.update_traces(textposition="outside", cliponaxis=False)
        f3.update_xaxes(range=[0, max(gp["dias"].max() * 1.15, 1)])
        f3.update_coloraxes(colorbar_tickformat=".0%", colorbar_title="% lentas")
        mostrar(f3, 440)
    with der:
        pv = p.pivot_table(index="region", columns="canal", values="dias", aggfunc="mean")
        f4 = px.imshow(pv, text_auto=".1f", aspect="auto", color_continuous_scale="YlOrRd",
                       title="Días promedio de entrega: región × canal")
        f4.update_coloraxes(colorbar_title="Días")
        mostrar(f4, 440)

    if len(nd):
        nd["tramo"] = pd.cut(nd["antig"], bins=[-1, 30, 90, 180, 100_000],
                             labels=["0–30 días", "31–90 días", "91–180 días", "Más de 180"])
        t = nd.groupby(["tramo", "estado"], observed=True).agg(
            pedidos=("ID_Pedido", "size"), ventas=("ventas", "sum")).reset_index()
        f5 = px.bar(t, x="tramo", y="pedidos", color="estado", barmode="stack",
                    color_discrete_map={"Pendiente": C4, "En tránsito": C3},
                    title="Pedidos sin entregar según su antigüedad desde la venta",
                    hover_data={"ventas": ":,.0f"})
        f5.update_xaxes(title="Antigüedad del pedido")
        mostrar(f5, 400)
    else:
        st.success("No hay pedidos pendientes o en tránsito con los filtros actuales.")

    rap, lent = gc.iloc[0], gc.iloc[-1]
    pais_prom = p.groupby("pais")["dias"].mean()
    rango_canal = gc["dias"].max() - gc["dias"].min()
    rango_pais = pais_prom.max() - pais_prom.min()
    if rango_canal > 2 * rango_pais:
        causa = f"la brecha se explica por el canal ({rango_canal:.1f} días de diferencia) y no por el país ({rango_pais:.1f} días entre el más rápido y el más lento)"
    else:
        causa = f"entre canales hay {rango_canal:.1f} días de diferencia y entre países {rango_pais:.1f}"
    hz = [
        f"El canal más rápido es **{rap['canal']}** ({rap['dias']:.1f} días) y el más lento **{lent['canal']}** ({lent['dias']:.1f} días): {causa}.",
        f"**{pct(p['lenta'].mean())}** de los pedidos tarda más de {umbral} días en entregarse.",
    ]
    if len(nd):
        hz.append(f"Hay **{len(nd):,}** pedidos sin entregar por **{usd(nd['ventas'].sum())}**; **{len(crit):,}** de ellos ({usd(crit['ventas'].sum())}) llevan {rezago} días o más desde la venta, lo que sugiere pedidos estancados o estados sin actualizar.")
    hallazgo(*hz)


# ----------------------------------------------------------------------
# 8. RENDER
# ----------------------------------------------------------------------
FUNCIONES = [tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9, tab10]
tabs = st.tabs(NOMBRES_TABS)
for i, (tab, fn) in enumerate(zip(tabs, FUNCIONES)):
    with tab:
        st.markdown(f'<div class="pregunta">❓ {i + 1}. {PREGUNTAS[i]}</div>', unsafe_allow_html=True)
        fn()

with st.expander("🔎 Ver una muestra de los datos filtrados"):
    st.dataframe(dff.head(500), **ANCHO)
