"""
Da Vinci - Gestión de Abastecimiento (dashboard ejecutivo, valorizado)
    Compras      : SO vs RN   (lo solicitado vs lo recepcionado en almacén)
    Despachos    : RN vs OF   (lo recepcionado vs lo despachado a sitio)
    Confirmación : OF vs RNO  (lo despachado vs lo confirmado por operaciones)

Ejecutar:  streamlit run app.py
"""
from __future__ import annotations

import io
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from plotly.subplots import make_subplots

# =====================================================
# CONFIG
# =====================================================
BASE_DIR = Path(r"D:\FILL RATE")
ARCHIVO = BASE_DIR / "OUTPUT" / "TRAZABILIDAD_BASE.xlsx"
MONEDA = "S/"

# Parámetros fijos (ya no hay barra lateral)
META_COMPRAS = 95        # % meta Fill Rate de compras
META_DESPACHO = 95       # % meta eficiencia de despacho
META_CONFIRMACION = 95   # % meta confirmación en operación
META_PROVEEDORES = 95    # % meta de POs atendidas completas
REPORTES = ["Fill rate - Trazabilidad", "Cumplimiento Proveedores"]
TOPE_LINEA = True        # un exceso en una línea no compensa el faltante de otra
SOLO_VENCIDAS_DEFECTO = True  # valor inicial del casillero "Solo entregas vencidas"

# Paleta tomada del dashboard de referencia
NAVY, RED, STEEL, PINK, GREEN = "#1F5C85", "#E8434B", "#7FA3BA", "#F4898D", "#2E9E6B"
AMBAR, SLATE, DARKRED = "#F2B84B", "#4A5A6A", "#B3262E"
PALETA = [NAVY, RED, STEEL, PINK, "#9AA5B1"]
INK, MUTED = "#1E2B3A", "#6B7A8C"

MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]

st.set_page_config(
    page_title="Da Vinci - Gestión de Abastecimiento", page_icon="📊", layout="wide",
    initial_sidebar_state="collapsed",
)

# ##### INICIO LÓGICA (sin Streamlit) #####
import unicodedata

NECESARIAS = [
    "SO Number", "SO Item Id", "SO Site", "SOE Site", "SO Family Name", "SO Group Name",
    "SO Suppliers", "PO Suppliers", "SO Delivery Date", "SO Creation Date",
    "RN Delivery Date", "OF Date", "RNO Delivery Date", "RNO Number", "RNO Item Id",
    "SO Total Price", "PO Total Price", "RN Total Price", "OF Total Price", "RNO Total Price",
    "SO Quantity", "RN Received Quantity", "OF Quantity", "RNO Received Quantity",
    # Reporte de cumplimiento de proveedores (PO vs RN)
    "PO Number", "PO Item Id", "PO Quantity", "PO Delivery Date", "PO Valid", "RN Number", "RN Is Valid",
    "PO Family Name", "PO Group Name", "PO Item Name", "PO Packing Unit",
]
STOCK_PROPIO = "ALMACEN LIMA"  # proveedor que representa despachos del stock del almacén
MONTOS = ["SO Total Price", "PO Total Price", "RN Total Price", "OF Total Price", "RNO Total Price"]
CANTIDADES = ["SO Quantity", "RN Received Quantity", "OF Quantity", "RNO Received Quantity"]
FECHAS = ["SO Delivery Date", "SO Creation Date", "RN Delivery Date", "OF Date", "RNO Delivery Date"]


def _norm(x) -> str:
    return unicodedata.normalize("NFKD", str(x)).encode("ascii", "ignore").decode().upper()


def es_almacen_como_sitio(serie: pd.Series) -> pd.Series:
    """Sitios que en realidad son los almacenes (p. ej. NEWREST ALMACEN LIMA / AREQUIPA): se eliminan del análisis."""
    mapa = {v: _norm(v) for v in serie.dropna().unique()}
    n = serie.map(mapa).fillna("")
    return n.str.contains("NEWREST", na=False) & n.str.contains("ALMACEN", na=False)


def leer_excel(origen) -> pd.DataFrame:
    # Solo se leen las columnas necesarias: acelera mucho la carga (110 columnas en el archivo)
    return pd.read_excel(origen, usecols=lambda c: str(c).strip() in NECESARIAS)


def consolidar(df: pd.DataFrame, tope_linea: bool = True):
    """
    Una fila por línea (SO Number + SO Item Id).
    Devuelve (lineas, fecha_corte_por_defecto).
    """
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    for c in MONTOS:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0) if c in df.columns else 0.0
    for c in CANTIDADES:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0) if c in df.columns else 0.0
    for c in FECHAS:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")

    # Los almacenes NO se analizan como sitios
    df = df[~es_almacen_como_sitio(df["SO Site"])].copy()

    # Un mismo RNO (RNO Number + RNO Item Id) puede confirmar varios OF de líneas SO distintas, pero la base
    # deja todo su monto en la primera fila. Se reparte entre las filas del RNO en proporción a su OF.
    if "RNO Number" in df.columns and "RNO Item Id" in df.columns:
        clave = [df["RNO Number"], df["RNO Item Id"]]
        tiene = df["RNO Number"].notna() & df["RNO Item Id"].notna()
        for col_rno, col_of in [("RNO Total Price", "OF Total Price"), ("RNO Received Quantity", "OF Quantity")]:
            peso = df[col_of].where(tiene, 0.0)
            suma_peso = peso.groupby(clave).transform("sum")
            total_doc = df[col_rno].where(tiene, 0.0).groupby(clave).transform("sum")
            reparto = total_doc * peso / suma_peso.replace(0, np.nan)
            df[col_rno] = np.where(tiene & (suma_peso > 0), reparto, df[col_rno])
            df[col_rno] = np.nan_to_num(df[col_rno].astype(float))

    # Proveedor: el de la OC; si falta, el del pedido
    if "PO Suppliers" in df.columns and "SO Suppliers" in df.columns:
        df["Proveedor"] = df["PO Suppliers"].fillna(df["SO Suppliers"])
    elif "PO Suppliers" in df.columns:
        df["Proveedor"] = df["PO Suppliers"]
    else:
        df["Proveedor"] = df.get("SO Suppliers", "SIN PROVEEDOR")
    df["Proveedor"] = df["Proveedor"].fillna("SIN PROVEEDOR").astype(str)
    # Despachos del stock propio: suman en los montos, pero no se analizan como un proveedor más
    df["Es_stock"] = df["Proveedor"].str.upper().str.contains(STOCK_PROPIO, na=False)

    # Almacén de origen (SOE Site): LIMA / AREQUIPA
    if "SOE Site" in df.columns:
        alm = df["SOE Site"].fillna("SIN DATO").astype(str).str.upper()
    else:
        alm = pd.Series("SIN DATO", index=df.index)
    df["Almacén"] = alm.str.replace("NEWREST ALMACEN", "", regex=False).str.strip().replace("", "SIN DATO")

    corte = max(
        df[c].max() for c in ["SO Creation Date", "RN Delivery Date", "OF Date", "RNO Delivery Date"]
        if c in df.columns and df[c].notna().any()
    )

    agg = {
        "SO Site": "first", "SO Family Name": "first", "SO Group Name": "first", "Proveedor": "first",
        "Es_stock": "max", "Almacén": "first",
        "SO Delivery Date": "min",
        **{c: "sum" for c in MONTOS},
        **{c: "sum" for c in CANTIDADES},
    }
    g = df.groupby(["SO Number", "SO Item Id"], sort=False).agg(agg).reset_index()
    g = g.rename(columns={
        "SO Number": "Pedido", "SO Site": "Sitio", "SO Family Name": "Familia",
        "SO Group Name": "Grupo", "SO Delivery Date": "Entrega",
        "SO Total Price": "SO", "PO Total Price": "PO", "RN Total Price": "RN",
        "OF Total Price": "OF", "RNO Total Price": "RNO",
        "SO Quantity": "SO_q", "RN Received Quantity": "RN_q",
        "OF Quantity": "OF_q", "RNO Received Quantity": "RNO_q",
    })
    for c in ["Sitio", "Familia", "Grupo"]:
        g[c] = g[c].fillna("SIN DATO").astype(str)

    for c in ["SO", "PO", "RN", "OF", "RNO", "SO_q", "RN_q", "OF_q", "RNO_q"]:
        g[c] = g[c].clip(lower=0)

    # Cumplimiento valorizado por línea (con tope al 100% para que un exceso no tape un faltante)
    if tope_linea:
        g["RN_c"] = np.minimum(g["RN"], g["SO"])
        g["OF_c"] = np.minimum(g["OF"], g["RN"])
        g["RNO_c"] = np.minimum(g["RNO"], g["OF"])
        g["RN_qc"] = np.minimum(g["RN_q"], g["SO_q"])
        g["OF_qc"] = np.minimum(g["OF_q"], g["RN_q"])
        g["RNO_qc"] = np.minimum(g["RNO_q"], g["OF_q"])
    else:
        g["RN_c"] = g["RN"]
        g["OF_c"] = g["OF"]
        g["RNO_c"] = g["RNO"]
        g["RN_qc"], g["OF_qc"], g["RNO_qc"] = g["RN_q"], g["OF_q"], g["RNO_q"]

    g["Mes"] = g["Entrega"].dt.to_period("M").dt.to_timestamp()
    return g, corte


def kpis(g: pd.DataFrame) -> dict:
    so, rn, of, rno = g["SO"].sum(), g["RN"].sum(), g["OF"].sum(), g["RNO"].sum()
    rn_c, of_c, rno_c = g["RN_c"].sum(), g["OF_c"].sum(), g["RNO_c"].sum()
    so_q, rn_q, of_q = g["SO_q"].sum(), g["RN_q"].sum(), g["OF_q"].sum()
    return {
        "so": so, "rn": rn, "of": of, "rno": rno, "rn_c": rn_c, "of_c": of_c, "rno_c": rno_c,
        "fr": rn_c / so if so > 0 else np.nan,
        "ef": of_c / rn if rn > 0 else np.nan,
        "ce": rno_c / of if of > 0 else np.nan,
        "brecha_c": so - rn_c,
        "brecha_d": rn - of_c,
        "brecha_e": of - rno_c,
        "lineas": len(g),
        "pedidos": g["Pedido"].nunique(),
        # Cumplimiento por cantidad: suma de unidades atendidas / suma de unidades pedidas (sin distinguir unidad)
        "cq_c": g["RN_qc"].sum() / so_q if so_q > 0 else np.nan,
        "cq_d": g["OF_qc"].sum() / rn_q if rn_q > 0 else np.nan,
        "cq_e": g["RNO_qc"].sum() / of_q if of_q > 0 else np.nan,
    }


def resumen_por(g: pd.DataFrame, col: str) -> pd.DataFrame:
    d = g.groupby(col).agg(SO=("SO", "sum"), RN=("RN", "sum"), RN_c=("RN_c", "sum"),
                           OF=("OF", "sum"), OF_c=("OF_c", "sum"),
                           RNO=("RNO", "sum"), RNO_c=("RNO_c", "sum")).reset_index()
    d["Fill Rate"] = np.where(d["SO"] > 0, d["RN_c"] / d["SO"], np.nan)
    d["Eficiencia"] = np.where(d["RN"] > 0, d["OF_c"] / d["RN"], np.nan)
    d["Brecha compras"] = d["SO"] - d["RN_c"]
    d["Brecha despacho"] = d["RN"] - d["OF_c"]
    d["Confirmación"] = np.where(d["OF"] > 0, d["RNO_c"] / d["OF"], np.nan)
    d["Brecha confirmación"] = d["OF"] - d["RNO_c"]
    return d


def base_dim(g: pd.DataFrame, col: str) -> pd.DataFrame:
    """El stock propio (ALMACEN LIMA) no se analiza como proveedor."""
    return g[~g["Es_stock"]] if col == "Proveedor" else g


def top_y_otros(d: pd.DataFrame, col: str, val: str, n: int = 4) -> pd.DataFrame:
    d = d[d[val] > 0].sort_values(val, ascending=False)
    if len(d) <= n:
        return d[[col, val]]
    top = d.head(n)[[col, val]]
    otros = pd.DataFrame({col: ["Otros"], val: [d.iloc[n:][val].sum()]})
    return pd.concat([top, otros], ignore_index=True)


def fmt_monto(x: float) -> str:
    if pd.isna(x):
        return "—"
    a = abs(x)
    if a >= 1_000_000:
        return f"{MONEDA} {x / 1_000_000:,.2f} M"
    if a >= 1_000:
        return f"{MONEDA} {x / 1_000:,.1f} K"
    return f"{MONEDA} {x:,.0f}"


def fmt_pct(x: float, dec: int = 1) -> str:
    return "—" if pd.isna(x) else f"{x * 100:.{dec}f}%"


def pct_de(valor: float, base: float, nombre: str) -> str:
    """Texto del comentario: qué % representa un valorizado respecto de su base."""
    return f"{valor / base * 100:.1f}% del {nombre}" if base and base > 0 else ""


def delta_html(actual: float, previo: float, tipo: str) -> str:
    """tipo='pct' -> variación relativa; tipo='pp' -> puntos porcentuales."""
    if previo is None or pd.isna(actual) or pd.isna(previo):
        return '<span class="dl dl-n">&nbsp;</span>'
    if tipo == "pct":
        if previo == 0:
            return '<span class="dl dl-n">&nbsp;</span>'
        v = (actual / previo - 1) * 100
        txt = f"{v:+.1f}%"
    else:
        v = (actual - previo) * 100
        txt = f"{v:+.1f} pp"
    cls = "dl-u" if v >= 0 else "dl-d"
    flecha = "▲" if v >= 0 else "▼"
    return f'<span class="dl {cls}">{flecha} {txt} vs mes ant.</span>'


def mes_label(ts: pd.Timestamp) -> str:
    return f"{MESES[ts.month - 1]} {ts.year}"

# ----- Cumplimiento de proveedores (PO vs RN) -----
TOL_PO = 0.001  # tolerancia para considerar "atendido" (0,1 %)

ESTADOS_ITEM = ["Atendido", "Atendido parcial", "Pendiente de atención"]
ESTADOS_PO = ["Atendida", "Atendida parcial", "Pendiente de atención"]


def _es_falso(serie: pd.Series) -> pd.Series:
    return serie.astype(str).str.strip().str.upper().isin(["FALSE", "0", "0.0"])


def construir_items_po(df: pd.DataFrame) -> pd.DataFrame:
    """
    Una fila por ítem de PO (PO Number + PO Item Id) con lo pedido (PO) y lo recibido en almacén (RN).
    Excluye los despachos del stock propio (ALMACEN LIMA) y las POs inválidas.
    """
    d = df.copy()
    d.columns = [str(c).strip() for c in d.columns]
    for c in ["PO Quantity", "PO Total Price", "RN Received Quantity", "RN Total Price"]:
        d[c] = pd.to_numeric(d[c], errors="coerce").fillna(0) if c in d.columns else 0.0
    for c in ["PO Delivery Date", "RN Delivery Date"]:
        d[c] = pd.to_datetime(d[c], errors="coerce").dt.normalize() if c in d.columns else pd.NaT

    prov = d["PO Suppliers"].fillna("SIN PROVEEDOR").astype(str)
    d["Proveedor"] = prov
    valida = ~_es_falso(d["PO Valid"]) if "PO Valid" in d.columns else True
    d = d[~prov.str.upper().str.contains(STOCK_PROPIO, na=False) & valida]
    d = d[d["PO Number"].notna() & d["PO Item Id"].notna()].copy()

    # Una RN inválida no cuenta como atención
    if "RN Is Valid" in d.columns:
        inval = _es_falso(d["RN Is Valid"]) & d["RN Number"].notna()
        d.loc[inval, ["RN Received Quantity", "RN Total Price"]] = 0.0

    d["_f_rn"] = d["RN Delivery Date"].where(d["RN Received Quantity"] > 0)

    if "SOE Site" in d.columns:
        alm = d["SOE Site"].fillna("SIN DATO").astype(str).str.upper()
    else:
        alm = pd.Series("SIN DATO", index=d.index)
    d["Almacén"] = alm.str.replace("NEWREST ALMACEN", "", regex=False).str.strip().replace("", "SIN DATO")
    for c, n in [("PO Family Name", "Familia"), ("PO Group Name", "Grupo"), ("PO Item Name", "Ítem"),
                 ("PO Packing Unit", "Unidad")]:
        d[n] = d[c].fillna("SIN DATO").astype(str) if c in d.columns else "SIN DATO"

    it = d.groupby(["PO Number", "PO Item Id"], sort=False).agg(
        Proveedor=("Proveedor", "first"), **{"Fecha PO": ("PO Delivery Date", "min")},
        Familia=("Familia", "first"), Grupo=("Grupo", "first"), Ítem=("Ítem", "first"),
        Unidad=("Unidad", "first"), Almacén=("Almacén", "first"),
        **{"Cant PO": ("PO Quantity", "sum"), "Monto PO": ("PO Total Price", "sum"),
           "Cant RN": ("RN Received Quantity", "sum"), "Monto RN": ("RN Total Price", "sum"),
           "F primera RN": ("_f_rn", "min"), "F última RN": ("_f_rn", "max")},
    ).reset_index().rename(columns={"PO Number": "PO", "PO Item Id": "Item Id"})
    it["Mes"] = it["Fecha PO"].dt.to_period("M").dt.to_timestamp()
    return it


def clasificar_items(it: pd.DataFrame, corte: pd.Timestamp) -> pd.DataFrame:
    """Evalúa cada ítem de las POs ya vencidas (PO Delivery Date <= corte) según lo recibido (RN)."""
    it = it[it["Fecha PO"] <= corte].copy()
    completo = it["Cant RN"] >= it["Cant PO"] * (1 - TOL_PO)
    parcial = (it["Cant RN"] > 0) & ~completo
    it["Estado ítem"] = np.select([completo, parcial], ["Atendido", "Atendido parcial"],
                                  default="Pendiente de atención")
    it["Estado ítem"] = pd.Categorical(it["Estado ítem"], categories=ESTADOS_ITEM, ordered=True)

    it["Cant atendida"] = np.minimum(it["Cant RN"], it["Cant PO"])
    it["Faltante"] = it["Cant PO"] - it["Cant atendida"]
    frac = (it["Cant atendida"] / it["Cant PO"].replace(0, np.nan)).fillna(1.0)
    it["% cumpl"] = frac
    it["Monto atendido"] = it["Monto PO"] * frac
    it["Monto pendiente"] = it["Monto PO"] - it["Monto atendido"]
    # Días transcurridos desde la fecha de la PO para lo que sigue sin atender completo
    it["Días vencida"] = np.where(completo, np.nan, (corte - it["Fecha PO"]).dt.days).astype(float)
    return it


def resumir_pos(it: pd.DataFrame) -> pd.DataFrame:
    """Una fila por PO con su estado global (¿se atendió el N° de PO completo?)."""
    e = it["Estado ítem"].astype(str)
    t = it.assign(_a=(e == "Atendido").astype(int), _p=(e == "Atendido parcial").astype(int),
                  _s=(e == "Pendiente de atención").astype(int))
    g = t.groupby("PO", sort=False).agg(
        Proveedor=("Proveedor", "first"), **{"Fecha PO": ("Fecha PO", "first")}, Almacén=("Almacén", "first"),
        Mes=("Mes", "first"), Ítems=("Item Id", "size"),
        **{"Ítems atendidos": ("_a", "sum"), "Ítems parciales": ("_p", "sum"), "Ítems pendientes": ("_s", "sum"),
           "Cant PO": ("Cant PO", "sum"), "Cant atendida": ("Cant atendida", "sum"),
           "Monto PO": ("Monto PO", "sum"), "Monto atendido": ("Monto atendido", "sum"),
           "Monto pendiente": ("Monto pendiente", "sum"), "Días vencida": ("Días vencida", "max")},
    ).reset_index()
    p, s = g["Ítems parciales"], g["Ítems pendientes"]
    estado = np.select([(p + s) == 0, s == g["Ítems"]], ["Atendida", "Pendiente de atención"],
                       default="Atendida parcial")
    g["Estado PO"] = pd.Categorical(estado, categories=ESTADOS_PO, ordered=True)
    # Por qué quedó parcial: ítems sin atender, ítems con menos cantidad, o ambas cosas
    g["Tipo de parcialidad"] = np.select(
        [(estado == "Atendida parcial") & (s > 0) & (p == 0),
         (estado == "Atendida parcial") & (s == 0) & (p > 0),
         (estado == "Atendida parcial")],
        ["Faltan ítems", "Cantidades incompletas", "Faltan ítems y cantidades incompletas"], default="—")
    g["% cantidad"] = np.where(g["Cant PO"] > 0, g["Cant atendida"] / g["Cant PO"], np.nan)
    return g


def kpis_po(pos: pd.DataFrame, items: pd.DataFrame) -> dict:
    n = len(pos)
    est = pos["Estado PO"].astype(str)
    at, pa, pe = int((est == "Atendida").sum()), int((est == "Atendida parcial").sum()), int((est == "Pendiente de atención").sum())
    div = lambda a, b: a / b if b > 0 else np.nan
    ie = items["Estado ítem"].astype(str)
    inc = pos.loc[est != "Atendida", "Días vencida"]
    return {
        "pos": n, "pos_at": at, "pos_parc": pa, "pos_pend": pe,
        "p_at": div(at, n), "p_parc": div(pa, n), "p_pend": div(pe, n),
        "items": len(items), "it_at": int((ie == "Atendido").sum()), "it_parc": int((ie == "Atendido parcial").sum()),
        "it_pend": int((ie == "Pendiente de atención").sum()),
        "p_items_at": div(int((ie == "Atendido").sum()), len(items)),
        "cq": div(pos["Cant atendida"].sum(), pos["Cant PO"].sum()),
        "cv": div(pos["Monto atendido"].sum(), pos["Monto PO"].sum()),
        "monto_po": pos["Monto PO"].sum(), "monto_pend": pos["Monto pendiente"].sum(),
        "dias_prom": inc.mean() if len(inc) else np.nan,
    }


def resumen_proveedor(pos: pd.DataFrame, col: str = "Proveedor") -> pd.DataFrame:
    e = pos["Estado PO"].astype(str)
    t = pos.assign(_a=(e == "Atendida"), _p=(e == "Atendida parcial"), _s=(e == "Pendiente de atención"),
                   _inc=(e != "Atendida"))
    d = t.groupby(col).agg(
        POs=("PO", "size"),
        **{"Atendidas": ("_a", "sum"), "Parciales": ("_p", "sum"), "Pendientes": ("_s", "sum"),
           "Incompletas": ("_inc", "sum"),
           "Cant PO": ("Cant PO", "sum"), "Cant atendida": ("Cant atendida", "sum"),
           "Monto PO": ("Monto PO", "sum"), "Monto atendido": ("Monto atendido", "sum"),
           "Monto pendiente": ("Monto pendiente", "sum")},
    ).reset_index()
    d["% atendidas"] = d["Atendidas"] / d["POs"]
    d["% pendientes"] = d["Pendientes"] / d["POs"]
    d["% cantidad"] = np.where(d["Cant PO"] > 0, d["Cant atendida"] / d["Cant PO"], np.nan)
    d["% valorizado"] = np.where(d["Monto PO"] > 0, d["Monto atendido"] / d["Monto PO"], np.nan)
    dias = t[t["_inc"]].groupby(col)["Días vencida"].mean().rename("Días vencida prom.")
    return d.merge(dias, on=col, how="left")

# ##### FIN LÓGICA #####


# =====================================================
# ESTILO (tarjetas como en el dashboard de referencia)
# =====================================================
CSS = f"""
<style>
.stApp {{ background:#EEF2F6; }}
header[data-testid="stHeader"] {{ background:transparent; }}
[data-testid="stSidebar"], [data-testid="stSidebarCollapsedControl"], [data-testid="collapsedControl"] {{ display:none; }}
.block-container {{ padding-top:1.2rem; padding-left:1.6rem; padding-right:14rem; max-width:100%; }}
div[class*="st-key-card"] {{
    background:#FFFFFF; border-radius:18px; padding:16px 22px;
    box-shadow:0 2px 12px rgba(30,60,90,.08); color:{INK};
}}
div[class*="st-key-card_hero"] {{
    background:linear-gradient(120deg,{NAVY} 0%,#2B7AA8 100%); color:#fff;
}}
.hero-t {{ font-size:1.9rem; font-weight:800; letter-spacing:.2px; margin:0; cursor:help; }}
.hero-s {{ font-size:.95rem; opacity:.9; margin-top:2px; }}
.hero-tag {{ float:right; background:rgba(255,255,255,.18); padding:6px 16px;
             border-radius:20px; font-weight:700; margin-top:-46px; }}
.ct {{ font-size:1.02rem; font-weight:800; color:{INK}; text-align:center; margin-bottom:4px; }}
.cs {{ font-size:.78rem; color:{MUTED}; text-align:center; margin-bottom:6px; }}
.khint {{ text-align:center; color:{MUTED}; font-size:.8rem; margin-bottom:8px; }}
.kitem {{ text-align:center; }}
.kico {{ width:44px; height:44px; border-radius:50%; margin:0 auto 4px; display:flex;
         align-items:center; justify-content:center; font-size:1.3rem; }}
.klab {{ font-weight:700; color:{INK}; font-size:.9rem; }}
.kval {{ font-weight:800; color:{INK}; font-size:1.4rem; line-height:1.5; }}
.kdelta {{ min-height:1.3rem; }}
.dl {{ font-size:.76rem; font-weight:700; }}
.dl-u {{ color:#1B9E5A; }} .dl-d {{ color:{RED}; }} .dl-n {{ color:{MUTED}; }}
div[class*="st-key-kbtn_"] {{ display:flex; justify-content:center; }}
div[class*="st-key-kbtn_"] button {{
    background:#F4F8FB; border:1px solid #D3DFEA; border-radius:12px; padding:0 14px; min-height:0;
}}
div[class*="st-key-kbtn_"] button p {{ font-size:1.4rem; font-weight:800; color:{NAVY}; line-height:1.5; margin:0; }}
div[class*="st-key-kbtn_"] button:hover {{ background:{NAVY}; border-color:{NAVY}; }}
div[class*="st-key-kbtn_"] button:hover p {{ color:#FFFFFF; }}
.srow {{ display:flex; justify-content:space-between; padding:7px 0; border-bottom:1px solid #EEF1F4;
         font-size:.92rem; }}
.srow span {{ color:{MUTED}; }} .srow b {{ color:{INK}; }}
.tip {{ cursor:help; border-bottom:1px dotted {MUTED}; }}
.pair {{ display:flex; justify-content:space-around; text-align:center; margin-top:2px; }}
.pair .pl {{ font-size:.8rem; color:{MUTED}; font-weight:700; }}
.pair .pv {{ font-size:1.15rem; font-weight:800; }}
.ksub {{ font-size:.76rem; color:{MUTED}; min-height:1.3rem; }}
div[class*="st-key-rail_reportes"] {{
    position:fixed; top:84px; right:16px; width:204px; z-index:1000; background:#FFFFFF;
    border-radius:16px; padding:14px 16px; box-shadow:0 2px 12px rgba(30,60,90,.14); color:{INK};
}}
.rail-t {{ font-weight:800; font-size:1rem; color:{NAVY}; margin-bottom:4px; }}
div[class*="st-key-rail_reportes"] label p {{ font-size:.92rem; font-weight:600; color:{INK}; }}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


# =====================================================
# CARGA (con caché)
# =====================================================
@st.cache_data(show_spinner="Cargando trazabilidad…")
def cargar(origen_bytes: bytes | None, ruta: str | None, tope: bool):
    origen = io.BytesIO(origen_bytes) if origen_bytes is not None else ruta
    crudo = leer_excel(origen)
    lineas_, corte_ = consolidar(crudo, tope_linea=tope)
    return lineas_, corte_, construir_items_po(crudo)


subido = st.file_uploader(
    "Subir TRAZABILIDAD_BASE.xlsx",
    type=["xlsx"]
)

if subido is None:
    st.info(
        "Sube TRAZABILIDAD_BASE.xlsx para continuar"
    )
    st.stop()

lineas, corte_def, items_po = cargar(
    subido.getvalue(),
    None,
    TOPE_LINEA
)

corte = pd.Timestamp(corte_def).normalize()


# =====================================================
# HELPERS DE GRÁFICOS
# =====================================================
def base_layout(fig, h=260, legend=False):
    fig.update_layout(
        height=h, margin=dict(l=8, r=8, t=8, b=8), paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)", font=dict(family="Segoe UI, Arial", color=INK),
        showlegend=legend,
    )
    return fig


def grafico(fig, key):
    st.plotly_chart(fig, key=key, config={"displayModeBar": False})


def gauge(valor, meta, key):
    v = 0 if pd.isna(valor) else valor * 100
    fig = go.Figure(go.Indicator(
        mode="gauge+number", value=v,
        number={"suffix": "%", "valueformat": ".1f", "font": {"size": 34, "color": INK}},
        gauge={
            "axis": {"range": [0, 100], "tickwidth": 0, "tickfont": {"size": 10, "color": MUTED}},
            "bar": {"color": NAVY, "thickness": 0.28},
            "bgcolor": "white", "borderwidth": 0,
            "steps": [
                {"range": [0, meta - 5], "color": "#F7BFC1"},
                {"range": [meta - 5, meta], "color": "#FBE7A8"},
                {"range": [meta, 100], "color": "#BDE3CB"},
            ],
            "threshold": {"line": {"color": INK, "width": 3}, "thickness": 0.8, "value": meta},
        },
    ))
    base_layout(fig, h=210)
    fig.update_layout(margin=dict(l=20, r=20, t=20, b=0))
    grafico(fig, key)


def donut(cumplido, pendiente, nombre_ok, key):
    tot = cumplido + pendiente
    pct = cumplido / tot if tot > 0 else 0
    fig = go.Figure(go.Pie(
        labels=[nombre_ok, "Pendiente"], values=[cumplido, max(pendiente, 0)], hole=0.74,
        marker=dict(colors=[NAVY, RED], line=dict(color="white", width=2)),
        textinfo="none", sort=False, direction="clockwise",
        hovertemplate="%{label}: " + MONEDA + " %{value:,.0f}<extra></extra>",
    ))
    fig.add_annotation(text=f"<b>{pct * 100:.0f}%</b>", x=0.5, y=0.54, showarrow=False,
                       font=dict(size=34, color=NAVY))
    fig.add_annotation(text=nombre_ok.lower(), x=0.5, y=0.38, showarrow=False,
                       font=dict(size=12, color=MUTED))
    base_layout(fig, h=235)
    grafico(fig, key)


def pie(d, col, val, key):
    fig = go.Figure(go.Pie(
        labels=d[col], values=d[val], hole=0.0, sort=False,
        marker=dict(colors=PALETA[: len(d)], line=dict(color="white", width=2)),
        textinfo="percent", textfont=dict(color="white", size=13),
        hovertemplate="%{label}<br>" + MONEDA + " %{value:,.0f}<extra></extra>",
    ))
    base_layout(fig, h=235, legend=True)
    fig.update_layout(legend=dict(orientation="h", y=-0.08, font=dict(size=11)),
                      margin=dict(l=4, r=4, t=4, b=34))
    grafico(fig, key)


def stats(filas):
    """Cada fila: (etiqueta, valor) o (etiqueta, valor, comentario al pasar el puntero)."""
    out = []
    for f in filas:
        a, b = f[0], f[1]
        t = f[2] if len(f) > 2 and f[2] else ""
        val = f'<b class="tip" title="{t}">{b}</b>' if t else f"<b>{b}</b>"
        out.append(f'<div class="srow"><span>{a}</span>{val}</div>')
    return "<div>" + "".join(out) + "</div>"


def par(a_lab, a_val, b_lab, b_val, a_tip="", b_tip=""):
    return (f'<div class="pair"><div><div class="pl">{a_lab}</div>'
            f'<div class="pv tip" title="{a_tip}" style="color:{NAVY}">{a_val}</div></div>'
            f'<div><div class="pl">{b_lab}</div>'
            f'<div class="pv tip" title="{b_tip}" style="color:{RED}">{b_val}</div></div></div>')


COL_ESTADO_PO = {"Atendida": GREEN, "Atendida parcial": AMBAR, "Pendiente de atención": RED}
COL_ESTADO_ITEM = {"Atendido": GREEN, "Atendido parcial": AMBAR, "Pendiente de atención": RED}


def donut_estados(etiquetas, valores, colores, centro, sub, key):
    fig = go.Figure(go.Pie(
        labels=etiquetas, values=valores, hole=0.68, sort=False, direction="clockwise",
        marker=dict(colors=colores, line=dict(color="white", width=2)), textinfo="none",
        hovertemplate="%{label}<br>%{value:,} (%{percent})<extra></extra>",
    ))
    fig.add_annotation(text=f"<b>{centro}</b>", x=0.5, y=0.56, showarrow=False, font=dict(size=30, color=NAVY))
    fig.add_annotation(text=sub, x=0.5, y=0.40, showarrow=False, font=dict(size=12, color=MUTED))
    base_layout(fig, h=290, legend=True)
    fig.update_layout(legend=dict(orientation="h", y=-0.05, font=dict(size=10)), margin=dict(l=4, r=4, t=4, b=70))
    grafico(fig, key)


# =====================================================
# ENCABEZADO (se llena después de leer los filtros) Y FILTROS
# =====================================================
def reporte_fill_rate():
    cont_hero = st.container(key="card_hero")

    with st.container(key="card_filtros"):
        f1, f2, f3, f4, f5, f6 = st.columns([1.35, 1, 1, 1, 1.4, 0.95])
        # El casillero se lee primero porque define las fechas y meses disponibles de los demás filtros
        with f6:
            st.markdown('<div style="height:1.75rem"></div>', unsafe_allow_html=True)
            solo_vencidas = st.checkbox(
                "Solo entregas vencidas", value=SOLO_VENCIDAS_DEFECTO, key="f_vencidas",
                help=f"Mide solo las líneas cuya fecha de entrega (SO Delivery Date) ya llegó a la fecha de corte "
                     f"({corte:%d/%m/%Y}). Desmárcalo para incluir también las entregas futuras.",
            )

        g_base = lineas[lineas["Entrega"] <= corte] if solo_vencidas else lineas
        if g_base.empty:
            st.warning("No hay datos para mostrar.")
            st.stop()

        f_min, f_max = g_base["Entrega"].min().date(), g_base["Entrega"].max().date()
        meses_base = sorted(g_base["Mes"].dropna().unique())
        opciones_mes = ["Todo el periodo"] + [mes_label(pd.Timestamp(m)) for m in reversed(meses_base)]
        mapa_mes = {mes_label(pd.Timestamp(m)): pd.Timestamp(m) for m in meses_base}
        tag = int(solo_vencidas)  # las claves cambian con el casillero para que fecha y periodo se reinicien

        with f1:
            rango = st.date_input(
                "Fecha de entrega (SO Delivery Date)", value=(f_min, f_max), min_value=f_min, max_value=f_max,
                key=f"f_fecha_{tag}", help="Día o rango. Para un día puntual haz clic dos veces en la misma fecha.",
            )
        with f2:
            periodo = st.selectbox("Periodo (mes de entrega)", opciones_mes, key=f"f_periodo_{tag}")
        with f3:
            sel_fam = st.multiselect("Familia", sorted(lineas["Familia"].unique()), key="f_fam", placeholder="Todas")
        with f4:
            sel_alm = st.multiselect("Almacén (SOE Site)", sorted(lineas["Almacén"].unique()), key="f_alm",
                                     placeholder="Todos")
        with f5:
            sel_site = st.multiselect("SO Site", sorted(lineas["Sitio"].unique()), key="f_site", placeholder="Todos")

    # --- Aplicación de filtros ---
    g_ctx = g_base
    if sel_fam:
        g_ctx = g_ctx[g_ctx["Familia"].isin(sel_fam)]
    if sel_alm:
        g_ctx = g_ctx[g_ctx["Almacén"].isin(sel_alm)]
    if sel_site:
        g_ctx = g_ctx[g_ctx["Sitio"].isin(sel_site)]

    if g_ctx.empty:
        st.warning("No hay datos con los filtros seleccionados.")
        st.stop()

    if periodo == "Todo el periodo":
        g_p, g_prev = g_ctx, None
    else:
        m_sel = mapa_mes[periodo]
        g_p = g_ctx[g_ctx["Mes"] == m_sel]
        prev = g_ctx[g_ctx["Mes"] == (m_sel - pd.DateOffset(months=1))]
        g_prev = prev if not prev.empty else None

    if isinstance(rango, (tuple, list)):
        d_ini = rango[0]
        d_fin = rango[1] if len(rango) == 2 else rango[0]  # con un solo clic se toma ese día
    else:
        d_ini = d_fin = rango

    etiqueta = periodo
    if (d_ini, d_fin) != (f_min, f_max):
        g_p = g_p[(g_p["Entrega"].dt.date >= d_ini) & (g_p["Entrega"].dt.date <= d_fin)]
        g_prev = None  # la comparación contra el mes anterior no aplica a un día o rango
        etiqueta = f"{d_ini:%d/%m/%Y}" if d_ini == d_fin else f"{d_ini:%d/%m/%Y} – {d_fin:%d/%m/%Y}"

    if g_p.empty:
        st.warning("No hay líneas con esa combinación de filtros (fecha de entrega, periodo, familia, almacén y site).")
        st.stop()

    K = kpis(g_p)
    KP = kpis(g_prev) if g_prev is not None else None


    def prev_val(k):
        return KP[k] if KP is not None else None


    # --- Encabezado ---
    partes = []
    if sel_fam:
        partes.append(f"Familia: {', '.join(sel_fam)}")
    if sel_alm:
        partes.append(f"Almacén: {', '.join(sel_alm)}")
    if sel_site:
        partes.append(f"Site: {', '.join(sel_site)}")
    filtros_txt = " · ".join(partes) or "Todas las familias, almacenes y sites"

    with cont_hero:
        st.markdown(
            f'<p class="hero-t" title="Data, Análisis, Visibilidad, Núcleo operativo, Control e Inteligencia logística">'
            f'📊 Da Vinci - Gestión de Abastecimiento</p>'
            f'<p class="hero-s">Reporte: {REPORTES[0]} · Compras (SO vs RN) · Despachos (RN vs OF) · Confirmación (OF vs RNO) · Montos en {MONEDA} · '
            f'Corte {corte:%d/%m/%Y} ({"solo entregas vencidas" if solo_vencidas else "incluye entregas futuras"}) · '
            f'{filtros_txt}</p>'
            f'<span class="hero-tag">{etiqueta}</span>',
            unsafe_allow_html=True,
        )


    # =====================================================
    # VENTANAS DE EFICIENCIA (se abren al hacer clic en el monto de RN, OF o RNO)
    # =====================================================
    def contenido_eficiencia(pref, sub, valor, meta, filas, ok, pend, nombre_ok, nombre_pend,
                             pie_df, pie_col, pie_val, pie_titulo, pie_vacio, tip_ok="", tip_pend=""):
        st.caption(sub)
        c_g, c_s = st.columns([1.15, 1])
        with c_g:
            gauge(valor, meta, f"{pref}_g")
        with c_s:
            st.markdown(stats(filas), unsafe_allow_html=True)
        c_d, c_p = st.columns(2)
        with c_d:
            st.markdown('<div class="ct">Cumplimiento valorizado</div>', unsafe_allow_html=True)
            donut(ok, pend, nombre_ok, f"{pref}_d")
            st.markdown(par(nombre_ok, fmt_monto(ok), nombre_pend, fmt_monto(pend), tip_ok, tip_pend),
                        unsafe_allow_html=True)
        with c_p:
            st.markdown(f'<div class="ct">{pie_titulo}</div>', unsafe_allow_html=True)
            if pie_df.empty:
                st.info(pie_vacio)
            else:
                pie(pie_df, pie_col, pie_val, f"{pref}_p")


    SUFIJO_TOPE = "; con tope por línea" if TOPE_LINEA else ""


    def tip_cantidad(atendidas, pedidas):
        return (f"Unidades {atendidas} ÷ unidades {pedidas}, sumando todas las líneas sin distinguir la unidad de medida"
                f"{SUFIJO_TOPE}")


    @st.dialog("Eficiencia de compras · SO vs RN", width="large")
    def dlg_compras():
        pg = top_y_otros(resumen_por(g_p, "Grupo"), "Grupo", "Brecha compras")
        t_ok = pct_de(K["rn_c"], K["so"], "solicitado (SO)")
        t_br = pct_de(K["brecha_c"], K["so"], "solicitado (SO)")
        contenido_eficiencia(
            "dc", "Fill Rate: SO solicitado vs RN recepcionado en almacén", K["fr"], META_COMPRAS,
            [("Solicitado (SO)", fmt_monto(K["so"])),
             ("Recepcionado (≤ SO)", fmt_monto(K["rn_c"]), t_ok),
             ("Brecha", fmt_monto(K["brecha_c"]), t_br),
             ("Cumplimiento por cantidad", fmt_pct(K["cq_c"]), tip_cantidad("recepcionadas", "solicitadas")),
             ("Meta", f"{META_COMPRAS}%")],
            K["rn_c"], K["brecha_c"], "Recepcionado", "Pendiente",
            pg, "Grupo", "Brecha compras", "Brecha de compras por grupo", "Sin brecha de compras en el periodo.",
            t_ok, t_br,
        )


    @st.dialog("Eficiencia de despacho · RN vs OF", width="large")
    def dlg_despacho():
        ps = top_y_otros(resumen_por(g_p, "Sitio"), "Sitio", "Brecha despacho")
        t_ok = pct_de(K["of_c"], K["rn"], "recepcionado (RN)")
        t_br = pct_de(K["brecha_d"], K["rn"], "recepcionado (RN)")
        contenido_eficiencia(
            "dd", "Eficiencia: RN recepcionado vs OF despachado a sitio", K["ef"], META_DESPACHO,
            [("Recepcionado (RN)", fmt_monto(K["rn"])),
             ("Despachado (OF)", fmt_monto(K["of_c"]), t_ok),
             ("Por despachar", fmt_monto(K["brecha_d"]), t_br),
             ("Cumplimiento por cantidad", fmt_pct(K["cq_d"]), tip_cantidad("despachadas", "recepcionadas")),
             ("Meta", f"{META_DESPACHO}%")],
            K["of_c"], K["brecha_d"], "Despachado", "Por despachar",
            ps, "Sitio", "Brecha despacho", "Stock por despachar por sitio", "Sin saldo por despachar en el periodo.",
            t_ok, t_br,
        )


    @st.dialog("Confirmación en operación · OF vs RNO", width="large")
    def dlg_confirmacion():
        pe = top_y_otros(resumen_por(g_p, "Sitio"), "Sitio", "Brecha confirmación")
        t_ok = pct_de(K["rno_c"], K["of"], "despachado (OF)")
        t_br = pct_de(K["brecha_e"], K["of"], "despachado (OF)")
        contenido_eficiencia(
            "de", "OF despachado por el almacén central vs RNO confirmado por operaciones", K["ce"], META_CONFIRMACION,
            [("Despachado (OF)", fmt_monto(K["of"])),
             ("Confirmado (RNO)", fmt_monto(K["rno_c"]), t_ok),
             ("Sin confirmar", fmt_monto(K["brecha_e"]), t_br),
             ("Cumplimiento por cantidad", fmt_pct(K["cq_e"]), tip_cantidad("confirmadas", "despachadas")),
             ("Meta", f"{META_CONFIRMACION}%")],
            K["rno_c"], K["brecha_e"], "Confirmado", "Sin confirmar",
            pe, "Sitio", "Brecha confirmación", "Despacho sin confirmar por sitio",
            "Todo lo despachado está confirmado en el periodo.",
            t_ok, t_br,
        )


    # =====================================================
    # FRANJA DE KPIs (clic en RN, OF o RNO para ver su eficiencia)
    # =====================================================
    def kcol(col, icono, color, titulo, valor, delta, boton=None, ayuda=None):
        with col:
            st.markdown(
                f'<div class="kitem"><div class="kico" style="background:{color}22">{icono}</div>'
                f'<div class="klab">{titulo}</div></div>', unsafe_allow_html=True)
            clic = False
            if boton:
                clic = st.button(valor, key=f"kbtn_{boton}", help=ayuda)
            else:
                st.markdown(f'<div class="kitem"><div class="kval">{valor}</div></div>', unsafe_allow_html=True)
            st.markdown(f'<div class="kitem kdelta">{delta}</div>', unsafe_allow_html=True)
        return clic


    with st.container(key="card_kpis"):
        st.markdown(
            '<div class="khint">Haz clic en el monto de <b>Recepcionado</b>, <b>Despachado</b> o <b>Confirmado</b> '
            'para ver su eficiencia</div>', unsafe_allow_html=True)
        kc = st.columns(7)
        kcol(kc[0], "🧾", NAVY, "Solicitado (SO)", fmt_monto(K["so"]), delta_html(K["so"], prev_val("so"), "pct"))
        abre_c = kcol(kc[1], "📥", STEEL, "Recepcionado (RN)", fmt_monto(K["rn"]),
                      delta_html(K["rn"], prev_val("rn"), "pct"), boton="rn", ayuda="Ver eficiencia de compras (SO vs RN)")
        abre_d = kcol(kc[2], "🚚", RED, "Despachado (OF)", fmt_monto(K["of"]),
                      delta_html(K["of"], prev_val("of"), "pct"), boton="of", ayuda="Ver eficiencia de despacho (RN vs OF)")
        abre_e = kcol(kc[3], "✅", GREEN, "Confirmado (RNO)", fmt_monto(K["rno"]),
                      delta_html(K["rno"], prev_val("rno"), "pct"), boton="rno", ayuda="Ver confirmación en operación (OF vs RNO)")
        kcol(kc[4], "🎯", NAVY, "Fill Rate Compras", fmt_pct(K["fr"]), delta_html(K["fr"], prev_val("fr"), "pp"))
        kcol(kc[5], "⚙️", RED, "Eficiencia Despacho", fmt_pct(K["ef"]), delta_html(K["ef"], prev_val("ef"), "pp"))
        kcol(kc[6], "🏭", GREEN, "Confirmación Operación", fmt_pct(K["ce"]), delta_html(K["ce"], prev_val("ce"), "pp"))

    if abre_c:
        dlg_compras()
    elif abre_d:
        dlg_despacho()
    elif abre_e:
        dlg_confirmacion()


    # =====================================================
    # RANKING DE BRECHAS
    # =====================================================
    dim = st.radio("Analizar brechas por:", ["Sitio", "Proveedor", "Grupo", "Familia"], horizontal=True, key="dim_rank")
    rd = resumen_por(base_dim(g_p, dim), dim)
    if dim == "Proveedor":
        st.caption("El análisis por proveedor excluye ALMACEN LIMA (despachos del stock); sus montos sí están en los totales.")


    def barras(d, col, valor, color, key, titulo):
        d = d[d[valor] > 0].sort_values(valor, ascending=False).head(10)
        fig = go.Figure(go.Bar(
            x=d[valor], y=d[col], orientation="h", marker_color=color,
            text=[fmt_monto(v) for v in d[valor]], textposition="outside", cliponaxis=False,
            customdata=np.stack([d["Fill Rate"] * 100, d["Eficiencia"] * 100, d["Confirmación"] * 100], axis=-1),
            hovertemplate="%{y}<br>Brecha: " + MONEDA + " %{x:,.0f}<br>Fill Rate: %{customdata[0]:.1f}%"
                          "<br>Eficiencia: %{customdata[1]:.1f}%<br>Confirmación: %{customdata[2]:.1f}%<extra></extra>",
        ))
        base_layout(fig, h=360)
        fig.update_yaxes(autorange="reversed", title=None, tickfont=dict(size=11))
        fig.update_xaxes(showgrid=True, gridcolor="#EEF1F4", showticklabels=False)
        fig.update_layout(margin=dict(l=8, r=70, t=8, b=8))
        st.markdown(f'<div class="ct">{titulo}</div>', unsafe_allow_html=True)
        grafico(fig, key)


    r1, r2, r3 = st.columns(3)
    with r1:
        with st.container(key="card_rank_compras"):
            barras(rd, dim, "Brecha compras", RED, "r_compras", f"Mayor brecha de compras (SO − RN) · Top 10 {dim}")
    with r2:
        with st.container(key="card_rank_desp"):
            barras(rd, dim, "Brecha despacho", NAVY, "r_desp", f"Mayor saldo por despachar (RN − OF) · Top 10 {dim}")
    with r3:
        with st.container(key="card_rank_conf"):
            barras(rd, dim, "Brecha confirmación", GREEN, "r_conf", f"Mayor saldo sin confirmar (OF − RNO) · Top 10 {dim}")

    # =====================================================
    # TENDENCIAS (ignoran periodo y fecha, respetan familia, almacén y site)
    # =====================================================
    tm = resumen_por(g_ctx.dropna(subset=["Mes"]), "Mes").sort_values("Mes")
    etq = [mes_label(pd.Timestamp(m)) for m in tm["Mes"]]

    d1, d2 = st.columns(2)
    with d1:
        with st.container(key="card_trend"):
            st.markdown('<div class="ct">Tendencia mensual de cumplimiento</div>', unsafe_allow_html=True)
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=etq, y=tm["Fill Rate"] * 100, name="Fill Rate Compras", mode="lines+markers",
                                     line=dict(color=NAVY, width=3), marker=dict(size=8)))
            fig.add_trace(go.Scatter(x=etq, y=tm["Eficiencia"] * 100, name="Eficiencia Despacho", mode="lines+markers",
                                     line=dict(color=RED, width=3), marker=dict(size=8)))
            fig.add_trace(go.Scatter(x=etq, y=tm["Confirmación"] * 100, name="Confirmación Operación", mode="lines+markers",
                                     line=dict(color=GREEN, width=3), marker=dict(size=8)))
            fig.add_hline(y=META_COMPRAS, line_dash="dot", line_color=MUTED, annotation_text=f"Meta {META_COMPRAS}%",
                          annotation_position="bottom right")
            base_layout(fig, h=340, legend=True)
            fig.update_yaxes(ticksuffix="%", gridcolor="#EEF1F4", rangemode="tozero")
            fig.update_layout(legend=dict(orientation="h", y=-0.18))
            grafico(fig, "t_rates")

    with d2:
        with st.container(key="card_combo"):
            st.markdown('<div class="ct">Montos mensuales (SO · RN · OF · RNO) y Fill Rate</div>', unsafe_allow_html=True)
            fig = make_subplots(specs=[[{"secondary_y": True}]])
            fig.add_trace(go.Bar(x=etq, y=tm["SO"], name="Solicitado", marker_color=STEEL), secondary_y=False)
            fig.add_trace(go.Bar(x=etq, y=tm["RN"], name="Recepcionado", marker_color=NAVY), secondary_y=False)
            fig.add_trace(go.Bar(x=etq, y=tm["OF_c"], name="Despachado", marker_color=PINK), secondary_y=False)
            fig.add_trace(go.Bar(x=etq, y=tm["RNO_c"], name="Confirmado", marker_color=GREEN), secondary_y=False)
            fig.add_trace(go.Scatter(x=etq, y=tm["Fill Rate"] * 100, name="Fill Rate %", mode="lines+markers",
                                     line=dict(color=RED, width=3),
                                     marker=dict(size=8, color="white", line=dict(color=RED, width=2))),
                          secondary_y=True)
            base_layout(fig, h=340, legend=True)
            fig.update_layout(barmode="group", legend=dict(orientation="h", y=-0.18))
            fig.update_yaxes(title_text=f"Monto ({MONEDA})", gridcolor="#EEF1F4", secondary_y=False)
            fig.update_yaxes(ticksuffix="%", showgrid=False, rangemode="tozero", secondary_y=True)
            grafico(fig, "t_combo")

    # =====================================================
    # TABLA DE DETALLE (con selector propio)
    # =====================================================
    VISTAS_TABLA = {
        "Por sitio": "Sitio",
        "Por proveedor": "Proveedor",
        "Por familia": "Familia",
        "Por almacén (SOE Site)": "Almacén",
        "Por mes": "Mes",
    }

    with st.container(key="card_tabla"):
        vista = st.radio("Ver detalle:", list(VISTAS_TABLA), horizontal=True, key="vista_tabla")
        col_t = VISTAS_TABLA[vista]
        st.markdown(f'<div class="ct">Detalle {vista.lower()}</div>', unsafe_allow_html=True)

        rt = resumen_por(base_dim(g_p, col_t), col_t)
        if col_t == "Mes":
            rt = rt.sort_values("Mes")
            rt["Mes"] = rt["Mes"].map(lambda m: mes_label(pd.Timestamp(m)))
        else:
            rt = rt.sort_values("SO", ascending=False)
        if col_t == "Proveedor":
            st.caption("Excluye ALMACEN LIMA (despachos del stock); sus montos sí están en los totales.")

        tabla = rt.copy()
        tabla["Fill Rate %"] = tabla["Fill Rate"] * 100
        tabla["Eficiencia %"] = tabla["Eficiencia"] * 100
        tabla["Confirmación %"] = tabla["Confirmación"] * 100
        tabla = tabla[[col_t, "SO", "RN", "Fill Rate %", "OF", "Eficiencia %", "RNO", "Confirmación %",
                       "Brecha compras", "Brecha despacho", "Brecha confirmación"]]
        tabla = tabla.rename(columns={"SO": "Solicitado", "RN": "Recepcionado", "OF": "Despachado", "RNO": "Confirmado"})
        st.dataframe(
            tabla, hide_index=True,
            column_config={
                "Solicitado": st.column_config.NumberColumn(f"Solicitado ({MONEDA})", format="%.0f"),
                "Recepcionado": st.column_config.NumberColumn(f"Recepcionado ({MONEDA})", format="%.0f"),
                "Despachado": st.column_config.NumberColumn(f"Despachado ({MONEDA})", format="%.0f"),
                "Brecha compras": st.column_config.NumberColumn(f"Brecha compras ({MONEDA})", format="%.0f"),
                "Brecha despacho": st.column_config.NumberColumn(f"Brecha despacho ({MONEDA})", format="%.0f"),
                "Confirmado": st.column_config.NumberColumn(f"Confirmado ({MONEDA})", format="%.0f"),
                "Brecha confirmación": st.column_config.NumberColumn(f"Sin confirmar ({MONEDA})", format="%.0f"),
                "Fill Rate %": st.column_config.ProgressColumn("Fill Rate", min_value=0, max_value=100, format="%.1f%%"),
                "Eficiencia %": st.column_config.ProgressColumn("Eficiencia", min_value=0, max_value=100, format="%.1f%%"),
                "Confirmación %": st.column_config.ProgressColumn("Confirmación", min_value=0, max_value=100, format="%.1f%%"),
            },
        )
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            tabla.to_excel(w, index=False, sheet_name=vista[:31])
        st.download_button("⬇️ Descargar tabla (Excel)", buf.getvalue(), f"detalle_{col_t.lower()}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with st.expander("¿Cómo se calculan los KPIs?"):
        st.markdown(
            f"""
- **Todo es valorizado** con las columnas *Total Price* ({MONEDA}), consolidadas por línea (pedido + ítem).
- **Fill Rate Compras** = RN recepcionado ÷ SO solicitado.
- **Eficiencia de Despacho** = OF despachado ÷ RN recepcionado.
- **Confirmación en Operación** = RNO confirmado por operaciones ÷ OF despachado por el almacén central.
- **Brecha de compras** = SO − RN. **Saldo por despachar** = RN − OF. **Despacho sin confirmar** = OF − RNO (puede incluir mercadería aún en tránsito).
- Tope 100% por línea: el RN de una línea no cuenta por encima de su SO, el OF no cuenta por encima de su RN y el RNO no cuenta por encima de su OF.
- **Cumplimiento por cantidad** = suma de unidades atendidas ÷ suma de unidades pedidas, sin distinguir la unidad de medida (con tope por línea). Ejemplo: si se piden 10 und de leche y 90 de aceite y se atienden 5 y 45, da 50 %. Compras: RN ÷ SO; despacho: OF ÷ RN; confirmación: RNO ÷ OF.
- **Solo entregas vencidas** (casillero en los filtros): mide únicamente las líneas cuya fecha de entrega (SO Delivery Date) es menor o igual a la fecha de corte ({corte:%d/%m/%Y}). Desmarcado, incluye las entregas futuras y los indicadores bajan porque aún no les corresponde llegar.
- Al pasar el puntero sobre los valorizados de Recepcionado, Brecha, Despachado, Por despachar, Confirmado y Sin confirmar (ventanas de eficiencia) se muestra el % que representan de su base.
- **Filtros:** fecha de entrega (día o rango), periodo (mes de entrega), familia, almacén (SOE Site), SO Site y casillero de entregas vencidas. Con filtro de fecha activo no se muestra la variación contra el mes anterior.
- **Sitios:** los almacenes (p. ej. NEWREST ALMACEN LIMA) se eliminan del análisis por sitio.
- **ALMACEN LIMA** (despachos del stock) cuenta en todos los montos y KPIs generales, pero no aparece como proveedor en el ranking ni en la tabla por proveedor.
- **Reparto del RNO:** un mismo RNO puede confirmar varios OF de distintas líneas de pedido, pero en la base su monto queda solo en la primera fila. El dashboard lo reparte entre esas líneas en proporción a su OF.
- **Ventanas de eficiencia:** haz clic en el monto de Recepcionado, Despachado o Confirmado (franja superior) para ver su gauge, donut y distribución.
            """
        )

    return locals()


# =====================================================
# REPORTE 2: CUMPLIMIENTO PROVEEDORES (PO vs RN)
# Estados: Atendida · Atendida parcial · Pendiente de atención
# =====================================================
def kpi_html(icono, color, titulo, valor, sub):
    return (f'<div class="kitem"><div class="kico" style="background:{color}22">{icono}</div>'
            f'<div class="klab">{titulo}</div><div class="kval">{valor}</div><div class="ksub">{sub}</div></div>')


def reporte_proveedores():
    cont_hero = st.container(key="card_hero")

    base = clasificar_items(items_po, corte)  # solo POs vencidas (PO Delivery Date <= corte)
    if base.empty:
        st.warning("No hay POs vencidas para evaluar.")
        st.stop()
    en_plazo = items_po.loc[items_po["Fecha PO"] > corte, "PO"].nunique()

    f_min, f_max = base["Fecha PO"].min().date(), base["Fecha PO"].max().date()
    meses_base = sorted(base["Mes"].dropna().unique())
    opciones_mes = ["Todo el periodo"] + [mes_label(pd.Timestamp(m)) for m in reversed(meses_base)]
    mapa_mes = {mes_label(pd.Timestamp(m)): pd.Timestamp(m) for m in meses_base}

    with st.container(key="card_filtros"):
        f1, f2, f3, f4, f5 = st.columns([1.35, 1, 1.7, 1, 1])
        with f1:
            rango = st.date_input(
                "Fecha de entrega (PO Delivery Date)", value=(f_min, f_max), min_value=f_min, max_value=f_max,
                key="p_fecha", help="Día o rango. Para un día puntual haz clic dos veces en la misma fecha.",
            )
        with f2:
            periodo = st.selectbox("Periodo (mes de entrega)", opciones_mes, key="p_periodo")
        with f3:
            sel_prov = st.multiselect("Proveedor", sorted(base["Proveedor"].unique()), key="p_prov", placeholder="Todos")
        with f4:
            sel_fam = st.multiselect("Familia", sorted(base["Familia"].unique()), key="p_fam", placeholder="Todas")
        with f5:
            sel_alm = st.multiselect("Almacén (SOE Site)", sorted(base["Almacén"].unique()), key="p_alm",
                                     placeholder="Todos")

    # --- Aplicación de filtros ---
    it_ctx = base
    if sel_prov:
        it_ctx = it_ctx[it_ctx["Proveedor"].isin(sel_prov)]
    if sel_fam:
        it_ctx = it_ctx[it_ctx["Familia"].isin(sel_fam)]
    if sel_alm:
        it_ctx = it_ctx[it_ctx["Almacén"].isin(sel_alm)]
    if it_ctx.empty:
        st.warning("No hay datos con los filtros seleccionados.")
        st.stop()

    it_f = it_ctx
    if periodo != "Todo el periodo":
        it_f = it_f[it_f["Mes"] == mapa_mes[periodo]]
    if isinstance(rango, (tuple, list)):
        d_ini = rango[0]
        d_fin = rango[1] if len(rango) == 2 else rango[0]  # con un solo clic se toma ese día
    else:
        d_ini = d_fin = rango
    etiqueta = periodo
    if (d_ini, d_fin) != (f_min, f_max):
        it_f = it_f[(it_f["Fecha PO"].dt.date >= d_ini) & (it_f["Fecha PO"].dt.date <= d_fin)]
        etiqueta = f"{d_ini:%d/%m/%Y}" if d_ini == d_fin else f"{d_ini:%d/%m/%Y} – {d_fin:%d/%m/%Y}"
    if it_f.empty:
        st.warning("No hay POs con esa combinación de filtros.")
        st.stop()

    pos = resumir_pos(it_f)
    K = kpis_po(pos, it_f)

    # --- Encabezado ---
    partes = []
    if sel_prov:
        partes.append(f"Proveedor: {', '.join(sel_prov)}")
    if sel_fam:
        partes.append(f"Familia: {', '.join(sel_fam)}")
    if sel_alm:
        partes.append(f"Almacén: {', '.join(sel_alm)}")
    filtros_txt = " · ".join(partes) or "Todos los proveedores, familias y almacenes"
    with cont_hero:
        st.markdown(
            f'<p class="hero-t" title="Data, Análisis, Visibilidad, Núcleo operativo, Control e Inteligencia logística">'
            f'📊 Da Vinci - Gestión de Abastecimiento</p>'
            f'<p class="hero-s">Reporte: {REPORTES[1]} · ¿Se atendió la PO completa? (RN vs PO) · '
            f'Corte {corte:%d/%m/%Y} · solo POs vencidas, es decir con PO Delivery Date ya cumplida '
            f'({en_plazo:,} POs aún en plazo no se evalúan) · {filtros_txt}</p>'
            f'<span class="hero-tag">{etiqueta}</span>',
            unsafe_allow_html=True,
        )

    # --- Franja de KPIs ---
    kpi_defs = [
        ("📋", NAVY, "POs evaluadas", f"{K['pos']:,}", f"{K['items']:,} ítems"),
        ("✅", GREEN, "POs atendidas", fmt_pct(K["p_at"]), f"{K['pos_at']:,} de {K['pos']:,} POs"),
        ("🧩", AMBAR, "POs atendidas parciales", fmt_pct(K["p_parc"]), f"{K['pos_parc']:,} POs"),
        ("⏳", RED, "POs pendientes de atención", fmt_pct(K["p_pend"]), f"{K['pos_pend']:,} POs sin RN"),
        ("📦", NAVY, "Cumplimiento por cantidad", fmt_pct(K["cq"]), "unidades atendidas ÷ pedidas"),
        ("💰", GREEN, "Cumplimiento valorizado", fmt_pct(K["cv"]), f"pendiente {fmt_monto(K['monto_pend'])}"),
    ]
    with st.container(key="card_kpis"):
        kc = st.columns(len(kpi_defs))
        for col, (ic, co, ti, va, su) in zip(kc, kpi_defs):
            with col:
                st.markdown(kpi_html(ic, co, ti, va, su), unsafe_allow_html=True)

    # --- Fila A: cumplimiento de POs, estado de POs, estado de ítems ---
    cnt_po = pos["Estado PO"].value_counts().reindex(ESTADOS_PO).fillna(0).astype(int)
    cnt_it = it_f["Estado ítem"].value_counts().reindex(ESTADOS_ITEM).fillna(0).astype(int)

    a1, a2, a3 = st.columns([1.5, 1, 1.1])
    with a1:
        with st.container(key="card_p_gauge"):
            st.markdown('<div class="ct">Cumplimiento de POs</div>'
                        '<div class="cs">PO atendida: todos sus ítems con RN y toda la cantidad pedida</div>',
                        unsafe_allow_html=True)
            c_g, c_s = st.columns([1.15, 1])
            with c_g:
                gauge(K["p_at"], META_PROVEEDORES, "pv_gauge")
            with c_s:
                st.markdown(stats([
                    ("POs evaluadas", f"{K['pos']:,}"),
                    ("Atendidas", f"{K['pos_at']:,}", pct_de(K["pos_at"], K["pos"], "total de POs")),
                    ("Atendidas parciales", f"{K['pos_parc']:,}", pct_de(K["pos_parc"], K["pos"], "total de POs")),
                    ("Pendientes de atención", f"{K['pos_pend']:,}", pct_de(K["pos_pend"], K["pos"], "total de POs")),
                    ("Meta", f"{META_PROVEEDORES}%"),
                ]), unsafe_allow_html=True)
    with a2:
        with st.container(key="card_p_estado_po"):
            st.markdown('<div class="ct">Estado de las POs</div>', unsafe_allow_html=True)
            donut_estados(list(cnt_po.index), list(cnt_po.values), [COL_ESTADO_PO[e] for e in cnt_po.index],
                          fmt_pct(K["p_at"], 0), "POs atendidas", "pv_estado_po")
    with a3:
        with st.container(key="card_p_estado_item"):
            st.markdown('<div class="ct">Estado de los ítems de las POs</div>', unsafe_allow_html=True)
            donut_estados(list(cnt_it.index), list(cnt_it.values), [COL_ESTADO_ITEM[e] for e in cnt_it.index],
                          fmt_pct(K["p_items_at"], 0), "ítems atendidos", "pv_estado_item")

    # --- Fila B: proveedores con más POs sin atender completas y mayor monto pendiente ---
    rp = resumen_proveedor(pos)
    b1, b2 = st.columns(2)
    with b1:
        with st.container(key="card_p_rank_inc"):
            st.markdown('<div class="ct">Proveedores con más POs sin atender completas · Top 10</div>',
                        unsafe_allow_html=True)
            d = rp[rp["Incompletas"] > 0].sort_values("Incompletas", ascending=False).head(10)
            fig = go.Figure()
            for nombre, col_d, color in [("Atendidas parciales", "Parciales", AMBAR),
                                         ("Pendientes de atención", "Pendientes", RED)]:
                fig.add_trace(go.Bar(
                    y=d["Proveedor"], x=d[col_d], name=nombre, orientation="h", marker_color=color,
                    customdata=np.stack([d["POs"], d["% atendidas"] * 100], axis=-1),
                    hovertemplate="%{y}<br>" + nombre + ": %{x}<br>POs evaluadas: %{customdata[0]}"
                                  "<br>% atendidas: %{customdata[1]:.1f}%<extra></extra>",
                ))
            base_layout(fig, h=380, legend=True)
            fig.update_layout(barmode="stack", legend=dict(orientation="h", y=-0.12),
                              margin=dict(l=8, r=20, t=8, b=40))
            fig.update_yaxes(autorange="reversed", title=None, tickfont=dict(size=11))
            fig.update_xaxes(showgrid=True, gridcolor="#EEF1F4", title="N° de POs")
            grafico(fig, "pv_rank_inc")
    with b2:
        with st.container(key="card_p_rank_monto"):
            st.markdown('<div class="ct">Mayor monto pendiente por proveedor · Top 10</div>', unsafe_allow_html=True)
            d = rp[rp["Monto pendiente"] > 0].sort_values("Monto pendiente", ascending=False).head(10)
            fig = go.Figure(go.Bar(
                x=d["Monto pendiente"], y=d["Proveedor"], orientation="h", marker_color=RED,
                text=[fmt_monto(v) for v in d["Monto pendiente"]], textposition="outside", cliponaxis=False,
                customdata=np.stack([d["% cantidad"] * 100, d["% valorizado"] * 100, d["POs"]], axis=-1),
                hovertemplate="%{y}<br>Pendiente: " + MONEDA + " %{x:,.0f}<br>Cumpl. cantidad: %{customdata[0]:.1f}%"
                              "<br>Cumpl. valorizado: %{customdata[1]:.1f}%<br>POs: %{customdata[2]}<extra></extra>",
            ))
            base_layout(fig, h=380)
            fig.update_layout(margin=dict(l=8, r=80, t=8, b=8))
            fig.update_yaxes(autorange="reversed", title=None, tickfont=dict(size=11))
            fig.update_xaxes(showgrid=True, gridcolor="#EEF1F4", showticklabels=False)
            grafico(fig, "pv_rank_monto")

    # --- Fila C: tendencia mensual (ignora periodo y fecha; respeta proveedor, familia y almacén) ---
    pos_ctx = resumir_pos(it_ctx)
    tm = resumen_proveedor(pos_ctx, "Mes").sort_values("Mes")
    etq = [mes_label(pd.Timestamp(m)) for m in tm["Mes"]]
    est_mes = (pos_ctx.groupby(["Mes", "Estado PO"], observed=False).size().unstack(fill_value=0)
               .reindex(columns=ESTADOS_PO, fill_value=0).sort_index())
    etq_mes = [mes_label(pd.Timestamp(m)) for m in est_mes.index]

    c1, c2 = st.columns(2)
    with c1:
        with st.container(key="card_p_trend"):
            st.markdown('<div class="ct">Tendencia mensual (por mes de PO Delivery Date)</div>', unsafe_allow_html=True)
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=etq, y=tm["% atendidas"] * 100, name="POs atendidas",
                                     mode="lines+markers", line=dict(color=GREEN, width=3), marker=dict(size=8)))
            fig.add_trace(go.Scatter(x=etq, y=tm["% cantidad"] * 100, name="Cumplimiento por cantidad",
                                     mode="lines+markers", line=dict(color=NAVY, width=3), marker=dict(size=8)))
            fig.add_trace(go.Scatter(x=etq, y=tm["% pendientes"] * 100, name="POs pendientes de atención",
                                     mode="lines+markers", line=dict(color=RED, width=3), marker=dict(size=8)))
            fig.add_hline(y=META_PROVEEDORES, line_dash="dot", line_color=MUTED,
                          annotation_text=f"Meta {META_PROVEEDORES}%", annotation_position="bottom right")
            base_layout(fig, h=360, legend=True)
            fig.update_yaxes(ticksuffix="%", gridcolor="#EEF1F4", rangemode="tozero")
            fig.update_layout(legend=dict(orientation="h", y=-0.2))
            grafico(fig, "pv_trend")
    with c2:
        with st.container(key="card_p_estados_mes"):
            st.markdown('<div class="ct">POs por estado y mes</div>', unsafe_allow_html=True)
            fig = go.Figure()
            for e in ESTADOS_PO:
                fig.add_trace(go.Bar(x=etq_mes, y=est_mes[e], name=e, marker_color=COL_ESTADO_PO[e]))
            base_layout(fig, h=360, legend=True)
            fig.update_layout(barmode="stack", legend=dict(orientation="h", y=-0.2))
            fig.update_yaxes(title="N° de POs", gridcolor="#EEF1F4")
            grafico(fig, "pv_estados_mes")

    # --- Fila D: tabla por proveedor ---
    with st.container(key="card_p_tabla"):
        st.markdown('<div class="ct">Cumplimiento por proveedor</div>', unsafe_allow_html=True)
        tp = rp.sort_values(["Monto pendiente", "POs"], ascending=False).copy()
        for c in ["% atendidas", "% cantidad", "% valorizado"]:
            tp[c] = tp[c] * 100
        tp = tp[["Proveedor", "POs", "Atendidas", "Parciales", "Pendientes", "% atendidas", "% cantidad",
                 "% valorizado", "Monto PO", "Monto pendiente", "Días vencida prom."]]
        tp = tp.rename(columns={"Parciales": "Atendidas parciales", "Pendientes": "Pendientes de atención"})
        st.dataframe(
            tp, hide_index=True,
            column_config={
                "% atendidas": st.column_config.ProgressColumn("% POs atendidas", min_value=0, max_value=100, format="%.1f%%"),
                "% cantidad": st.column_config.ProgressColumn("% cantidad", min_value=0, max_value=100, format="%.1f%%"),
                "% valorizado": st.column_config.ProgressColumn("% valorizado", min_value=0, max_value=100, format="%.1f%%"),
                "Monto PO": st.column_config.NumberColumn(f"Monto PO ({MONEDA})", format="%.0f"),
                "Monto pendiente": st.column_config.NumberColumn(f"Monto pendiente ({MONEDA})", format="%.0f"),
                "Días vencida prom.": st.column_config.NumberColumn("Días vencida prom.", format="%.1f"),
            },
        )
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            tp.to_excel(w, index=False, sheet_name="Por proveedor")
        st.download_button("⬇️ Descargar tabla (Excel)", buf.getvalue(), "cumplimiento_por_proveedor.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="p_dl_prov")

    # --- Fila E: detalle de POs y de sus ítems ---
    with st.container(key="card_p_detalle"):
        st.markdown('<div class="ct">Detalle de POs</div>', unsafe_allow_html=True)
        sel_estado = st.multiselect(
            "Estado de la PO", ESTADOS_PO, default=ESTADOS_PO[1:], key="p_estado",
            help="Por defecto se muestran las POs que no se atendieron completas.",
        )
        det = pos[pos["Estado PO"].astype(str).isin(sel_estado)].sort_values(
            ["Días vencida", "Monto pendiente"], ascending=False)
        tpo = det[["PO", "Proveedor", "Fecha PO", "Almacén", "Ítems", "Ítems atendidos", "Ítems parciales",
                   "Ítems pendientes", "Estado PO", "Tipo de parcialidad", "% cantidad", "Monto PO",
                   "Monto pendiente", "Días vencida"]].copy()
        tpo["Estado PO"] = tpo["Estado PO"].astype(str)
        tpo["% cantidad"] = tpo["% cantidad"] * 100
        tpo["Fecha PO"] = tpo["Fecha PO"].dt.date
        st.caption(f"{len(tpo):,} POs con los estados seleccionados")
        st.dataframe(
            tpo, hide_index=True,
            column_config={
                "% cantidad": st.column_config.ProgressColumn("% cantidad", min_value=0, max_value=100, format="%.1f%%"),
                "Monto PO": st.column_config.NumberColumn(f"Monto PO ({MONEDA})", format="%.0f"),
                "Monto pendiente": st.column_config.NumberColumn(f"Monto pendiente ({MONEDA})", format="%.0f"),
                "Días vencida": st.column_config.NumberColumn("Días vencida", format="%.0f"),
            },
        )
        buf2 = io.BytesIO()
        with pd.ExcelWriter(buf2, engine="openpyxl") as w:
            tpo.to_excel(w, index=False, sheet_name="POs")
        st.download_button("⬇️ Descargar POs (Excel)", buf2.getvalue(), "detalle_pos.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="p_dl_pos")

        if not det.empty:
            opciones_po = det["PO"].head(1000).tolist()
            info = det.set_index("PO")[["Proveedor", "Estado PO"]]
            po_sel = st.selectbox(
                "Ver los ítems de una PO", opciones_po, key=f"p_po_{len(opciones_po)}_{opciones_po[0]}",
                format_func=lambda x: f"{x} · {info.loc[x, 'Proveedor']} · {info.loc[x, 'Estado PO']}",
            )
            its = it_f[it_f["PO"] == po_sel].copy()
            tit = its[["Ítem", "Unidad", "Cant PO", "Cant RN", "Faltante", "% cumpl", "Estado ítem",
                       "F primera RN", "F última RN", "Días vencida"]].copy()
            tit["% cumpl"] = tit["% cumpl"] * 100
            tit["Estado ítem"] = tit["Estado ítem"].astype(str)
            tit["F primera RN"] = tit["F primera RN"].dt.date
            tit["F última RN"] = tit["F última RN"].dt.date
            st.dataframe(
                tit.sort_values("Faltante", ascending=False), hide_index=True,
                column_config={
                    "% cumpl": st.column_config.ProgressColumn("% cumplimiento", min_value=0, max_value=100,
                                                               format="%.1f%%"),
                    "Días vencida": st.column_config.NumberColumn("Días vencida", format="%.0f"),
                },
            )

    with st.expander("¿Cómo se evalúa el cumplimiento de proveedores?"):
        st.markdown(
            f"""
- **Qué se mide:** si cada proveedor atendió, con su **RN** (recepción en almacén), lo que pidió la **PO**.
- **Se evalúan solo las POs vencidas** (PO Delivery Date menor o igual a la fecha de corte, {corte:%d/%m/%Y}). Las POs aún en plazo no cuentan.
- **Estados de la PO:** *Atendida* (todos sus ítems con RN y por toda la cantidad pedida), *Atendida parcial* (se atendió algo, pero no todo) y *Pendiente de atención* (ningún ítem tiene RN).
- **Tipo de parcialidad** (columna del detalle): *Cantidades incompletas* (todos los ítems tienen RN, pero alguno por menos cantidad que la PO), *Faltan ítems* (algunos ítems sin RN y el resto completos) o *Faltan ítems y cantidades incompletas* (ambas situaciones).
- **Estados del ítem:** *Atendido* (RN ≥ PO), *Atendido parcial* (RN menor que la PO) y *Pendiente de atención* (sin RN).
- **Cumplimiento por cantidad** = unidades atendidas (RN, con tope en lo pedido por ítem) ÷ unidades pedidas, sin distinguir la unidad de medida. **Cumplimiento valorizado** = monto de la PO atendido (valorizado al precio de la PO) ÷ monto de la PO.
- **Días vencida** = días transcurridos entre el PO Delivery Date y la fecha de corte para lo que sigue sin atender completo.
- No se consideran proveedores los despachos del stock propio (ALMACEN LIMA), las POs inválidas ni las RN inválidas.
- Con el filtro de **Familia** o **Almacén**, la PO se evalúa solo con los ítems que cumplen ese filtro.
            """
        )
    return locals()


# =====================================================
# SELECTOR DE REPORTES (lista fija a la derecha)
# =====================================================
with st.container(key="rail_reportes"):
    st.markdown('<div class="rail-t">📑 Reportes</div>', unsafe_allow_html=True)
    reporte = st.radio("Reportes", REPORTES, key="reporte_sel", label_visibility="collapsed")

if reporte == REPORTES[0]:
    _resultado = reporte_fill_rate()
else:
    _resultado = reporte_proveedores()
