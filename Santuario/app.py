"""
Geovisualizador del Santuario de la Naturaleza Río Cruces (Valdivia, Chile)
Streamlit + Folium

Estructura esperada del repositorio:
    app.py
    requirements.txt
    data/
        Fuentes_y_proyectos_Santuario_Rio_Cruces.xlsx
        area_influencia_conaf2019_aprox.geojson
        santuario_rio_cruces.geojson      <- (opcional) límite real del santuario
"""

import html
import json
import re
from pathlib import Path

import folium
import pandas as pd
import streamlit as st
from folium.plugins import Fullscreen, MarkerCluster
from streamlit_folium import st_folium

# ──────────────────────────────────────────────────────────────────────────────
# 1. CONFIGURACIÓN
# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Santuario Río Cruces · Geovisualizador",
    page_icon="🦢",
    layout="wide",
)

BASE = Path(__file__).parent
ARCHIVO_EXCEL = "Fuentes_y_proyectos_Santuario_Rio_Cruces.xlsx"
ARCHIVO_INFLUENCIA = "area_influencia_conaf2019_aprox.geojson"
ARCHIVO_SANTUARIO = "santuario_rio_cruces.geojson"
HOJA_PROYECTOS = "Proyectos influencia"

CENTRO_SANTUARIO = [-39.69, -73.17]  # Valdivia, Santuario Río Cruces
ZOOM_INICIAL = 10

COLOR_SANTUARIO = "#87CEFA"   # celeste claro
COLOR_INFLUENCIA = "#4B0F7A"  # morado oscuro

# Límite PROVISORIO (aprox. 4.900 ha) usado solo si no existe data/santuario_rio_cruces.geojson
# Reemplázalo por tu capa real (DS 41/2015 o tu "Límite del santuario") en EPSG:4326.
SANTUARIO_APROX = [
    (-73.0927, -39.6125), (-73.1169, -39.6722), (-73.1502, -39.7058), (-73.1924, -39.7315),
    (-73.2242, -39.7461), (-73.2350, -39.7776), (-73.2443, -39.7815), (-73.2532, -39.7791),
    (-73.2572, -39.7707), (-73.2456, -39.7352), (-73.2070, -39.7161), (-73.1688, -39.6933),
    (-73.1393, -39.6637), (-73.1162, -39.6079), (-73.1114, -39.6039), (-73.1044, -39.6025),
    (-73.0959, -39.6056), (-73.0927, -39.6125),
]

# Solo Valdivia y Mariquina (se puede apagar desde la barra lateral)
COMUNAS_PERMITIDAS = ["Valdivia", "Mariquina"]

# ── Tipo de proyecto → (icono Font Awesome 6, color del pin, etiqueta) ─────────
# Para cambiar un icono basta con editar este diccionario
# (catálogo: https://fontawesome.com/search?o=r&m=free)
TIPOS = {
    "Inmobiliario de borde":    {"icono": "fa-house",              "color": "#E67E22", "etiqueta": "Inmobiliario"},
    "Acuicultura":              {"icono": "fa-fish",               "color": "#1ABC9C", "etiqueta": "Acuicultura"},
    "Agroindustria":            {"icono": "fa-tractor",            "color": "#7D6608", "etiqueta": "Agroindustria"},
    "Combustibles":             {"icono": "fa-fire-flame-curved",  "color": "#C0392B", "etiqueta": "Combustibles"},
    "Drenaje":                  {"icono": "fa-grip-lines",         "color": "#5D6D7E", "etiqueta": "Drenaje"},
    "Forestal industrial":      {"icono": "fa-tree",               "color": "#1E8449", "etiqueta": "Forestal"},
    "Oro":                      {"icono": "fa-coins",              "color": "#B7950B", "etiqueta": "Oro"},
    "PTAS":                     {"icono": "fa-droplet",            "color": "#2E86C1", "etiqueta": "PTAS (plantas de tratamiento)"},
    "Planta Valdivia (Arauco)": {"icono": "fa-industry",           "color": "#34495E", "etiqueta": "Planta (Arauco Valdivia)"},
    "RILes":                    {"icono": "RILES",                 "color": "#117A65", "etiqueta": "RILes"},
    "Residuos":                 {"icono": "fa-dumpster",           "color": "#6E2C00", "etiqueta": "Residuos"},
    "Solar":                    {"icono": "fa-sun",                "color": "#D4AC0D", "etiqueta": "Solar"},
    "Tendido eléctrico":        {"icono": "fa-bolt",               "color": "#8E44AD", "etiqueta": "Tendido eléctrico"},
    "Áridos":                   {"icono": "fa-hammer",             "color": "#A04000", "etiqueta": "Áridos"},
    "Sin vía directa":          {"icono": "",                      "color": "#FF69B4", "etiqueta": "Sin vía directa"},
}

ESTADOS = ["Aprobado", "No Admitido a Tramitación", "Desistido", "No calificado", "Rechazado"]
COLOR_ESTADO = {
    "Aprobado": "#27AE60",
    "No Admitido a Tramitación": "#7F8C8D",
    "Desistido": "#E67E22",
    "No calificado": "#2980B9",
    "Rechazado": "#C0392B",
}

SIN_OBJETO = "Sin objeto afectado"
# Si quieres unificar nombres de objetos de conservación, agrégalos aquí. Ej:
# {"Fondos someros": "Fondos someros y totorales"}
NORMALIZAR_OBJETOS = {}

FA_CSS = "https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css"

# ──────────────────────────────────────────────────────────────────────────────
# 2. CARGA DE DATOS
# ──────────────────────────────────────────────────────────────────────────────
def ubicar(nombre: str) -> Path | None:
    for carpeta in (BASE / "data", BASE):
        ruta = carpeta / nombre
        if ruta.exists():
            return ruta
    return None


def separar_objetos(texto) -> list[str]:
    """'Columna de agua, fondos someros (bajo)' -> ['Columna de agua', 'Fondos someros']"""
    if not isinstance(texto, str) or texto.strip() in ("", "—", "-"):
        return []
    texto = re.sub(r"\(.*?\)", "", texto)  # quita '(bajo)'
    salida = []
    for parte in texto.split(","):
        parte = parte.strip()
        if not parte:
            continue
        parte = parte[0].upper() + parte[1:]
        parte = NORMALIZAR_OBJETOS.get(parte, parte)
        if parte not in salida:
            salida.append(parte)
    return salida


@st.cache_data(show_spinner="Cargando proyectos…")
def cargar_proyectos() -> pd.DataFrame:
    ruta = ubicar(ARCHIVO_EXCEL)
    if ruta is None:
        st.error(f"No encuentro **{ARCHIVO_EXCEL}** en /data. Súbelo al repositorio.")
        st.stop()
    df = pd.read_excel(ruta, sheet_name=HOJA_PROYECTOS)
    df = df.rename(columns={
        "Nombre del proyecto": "nombre",
        "Tipo (clasificación del grupo)": "tipo",
        "Estado": "estado",
        "Titular": "titular",
        "Comuna": "comuna",
        "Latitud": "lat",
        "Longitud": "lon",
        "Vía hacia el santuario": "via",
        "Objetos de conservación afectados (PM, Tabla 17)": "objetos_txt",
        "Link expediente": "link",
    })
    df = df.dropna(subset=["lat", "lon"]).copy()
    for c in ("tipo", "estado", "comuna", "nombre", "titular", "via"):
        df[c] = df[c].astype(str).str.strip()
    df["objetos"] = df["objetos_txt"].apply(separar_objetos)
    return df


@st.cache_data
def cargar_geojson(nombre: str):
    ruta = ubicar(nombre)
    if ruta is None:
        return None
    with open(ruta, encoding="utf-8") as f:
        return json.load(f)


def geojson_santuario():
    real = cargar_geojson(ARCHIVO_SANTUARIO)
    if real is not None:
        return real, True
    aprox = {
        "type": "FeatureCollection",
        "features": [{
            "type": "Feature",
            "properties": {"nombre": "Santuario de la Naturaleza Río Cruces (límite provisorio)"},
            "geometry": {"type": "Polygon", "coordinates": [[list(p) for p in SANTUARIO_APROX]]},
        }],
    }
    return aprox, False


# ──────────────────────────────────────────────────────────────────────────────
# 3. ICONOS (pines con Font Awesome)
# ──────────────────────────────────────────────────────────────────────────────
def html_icono(icono: str) -> str:
    """HTML del símbolo que va dentro del pin."""
    if icono == "RILES":  # basurero con gota de agua
        return (
            '<span class="fa-stack" style="font-size:9px;width:2.2em;height:2.2em;line-height:2.2em;">'
            '<i class="fa-solid fa-trash-can fa-stack-2x"></i>'
            '<i class="fa-solid fa-droplet fa-stack-1x" style="color:#5DADE2;font-size:.55em;'
            'top:.18em;"></i></span>'
        )
    if not icono:
        return ""
    return f'<i class="fa-solid {icono}"></i>'


def pin(tipo: str) -> folium.DivIcon:
    cfg = TIPOS[tipo]
    cuerpo = (
        f'<div style="position:relative;width:34px;height:42px;">'
        f'<div style="position:absolute;left:3px;top:0;width:28px;height:28px;background:{cfg["color"]};'
        f'border:2px solid #fff;border-radius:50% 50% 50% 0;transform:rotate(-45deg);'
        f'box-shadow:0 2px 5px rgba(0,0,0,.45);"></div>'
        f'<div style="position:absolute;left:3px;top:0;width:28px;height:28px;display:flex;'
        f'align-items:center;justify-content:center;color:#fff;font-size:13px;">'
        f'{html_icono(cfg["icono"])}</div></div>'
    )
    return folium.DivIcon(html=cuerpo, icon_size=(34, 42), icon_anchor=(17, 40), popup_anchor=(0, -36))


def popup_html(r) -> str:
    esc = html.escape
    color = COLOR_ESTADO.get(r["estado"], "#555")
    objetos = ", ".join(r["objetos"]) if r["objetos"] else "Sin objeto afectado"
    link = ""
    if isinstance(r.get("link"), str) and r["link"].startswith("http"):
        link = f'<a href="{esc(r["link"])}" target="_blank">Ver expediente SEIA ↗</a>'
    return f"""
    <div style="font-family:Arial,sans-serif;font-size:12.5px;width:310px;line-height:1.4;">
      <div style="font-weight:700;font-size:13.5px;margin-bottom:6px;">{esc(r["nombre"])}</div>
      <div><b>Titular:</b> {esc(r["titular"])}</div>
      <div><b>Comuna:</b> {esc(r["comuna"])}</div>
      <div><b>Vía hacia el santuario:</b> {esc(r["via"])}</div>
      <hr style="margin:6px 0;border:0;border-top:1px solid #ddd;">
      <div><b>Tipo:</b> {esc(TIPOS.get(r["tipo"], {}).get("etiqueta", r["tipo"]))}</div>
      <div><b>Estado:</b> <span style="color:{color};font-weight:700;">{esc(r["estado"])}</span></div>
      <div><b>Objetos afectados:</b> {esc(objetos)}</div>
      <div style="margin-top:4px;">{link}</div>
    </div>"""


# ──────────────────────────────────────────────────────────────────────────────
# 4. ESTADO DE LOS FILTROS (leyendas con checkbox)
# ──────────────────────────────────────────────────────────────────────────────
def fijar_todos(prefijo: str, claves: list[str], valor: bool):
    for c in claves:
        st.session_state[f"{prefijo}_{c}"] = valor


def leyenda_checks(prefijo: str, claves: list[str], etiquetas=None, iconos=None, notas=None, conteos=None):
    """Dibuja una leyenda con checkboxes y devuelve las claves marcadas."""
    c1, c2 = st.columns(2)
    c1.button("✔ Marcar todo", key=f"{prefijo}_on", width="stretch",
              on_click=fijar_todos, args=(prefijo, claves, True))
    c2.button("✖ Desmarcar", key=f"{prefijo}_off", width="stretch",
              on_click=fijar_todos, args=(prefijo, claves, False))
    marcadas = []
    for c in claves:
        k = f"{prefijo}_{c}"
        if k not in st.session_state:
            st.session_state[k] = True
        col_i, col_c = st.columns([1, 8], vertical_alignment="center")
        if iconos:
            col_i.markdown(iconos[c], unsafe_allow_html=True)
        n = f" ({conteos.get(c, 0)})" if conteos is not None else ""
        etiqueta = (etiquetas or {}).get(c, c) + n
        if st.checkbox(etiqueta, key=k):
            marcadas.append(c)
        if notas and notas.get(c):
            st.caption(notas[c])
    return marcadas


def mini_pin_html(tipo: str) -> str:
    cfg = TIPOS[tipo]
    return (
        f'<span style="display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;'
        f'background:{cfg["color"]};color:#fff;border-radius:50%;font-size:10px;">{html_icono(cfg["icono"])}</span>'
    )


def punto_estado_html(estado: str) -> str:
    return (f'<span style="display:inline-block;width:14px;height:14px;border-radius:50%;'
            f'background:{COLOR_ESTADO[estado]};"></span>')


# ──────────────────────────────────────────────────────────────────────────────
# 5. APLICACIÓN
# ──────────────────────────────────────────────────────────────────────────────
st.markdown(f"<style>@import url('{FA_CSS}');</style>", unsafe_allow_html=True)
st.markdown(
    "<style>.block-container{padding-top:1.2rem;} "
    "[data-testid='stSidebar'] .stCheckbox{margin-bottom:-0.6rem;}</style>",
    unsafe_allow_html=True,
)

df_total = cargar_proyectos()

st.title("🦢 Santuario de la Naturaleza Río Cruces")
st.caption("Proyectos del SEIA en el área de influencia · Valdivia y San José de la Mariquina, Región de Los Ríos")

# ── Opciones generales ────────────────────────────────────────────────────────
with st.sidebar:
    st.header("Filtros")
    solo_vm = st.toggle("Solo Valdivia y Mariquina", value=True)
    agrupar = st.toggle("Agrupar pines cercanos", value=True,
                        help="Hay proyectos con coordenada genérica repetida; al acercar el mapa se separan.")

df = df_total.copy()
if solo_vm:
    patron = "|".join(COMUNAS_PERMITIDAS)
    df = df[df["comuna"].str.contains(patron, case=False, na=False)]

# Objetos de conservación presentes en los datos
objetos_unicos = sorted({o for lista in df["objetos"] for o in lista})
claves_objetos = objetos_unicos + [SIN_OBJETO]

with st.sidebar:
    tab_tipo, tab_estado, tab_obj = st.tabs(["🧭 Tipos", "📋 Estado", "🌿 Objetos"])

    # --- Leyenda 1: tipos de proyecto -----------------------------------------
    with tab_tipo:
        orden = [t for t in TIPOS if t in set(df["tipo"])]
        tipos_marcados = leyenda_checks(
            "tipo", orden,
            etiquetas={t: TIPOS[t]["etiqueta"] for t in orden},
            iconos={t: mini_pin_html(t) for t in orden},
            conteos=df["tipo"].value_counts().to_dict(),
        )

    # --- Leyenda 2: estado de tramitación -------------------------------------
    with tab_estado:
        estados_presentes = [e for e in ESTADOS if e in set(df["estado"])]
        estados_marcados = leyenda_checks(
            "estado", estados_presentes,
            iconos={e: punto_estado_html(e) for e in estados_presentes},
            conteos=df["estado"].value_counts().to_dict(),
        )

    # --- Leyenda 3: objetos de conservación afectados -------------------------
    with tab_obj:
        st.caption("Objetos de conservación del Plan de Manejo (Tabla 17) que afecta cada proyecto. "
                   "Debajo de cada uno se indican los tipos de proyecto que lo afectan.")
        notas, conteo_obj = {}, {}
        for o in objetos_unicos:
            sub = df[df["objetos"].apply(lambda lista, o=o: o in lista)]
            conteo_obj[o] = len(sub)
            tipos_o = sorted(TIPOS[t]["etiqueta"] for t in sub["tipo"].unique() if t in TIPOS)
            notas[o] = "Afectado por: " + ", ".join(tipos_o)
        sub_sin = df[df["objetos"].apply(len) == 0]
        conteo_obj[SIN_OBJETO] = len(sub_sin)
        notas[SIN_OBJETO] = "Proyectos dentro del área CONAF sin vía directa al agua."
        objetos_marcados = leyenda_checks(
            "obj", claves_objetos, notas=notas, conteos=conteo_obj,
        )

# ── Aplicar filtros (se combinan: tipo Y estado Y objeto) ─────────────────────
def pasa_objetos(lista) -> bool:
    if not lista:
        return SIN_OBJETO in objetos_marcados
    return any(o in objetos_marcados for o in lista)


filtrado = df[
    df["tipo"].isin(tipos_marcados)
    & df["estado"].isin(estados_marcados)
    & df["objetos"].apply(pasa_objetos)
]

m1, m2, m3 = st.columns(3)
m1.metric("Proyectos mostrados", len(filtrado))
m2.metric("Aprobados", int((filtrado["estado"] == "Aprobado").sum()))
m3.metric("Tipos activos", filtrado["tipo"].nunique())

# ── Mapa ──────────────────────────────────────────────────────────────────────
mapa = folium.Map(location=CENTRO_SANTUARIO, zoom_start=ZOOM_INICIAL, tiles=None, control_scale=True)
mapa.get_root().header.add_child(folium.CssLink(FA_CSS))

# Dos capas base
folium.TileLayer("OpenStreetMap", name="OpenStreetMap", show=True).add_to(mapa)
folium.TileLayer(
    tiles="https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}",
    attr="Esri, Maxar, Earthstar Geographics",
    name="Satélite (Esri)",
    show=False,
).add_to(mapa)

# Área de influencia (morado oscuro)
influencia = cargar_geojson(ARCHIVO_INFLUENCIA)
if influencia:
    folium.GeoJson(
        influencia,
        name="Área de influencia (CONAF 2019, aprox.)",
        style_function=lambda _: {"color": COLOR_INFLUENCIA, "weight": 2.5,
                                  "fillColor": COLOR_INFLUENCIA, "fillOpacity": 0.18},
        tooltip="Área de influencia (CONAF 2019, aproximada)",
    ).add_to(mapa)
else:
    st.sidebar.warning(f"Falta data/{ARCHIVO_INFLUENCIA}")

# Santuario (celeste claro)
gj_santuario, es_real = geojson_santuario()
folium.GeoJson(
    gj_santuario,
    name="Santuario Río Cruces",
    style_function=lambda _: {"color": "#3C9BD6", "weight": 2.5,
                              "fillColor": COLOR_SANTUARIO, "fillOpacity": 0.55},
    tooltip="Santuario de la Naturaleza Río Cruces",
).add_to(mapa)
if not es_real:
    st.sidebar.info("El polígono del santuario es **provisorio**. Agrega `data/santuario_rio_cruces.geojson` "
                    "con el límite oficial para reemplazarlo.")

# Pines
contenedor = MarkerCluster(name="Proyectos", options={"disableClusteringAtZoom": 14,
                                                       "spiderfyOnMaxZoom": True}) if agrupar \
    else folium.FeatureGroup(name="Proyectos")
for _, r in filtrado.iterrows():
    if r["tipo"] not in TIPOS:
        continue
    folium.Marker(
        [r["lat"], r["lon"]],
        icon=pin(r["tipo"]),
        tooltip=r["nombre"][:90],
        popup=folium.Popup(popup_html(r), max_width=340),
    ).add_to(contenedor)
contenedor.add_to(mapa)

Fullscreen().add_to(mapa)
folium.LayerControl(collapsed=True, position="topright").add_to(mapa)

st_folium(mapa, height=680, use_container_width=True, returned_objects=[])

with st.expander("📄 Tabla de proyectos mostrados"):
    tabla = filtrado[["nombre", "titular", "comuna", "tipo", "estado", "via", "objetos_txt"]].rename(columns={
        "nombre": "Proyecto", "titular": "Titular", "comuna": "Comuna", "tipo": "Tipo",
        "estado": "Estado", "via": "Vía hacia el santuario", "objetos_txt": "Objetos afectados"})
    st.dataframe(tabla, width="stretch", hide_index=True)

st.caption("Fuentes: SEIA (seia.sea.gob.cl); Plan de Manejo MN y SN Río Cruces y Chorocamayo, CONAF 2019. "
           "Los puntos son la coordenada representativa del SEIA, no necesariamente la faena. "
           "El área de influencia está digitalizada de forma aproximada: solo para exploración.")
