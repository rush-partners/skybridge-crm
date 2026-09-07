"""Skybridge ERP/CRM — Cotizador de Importaciones.

App Streamlit con SQLite para gestionar Clientes y Cotizaciones,
replicando el motor de cálculo de 'Simulador Skybridge.xlsx'.
"""
import html
import math
import os
import re
import uuid
from datetime import date, datetime
from urllib.parse import quote

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

import db
import calculo
import crm
import storage
from pdf_export import generar_pdf_cotizacion

st.set_page_config(page_title="Skybridge ERP/CRM", page_icon="🚢", layout="wide")
if "_db_inicializada" not in st.session_state:
    # Antes corría en CADA rerun (cada click, cada tecla) — con Turso
    # comparte una sola conexión con estado de transacción por sesión
    # (ver db.get_connection), así que si dos reruns se solapaban (típico
    # con clicks rápidos), sus sentencias de init_db() se entrelazaban en
    # la misma transacción y tiraban "cannot start a transaction within a
    # transaction". Al ejecutarlo una sola vez por sesión de navegador se
    # elimina la fuente más frecuente de ese solape.
    db.init_db()
    st.session_state["_db_inicializada"] = True

# Paleta e identidad tomadas de skybridgecomex.com: navy oscuro + acento
# naranja, tipografía Inter, botones rectos (radio 4px) en mayúscula.
def _inyectar_estilos():
    # Modo oscuro: mismo naranja de acento en los dos modos (identidad de
    # marca), navy/slate en vez de blanco — mismos valores que [theme.dark]
    # en .streamlit/config.toml, así el toggle de más abajo y el menú nativo
    # ☰ > Settings > Choose app theme quedan visualmente idénticos.
    # Todo el resto del CSS de acá abajo usa var(--sb-*) sin condicionales:
    # al cambiar estos valores, toda la hoja de estilos seguida se adapta sola.
    oscuro = st.session_state.get("tema_oscuro", False)
    if oscuro:
        tokens = {
            "navy": "#E2E8F0", "orange": "#E8652A", "orange-dark": "#F2854D",
            "text-secondary": "#94A3B8", "border": "#334155",
            "surface": "#1E293B", "zone": "#16202E", "bg": "#0B1220",
        }
    else:
        tokens = {
            "navy": "#0F172A", "orange": "#E8652A", "orange-dark": "#C04E18",
            "text-secondary": "#64748B", "border": "#E2E8F0",
            "surface": "#FFFFFF", "zone": "#F1F5F9", "bg": "#F8FAFC",
        }
    variables_css = "\n".join(f"    --sb-{nombre}: {valor};" for nombre, valor in tokens.items())

    # Borde izquierdo de la card de contacto del CRM, uno por etapa — se
    # arma acá desde crm.COLOR_ETAPA/ETAPA_SLUG (única fuente de verdad,
    # también usada por la badge de etapa) para que nunca queden
    # desincronizados entre sí.
    etapa_css = "\n".join(
        f'        [class*="st-key-filacrm_{crm.ETAPA_SLUG[etapa]}_"] {{ border-left: 3px solid {color} !important; }}'
        for etapa, color in crm.COLOR_ETAPA.items()
    )

    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');

        /* Paleta de skybridgecomex.com: los mismos tokens de marca (navy,
        naranja) invierten su rol en modo oscuro — el navy que en modo claro
        es el color de texto principal pasa a ser cercano al fondo, y
        viceversa. Ver el bloque Python de arriba para los dos juegos de
        valores. */
        :root {
__SB_VARS__
        }

        html, body, [class*="css"] { font-family: 'Inter', system-ui, sans-serif !important; }

        /* Streamlit deja ~6rem de aire arriba de CADA página por defecto —
        pensado para dejar lugar a un título grande, pero de sobra para el
        resto de las pantallas (y clave en el login: sumado al margin-top
        de .sb-login-logo de más abajo, tapaba el botón "Ingresar" y obligaba
        a scrollear en una pantalla que solo tiene 2 campos). Se achica una
        sola vez acá para toda la app. */
        .block-container { padding-top: 2.5rem !important; }

        /* ---- Chrome nativo de Streamlit: fondo de la app, sidebar, header,
        inputs, tabs, expanders, etc. — todo referenciando los mismos
        tokens de arriba, así responde solo al cambiar de modo. Sin esto el
        modo oscuro solo pintaría las cards/wordmark propios de la app y
        dejaría el resto (fondo, sidebar, inputs) en blanco. ---- */
        .stApp { background: var(--sb-bg); }
        [data-testid="stHeader"] { background: var(--sb-bg) !important; }
        [data-testid="stSidebar"] {
            background: var(--sb-zone) !important;
            border-right: 1px solid var(--sb-border);
        }
        [data-testid="stAppViewContainer"], [data-testid="stMain"],
        [data-testid="stMarkdownContainer"] p, [data-testid="stMarkdownContainer"] li,
        [data-testid="stMarkdownContainer"] span, [data-testid="stCaptionContainer"],
        [data-testid="stWidgetLabel"] p, label, .stApp, .stApp h1, .stApp h4, .stApp h5, .stApp h6 {
            color: var(--sb-navy);
        }
        [data-testid="stCaptionContainer"] { color: var(--sb-text-secondary) !important; }
        hr, [data-testid="stDivider"] { border-color: var(--sb-border) !important; }

        /* Inputs (texto, número, textarea) y el combobox del selectbox
        (esta versión de Streamlit ya no usa baseweb para el selectbox,
        sino un react-aria-ComboBox con un <input> normal adentro — se
        apunta directo al tag en vez de un atributo data-baseweb que ya no
        existe): mismo tratamiento surface + borde + texto que el resto de
        la app. */
        .stTextInput input, .stNumberInput input, .stTextArea textarea,
        [data-testid="stSelectbox"] input {
            background: var(--sb-surface) !important; color: var(--sb-navy) !important;
            border-color: var(--sb-border) !important;
        }
        [data-testid="stSelectbox"] [role="group"] {
            background: var(--sb-surface) !important; border-color: var(--sb-border) !important;
        }
        /* Fecha (react-aria-DateField): tampoco es un <input> — son 3
        <span role="spinbutton"> (día/mes/año) más los separadores "/" (o el
        "–" entre el inicio y el fin, en un date_input de rango) como spans
        [data-type="literal"] aparte — sin incluirlos acá quedaban con el
        color por defecto de Streamlit, invisible sobre fondo oscuro aunque
        los números sí se vieran bien. */
        [data-testid="stDateInputField"] {
            background: var(--sb-surface) !important; border-color: var(--sb-border) !important;
        }
        [data-testid="stDateInputField"] [role="spinbutton"],
        [data-testid="stDateInputField"] [data-type="literal"] {
            color: var(--sb-navy) !important;
        }
        /* Steppers +/- de los number_input: el SVG usa fill="currentColor"
        pero Streamlit les fija su propio "color" fijo (navy claro),
        independiente del tema — sin esto quedaban prácticamente invisibles
        (ícono oscuro sobre fondo oscuro) en modo oscuro. */
        [data-testid="stNumberInputStepDown"], [data-testid="stNumberInputStepUp"] {
            color: var(--sb-navy) !important;
        }
        /* El menú desplegable del selectbox se monta aparte, fuera del
        árbol del widget (portal a nivel body) — necesita su propio
        selector por rol ARIA (no depende de baseweb), si no queda con los
        colores por defecto de Streamlit. */
        [role="listbox"] {
            background: var(--sb-surface) !important; color: var(--sb-navy) !important;
            border: 1px solid var(--sb-border) !important;
        }
        [role="listbox"] [role="option"] { color: var(--sb-navy) !important; }

        /* Expanders (las 10 secciones del cotizador) y tabs: cabecera con
        superficie propia en vez del gris por defecto de Streamlit, que no
        se adapta solo al modo oscuro. */
        [data-testid="stExpander"] {
            background: var(--sb-surface) !important; border-color: var(--sb-border) !important;
        }
        [data-testid="stExpander"] summary { color: var(--sb-navy) !important; }
        .stTabs [data-baseweb="tab-list"] { border-bottom-color: var(--sb-border) !important; }
        .stTabs [data-baseweb="tab"] p { color: var(--sb-text-secondary) !important; }
        .stTabs [aria-selected="true"] p { color: var(--sb-orange-dark) !important; }

        /* Cargador de archivos: dropzone con la misma superficie. */
        [data-testid="stFileUploaderDropzone"] {
            background: var(--sb-zone) !important; border-color: var(--sb-border) !important;
        }
        /* Streamlit no tiene forma de localizar el texto nativo del
        uploader ("Upload" / "200MB per file") desde Python — queda fijo en
        inglés siempre, en las 5+ secciones de Documentación de la app.
        Se ocultan esos 2 textos (font-size: 0) y se reemplazan con ::after
        en español, mismo tamaño/línea que el original (14px/22.4px,
        verificado con el DOM real del uploader) para que no se note el
        cambio de tipografía. Selectores acotados dentro de
        stFileUploaderDropzone para no tocar otros botones "secondary" de
        la app. El texto de "200MB por archivo" queda genérico (sin listar
        los tipos de archivo aceptados) porque esa parte del texto original
        varía según el uploader — más seguro que armar la lista a mano acá
        y que se desincronice si el accept= cambia en Python. */
        [data-testid="stFileUploaderDropzone"] [data-testid="stBaseButton-secondary"] [data-testid="stMarkdownContainer"] p {
            font-size: 0;
        }
        [data-testid="stFileUploaderDropzone"] [data-testid="stBaseButton-secondary"] [data-testid="stMarkdownContainer"] p::after {
            content: "Subir archivo"; font-size: 14px; line-height: 22.4px;
        }
        [data-testid="stFileUploaderDropzoneInstructions"] span {
            font-size: 0;
        }
        [data-testid="stFileUploaderDropzoneInstructions"] span::after {
            content: "Máximo 200MB por archivo"; font-size: 14px; line-height: 22.4px;
        }

        /* Menú de navegación del sidebar (Panel de Control/Clientes/
        Cotizador/CRM): antes era un st.radio sin estilo propio, con los
        círculos nativos del navegador uno debajo del otro — se lo lleva al
        lenguaje visual de menú real: ítems de ancho completo, el activo
        resaltado con fondo propio + el mismo borde izquierdo naranja que
        usan el resto de las "cards hoja" de la app (cardwrap_/clientecard_/
        formrow_/cotfila_), sin el bullet del radio (ya no hace falta,
        el resaltado cumple esa función). El label "Navegación" se oculta
        (arriba del propio menú, con el wordmark SKYBRIDGE ya puesto, era
        redundante) pero queda en el DOM: stRadioGroup ya trae su propio
        aria-label="Navegación" para lectores de pantalla, así que ocultarlo
        visualmente no rompe accesibilidad. Verificado contra el DOM real
        del widget (data-testid, no clases de emotion, que cambian entre
        builds) para que sobreviva updates de Streamlit. */
        [data-testid="stSidebar"] [data-testid="stRadio"] > label {
            display: none;
        }
        [data-testid="stSidebar"] [data-testid="stRadioGroup"] {
            gap: 2px;
        }
        [data-testid="stSidebar"] [data-testid="stRadioOption"] {
            width: 100%; padding: 10px 12px; border-radius: 6px;
            border-left: 3px solid transparent;
            transition: background 0.15s ease, border-color 0.15s ease;
        }
        [data-testid="stSidebar"] [data-testid="stRadioOption"]:hover {
            background: var(--sb-surface);
        }
        [data-testid="stSidebar"] [data-testid="stRadioOption"][data-selected="true"] {
            background: var(--sb-surface); border-left-color: var(--sb-orange);
        }
        [data-testid="stSidebar"] [data-testid="stRadioOption"][data-selected="true"] [data-testid="stMarkdownContainer"] p {
            color: var(--sb-orange-dark) !important; font-weight: 700;
        }
        [data-testid="stSidebar"] [data-testid="stRadioOption"] [data-testid="stMarkdownContainer"] p {
            font-size: 14px; margin: 0;
        }
        [data-testid="stSidebar"] [data-testid="stRadioOption"] > div > div > div:has(+ [data-testid="stMarkdownContainer"]) {
            display: none;
        }

        /* NOTA: st.dataframe (Glide Data Grid) se renderiza en <canvas> y
        toma sus colores del theme activo de Streamlit del lado del
        servidor, no de CSS del navegador — probado con --gdg-bg-cell/etc.
        acá: la variable SÍ queda seteada en el DOM pero el canvas no la
        lee, así que este toggle no lo puede re-pintar. Sigue viéndose con
        la paleta clara en modo oscuro; para eso sí queda 100% acorde hace
        falta cambiar el theme nativo desde ☰ > Settings > Choose app
        theme (usa la paleta de [theme.dark] en config.toml). */

        /* Wordmark de marca en el sidebar (calco de "SKY" + "BRIDGE" del sitio) */
        .sb-logo {
            font-size: 21px; font-weight: 800; letter-spacing: 0.18em;
            text-transform: uppercase; color: var(--sb-navy); line-height: 1.3;
        }
        .sb-logo span { color: var(--sb-orange); }
        .sb-tagline {
            font-size: 11px; font-weight: 500; letter-spacing: 0.08em;
            text-transform: uppercase; color: var(--sb-text-secondary); margin-bottom: 1.25rem;
        }

        /* Títulos de sección: acento naranja con más contraste (orange-dark)
        sobre fondo claro, y una línea sutil que ordena la jerarquía visual. */
        .stApp h2, .stApp h3 {
            color: var(--sb-orange-dark) !important; font-weight: 800 !important;
            letter-spacing: 0.01em; border-bottom: 1px solid var(--sb-border);
            padding-bottom: 0.35rem;
        }

        /* Botones: mayúsculas + tracking, como los CTA del sitio.
        Selector DESCENDIENTE (" button", no "> button" de hijo directo):
        un st.button con help="..." hace que Streamlit meta el <button> real
        adentro de un wrapper .stTooltipHoverTarget/.stTooltipIcon extra —
        con ">" ese botón dejaba de matchear (ya no es hijo DIRECTO de
        .stButton) y se quedaba sin mayúscula ni tracking, aunque el botón
        de al lado (sin help=) sí los tuviera — se notaba en el CRM,
        "Cotizar"/"Descartado" (con help=) al lado de "Ver ficha →" (sin
        help=) en mayúscula. Afecta a las 8 reglas de .stButton/
        .stFormSubmitButton/.stDownloadButton de esta hoja de estilos. */
        .stButton button, .stFormSubmitButton button, .stDownloadButton button {
            text-transform: uppercase; letter-spacing: 0.08em;
            font-weight: 700; font-size: 12.5px; border-radius: 4px !important;
        }
        /* Botones "secondary" (todos menos los type="primary" naranjas, que
        ya toman bien el primaryColor del theme): sin esto quedan con el
        blanco/gris por defecto de Streamlit — invisibles sobre fondo
        oscuro. Los botones "unset" de las cards (cardwrap_/volver_ficha)
        pisan esto después por selector más específico.
        Los de dentro de un st.form (form_submit_button) tienen un testid
        distinto ("...secondaryFormSubmit", no "...secondary" a secas) —
        sin incluirlo acá quedaban con fondo blanco fijo y letra casi
        invisible en modo oscuro (ej. "Guardar cambios"/"Eliminar cliente"
        en la ficha de cliente). El trigger de st.popover ("Filtros
        avanzados") tiene su propio testid aparte ("stPopoverButton") con el
        mismo problema — fondo blanco fijo + letra clara de modo oscuro,
        prácticamente invisible. */
        [data-testid="stBaseButton-secondary"], [data-testid="stBaseButton-secondaryFormSubmit"],
        [data-testid="stPopoverButton"] {
            background: var(--sb-surface) !important; color: var(--sb-navy) !important;
            border: 1px solid var(--sb-border) !important;
        }

        /* "Botón" de link armado a mano (ancla <a>, no <button>): hace falta
        para poder fijar target="whatsapp_web" en el acceso rápido de
        WhatsApp (st.link_button no permite elegir el target, y por eso
        siempre abre pestaña nueva por cada contacto) — se estiliza igual
        que un botón secundario para que no se note la diferencia. */
        .sb-btn-link {
            display: flex; align-items: center; justify-content: center;
            width: 100%; box-sizing: border-box; padding: 0.5rem 1rem;
            border-radius: 4px; border: 1px solid var(--sb-border);
            background: var(--sb-surface); color: var(--sb-navy) !important;
            font-weight: 700; font-size: 12.5px; text-transform: uppercase;
            letter-spacing: 0.08em; text-decoration: none !important;
            margin-bottom: 0.5rem; cursor: pointer;
        }
        .sb-btn-link:hover { border-color: var(--sb-orange); color: var(--sb-orange) !important; }

        /* "Ver detalle →" de una importación (mismo <a> de arriba, clase
        extra): outline naranja como vercliente_/crmojo_, en vez del look
        de botón secundario neutro por defecto de sb-btn-link. */
        .sb-verdetalle-imp {
            height: 42px; border-radius: 6px !important;
            background: transparent !important; border: 1.5px solid var(--sb-orange) !important;
            color: var(--sb-orange-dark) !important;
        }
        .sb-verdetalle-imp:hover {
            background: var(--sb-orange) !important; border-color: var(--sb-orange) !important;
            color: #fff !important;
        }

        /* Fila de carga rápida de contactos (CRM): el botón "＋ Agregar" no
        tiene el renglón de label que sí reservan los text_input de al lado
        (aunque esté colapsado), así que sin este ajuste queda más alto y
        partido en 2 líneas si la columna es angosta — con nowrap se
        mantiene en una sola línea y con align-self se nivela con la altura
        real de los inputs vecinos en vez de estirarse. */
        [class*="st-key-formrow_crm_carga_rapida"] .stFormSubmitButton button {
            white-space: nowrap;
        }
        [class*="st-key-formrow_crm_carga_rapida"] .stFormSubmitButton {
            align-self: stretch;
        }
        /* Fila de contacto en la lista del CRM: mismo motivo que arriba —
        que "VER FICHA →" no se parta en 2 líneas si la columna queda algo
        angosta en pantallas más chicas. */
        [class*="st-key-contactcard_"] .stButton button {
            white-space: nowrap;
        }

        /* Métricas: etiqueta gris en mayúscula, valor en navy y negrita */
        [data-testid="stMetricLabel"] {
            text-transform: uppercase; letter-spacing: 0.05em;
            font-size: 11px !important; color: var(--sb-text-secondary) !important;
        }
        /* El tamaño gigante por defecto de st.metric trunca con "…" los
        valores en USD largos (ej. "USD 56,513…") apenas la columna se
        angosta — Se achica la letra y se permite el wrap en vez de cortar el
        texto, así el número completo siempre queda legible. */
        [data-testid="stMetricValue"] {
            font-weight: 700; color: var(--sb-navy);
            font-size: 1.35rem !important; line-height: 1.25 !important;
            white-space: normal !important; overflow: visible !important;
            text-overflow: clip !important; overflow-wrap: break-word;
        }

        /* Texto de las cards del Panel de Control y del listado de clientes:
        todas las líneas van en un único bloque para que no se acumule el
        margen que Streamlit agrega entre cada markdown/caption suelto —
        clave para que las cards no crezcan al sumar más registros. */
        .sb-card-title { font-weight: 700; font-size: 14px; margin: 0 0 2px 0; color: var(--sb-navy); }
        .sb-card-line {
            font-size: 12px; color: var(--sb-text-secondary); line-height: 1.45; margin: 0;
            white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
        }

        /* Header ejecutivo de la ficha de cliente: banner tipo tarjeta con
        flexbox (info a la izquierda, acciones a la derecha), construido con
        HTML/CSS propio en vez de los bloques por defecto de Streamlit. Un
        acento naranja a la izquierda lo empareja con el resto de las
        "cards hoja" de la app (clientecard_/cardwrap_/formrow_), y un
        tamaño más generoso le da el peso visual de encabezado de página en
        vez de una tira angosta. */
        [class*="st-key-fichahdr_"] {
            background: var(--sb-surface); border-radius: 8px; padding: 22px 28px;
            border-left: 4px solid var(--sb-orange);
            box-shadow: 0 1px 3px rgba(15, 23, 42, 0.08), 0 1px 2px rgba(15, 23, 42, 0.04);
            margin-bottom: -0.25rem;
        }
        [class*="st-key-fichahdr_"] > div[data-testid="stHorizontalBlock"] {
            display: flex; justify-content: space-between; align-items: center;
        }
        .sb-fichahdr-nombre { font-size: 26px; font-weight: 800; color: var(--sb-navy); margin: 0; line-height: 1.3; }
        .sb-fichahdr-datos { font-size: 14px; color: var(--sb-text-secondary); margin: 6px 0 0 0; }
        /* El email en la línea de datos se autoenlaza (mailto:) con el
        markdown de Streamlit; sin esto queda en azul default, fuera de la
        paleta de marca. */
        [class*="st-key-fichahdr_"] a {
            color: inherit; text-decoration: none;
        }
        /* Botones de acción rápida del header: compactos pero con presencia
        real (no el CTA de altura fija por defecto de Streamlit, tampoco la
        versión demasiado chica de 34px probada antes). */
        [class*="st-key-fichahdr_"] .stButton button {
            height: 40px; min-height: 40px; padding: 0 18px; font-size: 12px;
        }

        /* Botones "Volver a..." (ficha de cliente, listado del cotizador,
        panel, contacto de CRM): mismo botón outline naranja que "Ver ficha
        →"/"Ver detalle →" en vez del link de texto plano que tenían antes,
        para que se lean como el resto de los botones de la app. "sticky"
        en el wrapper (no en el <button>) para que se queden pegados arriba
        de la pantalla al bajar por una ficha larga, en vez de perderse
        arriba del todo — la superficie propia evita que el contenido se
        transparente por debajo al quedar "pegado". */
        /* top: 72px, no 12px — el header fijo de Streamlit (Deploy/menú)
        mide 60px y flota con z-index propio por encima de todo el
        contenido; con un offset menor el botón quedaba "pegado" pero
        tapado debajo de esa barra, sin verse. */
        [class*="st-key-volver_"] {
            position: sticky !important; top: 72px !important; z-index: 20;
            width: fit-content !important; margin-bottom: 12px;
        }
        [class*="st-key-volver_"] .stButton button {
            background: var(--sb-surface) !important;
            border: 1.5px solid var(--sb-orange) !important;
            color: var(--sb-orange-dark) !important;
            border-radius: 6px !important;
            height: 38px !important; min-height: 38px !important;
            padding: 0 16px !important; font-size: 13px !important; font-weight: 600 !important;
            white-space: nowrap !important;
            box-shadow: 0 1px 4px rgba(0,0,0,0.08);
            transition: background 0.15s ease, color 0.15s ease, box-shadow 0.15s ease;
        }
        [class*="st-key-volver_"] .stButton button:hover {
            background: var(--sb-orange) !important; color: #fff !important;
            box-shadow: 0 2px 8px rgba(232, 101, 42, 0.35) !important;
        }
        [class*="st-key-volver_"] .stButton button p { color: inherit !important; }

        /* Cards "hoja" (filas de clientes, cards del Panel de Control, filas
        de formulario del cotizador): superficie blanca sobre el canvas gris
        de la página, con sombra suave y un acento naranja a la izquierda —
        el anclaje visual que faltaba cuando todo era blanco sobre blanco.
        st.container(border=True) y st.columns() comparten el mismo testid en
        Streamlit, así que cada card se identifica por su propio
        `key="cardwrap_..."` / `key="formrow_..."` / `key="clientecard_..."` /
        `key="resultgroup_..."` (Streamlit lo vuelca como clase `st-key-<key>`)
        en vez de por ese testid genérico. Se usan prefijos distintos porque
        llevan un tratamiento de botón distinto: "cardwrap_" son cards de
        listado con acción de link discreto ("Ver más"); "formrow_" y
        "clientecard_" son filas con botones reales (formulario o
        "＋ Cotizar"/"Ver ficha →"); "resultgroup_" son grupos de métricas de
        solo lectura (Resultado de la operación / Simulación de venta), sin
        botones. */
        [class*="st-key-cardwrap_"], [class*="st-key-formrow_"], [class*="st-key-clientecard_"],
        [class*="st-key-resultgroup_"], [class*="st-key-cotfila_"], [class*="st-key-filacrm_"],
        [class*="st-key-impfila_"] {
            background: var(--sb-surface) !important;
            border: 1px solid var(--sb-border) !important;
            border-left: 3px solid var(--sb-orange) !important;
            border-radius: 8px !important;
            box-shadow: 0 1px 3px rgba(15, 23, 42, 0.08), 0 1px 2px rgba(15, 23, 42, 0.04) !important;
        }

        /* Card de cliente: más aire que el padding por defecto de Streamlit
        (15px parejo) — el bloque de 3 líneas de la columna de métricas
        quedaba con la última línea pegada al borde inferior de la card.
        Más padding vertical + un piso de alto le da al texto margen real
        sin agrandar el contenido en sí. */
        [class*="st-key-clientecard_"] {
            padding: 20px 24px 22px 24px !important;
            min-height: 96px;
        }
        /* Botones de la card de cliente: "＋ Cotizar" es la acción primaria
        (naranja relleno, type="primary" nativo de Streamlit — ya toma el
        primaryColor del theme en claro/oscuro sin CSS propio) y "Ver ficha"
        pasa de gris plano a un outline naranja con hover relleno, para que
        haya jerarquía real entre ambos en vez de dos botones grises iguales. */
        [class*="st-key-clientecard_"] .stButton button {
            height: 42px !important;
            border-radius: 6px !important;
            transition: background 0.15s ease, color 0.15s ease, box-shadow 0.15s ease;
        }
        [class*="st-key-cotizarcliente_"] button:hover {
            box-shadow: 0 2px 8px rgba(232, 101, 42, 0.35) !important;
        }
        [class*="st-key-vercliente_"] button {
            background: transparent !important;
            border: 1.5px solid var(--sb-orange) !important;
            color: var(--sb-orange-dark) !important;
        }
        [class*="st-key-vercliente_"] button:hover {
            background: var(--sb-orange) !important;
            border-color: var(--sb-orange) !important;
            color: #fff !important;
        }
        [class*="st-key-vercliente_"] button p { color: inherit !important; }

        /* Fila de cotización en el historial del Cotizador: misma card
        "hoja" que Clientes (ya sumada arriba a la lista compartida), con su
        propio padding — un poco menos alta que clientecard_ porque acá el
        bloque más largo es una sola línea de productos, no 3. */
        [class*="st-key-cotfila_"] {
            padding: 16px 22px 18px 22px !important;
            min-height: 84px;
        }
        /* Acciones de la fila (editar/PDF/detalle/duplicar/eliminar): 5
        íconos cuadrados alineados a la derecha en vez de un selector +
        botones de texto aparte debajo del listado — flex en el contenedor
        para que cada botón/download_button ocupe solo su propio ancho. */
        [class*="st-key-cotfilaacciones_"] {
            display: flex !important; flex-direction: row !important;
            justify-content: flex-end; align-items: center; gap: 6px;
        }
        [class*="st-key-cotfilaacciones_"] .stButton button,
        [class*="st-key-cotfilaacciones_"] .stDownloadButton button {
            width: 38px !important; min-width: 38px !important; height: 38px !important;
            min-height: 38px !important; padding: 0 !important; border-radius: 6px !important;
        }
        /* El wrapper de tooltip que agrega Streamlit cuando el botón lleva
        help= (stTooltipHoverTarget/stTooltipIcon) no debe angostar el
        ícono — sin esto quedaba con un padding/ancho propio que desalineaba
        los 5 botones entre sí. */
        [class*="st-key-cotfilaacciones_"] .stTooltipIcon,
        [class*="st-key-cotfilaacciones_"] .stTooltipHoverTarget {
            width: 38px !important; height: 38px !important;
        }
        /* "Eliminar" en rojo al hover, para que se note que es la acción
        destructiva del grupo — el resto se queda con el gris neutro. */
        [class*="st-key-coteli_"] button:hover {
            background: #FEE2E2 !important; border-color: #DC2626 !important; color: #DC2626 !important;
        }

        /* Columna de estado de la fila de cotización: el badge arriba y,
        debajo, el trigger del popover "Cambiar" — chico y sin wrapear,
        para que no compita en altura con los 5 íconos de acción al lado. */
        [class*="st-key-cotestado_"] {
            display: flex !important; flex-direction: column !important;
            align-items: flex-start !important; gap: 6px;
        }
        [class*="st-key-cotestado_"] .stPopover button {
            height: 22px !important; min-height: 22px !important; padding: 0 7px !important;
            font-size: 11px !important; white-space: nowrap !important; gap: 3px !important;
        }
        [class*="st-key-cotestado_"] .stPopover button svg {
            width: 12px !important; height: 12px !important;
        }

        /* Fila de importación en la ficha de cliente: mismo padding que
        cotfila_ (una sola línea de detalle además del número/proveedor).
        Antes esto era un st.expander con todo el formulario adentro; ahora
        es una card con un único botón "Ver detalle →" (outline, mismo
        tratamiento que vercliente_/crmojo_) que abre esa pantalla completa. */
        [class*="st-key-impfila_"] {
            padding: 16px 22px 18px 22px !important;
            min-height: 84px;
        }
        [class*="st-key-impfila_"] .stButton button {
            height: 42px !important; border-radius: 6px !important;
            background: transparent !important;
            border: 1.5px solid var(--sb-orange) !important;
            color: var(--sb-orange-dark) !important;
            transition: background 0.15s ease, color 0.15s ease;
        }
        [class*="st-key-impfila_"] .stButton button:hover {
            background: var(--sb-orange) !important;
            border-color: var(--sb-orange) !important;
            color: #fff !important;
        }
        [class*="st-key-impfila_"] .stButton button p { color: inherit !important; }

        /* Botón "Nueva cotización": mismo naranja primario, pero más chico
        y con el radio/sombra al hover del resto de los CTA de la app (ej.
        "＋ Cotizar" en Clientes), en vez del bloque plano por defecto de
        Streamlit. */
        [class*="st-key-nueva_cotizacion_historial"] button {
            height: 38px !important; border-radius: 6px !important;
            transition: box-shadow 0.15s ease;
        }
        [class*="st-key-nueva_cotizacion_historial"] button:hover {
            box-shadow: 0 2px 8px rgba(232, 101, 42, 0.35) !important;
        }

        /* Popover "Filtros avanzados": igual que el listbox del selectbox
        (ver [role="listbox"] arriba), Streamlit monta el panel del popover
        aparte, como hijo directo de <body> — NO como descendiente de la
        card que lo abre — así que un selector scopeado por key (ej.
        "[class*=st-key-formrow_cotizador_filtros] [data-testid=...]") nunca
        matchea nada ahí adentro. Necesita sus propios selectores globales,
        igual que el listbox. Sin el fondo/borde acá, quedaba con la
        superficie clara fija de Streamlit — un cuadro blanco encima de la
        página oscura en modo oscuro, con las labels (ya coloreadas para
        texto claro) invisibles sobre ese blanco. */
        [data-testid="stPopoverBody"] {
            background: var(--sb-surface) !important; border-color: var(--sb-border) !important;
            min-width: 320px;
        }
        [data-testid="stPopoverBody"] [data-testid="stDateInputField"] {
            border-radius: 6px !important; width: 100% !important;
        }

        /* Nombre de cliente con contacto vinculado en CRM: link discreto
        ("all:unset" salvo tipografía/color), en vez del botón gris cuadrado
        por defecto — así se distingue de un botón de acción sin dejar de
        leerse como el nombre en negrita que ya era antes. Selector por
        descendencia (no ">"): al llevar help=, Streamlit mete el <button>
        unas capas más adentro (stTooltipHoverTarget/stTooltipIcon) que un
        botón sin tooltip. */
        [class*="st-key-cotcliente_"] button {
            all: unset !important; cursor: pointer; font-size: 13px; font-weight: 600;
            color: var(--sb-navy);
        }
        [class*="st-key-cotcliente_"] button:hover {
            color: var(--sb-orange-dark) !important; text-decoration: underline;
        }
        [class*="st-key-cotcliente_"] button p { color: inherit !important; }

        /* Card de contacto del CRM: mismo padding/alto que clientecard_ (ya
        sumada arriba a la lista compartida para fondo/borde/sombra), pero
        con el borde izquierdo coloreado POR ETAPA en vez de naranja fijo —
        un vistazo a la barra ya dice en qué etapa está, antes solo lo decía
        la badge chica del medio. Colores generados desde crm.COLOR_ETAPA
        (__CRM_ETAPA_CSS__ más abajo), nunca hardcodeados acá, para que no
        se puedan desincronizar. */
        [class*="st-key-filacrm_"] {
            padding: 16px 22px 18px 22px !important;
            min-height: 84px;
        }
__CRM_ETAPA_CSS__
        /* Botones de la card: "Avanzar etapa" (sólido, primario nativo de
        Streamlit) + "Ver ficha →" (outline) — mismo par que "＋ Cotizar" /
        "Ver ficha →" en Clientes. Selector por descendencia (no ">"): al
        llevar help=, el wrapper de tooltip mete el <button> más adentro.
        white-space: nowrap — "Avanzar etapa" partía en 2 líneas dentro de
        columnas angostas, lo que hacía crecer el botón en alto y perder la
        alineación vertical con el resto de la fila (nombre/fecha/badge);
        junto con las columnas más anchas de c5 (ver más arriba), el texto
        entra cómodo en una sola línea. */
        [class*="st-key-filacrm_"] .stButton button {
            height: 42px !important;
            border-radius: 6px !important;
            white-space: nowrap !important;
            padding-left: 10px !important; padding-right: 10px !important;
            transition: background 0.15s ease, color 0.15s ease, box-shadow 0.15s ease;
        }
        /* white-space:nowrap por sí solo no alcanza: el <p> que Streamlit
        mete adentro del botón (vía stMarkdownContainer) trae su propio
        overflow-wrap:break-word de fábrica, que corta la palabra a la
        mitad en vez de respetar el nowrap del botón cuando el texto no
        entra en una columna angosta (pasó con "Marcar en negociación",
        el label que tenía antes "Cotizar") — mejor que desborde a que
        rompa una palabra por la mitad. */
        [class*="st-key-filacrm_"] .stButton button p {
            white-space: nowrap !important; overflow-wrap: normal !important; word-break: normal !important;
        }
        [class*="st-key-crmcotizar_"] button:hover {
            box-shadow: 0 2px 8px rgba(232, 101, 42, 0.35) !important;
        }
        [class*="st-key-crmojo_"] button {
            background: transparent !important;
            border: 1.5px solid var(--sb-orange) !important;
            color: var(--sb-orange-dark) !important;
        }
        [class*="st-key-crmojo_"] button:hover {
            background: var(--sb-orange) !important;
            border-color: var(--sb-orange) !important;
            color: #fff !important;
        }
        [class*="st-key-crmojo_"] button p { color: inherit !important; }
        /* "Eliminar contacto" en rojo al hover, mismo criterio que
        "Eliminar" en el historial del Cotizador (coteli_). */
        [class*="st-key-crmdel_"] button:hover {
            background: #FEE2E2 !important; border-color: #DC2626 !important; color: #DC2626 !important;
        }
        /* "Marcar como Descartado": gris apagado en reposo (no compite en
        peso visual con "Avanzar etapa"/"Ver ficha →" arriba), mismo rojo
        de "Eliminar" al hover — es una acción de salida del embudo, se
        distingue de las de avance sin gritar todo el tiempo. */
        [class*="st-key-crmdescartar_"] button {
            background: transparent !important;
            border: 1.5px solid var(--sb-border) !important;
            color: var(--sb-text-secondary) !important;
            height: 34px !important;
        }
        [class*="st-key-crmdescartar_"] button:hover {
            background: #FEE2E2 !important; border-color: #DC2626 !important; color: #DC2626 !important;
        }

        /* Íconos compactos de acceso rápido (WhatsApp / mail / ficha de
        cliente vinculado): mismo tratamiento que cotfilaacciones_ en el
        historial del Cotizador — cuadrados de 34px en fila, alineados a la
        izquierda (a diferencia de cotfilaacciones_, acá comparten columna
        con nada más a la derecha). */
        [class*="st-key-filacrmicons_"] {
            display: flex !important; flex-direction: row !important;
            align-items: center; gap: 6px;
        }
        [class*="st-key-filacrmicons_"] .stButton button,
        [class*="st-key-filacrmicons_"] .stLinkButton a,
        [class*="st-key-filacrmicons_"] .sb-btn-link {
            width: 34px !important; min-width: 34px !important; height: 34px !important;
            min-height: 34px !important; padding: 0 !important; border-radius: 6px !important;
        }
        /* El WhatsApp de esta fila es un .sb-btn-link (target="whatsapp_web"
        fijo, ver _link_button_html) en vez de un st.link_button nativo —
        mismo motivo que el de _render_acciones_rapidas: st.link_button
        fuerza target="_blank" y no deja elegirlo. Sin esto heredaba el
        texto en mayúscula/letter-spacing pensado para un botón de ancho
        completo, no para un cuadrado de ícono. */
        [class*="st-key-filacrmicons_"] .sb-btn-link {
            margin-bottom: 0 !important; text-transform: none; letter-spacing: normal; font-size: 15px;
        }
        [class*="st-key-filacrmicons_"] .stTooltipIcon,
        [class*="st-key-filacrmicons_"] .stTooltipHoverTarget {
            width: 34px !important; height: 34px !important;
        }

        /* Columnas del Kanban (Panel de Control, key="zona_..."): agrupan
        varias cards, así que se tratan como una "zona" (fondo gris tenue,
        sin sombra ni acento) en vez de una card más — evita que una columna
        con pocos ítems se vea como un agujero vacío en el layout. */
        [class*="st-key-zona_"] {
            background: var(--sb-zone) !important;
            border: 1px solid var(--sb-border) !important;
            border-left: 1px solid var(--sb-border) !important;
            box-shadow: none !important;
            border-radius: 12px !important;
        }

        /* Botón de acción dentro de una card del Panel de Control (Abrir /
        Ver detalle): el botón por defecto de Streamlit tiene una altura fija
        pensada para ser un CTA principal, y sobre 2-3 líneas de texto chico
        queda desproporcionado. Acá se lo trata como una acción secundaria
        liviana — chico, sin ocupar todo el ancho de la card. */
        [class*="st-key-cardwrap_"] .stButton {
            display: flex; justify-content: flex-start; margin: 4px 0 0 0;
        }
        [class*="st-key-cardwrap_"] .stButton button {
            all: unset !important; cursor: pointer; box-sizing: border-box;
            padding: 1px 2px; margin: 0; font-size: 12px; line-height: 1.45;
            font-weight: 600; font-family: inherit; letter-spacing: normal;
            text-transform: none; border-radius: 4px; text-align: center;
            color: var(--sb-orange-dark); white-space: nowrap; display: block;
        }
        [class*="st-key-cardwrap_"] .stButton button p {
            font-size: 12px; line-height: 1.45; white-space: nowrap; margin: 0;
        }
        [class*="st-key-cardwrap_"] .stButton button:hover {
            text-decoration: underline;
        }

        .sb-login-logo {
            font-size: 26px; font-weight: 800; letter-spacing: 0.18em;
            text-transform: uppercase; color: var(--sb-navy); text-align: center;
            /* Antes 8vh: sumado al padding-top del .block-container de arriba
            y al padding de la card de más abajo, el login (solo 2 campos y un
            botón) no entraba en una pantalla estándar sin scrollear. Con el
            .block-container ya achicado, esto alcanza para centrar el login
            sin quedar pegado arriba. */
            margin-top: 4vh; margin-bottom: 2px;
        }
        .sb-login-logo span { color: var(--sb-orange); }
        .sb-login-tagline {
            font-size: 11.5px; font-weight: 500; letter-spacing: 0.1em;
            text-transform: uppercase; color: var(--sb-text-secondary);
            text-align: center; margin-bottom: 1.25rem;
        }
        [class*="st-key-sb_login_card"] {
            background: var(--sb-surface); border: 1px solid var(--sb-border);
            border-radius: 10px; padding: 1.5rem 2rem 0.75rem;
            box-shadow: 0 4px 24px rgba(15, 23, 42, 0.10), 0 1px 3px rgba(15, 23, 42, 0.06);
            border-top: 4px solid var(--sb-orange);
        }
        </style>
        """.replace("__SB_VARS__", variables_css).replace("__CRM_ETAPA_CSS__", etapa_css),
        unsafe_allow_html=True,
    )


_inyectar_estilos()

MONEDAS = ["USD", "ARS"]
PRORRATEOS = ["FOB", "PESO"]
SI_NO = ["NO", "SI"]
ESTADOS = ["Borrador", "Enviada", "Aprobada", "Rechazada"]
COLOR_ESTADO_COTIZACION = {
    "Borrador": "#64748B", "Enviada": "#2563EB", "Aprobada": "#16A34A", "Rechazada": "#DC2626",
}

# Columnas porcentuales: se editan y muestran en puntos porcentuales (p.ej. 2.5
# significa 2.5%) para que cargar datos sea intuitivo, pero el motor de
# cálculo y la base de datos siguen trabajando en fracción (0.025). La
# conversión x100 / /100 se hace únicamente en esta capa de UI.
PCT_PROD_COLS = [
    "pct_derechos", "pct_tasa_estadistica", "pct_antidumping", "pct_iva",
    "pct_iva_adicional", "pct_ganancias", "pct_iibb", "margen_pct",
]
PCT_GASTO_COLS = ["pct_iva"]


def _fila_a_puntos(row, columnas):
    """Convierte los campos porcentuales de un dict de fracción (0.21) a puntos porcentuales (21.0) para editar."""
    row = dict(row)
    for col in columnas:
        row[col] = _f_local(row.get(col)) * 100
    return row


def _a_fraccion(records, columnas):
    """Convierte de puntos porcentuales (21.0) a fracción (0.21) para calcular/guardar."""
    out = []
    for r in records:
        r = dict(r)
        for col in columnas:
            r[col] = _f_local(r.get(col)) / 100
        out.append(r)
    return out


def _f_local(v):
    try:
        if v is None or v == "":
            return 0.0
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _parse_fecha(valor):
    """Convierte una fecha guardada como texto ISO (YYYY-MM-DD) al objeto
    date que espera st.date_input; None si está vacía o no se puede leer."""
    if not valor:
        return None
    try:
        return datetime.strptime(valor, "%Y-%m-%d").date()
    except (TypeError, ValueError):
        return None


def _fmt_fecha(valor):
    d = _parse_fecha(valor)
    return d.strftime("%d/%m/%Y") if d else None


def _link_button_html(label, href, target="_self", title=None):
    """'Botón' de link armado a mano: a diferencia de st.link_button, permite
    fijar el atributo target — necesario para que WhatsApp reutilice siempre
    la misma pestaña (ver .sb-btn-link en _inyectar_estilos). title= arma un
    tooltip nativo del navegador (el wrapper de tooltip de Streamlit,
    help=, no existe para HTML crudo)."""
    title_attr = f' title="{html.escape(title, quote=True)}"' if title else ""
    return (
        f'<a href="{html.escape(href, quote=True)}" target="{target}" rel="noopener"{title_attr} '
        f'class="sb-btn-link">{html.escape(label)}</a>'
    )


def _card_info_html(titulo, lineas):
    """HTML compacto (título + líneas) para el texto de una card del Panel de
    Control: todo en un solo bloque, para no acumular el margen que deja
    Streamlit entre cada markdown/caption suelto. .sb-card-line trunca con
    "..." (white-space: nowrap + text-overflow: ellipsis) para que la card
    no crezca con nombres/productos largos — se repite el texto en title=
    para que al pasar el mouse se pueda leer completo en vez de quedar
    cortado sin forma de verlo."""
    partes = [f'<div class="sb-card-title">{html.escape(str(titulo))}</div>']
    for linea in lineas:
        if linea:
            texto = html.escape(str(linea))
            partes.append(f'<div class="sb-card-line" title="{texto}">{texto}</div>')
    return "".join(partes)


def _badge(texto, color):
    """Pill de una palabra (etapa de CRM, estado de importación, "vencido")
    con el color de fondo al 13% y el texto al color pleno — un solo look
    para todos los badges de la app en vez de reinventarlo por sección."""
    return (
        f'<span style="background:{color}22; color:{color}; padding:2px 9px; '
        f'border-radius:4px; font-size:11px; font-weight:700; text-transform:uppercase; '
        f'letter-spacing:0.03em; white-space:nowrap;">{html.escape(texto)}</span>'
    )


def _fmt_tamano(num_bytes):
    num_bytes = num_bytes or 0
    valor = float(num_bytes)
    for unidad in ["B", "KB", "MB", "GB"]:
        if valor < 1024 or unidad == "GB":
            return f"{valor:.0f} {unidad}" if unidad == "B" else f"{valor:.1f} {unidad}"
        valor /= 1024


def _fmt_fecha_hora(valor):
    if not valor:
        return "-"
    try:
        return datetime.strptime(valor, "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y %H:%M")
    except (TypeError, ValueError):
        return valor


def _slug_archivo(texto):
    """Nombre de archivo seguro: conserva espacios y mayúsculas/minúsculas
    originales, solo saca los caracteres inválidos en un nombre de archivo."""
    return re.sub(r'[\\/:*?"<>|]+', "", (texto or "").strip()) or "sin datos"


def _nombre_pdf_cotizacion(cab, resultado):
    """Nombre de archivo común a los 2 PDFs de una cotización: (número)
    cliente - producto principal (el de mayor FOB total, si hay varios) —
    así cada PDF se identifica solo en la carpeta de descargas, sin tener
    que abrirlo. cab necesita 'numero' y 'cliente_nombre' (ya presentes en
    la fila de db.get_cotizacion / en el cab_full armado a mano del editor)."""
    productos_pdf = resultado.get("productos") or []
    producto_principal = max(productos_pdf, key=lambda p: p.get("fob_total", 0), default=None) if productos_pdf else None
    numero_pdf = f"({_slug_archivo(cab['numero'])})"
    resto_nombre = [
        cab.get("cliente_nombre") or "sin cliente",
        producto_principal.get("descripcion") if producto_principal else None,
    ]
    return numero_pdf + " " + " - ".join(_slug_archivo(p) for p in resto_nombre if p)


def _split_ext(nombre_archivo):
    """Separa 'factura.pdf' en ('factura', '.pdf'). Al renombrar solo se
    edita la base — la extensión se reengancha siempre, así no se rompe la
    detección de tipo (es_pdf/es_excel) que usan la vista previa y la
    conversión a PDF."""
    if "." in nombre_archivo:
        base, ext = nombre_archivo.rsplit(".", 1)
        return base, "." + ext
    return nombre_archivo, ""


def money(v, symbol="USD"):
    try:
        return f"{symbol} {v:,.2f}"
    except (TypeError, ValueError):
        return f"{symbol} 0.00"


def pct(v):
    try:
        return f"{v*100:.2f}%"
    except (TypeError, ValueError):
        return "0.00%"


def _gate_login():
    """Corre como primer paso de main(): si no hay sesión iniciada, muestra
    el login, centrado y con la misma identidad visual del resto de la
    app, y corta la ejecución con st.stop(). A propósito NO ofrece
    crear cuentas acá — eso es exclusivo de seed_usuario.py (por
    consola) para el primer usuario, y de Configuración (ya logueado)
    para los siguientes."""
    if st.session_state.get("usuario_autenticado"):
        return

    st.markdown("<style>[data-testid='stSidebar']{display:none;}</style>", unsafe_allow_html=True)

    _, col_centro, _ = st.columns([1, 1.1, 1])
    with col_centro:
        st.markdown(
            '<div class="sb-login-logo">SKY<span>BRIDGE</span></div>'
            '<div class="sb-login-tagline">ERP · Comercio Exterior</div>',
            unsafe_allow_html=True,
        )
        with st.container(key="sb_login_card"):
            if not db.hay_usuarios():
                setup_token_real = os.environ.get("SETUP_TOKEN")
                if setup_token_real:
                    st.markdown("**Restaurar base de datos**")
                    st.caption("Subí tu archivo skybridge.db y el código de configuración del servidor.")
                    with st.form("form_restaurar_db"):
                        archivo = st.file_uploader("Archivo skybridge.db", type=["db"])
                        token_ingresado = st.text_input("Código de configuración", type="password")
                        if st.form_submit_button("Restaurar", use_container_width=True):
                            if token_ingresado != setup_token_real:
                                st.error("Código incorrecto.")
                            elif archivo is None:
                                st.error("Subí el archivo primero.")
                            else:
                                db.restaurar_desde_sqlite_bytes(archivo.getvalue())
                                st.success("Base restaurada correctamente. Recargá la página para entrar.")
                                st.stop()
                else:
                    st.warning("Todavía no hay usuarios creados. Corré `python seed_usuario.py` en la consola del servidor para dar de alta el primero.")
            else:
                st.markdown("**Iniciar sesión**")
                with st.form("form_login"):
                    username = st.text_input("Usuario")
                    password = st.text_input("Contraseña", type="password")
                    if st.form_submit_button("Ingresar", use_container_width=True):
                        usuario, error = db.verificar_login(username, password)
                        if usuario:
                            st.session_state.usuario_autenticado = usuario
                            st.rerun()
                        else:
                            st.error(error)
    st.stop()


def _flash(msg, icon="✅"):
    """Guarda un mensaje para mostrarlo tras el próximo rerun.

    st.toast/st.success no pueden invocarse dentro de un callback (on_click),
    así que la mutación de estado se hace en el callback y el mensaje se
    muestra recién en el siguiente render normal del script.
    """
    st.session_state["_flash_msg"] = msg
    st.session_state["_flash_icon"] = icon


def _mostrar_flash():
    msg = st.session_state.pop("_flash_msg", None)
    if msg:
        st.toast(msg, icon=st.session_state.pop("_flash_icon", "✅"))


# ---------------------------------------------------------------- CLIENTES
# Documentación de una importación, agrupada en 2 apartados (provisorios /
# oficiales) dentro del tab "📎 Documentación" de la ficha de importación —
# ver _render_contenido_importacion. "Provisorios (sin clasificar)" es el
# nombre viejo de categoría ("Provisorios", una sola bolsa para los 4 tipos
# mezclados) de antes de esta separación más fina: NO se migran datos ni se
# adivina a cuál tipo nuevo corresponde cada archivo viejo — esa categoría
# sigue existiendo tal cual en la base y solo se muestra (dentro del grupo
# Provisorios) en las importaciones que ya tengan algo cargado ahí, así no
# se pierde ni se reclasifica a ciegas ni un archivo ya subido.
DOC_CATEGORIA_PROVISORIOS_LEGACY = "Provisorios"
DOC_CATEGORIAS_PROVISORIOS = [
    "Invoice Provisorio", "Packing List Provisorio", "BL Provisorio",
    "Despacho Provisorio", "Cotización Presentada",
]
DOC_CATEGORIAS_OFICIALES = ["Invoice", "Packing List", "Bill of Lading", "Despacho"]
DOC_CATEGORIA_FACTURAS = "Facturas"
DOC_CATEGORIA_PAGOS = "Pagos"
DOC_CATEGORIAS = (
    DOC_CATEGORIAS_PROVISORIOS + DOC_CATEGORIAS_OFICIALES
    + [DOC_CATEGORIA_FACTURAS, DOC_CATEGORIA_PAGOS, DOC_CATEGORIA_PROVISORIOS_LEGACY]
)
ESTADOS_IMPORTACION = ["En producción", "En coordinación", "En tránsito", "En puerto/Dep.Fiscal", "Entregada"]
COLOR_ESTADO_IMPORTACION = {
    "En producción": "#64748B", "En coordinación": "#2563EB", "En tránsito": "#D97706",
    "En puerto/Dep.Fiscal": "#7C3AED", "Entregada": "#16A34A",
}

OPCIONES_FORMATO_ENVIO = {
    "Envío aéreo": ["Aéreo c/despacho", "Courier Comercial"],
    "Envío marítimo": ["LCL", "FCL20", "FCL40", "Courier Comercial"],
}
TIPOS_ENVIO = list(OPCIONES_FORMATO_ENVIO.keys())
ICONO_TIPO_ENVIO = {"Envío aéreo": "✈️", "Envío marítimo": "🚢"}


def _guardar_archivos_subidos(imp_id, categoria, archivos):
    for f in archivos:
        db.add_documento(imp_id, categoria, f.name, f.getvalue())


def _eliminar_documento(doc_id):
    # El contenido vive en la misma fila (columna 'contenido' en Turso), no
    # en disco — borrar la fila alcanza, no hay ningún archivo aparte que
    # limpiar.
    db.delete_documento(doc_id)


def _eliminar_documento_cliente(doc_id):
    db.delete_documento_cliente(doc_id)


@st.dialog("Eliminar documento")
def _dialog_eliminar_documento(doc_id, nombre_archivo):
    st.warning(f"¿Confirmás eliminar **{nombre_archivo}**? Esta acción no se puede deshacer.")
    c1, c2 = st.columns(2)
    if c1.button("Cancelar", use_container_width=True, key=f"canceldeldoc_{doc_id}"):
        st.rerun()
    if c2.button("Sí, eliminar", type="primary", use_container_width=True, key=f"confirmdeldoc_{doc_id}"):
        try:
            _eliminar_documento(doc_id)
        except db.DBError as e:
            _flash(f"No se pudo eliminar el documento: {e}", icon="❌")
        st.rerun()


@st.dialog("Eliminar documento")
def _dialog_eliminar_documento_cliente(doc_id, nombre_archivo):
    st.warning(f"¿Confirmás eliminar **{nombre_archivo}**? Esta acción no se puede deshacer.")
    c1, c2 = st.columns(2)
    if c1.button("Cancelar", use_container_width=True, key=f"canceldelcliedoc_{doc_id}"):
        st.rerun()
    if c2.button("Sí, eliminar", type="primary", use_container_width=True, key=f"confirmdelcliedoc_{doc_id}"):
        try:
            _eliminar_documento_cliente(doc_id)
        except db.DBError as e:
            _flash(f"No se pudo eliminar el documento: {e}", icon="❌")
        st.rerun()


@st.dialog("Eliminar importación")
def _dialog_eliminar_importacion(imp_id, numero):
    st.warning(
        f"¿Confirmás eliminar **{numero}**? Esta acción no se puede deshacer: se borran todos "
        "sus documentos junto con la importación."
    )
    c1, c2 = st.columns(2)
    if c1.button("Cancelar", use_container_width=True, key=f"cancelimpdel_{imp_id}"):
        st.rerun()
    if c2.button("Sí, eliminar", type="primary", use_container_width=True, key=f"confirmimpdel_{imp_id}"):
        try:
            db.backup_antes_de_borrar("importacion")
            db.delete_importacion(imp_id)
            _flash("Importación eliminada.", icon="🗑️")
        except db.DBError as e:
            _flash(f"No se pudo eliminar la importación: {e}", icon="❌")
        st.rerun()


@st.dialog("Eliminar cliente")
def _dialog_eliminar_cliente(cliente_id, nombre):
    # cliente_id en importaciones/cotizaciones/contacts tiene ON DELETE
    # CASCADE/SET NULL en el schema (ver db.py) con foreign_keys=ON: al
    # borrar el cliente, SQLite se encarga solo de las importaciones
    # (CASCADE, se van, documentos incluidos) y las cotizaciones/contacto
    # vinculado (SET NULL, quedan pero desvinculados) — el contenido de los
    # documentos vive en la misma fila de la base (no en disco), así que el
    # cascade alcanza y no hace falta limpiar nada aparte.
    importaciones_cli = db.list_importaciones(cliente_id)
    total_docs_imp = sum(db.contar_documentos(imp["id"]) for imp in importaciones_cli)
    total_docs_cliente = db.contar_documentos_cliente(cliente_id)
    st.warning(
        f"¿Confirmás eliminar **{nombre}**? Esta acción no se puede deshacer, y también elimina "
        f"de forma permanente sus {len(importaciones_cli)} importación(es) y "
        f"{total_docs_imp + total_docs_cliente} documento(s) asociados.\n\n"
        "Las cotizaciones ya hechas para este cliente **no** se borran, pero quedan sin cliente "
        "vinculado. Si tiene un contacto de CRM vinculado, tampoco se borra: solo queda "
        "desvinculado del cliente."
    )
    c1, c2 = st.columns(2)
    if c1.button("Cancelar", use_container_width=True, key=f"cancelclidel_{cliente_id}"):
        st.rerun()
    if c2.button("Sí, eliminar", type="primary", use_container_width=True, key=f"confirmclidel_{cliente_id}"):
        try:
            db.backup_antes_de_borrar("cliente")
            db.delete_cliente(cliente_id)
            st.session_state.cliente_seleccionado = None
            _flash(f"Cliente '{nombre}' eliminado.", icon="🗑️")
        except db.DBError as e:
            _flash(f"No se pudo eliminar el cliente: {e}", icon="❌")
        st.rerun()


def _render_documentos_cliente(cliente_id):
    """Documentos generales del cliente (contratos, poderes, etc.), guardados
    en archivos_clientes/{cliente_id}/ — listado tipo tabla (Nombre, Fecha,
    Tamaño) con descarga/eliminación por archivo, más el cargador."""
    docs = db.list_documentos_cliente(cliente_id)
    if docs:
        h1, h2, h3, h4, h5, h6 = st.columns([2.9, 1.4, 0.9, 0.5, 1, 1])
        h1.markdown("**Nombre**")
        h2.markdown("**Fecha**")
        h3.markdown("**Tamaño**")
        h4.write("")
        h5.write("")
        h6.write("")
        for d in docs:
            c1, c2, c3, c4, c5, c6 = st.columns([2.9, 1.4, 0.9, 0.5, 1, 1])
            c1.write(f"📄 {d['nombre_archivo']}")
            c2.caption(_fmt_fecha_hora(d["subido_en"]))
            c3.caption(_fmt_tamano(d["tamano_bytes"]))
            with c4.popover("✏️", use_container_width=True):
                base_actual, ext_actual = _split_ext(d["nombre_archivo"])
                nuevo_base = st.text_input(
                    "Nuevo nombre", value=base_actual, key=f"renombrarcliedoc_{d['id']}",
                )
                if st.button("💾 Guardar", key=f"renombrarcliedocbtn_{d['id']}"):
                    if nuevo_base.strip():
                        db.rename_documento_cliente(d["id"], nuevo_base.strip() + ext_actual)
                        st.rerun()
                    else:
                        st.error("El nombre no puede quedar vacío.")
            c5.download_button(
                "Descargar", data=d["contenido"] or b"",
                file_name=d["nombre_archivo"], key=f"dlcliedoc_{d['id']}", use_container_width=True,
            )
            if c6.button("🗑️ Eliminar", key=f"delcliedoc_{d['id']}", use_container_width=True):
                _dialog_eliminar_documento_cliente(d["id"], d["nombre_archivo"])
    else:
        st.caption("Sin documentos cargados todavía.")

    st.divider()
    contador_key = f"upl_ctr_cliente_{cliente_id}"
    contador = st.session_state.get(contador_key, 0)
    subidos = st.file_uploader(
        "Subir documento(s)", accept_multiple_files=True, key=f"upl_cliente_{cliente_id}_{contador}",
    )
    if subidos:
        for f in subidos:
            contenido = f.getvalue()
            db.add_documento_cliente(cliente_id, f.name, contenido, len(contenido))
        st.session_state[contador_key] = contador + 1
        st.rerun()


def _render_proveedores_cliente(cliente_id):
    """Proveedores únicos detectados en las importaciones de este cliente,
    con el país de origen y cuántos embarques tiene cada uno."""
    importaciones = db.list_importaciones(cliente_id)
    proveedores = {}
    for imp in importaciones:
        nombre = (imp.get("proveedor_nombre") or "").strip()
        if not nombre:
            continue
        datos = proveedores.setdefault(nombre, {"pais": imp.get("proveedor_pais") or "-", "embarques": 0})
        datos["embarques"] += 1

    if not proveedores:
        st.info("Todavía no hay proveedores registrados en las importaciones de este cliente.")
        return

    filas = [
        {"Proveedor": nombre, "País / Origen": datos["pais"], "Total de Embarques": datos["embarques"]}
        for nombre, datos in sorted(proveedores.items())
    ]
    st.dataframe(pd.DataFrame(filas), hide_index=True, use_container_width=True)


def _render_documentos_categoria(imp_id, categoria):
    """Lista + descarga + carga de archivos de UNA categoría de documento —
    extraído de _render_documentos_importacion para poder usarse tanto
    dentro de un tab (varias categorías) como solo (una categoría propia,
    ej. el apartado de Facturas)."""
    if categoria == DOC_CATEGORIA_PROVISORIOS_LEGACY:
        st.caption(
            "Archivos cargados antes de separar los provisorios por tipo — sin clasificar entre "
            "Invoice/Packing List/BL/Despacho/Cotización presentada. Los nuevos que subas acá van "
            "a seguir juntándose en este mismo bucket viejo; para clasificar, subilos en la "
            "categoría correspondiente de arriba."
        )
    docs = db.list_documentos(imp_id, categoria)
    if docs:
        for d in docs:
            contenido = d["contenido"] or b""
            es_pdf = d["nombre_archivo"].lower().endswith(".pdf")
            es_excel = d["nombre_archivo"].lower().endswith((".xlsx", ".xls"))
            c1, c2, c3, c4, c5 = st.columns([2.7, 0.5, 1.1, 1.3, 0.5])
            c1.write(f"📄 {d['nombre_archivo']}")
            with c2.popover("✏️", use_container_width=True):
                base_actual, ext_actual = _split_ext(d["nombre_archivo"])
                nuevo_base = st.text_input(
                    "Nuevo nombre", value=base_actual, key=f"renombrardoc_{d['id']}",
                )
                if st.button("💾 Guardar", key=f"renombrardocbtn_{d['id']}"):
                    if nuevo_base.strip():
                        db.rename_documento(d["id"], nuevo_base.strip() + ext_actual)
                        st.rerun()
                    else:
                        st.error("El nombre no puede quedar vacío.")
            c3.download_button(
                "Descargar", data=contenido,
                file_name=d["nombre_archivo"], key=f"dl_{d['id']}", use_container_width=True,
            )
            if es_excel and c4.button("🔄 A PDF", key=f"conv_{d['id']}", use_container_width=True):
                with st.spinner("Convirtiendo a PDF…"):
                    pdf_bytes = storage.convertir_a_pdf(contenido, d["nombre_archivo"])
                if pdf_bytes:
                    nombre_pdf = f"{d['nombre_archivo'].rsplit('.', 1)[0]} (convertido).pdf"
                    db.add_documento(imp_id, categoria, nombre_pdf, pdf_bytes)
                    st.rerun()
                else:
                    st.error("No se pudo convertir el archivo a PDF. Verificá que LibreOffice esté instalado en el servidor.")
            if c5.button("🗑️", key=f"deldoc_{d['id']}", use_container_width=True):
                _dialog_eliminar_documento(d["id"], d["nombre_archivo"])
            if es_pdf and contenido:
                with st.expander("👁️ Vista previa", expanded=False):
                    st.pdf(contenido, key=f"pdfprev_{d['id']}")
    else:
        st.caption("Sin archivos cargados todavía.")

    # La key incluye un contador que se incrementa después de cada
    # carga: así el uploader "se vacía" y no reprocesa los mismos
    # archivos en el siguiente rerun.
    contador_key = f"upl_ctr_{imp_id}_{categoria}"
    contador = st.session_state.get(contador_key, 0)
    subidos = st.file_uploader(
        f"Subir {categoria.lower()}", accept_multiple_files=True,
        key=f"upl_{imp_id}_{categoria}_{contador}",
    )
    if subidos:
        _guardar_archivos_subidos(imp_id, categoria, subidos)
        st.session_state[contador_key] = contador + 1
        st.rerun()


def _render_documentos_importacion(imp_id, categorias=None):
    """Un tab por tipo de documento dentro de `categorias`, o el contenido
    directo si es una sola categoría (así se usa para los apartados de una
    sola categoría — Facturas, Pagos — sin una tira de un solo tab)."""
    categorias = categorias if categorias is not None else DOC_CATEGORIAS
    if len(categorias) == 1:
        _render_documentos_categoria(imp_id, categorias[0])
        return
    # key explícita (incluye la primera categoría, distinta entre el grupo
    # Provisorios y el grupo Oficiales): incluso dentro del mismo tab
    # "Documentación", dos st.tabs() sin key pueden confundirse entre sí al
    # navegar — mismo motivo que la key de imptabs_/fichaclitabs_ más arriba.
    tabs = st.tabs(categorias, key=f"doctabs_{imp_id}_{categorias[0]}")
    for tab, categoria in zip(tabs, categorias):
        with tab:
            _render_documentos_categoria(imp_id, categoria)


def _categorias_provisorios_para(imp_id):
    """Los 5 tipos provisorios granulares, más el bucket viejo 'Provisorios'
    (sin discriminar por tipo) SOLO si esta importación ya tiene archivos
    ahí — ver la nota junto a DOC_CATEGORIA_PROVISORIOS_LEGACY más arriba.
    Se agrega el nombre real de la categoría ("Provisorios"), no un alias,
    para que la carga/descarga siga apuntando a los mismos archivos ya
    guardados; el tab se distingue con una etiqueta propia dentro de
    _render_documentos_categoria en vez de cambiar la clave acá."""
    categorias = list(DOC_CATEGORIAS_PROVISORIOS)
    if db.list_documentos(imp_id, DOC_CATEGORIA_PROVISORIOS_LEGACY):
        categorias.append(DOC_CATEGORIA_PROVISORIOS_LEGACY)
    return categorias


def _render_detalle_final_importacion(imp, cotizacion):
    st.caption("Pagos reales realizados para esta carga.")
    c1, c2, c3 = st.columns(3)
    pago_proveedor = c1.number_input("Pago a proveedor (USD)", value=_f_local(imp.get("pago_proveedor_usd")), format="%.2f", key=f"impg_prov_{imp['id']}")
    costo_financiero = c2.number_input("Costo financiero (USD)", value=_f_local(imp.get("costo_financiero_usd")), format="%.2f", key=f"impg_cf_{imp['id']}")
    vep_tributos = c3.number_input("VEP Tributos (USD)", value=_f_local(imp.get("vep_tributos_usd")), format="%.2f", key=f"impg_vep_{imp['id']}")

    c4, c5 = st.columns(2)
    flete = c4.number_input("Flete (USD)", value=_f_local(imp.get("flete_usd")), format="%.2f", key=f"impg_flete_{imp['id']}")
    gastos_oper = c5.number_input("Gastos operativos (USD)", value=_f_local(imp.get("gastos_operativos_usd")), format="%.2f", key=f"impg_gop_{imp['id']}")

    total_real = pago_proveedor + costo_financiero + vep_tributos + flete + gastos_oper

    m1, m2 = st.columns(2)
    m1.metric("Total pagado final", money(total_real))
    if cotizacion:
        total_presupuestado = cotizacion.get("total_usd_civa") or 0
        m2.metric("Total presupuestado", money(total_presupuestado))
    else:
        m2.metric("Total presupuestado", "—")
        m2.caption("Sin cotización vinculada")

    return {
        "pago_proveedor_usd": pago_proveedor, "costo_financiero_usd": costo_financiero,
        "vep_tributos_usd": vep_tributos, "flete_usd": flete, "gastos_operativos_usd": gastos_oper,
    }


def _render_contenido_importacion(imp, cotizaciones_cliente):
    """Ficha completa de una importación en 3 carpetas — mismo patrón de
    organización que Cotizaciones/Importaciones en la ficha de Clientes:
    'Detalles' (todo lo operativo: estado, mercadería, tránsito, fechas,
    detalle final de costos), 'Documentación' (provisorios/oficiales, cada
    uno su propio apartado) y 'Facturas y pagos' (facturas/pagos, ídem).
    Sin envoltorio propio, para poder mostrarse tanto a pantalla completa
    (ficha de cliente, Panel de Control) como veníamos haciendo."""
    tab_detalles, tab_documentacion, tab_facturas = st.tabs(
        ["📁 Detalles", "📎 Documentación", "🧾 Facturas y pagos"], key=f"imptabs_{imp['id']}",
    )

    with tab_detalles:
        cot_opciones = {None: "— Sin vincular —"}
        cot_opciones.update({c["id"]: c["numero"] for c in cotizaciones_cliente})
        cot_ids = list(cot_opciones.keys())

        c1, c2 = st.columns(2)
        estado = c1.selectbox("Estado", ESTADOS_IMPORTACION, index=ESTADOS_IMPORTACION.index(imp.get("estado") or ESTADOS_IMPORTACION[0]), key=f"impe_{imp['id']}")
        cotizacion_id = c2.selectbox(
            "Cotización vinculada", options=cot_ids, format_func=lambda x: cot_opciones[x],
            index=cot_ids.index(imp.get("cotizacion_id")) if imp.get("cotizacion_id") in cot_ids else 0,
            key=f"impc_{imp['id']}",
        )

        st.markdown("**Mercadería y proveedor**")
        c3, c4, c5, c6 = st.columns(4)
        producto = c3.text_input("Producto", value=imp.get("producto") or "", key=f"impprod_{imp['id']}")
        prov_nombre = c4.text_input("Proveedor", value=imp.get("proveedor_nombre") or "", key=f"impp_{imp['id']}")
        prov_contacto = c5.text_input("Contacto", value=imp.get("proveedor_contacto") or "", key=f"impk_{imp['id']}")
        prov_pais = c6.text_input("Origen", value=imp.get("proveedor_pais") or "", key=f"imppa_{imp['id']}")

        st.markdown("**Tránsito**")
        c7, c8, c9, c10 = st.columns(4)
        tipo_envio = c7.selectbox(
            "Tipo de envío", TIPOS_ENVIO,
            index=TIPOS_ENVIO.index(imp.get("tipo_envio")) if imp.get("tipo_envio") in TIPOS_ENVIO else 0,
            key=f"imptenv_{imp['id']}",
        )
        opciones_formato = OPCIONES_FORMATO_ENVIO[tipo_envio]
        formato_guardado = imp.get("formato_envio")
        # La key incluye tipo_envio: si se cambia el tipo, el desplegable de
        # Formato nace de nuevo con las opciones correctas (evita un formato
        # guardado de aéreo, ej. "Courier Comercial", quedar "vivo" con opciones
        # de marítimo, o viceversa).
        formato = c8.selectbox(
            "Formato", opciones_formato,
            index=opciones_formato.index(formato_guardado) if formato_guardado in opciones_formato else 0,
            key=f"impfmt_{imp['id']}_{tipo_envio}",
        )
        etd = c9.date_input("ETD", value=_parse_fecha(imp.get("etd")), key=f"impetd_{imp['id']}", format="DD/MM/YYYY")
        eta = c10.date_input("ETA", value=_parse_fecha(imp.get("eta")), key=f"impeta_{imp['id']}", format="DD/MM/YYYY")
        notas = st.text_area("Detalles a considerar (notas de tránsito)", value=imp.get("notas_transito") or "", key=f"impnt_{imp['id']}")

        st.markdown("**Fechas del proceso**")
        c11, c12, c13 = st.columns(3)
        fecha_llegada_puerto = c11.date_input("Fecha de llegada", value=_parse_fecha(imp.get("fecha_llegada_puerto")), key=f"impfllp_{imp['id']}", format="DD/MM/YYYY")
        fecha_oficializacion = c12.date_input("Fecha oficialización", value=_parse_fecha(imp.get("fecha_oficializacion")), key=f"impfof_{imp['id']}", format="DD/MM/YYYY")
        fecha_entrega = c13.date_input("Fecha de entrega", value=_parse_fecha(imp.get("fecha_entrega")), key=f"impfent_{imp['id']}", format="DD/MM/YYYY")

        st.divider()
        st.markdown("**Detalle final**")
        cot = next((c for c in cotizaciones_cliente if c["id"] == cotizacion_id), None) if cotizacion_id else None
        pagos = _render_detalle_final_importacion(imp, cot)

        colA, colB = st.columns(2)
        if colA.button("💾 Guardar importación", key=f"impsave_{imp['id']}", type="primary", use_container_width=True):
            db.update_importacion(
                imp["id"], estado=estado, cotizacion_id=cotizacion_id, producto=producto,
                proveedor_nombre=prov_nombre, proveedor_contacto=prov_contacto, proveedor_pais=prov_pais,
                tipo_envio=tipo_envio, formato_envio=formato,
                etd=etd.isoformat() if etd else None, eta=eta.isoformat() if eta else None,
                notas_transito=notas,
                fecha_llegada_puerto=fecha_llegada_puerto.isoformat() if fecha_llegada_puerto else None,
                fecha_oficializacion=fecha_oficializacion.isoformat() if fecha_oficializacion else None,
                fecha_entrega=fecha_entrega.isoformat() if fecha_entrega else None,
                **pagos,
            )
            st.success("Importación guardada.")
            st.rerun()
        if colB.button("🗑️ Eliminar importación", key=f"impdel_{imp['id']}", use_container_width=True):
            _dialog_eliminar_importacion(imp["id"], imp["numero"])

    with tab_documentacion:
        # 2 apartados, como pediste: provisorios (incluye "Cotización
        # presentada") separados de los oficiales/definitivos. Ningún
        # archivo cambia de categoría ni se pierde acá, solo el agrupamiento
        # visual (ver _render_documentos_importacion / _categorias_provisorios_para).
        st.markdown("#### 📝 Provisorios")
        _render_documentos_importacion(imp["id"], _categorias_provisorios_para(imp["id"]))

        st.divider()
        st.markdown("#### ✅ Oficiales")
        _render_documentos_importacion(imp["id"], DOC_CATEGORIAS_OFICIALES)

    with tab_facturas:
        st.markdown("#### 🧾 Facturas")
        _render_documentos_importacion(imp["id"], [DOC_CATEGORIA_FACTURAS])

        st.divider()
        st.markdown("#### 💳 Pagos")
        _render_documentos_importacion(imp["id"], [DOC_CATEGORIA_PAGOS])


def _render_fila_importacion_cliente(imp):
    """Card 'hoja' de una importación en la ficha de cliente — mismo
    lenguaje visual que Cotizaciones/Clientes/CRM: número/tipo de envío,
    proveedor/origen, badge de estado y un único botón 'Ver detalle →' que
    abre la ficha completa (antes era un desplegable con todo el formulario
    adentro, incómodo con varias importaciones)."""
    icono = ICONO_TIPO_ENVIO.get(imp.get("tipo_envio"), "")
    with st.container(border=True, key=f"impfila_{imp['id']}"):
        c1, c2, c3, c4 = st.columns([2.2, 2.6, 1.6, 1.8], vertical_alignment="center")
        prefijo = f"{icono} " if icono else ""
        c1.markdown(
            f'<div style="font-weight:700; font-size:14px;">{prefijo}{html.escape(imp["numero"])}</div>'
            f'<div style="font-size:12px; color:#64748b; margin-top:2px;">{html.escape(imp.get("tipo_envio") or "—")}</div>',
            unsafe_allow_html=True,
        )
        c2.markdown(
            f'<div style="font-size:13px; font-weight:600;">{html.escape(imp.get("proveedor_nombre") or "sin proveedor")}</div>'
            f'<div style="font-size:12px; color:#64748b;">{html.escape(imp.get("proveedor_pais") or "—")}</div>',
            unsafe_allow_html=True,
        )
        c3.markdown(
            _badge_estado_importacion(imp.get("estado") or ESTADOS_IMPORTACION[0]), unsafe_allow_html=True,
        )
        # Link real armado a mano (target="_self", no target="_blank" fijo
        # de st.link_button) en vez de un st.button con on_click: navega
        # por query params, así arranca una sesión de Streamlit nueva de
        # punta a punta — main() lee cliente_ver/imp_ver al arrancar y
        # reconstruye el estado (ficha de cliente abierta en su
        # importación) antes de dibujar nada.
        c4.markdown(
            _link_button_html(
                "Ver detalle →", f"?cliente_ver={imp['cliente_id']}&imp_ver={imp['id']}",
            ).replace('class="sb-btn-link"', 'class="sb-btn-link sb-verdetalle-imp"'),
            unsafe_allow_html=True,
        )


def _render_ficha_importacion_cliente(imp):
    """Página completa de una importación, abierta desde su card en la
    ficha de cliente — mismo contenido que antes vivía adentro del
    desplegable (_render_contenido_importacion, sin cambios), solo que
    ahora es su propia pantalla en vez de expandirse inline entre las
    demás importaciones."""
    def _volver():
        st.session_state.importacion_seleccionada_cliente = None

    st.button(
        "Volver a la ficha del cliente", icon=":material/arrow_back:",
        on_click=_volver, key=f"volver_ficha_imp_{imp['id']}",
    )

    cliente = db.get_cliente(imp["cliente_id"])
    icono = ICONO_TIPO_ENVIO.get(imp.get("tipo_envio"), "")
    st.subheader(f"{icono + ' ' if icono else ''}{imp['numero']} — {cliente['nombre'] if cliente else 'sin cliente'}")

    cotizaciones_cliente = db.list_cotizaciones(cliente_id=imp["cliente_id"])
    _render_contenido_importacion(imp, cotizaciones_cliente)


def _render_importaciones_cliente(cliente_id):
    cotizaciones_cliente = db.list_cotizaciones(cliente_id=cliente_id)
    importaciones = db.list_importaciones(cliente_id)

    if not importaciones:
        st.info("Todavía no hay importaciones cargadas para este cliente.")
    else:
        st.caption(f"{len(importaciones)} importación(es)")
        for imp in importaciones:
            _render_fila_importacion_cliente(imp)

    st.markdown("---")
    cot_opciones = {None: "— Sin vincular —"}
    cot_opciones.update({c["id"]: c["numero"] for c in cotizaciones_cliente})
    cot_ids = list(cot_opciones.keys())
    c1, c2 = st.columns([3, 1])
    cot_para_nueva = c1.selectbox(
        "Vincular nueva importación a una cotización (opcional)",
        options=cot_ids, format_func=lambda x: cot_opciones[x], key=f"nueva_imp_cot_{cliente_id}",
    )
    c2.write("")
    c2.write("")
    if c2.button("➕ Nueva importación", key=f"nueva_imp_btn_{cliente_id}", use_container_width=True):
        # El tipo de envío, formato y ETD/ETA de la cotización (si la tuviera)
        # no se prellenan: se cargan directo en la importación con los
        # desplegables y el calendario.
        db.create_importacion(cliente_id, cotizacion_id=cot_para_nueva)
        _marcar_ganado_por_cliente(cliente_id)
        st.rerun()


@st.cache_data
def _generar_pdfs_cotizacion(cot_id, actualizado_en):
    """Recalcula el resultado de una cotización ya guardada y genera sus
    dos PDFs (simplificado y completo) — mismo criterio de nombre de
    archivo que usa el Cotizador al descargarlos. Se llama on-demand: en la
    ficha de cliente para la cotización elegida, y en el historial del
    Cotizador para cada fila de la página actual (máx. 10, nunca para el
    listado completo sin paginar).

    Cacheada por (cot_id, actualizado_en): antes regeneraba 2 PDFs con
    reportlab en CADA rerun del historial — incluida cada letra tipeada en
    el buscador — para filas que ni siquiera cambiaron. actualizado_en (ya
    presente en la fila de db.list_cotizaciones, se bump ea en cada guardado
    o cambio de estado) alcanza como parte de la key: cambia solo cuando la
    cotización realmente cambió, así el caché se invalida solo."""
    cab = db.get_cotizacion(cot_id)
    productos = db.get_productos(cot_id)
    gastos = db.get_gastos(cot_id)
    resultado = calculo.calcular(cab, productos, gastos)
    pdf_simple = generar_pdf_cotizacion(cab, resultado, completo=False)
    pdf_completo = generar_pdf_cotizacion(cab, resultado, completo=True)
    nombre_pdf = _nombre_pdf_cotizacion(cab, resultado)
    return pdf_simple, pdf_completo, nombre_pdf


def _render_ficha_cliente(cli):
    """Vista dedicada de un cliente (se abre al clickear 'Ver ficha' en el
    listado), en vez de expandirlo inline entre todos los demás clientes."""

    def _volver_al_listado():
        st.session_state.cliente_seleccionado = None

    st.button(
        "Volver a Clientes", icon=":material/arrow_back:",
        on_click=_volver_al_listado, key=f"volver_ficha_{cli['id']}",
    )

    # Header ejecutivo tipo banner: tarjeta blanca con flexbox (info a la
    # izquierda en HTML/CSS propio, acciones rápidas —botones reales de
    # Streamlit, compactos— a la derecha), pegada a las pestañas de abajo.
    with st.container(key=f"fichahdr_{cli['id']}"):
        c_datos, c_acciones = st.columns([3, 2])
        with c_datos:
            linea_cuit_rubro = "  •  ".join([
                f"CUIT: {html.escape(cli.get('cuit') or '—')}",
                f"Rubro: {html.escape(cli.get('rubro') or '—')}",
            ])
            linea_email_tel = "  •  ".join([
                f"Email: {html.escape(cli.get('email') or '—')}",
                f"Tel: {html.escape(cli.get('telefono') or '—')}",
            ])
            st.markdown(
                f'<div class="sb-fichahdr-nombre">{html.escape(cli["nombre"])}</div>'
                f'<div class="sb-fichahdr-datos">{linea_cuit_rubro}</div>'
                f'<div class="sb-fichahdr-datos">{linea_email_tel}</div>',
                unsafe_allow_html=True,
            )
        with c_acciones:
            b1, b2 = st.columns(2)
            b1.button(
                "➕ Cotizar", key=f"hdrcot_{cli['id']}", type="secondary",
                use_container_width=True, on_click=_cotizar_cliente, args=(cli["id"],),
            )
            if b2.button(
                "➕ Importación", key=f"hdrimp_{cli['id']}", type="secondary", use_container_width=True,
            ):
                db.create_importacion(cli["id"])
                # Evidencia dura de cliente ganado: si el contacto vinculado
                # (o recién creado) no estaba ya en una etapa posterior, pasa
                # a Ganado. Mismo criterio que "cotización Aprobada".
                _marcar_ganado_por_cliente(cli["id"])
                _flash("Importación creada.")
                st.rerun()

    # key explícita: sin ella, Streamlit puede confundir este grupo de tabs
    # con el de _render_contenido_importacion (que comparte la etiqueta
    # "📎 Documentación") al navegar entre la ficha del cliente y la ficha de
    # una importación — se vieron mezclados en un solo tab bar de 5+3 tabs.
    tab_cotizaciones, tab_importaciones, tab_datos, tab_documentos, tab_proveedores = st.tabs(
        ["📑 Cotizaciones", "📦 Importaciones", "📁 Datos del Cliente", "📎 Documentación", "🏭 Proveedores"],
    )

    with tab_cotizaciones:
        # Mismo card que el historial del Cotizador (_render_fila_cotizacion):
        # editar/PDF completo/ver detalle/duplicar/eliminar quedan resueltos
        # ahí mismo por fila, sin el desplegable "elegí una cotización" que
        # había antes para recién ahí descargar o previsualizar.
        cot_cli = [c for c in db.list_cotizaciones() if c["cliente_id"] == cli["id"]]
        if cot_cli:
            st.caption(f"{len(cot_cli)} cotización(es)")
            etapa_por_cliente = _mapa_etapa_por_cliente()
            for c in cot_cli:
                _render_fila_cotizacion(c, etapa_por_cliente=etapa_por_cliente)
        else:
            st.caption("Sin cotizaciones todavía.")

    with tab_importaciones:
        _render_importaciones_cliente(cli["id"])

    with tab_datos:
        with st.form(f"edit_cliente_{cli['id']}"):
            c1, c2 = st.columns(2)
            nombre = c1.text_input("Nombre", value=cli["nombre"], key=f"n_{cli['id']}")
            cuit = c2.text_input("CUIT", value=cli["cuit"] or "", key=f"c_{cli['id']}")
            c3, c4 = st.columns(2)
            email = c3.text_input("Email", value=cli["email"] or "", key=f"e_{cli['id']}")
            telefono = c4.text_input("Teléfono", value=cli["telefono"] or "", key=f"t_{cli['id']}")
            direccion = st.text_input("Dirección", value=cli["direccion"] or "", key=f"d_{cli['id']}")
            c5, c6 = st.columns(2)
            productos_interes = c5.text_input("Productos de interés", value=cli.get("productos_interes") or "", key=f"pi_{cli['id']}")
            rubro = c6.text_input("Rubro", value=cli.get("rubro") or "", key=f"ru_{cli['id']}")
            notas = st.text_area("Notas / información relevante", value=cli["notas"] or "", key=f"nt_{cli['id']}")
            b1, b2 = st.columns(2)
            if b1.form_submit_button("💾 Guardar cambios"):
                db.update_cliente(cli["id"], nombre, cuit, email, telefono, direccion, rubro, productos_interes, notas)
                st.success("Actualizado.")
                st.rerun()
            if b2.form_submit_button("🗑️ Eliminar cliente"):
                _dialog_eliminar_cliente(cli["id"], cli["nombre"])

    with tab_documentos:
        _render_documentos_cliente(cli["id"])

    with tab_proveedores:
        _render_proveedores_cliente(cli["id"])


CLIENTES_POR_PAGINA = 10


def _render_fila_cliente(cli, cot_por_cliente, en_proceso_por_cliente, completadas_por_cliente, etapa_por_cliente=None):
    with st.container(border=True, key=f"clientecard_{cli['id']}"):
        cot_count = cot_por_cliente.get(cli["id"], 0)
        en_proceso = en_proceso_por_cliente.get(cli["id"], 0)
        completadas = completadas_por_cliente.get(cli["id"], 0)

        c1, c2, c3, c4 = st.columns([4, 3, 1.5, 1.5], vertical_alignment="center")

        # Nombre + rubro + etapa CRM (si el cliente tiene un contacto
        # vinculado) en un único bloque HTML: si van en varios st.markdown
        # separados, el espacio entre elementos de Streamlit (no el margin
        # propio) los separa demasiado del nombre. La etapa se lee en vivo
        # de contacts.etapa (ver _mapa_etapa_por_cliente), así que un
        # cambio hecho desde la card del CRM se ve reflejado acá al toque.
        rubro_y_productos = " · ".join(v for v in (cli.get("rubro"), cli.get("productos_interes")) if v)
        etapa_crm = (etapa_por_cliente or {}).get(cli["id"])
        badge_html = f'<div style="margin-top:4px;">{_badge_etapa(etapa_crm)}</div>' if etapa_crm else ""
        # Antes esta línea se mostraba siempre, con "—" como fallback cuando
        # el cliente no tiene rubro ni producto cargado (frecuente en
        # "Cotizados", que son clientes recién pasados del CRM). Un renglón
        # completo solo para un guión huérfano se leía como un dato roto —
        # directamente no se renderiza la línea si no hay nada que mostrar.
        rubro_html = (
            f'<div style="font-size: 12px; color: #64748b; margin-top: 2px; line-height: 1.3;">'
            f'{html.escape(rubro_y_productos)}</div>'
        ) if rubro_y_productos else ""
        c1.markdown(
            f'<div style="font-weight: 700; font-size: 14px;">{html.escape(cli["nombre"])}</div>'
            f'{rubro_html}{badge_html}',
            unsafe_allow_html=True,
        )

        # Un único bloque HTML con line-height ajustado: 3 st.caption sueltos
        # dejaban un espacio muerto entre renglones que engordaba la card.
        c2.markdown(
            # Mismo ícono (🚢) para "en proceso" y "completadas": son la
            # misma cosa (una importación) en dos estados distintos, no dos
            # conceptos separados — antes usaban 📦 para "completadas" y se
            # leía como si fuera otra categoría de dato.
            f'<div style="font-size: 12px; color: #64748b; line-height: 1.3;">'
            f'<div>📑 Cotizaciones realizadas: <b>{cot_count}</b></div>'
            f'<div>🚢 Importaciones en proceso: <b>{en_proceso}</b></div>'
            f'<div>🚢 Importaciones completadas: <b>{completadas}</b></div>'
            f'</div>',
            unsafe_allow_html=True,
        )

        # Ambos botones ocupan el ancho de su propia columna: "+ Cotizar"
        # queda en la columna central (3) y "Ver ficha →" en la última (4),
        # pegado al borde derecho de la card.
        c3.button(
            "＋ Cotizar", key=f"cotizarcliente_{cli['id']}", use_container_width=True,
            on_click=_cotizar_cliente, args=(cli["id"],), type="primary",
        )
        if c4.button("Ver ficha →", key=f"vercliente_{cli['id']}", use_container_width=True, type="secondary"):
            st.session_state.cliente_seleccionado = cli["id"]
            st.rerun()


def _render_seccion_clientes_importaciones(
    todos_los_clientes, imp_por_cliente, cot_por_cliente, en_proceso_por_cliente, completadas_por_cliente,
    etapa_por_cliente,
):
    """Pestaña 'Importaciones': clientes con al menos una importación
    realizada, en cualquier estado (incluidas ya entregadas) — la vista por
    default al entrar a Clientes."""
    c1, c2 = st.columns([1, 1])
    search = c1.text_input(
        "Buscar por nombre / razón social", placeholder="Buscar por nombre / razón social...",
        label_visibility="collapsed", key="clientes_search_imp",
    )
    rubro_producto = c2.text_input(
        "Buscar por rubro o producto", placeholder="Buscar por rubro o producto...",
        label_visibility="collapsed", key="clientes_rubro_imp",
    )

    clientes = [c for c in todos_los_clientes if imp_por_cliente.get(c["id"], 0) > 0]
    if search:
        b = search.strip().lower()
        clientes = [c for c in clientes if b in (c.get("nombre") or "").lower()]
    if rubro_producto:
        b = rubro_producto.strip().lower()
        clientes = [
            c for c in clientes
            if b in (c.get("rubro") or "").lower() or b in (c.get("productos_interes") or "").lower()
        ]

    # Orden automático, sin intervención del usuario: de mayor a menor
    # cantidad de importaciones realizadas, y alfabético como desempate.
    clientes = sorted(
        clientes, key=lambda c: (-imp_por_cliente.get(c["id"], 0), (c.get("nombre") or "").lower()),
    )

    st.caption(f"{len(clientes)} cliente(s)")

    if not clientes:
        if search or rubro_producto:
            st.info("No se encontraron clientes con ese criterio de búsqueda.")
        else:
            st.info(
                "Todavía no hay clientes con importaciones realizadas. Los clientes 'Cotizados' en "
                "el CRM que todavía no importaron están en la pestaña '🧾 Cotizados'."
            )
        return

    # ---- Paginación: se reinicia a la página 1 si cambió algún filtro ----
    firma_filtro = (search, rubro_producto)
    if st.session_state.get("clientes_firma_filtro") != firma_filtro:
        st.session_state.clientes_firma_filtro = firma_filtro
        st.session_state.clientes_pagina = 1

    total_paginas = max(1, (len(clientes) - 1) // CLIENTES_POR_PAGINA + 1)
    pagina_actual = min(st.session_state.get("clientes_pagina", 1), total_paginas)
    inicio = (pagina_actual - 1) * CLIENTES_POR_PAGINA
    clientes_pagina = clientes[inicio:inicio + CLIENTES_POR_PAGINA]

    for cli in clientes_pagina:
        _render_fila_cliente(
            cli, cot_por_cliente, en_proceso_por_cliente, completadas_por_cliente, etapa_por_cliente,
        )

    if total_paginas > 1:
        cprev, cinfo, csig = st.columns([1, 2, 1])
        if cprev.button("← Anterior", disabled=(pagina_actual <= 1), key="clientes_pag_prev"):
            st.session_state.clientes_pagina = pagina_actual - 1
            st.rerun()
        cinfo.markdown(
            f'<p style="text-align:center; margin-top:0.5rem;">Página {pagina_actual} de {total_paginas}</p>',
            unsafe_allow_html=True,
        )
        if csig.button("Siguiente →", disabled=(pagina_actual >= total_paginas), key="clientes_pag_next"):
            st.session_state.clientes_pagina = pagina_actual + 1
            st.rerun()


def _render_seccion_clientes_cotizados(
    todos_los_clientes, imp_por_cliente, cot_por_cliente, en_proceso_por_cliente, completadas_por_cliente,
    etapa_por_cliente,
):
    """Pestaña 'Cotizados': clientes sin ninguna importación todavía, en
    etapa "Cotizado" o "Negociación" del CRM (crm.kanban_grupo las separa en
    dos columnas propias del tablero de Contactos, pero acá siguen
    contando como una sola bolsa: "cotizado, todavía sin importación") con
    un cliente vinculado. En cuanto tienen una importación pasan a la
    pestaña 'Importaciones' y dejan de listarse acá (no se duplican)."""
    c1, c2 = st.columns([1, 1])
    search = c1.text_input(
        "Buscar por nombre / razón social", placeholder="Buscar por nombre / razón social...",
        label_visibility="collapsed", key="clientes_search_cot",
    )
    rubro_producto = c2.text_input(
        "Buscar por rubro o producto", placeholder="Buscar por rubro o producto...",
        label_visibility="collapsed", key="clientes_rubro_cot",
    )

    contactos_cotizados = [
        c for c in db.list_contacts()
        if crm.kanban_grupo(c.get("etapa") or crm.STAGE_INICIAL) in ("cotizados", "negociacion")
        and c.get("cliente_id")
    ]
    ids_cotizados_crm = {c["cliente_id"] for c in contactos_cotizados}
    clientes = [
        c for c in todos_los_clientes
        if c["id"] in ids_cotizados_crm and imp_por_cliente.get(c["id"], 0) == 0
    ]
    if search:
        b = search.strip().lower()
        clientes = [c for c in clientes if b in (c.get("nombre") or "").lower()]
    if rubro_producto:
        b = rubro_producto.strip().lower()
        clientes = [
            c for c in clientes
            if b in (c.get("rubro") or "").lower() or b in (c.get("productos_interes") or "").lower()
        ]
    clientes = sorted(clientes, key=lambda c: (c.get("nombre") or "").lower())

    st.caption(f"{len(clientes)} cliente(s)")

    if not clientes:
        if search or rubro_producto:
            st.info("No se encontraron clientes con ese criterio de búsqueda.")
        else:
            st.info(
                "No hay clientes 'Cotizados' sin importaciones todavía. Esta lista se completa "
                "sola desde el embudo del CRM (pestaña '🤝 CRM' → columna 'Cotizados')."
            )
        return

    for cli in clientes:
        _render_fila_cliente(
            cli, cot_por_cliente, en_proceso_por_cliente, completadas_por_cliente, etapa_por_cliente,
        )


def vista_clientes():
    # La ficha de una importación (session_state.importacion_seleccionada_cliente)
    # se resuelve en main(), ANTES de llegar acá — no es una rama más de esta
    # función. Ver la nota grande en main() para el motivo (bug de
    # reconciliación de tabs de Streamlit al anidarla en esta misma función).
    cliente_sel_id = st.session_state.get("cliente_seleccionado")
    if cliente_sel_id is not None:
        cli = db.get_cliente(cliente_sel_id)
        if cli is None:
            st.session_state.cliente_seleccionado = None
            st.rerun()
        _render_ficha_cliente(cli)
        return

    # ---- Encabezado: título + botón "Nuevo cliente" (alterna el formulario) ----
    c_titulo, c_nuevo = st.columns([4, 1], vertical_alignment="center")
    c_titulo.title("Clientes")
    if c_nuevo.button("＋ Nuevo cliente", use_container_width=True, type="primary", key="clientes_toggle_nuevo"):
        st.session_state.clientes_mostrar_form_nuevo = not st.session_state.get("clientes_mostrar_form_nuevo", False)

    if st.session_state.get("clientes_mostrar_form_nuevo"):
        with st.form("form_nuevo_cliente", clear_on_submit=True):
            c1, c2 = st.columns(2)
            nombre = c1.text_input("Nombre / Razón social *")
            cuit = c2.text_input("CUIT")
            c3, c4 = st.columns(2)
            email = c3.text_input("Email")
            telefono = c4.text_input("Teléfono")
            direccion = st.text_input("Dirección")
            c5, c6 = st.columns(2)
            productos_interes = c5.text_input("Productos de interés")
            rubro = c6.text_input("Rubro")
            notas = st.text_area("Notas")
            if st.form_submit_button("Guardar cliente", type="primary"):
                if nombre.strip():
                    new_id = db.create_cliente(nombre.strip(), cuit, email, telefono, direccion, rubro, productos_interes, notas)
                    # Un cliente nace por acá o desde un contacto del CRM
                    # (_cotizar_contacto) — las dos puertas tienen que
                    # terminar con un contacto de CRM vinculado, para que
                    # nadie tenga que acordarse de crearlo a mano.
                    _asegurar_contacto_para_cliente(new_id)
                    st.session_state.cliente_seleccionado = new_id
                    st.session_state.clientes_mostrar_form_nuevo = False
                    st.success(f"Cliente '{nombre}' creado.")
                    st.rerun()
                else:
                    st.error("El nombre es obligatorio.")

    st.write("")

    # ---- Datos base para filtros y badges (una sola consulta cada uno) ----
    todos_los_clientes = db.list_clientes()
    todas_cotizaciones = db.list_cotizaciones()
    todas_importaciones = db.list_importaciones()

    cot_por_cliente, imp_por_cliente = {}, {}
    en_proceso_por_cliente, completadas_por_cliente = {}, {}
    for c in todas_cotizaciones:
        cid = c.get("cliente_id")
        if cid is not None:
            cot_por_cliente[cid] = cot_por_cliente.get(cid, 0) + 1
    for i in todas_importaciones:
        cid = i.get("cliente_id")
        if cid is None:
            continue
        imp_por_cliente[cid] = imp_por_cliente.get(cid, 0) + 1
        # "En proceso" = cualquier etapa de gestión activa, es decir todas
        # menos "Entregada" (el objetivo es saber cuántas importaciones
        # todavía requieren seguimiento, sin importar en qué etapa estén).
        # "Completadas" es el complemento: solo las ya entregadas.
        if (i.get("estado") or ESTADOS_IMPORTACION[0]) != "Entregada":
            en_proceso_por_cliente[cid] = en_proceso_por_cliente.get(cid, 0) + 1
        else:
            completadas_por_cliente[cid] = completadas_por_cliente.get(cid, 0) + 1

    # Dos pestañas (mismo patrón que "🧮 Cotizaciones"/"📦 Importaciones" en
    # Panel de Control): "Importaciones" es la vista por default, "Cotizados"
    # se abre a demanda. Cada una tiene su propio buscador (nombre/rubro),
    # su propia card (_render_fila_cliente, sin cambios) y su propio estado
    # vacío — están desacopladas a propósito, así un cliente que ya tiene
    # importaciones no compite por espacio/paginación con los que todavía
    # están cotizando.
    etapa_por_cliente = _mapa_etapa_por_cliente()
    tab_importaciones, tab_cotizados = st.tabs(["🚢 Importaciones", "🧾 Cotizados"])
    with tab_importaciones:
        _render_seccion_clientes_importaciones(
            todos_los_clientes, imp_por_cliente, cot_por_cliente,
            en_proceso_por_cliente, completadas_por_cliente, etapa_por_cliente,
        )
    with tab_cotizados:
        _render_seccion_clientes_cotizados(
            todos_los_clientes, imp_por_cliente, cot_por_cliente,
            en_proceso_por_cliente, completadas_por_cliente, etapa_por_cliente,
        )


# ---------------------------------------------------------------- COTIZADOR
COT_CONDICIONES_VENTA = ["FOB", "EXW", "FCA"]
COT_TIPOS_ENVIO = ["Marítimo", "Aéreo"]

# Prefijo de key de widget por campo. Se usa tanto para renderizar cada campo
# en la sección que le corresponde, como para "leer" su valor actual desde
# session_state ANTES de que su widget se renderice (ver _sync_desde_widgets):
# así, secciones que se muestran más arriba (ej. el CIF, que depende del
# Flete cargado más abajo en "Costos operativos") siempre reflejan el último
# valor cargado, igual que en una planilla de cálculo.
PROD_FIELD_PREFIX = {
    "descripcion": "pdesc", "ncm": "pncm", "fob_unit": "pfob", "cantidad": "pcant",
    "fob_decl_unit": "pfobd", "peso_kg": "ppesokg", "volumen_m3": "pvolm3",
    "pct_derechos": "pder", "pct_tasa_estadistica": "ptasa", "pct_antidumping": "panti",
    "pct_iva": "piva", "pct_iva_adicional": "pivaad", "pct_ganancias": "pgcia", "pct_iibb": "piibb",
    "margen_pct": "pmargen", "pv_final_usd": "pvfin",
}
GASTO_FIELD_PREFIX = {
    "concepto": "gcon", "moneda": "gmon", "monto": "gmonto",
    "prorrateo": "gpror", "iva_incluido": "gincl", "pct_iva": "gpctiva",
}
# Estos 4 gastos por defecto ya no se cargan en Costos operativos: se
# auto-sincronizan desde los campos de la sección "Tarifas flete" (Tarifa
# flete, Gastos en origen, Gastos locales, Seguro). Siguen existiendo como
# filas de `cot_gastos` (mismo cálculo de prorrateo/IVA que siempre
# tuvieron, sin duplicar nada), solo que ocultas de esa sección.
GASTOS_EN_TARIFAS_FLETE = {"Flete", "Seguro", "Gastos EXW", "Gastos locales"}


def _producto_default():
    return {
        "descripcion": "", "ncm": "", "fob_unit": 0.0, "cantidad": 0.0,
        "fob_decl_unit": 0.0, "peso_kg": 0.0, "volumen_m3": 0.0, "pct_derechos": 0.0,
        "pct_tasa_estadistica": 0.03, "pct_antidumping": 0.0, "pct_iva": 0.21,
        "pct_iva_adicional": 0.0, "pct_ganancias": 0.0, "pct_iibb": 0.05,
        "margen_pct": 0.3, "pv_final_usd": 0.0,
    }


def _gasto_default(concepto="", moneda="USD", prorrateo="FOB", iva_incluido="NO", pct_iva=0.21):
    return {
        "concepto": concepto, "moneda": moneda, "monto": 0.0,
        "prorrateo": prorrateo, "iva_incluido": iva_incluido, "pct_iva": pct_iva,
    }


def _con_uid(row):
    """Asigna un identificador estable por fila, usado como sufijo de las keys
    de los widgets (el id de la BD si ya existe, o uno nuevo)."""
    row = dict(row)
    if not row.get("_uid"):
        row["_uid"] = str(row["id"]) if row.get("id") else uuid.uuid4().hex[:8]
    return row


def _nueva_fila_producto():
    return _con_uid(_fila_a_puntos(_producto_default(), PCT_PROD_COLS))


def _nueva_fila_gasto(**kwargs):
    return _con_uid(_fila_a_puntos(_gasto_default(**kwargs), PCT_GASTO_COLS))


def _cargar_estado_cotizacion(cot_id):
    cab = db.get_cotizacion(cot_id)
    productos = db.get_productos(cot_id)
    gastos = db.get_gastos(cot_id)
    st.session_state.cot_cab = cab
    st.session_state.cot_productos = [_con_uid(_fila_a_puntos(p, PCT_PROD_COLS)) for p in productos]
    st.session_state.cot_gastos = [_con_uid(_fila_a_puntos(g, PCT_GASTO_COLS)) for g in gastos]


def _abrir_cotizacion_en_cotizador(cid):
    # Callback: corre antes del próximo rerun, así que es seguro fijar acá
    # tanto la cotización activa como la página a mostrar (se usa desde
    # Historial y desde el Panel de Control). También fuerza el modo
    # "editor" — si no, al venir desde el listado (que ahora es la vista por
    # default del Cotizador) la vista se quedaba en el listado sin mostrar
    # ningún cambio, aunque la cotización sí se hubiera cargado.
    _cargar_estado_cotizacion(cid)
    st.session_state.pagina_nav = "🧮 Cotizador"
    st.session_state.cotizador_modo = "editor"


def _cotizar_cliente(cliente_id):
    # Callback (mismo motivo que la anterior: toca "pagina_nav", que ya es
    # un widget instanciado en este run por el sidebar). Crea una cotización
    # nueva ya vinculada al cliente y abre el Cotizador directo en ella.
    new_id = db.create_cotizacion_borrador(cliente_id=cliente_id)
    # Mismo patrón que _cotizar_contacto en sentido inverso (cliente →
    # contacto en vez de contacto → cliente): asegura el vínculo de CRM y
    # sube a "Cotizado" si corresponde — sin esto, cotizar directo desde
    # Clientes (sin pasar por un contacto de CRM) dejaba el embudo sin
    # actualizar. _avanzar_etapa_si_corresponde nunca retrocede, así que no
    # hay riesgo de pisar una etapa más avanzada (ej. Ganado).
    contacto = _asegurar_contacto_para_cliente(cliente_id)
    if contacto:
        _avanzar_etapa_si_corresponde(contacto["id"], "Cotizado", st.session_state.get("crm_autor", ""))
    _cargar_estado_cotizacion(new_id)
    st.session_state.pagina_nav = "🧮 Cotizador"
    st.session_state.cotizador_modo = "editor"


def _sync_desde_widgets(filas, prefijos):
    """Actualiza cada dict con el valor MÁS RECIENTE de sus widgets, aunque esos
    widgets todavía no se hayan renderizado en esta pasada del script (su
    session_state ya existe desde la pasada anterior). Permite mostrar
    secciones calculadas (CIF, base imponible) antes de las secciones de
    carga de las que dependen, sin quedar "un paso atrás"."""
    for fila in filas:
        uid = fila["_uid"]
        for campo, prefijo in prefijos.items():
            key = f"{prefijo}_{uid}"
            if key in st.session_state:
                fila[campo] = st.session_state[key]


def _fila_campos_producto(p, campos):
    """Fila compacta para un producto con N campos numéricos en columnas cortas.
    campos: lista de (nombre_campo, etiqueta)."""
    uid = p["_uid"]
    with st.container(border=True, key=f"formrow_campos_{campos[0][0]}_{uid}"):
        cols = st.columns([2] + [1] * len(campos))
        cols[0].markdown(f"**{p.get('descripcion') or '(sin nombre)'}**")
        for col, (campo, etiqueta) in zip(cols[1:], campos):
            prefijo = PROD_FIELD_PREFIX[campo]
            p[campo] = col.number_input(etiqueta, value=_f_local(p.get(campo)), format="%.2f", key=f"{prefijo}_{uid}")


def _mercaderia_tiene_ceros(productos):
    """True si algún producto tiene FOB Unit. o Cantidad en 0 — señal de que
    todavía falta cargar su precio, no de que el producto vale 0 a propósito."""
    return any(_f_local(p.get("fob_unit")) == 0 or _f_local(p.get("cantidad")) == 0 for p in productos)


def _render_mercaderia():
    """Sección 2: datos básicos de cada producto (sin impuestos ni margen).
    Agregar un producto es opcional: la cotización arranca vacía."""
    productos = st.session_state.cot_productos

    if not productos:
        st.info("Todavía no cargaste productos. Usá '➕ Agregar producto' cuando lo necesites.")
    else:
        resumen = pd.DataFrame([{
            "#": i + 1,
            "Descripción": ("⚠️ " if (_f_local(p.get("fob_unit")) == 0 or _f_local(p.get("cantidad")) == 0) else "")
                + (p.get("descripcion") or "(sin nombre)"),
            "FOB Unit.": money(_f_local(p.get("fob_unit"))),
            "Cantidad": _f_local(p.get("cantidad")),
            "FOB Total": money(_f_local(p.get("fob_unit")) * _f_local(p.get("cantidad"))),
        } for i, p in enumerate(productos)])
        st.dataframe(resumen, hide_index=True, use_container_width=True)

    borrar_idx = None
    for i, p in enumerate(productos):
        uid = p["_uid"]
        with st.container(border=True, key=f"formrow_mercaderia_{uid}"):
            c0, c1, c2, c3 = st.columns([3, 2, 1.3, 1.3])
            p["descripcion"] = c0.text_input("Descripción", value=p.get("descripcion", ""), key=f"pdesc_{uid}")
            p["ncm"] = c1.text_input("NCM", value=p.get("ncm", ""), key=f"pncm_{uid}")
            p["fob_unit"] = c2.number_input("FOB Unit. (USD)", value=_f_local(p.get("fob_unit")), format="%.2f", key=f"pfob_{uid}")
            p["cantidad"] = c3.number_input("Cantidad", value=_f_local(p.get("cantidad")), format="%.2f", key=f"pcant_{uid}")
            if p["fob_unit"] == 0 or p["cantidad"] == 0:
                st.caption("⚠️ Falta cargar FOB Unit. y/o Cantidad — este producto no suma costo todavía.")

            c4, c5, c6, c7 = st.columns([1.6, 1, 1, 1.1])
            p["fob_decl_unit"] = c4.number_input("FOB Declarado Unit. (USD)", value=_f_local(p.get("fob_decl_unit")), format="%.2f", key=f"pfobd_{uid}", help="Si es distinto al FOB real de compra")
            p["peso_kg"] = c5.number_input("Peso (kg)", value=_f_local(p.get("peso_kg")), format="%.2f", key=f"ppesokg_{uid}")
            p["volumen_m3"] = c6.number_input("Volumen (m³)", value=_f_local(p.get("volumen_m3")), format="%.2f", key=f"pvolm3_{uid}")
            c7.write("")
            c7.write("")
            if c7.button("🗑️ Eliminar", key=f"pdel_{uid}", use_container_width=True):
                borrar_idx = i

    if borrar_idx is not None:
        productos.pop(borrar_idx)
        st.rerun()

    if st.button("➕ Agregar producto"):
        productos.append(_nueva_fila_producto())
        st.rerun()


def _render_cif(resultado):
    """Cómo se arma el CIF declarado (solo lectura)."""
    st.caption("FOB declarado + flete y seguro declarados. Sobre este valor se calculan los derechos.")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("FOB Declarado", money(resultado["fob_declarado"]))
    c2.metric("Flete Declarado", money(resultado["flete_declarado"]))
    c3.metric("Seguro Declarado", money(resultado["seguro_declarado"]))
    c4.metric("CIF Declarado", money(resultado["cif_declarado"]))


def _render_derechos():
    """% de derechos, tasa estadística y antidumping por producto."""
    productos = st.session_state.cot_productos
    if not productos:
        st.caption("Cargá productos en la sección 2 para definir sus derechos.")
        return
    st.caption("Dependen del NCM de cada producto — verificá en el nomenclador vigente.")
    for p in productos:
        _fila_campos_producto(p, [
            ("pct_derechos", "Derechos (%)"),
            ("pct_tasa_estadistica", "Tasa Estad. (%)"),
            ("pct_antidumping", "Antidump. (%)"),
        ])


def _render_base_imponible(resultado):
    """Base imponible de IVA por producto (solo lectura)."""
    productos = resultado["productos"]
    if not productos:
        st.caption("No hay productos cargados.")
        return
    st.caption("CIF de cada producto + sus derechos, tasa estadística y antidumping = base sobre la que se calculan IVA, IVA Adicional, Ganancias e IIBB.")
    rows = [{
        "Producto": p.get("descripcion") or "-", "CIF Producto": money(p["cif_producto"]),
        "Derechos+Tasas+Antidump.": money(p["subtotal_der_tasas"]), "Base IVA": money(p["base_iva_prod"]),
    } for p in productos]
    st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
    st.metric("Base IVA Total", money(resultado["base_iva_total"]))


def _render_impuestos():
    """% de IVA, IVA Adicional, Ganancias e IIBB por producto."""
    productos = st.session_state.cot_productos
    if not productos:
        st.caption("Cargá productos en la sección 2 para definir sus impuestos.")
        return
    for p in productos:
        _fila_campos_producto(p, [
            ("pct_iva", "IVA (%)"),
            ("pct_iva_adicional", "IVA Adic. (%)"),
            ("pct_ganancias", "Ganancias (%)"),
            ("pct_iibb", "IIBB (%)"),
        ])


def _gastos_visibles_tiene_ceros(gastos):
    """True si algún gasto operativo (fuera de Tarifas flete) quedó en 0 —
    señal de que falta cargar su monto, no de que sea gratis a propósito."""
    return any(
        _f_local(g.get("monto")) == 0
        for g in gastos if g.get("concepto") not in GASTOS_EN_TARIFAS_FLETE
    )


def _render_gastos():
    """Sección: costos operativos restantes (Flete, Seguro, Gastos en origen y
    Gastos locales ya no aparecen acá — se cargan en 'Tarifas flete', más
    arriba, para no duplicar la carga; siguen sumando a los totales igual
    que antes, solo que ahora se auto-sincronizan desde esa sección)."""
    st.caption("Honorarios, terminal portuaria, depósito fiscal y demás gastos del despacho. 'Prorrateo' define si se reparte entre productos por valor FOB o por peso/volumen.")

    gastos = st.session_state.cot_gastos
    visibles = [(i, g) for i, g in enumerate(gastos) if g.get("concepto") not in GASTOS_EN_TARIFAS_FLETE]

    if visibles:
        resumen = pd.DataFrame([{
            "#": n + 1,
            "Concepto": ("⚠️ " if _f_local(g.get("monto")) == 0 else "") + (g.get("concepto") or "(sin nombre)"),
            "Monto": money(_f_local(g.get("monto")), g.get("moneda") or "USD"),
            "Prorrateo": g.get("prorrateo"),
        } for n, (i, g) in enumerate(visibles)])
        st.dataframe(resumen, hide_index=True, use_container_width=True)

    borrar_idx = None
    for n, (i, g) in enumerate(visibles):
        uid = g["_uid"]
        titulo = g.get("concepto") or f"Gasto {n + 1} (sin nombre)"
        if _f_local(g.get("monto")) == 0:
            titulo = f"⚠️ {titulo} — sin monto cargado"
        with st.expander(f"{n + 1}. {titulo}", expanded=False):
            c1, c2, c3 = st.columns(3)
            g["concepto"] = c1.text_input("Concepto", value=g.get("concepto", ""), key=f"gcon_{uid}")
            g["moneda"] = c2.selectbox("Moneda", MONEDAS, index=MONEDAS.index(g.get("moneda") or "USD"), key=f"gmon_{uid}")
            g["monto"] = c3.number_input("Monto (s/IVA propio)", value=_f_local(g.get("monto")), format="%.2f", key=f"gmonto_{uid}")

            c4, c5, c6 = st.columns(3)
            g["prorrateo"] = c4.selectbox("Prorrateo", PRORRATEOS, index=PRORRATEOS.index(g.get("prorrateo") or "FOB"), key=f"gpror_{uid}", help="FOB: se reparte según el valor de cada producto. PESO: según su peso/volumen.")
            g["iva_incluido"] = c5.selectbox("IVA incluido en el monto", SI_NO, index=SI_NO.index(g.get("iva_incluido") or "NO"), key=f"gincl_{uid}")
            g["pct_iva"] = c6.number_input("IVA (%)", value=_f_local(g.get("pct_iva")), format="%.2f", key=f"gpctiva_{uid}", help="% de IVA que ya viene incluido en este gasto — se usa como crédito fiscal recuperable.")

            if st.button("🗑️ Eliminar gasto", key=f"gdel_{uid}"):
                borrar_idx = i

    if borrar_idx is not None:
        gastos.pop(borrar_idx)
        st.rerun()

    if st.button("➕ Agregar gasto"):
        gastos.append(_nueva_fila_gasto())
        st.rerun()


def _render_simulacion_venta(resultado):
    """Margen sugerido y PV final por producto, con el resultado de venta."""
    productos = st.session_state.cot_productos
    if not productos:
        st.caption("No hay productos cargados.")
        return
    for p, r in zip(productos, resultado["productos"]):
        uid = p["_uid"]
        with st.container(border=True, key=f"formrow_venta_{uid}"):
            c0, c1, c2, c3 = st.columns([2, 1, 1, 1])
            c0.markdown(f"**{p.get('descripcion') or '(sin nombre)'}**")
            c0.caption(f"Costo c/IVA unit.: {money(r['costo_civa_unit'])}")
            p["margen_pct"] = c1.number_input("Margen (%)", value=_f_local(p.get("margen_pct")), format="%.2f", key=f"pmargen_{uid}", help="Margen deseado sobre el costo, para calcular el PV sugerido")
            p["pv_final_usd"] = c2.number_input("PV Final (USD)", value=_f_local(p.get("pv_final_usd")), format="%.2f", key=f"pvfin_{uid}", help="Precio de venta real — dejalo en 0 si todavía no lo definiste")
            c3.metric("Margen real", pct(r["margen_real_pct"]))
            st.caption(f"PV sugerido: {money(r['pv_sugerido_usd'])} · Ganancia unit.: {money(r['ganancia_unit_usd'])} · Ganancia total: {money(r['ganancia_total_usd'])}")

            st.markdown("**Simulación (ARS)**")
            a0, a1, a2, a3 = st.columns([2, 1, 1, 1])
            a0.caption(f"Costo c/IVA unit.: {money(r['costo_civa_unit_ars'], 'ARS')}")
            a1.metric("PV sugerido", money(r["pv_sugerido_ars"], "ARS"))
            a2.metric("PV Final", money(r["pv_final_ars"], "ARS"))
            a3.metric("Margen real", pct(r["margen_real_pct"]))
            st.caption(f"Ganancia unit.: {money(r['ganancia_unit_ars'], 'ARS')} · Ganancia total: {money(r['ganancia_total_ars'], 'ARS')}")


def _volver_a_historial():
    st.session_state.cotizador_modo = "historial"


def _render_cotizador_editor():
    clientes = db.list_clientes()
    mapa_clientes = {c["id"]: c["nombre"] for c in clientes}

    # El editor ya no elige QUÉ cotización abrir (eso se decide desde el
    # listado, con "➕ Nueva cotización" o "✏️ Editar" en una fila) — acá solo
    # se edita la que ya quedó cargada en session_state.cot_cab.
    st.button(
        "Volver al listado", icon=":material/arrow_back:",
        key="volver_ficha_historial", on_click=_volver_a_historial,
    )

    cab = st.session_state.cot_cab
    cot_id = cab["id"]

    st.subheader(f"Cotización {cab['numero']}")

    # ============================================================
    # 1️⃣ Datos generales de la cotización
    # ============================================================
    with st.expander("1️⃣ Datos generales de la cotización", expanded=False):
        # Las keys incluyen cot_id para que, al cambiar de cotización, Streamlit
        # trate cada widget como uno nuevo y tome los valores recién cargados en
        # lugar de conservar el valor tipeado para la cotización anterior.
        with st.form(f"form_cabecera_{cot_id}"):
            c1, c2, c3 = st.columns(3)
            cliente_id = c1.selectbox(
                "Cliente", options=[None] + list(mapa_clientes.keys()),
                format_func=lambda x: "(sin cliente)" if x is None else mapa_clientes[x],
                index=(list(mapa_clientes.keys()).index(cab["cliente_id"]) + 1) if cab.get("cliente_id") in mapa_clientes else 0,
                key=f"cliente_id_{cot_id}",
            )
            estado = c2.selectbox("Estado", ESTADOS, index=ESTADOS.index(cab.get("estado") or "Borrador"), key=f"estado_{cot_id}")
            detalle_pedido = c3.text_input("Detalle pedido", value=cab.get("detalle_pedido") or "", key=f"detalle_{cot_id}")

            c4, c5, c6 = st.columns(3)
            origen = c4.text_input("Origen", value=cab.get("origen_cond_venta") or "", key=f"origen_{cot_id}")
            contenedor = c5.selectbox(
                "Condición de venta", COT_CONDICIONES_VENTA,
                index=COT_CONDICIONES_VENTA.index(cab.get("contenedor")) if cab.get("contenedor") in COT_CONDICIONES_VENTA else 0,
                key=f"contenedor_{cot_id}",
            )
            etd_eta = c6.selectbox(
                "Tipo de envío", COT_TIPOS_ENVIO,
                index=COT_TIPOS_ENVIO.index(cab.get("etd_eta")) if cab.get("etd_eta") in COT_TIPOS_ENVIO else 0,
                key=f"etdeta_{cot_id}",
            )

            c7, c8, c9 = st.columns(3)
            etd_date = c7.date_input("ETD", value=_parse_fecha(cab.get("carrier")), key=f"carrier_{cot_id}", format="DD/MM/YYYY")
            eta_date = c8.date_input("ETA", value=_parse_fecha(cab.get("freetime")), key=f"freetime_{cot_id}", format="DD/MM/YYYY")
            carrier = etd_date.isoformat() if etd_date else None
            freetime = eta_date.isoformat() if eta_date else None
            tc_tributos = c9.number_input("TC Tributos (despacho)", value=float(cab.get("tc_tributos") or 0), step=1.0, key=f"tctrib_{cot_id}")

            c10, c11, c12 = st.columns(3)
            cf_pct_input = c10.number_input(
                "Costo financiero (%) s/FOB", value=float(cab.get("costo_financiero_pct") or 0.025) * 100,
                format="%.2f", step=0.1, key=f"cfpct_{cot_id}", help="Ej: escribí 2.5 para 2,5%",
            )
            seguro_pct_input = c11.number_input(
                "Seguro (%) s/FOB", value=float(cab.get("seguro_pct") or 0.003) * 100,
                format="%.2f", step=0.05, key=f"segpct_{cot_id}", help="Ej: escribí 0.3 para 0,3%",
            )
            cf_pct = cf_pct_input / 100
            seguro_pct = seguro_pct_input / 100
            tc_operativos = c12.number_input("TC Operativos", value=float(cab.get("tc_operativos") or 0), step=1.0, key=f"tcoper_{cot_id}")

            st.form_submit_button("💾 Guardar datos generales")

    # El Arancel SIM se carga en la sección 6 (junto a IVA/Ganancias/IIBB, como
    # en el Excel), pero necesitamos su valor actual ya acá para poder
    # calcular una sola vez. Como todavía no se renderizó ese widget en esta
    # pasada, se lee directamente de session_state (ver _sync_desde_widgets).
    arancel_key = f"arancel_{cot_id}"
    arancel_sim = st.session_state.get(arancel_key, float(cab.get("arancel_sim") or 10))

    # Tarifas flete (sección 3): mismo patrón de prefetch que el Arancel SIM
    # — se lee el último valor tipeado desde session_state antes de calcular,
    # aunque esos widgets recién se dibujen más abajo en la página.
    tarifaflete_key = f"tarifaflete_{cot_id}"
    pctcert_key = f"pctcert_{cot_id}"
    gastosorigen_key = f"gastosorigen_{cot_id}"
    gastoslocaleshdr_key = f"gastoslocaleshdr_{cot_id}"
    tarifa_flete = st.session_state.get(tarifaflete_key, float(cab.get("tarifa_flete") or 0))
    pct_certificacion_input = st.session_state.get(pctcert_key, float(cab.get("pct_certificacion") or 0.5) * 100)
    gastos_origen = st.session_state.get(gastosorigen_key, float(cab.get("gastos_origen") or 0))
    gastos_locales_hdr = st.session_state.get(gastoslocaleshdr_key, float(cab.get("gastos_locales_hdr") or 0))
    pct_certificacion = pct_certificacion_input / 100

    # % IVA de Gastos locales y Seguro: son los únicos 2 de los 4 gastos
    # relocados que sí llevan IVA (Flete y Gastos en origen no pagan, no
    # hace falta editarlo) — se editan también acá para no reabrir la
    # sección "IVA" separada que se eliminó.
    _gasto_seguro = next((g for g in st.session_state.cot_gastos if g.get("concepto") == "Seguro"), None)
    _gasto_gastoslocales = next((g for g in st.session_state.cot_gastos if g.get("concepto") == "Gastos locales"), None)
    pctivaseguro_key = f"pctivaseguro_{cot_id}"
    pctivagastoslocales_key = f"pctivagastoslocales_{cot_id}"
    pct_iva_seguro_input = st.session_state.get(
        pctivaseguro_key, _f_local(_gasto_seguro.get("pct_iva")) if _gasto_seguro else 21.0)
    pct_iva_gastoslocales_input = st.session_state.get(
        pctivagastoslocales_key, _f_local(_gasto_gastoslocales.get("pct_iva")) if _gasto_gastoslocales else 21.0)

    # Sincronizamos productos y gastos con el último valor de sus widgets
    # (aunque esas secciones se rendericen más abajo) para poder calcular UNA
    # sola vez y mostrar secciones dependientes (CIF, base imponible) antes
    # en la página, igual que en el Excel.
    _sync_desde_widgets(st.session_state.cot_productos, PROD_FIELD_PREFIX)
    _sync_desde_widgets(st.session_state.cot_gastos, GASTO_FIELD_PREFIX)

    cab_calc = {
        "costo_financiero_pct": cf_pct, "seguro_pct": seguro_pct,
        "tc_tributos": tc_tributos, "tc_operativos": tc_operativos, "arancel_sim": arancel_sim,
        "tarifa_flete": tarifa_flete, "pct_certificacion": pct_certificacion,
    }
    # Las secciones se editan en puntos porcentuales (21 = 21%); acá se
    # convierten a fracción (0.21) antes de calcular y guardar.
    productos_list = _a_fraccion(st.session_state.cot_productos, PCT_PROD_COLS)
    gastos_list = _a_fraccion(st.session_state.cot_gastos, PCT_GASTO_COLS)
    resultado_previo = calculo.calcular(cab_calc, productos_list, gastos_list)

    # Flete, Gastos en origen, Gastos locales y Seguro se cargan una sola vez
    # en "Tarifas flete" (sección 3) y fluyen solos hacia su renglón interno
    # de Costos operativos — no hace falta tipear el mismo monto dos veces.
    # Sigue siendo EXACTAMENTE el mismo cálculo de siempre (prorrateo, IVA,
    # etc.), solo que la carga ya no pasa por esa sección.
    _campos_sincronizados = {
        "Flete": {"monto": tarifa_flete},
        "Gastos EXW": {"monto": gastos_origen},
        "Gastos locales": {"monto": gastos_locales_hdr, "pct_iva": pct_iva_gastoslocales_input},
        "Seguro": {"monto": resultado_previo["seguro_declarado"], "pct_iva": pct_iva_seguro_input},
    }
    for _g in st.session_state.cot_gastos:
        _cambios = _campos_sincronizados.get(_g.get("concepto"))
        if not _cambios:
            continue
        for _campo, _valor_nuevo in _cambios.items():
            _g[_campo] = _valor_nuevo
            # Solo pisamos session_state si el widget ya existe de un run
            # anterior — en el primer render alcanza con el valor de arriba
            # (fluye al widget vía su parámetro value=) y evita el warning de
            # Streamlit por setear session_state y value= a la vez.
            _gasto_key = f"{GASTO_FIELD_PREFIX[_campo]}_{_g['_uid']}"
            if _gasto_key in st.session_state:
                st.session_state[_gasto_key] = _valor_nuevo
    gastos_list = _a_fraccion(st.session_state.cot_gastos, PCT_GASTO_COLS)
    resultado = calculo.calcular(cab_calc, productos_list, gastos_list)

    # ============================================================
    # 2️⃣ Detalle de mercadería y precios
    # ============================================================
    titulo_mercaderia = "2️⃣ Detalle de mercadería y precios"
    if _mercaderia_tiene_ceros(st.session_state.cot_productos):
        titulo_mercaderia += " ⚠️ hay ítems sin cargar"
    with st.expander(titulo_mercaderia, expanded=False):
        _render_mercaderia()

    # ============================================================
    # 3️⃣ Tarifas flete
    # ============================================================
    titulo_tarifas = "3️⃣ Tarifas flete"
    if tarifa_flete == 0 or gastos_origen == 0 or gastos_locales_hdr == 0:
        titulo_tarifas += " ⚠️ hay ítems sin cargar"
    with st.expander(titulo_tarifas, expanded=False):
        tf1, tf2, tf3 = st.columns(3)
        tarifa_flete = tf1.number_input(
            "Tarifa flete (USD)", value=tarifa_flete, format="%.2f", step=10.0, key=tarifaflete_key,
            help="Costo total del flete pagado al forwarder/naviera.",
        )
        if tarifa_flete == 0:
            tf1.caption("⚠️ Sin cargar")
        pct_certificacion_input = tf2.number_input(
            "% Certificación", value=pct_certificacion_input, format="%.2f", min_value=0.0, max_value=100.0,
            step=5.0, key=pctcert_key, help="Porción del flete certificada por la naviera para declarar en el CIF.",
        )
        pct_certificacion = pct_certificacion_input / 100
        # Los campos disabled=True no refrescan su value= en reruns posteriores
        # (Streamlit los sigue leyendo de session_state, aunque el usuario nunca
        # los toque) — se fuerza acá, mismo patrón que el auto-sync del Seguro.
        tarifafletecert_key = f"tarifafletecert_{cot_id}"
        st.session_state[tarifafletecert_key] = tarifa_flete * pct_certificacion
        tf3.number_input(
            "Tarifa flete certificado (USD)", value=tarifa_flete * pct_certificacion, format="%.2f",
            disabled=True, key=tarifafletecert_key, help="= Tarifa flete × % Certificación. Se usa en el CIF.",
        )

        tf4, tf5, tf6 = st.columns(3)
        gastos_origen = tf4.number_input(
            "Gastos en origen (USD)", value=gastos_origen, format="%.2f", step=10.0, key=gastosorigen_key,
            help="Gastos EXW / en el país de origen (handling, documentación, etc.).",
        )
        if gastos_origen == 0:
            tf4.caption("⚠️ Sin cargar")
        gastos_locales_hdr = tf5.number_input(
            "Gastos locales (USD)", value=gastos_locales_hdr, format="%.2f", step=10.0, key=gastoslocaleshdr_key,
            help="Gastos locales en Argentina asociados al despacho.",
        )
        if gastos_locales_hdr == 0:
            tf5.caption("⚠️ Sin cargar")
        segurohdr_key = f"segurohdr_{cot_id}"
        st.session_state[segurohdr_key] = resultado["seguro_declarado"]
        tf6.number_input(
            "Seguro (USD)", value=resultado["seguro_declarado"], format="%.2f", disabled=True,
            key=segurohdr_key, help="= FOB total × Seguro (%) s/FOB, cargado en Datos generales.",
        )

        # Gastos en origen y Flete no pagan IVA, no hace falta editarlo — Gastos
        # locales y Seguro sí, así que su % IVA (crédito fiscal recuperable) se
        # edita acá en vez de en Costos operativos, donde ya no aparecen.
        _, tf8, tf9 = st.columns(3)
        pct_iva_gastoslocales_input = tf8.number_input(
            "IVA Gastos locales (%)", value=pct_iva_gastoslocales_input, format="%.2f", step=1.0,
            key=pctivagastoslocales_key, help="% de IVA incluido en Gastos locales — crédito fiscal recuperable.",
        )
        pct_iva_seguro_input = tf9.number_input(
            "IVA Seguro (%)", value=pct_iva_seguro_input, format="%.2f", step=1.0,
            key=pctivaseguro_key, help="% de IVA incluido en el Seguro — crédito fiscal recuperable.",
        )

    # ============================================================
    # 4️⃣ Valor CIF
    # ============================================================
    with st.expander("4️⃣ Valor CIF", expanded=False):
        _render_cif(resultado)

    # ============================================================
    # 5️⃣ Derechos, tasa estadística y antidumping
    # ============================================================
    with st.expander("5️⃣ Derechos, tasa estadística y antidumping (%)", expanded=False):
        _render_derechos()

    # ============================================================
    # 6️⃣ Base imponible para IVA
    # ============================================================
    with st.expander("6️⃣ Base imponible para el cálculo de IVA", expanded=False):
        _render_base_imponible(resultado)

    # ============================================================
    # 7️⃣ IVA, IVA Adicional, Ganancias, IIBB + Arancel SIM
    # ============================================================
    with st.expander("7️⃣ IVA, IVA Adicional, Ganancias e IIBB (%) + Arancel SIM", expanded=False):
        arancel_sim = st.number_input(
            "Arancel SIM (USD)", value=arancel_sim, step=1.0, key=arancel_key,
            help="Costo fijo del Sistema Informático María (trámite aduanero), no depende del producto.",
        )
        _render_impuestos()

    # ============================================================
    # 8️⃣ Costos operativos
    # ============================================================
    titulo_costos_op = "8️⃣ Costos operativos"
    if _gastos_visibles_tiene_ceros(st.session_state.cot_gastos):
        titulo_costos_op += " ⚠️ hay ítems sin cargar"
    with st.expander(titulo_costos_op, expanded=False):
        _render_gastos()

    # ============================================================
    # 9️⃣ Resultado de la operación (costo puro — sin venta ni margen,
    # eso vive en "Simulación de venta")
    # ============================================================
    with st.expander("9️⃣ Resultado de la operación (costos)", expanded=False):
        with st.container(border=True, key=f"resultgroup_desembolso_{cot_id}"):
            st.markdown('<div class="sb-card-title">💰 Desembolso total</div>', unsafe_allow_html=True)
            d1, d2 = st.columns(2)
            d1.metric("TOTAL desembolsado (c/IVA)", money(resultado["total_c_iva_usd"]))
            d2.metric("Total (ARS)", money(resultado["total_c_iva_ars"], "ARS"))

        with st.container(border=True, key=f"resultgroup_composicion_{cot_id}"):
            st.markdown('<div class="sb-card-title">🧮 Composición del costo</div>', unsafe_allow_html=True)
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("① Mercadería + fin.", money(resultado["fob_total_sum"] + resultado["costo_financiero"]))
            c2.metric("② Tributos VEP", money(resultado["total_tributos_vep_usd"]))
            c3.metric("③ Gastos oper. + IVA", money(resultado["subtotal_oper_usd"] + resultado["subtotal_ivater_usd"]))
            c4.metric("Incidencia s/FOB", pct(resultado["incidencia_fob"]))

        with st.container(border=True, key=f"resultgroup_recuperable_{cot_id}"):
            st.markdown('<div class="sb-card-title">♻️ Créditos fiscales recuperables</div>', unsafe_allow_html=True)
            e1, e2 = st.columns(2)
            e1.metric("IVA a recuperar", money(resultado["iva_a_recuperar"]))
            e2.metric("Percepciones a recuperar", money(resultado["percep_a_recuperar"]))

        with st.expander("Ver costeo unitario por producto"):
            rows = [{
                "Producto": p.get("descripcion") or "-",
                "Costo c/IVA Unit. (USD)": money(p["costo_civa_unit"]),
                "Costo c/IVA Unit. (ARS)": money(p["costo_civa_unit_ars"], "ARS"),
                "Costo s/IVA Unit. (USD)": money(p["costo_sviva_unit"]),
                "Costo s/IVA Unit. (ARS)": money(p["costo_sviva_unit_ars"], "ARS"),
            } for p in resultado["productos"]]
            st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
            check = resultado["check_diferencia"]
            if abs(check) < 0.5:
                st.success(f"✅ Check de consistencia OK (diferencia: {check:,.4f})")
            else:
                st.warning(f"⚠️ Diferencia de redondeo en el check: {money(check)}")

    # ============================================================
    # 🔟 Simulación de venta (todo lo que es venta, margen y ganancia
    # estimada vive acá, no en "Resultado de la operación")
    # ============================================================
    with st.expander("🔟 Simulación de venta", expanded=False):
        with st.container(border=True, key=f"resultgroup_venta_{cot_id}"):
            st.markdown('<div class="sb-card-title">📈 Resultado estimado de la venta</div>', unsafe_allow_html=True)
            v1, v2, v3 = st.columns(3)
            v1.metric("Venta total estimada", money(resultado["venta_total_usd"]))
            v2.metric("Venta total estimada (ARS)", money(resultado["venta_total_ars"], "ARS"))
            v3.metric("% Rentabilidad s/inversión", pct(resultado["rentabilidad_pct"]))
            v4, v5 = st.columns(2)
            v4.metric("Ganancia bruta", money(resultado["ganancia_bruta_usd"]))
            v5.metric("Ganancia bruta (ARS)", money(resultado["ganancia_bruta_ars"], "ARS"))

        _render_simulacion_venta(resultado)

    def _guardar_cotizacion():
        db.save_cotizacion(
            cot_id,
            {
                "cliente_id": cliente_id, "detalle_pedido": detalle_pedido,
                "origen_cond_venta": origen, "contenedor": contenedor, "etd_eta": etd_eta,
                "carrier": carrier, "freetime": freetime, "estado": estado,
                "costo_financiero_pct": cf_pct, "seguro_pct": seguro_pct,
                "tc_tributos": tc_tributos, "tc_operativos": tc_operativos, "arancel_sim": arancel_sim,
                "tarifa_flete": tarifa_flete, "pct_certificacion": pct_certificacion,
                "gastos_origen": gastos_origen, "gastos_locales_hdr": gastos_locales_hdr,
            },
            productos_list, gastos_list,
            {
                "total_c_iva_usd": resultado["total_c_iva_usd"], "total_c_iva_ars": resultado["total_c_iva_ars"],
                "total_s_iva_usd": resultado["total_s_iva_usd"], "total_s_iva_ars": resultado["total_s_iva_ars"],
                "venta_total_usd": resultado["venta_total_usd"], "ganancia_bruta_usd": resultado["ganancia_bruta_usd"],
                "rentabilidad_pct": resultado["rentabilidad_pct"],
            },
        )
        mensaje = "Cotización guardada."
        # Si la cotización quedó Aprobada, ese cliente "gana" el trato: pasa
        # a Ganado y deja de verse en el tablero de Contactos (se gestiona
        # desde Clientes). _marcar_ganado_por_cliente asegura el vínculo
        # cliente↔contacto si todavía no existía, en vez de no-opear en
        # silencio (que era el comportamiento antes de esto).
        if estado == "Aprobada" and cliente_id:
            contacto_vinculado = _marcar_ganado_por_cliente(cliente_id)
            if contacto_vinculado:
                mensaje += f" '{contacto_vinculado['nombre']}' pasó a Cliente."
        _cargar_estado_cotizacion(cot_id)
        _flash(mensaje)

    st.markdown("---")
    colA, colB, colC, colD = st.columns(4)
    colA.button("💾 Guardar cotización", type="primary", use_container_width=True, on_click=_guardar_cotizacion)

    cab_full = dict(cab)
    cab_full.update(cliente_nombre=mapa_clientes.get(cliente_id, ""), detalle_pedido=detalle_pedido,
                     origen_cond_venta=origen, contenedor=contenedor, etd_eta=etd_eta, carrier=carrier,
                     freetime=freetime)
    # Simplificado: solo página 1 (Resumen + Detalle de mercadería), para
    # mandarle al cliente sin exponer el desglose línea por línea de gastos.
    # Completo: página 1 + 2 (con el desglose), para uso interno.
    pdf_simple = generar_pdf_cotizacion(cab_full, resultado, completo=False)
    pdf_completo = generar_pdf_cotizacion(cab_full, resultado, completo=True)

    nombre_pdf = _nombre_pdf_cotizacion(cab_full, resultado)

    colB.download_button(
        "📄 PDF simplificado", data=pdf_simple, file_name=f"{nombre_pdf} - Simplificado.pdf",
        mime="application/pdf", use_container_width=True,
    )
    colC.download_button(
        "📑 PDF completo", data=pdf_completo, file_name=f"{nombre_pdf} - Completo.pdf",
        mime="application/pdf", use_container_width=True,
    )
    if colD.button("🗑️ Eliminar cotización", use_container_width=True):
        _dialog_eliminar_cotizacion(cot_id, cab["numero"])


# ---------------------------------------------------------------- PANEL DE CONTROL
# Tope de tarjetas dibujadas por columna: Streamlit crea un widget real por
# cada card (container + botón), así que con cientos de registros por estado
# el árbol de widgets se vuelve enorme y lento. Se muestran las más recientes
# y se pide afinar la búsqueda para ver el resto, en vez de renderizarlas todas.
LIMITE_TARJETAS_POR_COLUMNA = 8


def _render_panel_cotizaciones():
    cots_todas = db.list_cotizaciones()
    if not cots_todas:
        st.info("Todavía no hay cotizaciones cargadas.")
        return

    aprobadas = [c for c in cots_todas if (c.get("estado") or "Borrador") == "Aprobada"]
    sin_respuesta = [c for c in cots_todas if (c.get("estado") or "Borrador") == "Enviada"]
    rechazadas = [c for c in cots_todas if (c.get("estado") or "Borrador") == "Rechazada"]
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Cotizaciones realizadas", len(cots_todas))
    m2.metric("Cotizaciones aprobadas", len(aprobadas))
    m3.metric("Cotizaciones sin respuesta", len(sin_respuesta))
    m4.metric("Cotizaciones rechazadas", len(rechazadas))

    st.markdown("#### Seguimiento por estado")
    clientes_con_cot = sorted({c["cliente_nombre"] for c in cots_todas if c.get("cliente_nombre")})
    c1, c2 = st.columns([2.3, 1.3])
    busqueda = c1.text_input(
        "🔎 Buscar por número o cliente", key="panel_busqueda_cot",
        placeholder="Filtrar el tablero por número o cliente...",
    )
    cliente_filtro = c2.selectbox("Cliente", ["Todos"] + clientes_con_cot, key="panel_filtro_cliente_cot")

    cots = cots_todas
    if busqueda:
        b = busqueda.strip().lower()
        cots = [c for c in cots if b in (c.get("numero") or "").lower() or b in (c.get("cliente_nombre") or "").lower()]
    if cliente_filtro != "Todos":
        cots = [c for c in cots if c.get("cliente_nombre") == cliente_filtro]

    etapa_por_cliente = _mapa_etapa_por_cliente()
    cols = st.columns(len(ESTADOS))
    for col, estado in zip(cols, ESTADOS):
        with col, st.container(border=True, key=f"zona_cot_{estado}"):
            items = [c for c in cots if (c.get("estado") or "Borrador") == estado]
            st.markdown(f"**{estado}**")
            st.caption(f"{len(items)} cotización(es)")
            for c in items[:LIMITE_TARJETAS_POR_COLUMNA]:
                with st.container(border=True, key=f"cardwrap_cot_{c['id']}"):
                    st.markdown(
                        _card_info_html(c["numero"], [
                            c.get("cliente_nombre") or "sin cliente",
                            c.get("productos_desc") or "sin productos cargados",
                            money(c.get("total_usd_civa") or 0),
                        ]),
                        unsafe_allow_html=True,
                    )
                    # Etapa CRM del contacto vinculado, en vivo (ver
                    # _mapa_etapa_por_cliente) — mismo criterio que en el
                    # historial del Cotizador y en Clientes.
                    etapa_crm = etapa_por_cliente.get(c.get("cliente_id"))
                    if etapa_crm:
                        st.markdown(_badge_etapa(etapa_crm), unsafe_allow_html=True)
                    st.button(
                        "Abrir →", key=f"panelcot_{c['id']}", wrap=False,
                        on_click=_abrir_cotizacion_en_cotizador, args=(c["id"],),
                    )
            if len(items) > LIMITE_TARJETAS_POR_COLUMNA:
                st.caption(f"+{len(items) - LIMITE_TARJETAS_POR_COLUMNA} más — refiná la búsqueda para verlas")


def _render_panel_importaciones():
    imp_sel_id = st.session_state.get("panel_importacion_seleccionada")
    if imp_sel_id is not None:
        imp = db.get_importacion(imp_sel_id)
        if imp is None:
            st.session_state.panel_importacion_seleccionada = None
            st.rerun()

        def _volver_al_panel():
            st.session_state.panel_importacion_seleccionada = None

        st.button(
            "Volver al panel", icon=":material/arrow_back:",
            key="volver_panel", on_click=_volver_al_panel,
        )
        cliente = db.get_cliente(imp["cliente_id"])
        st.subheader(f"{imp['numero']} — {cliente['nombre'] if cliente else 'sin cliente'}")
        cotizaciones_cliente = db.list_cotizaciones(cliente_id=imp["cliente_id"])
        _render_contenido_importacion(imp, cotizaciones_cliente)
        return

    imps_todas = db.list_importaciones()
    if not imps_todas:
        st.info("Todavía no hay importaciones cargadas.")
        return

    # "En proceso" = cualquier etapa de gestión activa (En producción, En
    # coordinación, En tránsito, En puerto/Dep.Fiscal) — todas menos Entregada.
    en_proceso = [i for i in imps_todas if (i.get("estado") or ESTADOS_IMPORTACION[0]) != "Entregada"]
    completadas = [i for i in imps_todas if (i.get("estado") or ESTADOS_IMPORTACION[0]) == "Entregada"]
    m1, m2, m3 = st.columns(3)
    m1.metric("Importaciones totales", len(imps_todas))
    m2.metric("En proceso", len(en_proceso))
    m3.metric("Completadas", len(completadas))

    st.markdown("#### Seguimiento por estado")
    clientes_con_imp = sorted({i["cliente_nombre"] for i in imps_todas if i.get("cliente_nombre")})

    c1, c2, c3 = st.columns([2.2, 1.3, 1.3])
    busqueda = c1.text_input(
        "🔎 Buscar por número, cliente o producto", key="panel_busqueda_imp",
        placeholder="Filtrar el tablero por número, cliente o producto...",
    )
    cliente_filtro = c2.selectbox("Cliente", ["Todos"] + clientes_con_imp, key="panel_filtro_cliente_imp")
    tipo_filtro = c3.selectbox("Tipo de envío", ["Todos"] + TIPOS_ENVIO, key="panel_filtro_tipo_imp")

    c4, c5, c6 = st.columns([1.3, 1.3, 1.3])
    fecha_desde = c4.date_input("ETA desde", value=None, key="panel_filtro_eta_desde", format="DD/MM/YYYY")
    fecha_hasta = c5.date_input("ETA hasta", value=None, key="panel_filtro_eta_hasta", format="DD/MM/YYYY")
    c6.write("")
    ocultar_entregadas = c6.checkbox("Ocultar entregadas", key="panel_ocultar_entregadas")

    imps = imps_todas
    if busqueda:
        b = busqueda.strip().lower()
        imps = [
            i for i in imps
            if b in (i.get("numero") or "").lower()
            or b in (i.get("cliente_nombre") or "").lower()
            or b in (i.get("producto") or "").lower()
        ]
    if cliente_filtro != "Todos":
        imps = [i for i in imps if i.get("cliente_nombre") == cliente_filtro]
    if tipo_filtro != "Todos":
        imps = [i for i in imps if i.get("tipo_envio") == tipo_filtro]
    if fecha_desde:
        imps = [i for i in imps if (_parse_fecha(i.get("eta")) or date.min) >= fecha_desde]
    if fecha_hasta:
        imps = [i for i in imps if (_parse_fecha(i.get("eta")) or date.max) <= fecha_hasta]
    if ocultar_entregadas:
        imps = [i for i in imps if (i.get("estado") or ESTADOS_IMPORTACION[0]) != "Entregada"]

    estados_visibles = [e for e in ESTADOS_IMPORTACION if not (ocultar_entregadas and e == "Entregada")]
    cols = st.columns(len(estados_visibles))
    for col, estado in zip(cols, estados_visibles):
        with col, st.container(border=True, key=f"zona_imp_{estado}"):
            items = [i for i in imps if (i.get("estado") or ESTADOS_IMPORTACION[0]) == estado]
            st.markdown(f"**{estado}**")
            st.caption(f"{len(items)} importación(es)")
            for imp in items[:LIMITE_TARJETAS_POR_COLUMNA]:
                with st.container(border=True, key=f"cardwrap_imp_{imp['id']}"):
                    icono = ICONO_TIPO_ENVIO.get(imp.get("tipo_envio"), "")
                    titulo = f"{icono} {imp['numero']}" if icono else imp["numero"]
                    lineas = [imp.get("cliente_nombre") or "sin cliente"]
                    if imp.get("producto"):
                        lineas.append(f"📦 {imp['producto']}")
                    if imp.get("etd"):
                        lineas.append(f"ETD: {_fmt_fecha(imp['etd'])}")
                    if imp.get("eta"):
                        lineas.append(f"ETA: {_fmt_fecha(imp['eta'])}")
                    st.markdown(_card_info_html(titulo, lineas), unsafe_allow_html=True)
                    if st.button("Ver detalle →", key=f"panelimp_{imp['id']}", wrap=False):
                        st.session_state.panel_importacion_seleccionada = imp["id"]
                        st.rerun()
            if len(items) > LIMITE_TARJETAS_POR_COLUMNA:
                st.caption(f"+{len(items) - LIMITE_TARJETAS_POR_COLUMNA} más — refiná la búsqueda para verlas")


def _render_configuracion():
    st.caption(
        "Parámetros generales del sistema. Se guardan en la base y los usa toda la app — "
        "ningún valor de estos queda fijo en el código."
    )
    dias_actual = int(db.get_setting("dias_sin_respuesta", db.SETTINGS_DEFAULTS["dias_sin_respuesta"]))
    with st.form("form_config_general"):
        nuevo_valor = st.number_input(
            "CRM — Días sin respuesta para considerar una cotización 'Enviada' como vencida "
            "(usado en los sub-filtros de seguimiento del CRM)",
            min_value=1, value=dias_actual, step=1,
        )
        if st.form_submit_button("💾 Guardar", type="primary"):
            db.set_setting("dias_sin_respuesta", int(nuevo_valor))
            _flash("Configuración guardada.")
            st.rerun()

    st.divider()
    st.markdown("#### 🔁 Migración retroactiva de contactos CRM")
    st.caption(
        "Crea el contacto de CRM que le falte a cada cliente ya cargado (los que nacieron antes "
        "de que el alta de cliente generara uno solo), en la etapa que corresponde a su evidencia "
        "real: importación creada o cotización aprobada → Cliente; cualquier otra cotización → "
        "Cotizado; nada todavía → Nuevo. Segura de correr más de una vez: los clientes que ya "
        "tienen contacto vinculado se saltean solos."
    )
    if st.button("▶️ Ejecutar migración", key="btn_migrar_retroactivo"):
        with st.spinner("Migrando contactos..."):
            try:
                resultado = _migrar_contactos_retroactivos()
            except db.DBError as e:
                st.error(f"No se pudo completar la migración: {e}")
                resultado = None
        if resultado is None:
            pass
        elif not resultado:
            st.info("No había clientes sin contacto de CRM vinculado — nada para migrar.")
        else:
            st.success(f"Se creó el contacto de CRM para {len(resultado)} cliente(s):")
            for nombre, etapa in resultado:
                st.markdown(f"- **{nombre}** → {crm.ETAPA_LABEL.get(etapa, etapa)}")

    st.divider()
    st.markdown("#### 👤 Usuarios")
    for u in db.list_usuarios():
        cu1, cu2, cu3 = st.columns([2.5, 1, 1])
        cu1.markdown(f"**{u['nombre']}** (`{u['username']}`)")
        cu2.caption("Activo" if u["activo"] else "Desactivado")
        es_yo = u["id"] == st.session_state.usuario_autenticado.get("id")
        if not es_yo:
            if cu3.button("Desactivar" if u["activo"] else "Reactivar", key=f"toggleuser_{u['id']}"):
                db.set_usuario_activo(u["id"], 0 if u["activo"] else 1)
                st.rerun()
    with st.form("form_nuevo_usuario", clear_on_submit=True):
        st.caption("Agregar usuario nuevo")
        nu_username = st.text_input("Usuario", key="nu_username")
        nu_nombre = st.text_input("Nombre", key="nu_nombre")
        nu_password = st.text_input("Contraseña inicial", type="password", key="nu_password")
        if st.form_submit_button("Crear"):
            if nu_username and nu_nombre and nu_password:
                db.create_usuario(nu_username, nu_password, nu_nombre)
                _flash(f"Usuario '{nu_nombre}' creado.")
                st.rerun()
            else:
                st.error("Completá usuario, nombre y contraseña.")

    st.divider()
    st.markdown("#### 💾 Backup")
    # exportar_backup_sqlite() hace un SELECT * por cada tabla de la base
    # (11 en total) — con Turso eso tarda ~3s. st.download_button necesita
    # los bytes YA armados para poder dibujarse, así que antes corría en
    # CADA render de Configuración (y como st.tabs ejecuta las 3 pestañas
    # siempre, eso eran ~3s de más en CADA carga del Panel de Control,
    # aunque nadie tocara este botón). Separarlo en dos pasos — preparar,
    # después descargar — hace que esas 11 consultas corran solo cuando de
    # verdad se van a usar.
    if st.button("📦 Preparar backup para descargar"):
        with st.spinner("Armando el backup..."):
            st.session_state["_backup_bytes"] = db.exportar_backup_sqlite()
            st.session_state["_backup_nombre"] = f"skybridge_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
    if "_backup_bytes" in st.session_state:
        st.download_button(
            "⬇️ Descargar backup de la base",
            data=st.session_state["_backup_bytes"],
            file_name=st.session_state["_backup_nombre"],
            mime="application/octet-stream",
        )

    st.divider()
    st.markdown("#### ♻️ Restaurar backup")
    st.caption("Reemplaza la base de datos actual por un archivo .db subido acá. Esta acción no se puede deshacer.")
    with st.form("form_restaurar_backup_config"):
        archivo_restaurar = st.file_uploader("Archivo .db a restaurar", type=["db"], key="restaurar_backup_uploader")
        confirmar_restaurar = st.checkbox("Entiendo que esto reemplaza todos los datos actuales")
        if st.form_submit_button("Restaurar este backup"):
            if not archivo_restaurar:
                st.error("Subí un archivo primero.")
            elif not confirmar_restaurar:
                st.error("Tenés que confirmar el checkbox antes de restaurar.")
            else:
                db.backup_antes_de_borrar("restauracion_manual")
                db.restaurar_desde_sqlite_bytes(archivo_restaurar.getvalue())
                st.success("Base restaurada. La página se va a recargar.")
                st.rerun()


def vista_panel_control():
    st.header("📊 Panel de Control")
    st.caption("Seguimiento general e individual de cotizaciones e importaciones de todos los clientes.")

    tab_cot, tab_imp, tab_config = st.tabs(["🧮 Cotizaciones", "📦 Importaciones", "⚙️ Configuración"])
    with tab_cot:
        _render_panel_cotizaciones()
    with tab_imp:
        _render_panel_importaciones()
    with tab_config:
        _render_configuracion()


# ---------------------------------------------------------------- HISTORIAL
# Vista por default del Cotizador: el botón "Nueva cotización" y, debajo, el
# historial completo (filtros + paginado) en una única sección — antes eran
# dos pestañas ("Cotización en edición" / "Historial") que separaban crear de
# retomar, cuando en la práctica es el mismo flujo. Editar/crear pasa a un
# modo aparte (ver vista_cotizador) en vez de convivir como pestaña.
COTIZACIONES_POR_PAGINA = 10


def _crear_nueva_cotizacion_desde_historial():
    new_id = db.create_cotizacion_borrador()
    _cargar_estado_cotizacion(new_id)
    st.session_state.cotizador_modo = "editor"
    _flash("Nueva cotización creada.")


def _toggle_detalle_cotizacion(cid):
    st.session_state.cotizador_detalle_id = None if st.session_state.get("cotizador_detalle_id") == cid else cid


def _duplicar_cotizacion_historial(cid):
    db.duplicate_cotizacion(cid)
    _flash("Cotización duplicada.")


def _cambiar_estado_cotizacion(cot_id, cliente_id):
    # Control independiente en la card del listado — antes había que abrir
    # el editor entero y buscar "Estado" adentro del expander de datos
    # generales, junto a una docena de campos sin relación, para cambiarlo.
    nuevo_estado = st.session_state[f"cotestadosel_{cot_id}"]
    db.set_estado_cotizacion(cot_id, nuevo_estado)
    mensaje = "Estado actualizado."
    # Mismo criterio que al guardar desde el editor: "Aprobada" hace ganar
    # el trato en el CRM (ver _marcar_ganado_por_cliente).
    if nuevo_estado == "Aprobada" and cliente_id:
        contacto_vinculado = _marcar_ganado_por_cliente(cliente_id)
        if contacto_vinculado:
            mensaje += f" '{contacto_vinculado['nombre']}' pasó a Cliente."
    if st.session_state.get("cot_cab", {}).get("id") == cot_id:
        st.session_state.cot_cab["estado"] = nuevo_estado
    _flash(mensaje)


def _eliminar_cotizacion_historial(cid):
    # Compartida por el listado (historial) y el editor completo: si la
    # cotización borrada es la que estaba abierta en el editor, además de
    # limpiar su estado hay que rebotar a "historial" — si no, la pantalla
    # se queda intentando dibujar el editor de una cotización que ya no existe.
    db.backup_antes_de_borrar("cotizacion")
    db.delete_cotizacion(cid)
    if st.session_state.get("cot_cab", {}).get("id") == cid:
        for k in ["cot_cab", "cot_productos", "cot_gastos"]:
            st.session_state.pop(k, None)
        st.session_state.cotizador_modo = "historial"
    if st.session_state.get("cotizador_detalle_id") == cid:
        st.session_state.cotizador_detalle_id = None
    _flash("Cotización eliminada.", icon="🗑️")


@st.dialog("Confirmar cotización aprobada")
def _dialog_confirmar_aprobada(cot_id, cliente_id, nombre_cliente):
    # El control "Cambiar" es rápido a propósito (2 clics, sin abrir la
    # cotización) — pero "Aprobada" no es un estado más: dispara
    # _marcar_ganado_por_cliente, que pasa al cliente a "Cliente" (Ganado)
    # en el CRM de un solo sentido (no se deshace solo si después el
    # estado de la cotización vuelve para atrás). Justo ese combo — un
    # clic de más en el popover, corregido minutos después sin tocar el
    # CRM a mano — sacó una vez a un cliente real de "Cotizados" sin que
    # nadie se diera cuenta hasta que apareció mezclado en "Cliente". Este
    # paso extra es la forma de que no se repita.
    st.warning(
        f"Vas a marcar esta cotización como **Aprobada**. Eso pasa automáticamente a "
        f"**'{nombre_cliente}'** a la etapa **Cliente** en el CRM — es un pase de un solo "
        f"sentido: si después volvés el estado de la cotización para atrás, el cliente se "
        f"queda en 'Cliente' igual, no se revierte solo."
    )
    c1, c2 = st.columns(2)
    if c1.button("Cancelar", use_container_width=True, key=f"cancelaraprobada_{cot_id}"):
        st.rerun()
    if c2.button("Sí, aprobar", type="primary", use_container_width=True, key=f"confirmaraprobada_{cot_id}"):
        _cambiar_estado_cotizacion(cot_id, cliente_id)
        st.rerun()


@st.dialog("Eliminar cotización")
def _dialog_eliminar_cotizacion(cid, numero):
    st.warning(f"¿Confirmás eliminar **{numero}**? Esta acción no se puede deshacer.")
    c1, c2 = st.columns(2)
    if c1.button("Cancelar", use_container_width=True, key=f"cancelardel_{cid}"):
        st.rerun()
    if c2.button("Sí, eliminar", type="primary", use_container_width=True, key=f"confirmardel_{cid}"):
        _eliminar_cotizacion_historial(cid)
        st.rerun()


def _ir_a_ficha_cliente_desde_cotizacion(cliente_id):
    # Por ID de cliente, no por su contacto de CRM: antes enlazaba a la
    # ficha del contacto vinculado en CRM (db.get_contact_by_cliente_id), y
    # como la enorme mayoría de los clientes no tiene ningún contacto de CRM
    # vinculado (ver nota en vista_clientes), quedaba sin enlace para casi
    # todos (ej. Barwatex). cotizaciones.cliente_id siempre apunta a un
    # cliente real (ver create_cotizacion_borrador), así que este enlace
    # funciona siempre, sin depender del CRM ni de ninguna etapa.
    st.session_state.cliente_seleccionado = cliente_id
    st.session_state.pagina_nav = "📋 Clientes"


def _fmt_fecha_cotizacion(valor):
    # A diferencia de _fmt_fecha (fechas de CRM, guardadas como YYYY-MM-DD),
    # cotizaciones.fecha incluye hora ("YYYY-MM-DD HH:MM:SS").
    try:
        return datetime.strptime(valor, "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y")
    except (TypeError, ValueError):
        return valor or "—"


def _render_fila_cotizacion(c, etapa_por_cliente=None):
    """Fila 'hoja' de una cotización, con el mismo lenguaje visual que las
    cards de Clientes: número/fecha, cliente (con link a su ficha en el
    módulo Clientes, por cliente_id — funciona siempre, tenga o no ese
    cliente un contacto de CRM vinculado, más la etapa CRM de ese contacto
    si lo tiene), producto(s), estado (badge + control propio para
    cambiarlo sin entrar al editor) y 5 acciones (editar/PDF completo/ver
    detalle/duplicar/eliminar) con íconos de línea consistentes en vez de
    emojis."""
    pdf_simple, pdf_completo, nombre_pdf = _generar_pdfs_cotizacion(c["id"], c["actualizado_en"])
    estado_actual = c.get("estado") or "Borrador"
    if etapa_por_cliente is None:
        etapa_por_cliente = _mapa_etapa_por_cliente()

    with st.container(border=True, key=f"cotfila_{c['id']}"):
        c1, c2, c3, c_estado, c4 = st.columns([1.5, 2.0, 2.3, 1.5, 2.7], vertical_alignment="center")
        c1.markdown(
            f'<div style="font-weight:700; font-size:14px;">{html.escape(c["numero"])}</div>'
            f'<div style="font-size:12px; color:#64748b; margin-top:2px;">{_fmt_fecha_cotizacion(c.get("fecha"))}</div>',
            unsafe_allow_html=True,
        )
        with c2:
            nombre_cliente = c.get("cliente_nombre") or "sin cliente"
            if c.get("cliente_id"):
                st.button(
                    nombre_cliente, key=f"cotcliente_{c['id']}",
                    on_click=_ir_a_ficha_cliente_desde_cotizacion, args=(c["cliente_id"],),
                    help=f"Ver la ficha de '{nombre_cliente}' en Clientes",
                )
            else:
                st.markdown(
                    f'<div style="font-size:13px; font-weight:600;">{html.escape(nombre_cliente)}</div>',
                    unsafe_allow_html=True,
                )
            # Etapa CRM del contacto vinculado (si tiene) — se lee en vivo
            # de contacts.etapa en cada render, así que marcar "en
            # negociación" (o cualquier otra etapa) desde la card del CRM
            # se ve reflejado acá al toque, sin ningún paso extra.
            etapa_crm = etapa_por_cliente.get(c.get("cliente_id"))
            if etapa_crm:
                st.markdown(_badge_etapa(etapa_crm), unsafe_allow_html=True)
        c3.markdown(
            # text-align: left explícito: sin esto el "—" de fallback (un
            # solo carácter, a diferencia del texto real que ocupa varias
            # palabras) quedaba centrado en la columna en vez de alineado
            # al resto de las filas de la lista.
            f'<div style="font-size:12px; color:#64748b; line-height:1.35; text-align:left;">'
            f'{html.escape(c.get("productos_desc") or "—")}</div>',
            unsafe_allow_html=True,
        )
        with c_estado:
            with st.container(key=f"cotestado_{c['id']}"):
                st.markdown(_badge_estado_cotizacion(estado_actual), unsafe_allow_html=True)
                with st.popover("Cambiar", help="Cambiar estado sin abrir la cotización"):
                    nuevo_estado_sel = st.selectbox(
                        "Estado", ESTADOS, index=ESTADOS.index(estado_actual),
                        key=f"cotestadosel_{c['id']}",
                    )
                    # "Aprobada" no es un cambio más — ver _dialog_confirmar_aprobada.
                    # Sin este chequeo, el botón queda como on_click directo (igual
                    # que antes) para el resto de los estados, que no tienen ningún
                    # efecto de un solo sentido en el CRM.
                    contacto_vinculado = (
                        db.get_contact_by_cliente_id(c["cliente_id"]) if c.get("cliente_id") else None
                    )
                    dispara_ganado = (
                        nuevo_estado_sel == "Aprobada" and c.get("cliente_id")
                        and (not contacto_vinculado or (contacto_vinculado.get("etapa") or crm.STAGE_INICIAL) != "Ganado")
                    )
                    if st.button("Guardar", key=f"cotestadosave_{c['id']}", type="primary", use_container_width=True):
                        if dispara_ganado:
                            _dialog_confirmar_aprobada(c["id"], c["cliente_id"], c.get("cliente_nombre") or "el cliente")
                        else:
                            _cambiar_estado_cotizacion(c["id"], c.get("cliente_id"))
                            st.rerun()
        with c4:
            with st.container(key=f"cotfilaacciones_{c['id']}"):
                st.button(
                    "", icon=":material/edit:", key=f"cotedit_{c['id']}", help="Editar",
                    on_click=_abrir_cotizacion_en_cotizador, args=(c["id"],),
                )
                st.download_button(
                    "", icon=":material/picture_as_pdf:", data=pdf_completo,
                    file_name=f"{nombre_pdf} - Completo.pdf", mime="application/pdf",
                    key=f"cotpdf_{c['id']}", help="Descargar PDF completo",
                )
                st.button(
                    "", icon=":material/visibility:", key=f"cotdet_{c['id']}", help="Ver detalle",
                    on_click=_toggle_detalle_cotizacion, args=(c["id"],),
                )
                st.button(
                    "", icon=":material/content_copy:", key=f"cotdup_{c['id']}", help="Duplicar",
                    on_click=_duplicar_cotizacion_historial, args=(c["id"],),
                )
                if st.button(
                    "", icon=":material/delete:", key=f"coteli_{c['id']}", help="Eliminar",
                ):
                    _dialog_eliminar_cotizacion(c["id"], c["numero"])

    if st.session_state.get("cotizador_detalle_id") == c["id"]:
        with st.container(border=True, key=f"formrow_cotdetalle_{c['id']}"):
            st.caption(f"Vista previa (simplificado) — {c['numero']}")
            st.pdf(pdf_simple, key=f"cotdetpdf_{c['id']}")


def _render_historial_cotizaciones():
    clientes = db.list_clientes()
    mapa_clientes = {c["id"]: c["nombre"] for c in clientes}

    # Un buscador único (número/cliente/producto) siempre visible + un
    # popover "Filtros avanzados" con los 3 criterios combinables (cliente
    # exacto, producto, rango de fechas) para cuando hace falta cruzarlos —
    # antes los 3 vivían siempre abiertos en una fila, ocupando más lugar
    # del que se usa la mayoría de las veces. El date_input de rango, al
    # vivir en su propia fila dentro del popover (no apretado en 1 de 3
    # columnas de la card), ya no queda con el texto cortado.
    with st.container(border=True, key="formrow_cotizador_filtros"):
        col_buscar, col_avanzado, col_nueva = st.columns([3.2, 1.5, 1.6], vertical_alignment="bottom")
        busqueda = col_buscar.text_input(
            "Buscar", placeholder="Buscar por número, cliente o producto...",
            key="cotizador_busqueda",
        )
        with col_avanzado:
            with st.popover("Filtros avanzados", icon=":material/tune:", use_container_width=True):
                cliente_filtro = st.selectbox(
                    "Cliente", options=[c["id"] for c in clientes],
                    format_func=lambda x: mapa_clientes[x],
                    index=None, placeholder="Buscar cliente...",
                    key="cotizador_filtro_cliente",
                )
                producto_filtro = st.text_input(
                    "Producto", placeholder="Buscar por producto...", key="cotizador_filtro_producto",
                )
                rango_fechas = st.date_input(
                    "Rango de fechas", value=(), format="DD/MM/YYYY",
                    key="cotizador_filtro_fechas",
                )
        col_nueva.button(
            "Nueva cotización", icon=":material/add:", type="primary", key="nueva_cotizacion_historial",
            on_click=_crear_nueva_cotizacion_desde_historial, use_container_width=True,
        )

    fecha_desde = fecha_hasta = None
    if isinstance(rango_fechas, (tuple, list)) and len(rango_fechas) == 2:
        fecha_desde, fecha_hasta = rango_fechas

    # session_state.setdefault en vez de pisar el valor cada corrida: si no,
    # cualquier cambio de filtro (que dispara un rerun) resetearía siempre a
    # página 1 aunque el usuario ya hubiera avanzado — acá se resetea a 1
    # solo cuando cambia el conjunto de resultados, más abajo.
    st.session_state.setdefault("cotizador_pagina", 1)

    cots = db.list_cotizaciones(
        search=busqueda.strip(),
        cliente_id=cliente_filtro,
        producto=producto_filtro.strip(),
        fecha_desde=fecha_desde.isoformat() if fecha_desde else None,
        fecha_hasta=fecha_hasta.isoformat() if fecha_hasta else None,
    )
    total = len(cots)

    st.markdown(
        f'<div style="font-size:13px; font-weight:600; color:#64748b; margin:2px 0 10px 2px;">'
        f'{total} cotización(es) encontrada(s)</div>',
        unsafe_allow_html=True,
    )

    if not cots:
        st.info("No hay cotizaciones que coincidan con los filtros.")
        return

    total_paginas = max(1, math.ceil(total / COTIZACIONES_POR_PAGINA))
    # Clampeo (no reset a 1): permite que, al volver de otro filtro, la
    # página elegida se mantenga mientras siga siendo válida.
    st.session_state.cotizador_pagina = min(st.session_state.cotizador_pagina, total_paginas)
    pagina = st.session_state.cotizador_pagina
    inicio = (pagina - 1) * COTIZACIONES_POR_PAGINA
    cots_pagina = cots[inicio:inicio + COTIZACIONES_POR_PAGINA]

    etapa_por_cliente = _mapa_etapa_por_cliente()
    for c in cots_pagina:
        _render_fila_cotizacion(c, etapa_por_cliente=etapa_por_cliente)

    if total_paginas > 1:
        cnav1, cnav2, cnav3 = st.columns([1, 2, 1], vertical_alignment="center")
        if cnav1.button("← Anterior", disabled=pagina <= 1, use_container_width=True, key="cotizador_pag_ant"):
            st.session_state.cotizador_pagina = pagina - 1
            st.rerun()
        cnav2.markdown(
            f'<div style="text-align:center; font-size:13px; color:#64748b;">'
            f'Página {pagina} de {total_paginas}</div>',
            unsafe_allow_html=True,
        )
        if cnav3.button("Siguiente →", disabled=pagina >= total_paginas, use_container_width=True, key="cotizador_pag_sig"):
            st.session_state.cotizador_pagina = pagina + 1
            st.rerun()


def vista_cotizador():
    # Una sola sección a la vez: por default el listado (buscador + Nueva
    # cotización + historial paginado); "➕ Nueva cotización" o "✏️ Editar"
    # pasan a modo "editor" (ver _abrir_cotizacion_en_cotizador /
    # _cotizar_cliente / _crear_nueva_cotizacion_desde_historial), y
    # "← Volver al listado" regresa. El header + caption (mismo patrón que
    # "🤝 CRM — Contactos") solo se muestra en el listado — el editor ya
    # tiene su propio subheader con el número de cotización.
    st.session_state.setdefault("cotizador_modo", "historial")
    if st.session_state.cotizador_modo == "editor" and "cot_cab" in st.session_state:
        _render_cotizador_editor()
        return

    st.session_state.cotizador_modo = "historial"
    st.header("🧮 Cotizador de Importación")
    st.caption("Cotizaciones realizadas — armá una nueva o retomá cualquiera del historial completo.")
    _render_historial_cotizaciones()


# ---------------------------------------------------------------- CRM: CONTACTOS
def _seguimiento_vencido(c):
    d = _parse_fecha(c.get("proximo_seguimiento"))
    return d is not None and d < date.today()


def _dias_desde_actividad(c):
    dt = crm.parse_dt(c.get("ultima_actividad") or c.get("fecha_alta"))
    return (datetime.now() - dt).days if dt else None


def _relabel_transicion_etapas(texto):
    """Remapea un texto 'Anterior → Nueva' (formato de activity_log.texto
    para eventos cambio_etapa) a través de crm.ETAPA_LABEL — ej. 'Nuevo →
    Ganado' se muestra 'Nuevo → Cliente' — sin tocar lo que quedó guardado
    en la base."""
    if not texto or "→" not in texto:
        return texto
    izq, der = texto.split("→", 1)
    izq, der = izq.strip(), der.strip()
    return f"{crm.ETAPA_LABEL.get(izq, izq)} → {crm.ETAPA_LABEL.get(der, der)}"


def _badge_etapa(etapa):
    return _badge(crm.ETAPA_LABEL.get(etapa, etapa), crm.COLOR_ETAPA.get(etapa, "#64748B"))


def _mapa_etapa_por_cliente():
    """{cliente_id: etapa} de todos los contactos con cliente vinculado —
    para mostrar, junto al estado propio de una cotización o cliente, la
    etapa CRM del contacto vinculado (Cotizador, Clientes, Panel de
    Control). Se lee en vivo del contacto en cada render — sin duplicar el
    dato ni guardar nada nuevo — así que cambiar la etapa desde la card del
    CRM (ej. "Cotizado" → "Negociación") se ve reflejado en todos esos
    lugares de inmediato, sin ningún paso extra."""
    return {c["cliente_id"]: c.get("etapa") or crm.STAGE_INICIAL for c in db.list_contacts() if c.get("cliente_id")}


def _badge_estado_importacion(estado):
    return _badge(estado, COLOR_ESTADO_IMPORTACION.get(estado, "#64748B"))


def _badge_estado_cotizacion(estado):
    return _badge(estado, COLOR_ESTADO_COTIZACION.get(estado, "#64748B"))


def _ver_ficha_contacto(contact_id):
    st.session_state.contacto_seleccionado = contact_id


KANBAN_COLUMNAS = [
    ("nuevos", "🆕 Nuevos"),
    ("contactados", "✅ Contactados"),
    ("cotizados", "🧾 Cotizados"),
    ("negociacion", "🤝 Negociación"),
    ("cliente", "🏆 Cliente"),
    ("perdidos", "❌ Descartado"),
]
COTIZACION_SUBFILTROS = {
    "Todas": None,
    "A revisar": "a_revisar",
    "Activas": "activa",
    "Sin respuesta": "sin_respuesta",
    "Aprobadas": "aprobada_pendiente",
    "Estancadas": "estancado",
}


def _render_carga_rapida_contacto():
    """Alta rápida de un contacto (nombre/CUIT/teléfono + botón) — espacio
    propio y siempre visible arriba del buscador de "👥 Contactos", mismo
    lugar y misma idea que ya existía antes (colapsado detrás de "＋ Nuevo
    contacto"): para sumar un contacto suelto no hace falta abrir nada."""
    with st.container(border=True, key="formrow_crm_carga_rapida"):
        with st.form("form_carga_rapida", clear_on_submit=True):
            cr1, cr2, cr3, cr4, cr5 = st.columns([2.3, 2, 1.7, 2.2, 1.3], vertical_alignment="bottom")
            nombre_rapido = cr1.text_input("Nombre", label_visibility="collapsed", placeholder="Nombre de contacto")
            cuit_rapido = cr2.text_input("CUIT", label_visibility="collapsed", placeholder="CUIT (opcional)")
            # Placeholder acortado ("Teléfono / WhatsApp" completo no entraba
            # en esta columna, angosta a propósito porque un teléfono es
            # corto — se cortaba en "...WhatsAp").
            telefono_rapido = cr3.text_input("Teléfono", label_visibility="collapsed", placeholder="Tel. / WhatsApp")
            email_rapido = cr4.text_input("Email", label_visibility="collapsed", placeholder="Email (opcional)")
            if cr5.form_submit_button("＋ Agregar", use_container_width=True, type="primary"):
                email_norm = crm.normalizar_email(email_rapido)
                if not nombre_rapido.strip():
                    st.error("El nombre es obligatorio.")
                elif email_norm and not crm.email_valido(email_norm):
                    st.error("Ese email no parece válido.")
                elif not crm.whatsapp_valido(telefono_rapido):
                    st.error("Ese teléfono/WhatsApp no parece válido (revisá la cantidad de dígitos).")
                else:
                    db.create_contact(
                        nombre_rapido.strip(), "",
                        email_norm, crm.normalizar_whatsapp(telefono_rapido),
                        "Manual", crm.STAGE_INICIAL, "", None, cuit_rapido.strip(),
                    )
                    _flash(f"Contacto '{nombre_rapido}' agregado.")
                    st.rerun()


def _render_importar_contactos_masivo():
    """Importación masiva desde Excel/CSV — colapsada detrás de "＋ Nuevo
    contacto" (a diferencia de la carga rápida, esta sí vale la pena
    mantenerla oculta por default: mapeo de columnas, vista previa, etc.,
    mucho más pesada que completar 3 campos)."""
    with st.expander("📥 Importar contactos desde Excel/CSV"):
        _render_importar_contactos()
    st.write("")


LIMITE_FILAS_CONTACTOS = 25


def _badge_vencido():
    return _badge("⚠ Vencido", "#DC2626")


def _marcar_descartado_boton(contact_id):
    contacto = db.get_contact(contact_id)
    autor = st.session_state.get("crm_autor", "")
    db.change_etapa_contacto(contact_id, "Perdido", autor)
    _flash(f"'{contacto['nombre']}' marcado como Descartado.")


def _actualizar_etapa_contacto(contact_id, key_selectbox):
    # El popover "Actualizar estado" ya restringe el desplegable a las
    # etapas que son un avance real desde la actual (crm.es_avance) — acá
    # no hace falta re-chequear, a diferencia de _avanzar_etapa_si_corresponde.
    nueva_etapa = st.session_state[key_selectbox]
    contacto = db.get_contact(contact_id)
    autor = st.session_state.get("crm_autor", "")
    db.change_etapa_contacto(contact_id, nueva_etapa, autor)
    _flash(f"'{contacto['nombre']}' pasó a '{crm.ETAPA_LABEL.get(nueva_etapa, nueva_etapa)}'.")


@st.dialog("Eliminar contacto")
def _dialog_eliminar_contacto_crm(contact_id, nombre):
    st.warning(f"¿Confirmás eliminar a **{nombre}**? Esta acción no se puede deshacer.")
    c1, c2 = st.columns(2)
    if c1.button("Cancelar", use_container_width=True, key=f"crmcancelardel_{contact_id}"):
        st.rerun()
    if c2.button("Sí, eliminar", type="primary", use_container_width=True, key=f"crmconfirmardel_{contact_id}"):
        db.delete_contact(contact_id)
        if st.session_state.get("contacto_seleccionado") == contact_id:
            st.session_state.contacto_seleccionado = None
        _flash(f"Contacto '{nombre}' eliminado.", icon="🗑️")
        st.rerun()


def _render_fila_contacto_crm(c, en_curso_por_cliente=None):
    """Card de un contacto — mismo lenguaje visual que Clientes: barra de
    acento a la izquierda coloreada por etapa (crm.COLOR_ETAPA, vía el key
    f"filacrm_{slug}_{id}"), accesos directos de WhatsApp/mail/ficha de
    Cliente vinculado/eliminar (con confirmación) como íconos compactos,
    badge rojo si el seguimiento está vencido (antes solo un ⚠️ chiquito
    pegado a la fecha), y 4 botones en 2 filas de a 2 ('Cotizar'/'Ver
    ficha →' arriba, 'Actualizar estado'/'Marcar como Descartado' abajo).

    en_curso_por_cliente no-None significa "estamos en la columna Cliente"
    (etapa 'Ganado', ver crm.kanban_grupo): ahí no tiene sentido "avanzar
    etapa" ni "descartar" — en su lugar la card muestra un resumen
    operativo de solo lectura (importaciones en curso + última actividad)."""
    etapa = c.get("etapa") or crm.STAGE_INICIAL
    slug = crm.ETAPA_SLUG.get(etapa, "nuevo")

    with st.container(border=True, key=f"filacrm_{slug}_{c['id']}"):
        # c4 y c5 más anchas que en la primera versión: los botones
        # partían en 2 líneas dentro de una columna angosta, lo que hacía
        # crecer esa columna en alto y perder la alineación vertical con
        # el resto de la fila.
        c1, c2, c3, c4, c5 = st.columns([1.8, 1.3, 1.1, 1.6, 3.8], vertical_alignment="center")

        subt = " · ".join(v for v in (c.get("empresa"), c.get("cuit")) if v)
        linea_nombre = f'<div style="font-weight:700; font-size:14px;">{html.escape(c["nombre"])}</div>'
        if subt:
            linea_nombre += f'<div style="font-size:12px; color:#64748b;">{html.escape(subt)}</div>'
        c1.markdown(linea_nombre, unsafe_allow_html=True)

        with c2:
            seguimiento_txt = _fmt_fecha(c.get("proximo_seguimiento"))
            dias_act = _dias_desde_actividad(c)
            actividad_txt = f"hace {dias_act} día(s)" if dias_act is not None else "sin actividad"
            linea_seguimiento = f"📅 {seguimiento_txt} · " if seguimiento_txt else ""
            st.markdown(
                f'<div style="font-size:12px; color:#64748b;">{linea_seguimiento}🕓 {actividad_txt}</div>',
                unsafe_allow_html=True,
            )
            if _seguimiento_vencido(c):
                st.markdown(f'<div style="margin-top:3px;">{_badge_vencido()}</div>', unsafe_allow_html=True)

        c3.markdown(_badge_etapa(etapa), unsafe_allow_html=True)

        with c4:
            with st.container(key=f"filacrmicons_{c['id']}"):
                if c.get("whatsapp"):
                    # st.link_button fuerza target="_blank" (pestaña nueva
                    # por cada contacto) y no deja elegirlo — mismo link
                    # armado a mano que ya usa _render_acciones_rapidas,
                    # con target="whatsapp_web" fijo para reutilizar
                    # siempre la misma pestaña.
                    st.markdown(
                        _link_button_html(
                            "💬", f"https://wa.me/{c['whatsapp'].lstrip('+')}",
                            target="whatsapp_web", title="Abrir WhatsApp",
                        ),
                        unsafe_allow_html=True,
                    )
                if c.get("email"):
                    st.link_button(
                        "", icon=":material/mail:", url=f"mailto:{c['email']}",
                        key=f"crmmail_{c['id']}", help="Abrir mail",
                    )
                if c.get("cliente_id"):
                    st.button(
                        "", icon=":material/storefront:", key=f"crmcliente_{c['id']}",
                        help="Ver ficha de Cliente / Cotizador",
                        on_click=_ver_ficha_cliente_desde_contacto, args=(c["cliente_id"],),
                    )
                if st.button("", icon=":material/delete:", key=f"crmdel_{c['id']}", help="Eliminar contacto"):
                    _dialog_eliminar_contacto_crm(c["id"], c["nombre"])

        with c5:
            if en_curso_por_cliente is not None:
                # Columna "Cliente": nada de avanzar/descartar acá — es un
                # resumen operativo de solo lectura (mismo criterio "en
                # curso" = todo lo que no sea 'Entregada' que ya usa
                # vista_clientes()). El acceso a la ficha del cliente/
                # cotizador sigue disponible vía el ícono 🏬 en c4.
                n_en_curso = en_curso_por_cliente.get(c.get("cliente_id"), 0)
                dias_act = _dias_desde_actividad(c)
                actividad_txt = f"hace {dias_act} día(s)" if dias_act is not None else "sin actividad"
                st.markdown(
                    f'<div style="font-size:12px; color:#64748b; line-height:1.6; padding-top:10px;">'
                    f'📦 {n_en_curso} importación(es) en curso<br>🕓 última actividad {actividad_txt}</div>',
                    unsafe_allow_html=True,
                )
            else:
                # 4 acciones fijas en 2 filas de a 2, mismo lenguaje visual
                # que el resto de la app ("Cambiar" en el historial del
                # Cotizador: popover con selectbox + Guardar para lo que no
                # es un solo clic). "Cotizar" y "Ver ficha →" son directas;
                # "Actualizar estado" abre el desplegable con TODAS las
                # etapas de avance válidas desde acá (crm.es_avance, no solo
                # la inmediata siguiente); "Marcar como Descartado" solo
                # tiene sentido antes de "Cliente" (Ganado).
                b1, b2 = st.columns(2)
                b1.button(
                    "Cotizar", key=f"crmcotizar_{c['id']}", type="primary", use_container_width=True,
                    on_click=_cotizar_contacto, args=(c["id"],),
                    help="Crea (o reutiliza) el cliente vinculado y abre el Cotizador con una "
                    "cotización nueva para él.",
                )
                b2.button(
                    "Ver ficha →", key=f"crmojo_{c['id']}", type="secondary", use_container_width=True,
                    on_click=_ver_ficha_contacto, args=(c["id"],),
                )
                b3, b4 = st.columns(2)
                with b3:
                    opciones_avance = [s for s in crm.PROGRESO_STAGES if crm.es_avance(etapa, s)]
                    with st.popover("Actualizar estado", use_container_width=True):
                        key_sel = f"crmnuevaetapa_{c['id']}"
                        st.selectbox(
                            "Nueva etapa", opciones_avance,
                            format_func=lambda e: crm.ETAPA_LABEL.get(e, e),
                            key=key_sel,
                        )
                        st.button(
                            "Guardar", key=f"crmguardaretapa_{c['id']}", type="primary", use_container_width=True,
                            on_click=_actualizar_etapa_contacto, args=(c["id"], key_sel),
                        )
                if etapa in crm.PROGRESO_STAGES and etapa != "Ganado":
                    b4.button(
                        "Descartado", key=f"crmdescartar_{c['id']}", use_container_width=True,
                        on_click=_marcar_descartado_boton, args=(c["id"],),
                        help="Pasa a 'Descartado' — no se puede deshacer con un clic, pero sí "
                        "recuperar cambiando la etapa a mano desde la ficha.",
                    )


def _render_seccion_contactos(contactos, key_prefix, mostrar_filtro_cotizacion=False, en_curso_por_cliente=None):
    """Un 'sector' del CRM (una etapa): lista de filas, cada una con su
    propio botón de acceso directo a la ficha."""
    if mostrar_filtro_cotizacion:
        sub_filtro = st.selectbox(
            "Filtrar por seguimiento de la cotización", list(COTIZACION_SUBFILTROS.keys()),
            key=f"crm_cot_filtro_{key_prefix}",
        )
        clave_filtro = COTIZACION_SUBFILTROS[sub_filtro]
        if clave_filtro:
            # Agregado por cliente (2 consultas SQL agrupadas, no una por
            # contacto) — ver crm.estado_comercial para el criterio.
            dias_sin_respuesta = int(db.get_setting("dias_sin_respuesta", db.SETTINGS_DEFAULTS["dias_sin_respuesta"]))
            resumen_cot = db.resumen_cotizaciones_por_cliente(dias_sin_respuesta)
            imp_por_cliente = db.contar_importaciones_por_cliente()
            contactos = [
                c for c in contactos
                if crm.estado_comercial(
                    resumen_cot.get(c.get("cliente_id")),
                    imp_por_cliente.get(c.get("cliente_id"), 0) > 0,
                ) == clave_filtro
            ]

    st.caption(f"{len(contactos)} contacto(s)")
    if not contactos:
        st.info("No hay contactos en esta sección con ese criterio.")
        return

    for c in contactos[:LIMITE_FILAS_CONTACTOS]:
        _render_fila_contacto_crm(c, en_curso_por_cliente=en_curso_por_cliente)
    if len(contactos) > LIMITE_FILAS_CONTACTOS:
        st.caption(f"+{len(contactos) - LIMITE_FILAS_CONTACTOS} más — refiná la búsqueda para verlos")


def _render_lista_contactos():
    _render_carga_rapida_contacto()
    _render_importar_contactos_masivo()

    search = st.text_input(
        "Buscar por nombre, empresa o CUIT", placeholder="🔎 Buscar por nombre, empresa o CUIT...",
        label_visibility="collapsed", key="crm_search",
    )

    todos = db.list_contacts(search=search)
    # 'Ganado' ya tiene su propia columna ('cliente', ver crm.kanban_grupo)
    # con un resumen operativo de solo lectura en vez de excluirse de la
    # vista — antes se gestionaba solo desde el módulo Clientes.
    visibles = list(todos)

    grupos = {clave: [] for clave, _ in KANBAN_COLUMNAS}
    for c in visibles:
        grupo = crm.kanban_grupo(c.get("etapa") or crm.STAGE_INICIAL)
        if grupo:
            grupos[grupo].append(c)
    # Orden automático, sin checkbox: el más urgente arriba siempre — más
    # días desde el último contacto (o desde el alta, si nunca hubo
    # actividad — _dias_desde_actividad ya combina las dos fechas) primero.
    # Reemplaza a los filtros "vencido"/"inactivo 14+" que había antes: en
    # vez de tener que activarlos a mano, el que necesita atención ya
    # aparece arriba de su columna.
    for lista in grupos.values():
        lista.sort(key=lambda c: -(_dias_desde_actividad(c) or 0))

    # Chips (st.pills), no pestañas: las etapas del embudo son un nivel de
    # navegación DENTRO de "👥 Contactos", no el mismo nivel que "Contactos /
    # Seguimiento / Analíticas" — con st.tabs ambos niveles se veían igual
    # (misma línea subrayada) y se confundían entre sí.
    etapa_sel = st.pills(
        "Etapa", options=[clave for clave, _ in KANBAN_COLUMNAS],
        format_func=lambda clave: f"{dict(KANBAN_COLUMNAS)[clave]} ({len(grupos[clave])})",
        default="nuevos", required=True, key="crm_kanban_pill", label_visibility="collapsed",
    )
    en_curso_por_cliente = db.contar_importaciones_en_curso_por_cliente() if etapa_sel == "cliente" else None
    _render_seccion_contactos(
        grupos[etapa_sel], key_prefix=etapa_sel, mostrar_filtro_cotizacion=(etapa_sel == "cotizados"),
        en_curso_por_cliente=en_curso_por_cliente,
    )


def _render_acciones_rapidas(c, autor, key_prefix):
    """Panel de accesos rápidos de mail/WhatsApp (mensaje precargado + botón
    para registrar que se contactó) y de 'marcar respuesta recibida'.
    Reutilizado en la ficha completa del contacto Y en la lista 'A
    Contactar', para poder resolver todo sin tener que abrir la ficha."""
    col_mail, col_wsp = st.columns(2)
    with col_mail:
        st.markdown("**📧 Email**")
        if c.get("email"):
            asunto = st.text_input("Asunto", value=f"Skybridge Comex — {c['nombre']}", key=f"mailasunto_{key_prefix}")
            cuerpo = st.text_area("Mensaje", value=f"Hola {c['nombre']},\n\n", key=f"mailcuerpo_{key_prefix}", height=100)
            href = f"mailto:{c['email']}?subject={quote(asunto)}&body={quote(cuerpo)}"
            st.markdown(_link_button_html("📧 Abrir mail", href), unsafe_allow_html=True)
            if st.button("✔️ Marcar contactado (email)", key=f"regmail_{key_prefix}", use_container_width=True):
                db.add_activity(c["id"], "contacto_email", "Email enviado desde el CRM.", autor)
                _avanzar_etapa_si_corresponde(c["id"], "Contactado", autor)
                _flash("Actividad registrada.")
                st.rerun()
        else:
            st.caption("Este contacto no tiene email cargado — completalo en la pestaña 'Datos'.")
    with col_wsp:
        st.markdown("**💬 WhatsApp**")
        if c.get("whatsapp"):
            mensaje = st.text_area(
                "Mensaje", value=f"Hola {c['nombre']}, te escribimos desde Skybridge Comex.",
                key=f"wspmensaje_{key_prefix}", height=100,
            )
            numero = c["whatsapp"].lstrip("+")
            href = f"https://wa.me/{numero}?text={quote(mensaje)}"
            # target fijo ("whatsapp_web", no "_blank"): el navegador reutiliza
            # siempre la misma pestaña en vez de acumular una por contacto.
            st.markdown(_link_button_html("💬 Abrir WhatsApp", href, target="whatsapp_web"), unsafe_allow_html=True)
            if st.button("✔️ Marcar contactado (WhatsApp)", key=f"regwsp_{key_prefix}", use_container_width=True):
                db.add_activity(c["id"], "contacto_whatsapp", "WhatsApp enviado desde el CRM.", autor)
                _avanzar_etapa_si_corresponde(c["id"], "Contactado", autor)
                _flash("Actividad registrada.")
                st.rerun()
        else:
            st.caption("Este contacto no tiene WhatsApp cargado — completalo en la pestaña 'Datos'.")

    if st.button("✅ Marcar respuesta recibida", key=f"resp_{key_prefix}"):
        db.add_activity(c["id"], "respuesta_recibida", "", autor)
        _flash("Respuesta registrada.")
        st.rerun()

    st.button(
        "Nueva cotización", icon=":material/add:", key=f"nuevacotacciones_{key_prefix}",
        use_container_width=True, on_click=_cotizar_contacto, args=(c["id"],),
        help="Crea (o reutiliza) el cliente vinculado en el módulo Clientes y abre el Cotizador "
        "con una cotización nueva para él.",
    )


def _avanzar_etapa_si_corresponde(contact_id, etapa_objetivo, autor):
    """Sube la etapa del contacto a etapa_objetivo únicamente si eso
    representa un avance real (ver crm.es_avance) — nunca la retrocede."""
    contacto = db.get_contact(contact_id)
    etapa_actual = contacto.get("etapa") or crm.STAGE_INICIAL
    if crm.es_avance(etapa_actual, etapa_objetivo):
        db.change_etapa_contacto(contact_id, etapa_objetivo, autor)


def _asegurar_contacto_para_cliente(cliente_id):
    """Contacto de CRM vinculado a este cliente, creándolo si hace falta —
    mismo patrón que _cotizar_contacto pero al revés (cliente → contacto en
    vez de contacto → cliente). Sin esto, "Aprobada → Ganado" y "nueva
    importación → Ganado" quedaban en silencio para cualquier cliente sin
    contacto de CRM vinculado — la mayoría hoy (ver auditoría CRM: de 4
    clientes reales, solo 1 tenía el vínculo)."""
    contacto = db.get_contact_by_cliente_id(cliente_id)
    if contacto:
        return contacto
    cliente = db.get_cliente(cliente_id)
    if not cliente:
        return None
    new_id = db.create_contact(
        cliente["nombre"], "", cliente.get("email") or "", cliente.get("telefono") or "",
        "Cliente existente", crm.STAGE_INICIAL, "", None, cliente.get("cuit") or "",
    )
    db.link_contact_cliente(new_id, cliente_id)
    return db.get_contact(new_id)


def _migrar_contactos_retroactivos():
    """Migración de una sola corrida: crea el contacto de CRM para cada
    cliente que todavía no lo tiene (db.get_contact_by_cliente_id devuelve
    None), en la etapa que corresponde a la evidencia real que ya acumuló —
    no todos arrancan en "Nuevo". Misma prioridad de señales que
    crm.estado_comercial() (importación > cotización aprobada > cualquier
    cotización > nada), colapsada a las 3 etapas que hacen falta acá.
    Devuelve [(nombre_cliente, etapa_asignada), ...] para poder verificar el
    resultado. Pensada para correrse una vez (consola o script), no para
    llamarse en cada request."""
    dias_sin_respuesta = int(db.get_setting("dias_sin_respuesta", db.SETTINGS_DEFAULTS["dias_sin_respuesta"]))
    resumen_cot = db.resumen_cotizaciones_por_cliente(dias_sin_respuesta)
    imp_por_cliente = db.contar_importaciones_por_cliente()
    autor = "Migración retroactiva"
    migrados = []
    for cliente in db.list_clientes():
        if db.get_contact_by_cliente_id(cliente["id"]):
            continue
        resumen = resumen_cot.get(cliente["id"])
        tiene_importacion = imp_por_cliente.get(cliente["id"], 0) > 0
        if tiene_importacion or (resumen and resumen.get("aprobadas")):
            etapa = "Ganado"
        elif resumen and resumen.get("total"):
            etapa = "Cotizado"
        else:
            etapa = crm.STAGE_INICIAL
        contacto = _asegurar_contacto_para_cliente(cliente["id"])
        if contacto and etapa != crm.STAGE_INICIAL:
            _avanzar_etapa_si_corresponde(contacto["id"], etapa, autor)
        migrados.append((cliente["nombre"], etapa))
    return migrados


def _marcar_ganado_por_cliente(cliente_id, autor=""):
    """Asegura el vínculo cliente↔contacto y avanza ese contacto a 'Ganado'
    (sin retroceder si ya está más adelante) — evidencia dura de que el
    cliente es real: una cotización aprobada o una importación creada.
    Devuelve el contacto (para armar mensajes de confirmación) o None si el
    cliente no existe."""
    contacto = _asegurar_contacto_para_cliente(cliente_id)
    if contacto:
        _avanzar_etapa_si_corresponde(contacto["id"], "Ganado", autor or st.session_state.get("crm_autor", ""))
    return contacto


def _agregar_nota_contacto(contact_id, texto, autor):
    """Agrega una nota y, si el contacto todavía está en la etapa inicial
    ('Nuevo'), lo avanza automáticamente a la siguiente etapa del embudo —
    una nota es evidencia de que alguien ya se ocupó del contacto. No toca
    contactos que ya están en una etapa posterior."""
    db.add_activity(contact_id, "nota", texto, autor)
    contacto = db.get_contact(contact_id)
    etapa_actual = contacto.get("etapa") or crm.STAGE_INICIAL
    if etapa_actual == crm.STAGE_INICIAL:
        siguiente = crm.siguiente_etapa(etapa_actual)
        if siguiente:
            db.change_etapa_contacto(contact_id, siguiente, autor)


def _cotizar_contacto(contact_id):
    """Callback del botón '➕ Cotizar' en la ficha de un contacto del CRM.
    Si el contacto todavía no está vinculado a un cliente del módulo
    Clientes, crea uno a partir de sus datos (nombre/empresa, email,
    WhatsApp) y lo vincula — así se reutiliza tal cual el motor de
    cotizador que ya existe para clientes, en vez de duplicarlo para
    contactos. Además, generar una cotización es evidencia de avance real
    en el embudo, así que sube la etapa a 'Cotizado' si corresponde."""
    contacto = db.get_contact(contact_id)
    cliente_id = contacto.get("cliente_id")
    if not cliente_id:
        cliente_id = db.create_cliente(
            nombre=contacto.get("empresa") or contacto["nombre"],
            email=contacto.get("email") or "",
            telefono=contacto.get("whatsapp") or "",
            notas=f"Cliente creado automáticamente desde el contacto CRM '{contacto['nombre']}'.",
        )
        db.link_contact_cliente(contact_id, cliente_id)
    autor = st.session_state.get("crm_autor", "")
    _avanzar_etapa_si_corresponde(contact_id, "Cotizado", autor)
    _cotizar_cliente(cliente_id)


def _ver_ficha_cliente_desde_contacto(cliente_id):
    st.session_state.cliente_seleccionado = cliente_id
    st.session_state.pagina_nav = "📋 Clientes"


def _render_ficha_contacto(c):
    """Vista dedicada de un contacto: datos, cambio de etapa, accesos
    rápidos a mail/WhatsApp (solo arman el link — el envío lo hace la
    persona), timeline de activity_log y edición/baja."""

    def _volver():
        st.session_state.contacto_seleccionado = None

    st.button(
        "Volver a Contactos", icon=":material/arrow_back:",
        on_click=_volver, key=f"volver_ficha_contacto_{c['id']}",
    )

    autor = st.text_input(
        "Tu nombre (autor de las actividades que cargues)", key="crm_autor",
        help="Sin login todavía — este nombre queda guardado en cada nota/cambio que cargues, "
        "para cuando se sume el multiusuario.",
    )

    with st.container(key=f"fichahdr_contacto_{c['id']}"):
        # bottom-alineado: "Etapa" trae su propio label arriba del
        # desplegable y "➕ Cotizar" no tiene label — sin esto el botón queda
        # a la altura del label en vez de a la altura del control.
        c_datos, c_etapa, c_cotizar = st.columns([2.6, 1.3, 1.1], vertical_alignment="bottom")
        with c_datos:
            linea1 = "  •  ".join([
                f"Empresa: {html.escape(c.get('empresa') or '—')}",
                f"CUIT: {html.escape(c.get('cuit') or '—')}",
            ])
            linea2 = "  •  ".join([
                f"Email: {html.escape(c.get('email') or '—')}",
                f"WhatsApp: {html.escape(c.get('whatsapp') or '—')}",
            ])
            linea3 = "  •  ".join([
                f"Origen: {html.escape(c.get('origen') or '—')}",
                f"Asignado a: {html.escape(c.get('asignado_a') or '—')}",
            ])
            st.markdown(
                f'<div class="sb-fichahdr-nombre">{html.escape(c["nombre"])}</div>'
                f'<div class="sb-fichahdr-datos">{linea1}</div>'
                f'<div class="sb-fichahdr-datos">{linea2}</div>'
                f'<div class="sb-fichahdr-datos">{linea3}</div>',
                unsafe_allow_html=True,
            )
        with c_etapa:
            etapa_actual = c.get("etapa") or crm.STAGE_INICIAL
            nueva_etapa = st.selectbox(
                "Etapa", crm.STAGES,
                index=crm.STAGES.index(etapa_actual) if etapa_actual in crm.STAGES else 0,
                format_func=lambda e: crm.ETAPA_LABEL.get(e, e),
                key=f"etapasel_{c['id']}",
            )
            if nueva_etapa != etapa_actual:
                if st.button("💾 Guardar etapa", key=f"etapaguardar_{c['id']}", type="primary", use_container_width=True):
                    db.change_etapa_contacto(c["id"], nueva_etapa, autor)
                    _flash(f"Etapa actualizada a '{crm.ETAPA_LABEL.get(nueva_etapa, nueva_etapa)}'.")
                    st.rerun()
        with c_cotizar:
            st.button(
                "➕ Cotizar", key=f"hdrcotcontacto_{c['id']}", type="secondary", use_container_width=True,
                on_click=_cotizar_contacto, args=(c["id"],),
                help="Crea (o reutiliza) el cliente vinculado en el módulo Clientes y abre una "
                "cotización nueva para él.",
            )

    tab_datos_cliente, tab_cotizaciones = st.tabs(["📁 Datos Cliente", "🧾 Cotizaciones"])

    with tab_datos_cliente:
        st.markdown("#### 📁 Datos")
        with st.form(f"edit_contacto_{c['id']}"):
            c1, c2 = st.columns(2)
            nombre = c1.text_input("Nombre de contacto", value=c["nombre"], key=f"cn_{c['id']}")
            empresa = c2.text_input("Empresa", value=c.get("empresa") or "", key=f"ce_{c['id']}")
            c3, c4 = st.columns(2)
            cuit = c3.text_input("CUIT", value=c.get("cuit") or "", key=f"ccuit_{c['id']}")
            email = c4.text_input("Email", value=c.get("email") or "", key=f"cem_{c['id']}")
            c5, c6 = st.columns(2)
            whatsapp = c5.text_input("WhatsApp", value=c.get("whatsapp") or "", key=f"cw_{c['id']}")
            origen = c6.text_input("Origen", value=c.get("origen") or "", key=f"co_{c['id']}")
            c7, c8 = st.columns(2)
            asignado_a = c7.text_input("Asignado a", value=c.get("asignado_a") or "", key=f"ca_{c['id']}")
            proximo_seguimiento = c8.date_input(
                "Próximo seguimiento", value=_parse_fecha(c.get("proximo_seguimiento")),
                format="DD/MM/YYYY", key=f"cps_{c['id']}",
            )
            b1, b2 = st.columns(2)
            if b1.form_submit_button("💾 Guardar cambios"):
                email_norm = crm.normalizar_email(email)
                if email_norm and not crm.email_valido(email_norm):
                    st.error("Ese email no parece válido.")
                elif not crm.whatsapp_valido(whatsapp):
                    st.error("Ese teléfono/WhatsApp no parece válido (revisá la cantidad de dígitos).")
                else:
                    db.update_contact(
                        c["id"], nombre, empresa, email_norm, crm.normalizar_whatsapp(whatsapp),
                        origen, asignado_a, proximo_seguimiento.isoformat() if proximo_seguimiento else None,
                        cuit,
                    )
                    st.success("Actualizado.")
                    st.rerun()
            if b2.form_submit_button("🗑️ Eliminar contacto"):
                db.delete_contact(c["id"])
                st.session_state.contacto_seleccionado = None
                st.warning("Contacto eliminado.")
                st.rerun()

        st.divider()
        st.markdown("#### 🕓 Timeline")
        with st.form(f"form_nota_{c['id']}", clear_on_submit=True):
            texto_nota = st.text_area("Nueva nota")
            if st.form_submit_button("➕ Agregar nota"):
                if texto_nota.strip():
                    _agregar_nota_contacto(c["id"], texto_nota.strip(), autor)
                    st.rerun()
        st.write("")
        eventos = db.list_activity(c["id"])
        if not eventos:
            st.caption("Sin actividad registrada todavía.")
        for ev in eventos:
            etiqueta = crm.TIPOS_ACTIVIDAD.get(ev["tipo"], ev["tipo"])
            with st.container(border=True, key=f"actev_{ev['id']}"):
                autor_txt = f" · {html.escape(ev['autor'])}" if ev.get("autor") else ""
                st.markdown(
                    f'<div style="font-size:12px; color:#64748b;">{etiqueta} · '
                    f'{_fmt_fecha_hora(ev["fecha"])}{autor_txt}</div>',
                    unsafe_allow_html=True,
                )
                if ev.get("texto"):
                    texto_mostrado = (
                        _relabel_transicion_etapas(ev["texto"]) if ev["tipo"] == "cambio_etapa" else ev["texto"]
                    )
                    st.markdown(html.escape(texto_mostrado).replace("\n", "<br>"), unsafe_allow_html=True)

        st.divider()
        st.markdown("#### 🚀 Acciones rápidas")
        st.caption(
            "Estos botones solo abren tu cliente de mail o WhatsApp con todo precargado — "
            "el envío lo hacés vos, acá no se manda nada solo."
        )
        _render_acciones_rapidas(c, autor, key_prefix=str(c["id"]))

    with tab_cotizaciones:
        cliente_vinculado = db.get_cliente(c["cliente_id"]) if c.get("cliente_id") else None
        if cliente_vinculado:
            st.caption(f"Vinculado al cliente '{cliente_vinculado['nombre']}' del módulo Clientes.")
            cots = [x for x in db.list_cotizaciones() if x["cliente_id"] == cliente_vinculado["id"]]
            if cots:
                st.dataframe(
                    pd.DataFrame(cots)[["numero", "fecha", "estado", "total_usd_civa"]],
                    hide_index=True, use_container_width=True,
                )
            else:
                st.caption("Todavía no hay cotizaciones para este cliente.")
            bc1, bc2 = st.columns(2)
            bc1.button(
                "➕ Nueva cotización", key=f"tabcotcontacto_{c['id']}", type="primary", use_container_width=True,
                on_click=_cotizar_contacto, args=(c["id"],),
            )
            bc2.button(
                "🏢 Ver ficha completa de cliente →", key=f"verclidesdecontacto_{c['id']}", use_container_width=True,
                on_click=_ver_ficha_cliente_desde_contacto, args=(cliente_vinculado["id"],),
            )
        else:
            st.caption(
                "Todavía no tiene ninguna cotización. Al generar la primera se crea automáticamente "
                "un cliente vinculado en el módulo Clientes (con el nombre, email y WhatsApp de este "
                "contacto) para poder cotizarlo, y más adelante hacerle seguimiento de importaciones "
                "y guardar su documentación."
            )
            st.button(
                "➕ Cotizar", key=f"tabcotcontacto_{c['id']}", type="primary",
                on_click=_cotizar_contacto, args=(c["id"],),
            )


def _procesar_importacion(filas, actualizar_duplicados, ejecutar, autor=""):
    """Clasifica (y opcionalmente ejecuta) la importación de una lista de
    filas ya mapeadas a campos del contacto. Se usa tanto para la vista
    previa (ejecutar=False) como para la importación real (ejecutar=True),
    así ambas quedan garantizadas 100% consistentes entre sí."""
    resultados = []
    vistos = {}  # clave (email o whatsapp normalizado) -> dict del contacto ya visto en este archivo
    for fila_cruda in filas:
        f = crm.preparar_fila_contacto(fila_cruda)
        if not f["nombre"]:
            resultados.append({**f, "estado": "Sin nombre — omitido", "contacto_id": None})
            continue

        clave = f["email"] or f["whatsapp"]
        dup = vistos.get(clave) if clave else None
        if dup is None:
            dup = db.find_contacto_duplicado(f["email"], f["whatsapp"])

        if dup:
            if ejecutar and actualizar_duplicados:
                db.update_contact(
                    dup["id"],
                    nombre=f["nombre"] or dup.get("nombre") or "",
                    empresa=f["empresa"] or dup.get("empresa") or "",
                    email=f["email"] or dup.get("email") or "",
                    whatsapp=f["whatsapp"] or dup.get("whatsapp") or "",
                    origen=f["origen"] or dup.get("origen") or "",
                    asignado_a=f["asignado_a"] or dup.get("asignado_a") or "",
                    proximo_seguimiento=dup.get("proximo_seguimiento"),
                )
            estado = "Duplicado — actualizado" if (ejecutar and actualizar_duplicados) else "Duplicado — se omite"
            resultados.append({**f, "estado": estado, "contacto_id": dup["id"]})
            if clave:
                vistos[clave] = dup
        else:
            new_id = None
            if ejecutar:
                new_id = db.create_contact(
                    f["nombre"], f["empresa"], f["email"], f["whatsapp"], f["origen"],
                    crm.STAGE_INICIAL, f["asignado_a"],
                )
                db.add_activity(new_id, "nota", "Alta por importación de archivo.", autor)
            resultados.append({**f, "estado": "Nuevo", "contacto_id": new_id})
            if clave:
                vistos[clave] = {"id": new_id, **f}
    return resultados


def _render_importar_contactos():
    st.caption(
        "Subí un Excel o CSV con tus contactos. En el paso siguiente mapeás las columnas del "
        "archivo a los campos del CRM — no hace falta que los nombres coincidan exactamente. "
        "Si reimportás el mismo archivo (o uno con contactos repetidos), los que ya existen por "
        "email o WhatsApp no se duplican."
    )
    archivo = st.file_uploader("Excel o CSV de contactos", type=["xlsx", "xls", "csv"], key="crm_import_file")
    if archivo is None:
        return

    try:
        if archivo.name.lower().endswith(".csv"):
            df = pd.read_csv(archivo)
        else:
            df = pd.read_excel(archivo)
    except Exception as e:
        st.error(f"No se pudo leer el archivo: {e}")
        return

    if df.empty:
        st.warning("El archivo no tiene filas.")
        return

    st.markdown("**Mapeo de columnas**")
    st.caption("Elegí qué columna del archivo corresponde a cada campo. 'Nombre' es obligatorio.")
    sugerido = crm.sugerir_mapeo_columnas(list(df.columns))
    opciones_col = ["— No mapear —"] + [str(c) for c in df.columns]
    etiquetas = {
        "nombre": "Nombre *", "empresa": "Empresa", "email": "Email",
        "whatsapp": "WhatsApp / Teléfono", "origen": "Origen", "asignado_a": "Asignado a",
    }
    mapeo = {}
    cols_form = st.columns(3)
    for i, (campo, etiqueta) in enumerate(etiquetas.items()):
        default = sugerido.get(campo)
        idx = opciones_col.index(str(default)) if default is not None and str(default) in opciones_col else 0
        elegido = cols_form[i % 3].selectbox(etiqueta, opciones_col, index=idx, key=f"crm_map_{campo}")
        mapeo[campo] = None if elegido == "— No mapear —" else elegido

    if not mapeo.get("nombre"):
        st.warning("Mapeá al menos la columna Nombre para poder importar.")
        return

    columnas_por_str = {str(c): c for c in df.columns}
    filas = []
    for _, row in df.iterrows():
        fila = {campo: row[columnas_por_str[col]] for campo, col in mapeo.items() if col}
        filas.append(fila)

    actualizar_duplicados = st.checkbox(
        "Actualizar los datos de los contactos que ya existan (en vez de omitirlos)",
        key="crm_import_actualizar",
    )

    resultados = _procesar_importacion(filas, actualizar_duplicados, ejecutar=False)
    df_preview = pd.DataFrame(resultados)[["nombre", "empresa", "email", "whatsapp", "origen", "asignado_a", "estado"]]
    df_preview.columns = ["Nombre", "Empresa", "Email", "WhatsApp", "Origen", "Asignado a", "Estado"]
    st.markdown("**Vista previa**")
    st.dataframe(df_preview, hide_index=True, use_container_width=True)

    n_nuevos = sum(1 for r in resultados if r["estado"] == "Nuevo")
    n_dup = sum(1 for r in resultados if r["estado"].startswith("Duplicado"))
    n_sin_nombre = sum(1 for r in resultados if r["estado"].startswith("Sin nombre"))
    st.caption(f"{n_nuevos} nuevo(s) · {n_dup} duplicado(s) · {n_sin_nombre} sin nombre (se omiten)")

    autor = st.session_state.get("crm_autor", "")
    if st.button(f"📥 Importar {n_nuevos + n_dup} contacto(s)", type="primary", disabled=(n_nuevos + n_dup == 0)):
        _procesar_importacion(filas, actualizar_duplicados, ejecutar=True, autor=autor)
        _flash(f"Importación completa: {n_nuevos} nuevo(s), {n_dup} duplicado(s).")
        st.rerun()


def _tema_font_color():
    return "#E2E8F0" if st.session_state.get("tema_oscuro") else "#0F172A"


def _tema_border_color():
    return "#334155" if st.session_state.get("tema_oscuro") else "#E2E8F0"


def _layout_transparente(fig, height=300):
    fig.update_layout(
        margin=dict(l=10, r=10, t=10, b=10), height=height,
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font_color=_tema_font_color(),
    )
    return fig


def _render_analitica_crm():
    """Analíticas del embudo — 4 KPIs (una sola card, mismo patrón que
    'Composición del costo' en el Cotizador) + 3 gráficos, todos con la
    paleta real de la app (crm.COLOR_ETAPA por etapa, naranja de marca para
    el resto) en vez de un azul/verde genérico de Plotly, y sin grillas ni
    fondo — coherente con el resto de la web en vez de "otro dashboard"."""
    contactos = db.list_contacts()
    if not contactos:
        st.info("Todavía no hay contactos cargados — no hay datos para mostrar analíticas.")
        return
    actividad_por_contacto = crm.agrupar_actividad_por_contacto(db.list_all_activity())

    conteo_etapa = crm.contactos_por_etapa(contactos)
    embudo = crm.funnel_progreso(contactos, actividad_por_contacto)
    tiempos = crm.tiempo_promedio_por_etapa(contactos, actividad_por_contacto)
    respuesta = crm.tasa_respuesta(contactos, actividad_por_contacto)
    origenes = crm.desglose_por_origen(contactos)

    total = len(contactos)
    # "Pipeline activo" = todavía en juego (ni ganado ni descartado) — la
    # cifra que de verdad hay que mirar día a día, más que el total bruto.
    pipeline_activo = sum(conteo_etapa.get(s, 0) for s in ("Nuevo", "Contactado", "Cotizado", "Negociación"))
    tasa_conversion = (conteo_etapa.get("Ganado", 0) / total) if total else 0.0

    with st.container(border=True, key="resultgroup_crm_kpis"):
        st.markdown('<div class="sb-card-title">📊 Resumen del embudo</div>', unsafe_allow_html=True)
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Contactos totales", total)
        k2.metric("Pipeline activo", pipeline_activo)
        k3.metric("Conversión a Cliente", pct(tasa_conversion))
        k4.metric("Tasa de respuesta", pct(respuesta["tasa"]))
        st.caption(f"{respuesta['respondieron']} de {respuesta['contactados']} contacto(s) contactados respondieron.")

    col_funnel, col_tiempos = st.columns([1.3, 1])
    with col_funnel:
        with st.container(border=True, key="resultgroup_crm_funnel"):
            st.markdown('<div class="sb-card-title">🔻 Embudo de ventas</div>', unsafe_allow_html=True)
            etapas_funnel = list(embudo.keys())
            fig = go.Figure(go.Funnel(
                y=[crm.ETAPA_LABEL.get(e, e) for e in etapas_funnel],
                x=list(embudo.values()),
                marker={"color": [crm.COLOR_ETAPA.get(e, "#64748B") for e in etapas_funnel]},
                textinfo="value+percent initial",
                connector={"line": {"color": _tema_border_color(), "width": 1}},
            ))
            st.plotly_chart(_layout_transparente(fig, 340), use_container_width=True)

    with col_tiempos:
        with st.container(border=True, key="resultgroup_crm_tiempos"):
            st.markdown('<div class="sb-card-title">⏱️ Tiempo promedio por etapa</div>', unsafe_allow_html=True)
            etapas_tiempo = [s for s in crm.STAGES if tiempos.get(s) is not None]
            if not etapas_tiempo:
                st.caption("Sin datos suficientes todavía.")
            else:
                valores = [tiempos[s] for s in etapas_tiempo]
                figt = go.Figure(go.Bar(
                    x=valores, y=[crm.ETAPA_LABEL.get(s, s) for s in etapas_tiempo], orientation="h",
                    marker={"color": [crm.COLOR_ETAPA.get(s, "#64748B") for s in etapas_tiempo]},
                    text=[f"{v:.1f} d" for v in valores], textposition="outside", cliponaxis=False,
                ))
                # cliponaxis=False evita que la etiqueta se corte JUSTO en el
                # borde del eje, pero el eje (oculto) igual autoescalaba su
                # rango al valor máximo de la barra más larga — sin aire
                # extra a la derecha, la etiqueta de esa barra quedaba fuera
                # del área visible del gráfico (ej. "3.1 días" se veía como
                # "3.1 d"). 25% de margen extra le da lugar al texto más
                # largo sin afectar la escala real de las barras.
                figt.update_xaxes(visible=False, range=[0, max(valores) * 1.25])
                figt.update_yaxes(autorange="reversed")
                st.plotly_chart(_layout_transparente(figt, 340), use_container_width=True)

    with st.container(border=True, key="resultgroup_crm_nuevos"):
        c_titulo, c_freq = st.columns([3, 1.3], vertical_alignment="center")
        c_titulo.markdown('<div class="sb-card-title">📈 Contactos nuevos en el tiempo</div>', unsafe_allow_html=True)
        freq_label = c_freq.radio(
            "Agrupar por", ["Semana", "Mes"], horizontal=True, key="crm_freq_nuevos", label_visibility="collapsed",
        )
        # "ME" (fin de mes), no "M" — pandas dejó de soportar el alias viejo.
        df_nuevos = crm.nuevos_por_periodo(contactos, "W" if freq_label == "Semana" else "ME")
        if df_nuevos.empty:
            st.caption("Sin datos suficientes.")
        else:
            # Etiquetas de texto + eje categórico, no un eje de fecha
            # continuo: son baldes discretos (semana/mes), no una medición
            # continua — con muy pocos períodos (a veces uno solo) un eje
            # de fecha autoescalaba a una ventana de microsegundos y el
            # gráfico quedaba en blanco. "lines+markers" para que un único
            # período siga mostrando un punto, no una línea invisible.
            fmt = "%d/%m" if freq_label == "Semana" else "%b %Y"
            etiquetas_periodo = df_nuevos["periodo"].dt.strftime(fmt)
            fig2 = go.Figure(go.Scatter(
                x=etiquetas_periodo, y=df_nuevos["cantidad"], mode="lines+markers",
                line={"color": "#E8652A", "width": 2.5, "shape": "spline"},
                marker={"color": "#E8652A", "size": 6},
                fill="tozeroy", fillcolor="rgba(232, 101, 42, 0.15)",
            ))
            fig2.update_xaxes(type="category", showgrid=False)
            fig2.update_yaxes(showgrid=False, zeroline=False, rangemode="tozero")
            st.plotly_chart(_layout_transparente(fig2, 260), use_container_width=True)

    if any(o != "Sin especificar" for o in origenes):
        with st.container(border=True, key="resultgroup_crm_origenes"):
            st.markdown('<div class="sb-card-title">🧭 Desglose por origen</div>', unsafe_allow_html=True)
            df_origen = pd.DataFrame(sorted(origenes.items(), key=lambda x: x[1]), columns=["Origen", "Cantidad"])
            fig3 = go.Figure(go.Bar(
                x=df_origen["Cantidad"], y=df_origen["Origen"], orientation="h",
                marker={"color": "#E8652A"},
                text=df_origen["Cantidad"], textposition="outside", cliponaxis=False,
            ))
            # Mismo margen extra que en "Tiempo promedio por etapa" — sin
            # esto la etiqueta de la barra más larga puede quedar cortada
            # contra el borde del gráfico.
            fig3.update_xaxes(visible=False, range=[0, df_origen["Cantidad"].max() * 1.25])
            st.plotly_chart(_layout_transparente(fig3, max(180, 40 * len(df_origen))), use_container_width=True)


def vista_crm():
    contacto_sel_id = st.session_state.get("contacto_seleccionado")
    if contacto_sel_id is not None:
        contacto = db.get_contact(contacto_sel_id)
        if contacto is None:
            st.session_state.contacto_seleccionado = None
            st.rerun()
        _render_ficha_contacto(contacto)
        return

    # A diferencia de "📋 Clientes"/"🧮 Cotizador" (título + botón que abre un
    # formulario colapsado), acá la alta rápida ya vive siempre visible
    # arriba del buscador de "👥 Contactos" (_render_carga_rapida_contacto),
    # así que no hace falta ningún botón de header para mostrarla.
    st.title("CRM")

    tab_lista, tab_analitica = st.tabs(["👥 Contactos", "📈 Analíticas"])
    with tab_lista:
        _render_lista_contactos()
    with tab_analitica:
        _render_analitica_crm()


PAGINAS = ["📊 Panel de Control", "📋 Clientes", "🧮 Cotizador", "🤝 CRM"]


def _resetear_subpaginas():
    # on_change del radio del sidebar: sólo se dispara con un clic real del
    # usuario en el menú (a diferencia de fijar st.session_state.pagina_nav
    # a mano desde otro callback, como hacen los "deep links" de cliente/
    # cotización — eso no cuenta como interacción del widget y no entra
    # acá), así que es el punto exacto para "al cambiar de sección, empezar
    # de nuevo en su página principal" sin romper esos deep links.
    st.session_state.cliente_seleccionado = None
    st.session_state.importacion_seleccionada_cliente = None
    st.session_state.cotizador_modo = "historial"
    st.session_state.cotizador_detalle_id = None
    st.session_state.contacto_seleccionado = None
    st.session_state.panel_importacion_seleccionada = None


# ---------------------------------------------------------------- MAIN
def main():
    _gate_login()

    # Navegación a la ficha de una importación desde su card en la ficha de
    # cliente: llega acá por query params (?cliente_ver=X&imp_ver=Y), no por
    # session_state seteado en un callback — es una navegación de browser
    # real (ver el link armado a mano en _render_fila_importacion_cliente),
    # así que arranca una sesión de Streamlit nueva de punta a punta y hay
    # que reconstruir el estado leyendo la URL antes de dibujar nada.
    q_imp = st.query_params.get("imp_ver")
    if q_imp:
        st.session_state.importacion_seleccionada_cliente = int(q_imp)
        q_cliente = st.query_params.get("cliente_ver")
        if q_cliente:
            st.session_state.cliente_seleccionado = int(q_cliente)
        st.session_state.pagina_nav = "📋 Clientes"
        st.query_params.clear()

    st.sidebar.markdown(
        '<div class="sb-logo">SKY<span>BRIDGE</span></div>'
        '<div class="sb-tagline">ERP · Comercio Exterior</div>',
        unsafe_allow_html=True,
    )
    # El toggle en sí solo guarda el booleano en session_state — el cambio de
    # paleta lo hace _inyectar_estilos() (ya corrido al importar el módulo,
    # arriba del todo) leyendo ese mismo valor en el siguiente rerun, que
    # Streamlit dispara automáticamente al tocar el widget.
    st.sidebar.toggle("🌙 Modo oscuro", key="tema_oscuro")

    st.sidebar.caption(f"👤 {st.session_state.usuario_autenticado['nombre']}")
    if st.sidebar.button("Cerrar sesión", use_container_width=True):
        st.session_state.usuario_autenticado = None
        st.rerun()

    # Fijamos el valor por defecto ANTES de instanciar el widget (única forma
    # válida de asignarle un valor inicial): así, cuando un callback de otra
    # página fija st.session_state.pagina_nav = "🧮 Cotizador", el radio nace
    # ya apuntando ahí en el siguiente rerun.
    st.session_state.setdefault("pagina_nav", PAGINAS[0])
    pagina = st.sidebar.radio(
        "Navegación", PAGINAS, key="pagina_nav", on_change=_resetear_subpaginas,
    )

    _mostrar_flash()

    # La ficha de una importación (abierta con "Ver detalle →" en la
    # pestaña "📦 Importaciones" de la ficha de un cliente) se resuelve ACÁ,
    # como un swap de página completo — igual que Clientes↔Cotizador↔CRM,
    # que sí reconcilian bien — en vez de como una rama más adentro de
    # vista_clientes(). Anidada ahí (misma posición de página que la ficha
    # de cliente, con una cantidad de tabs distinta), Streamlit dejaba tabs
    # viejos de la ficha de cliente pegados en el DOM al navegar entre
    # ambas — un bug de reconciliación de React confirmado con el árbol de
    # Python correcto (la función de la ficha de cliente no se ejecuta en
    # ese run), no arreglable con key=, containers ni st.empty(). Tratarla
    # como página propia lo esquiva del todo.
    imp_sel_id = st.session_state.get("importacion_seleccionada_cliente")
    if imp_sel_id is not None:
        imp = db.get_importacion(imp_sel_id)
        if imp is None:
            st.session_state.importacion_seleccionada_cliente = None
            st.rerun()
        with st.empty().container():
            _render_ficha_importacion_cliente(imp)
        return

    if pagina == "📊 Panel de Control":
        vista_panel_control()
    elif pagina == "📋 Clientes":
        vista_clientes()
    elif pagina == "🧮 Cotizador":
        vista_cotizador()
    else:
        vista_crm()


if __name__ == "__main__":
    main()
