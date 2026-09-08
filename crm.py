"""Motor del CRM: normalización de contacto, importación con dedupe, y
analíticas del embudo de ventas.

El envío de mails/WhatsApp lo sigue haciendo una persona a mano — acá solo
se arma el link (mailto: / wa.me) con todo precargado. Nada de este módulo
envía nada por sí solo.

Las analíticas del embudo (funnel, tasa de conversión, tiempo por etapa,
tasa de respuesta) se calculan siempre a partir de `activity_log` — es la
única fuente de verdad del historial de un contacto, así que nunca hay que
mantener sincronizado un contador aparte.
"""
import re
from datetime import datetime
from statistics import mean

# Etapas del embudo de ventas. Lista simple y editable (no una tabla en la
# base): agregar/renombrar una etapa acá alcanza, no hay que migrar nada.
# "Ganado" y "Perdido" son las dos salidas del embudo: Ganado es el último
# escalón del progreso (cuenta para el funnel), Perdido es una salida sin
# perder el progreso ya alcanzado antes de perderse.
STAGES = ["Nuevo", "Contactado", "Cotizado", "Negociación", "Ganado", "Perdido"]
STAGE_INICIAL = STAGES[0]

# Orden de progreso real del embudo (sin "Perdido", que es una salida, no un
# escalón de avance).
PROGRESO_STAGES = [s for s in STAGES if s != "Perdido"]

TIPOS_ACTIVIDAD = {
    "nota": "📝 Nota",
    "contacto_email": "📧 Contacto por email",
    "contacto_whatsapp": "💬 Contacto por WhatsApp",
    "respuesta_recibida": "✅ Respuesta recibida",
    "cambio_etapa": "🔀 Cambio de etapa",
    # Nuevos (Fase 5): necesidad_confirmada es el disparador de
    # Contactado→Calificado (señal propia, no se fusiona con el invoice —
    # ver caso real de sourcing antes de tener invoice). Los otros 6 son los
    # pasos del "pipeline operativo" (PIPELINE_PASOS más abajo) que se
    # muestran una vez Calificado, mientras la cotización todavía no se
    # mandó — relevamiento_confirmado y documentacion_recibida reemplazan al
    # 'nota' genérico que tenían antes esos 2 pasos, para que el contador
    # "N/7" sea confiable.
    "necesidad_confirmada": "🎯 Necesidad confirmada",
    "relevamiento_confirmado": "📋 Relevamiento del pedido",
    "proveedor_confirmado": "🏭 Proveedor confirmado",
    "productos_confirmados": "📦 Productos confirmados",
    "documentacion_recibida": "📄 Documentación recibida",
    "consulta_despachante": "🛃 Consulta a despachante",
    "consulta_forwarder": "🚢 Consulta a forwarder",
}

# Orden operativo de los 7 pasos del "pipeline" que se muestra en la card de
# un contacto Calificado sin cotización enviada todavía (ver Fase 7 — acá
# solo la lógica, todavía sin engancharse a ninguna pantalla). OJO: este
# orden es una reconstrucción mía a partir del resumen de la charla, no
# tengo el texto original del ticket con la numeración exacta — confirmar
# con Tom antes de dar por cerrada la Fase 7.
PIPELINE_PASOS = [
    "relevamiento_confirmado",
    "necesidad_confirmada",
    "proveedor_confirmado",
    "documentacion_recibida",
    "productos_confirmados",
    "consulta_despachante",
    "consulta_forwarder",
]

COLOR_ETAPA = {
    "Nuevo": "#64748B", "Contactado": "#2563EB", "Cotizado": "#D97706",
    "Negociación": "#7C3AED", "Ganado": "#16A34A", "Perdido": "#DC2626",
}

# Versión "CSS-safe" (sin acentos/espacios) de cada etapa, para usarla como
# fragmento de key de Streamlit (ej. key=f"filacrm_{ETAPA_SLUG[etapa]}_{id}")
# y así poder pintar el borde de cada card según su etapa con un selector
# CSS por prefijo, sin depender de estilos inline por fila.
ETAPA_SLUG = {
    "Nuevo": "nuevo", "Contactado": "contactado", "Cotizado": "cotizado",
    "Negociación": "negociacion", "Ganado": "ganado", "Perdido": "perdido",
}

# Capa de etiqueta de visualización: el valor guardado en contacts.etapa
# (y todas las comparaciones de este módulo/app.py — STAGES, COLOR_ETAPA,
# ETAPA_SLUG, los if etapa == "Ganado", etc.) sigue siendo literalmente
# "Ganado"/"Perdido", sin migrar nada. Solo lo que se le muestra al usuario
# pasa por acá (badges, chip del kanban, selectbox de etapa, texto de
# cambio_etapa en el timeline, analíticas) — mismo criterio que ya se usa
# en cotizaciones.contenedor/etd_eta/carrier/freetime (columnas que guardan
# algo distinto de su nombre literal). Identidad (.get(etapa, etapa)) para
# el resto de las etapas, que no cambian de nombre.
ETAPA_LABEL = {
    "Nuevo": "Prospecto",
    "Cotizado": "Calificado",
    "Negociación": "En negociación",
    "Ganado": "Cliente",
    "Perdido": "Descartado",
}

_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def normalizar_email(valor) -> str:
    return (str(valor) if valor is not None else "").strip().lower()


def email_valido(valor: str) -> bool:
    return bool(valor) and bool(_EMAIL_RE.match(valor))


def normalizar_whatsapp(valor) -> str:
    """Normaliza a formato internacional +549... que espera wa.me.

    No existe forma 100% confiable de saber la longitud del código de área
    argentino (2 a 4 dígitos según la región) sin un padrón completo, así
    que esto es una heurística razonable (saca 0 de larga distancia, 15 de
    celular, antepone 549) — el resultado siempre se muestra en el paso de
    importación para que la persona lo corrija a mano si hace falta.
    """
    if not valor:
        return ""
    digitos = re.sub(r"\D", "", str(valor))
    if not digitos:
        return ""
    if digitos.startswith("00"):
        digitos = digitos[2:]
    if digitos.startswith("54"):
        digitos = digitos[2:]
    if digitos.startswith("9"):
        digitos = digitos[1:]
    if digitos.startswith("0"):
        digitos = digitos[1:]
    if len(digitos) >= 10:
        resto, ultimos8 = digitos[:-8], digitos[-8:]
        if resto.endswith("15"):
            digitos = resto[:-2] + ultimos8
    return "+549" + digitos


def whatsapp_valido(valor) -> bool:
    """Chequeo laxo de plausibilidad (no un validador estricto de números
    argentinos — normalizar_whatsapp ya es una heurística sin garantías):
    solo bloquea valores claramente erróneos, con muy pocos o demasiados
    dígitos para ser un teléfono real (8 a 15, el máximo de un E.164).
    Vacío se considera válido (el campo es opcional)."""
    if not valor:
        return True
    digitos = re.sub(r"\D", "", str(valor))
    return 8 <= len(digitos) <= 15


def parse_dt(valor):
    if not valor:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(valor, fmt)
        except (TypeError, ValueError):
            continue
    return None


def _etapa_destino(texto_evento: str):
    """De un texto de activity_log tipo 'Anterior → Nueva' devuelve 'Nueva'."""
    if not texto_evento or "→" not in texto_evento:
        return None
    return texto_evento.split("→")[-1].strip()


def agrupar_actividad_por_contacto(actividad: list) -> dict:
    por_contacto = {}
    for ev in actividad:
        por_contacto.setdefault(ev["contact_id"], []).append(ev)
    return por_contacto


def _stage_maxima_alcanzada(contacto, eventos_cambio_etapa):
    """Índice (en PROGRESO_STAGES) más alto que el contacto alcanzó alguna
    vez: mira la etapa actual Y el historial de cambio_etapa, así un
    contacto que terminó 'Perdido' sigue contando el progreso real que hizo
    antes de perderse (en vez de aparecer como si nunca hubiera avanzado)."""
    etapa_actual = contacto.get("etapa") or STAGE_INICIAL
    maximo = PROGRESO_STAGES.index(etapa_actual) if etapa_actual in PROGRESO_STAGES else 0
    for ev in eventos_cambio_etapa:
        destino = _etapa_destino(ev.get("texto"))
        if destino in PROGRESO_STAGES:
            maximo = max(maximo, PROGRESO_STAGES.index(destino))
    return maximo


def contactos_por_etapa(contactos: list) -> dict:
    conteo = {s: 0 for s in STAGES}
    for c in contactos:
        etapa = c.get("etapa") or STAGE_INICIAL
        conteo[etapa] = conteo.get(etapa, 0) + 1
    return conteo


def funnel_progreso(contactos: list, actividad_por_contacto: dict) -> dict:
    """Para cada etapa del progreso (excluyendo Perdido), cuántos contactos
    llegaron a alcanzarla alguna vez — base del gráfico funnel y de las
    tasas de conversión entre etapas."""
    alcanzados = {s: 0 for s in PROGRESO_STAGES}
    for c in contactos:
        eventos = [e for e in actividad_por_contacto.get(c["id"], []) if e["tipo"] == "cambio_etapa"]
        maximo = _stage_maxima_alcanzada(c, eventos)
        for i in range(maximo + 1):
            alcanzados[PROGRESO_STAGES[i]] += 1
    return alcanzados


def tiempo_promedio_por_etapa(contactos: list, actividad_por_contacto: dict) -> dict:
    """Promedio de días que los contactos pasaron en cada etapa, contando
    desde su alta (etapa inicial) o desde el cambio_etapa que los llevó ahí,
    hasta el siguiente cambio de etapa (o hasta hoy, si siguen en esa etapa)."""
    duraciones = {s: [] for s in STAGES}
    ahora = datetime.now()
    for c in contactos:
        eventos = sorted(
            [e for e in actividad_por_contacto.get(c["id"], []) if e["tipo"] == "cambio_etapa"],
            key=lambda e: e.get("fecha") or "",
        )
        puntos = [(parse_dt(c.get("fecha_alta")), STAGE_INICIAL)]
        for ev in eventos:
            destino = _etapa_destino(ev.get("texto"))
            if destino in STAGES:
                puntos.append((parse_dt(ev.get("fecha")), destino))
        for i, (inicio, etapa) in enumerate(puntos):
            fin = puntos[i + 1][0] if i + 1 < len(puntos) else ahora
            if inicio and fin and fin >= inicio:
                duraciones[etapa].append((fin - inicio).total_seconds() / 86400)
    return {etapa: (mean(vals) if vals else None) for etapa, vals in duraciones.items()}


def siguiente_etapa(etapa_actual):
    """Etapa que sigue en el progreso del embudo (excluyendo 'Perdido', que
    es una salida y no un avance). None si ya está en la última etapa."""
    if etapa_actual not in PROGRESO_STAGES:
        return None
    idx = PROGRESO_STAGES.index(etapa_actual)
    return PROGRESO_STAGES[idx + 1] if idx + 1 < len(PROGRESO_STAGES) else None


def es_avance(etapa_actual, etapa_candidata):
    """True si etapa_candidata representa un progreso real respecto de
    etapa_actual dentro del embudo (nunca hacia atrás). Un contacto en
    'Perdido' (fuera del progreso) sí puede "avanzar" a cualquier etapa de
    progreso — cuenta como reactivar un contacto perdido."""
    if etapa_candidata not in PROGRESO_STAGES:
        return False
    if etapa_actual not in PROGRESO_STAGES:
        return True
    return PROGRESO_STAGES.index(etapa_candidata) > PROGRESO_STAGES.index(etapa_actual)


def kanban_grupo(etapa):
    """Mapea la etapa de un contacto a la columna del tablero kanban de
    'Contactos'. 'Negociación' tiene su propia columna ('negociacion'), y
    'Ganado' también ('cliente') — ahí las tarjetas muestran un resumen
    operativo de solo lectura en vez de los botones normales de avance,
    ver _render_fila_contacto_crm en app.py."""
    if etapa == STAGE_INICIAL:
        return "nuevos"
    if etapa == "Contactado":
        return "contactados"
    if etapa == "Cotizado":
        return "cotizados"
    if etapa == "Negociación":
        return "negociacion"
    if etapa == "Ganado":
        return "cliente"
    if etapa == "Perdido":
        return "perdidos"
    return None


def estado_comercial(resumen_cot, tiene_importacion: bool):
    """Estado comercial de un cliente, derivado de TODO su historial de
    cotizaciones (resumen_cot: una fila de
    db.resumen_cotizaciones_por_cliente(), o None si nunca cotizó) más si
    ya tiene alguna importación real — no solo la cotización más reciente.
    Con cientos de cotizaciones acumuladas por cliente, mirar nada más la
    última daba una foto parcial (ver nota en resumen_cotizaciones_por_cliente).

    Prioridad de señales, de más a menos fuerte:
    'ganado' (ya tiene una importación — evidencia dura, pisa cualquier
    otra señal) > 'aprobada_pendiente' (cotización aprobada, todavía sin
    importación creada) > 'enviada' (hay al menos una enviada — sin
    distinguir vigente/vencida por fecha: esa distinción la da ahora el
    badge de inactividad del contacto, no el estado de la cotización) >
    'a_revisar' (nada resuelto todavía: borradores y/o rechazadas, pero no
    el 100% rechazadas) > 'rechazada' (TODAS sus cotizaciones están
    rechazadas, sin ninguna pendiente ni aprobada — candidato a marcar
    como Descartado)."""
    if tiene_importacion:
        return "ganado"
    if not resumen_cot or not resumen_cot.get("total"):
        return None
    if resumen_cot.get("aprobadas"):
        return "aprobada_pendiente"
    if resumen_cot.get("enviadas"):
        return "enviada"
    if resumen_cot.get("rechazadas") == resumen_cot.get("total"):
        return "rechazada"
    return "a_revisar"


def pipeline_operativo(tipos_logueados: set):
    """Progreso del "pipeline operativo" (7 pasos, PIPELINE_PASOS) de un
    contacto Calificado que todavía no mandó cotización — pura lógica
    (activity_log → paso), sin tocar ninguna pantalla todavía (se engancha
    recién en la Fase 7). tipos_logueados: set de c["tipo"] de
    activity_log para ese contacto (cualquier cantidad de veces cada uno,
    acá solo importa si pasó o no al menos una vez).

    Devuelve (n, etiqueta): n = pasos completados (0-7); etiqueta = nombre
    del PRÓXIMO paso pendiente (lo que falta para seguir), o el del último
    paso completado si ya están los 7. None si no hay ningún paso
    completado todavía (contacto recién Calificado, sin nada logueado)."""
    completados = [p for p in PIPELINE_PASOS if p in tipos_logueados]
    n = len(completados)
    if n == 0:
        return None
    if n == len(PIPELINE_PASOS):
        return n, TIPOS_ACTIVIDAD[PIPELINE_PASOS[-1]]
    siguiente = next(p for p in PIPELINE_PASOS if p not in tipos_logueados)
    return n, TIPOS_ACTIVIDAD[siguiente]


def tasa_respuesta(contactos: list, actividad_por_contacto: dict) -> dict:
    contactados = respondieron = 0
    for c in contactos:
        tipos = {e["tipo"] for e in actividad_por_contacto.get(c["id"], [])}
        if tipos & {"contacto_email", "contacto_whatsapp"}:
            contactados += 1
            if "respuesta_recibida" in tipos:
                respondieron += 1
    return {
        "contactados": contactados,
        "respondieron": respondieron,
        "tasa": (respondieron / contactados) if contactados else 0.0,
    }


def desglose_por_origen(contactos: list) -> dict:
    conteo = {}
    for c in contactos:
        origen = (c.get("origen") or "").strip() or "Sin especificar"
        conteo[origen] = conteo.get(origen, 0) + 1
    return conteo


def nuevos_por_periodo(contactos: list, freq: str = "W"):
    """DataFrame [periodo, cantidad] de altas agrupadas por semana ('W') o
    mes ('ME', fin de mes — el alias viejo 'M' ya no existe en pandas),
    para el gráfico de contactos nuevos en el tiempo."""
    import pandas as pd
    if not contactos:
        return pd.DataFrame(columns=["periodo", "cantidad"])
    df = pd.DataFrame(contactos)
    df["fecha_alta_dt"] = pd.to_datetime(df["fecha_alta"], errors="coerce")
    df = df.dropna(subset=["fecha_alta_dt"])
    if df.empty:
        return pd.DataFrame(columns=["periodo", "cantidad"])
    agrupado = df.set_index("fecha_alta_dt").resample(freq).size().reset_index(name="cantidad")
    agrupado.columns = ["periodo", "cantidad"]
    return agrupado


# ---------------------------------------------------------------- IMPORTACIÓN

# Campo destino -> palabras clave para adivinar la columna correspondiente
# en un Excel/CSV subido (comparación case-insensitive, sin acentos).
CAMPOS_IMPORTABLES = {
    "nombre": ["nombre", "contacto", "name"],
    "empresa": ["empresa", "compania", "company", "razon social"],
    "email": ["email", "mail", "correo"],
    "whatsapp": ["whatsapp", "telefono", "celular", "phone", "tel"],
    "origen": ["origen", "fuente", "source"],
    "asignado_a": ["asignado", "responsable", "vendedor", "owner"],
    # Fase 6 — ficha ampliada, también importables desde Excel/CSV.
    "provincia": ["provincia", "province", "estado"],
    "localidad": ["localidad", "ciudad", "city"],
    "rubro": ["rubro", "sector", "industria", "industry"],
    "cargo_contacto": ["cargo", "puesto", "rol", "position", "title"],
}


def _sin_acentos(t: str) -> str:
    tabla = str.maketrans("áéíóúñÁÉÍÓÚÑ", "aeiounAEIOUN")
    return t.translate(tabla)


def sugerir_mapeo_columnas(columnas: list) -> dict:
    """Adivina, para cada campo del contacto, cuál columna del archivo
    subido le corresponde — para no obligar a mapear todo a mano cuando los
    nombres de columna ya son razonablemente parecidos."""
    normalizadas = {col: _sin_acentos(str(col)).strip().lower() for col in columnas}
    sugerido = {}
    for campo, claves in CAMPOS_IMPORTABLES.items():
        match = next((col for col, norm in normalizadas.items() if any(k in norm for k in claves)), None)
        sugerido[campo] = match
    return sugerido


def preparar_fila_contacto(fila: dict) -> dict:
    """Normaliza una fila ya mapeada (claves = campos del contacto) para
    insertar/comparar: email en minúsculas, whatsapp en formato +549..."""
    return {
        "nombre": (fila.get("nombre") or "").strip(),
        "empresa": (fila.get("empresa") or "").strip(),
        "email": normalizar_email(fila.get("email")),
        "whatsapp": normalizar_whatsapp(fila.get("whatsapp")),
        "origen": (fila.get("origen") or "").strip(),
        "asignado_a": (fila.get("asignado_a") or "").strip(),
        "provincia": (fila.get("provincia") or "").strip(),
        "localidad": (fila.get("localidad") or "").strip(),
        "rubro": (fila.get("rubro") or "").strip(),
        "cargo_contacto": (fila.get("cargo_contacto") or "").strip(),
    }
