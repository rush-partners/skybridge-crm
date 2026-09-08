"""Capa de acceso a datos para Skybridge ERP/CRM.

La base de datos vive en Turso (servicio de base de datos en la nube,
compatible con SQLite vía SQL sobre HTTP) — no como archivo local. Así los
datos sobreviven un reinicio o redeploy del hosting, que es justo lo que un
archivo .db suelto en el disco NO garantiza en un plan gratuito."""
import http.client
import json
import os
import sqlite3
import tempfile
import threading
from pathlib import Path
from urllib.parse import urlsplit
from datetime import datetime, timedelta

import bcrypt
import streamlit as st
import turso_serverless
from turso_serverless.session import ProtocolError


def _reintentar_stream_perdido(metodo_original):
    """Turso cierra del lado del servidor cualquier stream que quedó
    inactivo un rato (unos segundos/minutos sin consultas) — típico
    después de un redeploy, o simplemente cuando el usuario tarda en leer
    la pantalla antes del siguiente click. La conexión igual sigue viva
    (el socket keep-alive de _post_keepalive no se cortó), pero el
    servidor ya no reconoce el "baton" que arrastrábamos, y responde con
    "stream not found" — eso se propagaba tal cual hasta la pantalla del
    usuario como un error fatal, obligando a recargar TODO el navegador
    (perdiendo el login) para que se abriera una sesión nueva.

    Es seguro reintentar UNA vez en este caso puntual: cuando el server
    rechaza el baton, rechaza el pedido ANTES de tocar la base — no llegó
    a ejecutar nada, así que no hay riesgo de duplicar un INSERT/UPDATE
    (a diferencia de un timeout de red genérico, donde no se sabe si el
    pedido llegó a aplicarse o no, y por eso ESE caso no se reintenta acá).
    Para cuando este wrapper ve la excepción, Session._post ya dejó
    self._baton en None (via _reset_stream, ver session.py de la
    librería), así que el reintento sale pidiendo un stream nuevo solo."""
    def envoltorio(self, *args, **kwargs):
        try:
            return metodo_original(self, *args, **kwargs)
        except ProtocolError as e:
            if "stream not found" not in str(e).lower():
                raise
            return metodo_original(self, *args, **kwargs)
    return envoltorio


def _instalar_keepalive_turso():
    """turso_serverless.Session (la librería instalada, único release
    disponible: 0.1.0) tiene tres problemas de rendimiento/robustez que se
    parchean acá porque no hay forma de arreglarlos sin editar
    site-packages (que se pisa en cada `pip install -r requirements.txt`):

    1. _post abre una conexión TCP+TLS NUEVA en cada consulta HTTP (usa
       urllib.request.urlopen, que no reutiliza sockets) — cada
       SELECT/INSERT individual pagaba un handshake completo contra Turso
       (~1-3s cada uno, medido), y una pantalla con 5-10 consultas se
       sentía inusable (20-30s). Se reemplaza el transporte por una
       conexión HTTP keep-alive guardada en la sesión, reconectando sola
       si el socket se cae.

    2. Streamlit no cancela un rerun en curso al instante: si el usuario
       interactúa de nuevo mientras el script anterior todavía está
       terminando, los dos hilos corren un rato en simultáneo — y como
       get_connection() guarda UNA sesión por sesión de navegador (no por
       hilo, justamente para que sobreviva entre reruns), ambos hilos
       pueden terminar usando la MISMA sesión (mismo socket HTTP y mismo
       "baton", el token de continuidad del stream de libSQL) al mismo
       tiempo. Sin serializar, eso rompe de dos formas: el socket
       reventaba con "ResponseNotReady", y aun arreglando eso, dos
       pedidos que leen el mismo baton antes de que el primero lo
       actualice hacen que el servidor rechace el segundo con
       "generation mismatch". Por eso el lock (RLock: execute_stmt y
       execute_pipeline pueden llamarse entre sí en el mismo hilo, ej. en
       _refresh_autocommit) envuelve el método entero — desde que se lee
       el baton hasta que se actualiza con la respuesta — y no sólo la
       llamada HTTP.

    3. Un stream que Turso ya cerró por inactividad ("stream not found")
       tumbaba toda la sesión del usuario en vez de reconectar sola — ver
       _reintentar_stream_perdido arriba."""
    from turso_serverless.session import Session, ProtocolError

    def _post_keepalive(self, path, body):
        url = f"{self._base_url}{path}"
        partes = urlsplit(url)
        host, port = partes.hostname, partes.port or 443
        datos = json.dumps(body, allow_nan=False).encode("utf-8")
        cabeceras = self._headers()
        cabeceras["Connection"] = "keep-alive"

        def _intentar(conn):
            conn.request("POST", partes.path or "/", body=datos, headers=cabeceras)
            resp = conn.getresponse()
            return resp, resp.read()

        conn = getattr(self, "_http_conn", None)
        if conn is None or getattr(self, "_http_conn_host", None) != (host, port):
            conn = http.client.HTTPSConnection(host, port, timeout=30)
            self._http_conn, self._http_conn_host = conn, (host, port)

        try:
            resp, crudo = _intentar(conn)
        except (http.client.HTTPException, OSError):
            # El socket reusado puede haber sido cerrado por el servidor
            # (idle timeout) — se reintenta una vez con una conexión nueva.
            try:
                conn.close()
            except Exception:
                pass
            conn = http.client.HTTPSConnection(host, port, timeout=30)
            self._http_conn = conn
            try:
                resp, crudo = _intentar(conn)
            except (http.client.HTTPException, OSError) as e:
                self._reset_stream()
                raise ProtocolError(f"request to {url} failed: {e!r}") from None

        if resp.status >= 400:
            self._reset_stream()
            mensaje = None
            try:
                parseado = json.loads(crudo.decode("utf-8", errors="replace"))
                if isinstance(parseado, dict):
                    for clave in ("error", "message"):
                        if isinstance(parseado.get(clave), str):
                            mensaje = parseado[clave]
                            break
            except ValueError:
                pass
            if mensaje is not None:
                raise ProtocolError(f"HTTP status {resp.status}: {mensaje}") from None
            raise ProtocolError(f"HTTP status {resp.status}") from None
        return crudo

    Session._post = _post_keepalive

    def _con_lock(metodo_original):
        def envoltorio(self, *args, **kwargs):
            lock = getattr(self, "_stream_lock", None)
            if lock is None:
                lock = threading.RLock()
                self._stream_lock = lock
            with lock:
                return metodo_original(self, *args, **kwargs)
        return envoltorio

    Session.execute_stmt = _con_lock(_reintentar_stream_perdido(Session.execute_stmt))
    Session.execute_pipeline = _con_lock(_reintentar_stream_perdido(Session.execute_pipeline))


_instalar_keepalive_turso()


def _secreto(nombre: str):
    """Busca primero en st.secrets (.streamlit/secrets.toml en local, panel
    de Secrets en Streamlit Community Cloud) y si no está, en el entorno —
    así funciona igual sin importar dónde corra la app."""
    try:
        if nombre in st.secrets:
            return st.secrets[nombre]
    except Exception:
        pass
    return os.environ.get(nombre)


TURSO_DATABASE_URL = _secreto("TURSO_DATABASE_URL")
TURSO_AUTH_TOKEN = _secreto("TURSO_AUTH_TOKEN")

DATA_DIR = Path(os.environ.get("DATA_DIR", Path(__file__).parent))
DATA_DIR.mkdir(parents=True, exist_ok=True)
BACKUPS_DIR = DATA_DIR / "backups"
MAX_BACKUPS = 30

SCHEMA = """
CREATE TABLE IF NOT EXISTS clientes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT NOT NULL,
    cuit TEXT,
    email TEXT,
    telefono TEXT,
    direccion TEXT,
    rubro TEXT,
    productos_interes TEXT,
    notas TEXT,
    fecha_alta TEXT DEFAULT (datetime('now','localtime'))
);

CREATE TABLE IF NOT EXISTS cotizaciones (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    numero TEXT,
    cliente_id INTEGER REFERENCES clientes(id) ON DELETE SET NULL,
    fecha TEXT DEFAULT (datetime('now','localtime')),
    detalle_pedido TEXT,
    origen_cond_venta TEXT,
    contenedor TEXT,
    etd_eta TEXT,
    carrier TEXT,
    freetime TEXT,
    estado TEXT DEFAULT 'Borrador',
    costo_financiero_pct REAL DEFAULT 0.025,
    seguro_pct REAL DEFAULT 0.003,
    tc_tributos REAL DEFAULT 0,
    tc_operativos REAL DEFAULT 0,
    tc_venta_ars REAL DEFAULT 0,
    arancel_sim REAL DEFAULT 10,
    total_usd_civa REAL,
    total_ars_civa REAL,
    total_usd_sviva REAL,
    total_ars_sviva REAL,
    venta_total_usd REAL,
    ganancia_bruta_usd REAL,
    rentabilidad_pct REAL,
    creado_en TEXT DEFAULT (datetime('now','localtime')),
    actualizado_en TEXT DEFAULT (datetime('now','localtime'))
);

CREATE TABLE IF NOT EXISTS cotizacion_productos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    cotizacion_id INTEGER NOT NULL REFERENCES cotizaciones(id) ON DELETE CASCADE,
    orden INTEGER,
    descripcion TEXT,
    ncm TEXT,
    fob_unit REAL DEFAULT 0,
    cantidad REAL DEFAULT 0,
    fob_decl_unit REAL DEFAULT 0,
    peso_vol REAL DEFAULT 0,
    pct_derechos REAL DEFAULT 0,
    pct_tasa_estadistica REAL DEFAULT 0.03,
    pct_antidumping REAL DEFAULT 0,
    pct_iva REAL DEFAULT 0.21,
    pct_iva_adicional REAL DEFAULT 0,
    pct_ganancias REAL DEFAULT 0,
    pct_iibb REAL DEFAULT 0.05,
    margen_pct REAL DEFAULT 0.3,
    pv_final_usd REAL DEFAULT 0,
    pv_mercado_usd REAL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS cotizacion_gastos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    cotizacion_id INTEGER NOT NULL REFERENCES cotizaciones(id) ON DELETE CASCADE,
    orden INTEGER,
    concepto TEXT,
    moneda TEXT DEFAULT 'USD',
    monto REAL DEFAULT 0,
    prorrateo TEXT DEFAULT 'FOB',
    iva_incluido TEXT DEFAULT 'NO',
    pct_iva REAL DEFAULT 0.21
);

CREATE TABLE IF NOT EXISTS importaciones (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    numero TEXT,
    cliente_id INTEGER NOT NULL REFERENCES clientes(id) ON DELETE CASCADE,
    cotizacion_id INTEGER REFERENCES cotizaciones(id) ON DELETE SET NULL,
    estado TEXT DEFAULT 'En producción',
    producto TEXT,
    proveedor_nombre TEXT,
    proveedor_contacto TEXT,
    proveedor_pais TEXT,
    contenedor TEXT,
    carrier TEXT,
    tipo_envio TEXT,
    formato_envio TEXT,
    etd TEXT,
    eta TEXT,
    deposito_destino TEXT,
    notas_transito TEXT,
    fecha_llegada_puerto TEXT,
    fecha_oficializacion TEXT,
    fecha_entrega TEXT,
    pago_proveedor_usd REAL DEFAULT 0,
    costo_financiero_usd REAL DEFAULT 0,
    vep_tributos_usd REAL DEFAULT 0,
    flete_usd REAL DEFAULT 0,
    gastos_operativos_usd REAL DEFAULT 0,
    intervenciones_usd REAL DEFAULT 0,
    otros_pagos_usd REAL DEFAULT 0,
    creado_en TEXT DEFAULT (datetime('now','localtime')),
    actualizado_en TEXT DEFAULT (datetime('now','localtime'))
);

CREATE TABLE IF NOT EXISTS importacion_documentos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    importacion_id INTEGER NOT NULL REFERENCES importaciones(id) ON DELETE CASCADE,
    categoria TEXT NOT NULL,
    nombre_archivo TEXT NOT NULL,
    ruta_archivo TEXT NOT NULL,
    subido_en TEXT DEFAULT (datetime('now','localtime'))
);

CREATE TABLE IF NOT EXISTS cliente_documentos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    cliente_id INTEGER NOT NULL REFERENCES clientes(id) ON DELETE CASCADE,
    nombre_archivo TEXT NOT NULL,
    ruta_archivo TEXT NOT NULL,
    tamano_bytes INTEGER DEFAULT 0,
    subido_en TEXT DEFAULT (datetime('now','localtime'))
);

CREATE TABLE IF NOT EXISTS contacts (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT NOT NULL,
    empresa TEXT,
    cuit TEXT,
    email TEXT,
    whatsapp TEXT,
    origen TEXT,
    etapa TEXT DEFAULT 'Nuevo',
    asignado_a TEXT,
    proximo_seguimiento TEXT,
    fecha_alta TEXT DEFAULT (datetime('now','localtime')),
    cliente_id INTEGER REFERENCES clientes(id) ON DELETE SET NULL
);

-- Única fuente de verdad del historial de un contacto: timeline en la ficha
-- Y datos crudos de las analíticas del embudo (funnel, conversión, tiempo
-- por etapa, tasa de respuesta) salen todos de acá, no de una tabla aparte.
CREATE TABLE IF NOT EXISTS activity_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    contact_id INTEGER NOT NULL REFERENCES contacts(id) ON DELETE CASCADE,
    tipo TEXT NOT NULL,
    fecha TEXT DEFAULT (datetime('now','localtime')),
    autor TEXT,
    texto TEXT
);

-- Parámetros generales editables desde la UI (Panel de Control >
-- Configuración), para que ningún valor de estos quede fijo en el código.
CREATE TABLE IF NOT EXISTS settings (
    key TEXT PRIMARY KEY,
    value TEXT
);

CREATE INDEX IF NOT EXISTS idx_activity_contact ON activity_log(contact_id);
CREATE INDEX IF NOT EXISTS idx_cotizaciones_cliente ON cotizaciones(cliente_id);
CREATE INDEX IF NOT EXISTS idx_importaciones_cliente ON importaciones(cliente_id);
CREATE INDEX IF NOT EXISTS idx_cotizacion_productos_cot ON cotizacion_productos(cotizacion_id);
CREATE INDEX IF NOT EXISTS idx_cotizacion_gastos_cot ON cotizacion_gastos(cotizacion_id);

CREATE TABLE IF NOT EXISTS usuarios (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT UNIQUE NOT NULL,
    password_hash TEXT NOT NULL,
    nombre TEXT NOT NULL,
    activo INTEGER DEFAULT 1,
    creado_en TEXT DEFAULT (datetime('now','localtime'))
);
"""

# Valores por defecto de `settings`: se insertan una sola vez (INSERT OR
# IGNORE) en init_db(), así la fila siempre existe y la UI de Configuración
# no tiene que manejar el caso "todavía no se guardó nada".
SETTINGS_DEFAULTS = {
    "dias_sin_respuesta": "3",
}

DEFAULT_GASTOS = [
    ("Flete", "USD", "PESO", "NO", 0.0),
    ("Seguro", "USD", "FOB", "NO", 0.21),
    ("Gastos EXW", "USD", "FOB", "NO", 0.0),
    ("Gastos locales", "USD", "FOB", "NO", 0.21),
    ("Terminal portuaria", "USD", "PESO", "NO", 0.21),
    ("Depósito fiscal", "USD", "PESO", "NO", 0.21),
    ("Gastos adicionales", "USD", "PESO", "NO", 0.21),
    ("Gestiones terceros organismos", "USD", "PESO", "NO", 0.21),
    ("Traslado interno (en ARS)", "ARS", "PESO", "NO", 0.21),
    ("Honorarios despachante", "USD", "FOB", "NO", 0.21),
    ("Honorarios consultora", "USD", "FOB", "NO", 0.21),
]


_TIMEOUT_LOCK_CONEXION = 5


def get_connection():
    """Cada función de este archivo hace get_connection() y al final
    conn.close() — con Turso, abrir conexión nueva por cada consulta suma un
    viaje de red (ida y vuelta a Virginia) por cada una, y una pantalla que
    antes hacía 5-10 consultas locales instantáneas ahora se sentía lenta.
    Para no tener que tocar las ~30 funciones de abajo, get_connection()
    guarda UNA conexión en st.session_state (NO en threading.local: Streamlit
    corre cada rerun del script — cada click, cada tecla — en un hilo
    NUEVO, así que threading.local no sobrevive entre interacciones y
    terminaba abriendo una conexión distinta en cada una, sin ninguna
    mejora real; st.session_state en cambio persiste mientras dure la
    sesión del navegador, que es la unidad correcta acá) y la reutiliza en
    cada llamada dentro de esa sesión. Con esto, sólo la primera consulta de
    toda la sesión del usuario paga el viaje de red de abrir conexión — el
    resto son mucho más rápidas, incluso cambiando de pantalla.

    Como Streamlit puede solapar dos reruns (uno terminando justo cuando el
    siguiente ya arrancó — típico de un doble click en "Guardar"), y ambos
    comparten esta misma conexión, una función de varias sentencias (ej.
    save_cotizacion: UPDATE + varios INSERT + commit) puede quedar a medio
    terminar cuando la OTRA corrida mete su propia sentencia en el medio —
    eso rompe la transacción ("cannot start/commit a transaction..."). Por
    eso conn.close() (que cada función de abajo ya llama al final, sea
    lectura o escritura) queda pisado para liberar un lock que get_connection()
    toma ACÁ, antes de devolver la conexión: así cada función de este
    archivo pasa a ser, sin tocarla, una sección crítica completa. Si algo
    corta una función a mitad de camino sin llegar a su conn.close() (una
    excepción no atrapada), el lock quedaría trabado para siempre — por
    eso el acquire tiene timeout: pasados _TIMEOUT_LOCK_CONEXION segundos
    sin poder tomarlo, se abandona esa conexión y se abre una nueva en vez
    de dejar la sesión colgada."""
    conn = st.session_state.get("_db_conn")
    if conn is not None:
        if not conn._sb_lock.acquire(timeout=_TIMEOUT_LOCK_CONEXION):
            conn = None
    if conn is None:
        conn = turso_serverless.connect(TURSO_DATABASE_URL, auth_token=TURSO_AUTH_TOKEN)
        conn.row_factory = turso_serverless.Row
        conn._sb_lock = threading.RLock()
        conn._sb_lock.acquire()
        conn.close = conn._sb_lock.release
        conn.execute("PRAGMA foreign_keys = ON")
        st.session_state["_db_conn"] = conn
    return conn


DBError = turso_serverless.Error


def hay_usuarios():
    conn = get_connection()
    n = conn.execute("SELECT COUNT(*) AS n FROM usuarios").fetchone()["n"]
    conn.close()
    return n > 0


def create_usuario(username, password, nombre):
    conn = get_connection()
    pw_hash = bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
    cur = conn.execute(
        "INSERT INTO usuarios (username, password_hash, nombre) VALUES (?, ?, ?)",
        (username, pw_hash, nombre),
    )
    conn.commit()
    nuevo_id = cur.lastrowid
    conn.close()
    return nuevo_id


INTENTOS_MAX = 5
BLOQUEO_MINUTOS = 15


def verificar_login(username, password):
    """Devuelve (usuario_o_None, mensaje_o_None). Bloquea la cuenta
    BLOQUEO_MINUTOS después de INTENTOS_MAX fallos seguidos, para que no
    se puedan probar contraseñas sin límite."""
    conn = get_connection()
    row = conn.execute(
        "SELECT * FROM usuarios WHERE username = ? AND activo = 1", (username,)
    ).fetchone()

    if row and row["bloqueado_hasta"]:
        bloqueado_hasta = datetime.strptime(row["bloqueado_hasta"], "%Y-%m-%d %H:%M:%S")
        if datetime.now() < bloqueado_hasta:
            conn.close()
            minutos_restantes = max(1, int((bloqueado_hasta - datetime.now()).total_seconds() // 60) + 1)
            return None, f"Demasiados intentos fallidos. Probá de nuevo en {minutos_restantes} minuto(s)."

    if not row or not bcrypt.checkpw(password.encode("utf-8"), row["password_hash"].encode("utf-8")):
        if row:
            nuevos_intentos = (row["intentos_fallidos"] or 0) + 1
            if nuevos_intentos >= INTENTOS_MAX:
                bloqueo = (datetime.now() + timedelta(minutes=BLOQUEO_MINUTOS)).strftime("%Y-%m-%d %H:%M:%S")
                conn.execute(
                    "UPDATE usuarios SET intentos_fallidos = 0, bloqueado_hasta = ? WHERE id = ?",
                    (bloqueo, row["id"]),
                )
            else:
                conn.execute(
                    "UPDATE usuarios SET intentos_fallidos = ? WHERE id = ?", (nuevos_intentos, row["id"]),
                )
            conn.commit()
        conn.close()
        return None, "Usuario o contraseña incorrectos."

    if row["intentos_fallidos"] or row["bloqueado_hasta"]:
        conn.execute(
            "UPDATE usuarios SET intentos_fallidos = 0, bloqueado_hasta = NULL WHERE id = ?", (row["id"],),
        )
        conn.commit()
    conn.close()
    return {"id": row["id"], "username": row["username"], "nombre": row["nombre"]}, None


def list_usuarios():
    conn = get_connection()
    rows = conn.execute("SELECT * FROM usuarios ORDER BY nombre").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def set_usuario_activo(usuario_id, activo):
    conn = get_connection()
    conn.execute("UPDATE usuarios SET activo = ? WHERE id = ?", (int(activo), usuario_id))
    conn.commit()
    conn.close()


def _nombres_tablas(conn) -> list[str]:
    return [
        r["name"] for r in conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'"
        ).fetchall()
    ]


def exportar_backup_sqlite() -> bytes:
    """Arma un archivo .db (SQLite clásico) con el contenido ACTUAL de
    Turso, tabla por tabla. El resultado es un archivo normal que se puede
    abrir con cualquier visor de SQLite, y que restaurar_desde_sqlite_bytes()
    sabe leer para restaurar — así el botón de backup y el de restaurar
    siguen hablando el mismo formato de siempre."""
    conn_turso = get_connection()
    tablas = _nombres_tablas(conn_turso)
    with tempfile.NamedTemporaryFile(suffix=".db", delete=False) as tmp:
        ruta_tmp = Path(tmp.name)
    try:
        conn_local = sqlite3.connect(ruta_tmp)
        conn_local.row_factory = sqlite3.Row
        conn_local.executescript(SCHEMA)
        # SCHEMA por si sola no alcanza: las columnas agregadas después de la
        # creación original de cada tabla (COLUMNAS_NUEVAS, vía ALTER TABLE en
        # producción) no están ahí — sin este paso, el INSERT de abajo falla
        # apenas Turso tiene una fila con alguna de esas columnas.
        _migrar_columnas_faltantes(conn_local)
        conn_local.commit()
        for tabla in tablas:
            filas = conn_turso.execute(f"SELECT * FROM {tabla}").fetchall()
            if not filas:
                continue
            columnas = filas[0].keys()
            placeholders = ",".join("?" for _ in columnas)
            conn_local.executemany(
                f"INSERT INTO {tabla} ({','.join(columnas)}) VALUES ({placeholders})",
                [tuple(f[c] for c in columnas) for f in filas],
            )
        conn_local.commit()
        conn_local.close()
        conn_turso.close()
        return ruta_tmp.read_bytes()
    finally:
        ruta_tmp.unlink(missing_ok=True)


def restaurar_desde_sqlite_bytes(contenido: bytes):
    """Reemplaza TODO el contenido de Turso por el de un archivo .db subido
    (por ejemplo, un backup descargado antes, o la base real la primera
    vez). Borra cada tabla y vuelve a insertar fila por fila. Se usa desde
    el login inicial (primera carga, sin usuarios todavía) y desde
    Configuración ▸ Restaurar backup."""
    with tempfile.NamedTemporaryFile(suffix=".db", delete=False) as tmp:
        tmp.write(contenido)
        ruta_tmp = Path(tmp.name)
    try:
        conn_local = sqlite3.connect(ruta_tmp)
        conn_local.row_factory = sqlite3.Row
        tablas = _nombres_tablas(conn_local)

        conn_turso = get_connection()
        conn_turso.execute("PRAGMA foreign_keys = OFF")
        for tabla in tablas:
            conn_turso.execute(f"DELETE FROM {tabla}")
        for tabla in tablas:
            filas = conn_local.execute(f"SELECT * FROM {tabla}").fetchall()
            if not filas:
                continue
            columnas = filas[0].keys()
            placeholders = ",".join("?" for _ in columnas)
            for f in filas:
                conn_turso.execute(
                    f"INSERT INTO {tabla} ({','.join(columnas)}) VALUES ({placeholders})",
                    tuple(f[c] for c in columnas),
                )
        conn_turso.commit()
        conn_turso.execute("PRAGMA foreign_keys = ON")
        conn_turso.close()
        conn_local.close()
    finally:
        ruta_tmp.unlink(missing_ok=True)


def backup_antes_de_borrar(motivo: str):
    """Snapshot de seguridad ANTES de una baja irreversible (cotización,
    importación o cliente — lo real no puede depender de un solo clic sin
    red de contención). Queda en disco local, útil dentro de la sesión
    actual del servidor. Rota sola: conserva las últimas MAX_BACKUPS, borra
    el resto."""
    BACKUPS_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    destino = BACKUPS_DIR / f"skybridge_{timestamp}_{motivo}.db"
    destino.write_bytes(exportar_backup_sqlite())
    backups = sorted(BACKUPS_DIR.glob("skybridge_*.db"), key=lambda p: p.stat().st_mtime)
    for viejo in backups[:-MAX_BACKUPS]:
        viejo.unlink()


# Columnas agregadas después de la creación original de la tabla. SQLite no
# actualiza una tabla ya existente cuando se edita el CREATE TABLE de arriba,
# así que en cada arranque se agregan (ALTER TABLE) las que falten.
COLUMNAS_NUEVAS = {
    "importaciones": {
        "producto": "TEXT",
        "fecha_llegada_puerto": "TEXT",
        "fecha_oficializacion": "TEXT",
        "fecha_entrega": "TEXT",
        "costo_financiero_usd": "REAL DEFAULT 0",
        "flete_usd": "REAL DEFAULT 0",
        "tipo_envio": "TEXT",
        "formato_envio": "TEXT",
    },
    "clientes": {
        "rubro": "TEXT",
        "productos_interes": "TEXT",
    },
    "cotizaciones": {
        "tarifa_flete": "REAL DEFAULT 0",
        "pct_certificacion": "REAL DEFAULT 0.5",
        "gastos_origen": "REAL DEFAULT 0",
        "gastos_locales_hdr": "REAL DEFAULT 0",
        "seguro_modo": "TEXT DEFAULT 'auto'",
        "seguro_manual_usd": "REAL DEFAULT 0",
        "tc_venta": "REAL DEFAULT 0",
    },
    "cotizacion_productos": {
        "peso_kg": "REAL DEFAULT 0",
        "volumen_m3": "REAL DEFAULT 0",
    },
    "contacts": {
        "cliente_id": "INTEGER REFERENCES clientes(id) ON DELETE SET NULL",
        "cuit": "TEXT",
    },
    "usuarios": {
        "intentos_fallidos": "INTEGER DEFAULT 0",
        "bloqueado_hasta": "TEXT",
    },
    "importacion_documentos": {
        # El archivo en sí, no solo su nombre — antes vivía en disco local
        # (ruta_archivo apuntaba ahí), pero eso se pierde en cada redeploy
        # del hosting (filesystem efímero). Guardarlo acá adentro hace que
        # viaje junto con el resto de los datos, igual que ya pasa con
        # todo lo demás desde la migración a Turso.
        "contenido": "BLOB",
    },
    "cliente_documentos": {
        "contenido": "BLOB",
    },
}


def _migrar_columnas_faltantes(conn):
    for tabla, columnas in COLUMNAS_NUEVAS.items():
        existentes = {row["name"] for row in conn.execute(f"PRAGMA table_info({tabla})").fetchall()}
        for columna, tipo in columnas.items():
            if columna not in existentes:
                conn.execute(f"ALTER TABLE {tabla} ADD COLUMN {columna} {tipo}")


def init_db():
    conn = get_connection()
    conn.executescript(SCHEMA)
    _migrar_columnas_faltantes(conn)
    for k, v in SETTINGS_DEFAULTS.items():
        conn.execute("INSERT OR IGNORE INTO settings (key, value) VALUES (?, ?)", (k, v))
    conn.commit()
    conn.close()


# ---------- Clientes ----------

def list_clientes(search: str = ""):
    conn = get_connection()
    if search:
        rows = conn.execute(
            "SELECT * FROM clientes WHERE nombre LIKE ? OR cuit LIKE ? ORDER BY nombre",
            (f"%{search}%", f"%{search}%"),
        ).fetchall()
    else:
        rows = conn.execute("SELECT * FROM clientes ORDER BY nombre").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_cliente(cliente_id: int):
    conn = get_connection()
    row = conn.execute("SELECT * FROM clientes WHERE id = ?", (cliente_id,)).fetchone()
    conn.close()
    return dict(row) if row else None


def create_cliente(nombre, cuit="", email="", telefono="", direccion="", rubro="", productos_interes="", notas=""):
    conn = get_connection()
    cur = conn.execute(
        """INSERT INTO clientes (nombre, cuit, email, telefono, direccion, rubro, productos_interes, notas)
           VALUES (?,?,?,?,?,?,?,?)""",
        (nombre, cuit, email, telefono, direccion, rubro, productos_interes, notas),
    )
    conn.commit()
    new_id = cur.lastrowid
    conn.close()
    return new_id


def update_cliente(cliente_id, nombre, cuit="", email="", telefono="", direccion="", rubro="", productos_interes="", notas=""):
    conn = get_connection()
    conn.execute(
        """UPDATE clientes SET nombre=?, cuit=?, email=?, telefono=?, direccion=?,
           rubro=?, productos_interes=?, notas=? WHERE id=?""",
        (nombre, cuit, email, telefono, direccion, rubro, productos_interes, notas, cliente_id),
    )
    conn.commit()
    conn.close()


def delete_cliente(cliente_id):
    conn = get_connection()
    conn.execute("DELETE FROM clientes WHERE id=?", (cliente_id,))
    conn.commit()
    conn.close()


# ---------- Cotizaciones ----------

def next_numero():
    # MAX del sufijo numérico ya usado, no COUNT(*): con COUNT(*) dos
    # cotizaciones borradas (o cualquier gap) hacen que el conteo quede por
    # debajo del máximo real ya asignado, y el siguiente número generado
    # choca con uno existente (ver COT-00003 duplicado en id=52 e id=85).
    # Se incluye también el prefijo viejo 'COT-' en el MAX (aunque ya no se
    # genera) para no repetir un número si quedara alguna cotización sin
    # migrar al nuevo prefijo 'SKY-'.
    conn = get_connection()
    row = conn.execute(
        "SELECT MAX(CAST(SUBSTR(numero, 5) AS INTEGER)) AS maxn FROM cotizaciones "
        "WHERE numero LIKE 'SKY-%' OR numero LIKE 'COT-%'"
    ).fetchone()
    conn.close()
    return f"SKY-{(row['maxn'] or 0) + 1:05d}"


def create_cotizacion_borrador(cliente_id=None):
    """Crea una cotización vacía con gastos operativos por defecto y devuelve su id."""
    conn = get_connection()
    numero = next_numero()
    cur = conn.execute(
        "INSERT INTO cotizaciones (numero, cliente_id) VALUES (?, ?)",
        (numero, cliente_id),
    )
    cot_id = cur.lastrowid
    for i, (concepto, moneda, prorrateo, iva_incl, pct_iva) in enumerate(DEFAULT_GASTOS):
        conn.execute(
            """INSERT INTO cotizacion_gastos
               (cotizacion_id, orden, concepto, moneda, monto, prorrateo, iva_incluido, pct_iva)
               VALUES (?,?,?,?,0,?,?,?)""",
            (cot_id, i, concepto, moneda, prorrateo, iva_incl, pct_iva),
        )
    conn.commit()
    conn.close()
    return cot_id


def list_cotizaciones(search: str = "", cliente_id=None, producto: str = "", fecha_desde: str = None, fecha_hasta: str = None):
    conn = get_connection()
    q = """SELECT c.*, cl.nombre AS cliente_nombre,
           (SELECT GROUP_CONCAT(p.descripcion, ' · ') FROM cotizacion_productos p WHERE p.cotizacion_id = c.id) AS productos_desc
           FROM cotizaciones c LEFT JOIN clientes cl ON cl.id = c.cliente_id"""
    condiciones = []
    params = []
    if search:
        condiciones.append("(c.numero LIKE ? OR cl.nombre LIKE ? OR c.detalle_pedido LIKE ?)")
        params += [f"%{search}%", f"%{search}%", f"%{search}%"]
    if cliente_id is not None:
        condiciones.append("c.cliente_id = ?")
        params.append(cliente_id)
    if producto:
        condiciones.append(
            "EXISTS (SELECT 1 FROM cotizacion_productos p WHERE p.cotizacion_id = c.id AND p.descripcion LIKE ?)"
        )
        params.append(f"%{producto}%")
    if fecha_desde:
        condiciones.append("date(c.fecha) >= date(?)")
        params.append(fecha_desde)
    if fecha_hasta:
        condiciones.append("date(c.fecha) <= date(?)")
        params.append(fecha_hasta)
    if condiciones:
        q += " WHERE " + " AND ".join(condiciones)
    q += " ORDER BY c.fecha DESC, c.id DESC"
    rows = conn.execute(q, params).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def resumen_cotizaciones_por_cliente(dias_sin_respuesta: int):
    """Agregado del historial COMPLETO de cotizaciones de cada cliente (una
    sola consulta SQL agrupada), no solo la más reciente — con cientos de
    cotizaciones acumuladas por cliente, mirar nada más la última da una
    foto parcial (ej. un cliente con 40 rechazadas y 1 recién enviada se
    clasificaría por la enviada, ignorando el patrón). El corte de
    "enviada vigente vs. vencida" usa el umbral dias_sin_respuesta ya
    configurado en Ajustes, así queda consistente con "⏰ Seguimiento".
    Ver crm.estado_comercial(), que consume este resumen."""
    conn = get_connection()
    corte = (datetime.now() - timedelta(days=dias_sin_respuesta)).strftime("%Y-%m-%d %H:%M:%S")
    rows = conn.execute(
        """SELECT cliente_id,
                  COUNT(*) AS total,
                  SUM(CASE WHEN estado='Aprobada' THEN 1 ELSE 0 END) AS aprobadas,
                  SUM(CASE WHEN estado='Rechazada' THEN 1 ELSE 0 END) AS rechazadas,
                  SUM(CASE WHEN estado='Enviada' AND actualizado_en >= ? THEN 1 ELSE 0 END) AS enviadas_vigentes,
                  SUM(CASE WHEN estado='Enviada' AND actualizado_en < ? THEN 1 ELSE 0 END) AS enviadas_vencidas,
                  SUM(CASE WHEN estado IS NULL OR estado='Borrador' THEN 1 ELSE 0 END) AS borradores
           FROM cotizaciones
           WHERE cliente_id IS NOT NULL
           GROUP BY cliente_id""",
        (corte, corte),
    ).fetchall()
    conn.close()
    return {r["cliente_id"]: dict(r) for r in rows}


def contar_importaciones_por_cliente():
    """{cliente_id: cantidad de importaciones}, cualquier estado — una sola
    consulta agrupada, para no recorrer todas las importaciones en Python
    por cada cliente (mismo criterio de escala que resumen_cotizaciones_por_cliente)."""
    conn = get_connection()
    rows = conn.execute(
        "SELECT cliente_id, COUNT(*) AS total FROM importaciones "
        "WHERE cliente_id IS NOT NULL GROUP BY cliente_id"
    ).fetchall()
    conn.close()
    return {r["cliente_id"]: r["total"] for r in rows}


def contar_importaciones_en_curso_por_cliente():
    """{cliente_id: cantidad de importaciones NO entregadas} — para el
    resumen operativo de solo lectura de la columna 'Cliente' del CRM
    (mismo criterio de "en proceso" que ya usa vista_clientes() en app.py:
    todo lo que no sea 'Entregada')."""
    conn = get_connection()
    rows = conn.execute(
        "SELECT cliente_id, COUNT(*) AS total FROM importaciones "
        "WHERE cliente_id IS NOT NULL AND (estado IS NULL OR estado != 'Entregada') "
        "GROUP BY cliente_id"
    ).fetchall()
    conn.close()
    return {r["cliente_id"]: r["total"] for r in rows}


def get_cotizacion(cot_id):
    conn = get_connection()
    row = conn.execute(
        """SELECT c.*, cl.nombre AS cliente_nombre, cl.cuit AS cliente_cuit
           FROM cotizaciones c LEFT JOIN clientes cl ON cl.id = c.cliente_id
           WHERE c.id = ?""",
        (cot_id,),
    ).fetchone()
    conn.close()
    return dict(row) if row else None


def get_productos(cot_id):
    conn = get_connection()
    rows = conn.execute(
        "SELECT * FROM cotizacion_productos WHERE cotizacion_id=? ORDER BY orden, id", (cot_id,)
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_gastos(cot_id):
    conn = get_connection()
    rows = conn.execute(
        "SELECT * FROM cotizacion_gastos WHERE cotizacion_id=? ORDER BY orden, id", (cot_id,)
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def save_cotizacion(cot_id, cabecera: dict, productos: list, gastos: list, totales: dict):
    """Sobrescribe cabecera, líneas de producto y gastos de una cotización, y guarda los totales calculados."""
    conn = get_connection()
    conn.execute(
        """UPDATE cotizaciones SET
             cliente_id=?, detalle_pedido=?, origen_cond_venta=?, contenedor=?, etd_eta=?,
             carrier=?, freetime=?, estado=?, costo_financiero_pct=?, seguro_pct=?,
             tc_tributos=?, tc_operativos=?, arancel_sim=?,
             tarifa_flete=?, pct_certificacion=?, gastos_origen=?, gastos_locales_hdr=?,
             seguro_modo=?, seguro_manual_usd=?, tc_venta=?,
             total_usd_civa=?, total_ars_civa=?, total_usd_sviva=?, total_ars_sviva=?,
             venta_total_usd=?, ganancia_bruta_usd=?, rentabilidad_pct=?,
             actualizado_en=datetime('now','localtime')
           WHERE id=?""",
        (
            cabecera.get("cliente_id"), cabecera.get("detalle_pedido"), cabecera.get("origen_cond_venta"),
            cabecera.get("contenedor"), cabecera.get("etd_eta"), cabecera.get("carrier"),
            cabecera.get("freetime"), cabecera.get("estado", "Borrador"),
            cabecera.get("costo_financiero_pct", 0.025), cabecera.get("seguro_pct", 0.003),
            cabecera.get("tc_tributos", 0), cabecera.get("tc_operativos", 0),
            cabecera.get("arancel_sim", 10),
            cabecera.get("tarifa_flete", 0), cabecera.get("pct_certificacion", 0.5),
            cabecera.get("gastos_origen", 0), cabecera.get("gastos_locales_hdr", 0),
            cabecera.get("seguro_modo", "auto"), cabecera.get("seguro_manual_usd", 0),
            cabecera.get("tc_venta", 0),
            totales.get("total_c_iva_usd"), totales.get("total_c_iva_ars"),
            totales.get("total_s_iva_usd"), totales.get("total_s_iva_ars"),
            totales.get("venta_total_usd"), totales.get("ganancia_bruta_usd"),
            totales.get("rentabilidad_pct"),
            cot_id,
        ),
    )
    conn.execute("DELETE FROM cotizacion_productos WHERE cotizacion_id=?", (cot_id,))
    for i, p in enumerate(productos):
        conn.execute(
            """INSERT INTO cotizacion_productos
               (cotizacion_id, orden, descripcion, ncm, fob_unit, cantidad, fob_decl_unit, peso_kg, volumen_m3,
                pct_derechos, pct_tasa_estadistica, pct_antidumping, pct_iva, pct_iva_adicional,
                pct_ganancias, pct_iibb, margen_pct, pv_final_usd, pv_mercado_usd)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (
                cot_id, i, p.get("descripcion", ""), p.get("ncm", ""),
                p.get("fob_unit", 0) or 0, p.get("cantidad", 0) or 0,
                p.get("fob_decl_unit", 0) or 0, p.get("peso_kg", 0) or 0, p.get("volumen_m3", 0) or 0,
                p.get("pct_derechos", 0) or 0, p.get("pct_tasa_estadistica", 0.03) or 0,
                p.get("pct_antidumping", 0) or 0, p.get("pct_iva", 0.21) or 0,
                p.get("pct_iva_adicional", 0) or 0, p.get("pct_ganancias", 0) or 0,
                p.get("pct_iibb", 0.05) or 0, p.get("margen_pct", 0.3) or 0,
                p.get("pv_final_usd", 0) or 0, p.get("pv_mercado_usd", 0) or 0,
            ),
        )
    conn.execute("DELETE FROM cotizacion_gastos WHERE cotizacion_id=?", (cot_id,))
    for i, g in enumerate(gastos):
        conn.execute(
            """INSERT INTO cotizacion_gastos
               (cotizacion_id, orden, concepto, moneda, monto, prorrateo, iva_incluido, pct_iva)
               VALUES (?,?,?,?,?,?,?,?)""",
            (
                cot_id, i, g.get("concepto", ""), g.get("moneda", "USD"), g.get("monto", 0) or 0,
                g.get("prorrateo", "FOB"), g.get("iva_incluido", "NO"), g.get("pct_iva", 0.21) or 0,
            ),
        )
    conn.commit()
    conn.close()


def set_estado_cotizacion(cot_id, estado):
    """Cambia únicamente el estado de una cotización, sin tocar productos,
    gastos ni totales — para el control independiente en la card del listado."""
    conn = get_connection()
    conn.execute(
        "UPDATE cotizaciones SET estado=?, actualizado_en=datetime('now','localtime') WHERE id=?",
        (estado, cot_id),
    )
    conn.commit()
    conn.close()


def delete_cotizacion(cot_id):
    conn = get_connection()
    conn.execute("DELETE FROM cotizaciones WHERE id=?", (cot_id,))
    conn.commit()
    conn.close()


def duplicate_cotizacion(cot_id):
    cab = get_cotizacion(cot_id)
    productos = get_productos(cot_id)
    gastos = get_gastos(cot_id)
    new_id = create_cotizacion_borrador(cliente_id=cab.get("cliente_id"))
    save_cotizacion(
        new_id,
        {
            "cliente_id": cab.get("cliente_id"),
            "detalle_pedido": cab.get("detalle_pedido"),
            "origen_cond_venta": cab.get("origen_cond_venta"),
            "contenedor": cab.get("contenedor"),
            "etd_eta": cab.get("etd_eta"),
            "carrier": cab.get("carrier"),
            "freetime": cab.get("freetime"),
            "estado": "Borrador",
            "costo_financiero_pct": cab.get("costo_financiero_pct"),
            "seguro_pct": cab.get("seguro_pct"),
            "tc_tributos": cab.get("tc_tributos"),
            "tc_operativos": cab.get("tc_operativos"),
            "arancel_sim": cab.get("arancel_sim"),
            "tarifa_flete": cab.get("tarifa_flete"),
            "pct_certificacion": cab.get("pct_certificacion"),
            "gastos_origen": cab.get("gastos_origen"),
            "gastos_locales_hdr": cab.get("gastos_locales_hdr"),
        },
        productos,
        gastos,
        {},
    )
    return new_id


# ---------- Importaciones ----------

def next_numero_importacion():
    # Mismo fix que next_numero(): MAX del sufijo ya usado, no COUNT(*).
    conn = get_connection()
    row = conn.execute(
        "SELECT MAX(CAST(SUBSTR(numero, 5) AS INTEGER)) AS maxn FROM importaciones WHERE numero LIKE 'IMP-%'"
    ).fetchone()
    conn.close()
    return f"IMP-{(row['maxn'] or 0) + 1:05d}"


def create_importacion(cliente_id, cotizacion_id=None):
    conn = get_connection()
    numero = next_numero_importacion()
    cur = conn.execute(
        # El estado se fija explícito (no se depende del DEFAULT de la
        # columna): un DEFAULT en SQLite queda fijo en el archivo .db ya
        # creado y no se actualiza solo si luego se edita el schema acá.
        # Debe coincidir con ESTADOS_IMPORTACION[0] en app.py.
        "INSERT INTO importaciones (numero, cliente_id, cotizacion_id, estado) VALUES (?,?,?,?)",
        (numero, cliente_id, cotizacion_id, "En producción"),
    )
    conn.commit()
    new_id = cur.lastrowid
    conn.close()
    return new_id


def list_importaciones(cliente_id=None):
    conn = get_connection()
    if cliente_id is not None:
        rows = conn.execute(
            "SELECT * FROM importaciones WHERE cliente_id=? ORDER BY id DESC", (cliente_id,)
        ).fetchall()
    else:
        rows = conn.execute(
            """SELECT i.*, cl.nombre AS cliente_nombre
               FROM importaciones i LEFT JOIN clientes cl ON cl.id = i.cliente_id
               ORDER BY i.id DESC"""
        ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_importacion(imp_id):
    conn = get_connection()
    row = conn.execute("SELECT * FROM importaciones WHERE id=?", (imp_id,)).fetchone()
    conn.close()
    return dict(row) if row else None


def update_importacion(imp_id, **campos):
    if not campos:
        return
    sets = ", ".join(f"{k}=?" for k in campos)
    valores = list(campos.values())
    conn = get_connection()
    conn.execute(
        f"UPDATE importaciones SET {sets}, actualizado_en=datetime('now','localtime') WHERE id=?",
        (*valores, imp_id),
    )
    conn.commit()
    conn.close()


def delete_importacion(imp_id):
    conn = get_connection()
    conn.execute("DELETE FROM importaciones WHERE id=?", (imp_id,))
    conn.commit()
    conn.close()


def add_documento(importacion_id, categoria, nombre_archivo, contenido: bytes):
    # ruta_archivo ya no se usa para ubicar el archivo (eso ahora es
    # 'contenido', ver COLUMNAS_NUEVAS) — se sigue completando solo porque
    # la columna es NOT NULL desde la versión vieja del schema.
    conn = get_connection()
    conn.execute(
        """INSERT INTO importacion_documentos (importacion_id, categoria, nombre_archivo, ruta_archivo, contenido)
           VALUES (?,?,?,?,?)""",
        (importacion_id, categoria, nombre_archivo, nombre_archivo, contenido),
    )
    conn.commit()
    conn.close()


def contar_documentos(importacion_id):
    """COUNT liviano (sin traer 'contenido') — para mostrar cuántos
    documentos tiene una importación sin pagar el costo de bajar todos los
    archivos, como pasaría con list_documentos()."""
    conn = get_connection()
    row = conn.execute(
        "SELECT COUNT(*) AS n FROM importacion_documentos WHERE importacion_id=?", (importacion_id,)
    ).fetchone()
    conn.close()
    return row["n"] if row else 0


def list_documentos(importacion_id, categoria=None):
    conn = get_connection()
    if categoria:
        rows = conn.execute(
            "SELECT * FROM importacion_documentos WHERE importacion_id=? AND categoria=? ORDER BY id",
            (importacion_id, categoria),
        ).fetchall()
    else:
        rows = conn.execute(
            "SELECT * FROM importacion_documentos WHERE importacion_id=? ORDER BY id", (importacion_id,)
        ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_documento(doc_id):
    conn = get_connection()
    row = conn.execute("SELECT * FROM importacion_documentos WHERE id=?", (doc_id,)).fetchone()
    conn.close()
    return dict(row) if row else None


def delete_documento(doc_id):
    conn = get_connection()
    conn.execute("DELETE FROM importacion_documentos WHERE id=?", (doc_id,))
    conn.commit()
    conn.close()


def rename_documento(doc_id, nuevo_nombre):
    # Solo cambia el nombre "de vidriera" (nombre_archivo) — el contenido
    # (columna 'contenido') no se toca, así que no hay ningún I/O aparte
    # del UPDATE.
    conn = get_connection()
    conn.execute("UPDATE importacion_documentos SET nombre_archivo=? WHERE id=?", (nuevo_nombre, doc_id))
    conn.commit()
    conn.close()


# ---------- Documentos de cliente ----------

def contar_documentos_cliente(cliente_id):
    conn = get_connection()
    row = conn.execute(
        "SELECT COUNT(*) AS n FROM cliente_documentos WHERE cliente_id=?", (cliente_id,)
    ).fetchone()
    conn.close()
    return row["n"] if row else 0


def add_documento_cliente(cliente_id, nombre_archivo, contenido: bytes, tamano_bytes=0):
    conn = get_connection()
    conn.execute(
        """INSERT INTO cliente_documentos (cliente_id, nombre_archivo, ruta_archivo, tamano_bytes, contenido)
           VALUES (?,?,?,?,?)""",
        (cliente_id, nombre_archivo, nombre_archivo, tamano_bytes, contenido),
    )
    conn.commit()
    conn.close()


def list_documentos_cliente(cliente_id):
    conn = get_connection()
    rows = conn.execute(
        "SELECT * FROM cliente_documentos WHERE cliente_id=? ORDER BY id DESC", (cliente_id,)
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_documento_cliente(doc_id):
    conn = get_connection()
    row = conn.execute("SELECT * FROM cliente_documentos WHERE id=?", (doc_id,)).fetchone()
    conn.close()
    return dict(row) if row else None


def delete_documento_cliente(doc_id):
    conn = get_connection()
    conn.execute("DELETE FROM cliente_documentos WHERE id=?", (doc_id,))
    conn.commit()
    conn.close()


def rename_documento_cliente(doc_id, nuevo_nombre):
    conn = get_connection()
    conn.execute("UPDATE cliente_documentos SET nombre_archivo=? WHERE id=?", (nuevo_nombre, doc_id))
    conn.commit()
    conn.close()


# ---------------------------------------------------------------- CRM: CONTACTOS

def list_contacts(search: str = ""):
    # ultima_actividad sale de activity_log (fuente única de verdad del
    # timeline) via subquery, para poder filtrar "sin actividad hace más de
    # X días" en la capa de arriba sin otra consulta por contacto.
    q = """SELECT c.*,
                  (SELECT MAX(a.fecha) FROM activity_log a WHERE a.contact_id = c.id) AS ultima_actividad
           FROM contacts c"""
    conn = get_connection()
    if search:
        b = f"%{search}%"
        rows = conn.execute(
            q + " WHERE c.nombre LIKE ? OR c.empresa LIKE ? OR c.email LIKE ? OR c.cuit LIKE ? ORDER BY c.fecha_alta DESC",
            (b, b, b, b),
        ).fetchall()
    else:
        rows = conn.execute(q + " ORDER BY c.fecha_alta DESC").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_contact(contact_id):
    conn = get_connection()
    row = conn.execute("SELECT * FROM contacts WHERE id=?", (contact_id,)).fetchone()
    conn.close()
    return dict(row) if row else None


def find_contacto_duplicado(email: str = "", whatsapp: str = ""):
    """Busca un contacto ya existente por email o whatsapp normalizados,
    para poder reimportar un archivo sin generar duplicados."""
    conn = get_connection()
    row = None
    if email:
        row = conn.execute(
            "SELECT * FROM contacts WHERE email <> '' AND lower(email) = lower(?)", (email,)
        ).fetchone()
    if row is None and whatsapp:
        row = conn.execute(
            "SELECT * FROM contacts WHERE whatsapp <> '' AND whatsapp = ?", (whatsapp,)
        ).fetchone()
    conn.close()
    return dict(row) if row else None


def create_contact(nombre, empresa="", email="", whatsapp="", origen="", etapa="Nuevo",
                    asignado_a="", proximo_seguimiento=None, cuit=""):
    conn = get_connection()
    cur = conn.execute(
        """INSERT INTO contacts (nombre, empresa, cuit, email, whatsapp, origen, etapa, asignado_a, proximo_seguimiento)
           VALUES (?,?,?,?,?,?,?,?,?)""",
        (nombre, empresa, cuit, email, whatsapp, origen, etapa, asignado_a, proximo_seguimiento),
    )
    conn.commit()
    new_id = cur.lastrowid
    conn.close()
    return new_id


def update_contact(contact_id, nombre, empresa="", email="", whatsapp="", origen="",
                    asignado_a="", proximo_seguimiento=None, cuit=""):
    # La etapa NO se actualiza acá: cambia únicamente por change_etapa_contacto,
    # que además deja registro en activity_log — así el timeline nunca queda
    # desincronizado de la etapa actual del contacto.
    conn = get_connection()
    conn.execute(
        """UPDATE contacts SET nombre=?, empresa=?, cuit=?, email=?, whatsapp=?, origen=?,
           asignado_a=?, proximo_seguimiento=? WHERE id=?""",
        (nombre, empresa, cuit, email, whatsapp, origen, asignado_a, proximo_seguimiento, contact_id),
    )
    conn.commit()
    conn.close()


def delete_contact(contact_id):
    conn = get_connection()
    conn.execute("DELETE FROM contacts WHERE id=?", (contact_id,))
    conn.commit()
    conn.close()


def change_etapa_contacto(contact_id, nueva_etapa, autor=""):
    conn = get_connection()
    row = conn.execute("SELECT etapa FROM contacts WHERE id=?", (contact_id,)).fetchone()
    etapa_anterior = row["etapa"] if row else None
    conn.execute("UPDATE contacts SET etapa=? WHERE id=?", (nueva_etapa, contact_id))
    if etapa_anterior != nueva_etapa:
        conn.execute(
            "INSERT INTO activity_log (contact_id, tipo, autor, texto) VALUES (?,?,?,?)",
            (contact_id, "cambio_etapa", autor, f"{etapa_anterior or '—'} → {nueva_etapa}"),
        )
    conn.commit()
    conn.close()


def add_activity(contact_id, tipo, texto="", autor=""):
    conn = get_connection()
    conn.execute(
        "INSERT INTO activity_log (contact_id, tipo, autor, texto) VALUES (?,?,?,?)",
        (contact_id, tipo, autor, texto),
    )
    conn.commit()
    conn.close()


def list_activity(contact_id):
    conn = get_connection()
    rows = conn.execute(
        "SELECT * FROM activity_log WHERE contact_id=? ORDER BY fecha DESC, id DESC", (contact_id,)
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def list_all_activity():
    """Todo activity_log de todos los contactos, para las analíticas del
    embudo (funnel, conversión, tiempo por etapa, tasa de respuesta)."""
    conn = get_connection()
    rows = conn.execute("SELECT * FROM activity_log ORDER BY fecha").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def link_contact_cliente(contact_id, cliente_id):
    """Vincula un contacto del CRM a un registro de `clientes` — a partir de
    acá el contacto puede cotizar (reutiliza el motor de cotizador ya
    existente para clientes) y acceder a su ficha completa."""
    conn = get_connection()
    conn.execute("UPDATE contacts SET cliente_id=? WHERE id=?", (cliente_id, contact_id))
    conn.commit()
    conn.close()


def get_contact_by_cliente_id(cliente_id):
    """Contacto del CRM vinculado a un cliente (si lo hay) — se usa para
    propagar 'cotización aprobada' hacia 'contacto Ganado'."""
    conn = get_connection()
    row = conn.execute("SELECT * FROM contacts WHERE cliente_id=?", (cliente_id,)).fetchone()
    conn.close()
    return dict(row) if row else None


# ---------------------------------------------------------------- CONFIGURACIÓN

def get_setting(key, default=None):
    conn = get_connection()
    row = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    conn.close()
    return row["value"] if row else default


def set_setting(key, value):
    conn = get_connection()
    conn.execute(
        "INSERT INTO settings (key, value) VALUES (?, ?) "
        "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
        (key, str(value)),
    )
    conn.commit()
    conn.close()
