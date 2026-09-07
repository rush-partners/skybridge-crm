"""Migración única de documentos existentes: copia el CONTENIDO de cada
archivo (que hasta ahora vivía solo en disco local, en la ruta guardada en
'ruta_archivo') hacia la columna 'contenido' de Turso, para que a partir de
ahora viajen junto con el resto de los datos y no se pierdan en cada
redeploy del hosting (que no tiene disco persistente).

Correr UNA sola vez, DESDE ESTA MÁQUINA (donde están los archivos originales
en disco), DESPUÉS de actualizar db.py/app.py/storage.py pero ANTES (o
después, da igual, es idempotente) de subir los cambios a GitHub:

    python migrar_documentos_a_turso.py

Es seguro correrlo más de una vez: solo toca filas con contenido todavía
vacío (NULL), así que un documento ya migrado no se vuelve a tocar.

No hace falta Streamlit para correrlo (es un script de línea de comandos
suelto), así que lee TURSO_DATABASE_URL / TURSO_AUTH_TOKEN directamente de
.streamlit/secrets.toml en vez de pasar por st.secrets."""
import tomllib
from pathlib import Path

import turso_serverless

RAIZ = Path(__file__).parent
SECRETS_PATH = RAIZ / ".streamlit" / "secrets.toml"


def _cargar_credenciales():
    if not SECRETS_PATH.exists():
        raise SystemExit(
            f"No encontré {SECRETS_PATH}. Corré este script desde la carpeta del "
            "proyecto (donde está .streamlit/secrets.toml)."
        )
    secretos = tomllib.loads(SECRETS_PATH.read_text(encoding="utf-8"))
    url = secretos.get("TURSO_DATABASE_URL")
    token = secretos.get("TURSO_AUTH_TOKEN")
    if not url or not token:
        raise SystemExit(f"{SECRETS_PATH} no tiene TURSO_DATABASE_URL / TURSO_AUTH_TOKEN.")
    return url, token


def _migrar_tabla(conn, tabla: str, columna_id: str = "id"):
    filas = conn.execute(
        f"SELECT {columna_id} AS id, ruta_archivo, nombre_archivo FROM {tabla} WHERE contenido IS NULL"
    ).fetchall()
    migrados, faltantes = 0, []
    for fila in filas:
        ruta = Path(fila["ruta_archivo"])
        if not ruta.exists():
            faltantes.append((fila["id"], fila["nombre_archivo"], str(ruta)))
            continue
        contenido = ruta.read_bytes()
        conn.execute(
            f"UPDATE {tabla} SET contenido = ? WHERE {columna_id} = ?",
            (contenido, fila["id"]),
        )
        migrados += 1
        print(f"  ✓ [{tabla}] id={fila['id']} '{fila['nombre_archivo']}' ({len(contenido)} bytes)")
    return migrados, faltantes


def main():
    url, token = _cargar_credenciales()
    print(f"Conectando a Turso ({url})…")
    conn = turso_serverless.connect(url, auth_token=token)
    conn.row_factory = turso_serverless.Row

    print("\nDocumentos de importación (importacion_documentos):")
    m1, f1 = _migrar_tabla(conn, "importacion_documentos")

    print("\nDocumentos de cliente (cliente_documentos):")
    m2, f2 = _migrar_tabla(conn, "cliente_documentos")

    conn.commit()
    conn.close()

    print(f"\nTotal migrado: {m1 + m2} documento(s).")
    faltantes = f1 + f2
    if faltantes:
        print(
            f"\n⚠ {len(faltantes)} documento(s) NO se pudieron migrar porque ya no "
            "encontré el archivo en disco (se perdió antes de esta migración, o la "
            "ruta guardada corresponde a otra máquina). Quedan con vista previa vacía:"
        )
        for doc_id, nombre, ruta in faltantes:
            print(f"  ✗ id={doc_id} '{nombre}' — buscado en: {ruta}")
    else:
        print("Todos los documentos existentes tenían su archivo en disco y se migraron sin problemas.")


if __name__ == "__main__":
    main()
