"""Migración única: convierte el texto libre que había en
importaciones.notas_transito ("Detalles a considerar (notas de tránsito)")
en notas individuales de la nueva bitácora importacion_log — una fila por
línea, cada una con su propia fecha, en vez de un solo bloque de texto que
Tom mantenía a mano escribiendo él mismo la fecha de cada línea (ej.
"19/08 confirmación de la carga").

Cada línea de notas_transito debe empezar con una fecha DD/MM (el formato
que Tom viene usando). Como esas fechas no llevan año, se infiere uno:
arranca en el año de creación de la importación (creado_en) y avanza un
año cada vez que la fecha de una línea "retrocede" respecto de la
anterior (ej. de diciembre a enero) — así una racha de fechas ascendentes
dentro de la misma importación no cruza de año de pedo, pero si la racha
da la vuelta del calendario, el año sube solo. Una línea que NO arranca
con fecha se pega como continuación de la nota anterior (texto multilínea);
si es la primera línea de todas y no tiene fecha, se migra igual con la
fecha de creación de la importación, marcada como "SIN FECHA" en el
reporte para que Tom la revise a mano.

notas_transito NO se borra ni se toca — queda en la tabla importaciones
como archivo de lo ya cargado, pero la ficha ya no la lee ni la edita
(ver _render_timeline_importacion en app.py). Es idempotente: una
importación que ya tiene filas en importacion_log se saltea (evita migrar
dos veces si el script se corre de nuevo).

Correr DESPUÉS de haber actualizado db.py/app.py (necesita la tabla
importacion_log ya creada — alcanza con abrir la app una vez, o correr
init_db(), para que se cree sola) y DESPUÉS de subir los cambios:

    python migrar_notas_transito.py             # dry-run: solo muestra qué haría
    python migrar_notas_transito.py --commit     # migra de verdad

No hace falta Streamlit para correrlo — lee TURSO_DATABASE_URL /
TURSO_AUTH_TOKEN directo de .streamlit/secrets.toml, igual que
migrar_documentos_a_turso.py."""
import re
import sys
import tomllib
from datetime import date
from pathlib import Path

import turso_serverless

RAIZ = Path(__file__).parent
SECRETS_PATH = RAIZ / ".streamlit" / "secrets.toml"

LINEA_FECHA_RE = re.compile(r"^(\d{1,2})/(\d{1,2})\.?\s*(.*)$")


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


def _parsear_lineas(texto):
    """Devuelve una lista de {dia, mes, texto, sin_fecha} — una por nota,
    juntando líneas de continuación (sin fecha propia) a la nota anterior."""
    entradas = []
    for linea_cruda in texto.splitlines():
        linea = linea_cruda.strip()
        if not linea:
            continue
        m = LINEA_FECHA_RE.match(linea)
        dia = mes = None
        if m:
            posible_dia, posible_mes = int(m.group(1)), int(m.group(2))
            if 1 <= posible_dia <= 31 and 1 <= posible_mes <= 12:
                dia, mes = posible_dia, posible_mes
                resto = m.group(3).strip()
        if dia is not None:
            entradas.append({"dia": dia, "mes": mes, "texto": resto, "sin_fecha": False})
        elif entradas:
            entradas[-1]["texto"] = (entradas[-1]["texto"] + "\n" + linea).strip()
        else:
            entradas.append({"dia": None, "mes": None, "texto": linea, "sin_fecha": True})
    return entradas


def _resolver_fechas(entradas, creado_en_str):
    """Completa el año de cada entrada con fecha (día/mes) usando la regla
    descripta arriba, y arma el string 'YYYY-MM-DD HH:MM:SS' final (mismo
    formato que usa el resto de la app) para cada una."""
    try:
        creado_en = date.fromisoformat((creado_en_str or "")[:10])
    except ValueError:
        creado_en = date.today()

    year_actual = creado_en.year
    prev_md = (creado_en.day, creado_en.month)
    resultado = []
    for e in entradas:
        if e["sin_fecha"]:
            fecha_iso = f"{creado_en.isoformat()} 12:00:00"
        else:
            if (e["mes"], e["dia"]) < (prev_md[1], prev_md[0]):
                year_actual += 1
            try:
                fecha = date(year_actual, e["mes"], e["dia"])
            except ValueError:
                # 31/04, 30/02, etc. — fecha inválida en el calendario real;
                # se migra igual con la fecha de creación, marcada para revisar.
                resultado.append({**e, "fecha_iso": f"{creado_en.isoformat()} 12:00:00", "sin_fecha": True})
                continue
            fecha_iso = f"{fecha.isoformat()} 12:00:00"
            prev_md = (e["dia"], e["mes"])
        resultado.append({**e, "fecha_iso": fecha_iso})
    return resultado


def main():
    commit = "--commit" in sys.argv
    url, token = _cargar_credenciales()
    print(f"Conectando a Turso ({url})…")
    conn = turso_serverless.connect(url, auth_token=token)
    conn.row_factory = turso_serverless.Row

    filas = conn.execute(
        "SELECT id, numero, creado_en, notas_transito FROM importaciones "
        "WHERE notas_transito IS NOT NULL AND TRIM(notas_transito) != ''"
    ).fetchall()

    total_notas = 0
    total_sin_fecha = 0
    saltadas = 0

    for fila in filas:
        imp_id, numero, creado_en, texto = fila["id"], fila["numero"], fila["creado_en"], fila["notas_transito"]

        ya_migrada = conn.execute(
            "SELECT COUNT(*) AS n FROM importacion_log WHERE importacion_id=?", (imp_id,)
        ).fetchone()["n"]
        if ya_migrada:
            saltadas += 1
            continue

        entradas = _resolver_fechas(_parsear_lineas(texto), creado_en)
        if not entradas:
            continue

        print(f"\n{numero or f'#{imp_id}'} (creada {creado_en}):")
        for e in entradas:
            marca = "  ⚠ SIN FECHA (revisar)" if e["sin_fecha"] else ""
            primera_linea = e["texto"].splitlines()[0] if e["texto"] else ""
            print(f"  {e['fecha_iso'][:10]} — {primera_linea}{marca}")
            if "\n" in e["texto"]:
                for cont in e["texto"].splitlines()[1:]:
                    print(f"             {cont}")
            total_notas += 1
            if e["sin_fecha"]:
                total_sin_fecha += 1

        if commit:
            for e in entradas:
                conn.execute(
                    "INSERT INTO importacion_log (importacion_id, autor, texto, fecha) VALUES (?,?,?,?)",
                    (imp_id, "", e["texto"], e["fecha_iso"]),
                )

    if commit:
        conn.commit()
    conn.close()

    print(f"\n{'Migradas' if commit else '[DRY RUN] Se migrarían'}: {total_notas} nota(s) "
          f"en {len(filas) - saltadas} importación(es).")
    if saltadas:
        print(f"Salteadas (ya tenían notas en la bitácora nueva): {saltadas} importación(es).")
    if total_sin_fecha:
        print(f"\n⚠ {total_sin_fecha} nota(s) sin fecha reconocible — revisalas en la ficha después de migrar.")
    if not commit:
        print("\nEsto fue un dry-run, no se escribió nada. Si el reporte de arriba está bien, "
              "corré:\n    python migrar_notas_transito.py --commit")


if __name__ == "__main__":
    main()
