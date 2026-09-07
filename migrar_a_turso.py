"""Migra los datos reales de skybridge.db (local) a Turso. Se corre UNA
SOLA VEZ, después de que db.py ya apunta a Turso — así la primera vez que
la app se conecta, ya encuentra los datos reales adentro en vez de una
base vacía.
Uso: python migrar_a_turso.py"""
from pathlib import Path
import db

ORIGEN = Path(__file__).parent / "skybridge.db"

if not ORIGEN.exists():
    print("No se encontró skybridge.db en esta carpeta — no hay nada para migrar.")
else:
    print("Creando el esquema en Turso...")
    db.init_db()
    print("Copiando los datos reales a Turso (puede tardar unos segundos)...")
    db.restaurar_desde_sqlite_bytes(ORIGEN.read_bytes())
    print("Listo. Verificando...")
    print(f"  Clientes: {len(db.list_clientes())}")
    print(f"  Cotizaciones: {len(db.list_cotizaciones())}")
    print(f"  Importaciones: {len(db.list_importaciones())}")
    print(f"  Contactos: {len(db.list_contacts())}")
    print(f"  Usuarios: {len(db.list_usuarios())}")
    print("\nComparalo con lo que tenías antes de migrar. Si los números coinciden, salió bien.")
