"""Crea un usuario directo en la base, por consola. Uso: python seed_usuario.py
La pantalla de login NUNCA permite crear cuentas por sí sola — esta es la
forma de dar de alta un usuario sin pasar por el navegador."""
import getpass
import db

db.init_db()
username = input("Usuario: ").strip()
nombre = input("Nombre: ").strip()
password = getpass.getpass("Contraseña: ")
password2 = getpass.getpass("Repetir contraseña: ")
if not username or not nombre or not password:
    print("Completá todos los campos.")
elif password != password2:
    print("Las contraseñas no coinciden.")
else:
    db.create_usuario(username, password, nombre)
    print(f"Usuario '{username}' creado.")
