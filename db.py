import sqlite3
import hashlib


conecion= sqlite3.connect("autocloudparking.db")
conecion.execute("CREATE TABLE IF NOT EXISTS vehiculo (placa TEXT PRIMARY KEY, tipo TEXT)")
conecion.execute("CREATE TABLE IF NOT EXISTS tarifa (tipo TEXT PRIMARY KEY, valor_hora INTEGER)")
conecion.execute("CREATE TABLE IF NOT EXISTS registro(id INTEGER PRIMARY KEY AUTOINCREMENT, placa TEXT, entrada TEXT, salida TEXT, valor INTEGER)")
conecion.execute("CREATE TABLE IF NOT EXISTS usuario (usuario TEXT PRIMARY KEY, clave TEXT)")
conecion.execute("INSERT OR IGNORE INTO usuario values('admin' ,?)",(hashlib.sha256(b"123").hexdigest(),))
conecion.execute("INSERT OR IGNORE INTO tarifa values('Carro', 2500)")
conecion.execute("INSERT OR IGNORE INTO tarifa values('Moto', 2000)")
conecion.execute("INSERT OR IGNORE INTO tarifa values('Bicicleta', 1000)")
conecion.commit()
conecion.close()

print("Conectado")
