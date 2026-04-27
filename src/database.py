import sqlite3
import os
from datetime import datetime


# =============================================================================
#  SECCIÓN 1 — INICIALIZACIÓN
# =============================================================================

def ruta_bd():
    """Genera una ruta persistente y semi-oculta en AppData para la base de datos."""
    # Esto apunta a C:\Users\[USER]\AppData\Local\Abacus_Data
    carpeta_appdata = os.path.join(os.getenv('LOCALAPPDATA'), 'Abacus_Data')
    os.makedirs(carpeta_appdata, exist_ok=True)
    return os.path.join(carpeta_appdata, 'arqueo_local.db')

def inicializar_db():
    """
    Crea las tablas principales si no existen:
      · bolsas     — registros de efectivo en sobre/bolsa individual
      · monedero   — saldo acumulado por denominación de moneda
    """
    conexion = sqlite3.connect(ruta_bd())
    cursor   = conexion.cursor()

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS bolsas (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha_ingreso TEXT    NOT NULL,
            nif           TEXT,
            contrato      TEXT,
            importe       REAL    NOT NULL,
            estado        TEXT    DEFAULT 'IN-CAJA FUERTE',
            lote_envio    TEXT
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS monedero (
            denominacion   TEXT PRIMARY KEY,
            cantidad_total REAL DEFAULT 0.0
        )
    ''')

    conexion.commit()
    conexion.close()


def inicializar_snapshot_table():
    """
    Crea la tabla de snapshots diarios si no existe.
    Se llama por separado de inicializar_db() para permitir
    añadir esta funcionalidad sin alterar el esquema principal.
    """
    conexion = sqlite3.connect(ruta_bd())
    cursor   = conexion.cursor()

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS snapshot_diario (
            fecha             TEXT,
            denominacion      TEXT,
            cantidad_inicial  REAL,
            PRIMARY KEY (fecha, denominacion)
        )
    ''')

    conexion.commit()
    conexion.close()


# =============================================================================
#  SECCIÓN 2 — MONEDERO
# =============================================================================

def modificar_monedero(denominacion, variacion):
    """
    Aplica un delta (+/-) al saldo de una denominación (UPSERT).
    Si la denominación no existe aún en la tabla, la crea con ese valor inicial.
    Usado por las transacciones en tiempo real desde los campos +/- de la UI.
    """
    conexion = sqlite3.connect(ruta_bd())
    cursor   = conexion.cursor()

    cursor.execute('''
        INSERT INTO monedero (denominacion, cantidad_total)
        VALUES (?, ?)
        ON CONFLICT(denominacion)
        DO UPDATE SET cantidad_total = cantidad_total + excluded.cantidad_total
    ''', (denominacion, variacion))

    conexion.commit()
    conexion.close()


def guardar_saldo_total(denominacion, total):
    """
    Sobrescribe el saldo de una denominación con un valor absoluto (UPSERT).
    A diferencia de modificar_monedero(), reemplaza el total en lugar de sumarlo.
    Usado exclusivamente por el modo de edición directa de totales.
    """
    conexion = sqlite3.connect(ruta_bd())
    cursor   = conexion.cursor()

    cursor.execute('''
        INSERT INTO monedero (denominacion, cantidad_total)
        VALUES (?, ?)
        ON CONFLICT(denominacion)
        DO UPDATE SET cantidad_total = excluded.cantidad_total
    ''', (denominacion, total))

    conexion.commit()
    conexion.close()


def obtener_todos_los_saldos():
    """
    Devuelve todos los registros de la tabla monedero.
    Usado al arrancar la aplicación para poblar los labels de la UI.
    Retorna: lista de tuplas (denominacion, cantidad_total)
    """
    conexion = sqlite3.connect(ruta_bd())
    cursor   = conexion.cursor()

    cursor.execute("SELECT denominacion, cantidad_total FROM monedero")
    datos = cursor.fetchall()

    conexion.close()
    return datos


# =============================================================================
#  SECCIÓN 3 — BOLSAS
#  Una bolsa puede estar en uno de dos estados: 'IN-CAJA FUERTE' o 'ENVIADA'.
#  El flujo normal es: registro → envío (agrupado en lote) → [reversión opcional]
# =============================================================================

def registrar_bolsa(nif, contrato, importe):
    """
    Inserta una nueva bolsa con estado 'IN-CAJA FUERTE' (valor por defecto).
    La marca de tiempo se genera automáticamente en el momento del registro.
    """
    fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    conexion = sqlite3.connect(ruta_bd())
    cursor   = conexion.cursor()

    cursor.execute('''
        INSERT INTO bolsas (fecha_ingreso, nif, contrato, importe)
        VALUES (?, ?, ?, ?)
    ''', (fecha_actual, nif, contrato, importe))

    conexion.commit()
    conexion.close()


def obtener_bolsas_in_caja():
    """
    Devuelve todas las bolsas con estado 'IN-CAJA FUERTE' para poblar la tabla
    principal de la UI. El campo 'id' se incluye aunque está oculto en la vista;
    es necesario para identificar la fila en operaciones posteriores (envío, borrado).
    Retorna: lista de tuplas (id, importe, contrato, nif)
    """
    conexion = sqlite3.connect(ruta_bd())
    cursor   = conexion.cursor()

    cursor.execute('''
        SELECT id, importe, contrato, nif
        FROM bolsas
        WHERE estado = 'IN-CAJA FUERTE'
    ''')
    bolsas = cursor.fetchall()

    conexion.close()
    return bolsas


def borrar_bolsa(id_bolsa):
    """
    Elimina permanentemente una bolsa por su ID.
    Esta operación no tiene vuelta atrás; si se necesita deshacer un envío,
    usar revertir_envio_bolsa() en su lugar.
    """
    conexion = sqlite3.connect(ruta_bd())
    cursor   = conexion.cursor()

    cursor.execute("DELETE FROM bolsas WHERE id = ?", (id_bolsa,))

    conexion.commit()
    conexion.close()


def registrar_envio_bolsas(ids_bolsas, fecha_lote):
    """
    Marca un conjunto de bolsas como 'ENVIADA' y las agrupa bajo un mismo lote.
    El nombre del lote sigue el formato "Bolsas - DD/MM/YYYY", que se usa
    como identificador en el historial maestro-detalle de la UI.
    """
    nombre_lote = f"Bolsas - {fecha_lote}"

    conexion = sqlite3.connect(ruta_bd())
    cursor   = conexion.cursor()

    for id_b in ids_bolsas:
        cursor.execute('''
            UPDATE bolsas
            SET estado = 'ENVIADA', lote_envio = ?
            WHERE id = ?
        ''', (nombre_lote, id_b))

    conexion.commit()
    conexion.close()


def revertir_envio_bolsa(id_bolsa):
    """
    Devuelve una bolsa enviada al estado 'IN-CAJA FUERTE' y borra su lote.
    Operación de seguridad para corregir envíos incorrectos desde el historial.
    """
    conexion = sqlite3.connect(ruta_bd())
    cursor   = conexion.cursor()

    cursor.execute('''
        UPDATE bolsas
        SET estado = 'IN-CAJA FUERTE', lote_envio = NULL
        WHERE id = ?
    ''', (id_bolsa,))

    conexion.commit()
    conexion.close()


# =============================================================================
#  SECCIÓN 4 — HISTORIAL DE LOTES
# =============================================================================

def obtener_lotes_enviados():
    """
    Devuelve los últimos 10 lotes de envío distintos, ordenados cronológicamente
    (el más reciente primero). Cada lote es una agrupación de bolsas enviadas
    el mismo día con el mismo identificador de lote.
    Retorna: lista de tuplas (nombre_lote,)
    """
    conexion = sqlite3.connect(ruta_bd())
    cursor   = conexion.cursor()

    cursor.execute('''
        SELECT lote_envio
        FROM bolsas
        WHERE estado = 'ENVIADA' AND lote_envio IS NOT NULL
        GROUP BY lote_envio
        ORDER BY MAX(fecha_ingreso) DESC
        LIMIT 10
    ''')
    lotes = cursor.fetchall()

    conexion.close()
    return lotes


def obtener_bolsas_por_lote(nombre_lote):
    """
    Devuelve el detalle de todas las bolsas pertenecientes a un lote concreto.
    Se usa para poblar la tabla inferior del historial al seleccionar un lote.
    Retorna: lista de tuplas (id, importe, contrato, nif)
    """
    conexion = sqlite3.connect(ruta_bd())
    cursor   = conexion.cursor()

    cursor.execute('''
        SELECT id, importe, contrato, nif
        FROM bolsas
        WHERE lote_envio = ?
    ''', (nombre_lote,))
    bolsas = cursor.fetchall()

    conexion.close()
    return bolsas


# =============================================================================
#  SECCIÓN 5 — ARQUEO
# =============================================================================

def obtener_datos_arqueo():
    """
    Calcula los tres valores que muestra la barra azul inferior de la UI:
      · desglose_monedas — lista de (denominacion, cantidad_total) del monedero
      · total_bolsas     — suma de importes de bolsas 'IN-CAJA FUERTE'
      · gran_total       — suma de ambos conceptos
    """
    conexion = sqlite3.connect(ruta_bd())
    cursor   = conexion.cursor()

    cursor.execute("SELECT denominacion, cantidad_total FROM monedero")
    desglose_monedas = cursor.fetchall()

    cursor.execute("SELECT SUM(importe) FROM bolsas WHERE estado = 'IN-CAJA FUERTE'")
    resultado_bolsas = cursor.fetchone()[0]
    total_bolsas     = resultado_bolsas if resultado_bolsas is not None else 0.0

    total_suelto = sum(cantidad for _, cantidad in desglose_monedas)
    gran_total   = total_suelto + total_bolsas

    conexion.close()
    return desglose_monedas, total_bolsas, gran_total


# =============================================================================
#  SECCIÓN 6 — SNAPSHOT DIARIO
# =============================================================================

def capturar_snapshot_diario():
    """
    Guarda una copia de los saldos actuales del monedero la primera vez que
    se ejecuta la aplicación en el día. Si ya existe una foto de hoy, no hace nada.
    Esto garantiza que el snapshot refleje siempre el estado de apertura real,
    sin importar cuántas veces se reinicie el programa durante la jornada.
    """
    fecha_hoy = datetime.now().strftime("%Y-%m-%d")

    conexion = sqlite3.connect(ruta_bd())
    cursor   = conexion.cursor()

    # Comprobar si ya existe la foto de hoy antes de escribir nada
    cursor.execute("SELECT 1 FROM snapshot_diario WHERE fecha = ?", (fecha_hoy,))
    if cursor.fetchone():
        conexion.close()
        return

    # Primera apertura del día: copiar el estado actual del monedero
    cursor.execute("SELECT denominacion, cantidad_total FROM monedero")
    saldos_actuales = cursor.fetchall()

    for den, cantidad in saldos_actuales:
        cursor.execute('''
            INSERT INTO snapshot_diario (fecha, denominacion, cantidad_inicial)
            VALUES (?, ?, ?)
        ''', (fecha_hoy, den, cantidad))

    conexion.commit()
    conexion.close()


def obtener_snapshot_hoy():
    """
    Recupera la foto de apertura del día actual como diccionario.
    Retorna: dict {denominacion: cantidad_inicial}
    Si no existe snapshot para hoy (p.ej. primer uso del día sin monedas),
    devuelve un diccionario vacío.
    """
    fecha_hoy = datetime.now().strftime("%Y-%m-%d")

    conexion = sqlite3.connect(ruta_bd())
    cursor   = conexion.cursor()

    cursor.execute('''
        SELECT denominacion, cantidad_inicial
        FROM snapshot_diario
        WHERE fecha = ?
    ''', (fecha_hoy,))
    datos = dict(cursor.fetchall())

    conexion.close()
    return datos


# =============================================================================
#  ARRANQUE DIRECTO
# =============================================================================

if __name__ == '__main__':
    inicializar_db()
    inicializar_snapshot_table()
    print("Base de datos creada y lista para operar.")