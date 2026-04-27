import tkinter as tk
import database
from tkinter import ttk
from datetime import datetime, timedelta
import sqlite3
import csv
from tkinter import filedialog
import sys
import os
import socket


# =============================================================================
#  SECCIÓN 1 — ESTADO GLOBAL
#  Variables compartidas por toda la aplicación. 
# =============================================================================

# Denominaciones de moneda 
DENOMINACIONES = ["2 €", "1 €", "0,50 €", "0,20 €", "0,10 €", "0,05 €", "0,02 €", "0,01 €"]

ESTRUCTURA_MONEDAS = {
    "2 €": (500, 50), "1 €": (250, 25), "0,50 €": (200, 20), "0,20 €": (80, 8),
    "0,10 €": (40, 4), "0,05 €": (25, 2.5), "0,02 €": (10, 1), "0,01 €": (5, 0.5)
}

# --- Estado de la pestaña Monedas ---
variables_saldos    = {den: 0.0 for den in DENOMINACIONES}   # Saldo en memoria por denominación
labels_saldos       = {}                                     # Widgets Label de solo lectura (columna central)
entradas_edicion    = {}                                     # Widgets Entry del modo edición
lista_cajas_input   = []                                     # Widgets Entry de transacciones (+/-) por fila
lista_cajas_edicion = []                                     # Widgets Entry de edición directa por fila
modo_edicion        = False                                  # True = modo edición activo en la pestaña Monedas

# --- Estado de la pestaña Bolsas ---
modo_vista_enviadas     = False   # True = mostrando historial de envíos (en lugar del formulario)
indice_inicio_arrastre  = None    # Fila donde empezó el arrastre de selección en la tabla


# =============================================================================
#  SECCIÓN 2 — UTILIDADES GENÉRICAS
#  Funciones auxiliares de uso transversal: formato, posicionamiento
#  y construcción de ventanas emergentes reutilizables.
# =============================================================================

def ruta_recurso(relativa):
    """
    Resuelve la ruta a un archivo tanto en desarrollo como dentro del .exe.
    PyInstaller extrae los recursos a una carpeta temporal (_MEIPASS) al ejecutar.
    """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relativa)
    base_path = os.path.dirname(os.path.abspath(__file__))
    if os.path.basename(base_path) == 'src':
        base_path = os.path.dirname(base_path)
        
    return os.path.join(base_path, relativa)


def formatear_euro(valor):
    """
    Convierte un número a cadena de divisa.
    Omite decimales si el valor es entero para mayor legibilidad.
      50.0  →  "50 €"
      50.5  →  "50.50 €"
    """
    if valor == int(valor):
        texto = f"{int(valor):,}".replace(',','.')
        return f"{texto} €"
    else:
        texto = f"{valor:,.2f}"
        texto = texto.replace(',','_').replace('.',',').replace('_','.')
        return f"{texto} €"


def centrar_ventana(ventana_hija, ventana_padre, ancho, alto):
    """Calcula el centro de la ventana padre y sitúa la hija exactamente ahí."""
    ventana_padre.update_idletasks()
    x = ventana_padre.winfo_rootx() + (ventana_padre.winfo_width()  // 2) - (ancho // 2)
    y = ventana_padre.winfo_rooty() + (ventana_padre.winfo_height() // 2) - (alto  // 2)
    ventana_hija.geometry(f"{ancho}x{alto}+{x}+{y}")


# --- Popups reutilizables con diseño coherente con el resto de la app ---

def popup_silencioso_askokcancel(titulo, mensaje, parent):
    """
    Diálogo de confirmación (Aceptar / Cancelar) al estilo Windows:
    fondo blanco para el mensaje, franja gris inferior para los botones,
    botón Aceptar en azul a la izquierda del Cancelar gris.
    Devuelve True si el usuario acepta, False si cancela o cierra.
    """
    resultado = tk.BooleanVar(value=False)
    pop = tk.Toplevel(parent)
    pop.withdraw()
    pop.title(titulo)
    pop.configure(bg="#ffffff")
    pop.resizable(False, False)
    pop.transient(parent)
    pop.grab_set()

    # Zona blanca con el mensaje
    frame_mensaje = tk.Frame(pop, bg="#ffffff")
    frame_mensaje.pack(expand=True, fill="both", padx=20, pady=20)
    tk.Label(
        frame_mensaje, text=mensaje,
        font=("Franklin Gothic Demi", 11), bg="#ffffff",
        wraplength=280, justify="left"
    ).pack(anchor="w")

    # Franja gris inferior con los botones
    frame_btns = tk.Frame(pop, bg="#f0f0f0", pady=12, padx=15)
    frame_btns.pack(side="bottom", fill="x")

    def _aceptar(event=None):
        resultado.set(True)
        pop.destroy()

    # Orden Windows: Cancelar a la derecha, Aceptar a su izquierda
    tk.Button(
        frame_btns, text="Cancelar",
        font=("Franklin Gothic Demi", 10), bg="#e1e1e1", fg="#111111",
        relief="flat", bd=0, padx=15, pady=4, cursor="hand2",
        command=pop.destroy
    ).pack(side="right", padx=(10, 0))

    tk.Button(
        frame_btns, text="Aceptar",
        font=("Franklin Gothic Demi", 10, "bold"), bg="#066dff", fg="white",
        relief="flat", bd=0, padx=15, pady=4, cursor="hand2",
        command=_aceptar
    ).pack(side="right")

    pop.bind("<Return>", _aceptar)
    centrar_ventana(pop, parent, 350, 160)
    pop.deiconify()
    pop.focus_force()
    parent.wait_window(pop)
    return resultado.get()


def popup_silencioso_mensaje(titulo, mensaje, parent, es_error=False):
    """
    Diálogo informativo de un solo botón ("Entendido").
    """
    pop = tk.Toplevel(parent)
    pop.withdraw()
    pop.title(titulo)
    pop.configure(bg="#ffffff")
    pop.resizable(False, False)
    pop.transient(parent)
    pop.grab_set()

    color_btn = "#D7263D" if es_error else "#0A8754"

    frame_mensaje = tk.Frame(pop, bg="#ffffff")
    frame_mensaje.pack(expand=True, fill="both", padx=20, pady=20)
    tk.Label(
        frame_mensaje, text=mensaje,
        font=("Franklin Gothic Demi", 11), bg="#ffffff",
        wraplength=280, justify="left"
    ).pack(anchor="w")

    frame_btns = tk.Frame(pop, bg="#f0f0f0", pady=12, padx=15)
    frame_btns.pack(side="bottom", fill="x")

    tk.Button(
        frame_btns, text="Entendido",
        font=("Franklin Gothic Demi", 10, "bold"), bg=color_btn, fg="white",
        relief="flat", bd=0, padx=15, pady=4, cursor="hand2",
        command=pop.destroy
    ).pack(side="right")

    pop.bind("<Return>", lambda e: pop.destroy())
    centrar_ventana(pop, parent, 350, 160)
    pop.deiconify()
    pop.focus_force()
    parent.wait_window(pop)


def popup_silencioso_logs_diarios(parent):
    """
    Muestra los saldos registrados en el snapshot de apertura del día actual.
    Consulta la BD a través de database.obtener_snapshot_hoy().
    """
    pop = tk.Toplevel(parent)
    pop.withdraw()
    pop.title("Registros Varios")
    pop.configure(bg="#ffffff")
    pop.resizable(False, False)
    pop.transient(parent)
    pop.grab_set()

    datos_hoy    = database.obtener_snapshot_hoy()
    fecha_format = datetime.now().strftime("%d/%m/%Y")

    tk.Label(
        pop, text=f"Saldos iniciales ({fecha_format})",
        font=("Franklin Gothic Demi", 11, "bold"), bg="#ffffff", fg="#066dff"
    ).pack(pady=(15, 10))

    frame_lista = tk.Frame(pop, bg="#ffffff")
    frame_lista.pack(padx=30, pady=5)

    for i, den in enumerate(DENOMINACIONES):
        saldo_inicial = datos_hoy.get(den, 0.0)
        tk.Label(
            frame_lista, text=den,
            font=("Franklin Gothic Demi", 10), bg="#ffffff", fg="#555555",
            width=10, anchor="w"
        ).grid(row=i, column=0, pady=2)
        tk.Label(
            frame_lista, text=formatear_euro(saldo_inicial),
            font=("Franklin Gothic Demi", 10), bg="#ffffff", fg="#111111",
            width=12, anchor="e"
        ).grid(row=i, column=1, pady=2)
    
    # Separador visual y Título de sección actual
    tk.Frame(pop, height=1, bg="#e1e1e1").pack(fill="x", padx=30, pady=10)
    tk.Label(pop, text="EN FIJO\nPaquetes (p) y Blísteres (b)", 
             font=("Franklin Gothic Demi", 10, "bold"), bg="#ffffff", fg="#066dff").pack(pady=(0, 10))

    # Frame para las 4 columnas simétricas
    frame_desglose = tk.Frame(pop, bg="#ffffff")
    frame_desglose.pack(padx=20, pady=5)

    for i, den in enumerate(DENOMINACIONES):
        # Cálculo logístico
        p_val, b_val = ESTRUCTURA_MONEDAS[den]
        total_actual = variables_saldos[den]
        paqs = int(total_actual // p_val)
        blist = int((total_actual % p_val) // b_val)
        
        # Posicionamiento en rejilla de 4 columnas (2 filas x 4 columnas)
        fila_d = i // 2
        col_d  = i % 2
    
        tk.Label(
            frame_desglose, text=f"{den}: {paqs}p {blist}b",
            font=("Franklin Gothic Demi", 9), bg="#ffffff", fg="#333333",
            padx=5, pady=2, width=16, anchor="w"
        ).grid(row=fila_d, column=col_d, padx=2, pady=1)

    frame_btns = tk.Frame(pop, bg="#f0f0f0", pady=10, padx=15)
    frame_btns.pack(side="bottom", fill="x", pady=(15, 0))

    tk.Button(
        frame_btns, text="Cerrar",
        font=("Franklin Gothic Demi", 10), bg="#e1e1e1", fg="#111111",
        relief="flat", bd=0, padx=15, pady=4, cursor="hand2",
        command=pop.destroy
    ).pack(side="right")

    pop.bind("<Return>", lambda e: pop.destroy())
    centrar_ventana(pop, parent, 380, 480)
    pop.deiconify()
    pop.focus_force()


# =============================================================================
#  SECCIÓN 3 — VALIDACIONES DE ENTRADA
#  Funciones usadas como 'validatecommand' en los widgets Entry de tkinter.
#  Se llaman de forma proactiva en cada pulsación de tecla (validate="key").
#  Devuelven True para permitir el carácter, False para bloquearlo.
# =============================================================================

def validar_max_18_caracteres(texto_propuesto):
    """Permite cualquier carácter, pero con un límite máximo de 18."""
    return len(texto_propuesto) <= 18


def validar_pulsacion(texto_propuesto):
    """
    Validador para campos monetarios con decimales y signo +/-.
    Permite: dígitos, un punto decimal, coma (→ punto), apóstrofe (→ punto),
    y un signo + o - al inicio. Limita a 6 enteros y 2 decimales.
    """
    if texto_propuesto in ("", "+", "-"):
        return True

    # Normalizar separadores decimales alternativos
    t_norm = texto_propuesto.replace(',', '.').replace("'", '.')

    if t_norm.count('.') > 1:
        return False

    # Analizar solo la parte numérica (sin el signo inicial)
    t_eval = t_norm[1:] if t_norm.startswith(('+', '-')) else t_norm

    for char in t_eval:
        if not char.isdigit() and char != '.':
            return False

    partes = t_eval.split('.')
    if len(partes[0]) > 6:
        return False
    if len(partes) > 1 and len(partes[1]) > 2:
        return False

    return True


def validar_4_digitos(texto_propuesto):
    """Solo dígitos, máximo 4 caracteres. Usado para el campo Oficina."""
    if texto_propuesto == "":
        return True
    return texto_propuesto.isdigit() and len(texto_propuesto) <= 4


def validar_enteros(texto_propuesto):
    """Solo dígitos, máximo 10 caracteres. Usado para el campo Contrato."""
    if texto_propuesto == "":
        return True
    return texto_propuesto.isdigit() and len(texto_propuesto) <= 10


def validar_formato_fecha(texto_propuesto):
    """
    Bloquea letras y cualquier carácter que no sea dígito o '/'.
    Limita la longitud a 10 caracteres (formato DD/MM/YYYY).
    """
    if texto_propuesto == "":
        return True
    if len(texto_propuesto) > 10:
        return False
    for char in texto_propuesto:
        if not char.isdigit() and char != '/':
            return False
    return True


# =============================================================================
#  SECCIÓN 4 — LÓGICA DE NEGOCIO: MONEDAS
#  Gestión completa de la pestaña Monedas:
#    4a) Helpers internos de validación de valores monetarios
#    4b) Carga inicial desde BD
#    4c) Procesamiento de transacciones (+/-)
#    4d) Modo edición directa de totales
#    4e) Actualización de la barra de totales inferior
# =============================================================================

# --- 4a) Helpers de validación monetaria ---

def limpiar_y_validar_formato(texto):
    """
    Limpia separadores alternativos y valida que el texto sea un número
    positivo con máximo 2 decimales y valor entre 0 y 100.000.
    Devuelve el float limpio o lanza ValueError con descripción.
    """
    if not texto:
        return 0.0

    texto = texto.replace(" ", "").replace(',', '.').replace("'", '.')

    if '.' in texto and len(texto.split('.')[1]) > 2:
        raise ValueError("Solo dos decimales")

    valor = float(texto)

    if valor < 0:
        raise ValueError("Negativo imposible")
    if valor > 100000:
        raise ValueError("Límite superado")

    return valor


def validar_valor_moneda(texto, den):
    """
    Además del formato base (limpiar_y_validar_formato), verifica que el
    importe sea múltiplo físico de la denominación indicada.
    Ej: para "0.20 €", solo se aceptan 0.20, 0.40, 0.60...
    """
    valor = limpiar_y_validar_formato(texto)
    if valor == 0:
        return 0.0

    # Obtenemos el valor del blister de nuestra constante (segundo valor de la tupla)
    _, valor_blister      = ESTRUCTURA_MONEDAS[den] 
    input_centimos        = int(round(valor * 100))
    blister_centimos      = int(round(valor_blister * 100))

    if input_centimos % blister_centimos != 0:
        raise ValueError(f"No es múltiplo del blister ({valor_blister} €)")

    return valor


# --- 4b) Carga inicial desde la base de datos ---

def cargar_datos_iniciales():
    """
    Lee los saldos actuales de la BD y los vuelca en los labels de la interfaz.
    Se llama una sola vez al arrancar la aplicación.
    """
    saldos_bd = database.obtener_todos_los_saldos()
    for den, total in saldos_bd:
        if den in variables_saldos:
            variables_saldos[den] = total
            labels_saldos[den].config(text=formatear_euro(total))
    actualizar_totales()


# --- 4c) Procesamiento de transacciones (+/-) ---

def procesar_input_moneda(event, denominacion, entry_widget):
    """
    Controlador del campo de transacción por denominación.
    Lee el texto del Entry, detecta el signo (+/-), valida el valor,
    actualiza el saldo en memoria y persiste el delta en la BD.
    Pinta el campo de rojo si hay error; lo limpia si la operación es válida.
    """
    texto = entry_widget.get().strip()
    if not texto:
        # Campo vacío: restaurar aspecto neutro sin procesar nada
        entry_widget.config(highlightbackground="#cccccc", highlightcolor="#066dff", bg="white")
        return

    try:
        es_suma        = texto.startswith('+')
        texto_limpio   = texto.lstrip('+-')
        valor_numerico = validar_valor_moneda(texto_limpio, denominacion)

        nuevo_saldo = round(
            variables_saldos[denominacion] + valor_numerico if es_suma
            else variables_saldos[denominacion] - valor_numerico,
            2
        )

        if nuevo_saldo < 0:
            raise ValueError("Saldo negativo imposible")
        if nuevo_saldo > 99999.99:
            raise ValueError("Límite máximo superado")

        # Actualizar memoria, BD e interfaz
        variables_saldos[denominacion] = nuevo_saldo
        database.modificar_monedero(denominacion, valor_numerico if es_suma else -valor_numerico)
        labels_saldos[denominacion].config(text=formatear_euro(nuevo_saldo))

        # Limpiar el campo y restaurar aspecto neutro
        entry_widget.delete(0, 'end')
        entry_widget.config(highlightbackground="#cccccc", highlightcolor="#066dff", bg="white")
        actualizar_totales()

    except ValueError:
        # Marcar el campo en rojo y seleccionar el contenido para fácil corrección
        entry_widget.config(highlightbackground="#ff4d4d", highlightcolor="#ff4d4d", bg="#fff0f0")
        entry_widget.select_range(0, 'end')
        return "break"


def saltar_fila(event, direccion, lista_entries, indice_actual):
    """
    Mueve el foco al Entry anterior o siguiente dentro de una lista vertical.
    Usado para navegar con Tab / Shift-Tab / Arriba / Abajo entre filas.
    """
    if direccion == "abajo" and indice_actual < len(lista_entries) - 1:
        lista_entries[indice_actual + 1].focus()
    elif direccion == "arriba" and indice_actual > 0:
        lista_entries[indice_actual - 1].focus()
    return "break"


# --- 4d) Modo edición directa de totales ---

def _marcar_error_edicion(den):
    """Pinta de rojo la caja de edición de la denominación dada."""
    entradas_edicion[den].config(bg="#fff0f0", highlightbackground="#cc0000", highlightcolor="#cc0000")


def _limpiar_error_edicion(den):
    """Restaura el aspecto neutro de la caja de edición de la denominación dada."""
    entradas_edicion[den].config(bg="white", highlightbackground="#cccccc", highlightcolor="#066dff")


def auditar_celda_edicion(event, den):
    """
    Valida en tiempo real la celda de edición al perder el foco (FocusOut / Tab).
    Si el valor no es válido, pinta la celda en rojo y devuelve el foco a ella.
    """
    texto = entradas_edicion[den].get().strip()
    try:
        validar_valor_moneda(texto, den)
        _limpiar_error_edicion(den)
    except ValueError:
        _marcar_error_edicion(den)
        entradas_edicion[den].select_range(0, 'end')


def alternar_modo_edicion():
    """
    Máquina de estados de dos fases para la edición directa de totales:

    LECTURA → EDICIÓN:
      Oculta los Labels y muestra los Entry pre-rellenos con el valor actual.

    EDICIÓN → LECTURA (Guardar):
      Valida todas las celdas. Si hay errores, aborta.
      Si no hay cambios, cancela silenciosamente.
      Si hay cambios, pide confirmación y persiste en BD.
    """
    global modo_edicion

    if not modo_edicion:
        # ── Lectura → Edición ──────────────────────────────────────────────
        modo_edicion = True
        btn_editar.config(
            text="Guardar cambios", relief="solid",
            bg="#066dff", fg="white", font=("Franklin Gothic Demi", 10)
        )

        for caja_input in lista_cajas_input:
            caja_input.config(state="disabled")

        for den in DENOMINACIONES:
            val = variables_saldos[den]
            # Mostrar vacío si es 0; entero si no tiene decimales significativos
            if val == 0:
                texto_mostrar = ""
            elif val == int(val):
                texto_mostrar = str(int(val))
            else:
                texto_mostrar = str(val)

            labels_saldos[den].grid_remove()
            entradas_edicion[den].delete(0, 'end')
            entradas_edicion[den].insert(0, texto_mostrar)
            _limpiar_error_edicion(den)
            entradas_edicion[den].grid()

        # Colocar el foco en la primera fila con todo seleccionado
        entradas_edicion[DENOMINACIONES[0]].focus()
        entradas_edicion[DENOMINACIONES[0]].select_range(0, 'end')

    else:
        # ── Edición → Lectura (Guardar) ────────────────────────────────────
        nuevos_valores = {}
        hay_errores    = False

        for den in DENOMINACIONES:
            _limpiar_error_edicion(den)
            try:
                nuevos_valores[den] = validar_valor_moneda(entradas_edicion[den].get().strip(), den)
            except ValueError:
                _marcar_error_edicion(den)
                hay_errores = True

        if hay_errores:
            return  
        
        sin_cambios = all(nuevos_valores[den] == variables_saldos[den] for den in DENOMINACIONES)
        if sin_cambios:
            for den in DENOMINACIONES:
                entradas_edicion[den].grid_remove()
                labels_saldos[den].grid()
            modo_edicion = False
            btn_editar.config(
                text="Editar totales", relief="flat",
                bg="#f0f0f0", fg="#888888", font=("Franklin Gothic Demi", 10)
            )

            for caja_input in lista_cajas_input:
                caja_input.config(state="normal")


            return

        # Confirmar antes de sobrescribir
        if not popup_silencioso_askokcancel("Confirmar Totales", "¿Quieres cambiar los totales?", ventana):
            resetear_pestana_monedas()
            return

        # Persistir y actualizar la interfaz
        for den in DENOMINACIONES:
            nuevo_total = round(nuevos_valores[den], 2)
            variables_saldos[den] = nuevo_total
            database.guardar_saldo_total(den, nuevo_total)
            labels_saldos[den].config(text=formatear_euro(nuevo_total))
            entradas_edicion[den].grid_remove()
            labels_saldos[den].grid()

        actualizar_totales()
        modo_edicion = False
        btn_editar.config(
            text="Editar totales", relief="flat",
            bg="#f0f0f0", fg="#888888", font=("Franklin Gothic Demi", 10)
        )

        for caja_input in lista_cajas_input:
            caja_input.config(state="normal")


# --- 4e) Barra de totales inferior ---

def actualizar_totales():
    """
    Recalcula y muestra los tres totales de la barra azul inferior:
      · Total Monedas  — suma de variables_saldos (en memoria, tiempo real)
      · Total Bolsas   — bolsas con estado IN-CAJA FUERTE (consulta BD)
      · Gran Total     — suma de ambos
    """
    suma_monedas         = sum(variables_saldos.values())
    _, total_bolsas, _   = database.obtener_datos_arqueo()
    gran_total           = suma_monedas + total_bolsas

    var_total_monedas.set(formatear_euro(suma_monedas))
    var_total_bolsas.set(formatear_euro(total_bolsas))
    var_gran_total.set(formatear_euro(gran_total))


# =============================================================================
#  SECCIÓN 5 — LÓGICA DE NEGOCIO: BOLSAS
#  Gestión completa de la pestaña Bolsas:
#    5a) Registro de nuevas bolsas
#    5b) Envío de bolsas (con popup de fecha)
#    5c) Vista historial (maestro-detalle por lotes)
#    5d) Borrado y reversión
#    5e) Exportación a CSV
#    5f) Selección por arrastre en tablas
# =============================================================================

# --- 5a) Registro de nuevas bolsas ---

def accion_registrar_bolsa():
    """
    Lee el formulario, valida los campos obligatorios (Importe + Contrato)
    y persiste la nueva bolsa en la BD. Limpia el formulario tras el éxito.
    El campo NIF y Oficina son opcionales.
    """
    texto_importe = caja_bolsa_importe.get().strip()
    oficina       = caja_bolsa_oficina.get().strip()
    contrato_num  = caja_bolsa_contrato.get().strip()
    nif           = caja_bolsa_dni.get().strip()

    # Resetear estado visual de los campos obligatorios
    caja_bolsa_importe.config(bg="white",  highlightbackground="#cccccc", highlightcolor="#066dff")
    caja_bolsa_contrato.config(bg="white", highlightbackground="#cccccc", highlightcolor="#066dff")

    if not any([texto_importe, oficina, contrato_num, nif]):
        return

    hay_errores    = False
    importe_limpio = 0.0

    # Importe y Contrato son mutuamente obligatorios
    if not texto_importe:
        caja_bolsa_importe.config(bg="#fff0f0", highlightbackground="#cc0000", highlightcolor="#cc0000")
        hay_errores = True
    if not contrato_num:
        caja_bolsa_contrato.config(bg="#fff0f0", highlightbackground="#cc0000", highlightcolor="#cc0000")
        hay_errores = True

    if texto_importe:
        try:
            importe_limpio = limpiar_y_validar_formato(texto_importe)
        except ValueError:
            caja_bolsa_importe.config(bg="#fff0f0", highlightbackground="#cc0000", highlightcolor="#cc0000")
            hay_errores = True

    if hay_errores:
        return

    # Combinar oficina y contrato en una sola cadena si hay oficina
    contrato_combinado = f"{oficina} - {contrato_num}" if oficina else contrato_num

    try:
        database.registrar_bolsa(nif, contrato_combinado, importe_limpio)
        # Limpiar el formulario y devolver el foco al campo Importe
        caja_bolsa_importe.delete(0, 'end')
        caja_bolsa_oficina.delete(0, 'end')
        caja_bolsa_contrato.delete(0, 'end')
        caja_bolsa_dni.delete(0, 'end')
        caja_bolsa_importe.focus()
        refrescar_tabla_bolsas()
    except Exception as e:
        popup_silencioso_mensaje("Error", f"No se pudo guardar: {e}", ventana, es_error=True)


# --- 5b) Envío de bolsas ---

def formatear_fecha_al_vuelo(event):
    """
    Auto-inserta las barras '/' mientras se escribe en el campo de fecha,
    construyendo el formato DD/MM/YYYY de forma transparente para el usuario.
    Ignora las teclas de borrado y navegación para no interferir con la edición.
    """
    if event.keysym in ('BackSpace', 'Delete', 'Left', 'Right'):
        return

    texto     = caja_fecha.get().replace("/", "")
    nuevo_txt = ""
    for i, char in enumerate(texto[:8]):
        if i in (2, 4):
            nuevo_txt += "/"
        nuevo_txt += char

    if caja_fecha.get() != nuevo_txt:
        caja_fecha.delete(0, 'end')
        caja_fecha.insert(0, nuevo_txt)


def confirmar_envio_lote(ids, fecha_str):
    """
    Valida que la fecha esté dentro de la ventana de ±15 días respecto a hoy
    y, si es válida, registra el envío en la BD y refresca la tabla.
    Se llama tanto desde los botones rápidos (HOY/AYER/MAÑANA) como desde
    el campo de fecha manual.
    """
    try:
        fecha_dt      = datetime.strptime(fecha_str, "%d/%m/%Y")
        hoy           = datetime.now()
        margen_pasado = hoy - timedelta(days=15)
        margen_futuro = hoy + timedelta(days=15)

        if not (margen_pasado <= fecha_dt <= margen_futuro):
            popup_silencioso_mensaje(
                "Error de Fecha",
                "La fecha debe encontrarse en un intervalo de ±15 días respecto a la fecha de hoy.",
                popup, es_error=True
            )
            return

        database.registrar_envio_bolsas(ids, fecha_str)
        popup.destroy()
        refrescar_tabla_bolsas()

    except ValueError:
        popup_silencioso_mensaje("Error", "Formato de fecha inválido.", popup, es_error=True)


def abrir_popup_envio():
    """
    Abre la ventana de confirmación de envío para las bolsas seleccionadas.
    Ofrece tres botones rápidos (HOY / AYER / MAÑANA) y un campo de fecha
    manual con auto-formato. Solo se activa si hay filas seleccionadas.
    """
    global popup, caja_fecha

    seleccion = tabla_bolsas.selection()
    if not seleccion:
        return

    ids_a_enviar = [tabla_bolsas.item(i)['values'][0] for i in seleccion]

    popup = tk.Toplevel(ventana)
    popup.withdraw()
    popup.title("Confirmar Envío")
    popup.configure(bg="#f0f0f0")
    popup.resizable(False, False)
    popup.grab_set()
    popup.transient(ventana)
    popup.protocol("WM_DELETE_WINDOW", popup.destroy)

    centrar_ventana(popup, ventana, 350, 320)
    popup.focus_force()

    palabra_bolsa = "bolsa" if len(ids_a_enviar) == 1 else "bolsas"
    tk.Label(
        popup, text=f"Enviar {len(ids_a_enviar)} {palabra_bolsa}",
        font=("Franklin Gothic Demi", 12, "bold"), bg="#f0f0f0"
    ).pack(pady=(15, 10))

    # Botones de fecha rápida
    hoy   = datetime.now()
    ayer  = hoy - timedelta(days=1)
    manana = hoy + timedelta(days=1)

    for texto, fecha in [("HOY", hoy), ("AYER", ayer), ("MAÑANA", manana)]:
        f_str = fecha.strftime("%d/%m/%Y")
        tk.Button(
            popup, text=texto,
            font=("Franklin Gothic Demi", 10, "bold"), width=25,
            bg="#ffffff", fg="#333333", activebackground="#e6e6e6",
            relief="solid", bd=1, highlightthickness=0, cursor="hand2",
            command=lambda f=f_str: confirmar_envio_lote(ids_a_enviar, f)
        ).pack(pady=4)

    tk.Label(
        popup, text="O introducir manualmente:",
        font=("Franklin Gothic Demi", 9), fg="#555555", bg="#f0f0f0"
    ).pack(pady=(15, 5))

    caja_fecha = tk.Entry(
        popup, font=("Franklin Gothic Demi", 12), justify="center", width=15,
        relief="flat", bd=1, bg="white", fg="#111111",
        highlightthickness=1, highlightbackground="#cccccc", highlightcolor="#066dff",
        validate="key", validatecommand=validacion_fecha
    )
    caja_fecha.pack(pady=5)
    caja_fecha.bind("<KeyRelease>", formatear_fecha_al_vuelo)
    caja_fecha.bind("<Return>", lambda e: confirmar_envio_lote(ids_a_enviar, caja_fecha.get()))
    caja_fecha.bind("<FocusIn>", lambda e: e.widget.after(10, lambda: e.widget.select_range(0, 'end')))

    tk.Button(
        popup, text="Confirmar Fecha",
        font=("Franklin Gothic Demi", 10, "bold"),
        bg="#066dff", fg="white", activebackground="#0555c2", activeforeground="white",
        relief="flat", bd=0, highlightthickness=0, padx=15, pady=4, cursor="hand2",
        command=lambda: confirmar_envio_lote(ids_a_enviar, caja_fecha.get())
    ).pack(pady=(15, 10))

    popup.deiconify()
    popup.grab_set()


# --- 5c) Vista historial: maestro-detalle por lotes ---

def refrescar_tabla_lotes():
    """
    Consulta la BD y rellena la tabla superior del historial con los
    nombres de lote de envíos registrados (una fila por lote).
    """
    for fila in tabla_lotes.get_children():
        tabla_lotes.delete(fila)

    for lote in database.obtener_lotes_enviados():
        if lote[0]:
            tabla_lotes.insert('', 'end', values=(lote[0],))


def cargar_detalle_lote(event):
    """
    Reacción al clic en la tabla de lotes (superior del historial).
    Carga en la tabla inferior todas las bolsas pertenecientes al lote seleccionado.
    """
    seleccion = tabla_lotes.selection()
    if not seleccion:
        return

    nombre_lote = tabla_lotes.item(seleccion[0])['values'][0]

    for fila in tabla_bolsas.get_children():
        tabla_bolsas.delete(fila)

    for bolsa in database.obtener_bolsas_por_lote(nombre_lote):
        tabla_bolsas.insert('', 'end', values=(bolsa[0], formatear_euro(bolsa[1]), bolsa[2], bolsa[3]))


def refrescar_tabla_bolsas():
    """
    Vacía la tabla principal de bolsas y la repuebla con las bolsas
    que actualmente tienen estado IN-CAJA FUERTE. Actualiza los totales.
    """
    for fila in tabla_bolsas.get_children():
        tabla_bolsas.delete(fila)

    for bolsa in database.obtener_bolsas_in_caja():
        tabla_bolsas.insert('', 'end', values=(bolsa[0], formatear_euro(bolsa[1]), bolsa[2], bolsa[3]))

    actualizar_totales()


def alternar_vista_bolsas():
    """
    Alterna entre el formulario de registro (modo por defecto) y el historial
    de envíos por lotes. Actualiza el botón y cambia el frame visible.
    """
    global modo_vista_enviadas

    if not modo_vista_enviadas:
        # ── Registro → Historial ──────────────────────────────────────────
        modo_vista_enviadas = True
        btn_ver_enviadas.config(text="VOLVER AL REGISTRO", bg="#066dff")

        frame_bolsas_form.grid_remove()
        frame_bolsas_lotes.grid(row=0, column=0, sticky='nsew', padx=10, pady=(10, 5))

        for fila in tabla_bolsas.get_children():
            tabla_bolsas.delete(fila)
        refrescar_tabla_lotes()

    else:
        # ── Historial → Registro ──────────────────────────────────────────
        modo_vista_enviadas = False
        btn_ver_enviadas.config(text="BOLSAS ENVIADAS", bg="black")

        frame_bolsas_lotes.grid_remove()
        frame_bolsas_form.grid(row=0, column=0, sticky='nsew', padx=10, pady=(10, 5))

        refrescar_tabla_bolsas()


# --- 5d) Borrado y reversión ---

def accion_borrar_o_revertir(event=None):
    """
    Comportamiento dual según el modo activo:

    MODO REGISTRO  → Borra las bolsas seleccionadas permanentemente de la BD.

    MODO HISTORIAL → Reversión de seguridad: devuelve las bolsas (o el lote
                     entero) al estado IN-CAJA FUERTE. Prioriza la selección
                     de bolsas individuales sobre la selección de lote.
    """
    if not modo_vista_enviadas:
        # ── Borrado definitivo (modo registro) ────────────────────────────
        seleccionados = tabla_bolsas.selection()
        if not seleccionados:
            return

        msg = (
            "¿Borrar permanentemente la bolsa?"
            if len(seleccionados) == 1
            else f"¿Borrar las {len(seleccionados)} bolsas?"
        )
        if not popup_silencioso_askokcancel("Confirmar borrado", msg, ventana):
            return

        for item_id in seleccionados:
            id_bolsa = tabla_bolsas.item(item_id)['values'][0]
            database.borrar_bolsa(id_bolsa)
        refrescar_tabla_bolsas()

    else:
        # ── Reversión (modo historial) ────────────────────────────────────
        seleccion_bolsas = tabla_bolsas.selection()
        seleccion_lote   = tabla_lotes.selection()

        # Prioridad: bolsas individuales > lote completo
        if seleccion_bolsas:
            msg = (
                "¿Devolver esta bolsa a CAJA FUERTE?"
                if len(seleccion_bolsas) == 1
                else "¿Devolver estas bolsas a CAJA FUERTE?"
            )
            if not popup_silencioso_askokcancel("Revertir Bolsa", msg, ventana):
                return

            for item_id in seleccion_bolsas:
                id_bolsa = tabla_bolsas.item(item_id)['values'][0]
                database.revertir_envio_bolsa(id_bolsa)
                tabla_bolsas.delete(item_id)  # Eliminación visual instantánea

            # Recargar tabla superior por si el lote quedó vacío
            refrescar_tabla_lotes()
            actualizar_totales()

        elif seleccion_lote:
            msg = (
                "¿Devolver todo este lote a CAJA FUERTE?"
                if len(seleccion_lote) == 1
                else "¿Devolver estos lotes a CAJA FUERTE?"
            )
            if not popup_silencioso_askokcancel("Revertir Lote", msg, ventana):
                return

            for lote_id in seleccion_lote:
                nombre_lote = tabla_lotes.item(lote_id)['values'][0]
                for bolsa in database.obtener_bolsas_por_lote(nombre_lote):
                    database.revertir_envio_bolsa(bolsa[0])

            refrescar_tabla_lotes()
            for fila in tabla_bolsas.get_children():
                tabla_bolsas.delete(fila)
            actualizar_totales()


# --- 5e) Exportación a CSV ---

def exportar_historico_completo():
    """
    Vuelca toda la tabla de bolsas (todos los estados) a un archivo CSV
    con delimitador ';' (compatible con Excel europeo / codificación UTF-8 BOM).
    Abre el diálogo nativo de guardado para elegir ruta y nombre de archivo.
    """
    ruta = filedialog.asksaveasfilename(
        defaultextension=".csv",
        filetypes=[("Archivo CSV", "*.csv")],
        title="Exportar Historial Completo",
        initialfile=f"Historico_Abaco_{datetime.now().strftime('%d%m%Y')}.csv"
    )

    if not ruta:
        return

    conexion = sqlite3.connect(database.ruta_bd())
    cursor   = conexion.cursor()
    cursor.execute("SELECT * FROM bolsas ORDER BY fecha_ingreso DESC")
    datos    = cursor.fetchall()
    columnas = [desc[0].upper() for desc in cursor.description]
    conexion.close()

    try:
        with open(ruta, mode='w', newline='', encoding='utf-8-sig') as f:
            escritor = csv.writer(f, delimiter=';')
            escritor.writerow(columnas)
            escritor.writerows(datos)
        popup_silencioso_mensaje("Éxito", "Historial exportado correctamente.", ventana)
    except Exception as e:
        popup_silencioso_mensaje("Error", f"Fallo al exportar: {e}", ventana, es_error=True)


# --- 5f) Selección por arrastre en tablas ---

def iniciar_clic_tabla(event):
    """
    Registra la fila inicial del arrastre y permite deseleccionar
    haciendo clic en el área vacía de cualquier tabla.
    En modo historial, un clic en tabla_bolsas deselecciona tabla_lotes.
    """
    global indice_inicio_arrastre
    tabla  = event.widget
    region = tabla.identify_region(event.x, event.y)

    # Ignorar clics en el separador de columnas para no interferir con el resize
    if region == "separator":
        return "break"

    if modo_vista_enviadas and tabla == tabla_bolsas:
        tabla_lotes.selection_remove(tabla_lotes.selection())

    item  = tabla.identify_row(event.y)
    filas = tabla.get_children()

    if not item:
        # Clic en zona vacía → deseleccionar todo
        tabla.selection_remove(tabla.selection())
        indice_inicio_arrastre = len(filas)
    else:
        indice_inicio_arrastre = filas.index(item)


def arrastrar_seleccion(event):
    """
    Extiende la selección de filas mientras se mantiene pulsado el botón
    del ratón, desde la fila inicial del clic hasta la posición actual.
    """
    global indice_inicio_arrastre
    if indice_inicio_arrastre is None:
        return

    tabla        = event.widget
    filas        = tabla.get_children()
    item_actual  = tabla.identify_row(event.y)

    if item_actual:
        indice_actual = filas.index(item_actual)
    else:
        # Cursor fuera de la tabla: anclar al extremo más cercano
        indice_actual = 0 if event.y < 30 else len(filas)

    inicio = min(indice_inicio_arrastre, indice_actual)
    fin    = max(indice_inicio_arrastre, indice_actual)
    tabla.selection_set(filas[inicio:fin + 1])


# =============================================================================
#  SECCIÓN 6 — GESTIÓN DE PESTAÑAS
# =============================================================================

def resetear_pestana_monedas():
    """
    Si el modo edición está activo, lo cancela sin guardar y restaura
    los Labels originales, descartando cualquier cambio no confirmado.
    """
    global modo_edicion
    if modo_edicion:
        modo_edicion = False
        btn_editar.config(
            text="Editar totales", relief="flat",
            bg="#f0f0f0", fg="#888888", font=("Franklin Gothic Demi", 9)
        )
        for den in DENOMINACIONES:
            _limpiar_error_edicion(den)
            entradas_edicion[den].grid_remove()
            labels_saldos[den].grid()
            
        
        for caja_input in lista_cajas_input:
            caja_input.config(state="normal")


def resetear_pestana_bolsas():
    """
    Si el historial está visible, vuelve al formulario de registro
    y recarga la tabla principal de bolsas.
    """
    global modo_vista_enviadas
    if modo_vista_enviadas:
        modo_vista_enviadas = False
        btn_ver_enviadas.config(text="BOLSAS ENVIADAS", bg="black")
        frame_bolsas_lotes.grid_remove()
        frame_bolsas_form.grid(row=0, column=0, sticky='nsew', padx=10, pady=(10, 5))
        refrescar_tabla_bolsas()


def evento_cambio_pestana(event):
    """
    Callback del evento <<NotebookTabChanged>>.
    Resetea ambas pestañas al navegar, independientemente del origen.
    """
    resetear_pestana_monedas()
    resetear_pestana_bolsas()


# =============================================================================
#  SECCIÓN 7 — CONSTRUCCIÓN DE LA INTERFAZ
#    7a) Ventana principal
#    7b) Registro de comandos de validación
#    7c) Barra de totales inferior
#    7d) Gestor de pestañas (Notebook)
#    7e) Pestaña Monedas
#    7f) Pestaña Bolsas
# =============================================================================

# --- 7a) Ventana principal ---

ventana = tk.Tk()
ventana.withdraw()   # Ocultar hasta que todo esté construido
ventana.title("Abaco")
ventana.configure(bg="#f0f0f0")
ventana.resizable(False, False)
# <inciso> ruta relativa para logo.png
ruta_logo = ruta_recurso('img/logo.png')
icono = tk.PhotoImage(file=ruta_logo)
# </inciso>
ventana.iconphoto(True, icono)

# Centrar en cualquier monitor
ancho_ventana, alto_ventana = 700, 500
ancho_pantalla = ventana.winfo_screenwidth()
alto_pantalla  = ventana.winfo_screenheight()
x_ventana      = (ancho_pantalla // 2) - (ancho_ventana // 2)
y_ventana      = (alto_pantalla  // 2) - (alto_ventana  // 2)
ventana.geometry(f"{ancho_ventana}x{alto_ventana}+{x_ventana}+{y_ventana}")


# --- 7b) Registro de comandos de validación ---
# Se registran aquí, tras crear la ventana raíz, para que puedan
# referenciarse en los parámetros 'validatecommand' de los Entry.

validacion_estricta  = (ventana.register(validar_pulsacion),      '%P')
validacion_4_digitos = (ventana.register(validar_4_digitos),      '%P')
validacion_enteros   = (ventana.register(validar_enteros),        '%P')
validacion_fecha     = (ventana.register(validar_formato_fecha),  '%P')
validacion_nif       = (ventana.register(validar_max_18_caracteres), '%P')

# --- 7c) Barra de totales inferior ---

barra_inferior = tk.Frame(ventana, bg="#066dff", pady=10)
barra_inferior.pack(side="bottom", fill="x")

var_total_monedas = tk.StringVar(value="0 €")
var_total_bolsas  = tk.StringVar(value="0 €")
var_gran_total    = tk.StringVar(value="0 €")

for i in range(6):
    barra_inferior.columnconfigure(i, weight=1)

tk.Label(barra_inferior, text="MONEDAS:", fg="#000000", bg="#066dff",
         font=("Franklin Gothic Demi", 11, "bold")).grid(row=0, column=0, sticky="e", padx=2)
tk.Label(barra_inferior, textvariable=var_total_monedas, fg="#000000", bg="#066dff",
         font=("Franklin Gothic Demi", 12, "bold")).grid(row=0, column=1, sticky="w")

tk.Label(barra_inferior, text="BOLSAS:",  fg="#000000", bg="#066dff",
         font=("Franklin Gothic Demi", 11, "bold")).grid(row=0, column=2, sticky="e", padx=2)
tk.Label(barra_inferior, textvariable=var_total_bolsas, fg="#000000", bg="#066dff",
         font=("Franklin Gothic Demi", 12, "bold")).grid(row=0, column=3, sticky="w")

tk.Label(barra_inferior, text="TOTAL:",   fg="#000000", bg="#066dff",
         font=("Franklin Gothic Demi", 11, "bold")).grid(row=0, column=4, sticky="e", padx=2)
tk.Label(barra_inferior, textvariable=var_gran_total,   fg="#000000", bg="#066dff",
         font=("Franklin Gothic Demi", 12, "bold")).grid(row=0, column=5, sticky="w")


# --- 7d) Gestor de pestañas (Notebook) ---

estilo = ttk.Style()
estilo.configure('TNotebook.Tab', font=('Franklin Gothic Demi', 12, 'bold'), padding=[15, 5])
estilo.layout("Tab", [
    ('Notebook.tab', {'sticky': 'nswe', 'children': [
        ('Notebook.padding', {'side': 'top', 'sticky': 'nswe', 'children': [
            ('Notebook.label', {'side': 'top', 'sticky': ''})
        ]})
    ]})
])
estilo.configure("Treeview",         font=("Franklin Gothic Demi", 12), rowheight=30)
estilo.configure("Treeview.Heading", font=("Franklin Gothic Demi", 11, "bold"))

panel_pestanas = ttk.Notebook(ventana)
panel_pestanas.pack(side="top", expand=True, fill="both", padx=0, pady=10)

pestana_monedas = ttk.Frame(panel_pestanas)
pestana_bolsas  = ttk.Frame(panel_pestanas)

panel_pestanas.add(pestana_monedas, text="MONEDAS")
panel_pestanas.add(pestana_bolsas,  text="BOLSAS")

panel_pestanas.bind("<<NotebookTabChanged>>", evento_cambio_pestana)


# --- 7e) Pestaña Monedas ---

pestana_monedas.columnconfigure(0, weight=1)
pestana_monedas.rowconfigure(0, weight=1)
pestana_monedas.rowconfigure(1, weight=0)

frame_contenido = tk.Frame(pestana_monedas, bg="#f0f0f0")
frame_contenido.grid(row=0, column=0, sticky='nsew')
frame_contenido.columnconfigure(0, weight=1)
frame_contenido.columnconfigure(1, weight=1)
frame_contenido.columnconfigure(2, weight=1)
for fila in range(9):
    frame_contenido.rowconfigure(fila, weight=1)

# Una fila por denominación: Label denominación | Label/Entry saldo | Entry transacción
for i, den in enumerate(DENOMINACIONES):
    fila = i + 1

    # Columna izquierda: nombre de la denominación
    tk.Label(
        frame_contenido, text=den,
        font=("Franklin Gothic Demi", 16)
    ).grid(row=fila, column=0, pady=8)

    # Columna central (modo lectura): label con el saldo actual
    lbl = tk.Label(
        frame_contenido, text=formatear_euro(0),
        font=("Franklin Gothic Demi", 14), width=10, anchor="e"
    )
    lbl.grid(row=fila, column=1, pady=8)
    labels_saldos[den] = lbl
    lbl.bind("<Double-Button-1>", lambda e: alternar_modo_edicion() if not modo_edicion else None)

    # Columna central (modo edición): Entry para editar el total directamente
    caja_edicion = tk.Entry(
        frame_contenido, font=("Franklin Gothic Demi", 14), width=10, justify="right",
        relief="flat", bd=1, bg="white", fg="#111111",
        highlightthickness=1, highlightbackground="#cccccc",
        validate="key", validatecommand=validacion_estricta
    )
    caja_edicion.grid(row=fila, column=1, pady=8)
    caja_edicion.grid_remove()   # Oculto por defecto; visible solo en modo edición
    entradas_edicion[den]  = caja_edicion
    lista_cajas_edicion.append(caja_edicion)

    # Columna derecha: Entry de transacciones +/- en tiempo real
    caja_input = tk.Entry(
        frame_contenido, font=("Franklin Gothic Demi", 14), width=10, justify="center",
        relief="flat", bd=1, bg="white", fg="#111111",
        highlightthickness=1, highlightbackground="#cccccc", highlightcolor="#066dff",
        validate="key", validatecommand=validacion_estricta
    )
    caja_input.grid(row=fila, column=2, pady=8)
    lista_cajas_input.append(caja_input)

    # Binds del campo de transacción
    caja_input.bind("<Return>",
        lambda e, d=den, c=caja_input: procesar_input_moneda(e, d, c))
    caja_input.bind("<Tab>",
        lambda e, d=den, c=caja_input, idx=i:
            procesar_input_moneda(e, d, c) or saltar_fila(e, "abajo", lista_cajas_input, idx))
    caja_input.bind("<Shift-Tab>",
        lambda e, d=den, c=caja_input, idx=i:
            procesar_input_moneda(e, d, c) or saltar_fila(e, "arriba", lista_cajas_input, idx))
    caja_input.bind("<Down>",   lambda e, idx=i: saltar_fila(e, "abajo",  lista_cajas_input, idx))
    caja_input.bind("<Up>",     lambda e, idx=i: saltar_fila(e, "arriba", lista_cajas_input, idx))
    caja_input.bind("<FocusOut>",
        lambda e, c=caja_input: c.config(highlightbackground="#cccccc", highlightcolor="#066dff", bg="white"))
    caja_input.bind("<FocusIn>",  lambda e: e.widget.after(10, lambda: e.widget.select_range(0, 'end')))

    # Binds del campo de edición directa
    caja_edicion.bind("<Tab>",
        lambda e, d=den, idx=i:
            auditar_celda_edicion(e, d) or saltar_fila(e, "abajo", lista_cajas_edicion, idx))
    caja_edicion.bind("<Shift-Tab>",
        lambda e, idx=i: saltar_fila(e, "arriba", lista_cajas_edicion, idx))
    caja_edicion.bind("<Return>",   lambda e: alternar_modo_edicion())
    caja_edicion.bind("<FocusOut>", lambda e, d=den: auditar_celda_edicion(e, d))
    caja_edicion.bind("<FocusIn>",  lambda e: e.widget.after(10, lambda: e.widget.select_range(0, 'end')))

# Botones de acción (centrados con columnspan=3 para simetría absoluta)
frame_acciones_monedas = tk.Frame(frame_contenido, bg="#f0f0f0")
frame_acciones_monedas.grid(row=len(DENOMINACIONES) + 1, column=0, columnspan=3, pady=(25, 10))

btn_editar = tk.Button(
    frame_acciones_monedas, text="Editar totales",
    font=("Franklin Gothic Demi", 10), fg="#888888", bg="#f0f0f0",
    relief="flat", overrelief="flat", activebackground="#f0f0f0", activeforeground="#555555",
    bd=0, padx=5, pady=0, cursor="hand2",
    command=alternar_modo_edicion
)
btn_editar.pack(side="left")

btn_logs_diarios = tk.Button(
    frame_acciones_monedas, text="📜",
    font=("Segoe UI Emoji", 11), fg="#888888", bg="#f0f0f0",
    relief="flat", overrelief="flat", activebackground="#f0f0f0", activeforeground="#555555",
    bd=0, padx=0, pady=0, cursor="hand2",
    command=lambda: popup_silencioso_logs_diarios(ventana)
)
btn_logs_diarios.pack(side="left")


# --- 7f) Pestaña Bolsas ---

pestana_bolsas.rowconfigure(0, weight=0)
pestana_bolsas.rowconfigure(1, weight=1)
pestana_bolsas.rowconfigure(2, weight=0)
pestana_bolsas.columnconfigure(0, weight=1)

# ── Frame superior: formulario de registro (modo por defecto) ─────────────────

frame_bolsas_form = tk.Frame(pestana_bolsas)
frame_bolsas_form.grid(row=0, column=0, sticky='nsew', padx=10, pady=(10, 5))
frame_bolsas_form.columnconfigure(0, weight=1)
frame_bolsas_form.columnconfigure(1, weight=0)
frame_bolsas_form.columnconfigure(2, weight=1)

frame_inputs = tk.Frame(frame_bolsas_form)
frame_inputs.grid(row=1, column=1, pady=5)

tk.Label(frame_inputs, text="Importe:", font=("Franklin Gothic Demi", 12)).pack(side="left", padx=(0, 4))
caja_bolsa_importe = tk.Entry(
    frame_inputs, font=("Franklin Gothic Demi", 12), justify="right", width=10,
    relief="flat", bd=1, bg="white", fg="#111111",
    highlightthickness=1, highlightbackground="#cccccc", highlightcolor="#066dff",
    validate="key", validatecommand=validacion_estricta
)
caja_bolsa_importe.pack(side="left")

tk.Label(frame_inputs, text=" | ", font=("Franklin Gothic Demi", 12), fg="#888888").pack(side="left", padx=6)

tk.Label(frame_inputs, text="Contrato:", font=("Franklin Gothic Demi", 12)).pack(side="left", padx=(0, 4))
caja_bolsa_oficina = tk.Entry(
    frame_inputs, font=("Franklin Gothic Demi", 12), justify="center", width=6,
    relief="flat", bd=1, bg="white", fg="#111111",
    highlightthickness=1, highlightbackground="#cccccc", highlightcolor="#066dff",
    validate="key", validatecommand=validacion_4_digitos
)
caja_bolsa_oficina.pack(side="left")

tk.Label(frame_inputs, text=" - ", font=("Franklin Gothic Demi", 12)).pack(side="left", padx=2)

caja_bolsa_contrato = tk.Entry(
    frame_inputs, font=("Franklin Gothic Demi", 12), width=10,
    relief="flat", bd=1, bg="white", fg="#111111",
    highlightthickness=1, highlightbackground="#cccccc", highlightcolor="#066dff",
    validate="key", validatecommand=validacion_enteros
)
caja_bolsa_contrato.pack(side="left")

tk.Label(frame_inputs, text=" | ", font=("Franklin Gothic Demi", 12), fg="#888888").pack(side="left", padx=6)

tk.Label(frame_inputs, text="NIF:", font=("Franklin Gothic Demi", 12)).pack(side="left", padx=(0, 4))
caja_bolsa_dni = tk.Entry(
    frame_inputs, font=("Franklin Gothic Demi", 12), width=12,
    relief="flat", bd=1, bg="white", fg="#111111",
    highlightthickness=1, highlightbackground="#cccccc", highlightcolor="#066dff",
    validate="key", validatecommand=validacion_nif
)
caja_bolsa_dni.pack(side="left")

# Autoselección de texto al enfocar cualquier campo del formulario
for caja in (caja_bolsa_importe, caja_bolsa_oficina, caja_bolsa_contrato, caja_bolsa_dni):
    caja.bind("<FocusIn>", lambda e: e.widget.after(10, lambda: e.widget.select_range(0, 'end')))

# Botones de acción del formulario
frame_botones_accion = tk.Frame(frame_bolsas_form)
frame_botones_accion.grid(row=2, column=0, columnspan=3, pady=(15, 5))

tk.Button(
    frame_botones_accion, text="➕",
    font=("Segoe UI Emoji", 14), fg="#066dff", bg="#f0f0f0", activebackground="#f0f0f0",
    relief="flat", overrelief="flat", bd=0, highlightthickness=0, cursor="hand2", width=4,
    command=accion_registrar_bolsa
).pack(side="left", padx=10)

tk.Button(
    frame_botones_accion, text="➡️",
    font=("Segoe UI Emoji", 14), fg="#0F8A5F", bg="#f0f0f0", activebackground="#f0f0f0",
    relief="flat", overrelief="flat", bd=0, highlightthickness=0, cursor="hand2", width=4,
    command=abrir_popup_envio
).pack(side="left", padx=10)

btn_borrar_bolsas = tk.Button(
    frame_botones_accion, text="➖",
    font=("Segoe UI Emoji", 14), fg="#D7263D", bg="#f0f0f0", activebackground="#f0f0f0",
    relief="flat", overrelief="flat", bd=0, highlightthickness=0, cursor="hand2", width=4,
    command=accion_borrar_o_revertir
)
btn_borrar_bolsas.pack(side="left", padx=10)

# ── Frame superior alternativo: historial de lotes enviados ──────────────────
# Se muestra en lugar del formulario al pulsar "BOLSAS ENVIADAS".

frame_bolsas_lotes = tk.Frame(pestana_bolsas)
frame_bolsas_lotes.rowconfigure(0, weight=1)
frame_bolsas_lotes.columnconfigure(0, weight=1)

tabla_lotes = ttk.Treeview(
    frame_bolsas_lotes, columns=("lote",), show="headings",
    height=4, selectmode="browse"
)
tabla_lotes.heading("lote", text="HISTORIAL DE ENVÍOS", anchor="center")
tabla_lotes.column("lote", anchor="center")

scrollbar_lotes = ttk.Scrollbar(frame_bolsas_lotes, orient="vertical", command=tabla_lotes.yview)
tabla_lotes.configure(yscrollcommand=scrollbar_lotes.set)
tabla_lotes.grid(row=0, column=0, sticky="nsew")
scrollbar_lotes.grid(row=0, column=1, sticky="ns")
tabla_lotes.bind("<<TreeviewSelect>>", cargar_detalle_lote)

tk.Button(
    frame_bolsas_lotes, text="Exportar Todo",
    font=("Franklin Gothic Demi", 9), 
    fg="#888888", 
    bg="#f0f0f0",
    cursor="hand2",
    command=exportar_historico_completo,
    highlightthickness=0,
    relief="flat",
    overrelief="flat",
    bd=0,
    activebackground="#f0f0f0"
).grid(row=1, column=0, sticky="e", pady=2)

# ── Frame inferior: tabla principal de bolsas (compartida por ambos modos) ───

frame_bolsas_tabla = tk.Frame(pestana_bolsas, bg="#cce5ff")
frame_bolsas_tabla.grid(row=1, column=0, sticky='nsew', padx=10, pady=(5, 10))
frame_bolsas_tabla.rowconfigure(0, weight=1)
frame_bolsas_tabla.columnconfigure(0, weight=1)

tabla_bolsas = ttk.Treeview(
    frame_bolsas_tabla,
    columns=("id", "importe", "contrato", "nif"),
    show="headings",
    displaycolumns=("importe", "contrato", "nif"),   # La columna id queda oculta
    selectmode="extended"
)

tabla_bolsas.heading("importe",  text="IMPORTE",  anchor="center", command=lambda: None)
tabla_bolsas.heading("contrato", text="CONTRATO", anchor="center", command=lambda: None)
tabla_bolsas.heading("nif",      text="NIF",      anchor="center", command=lambda: None)

tabla_bolsas.column("importe",  stretch=False, anchor="e")
tabla_bolsas.column("contrato", stretch=False, anchor="center")
tabla_bolsas.column("nif",      stretch=False, anchor="center")

def ajustar_anchos_tabla(event=None):
    """Distribuye el ancho total de la tabla entre las tres columnas visibles."""
    ancho_total = tabla_bolsas.winfo_width()
    if ancho_total < 50:
        return
    tabla_bolsas.column("importe",  width=int(ancho_total * 0.30))
    tabla_bolsas.column("contrato", width=int(ancho_total * 0.40))
    tabla_bolsas.column("nif",      width=int(ancho_total * 0.30))

tabla_bolsas.bind("<Configure>", ajustar_anchos_tabla)

scrollbar = ttk.Scrollbar(frame_bolsas_tabla, orient="vertical", command=tabla_bolsas.yview)
tabla_bolsas.configure(yscrollcommand=scrollbar.set)
tabla_bolsas.grid(row=0, column=0, sticky="nsew")
scrollbar.grid(row=0, column=1, sticky="ns")

# Binds de interacción con las tablas
tabla_lotes.bind("<ButtonPress-1>", iniciar_clic_tabla)
tabla_lotes.bind("<Delete>",        accion_borrar_o_revertir)
tabla_lotes.bind("<BackSpace>",     accion_borrar_o_revertir)

tabla_bolsas.bind("<ButtonPress-1>", iniciar_clic_tabla)
tabla_bolsas.bind("<B1-Motion>",     arrastrar_seleccion)
tabla_bolsas.bind("<Delete>",        accion_borrar_o_revertir)
tabla_bolsas.bind("<BackSpace>",     accion_borrar_o_revertir)


# ── Botón de alternancia de vista (debajo de la tabla, fila 2) ───────────────

btn_ver_enviadas = tk.Button(
    pestana_bolsas,
    text="BOLSAS ENVIADAS",
    font=("Franklin Gothic Demi", 10, "bold"),
    bg="black", fg="white",
    activebackground="#666666", activeforeground="white",
    relief="solid", bd=1, highlightthickness=0,
    cursor="hand2", takefocus=0,
    command=alternar_vista_bolsas
)
btn_ver_enviadas.grid(row=2, column=0, pady=(5, 10))



# =============================================================================
#  SECCIÓN 8 — ARRANQUE
# =============================================================================

if __name__ == '__main__':
    # Evitar múltiples instancias
    try:
        lock_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        lock_socket.bind(("127.0.0.1", 65432))
    except socket.error:
        sys.exit() 
    database.inicializar_db()
    database.inicializar_snapshot_table()
    database.capturar_snapshot_diario()
    cargar_datos_iniciales()
    refrescar_tabla_bolsas()
    ventana.deiconify()   # Mostrar la ventana solo cuando todo está listo
    ventana.mainloop()