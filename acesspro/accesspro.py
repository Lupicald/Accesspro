import threading
import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk, simpledialog
import gspread
import os
import time
import datetime
import pandas as pd
import json
import sys
from zk import ZK

# --- LIBRERÍAS GOOGLE OAUTH ---
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

# ==========================================
# ⚙️ CONFIGURACIÓN GLOBAL
# ==========================================
ARCHIVO_CONFIG = "config_app.json"
ARCHIVO_EXCEL_LOCAL = "Reporte_Asistencia.xlsx"
NOMBRE_HOJA_NUBE = "Asistencia_ZKTeco"

# Nuevos archivos
ARCHIVO_USUARIOS = "usuarios_config.json"   # perfiles editables
ARCHIVO_ESTADOS = "estados.json"            # persistir últimos estados

# --- MEMORIA RAM (Para evitar duplicados) ---
HISTORIAL_PROCESADO = set()
ESTADOS_USUARIOS = {}  # { uid: {"nombre": str, "tipo": str, "ultimo_estado": str, "ultima_actividad": "YYYY-MM-DD HH:MM:SS", "alerta": bool} }

HORARIOS_CONFIG = {
    'entrada': datetime.time(9, 0),
    'tolerancia': 15,
    'salida': datetime.time(18, 0)
}

# ALERTAS: visitantes sin salida después de X horas
ALERTA_HORAS_SIN_SALIDA = 4
ALERTA_CHECK_SECONDS = 60  # cada cuánto checar visitantes

SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

# ==========================================
# 🛠️ FUNCIONES DE UTILIDAD
# ==========================================
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def cargar_config():
    if os.path.exists(ARCHIVO_CONFIG):
        try:
            with open(ARCHIVO_CONFIG, 'r') as f:
                return json.load(f)
        except:
            pass
    return {"ip": "192.168.1.201", "sucursal": "Matriz"}

def guardar_config(ip, sucursal):
    try:
        data = {"ip": ip, "sucursal": sucursal}
        with open(ARCHIVO_CONFIG, 'w') as f:
            json.dump(data, f)
    except:
        pass

# ==========================================
# 🗂️ CREAR ARCHIVOS NECESARIOS (si faltan)
# ==========================================
def ensure_files_exist():
    """Crea archivos base si no existen para que la app sea portable/exe-friendly."""
    # usuarios_config.json
    if not os.path.exists(ARCHIVO_USUARIOS):
        try:
            ejemplo = {
                "1": {"nombre": "Admin Ejemplo", "tipo": "empleado", "hora_entrada": "09:00", "hora_salida": "18:00"}
            }
            with open(ARCHIVO_USUARIOS, 'w', encoding='utf-8') as f:
                json.dump(ejemplo, f, indent=2, ensure_ascii=False)
        except:
            pass
    # estados.json
    if not os.path.exists(ARCHIVO_ESTADOS):
        try:
            with open(ARCHIVO_ESTADOS, 'w', encoding='utf-8') as f:
                json.dump({}, f, indent=2, ensure_ascii=False)
        except:
            pass
    # config_app.json
    if not os.path.exists(ARCHIVO_CONFIG):
        try:
            with open(ARCHIVO_CONFIG, 'w', encoding='utf-8') as f:
                json.dump({"ip": "192.168.1.201", "sucursal": "Matriz"}, f, indent=2)
        except:
            pass
    # Reporte Excel: opcional, no lo creamos vacío aquí (se crea al guardar).

# ==========================================
# 👥 GESTIÓN DE USUARIOS (JSON editable)
# ==========================================
def cargar_usuarios():
    """Carga o crea archivo de usuarios. Formato:
    {
        "1": {"nombre":"Luis","tipo":"empleado","hora_entrada":"09:00","hora_salida":"18:00"},
        "99": {"nombre":"Visitante X","tipo":"visitante"}
    }
    """
    if os.path.exists(ARCHIVO_USUARIOS):
        try:
            with open(ARCHIVO_USUARIOS, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return {}
    else:
        return {}

def guardar_usuarios(usuarios):
    try:
        with open(ARCHIVO_USUARIOS, 'w', encoding='utf-8') as f:
            json.dump(usuarios, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        print("Error guardando usuarios:", e)
        return False

def cargar_estados():
    global ESTADOS_USUARIOS
    if os.path.exists(ARCHIVO_ESTADOS):
        try:
            with open(ARCHIVO_ESTADOS, 'r', encoding='utf-8') as f:
                ESTADOS_USUARIOS = json.load(f)
        except:
            ESTADOS_USUARIOS = {}
    else:
        ESTADOS_USUARIOS = {}

def guardar_estados():
    try:
        with open(ARCHIVO_ESTADOS, 'w', encoding='utf-8') as f:
            json.dump(ESTADOS_USUARIOS, f, indent=2, ensure_ascii=False)
    except Exception as e:
        print("Error guardando estados:", e)

def sync_nombres_con_usuarios(usuarios_local):
    """
    Si el administrador cambia el nombre o tipo de un UID en usuarios_config.json,
    sincronizamos ESTADOS_USUARIOS para que muestre el nombre y tipo configurado.
    """
    changed = False
    for uid, info in list(ESTADOS_USUARIOS.items()):
        if uid in usuarios_local:
            nombre_conf = usuarios_local[uid].get("nombre")
            tipo_conf = usuarios_local[uid].get("tipo")
            updated = False
            if nombre_conf and ESTADOS_USUARIOS[uid].get("nombre") != nombre_conf:
                ESTADOS_USUARIOS[uid]["nombre"] = nombre_conf
                updated = True
            if tipo_conf and ESTADOS_USUARIOS[uid].get("tipo") != tipo_conf:
                ESTADOS_USUARIOS[uid]["tipo"] = tipo_conf
                updated = True
            if updated:
                changed = True
    if changed:
        guardar_estados()

# ==========================================
# 🧠 LÓGICA ANTI-DUPLICADOS (NUEVO)
# ==========================================
def cargar_historial_existente(log_func):
    """ Lee el Excel local al inicio para saber qué ya tenemos """
    global HISTORIAL_PROCESADO
    if os.path.exists(ARCHIVO_EXCEL_LOCAL):
        try:
            df = pd.read_excel(ARCHIVO_EXCEL_LOCAL)
            for index, row in df.iterrows():
                llave = f"{row['ID']}_{row['Fecha']}"
                HISTORIAL_PROCESADO.add(llave)
            log_func(f"🧠 Memoria cargada: {len(HISTORIAL_PROCESADO)} registros previos ignorados.")
        except Exception as e:
            log_func(f"⚠️ No se pudo leer historial previo: {e}")

# ==========================================
# 🧠 LÓGICA DE NEGOCIO (EL JUEZ)
# ==========================================
def analizar_registro(uid, fecha, punch, usuarios_local):
    """
    Determina modo (Entrada/Salida), estado (A tiempo / Retardo / Jornada cumplida / Salida anticipada)
    usando horario global por defecto o horario individual si existe en usuarios_local.
    """
    hora_registro = fecha.time()
    modo = "Entrada" if punch in [0, 4, 255] else "Salida"

    if punch == 255:
        # heurística: por la mañana es entrada, por la tarde salida
        if hora_registro < datetime.time(12, 0):
            modo = "Entrada"
        else:
            modo = "Salida"

    estado = "Normal"

    # Obtener horario preferente: usuario o global
    horario_entrada = HORARIOS_CONFIG['entrada']
    horario_salida = HORARIOS_CONFIG['salida']
    tolerancia = HORARIOS_CONFIG['tolerancia']

    user_info = usuarios_local.get(uid)
    if user_info:
        he = user_info.get("hora_entrada")
        hs = user_info.get("hora_salida")
        try:
            if he:
                h, m = map(int, he.split(":"))
                horario_entrada = datetime.time(h, m)
            if hs:
                h2, m2 = map(int, hs.split(":"))
                horario_salida = datetime.time(h2, m2)
        except:
            pass

    if modo == "Entrada":
        limite = (datetime.datetime.combine(datetime.date.today(), horario_entrada) +
                  datetime.timedelta(minutes=tolerancia)).time()
        estado = "✅ A Tiempo" if hora_registro <= limite else "⚠️ Retardo"
    elif modo == "Salida":
        estado = "✅ Jornada Cumplida" if hora_registro >= horario_salida else "⚠️ Salida Anticipada"

    return modo, estado

def actualizar_estado_usuario(uid, nombre, tipo, modo, fecha_str):
    """ Actualiza ESTADOS_USUARIOS y persiste """
    # nombre ya debe venir con preferencia a usuarios_config si aplica
    ESTADOS_USUARIOS[uid] = {
        "nombre": nombre,
        "tipo": tipo,
        "ultimo_estado": modo,
        "ultima_actividad": fecha_str,
        "alerta": ESTADOS_USUARIOS.get(uid, {}).get("alerta", False)
    }
    guardar_estados()

def guardar_excel_local(datos):
    """
    Ahora columnas extendidas:
    ID, Nombre, Fecha, Modo, Estado, Sucursal, Tipo, Ultimo_Estado, Ultima_Actividad
    datos viene como [uid, nom, fecha_str, modo, est, sucursal, tipo, modo, fecha_str]
    """
    try:
        col = ["ID", "Nombre", "Fecha", "Modo", "Estado", "Sucursal", "Tipo", "Ultimo_Estado", "Ultima_Actividad"]
        df_new = pd.DataFrame(datos, columns=col)
        if os.path.exists(ARCHIVO_EXCEL_LOCAL):
            df_old = pd.read_excel(ARCHIVO_EXCEL_LOCAL)
            df_final = pd.concat([df_old, df_new], ignore_index=True)
        else:
            df_final = df_new
        df_final.to_excel(ARCHIVO_EXCEL_LOCAL, index=False)
        return True
    except Exception as e:
        print("Error guardando excel:", e)
        return False

# ==========================================
# 🔐 CONEXIÓN GOOGLE
# ==========================================
def conectar_google(log_func):
    creds = None
    token_path = resource_path('token.json')
    client_secret_path = resource_path('client_secret.json')

    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except:
                creds = None
        if not creds:
            if os.path.exists(client_secret_path):
                log_func("🌐 Iniciando sesión Google...")
                flow = InstalledAppFlow.from_client_secrets_file(client_secret_path, SCOPES)
                creds = flow.run_local_server(port=0)
                try:
                    with open(token_path, 'w') as token:
                        token.write(creds.to_json())
                except:
                    pass
            else:
                log_func("❌ ERROR: Falta 'client_secret.json'")
                return None
    try:
        client = gspread.authorize(creds)
        try:
            sheet = client.open(NOMBRE_HOJA_NUBE).sheet1
            # asegurarnos de encabezados (en caso de que falten columnas nuevas)
            current = sheet.row_values(1)
            required = ["ID", "Nombre", "Fecha y Hora", "Evento", "Análisis", "Sucursal", "Tipo", "Ultimo_Estado", "Ultima_Actividad"]
            if current != required:
                try:
                    # reescribir cabecera (si es necesario)
                    sheet.delete_rows(1)
                    sheet.insert_row(required, 1)
                except:
                    pass
            return sheet
        except Exception:
            sh = client.create(NOMBRE_HOJA_NUBE)
            sh.sheet1.append_row(["ID", "Nombre", "Fecha y Hora", "Evento", "Análisis", "Sucursal", "Tipo", "Ultimo_Estado", "Ultima_Actividad"])
            return sh.sheet1
    except Exception as e:
        log_func(f"Error Google: {e}")
        return None

# ==========================================
# 🔔 Monitor de visitantes (alertas)
# ==========================================
def monitor_visitantes(log_func, stop_event):
    """Revisa periódicamente visitantes/reclusos que estén 'Dentro' sin registrar salida > ALERTA_HORAS_SIN_SALIDA."""
    while not stop_event.is_set():
        try:
            now = datetime.datetime.now()
            changed = False
            for uid, info in list(ESTADOS_USUARIOS.items()):
                tipo = info.get("tipo", "")
                ultimo_estado = info.get("ultimo_estado", "")
                ultima_actividad = info.get("ultima_actividad")
                alerta = info.get("alerta", False)

                if tipo in ("visitante", "recluso") and ultimo_estado == "Entrada" and ultima_actividad:
                    try:
                        dt = datetime.datetime.strptime(ultima_actividad, "%Y-%m-%d %H:%M:%S")
                        diff_hours = (now - dt).total_seconds() / 3600.0
                        if diff_hours >= ALERTA_HORAS_SIN_SALIDA and not alerta:
                            # marcar alerta y loggear
                            ESTADOS_USUARIOS[uid]["alerta"] = True
                            changed = True
                            log_func(f"🚨 ALERTA: UID {uid} ({info.get('nombre')}) sin salida registrada desde {ultima_actividad} ({diff_hours:.1f}h).")
                    except Exception:
                        pass
            if changed:
                guardar_estados()
        except Exception as e:
            # no queremos que el monitor mate el programa
            try:
                log_func(f"Monitor error: {e}")
            except:
                print("Monitor error:", e)
        # dormir
        stop_event.wait(ALERTA_CHECK_SECONDS)

# ==========================================
# 🔌 HILO PRINCIPAL
# ==========================================
def hilo_proceso(ip, sucursal, log_func, update_status_func, add_row_func, stop_event=None):
    guardar_config(ip, sucursal)
    log_func("--- SISTEMA INICIADO ---")

    # asegurar archivos base
    ensure_files_exist()

    # Cargar usuarios y estados persistidos
    usuarios_local = cargar_usuarios()
    cargar_estados()
    # sincronizar nombres y tipos en caso de cambios manuales
    sync_nombres_con_usuarios(usuarios_local)

    # 1. Cargar historial previo para anti-duplicados
    cargar_historial_existente(log_func)

    sheet = conectar_google(log_func)
    if sheet:
        update_status_func("google", True)
        log_func("☁️ Nube Conectada")
    else:
        update_status_func("google", False)
        log_func("⚠️ MODO OFFLINE")

    while True:
        try:
            conn = ZK(ip, port=4370, timeout=10, password=0, force_udp=True, ommit_ping=True)
            conn.connect()
            update_status_func("reloj", True)

            conn.disable_device()
            users = conn.get_users()
            mapa = {str(u.user_id): u.name for u in users}
            att = conn.get_attendance()
            conn.enable_device()

            if att:
                batch_local = []
                batch_nube = []
                nuevos_contador = 0

                # reload usuarios each loop so GUI edits are respected
                usuarios_local = cargar_usuarios()
                # sync names and types
                sync_nombres_con_usuarios(usuarios_local)

                for a in att:
                    fecha_str = a.timestamp.strftime("%Y-%m-%d %H:%M:%S")
                    uid = str(a.user_id)

                    # --- FILTRO MAESTRO ---
                    llave_unica = f"{uid}_{fecha_str}"

                    if llave_unica in HISTORIAL_PROCESADO:
                        continue  # ¡YA EXISTE! Lo saltamos

                    # Si llegamos aquí, es NUEVO
                    HISTORIAL_PROCESADO.add(llave_unica)
                    nuevos_contador += 1

                    raw_nom = mapa.get(uid, "Desconocido")
                    user_info = usuarios_local.get(uid)
                    # PRIORIDAD: nombre del sistema (usuarios_config) > nombre del dispositivo
                    display_name = raw_nom
                    tipo = "visitante"
                    if user_info:
                        nombre_conf = user_info.get("nombre")
                        if nombre_conf:
                            display_name = nombre_conf
                        tipo = user_info.get("tipo", "visitante")

                    modo, est = analizar_registro(uid, a.timestamp, a.punch, usuarios_local)

                    # Actualizar estado en RAM y persistir (usa display_name)
                    actualizar_estado_usuario(uid, display_name, tipo, modo, fecha_str)

                    # Agregar a GUI (Visual) - add_row_func mostrará los datos actualizados desde ESTADOS_USUARIOS
                    add_row_func(uid, display_name, fecha_str, modo, est)

                    # Agregar a listas de guardado (notar columnas extendidas)
                    batch_local.append([uid, display_name, fecha_str, modo, est, sucursal, tipo, modo, fecha_str])
                    batch_nube.append([uid, display_name, fecha_str, modo, est, sucursal, tipo, modo, fecha_str])

                if nuevos_contador > 0:
                    log_func(f"✅ Se detectaron {nuevos_contador} registros NUEVOS.")
                    guardar_excel_local(batch_local)
                    if sheet:
                        try:
                            sheet.append_rows(batch_nube)
                        except Exception as e:
                            log_func(f"⚠️ Error subiendo a nube: {e}")
                else:
                    pass

            conn.disconnect()
            # permitir salida ordenada si stop_event está activo (útil en pruebas)
            if stop_event and stop_event.is_set():
                break
            time.sleep(60)

        except Exception as e:
            update_status_func("reloj", False)
            log_func(f"Reintentando: {e}")
            time.sleep(20)

# ==========================================
# 🖥️ INTERFAZ GRÁFICA + GESTOR DE USUARIOS
# ==========================================
def start_gui():
    # Asegurar archivos antes de todo (para que el exe + credentials funcione en otra máquina)
    ensure_files_exist()

    usuarios_local = cargar_usuarios()
    cargar_estados()
    # sincronizar por si el archivo de usuarios trae cambios
    sync_nombres_con_usuarios(usuarios_local)

    root = tk.Tk()
    root.title("ZKTECO SYNC PRO v7.2")
    root.geometry("1000x700")

    style = ttk.Style()
    style.theme_use("clam")
    style.configure("Treeview", rowheight=25, font=('Arial', 9))
    style.configure("Treeview.Heading", font=('Arial', 10, 'bold'), background="#ddd")

    # Header
    frame_top = tk.Frame(root, bg="#2c3e50", height=60)
    frame_top.pack(fill="x")
    tk.Label(frame_top, text="ZKTECO MANAGER", fg="white", bg="#2c3e50", font=("Segoe UI", 16, "bold")).pack(side="left", padx=20, pady=10)

    # Config
    cfg = cargar_config()
    frame_cfg = tk.Frame(root, pady=10, bg="#ecf0f1")
    frame_cfg.pack(fill="x")

    tk.Label(frame_cfg, text="IP Reloj:", bg="#ecf0f1").pack(side="left", padx=(20,5))
    entry_ip = tk.Entry(frame_cfg, width=15)
    entry_ip.insert(0, cfg.get("ip"))
    entry_ip.pack(side="left")

    tk.Label(frame_cfg, text="Sucursal:", bg="#ecf0f1").pack(side="left", padx=(20,5))
    entry_suc = tk.Entry(frame_cfg, width=20)
    entry_suc.insert(0, cfg.get("sucursal"))
    entry_suc.pack(side="left")

    # Botón para gestionar usuarios
    def abrir_gestor_usuarios():
        usuarios = cargar_usuarios()
        win = tk.Toplevel(root)
        win.title("Gestor de Usuarios")
        win.geometry("700x450")

        cols = ("ID", "Nombre", "Tipo", "Entrada", "Salida")
        tree_u = ttk.Treeview(win, columns=cols, show="headings")
        for c in cols:
            tree_u.heading(c, text=c)
            tree_u.column(c, width=120, anchor="center")
        tree_u.pack(fill="both", expand=True, padx=10, pady=10)

        def refrescar_tree():
            for i in tree_u.get_children():
                tree_u.delete(i)
            u = cargar_usuarios()
            for k, v in u.items():
                tree_u.insert("", "end", values=(k, v.get("nombre",""), v.get("tipo","visitante"), v.get("hora_entrada",""), v.get("hora_salida","")))

        def validar_hhmm(h):
            if not h:
                return True
            try:
                parts = h.split(":")
                if len(parts) != 2:
                    return False
                hh = int(parts[0])
                mm = int(parts[1])
                return 0 <= hh < 24 and 0 <= mm < 60
            except:
                return False

        def add_user():
            uid = simpledialog.askstring("ID usuario", "ID único (ej. 23):", parent=win)
            if not uid:
                return
            nombre = simpledialog.askstring("Nombre", "Nombre completo:", parent=win)
            if not nombre:
                return
            tipo = simpledialog.askstring("Tipo", "Tipo (empleado/visitante/recluso/externo):", parent=win, initialvalue="empleado")
            hora_entrada = simpledialog.askstring("Hora entrada", "Formato HH:MM (dejar vacío para visitante):", parent=win)
            hora_salida = simpledialog.askstring("Hora salida", "Formato HH:MM:", parent=win)
            if hora_entrada and not validar_hhmm(hora_entrada):
                messagebox.showerror("Error", "Formato hora entrada inválido. Usa HH:MM", parent=win)
                return
            if hora_salida and not validar_hhmm(hora_salida):
                messagebox.showerror("Error", "Formato hora salida inválido. Usa HH:MM", parent=win)
                return
            usuarios = cargar_usuarios()
            usuarios[uid] = {"nombre": nombre, "tipo": tipo}
            if hora_entrada: usuarios[uid]["hora_entrada"] = hora_entrada
            if hora_salida: usuarios[uid]["hora_salida"] = hora_salida
            guardar_usuarios(usuarios)
            # sincronizar nombres y tipos en estados
            sync_nombres_con_usuarios(usuarios)
            refrescar_tree()

        def edit_user():
            sel = tree_u.selection()
            if not sel:
                messagebox.showinfo("Editar", "Selecciona un usuario", parent=win)
                return
            vals = tree_u.item(sel[0])["values"]
            uid = str(vals[0])
            usuarios = cargar_usuarios()
            u = usuarios.get(uid, {})
            nombre = simpledialog.askstring("Nombre", "Nombre completo:", parent=win, initialvalue=u.get("nombre",""))
            tipo = simpledialog.askstring("Tipo", "Tipo (empleado/visitante/recluso/externo):", parent=win, initialvalue=u.get("tipo","visitante"))
            hora_entrada = simpledialog.askstring("Hora entrada", "Formato HH:MM (dejar vacío para visitante):", parent=win, initialvalue=u.get("hora_entrada",""))
            hora_salida = simpledialog.askstring("Hora salida", "Formato HH:MM:", parent=win, initialvalue=u.get("hora_salida",""))
            if hora_entrada and not validar_hhmm(hora_entrada):
                messagebox.showerror("Error", "Formato hora entrada inválido. Usa HH:MM", parent=win)
                return
            if hora_salida and not validar_hhmm(hora_salida):
                messagebox.showerror("Error", "Formato hora salida inválido. Usa HH:MM", parent=win)
                return
            usuarios[uid] = {"nombre": nombre, "tipo": tipo}
            if hora_entrada: usuarios[uid]["hora_entrada"] = hora_entrada
            elif "hora_entrada" in usuarios[uid]:
                usuarios[uid].pop("hora_entrada", None)
            if hora_salida: usuarios[uid]["hora_salida"] = hora_salida
            elif "hora_salida" in usuarios[uid]:
                usuarios[uid].pop("hora_salida", None)
            guardar_usuarios(usuarios)
            # sincronizar nombres y tipos en estados (inmediato)
            sync_nombres_con_usuarios(usuarios)
            refrescar_tree()

        def del_user():
            sel = tree_u.selection()
            if not sel:
                messagebox.showinfo("Eliminar", "Selecciona un usuario", parent=win)
                return
            vals = tree_u.item(sel[0])["values"]
            uid = str(vals[0])
            usuarios = cargar_usuarios()
            if uid in usuarios:
                if messagebox.askyesno("Confirmar", f"Eliminar usuario {uid} - {usuarios[uid].get('nombre')} ?", parent=win):
                    usuarios.pop(uid)
                    guardar_usuarios(usuarios)
                    refrescar_tree()

        btn_frame = tk.Frame(win)
        btn_frame.pack(fill="x", padx=10, pady=(0,10))
        tk.Button(btn_frame, text="Agregar", command=add_user).pack(side="left", padx=5)
        tk.Button(btn_frame, text="Editar", command=edit_user).pack(side="left", padx=5)
        tk.Button(btn_frame, text="Eliminar", command=del_user).pack(side="left", padx=5)
        tk.Button(btn_frame, text="Actualizar", command=refrescar_tree).pack(side="left", padx=5)

        refrescar_tree()

    tk.Button(frame_cfg, text="Gestionar Usuarios", command=abrir_gestor_usuarios, bg="#f39c12", fg="white").pack(side="right", padx=10)

    # Status
    frame_status = tk.Frame(frame_cfg, bg="#ecf0f1")
    frame_status.pack(side="right", padx=20)
    lbl_reloj = tk.Label(frame_status, text="● RELOJ", fg="#bdc3c7", bg="#ecf0f1", font=("Arial", 10, "bold"))
    lbl_reloj.pack(side="left", padx=10)
    lbl_cloud = tk.Label(frame_status, text="● NUBE", fg="#bdc3c7", bg="#ecf0f1", font=("Arial", 10, "bold"))
    lbl_cloud.pack(side="left", padx=10)

    def update_status(tipo, online):
        color = "#27ae60" if online else "#c0392b"
        if tipo == "reloj": lbl_reloj.config(fg=color)
        if tipo == "google": lbl_cloud.config(fg=color)

    # Tabla principal
    frame_table = tk.Frame(root)
    frame_table.pack(fill="both", expand=True, padx=10, pady=5)

    cols = ("ID", "Nombre", "Hora", "Evento", "Análisis", "Tipo", "Ult_Estado", "Ult_Actividad")
    tree = ttk.Treeview(frame_table, columns=cols, show="headings")

    for c in cols:
        tree.heading(c, text=c)
        if c == "Nombre":
            tree.column(c, width=220)
        else:
            tree.column(c, width=120, anchor="center")

    tree.pack(side="left", fill="both", expand=True)

    scrollbar = ttk.Scrollbar(frame_table, orient="vertical", command=tree.yview)
    tree.configure(yscroll=scrollbar.set)
    scrollbar.pack(side="right", fill="y")
    tree.tag_configure("late", foreground="red")
    tree.tag_configure("ok", foreground="green")

    def add_row_to_table(uid, nom, hora, evento, estado):
        # Mostrar info extendida si existe estado en RAM
        s = ESTADOS_USUARIOS.get(str(uid), {})
        tipo = s.get("tipo", "")
        ultimo_estado = s.get("ultimo_estado", "")
        ultima_act = s.get("ultima_actividad", "")
        # LIMITADOR DE FILAS
        if len(tree.get_children()) > 200:
            tree.delete(tree.get_children()[-1])
        tag = "late" if "Retardo" in estado or "Anticipada" in estado else "ok"
        tree.insert("", 0, values=(uid, nom, hora, evento, estado, tipo, ultimo_estado, ultima_act), tags=(tag,))

    # Log
    frame_log = tk.Frame(root, height=100)
    frame_log.pack(fill="x", padx=10, pady=5)
    log_widget = scrolledtext.ScrolledText(frame_log, height=6, font=("Consolas", 8))
    log_widget.pack(fill="both")

    def log(msg):
        t = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_widget.insert(tk.END, f"[{t}] {msg}\n")
        log_widget.see(tk.END)

    # Run
    def run():
        btn_start.config(state="disabled", text="SERVICIO ACTIVO", bg="#27ae60")
        stop_event = threading.Event()
        # hilo principal
        t = threading.Thread(target=hilo_proceso, args=(
            entry_ip.get(), entry_suc.get(), log, update_status, add_row_to_table, stop_event
        ))
        t.daemon = True
        t.start()
        # monitor visitantes
        t2 = threading.Thread(target=monitor_visitantes, args=(log, stop_event))
        t2.daemon = True
        t2.start()

    btn_start = tk.Button(root, text="INICIAR SISTEMA", command=run, bg="#2980b9", fg="white", font=("Arial", 11, "bold"), height=2)
    btn_start.pack(fill="x", padx=20, pady=10)

    # Mostrar panel lateral con resumen rápido de estados
    def abrir_panel_estados():
        win = tk.Toplevel(root)
        win.title("Resumen de Estados")
        win.geometry("500x400")
        cols = ("ID", "Nombre", "Tipo", "Ult_Estado", "Ult_Actividad", "Alerta")
        tree_s = ttk.Treeview(win, columns=cols, show="headings")
        for c in cols:
            tree_s.heading(c, text=c)
            tree_s.column(c, width=100, anchor="center")
        tree_s.pack(fill="both", expand=True, padx=10, pady=10)

        def refrescar():
            for i in tree_s.get_children():
                tree_s.delete(i)
            for uid, info in ESTADOS_USUARIOS.items():
                tree_s.insert("", "end", values=(uid, info.get("nombre",""), info.get("tipo",""), info.get("ultimo_estado",""), info.get("ultima_actividad",""), info.get("alerta", False)))
        tk.Button(win, text="Actualizar", command=refrescar).pack(pady=5)

        # botón extra: marcar salida manual para un UID seleccionado
        def marcar_salida_manual():
            sel = tree_s.selection()
            if not sel:
                messagebox.showinfo("Marcar salida", "Selecciona un registro", parent=win)
                return
            vals = tree_s.item(sel[0])["values"]
            uid = str(vals[0])
            now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            # si no existe en usuarios, tratar como visitante por defecto
            usuarios = cargar_usuarios()
            user_info = usuarios.get(uid, {})
            tipo = user_info.get("tipo", ESTADOS_USUARIOS.get(uid, {}).get("tipo", "visitante"))
            nombre = user_info.get("nombre", ESTADOS_USUARIOS.get(uid, {}).get("nombre", uid))
            # marcar salida
            actualizar_estado_usuario(uid, nombre, tipo, "Salida", now)
            messagebox.showinfo("Salida marcada", f"Salida manual marcada para {uid} - {nombre} a las {now}", parent=win)
            refrescar()

        tk.Button(win, text="Marcar salida manual", command=marcar_salida_manual).pack(pady=(0,5))

        refrescar()

    tk.Button(frame_cfg, text="Panel Estados", command=abrir_panel_estados, bg="#16a085", fg="white").pack(side="right", padx=5)

    root.mainloop()

if __name__ == "__main__":
    start_gui()
