import streamlit as st
import pandas as pd
import psycopg2
# import hashlib  <-- YA NO LO NECESITAS (Supabase se encarga de la encriptación)
from datetime import date, timedelta
import time
import io
from supabase import create_client, Client  # <--- NUEVO
from streamlit_option_menu import option_menu # <--- NUEVO

# ==========================================
# REGIÓN 1: CONFIGURACIÓN Y ESTILOS
# ==========================================

# Configuración de la página (ESTO SE QUEDA IGUAL, SIEMPRE PRIMERO)
try:
    st.set_page_config(page_title="Kilaco Intranet", layout="wide", page_icon="logo.png")
except:
    st.set_page_config(page_title="Kilaco Intranet", layout="wide", page_icon="🥖")

# --- CONEXIÓN SUPABASE (NUEVO BLOQUE) ---
# Inicializamos el cliente aquí para que esté disponible en toda la app
try:
    # Usamos st.secrets para leer lo que pusiste en secrets.toml
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    supabase: Client = create_client(url, key)
except Exception as e:
    st.error(f"Error crítico conectando a Supabase: {e}")
    st.stop() # Detiene la app si no hay conexión segura

# --- FIN DE LA ZONA DE CONFIGURACIÓN ---


def aplicar_estilos_kilaco():
    """Inyecta CSS para tipografía, pestañas y optimización ergonómica táctil (MILL)."""
    st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600&display=swap');
        .stApp {font-family: 'Inter', sans-serif;}
        
        /* 1. OPTIMIZACIÓN DE ESPACIO BASE (Escritorio) */
        .block-container {
            padding-top: 2rem !important;
            padding-bottom: 2rem !important;
        }
        
        /* 2. JERARQUÍA EN PESTAÑAS */
        button[data-baseweb="tab"] > div[data-testid="stMarkdownContainer"] > p {
            font-size: 1.05rem !important;
            font-weight: 600 !important;
            color: #333333 !important;
        }

        /* 3. ERGONOMÍA TÁCTIL (Media Queries para Móviles y Tablets) */
        @media (max-width: 768px) {
            /* Reducir márgenes laterales muertos para ganar espacio útil en la grilla */
            .block-container {
                padding-left: 0.5rem !important;
                padding-right: 0.5rem !important;
                padding-top: 1rem !important;
            }
            
            /* Áreas de impacto (Tap Targets) ampliadas para los dedos */
            .stButton > button {
                min-height: 3.2rem !important;
                font-size: 1.1rem !important;
                border-radius: 8px !important;
            }
            
            /* Bloqueo del Zoom de iOS: Exige que los inputs sean de 16px mínimo */
            input, select, textarea, div[data-baseweb="select"] > div {
                font-size: 16px !important;
                min-height: 3rem !important;
            }
            
            /* Ajuste de jerarquía de textos para no saturar pantallas pequeñas */
            h3 { font-size: 1.5rem !important; }
            h4 { font-size: 1.3rem !important; }
            h5 { font-size: 1.1rem !important; }
        }
        </style>
    """, unsafe_allow_html=True)

# ==========================================
# REGIÓN 2: CONSTANTES Y UTILIDADES
# ==========================================

DB_URI = st.secrets["DB_URI"]

from psycopg2 import pool

@st.cache_resource
def get_connection_pool():
    """Vena Principal: Mantiene un túnel abierto hacia Supabase."""
    return pool.ThreadedConnectionPool(1, 20, DB_URI)

class ConexionFantasma:
    """
    Wrapper (Envoltorio) ontológico. Actúa como un espejo de la conexión real en C, 
    pero intercepta la orden de destrucción para preservar la vida del túnel.
    """
    def __init__(self, piscina, conexion_real):
        self.piscina = piscina
        self.conexion_real = conexion_real

    def cursor(self, *args, **kwargs):
        return self.conexion_real.cursor(*args, **kwargs)

    def commit(self):
        self.conexion_real.commit()

    def rollback(self):
        self.conexion_real.rollback()

    def close(self):
        # MAGIA: En lugar de destruir la conexión, la liberamos de vuelta a la piscina
        try:
            self.piscina.putconn(self.conexion_real)
        except:
            pass

    def __getattr__(self, nombre):
        # Delega cualquier otro comando directamente al motor de C (ej: consultas de Pandas)
        return getattr(self.conexion_real, nombre)

def get_conn():
    """Obtiene una conexión instantánea y la entrega disfrazada."""
    try:
        p = get_connection_pool()
        conn = p.getconn()
        return ConexionFantasma(p, conn)
    except Exception as e:
        st.error(f"Error crítico en el túnel de conexión: {e}")
        st.stop()

def safe_float(val):
    if val is None: return 0.0
    try: return float(val)
    except: return 0.0

def safe_int(val):
    if val is None: return 0
    try: return int(float(val))
    except: return 0

def fmt_clp(valor):
    return "$ 0" if valor is None else "$ " + "{:,.0f}".format(valor).replace(",", ".")

def val_gui(val):
    return val if val and val > 0 else None

def val_db(val):
    return int(val) if val else 0

def descargar_respaldo_completo():
    """Genera un archivo Excel en memoria con todas las tablas clave."""
    conn = get_conn()
    buffer = io.BytesIO()
    
    tablas = [
        "usuarios", "vendedores", "productos", "productos_bolsas", 
        "comunas", "bancos", "clientes", "clientes_corriente",
        "stock", "despacho", "despacho_corriente", 
        "finanzas", "finanzas_corriente", 
        "control_bandejas", "control_bolsas", 
        "movimientos_credito", "transferencias", 
        "caja_movimientos", "produccion_extras", "produccion_corriente"
    ]
    
    try:
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            for tabla in tablas:
                try:
                    df = pd.read_sql(f"SELECT * FROM {tabla}", conn)
                    if not df.empty:
                        df.to_excel(writer, sheet_name=tabla[:31], index=False)
                except Exception as e:
                    pass
    finally:
        conn.close()
        
    return buffer

# ==========================================
# REGIÓN 3: ACCESO A DATOS (DAO) Y LÓGICA
# ==========================================

# --- Autenticación y Sesión ---
def check_login(email, password):
    """Autentica con Supabase Auth y recupera el perfil y rol."""
    try:
        # 1. Login en Supabase (Auth)
        session = supabase.auth.sign_in_with_password({"email": email, "password": password})
        user_id = session.user.id
        
        # 2. Consultar el ROL en la tabla 'profiles'
        response = supabase.table("profiles").select("*").eq("id", user_id).execute()
        
        if response.data:
            perfil = response.data[0]
            return {
                "id": perfil["id"],
                "nombre": perfil.get("nombre_completo", email),
                "rol": perfil.get("rol", "repartidor"),
                "id_vendedor": perfil.get("id_vendedor")
            }
        return None
    except Exception as e:
        print(f"Login error: {e}")
        return None

def init_session():
    if "logged_in" not in st.session_state: st.session_state.logged_in = False
    if "user_name" not in st.session_state: st.session_state.user_name = None
    if "user_role" not in st.session_state: st.session_state.user_role = None
    if "user_id" not in st.session_state: st.session_state.user_id = None
    if "id_vendedor" not in st.session_state: st.session_state.id_vendedor = None 
    if "current_module" not in st.session_state: st.session_state.current_module = "menu"

# --- Referencias Generales ---
@st.cache_data(ttl=600)
def get_referencias():
    conn = get_conn()
    try:
        with conn.cursor() as c:
            c.execute("SELECT nombre FROM comunas ORDER BY nombre")
            # Truco MILL: Inyectamos el "-" en la primera posición (Índice 0)
            comunas = ["-"] + [r[0] for r in c.fetchall()]
            
            c.execute("SELECT nombre FROM bancos ORDER BY nombre")
            bancos = [r[0] for r in c.fetchall()]
        
        # Ahora extraemos TODO el perfil del vendedor desde la DB
        query_vend = """
            SELECT id, nombre, area, bencina, comision, dscto_esp 
            FROM vendedores 
            WHERE nombre NOT IN ('Vendedor 1', 'Vendedor 2', 'nan', 'NaN') 
            AND nombre IS NOT NULL 
            ORDER BY nombre
        """
        vendedores = pd.read_sql(query_vend, conn)
        return comunas, bancos, vendedores
    finally:
        conn.close()

# --- ESCUDOS TÉRMICOS (CACHÉ) PARA LA INTERFAZ ---

@st.cache_data(ttl=300)
def obtener_productos_activos():
    """Caché de productos para evitar consultas al cambiar de pestañas en Despacho."""
    conn = get_conn()
    try: return pd.read_sql("SELECT * FROM productos ORDER BY orden_visual", conn)
    finally: conn.close()

@st.cache_data(ttl=60)
def obtener_despacho_vehiculo(fecha, id_vendedor):
    """Caché de la carga actual del vehículo."""
    conn = get_conn()
    try: return pd.read_sql("SELECT d.id, p.nombre, d.saldo_anterior, d.carga FROM despacho d JOIN productos p ON d.id_producto=p.id WHERE d.fecha=%s AND d.id_vendedor=%s AND d.carga > 0 ORDER BY p.orden_visual ASC", conn, params=(fecha, id_vendedor))
    finally: conn.close()

@st.cache_data(ttl=60)
def obtener_transferencias_recientes():
    """Caché para la pestaña de Conciliación."""
    conn = get_conn()
    try: return pd.read_sql("""SELECT t.id, TO_CHAR(t.fecha, 'DD/MM/YYYY') as fecha, v.nombre as "Repartidor", t.monto, t.banco_emisor, t.tipo_transferencia, t.verificado FROM transferencias t JOIN vendedores v ON t.id_vendedor=v.id ORDER BY t.fecha DESC, t.id DESC LIMIT 50""", conn)
    finally: conn.close()

@st.cache_data(ttl=60)
def obtener_caja_del_dia(fecha, area):
    """Caché de la auditoría diaria de caja (Aplica para Especial y Corriente)."""
    conn = get_conn()
    try: return pd.read_sql("SELECT id, area, descripcion as entidad, item as detalle, ingreso_efectivo, ingreso_transferencia, egreso as egreso_efectivo, egreso_transferencia, rol_creador FROM caja_movimientos WHERE fecha=%s AND area=%s ORDER BY id DESC", conn, params=(fecha, area))
    finally: conn.close()

@st.cache_data(ttl=60)
def obtener_todas_entidades():
    """Caché para el mantenedor de Ajustes de Caja."""
    conn = get_conn()
    try: return pd.read_sql("SELECT * FROM entidades WHERE estado='ACTIVO' ORDER BY nombre", conn)
    finally: conn.close()

# --- Consultas a la base de datos ---
def obtener_diccionario_vendedores():
    """Retorna un diccionario { 'Nombre': ID } de todos los vendedores."""
    conn = get_conn()
    try:
        df = pd.read_sql("SELECT id, nombre FROM vendedores ORDER BY nombre", conn)
        return dict(zip(df['nombre'], df['id']))
    finally:
        conn.close()

def obtener_diccionario_clientes_corriente():
    """Retorna un diccionario { 'Nombre': ID } de los clientes del corriente."""
    conn = get_conn()
    try:
        df = pd.read_sql("SELECT id, nombre FROM clientes_corriente ORDER BY nombre", conn)
        return dict(zip(df['nombre'], df['id']))
    finally:
        conn.close()

# --- Funciones Pan Especial ---
def obtener_datos_stock(fecha):
    conn = get_conn()
    try:
        query = """
            SELECT p.id, p.nombre, p.produccion_standar, p.bolsas_por_caja, p.rendimiento_por_saco,
            COALESCE((SELECT stock_final FROM stock WHERE id_producto=p.id AND fecha<%s ORDER BY fecha DESC LIMIT 1), 0) as stock_inicial,
            COALESCE(s.fabricacion, 0) as fabricacion,
            COALESCE((SELECT SUM(d.carga) FROM despacho d WHERE d.fecha=%s AND d.id_producto=p.id), 0) as salida_calculada
            FROM productos p
            LEFT JOIN stock s ON p.id=s.id_producto AND s.fecha=%s
            ORDER BY p.orden_visual ASC
        """
        df = pd.read_sql(query, conn, params=(fecha, fecha, fecha))
        # Filtro de seguridad por si quedaron productos basura
        df = df[~df['nombre'].isin(['Molde Integral', 'Molde Integral XL'])]
        
        df.fillna(0, inplace=True)
        df['stock_final'] = df['stock_inicial'] + df['fabricacion'] - df['salida_calculada']
        df['bolsas_necesarias'] = (df['produccion_standar'] - df['stock_final']).clip(lower=0) 
        df['rendimiento_por_saco'] = df['rendimiento_por_saco'].replace(0, 1)
        df['sacos_manana'] = df['bolsas_necesarias'] / df['rendimiento_por_saco']
        return df
    finally:
        conn.close()

def registrar_produccion(fecha, id_prod, stock_ini, fab, stock_fin, bolsas_nec):
    conn = get_conn()
    try:
        with conn.cursor() as c:
            c.execute("SELECT id FROM stock WHERE fecha=%s AND id_producto=%s", (fecha, id_prod))
            ex = c.fetchone()
            if ex:
                c.execute("UPDATE stock SET stock_inicial=%s, fabricacion=%s, stock_final=%s, bolsas_necesarias=%s WHERE id=%s", (stock_ini, fab, stock_fin, bolsas_nec, ex[0]))
            else:
                c.execute("INSERT INTO stock (fecha, id_producto, stock_inicial, fabricacion, stock_final, bolsas_necesarias) VALUES (%s,%s,%s,%s,%s,%s)", (fecha, id_prod, stock_ini, fab, stock_fin, bolsas_nec))
        conn.commit()
    finally:
        conn.close()

def obtener_bandejas(fecha, id_v):
    conn = get_conn()
    try:
        with conn.cursor() as c:
            c.execute("SELECT * FROM control_bandejas WHERE fecha=%s AND id_vendedor=%s", (fecha, id_v))
            hoy = c.fetchone()
            if hoy: return {"ant": hoy[3], "sal": hoy[4], "ret": hoy[5], "fin": hoy[6], "existe": True}
            
            c.execute("SELECT saldo_final FROM control_bandejas WHERE id_vendedor=%s AND fecha < %s ORDER BY fecha DESC LIMIT 1", (id_v, fecha))
            hist = c.fetchone()
            ant = hist[0] if hist else 0
        return {"ant": ant, "sal": 0, "ret": 0, "fin": ant, "existe": False}
    finally:
        conn.close()

def guardar_bandejas(fecha, id_v, ant, sal, ret):
    fin = ant - sal + ret
    conn = get_conn()
    try:
        with conn.cursor() as c:
            c.execute("SELECT id FROM control_bandejas WHERE fecha=%s AND id_vendedor=%s", (fecha, id_v))
            ex = c.fetchone()
            if ex:
                c.execute("UPDATE control_bandejas SET saldo_anterior=%s, salida=%s, retorno=%s, saldo_final=%s WHERE id=%s", (ant, sal, ret, fin, ex[0]))
            else:
                c.execute("INSERT INTO control_bandejas (fecha, id_vendedor, saldo_anterior, salida, retorno, saldo_final) VALUES (%s,%s,%s,%s,%s,%s)", (fecha, id_v, ant, sal, ret, fin))
        conn.commit()
    finally:
        conn.close()

def obtener_bolsas_manual(fecha):
    conn = get_conn()
    try:
        query = """
            SELECT pb.id, pb.nombre, pb.unidades_por_bolsa as factor,
            COALESCE((SELECT stock_cajas_final FROM control_bolsas WHERE id_producto=pb.id AND fecha < %s ORDER BY fecha DESC LIMIT 1), 0.0) as stock_inicial_cajas,
            cb.id as id_cb, 
            COALESCE(cb.ingreso_cajas, 0.0) as ingreso_cajas, 
            COALESCE(cb.produccion_ayer, 0) as produccion_hoy_unidades
            FROM productos_bolsas pb
            LEFT JOIN control_bolsas cb ON pb.id = cb.id_producto AND cb.fecha = %s
            ORDER BY pb.id ASC
        """
        df = pd.read_sql(query, conn, params=(fecha, fecha))
        df['stock_inicial_cajas'] = df['stock_inicial_cajas'].astype(float)
        df['ingreso_cajas'] = df['ingreso_cajas'].astype(float)
        df['factor'] = df['factor'].replace(0, 1)
        df['stock_inicial_bolsas'] = df['stock_inicial_cajas'] * df['factor']
        df['gasto_cajas'] = df['produccion_hoy_unidades'] / df['factor']
        df['stock_cajas_final'] = df['stock_inicial_cajas'] + df['ingreso_cajas'] - df['gasto_cajas']
        df['stock_bolsas_final'] = df['stock_cajas_final'] * df['factor']
        return df
    finally:
        conn.close()

def guardar_bolsas_manual(fecha, df):
    conn = get_conn()
    try:
        with conn.cursor() as c:
            for i, r in df.iterrows():
                if pd.notna(r['id_cb']):
                    c.execute("UPDATE control_bolsas SET stock_cajas_ayer=%s, ingreso_cajas=%s, produccion_ayer=%s, stock_cajas_final=%s WHERE id=%s", (r['stock_inicial_cajas'], r['ingreso_cajas'], r['produccion_hoy_unidades'], r['stock_cajas_final'], r['id_cb']))
                else:
                    c.execute("INSERT INTO control_bolsas (fecha, id_producto, stock_cajas_ayer, ingreso_cajas, produccion_ayer, stock_cajas_final) VALUES (%s,%s,%s,%s,%s,%s)", (fecha, r['id'], r['stock_inicial_cajas'], r['ingreso_cajas'], r['produccion_hoy_unidades'], r['stock_cajas_final']))
        conn.commit()
    finally:
        conn.close()

def registrar_carga(fecha, id_v, id_p, cant):
    conn = get_conn()
    try:
        with conn.cursor() as c:
            c.execute("SELECT saldo_actual FROM despacho WHERE id_vendedor=%s AND id_producto=%s AND fecha < %s ORDER BY fecha DESC LIMIT 1", (id_v, id_p, fecha))
            row = c.fetchone()
            saldo_ant = safe_int(row[0] if row else 0)
            
            c.execute("SELECT id FROM despacho WHERE fecha=%s AND id_vendedor=%s AND id_producto=%s", (fecha, id_v, id_p))
            ex = c.fetchone()
            
            saldo_proj = saldo_ant + cant
            if ex:
                c.execute("UPDATE despacho SET carga = carga + %s, saldo_actual = saldo_actual + %s WHERE id=%s", (cant, cant, ex[0]))
            else:
                c.execute("INSERT INTO despacho (fecha, id_vendedor, id_producto, saldo_anterior, carga, devolucion_muestra, saldo_actual, venta_unidades) VALUES (%s,%s,%s,%s,%s,0,%s,0)", (fecha, id_v, id_p, saldo_ant, cant, saldo_proj))
        conn.commit()
        st.cache_data.clear() # Limpiamos la caché para que la tabla se actualice al instante
        st.toast("Carga registrada")
    finally:
        conn.close()

def actualizar_carga_masiva(df_editado):
    conn = get_conn()
    try:
        with conn.cursor() as c:
            for i, r in df_editado.iterrows():
                nuevo_saldo = r['saldo_anterior'] + r['carga']
                c.execute("UPDATE despacho SET carga=%s, saldo_actual=%s WHERE id=%s", (r['carga'], nuevo_saldo, r['id']))
        conn.commit()
        st.cache_data.clear() # Limpiamos la caché para que la tabla se actualice al instante
        st.toast("Carga corregida")
    finally:
        conn.close()

def obtener_planilla(fecha, id_v):
    conn = get_conn()
    try:
        df_prod = pd.read_sql("SELECT id, nombre, precio_estandar, orden_visual FROM productos", conn)
        df_hoy = pd.read_sql("SELECT id as id_despacho, id_producto, saldo_anterior, carga, devolucion_muestra, saldo_actual FROM despacho WHERE fecha=%s AND id_vendedor=%s", conn, params=(fecha, id_v))
        
        with conn.cursor() as c:
            c.execute("SELECT dscto_esp, nombre FROM vendedores WHERE id=%s", (id_v,))
            res_desc = c.fetchone()
            descuento = float(res_desc[0]) if res_desc and res_desc[0] else 0.0
            vendedor_nombre = res_desc[1] if res_desc else ""
            
        # --- MOTOR DE PRECIOS PERSONALIZADOS (HUGO PALACIOS) ---
        if vendedor_nombre == "Hugo Palacios":
            precios_hugo = {
                "Lengua": 1150, "Lengua 6": 1000, "Frica": 1250, "Lengua XL (25)": 1350,
                "Molde XL": 1400, "Molde": 1300, "Pizza Individual": 1050, "Frica XL": 1350,
                "Hallulla": 1050, "Tapadito": 1400, "Pan Rallado": 550, "Pizza Familiar": 1300,
                "Lengua XXL (30)": 1950, "Lengua XXXL (35)": 2350, "Lengua XXXXL (40)": 2650
            }
            def set_precio_hugo(row):
                nombre_limpio = str(row['nombre']).strip()
                return precios_hugo.get(nombre_limpio, row['precio_estandar'])

            df_prod['precio_estandar'] = df_prod.apply(set_precio_hugo, axis=1)
            
        # Para el resto de los mortales, aplicamos el descuento normal
        elif descuento > 0:
            df_prod['precio_estandar'] = df_prod['precio_estandar'] * (1 - descuento)
        
        filas = []
        with conn.cursor() as c:
            for i, prod in df_prod.iterrows():
                id_p = prod['id']
                reg_hoy = df_hoy[df_hoy['id_producto'] == id_p]
                if not reg_hoy.empty:
                    r = reg_hoy.iloc[0]
                    filas.append({"id": r['id_despacho'], "id_producto": id_p, "nombre": prod['nombre'], "precio_estandar": prod['precio_estandar'], "orden_visual": prod['orden_visual'], "saldo_anterior": r['saldo_anterior'], "carga": r['carga'], "devolucion_muestra": r['devolucion_muestra'], "saldo_actual": r['saldo_actual']})
                else:
                    c.execute("SELECT saldo_actual FROM despacho WHERE id_vendedor=%s AND id_producto=%s AND fecha < %s ORDER BY fecha DESC LIMIT 1", (id_v, id_p, fecha))
                    last = c.fetchone()
                    saldo_ayer = last[0] if last else 0
                    if saldo_ayer > 0:
                        filas.append({"id": None, "id_producto": id_p, "nombre": prod['nombre'], "precio_estandar": prod['precio_estandar'], "orden_visual": prod['orden_visual'], "saldo_anterior": saldo_ayer, "carga": 0, "devolucion_muestra": 0, "saldo_actual": saldo_ayer})
        
        if not filas: return pd.DataFrame()
        df = pd.DataFrame(filas)
        df = df.sort_values('orden_visual') 
        df['disp'] = df['saldo_anterior'] + df['carga']
        df['venta'] = (df['disp'] - df['devolucion_muestra'] - df['saldo_actual']).clip(lower=0)
        df['total'] = df['venta'] * df['precio_estandar']
        return df
    finally:
        conn.close()

def guardar_oficina(df, fecha, id_v):
    conn = get_conn()
    try:
        with conn.cursor() as c:
            for i, r in df.iterrows():
                vta = max(0, r['saldo_anterior'] + r['carga'] - r['devolucion'] - r['saldo_final']) 
                if pd.notna(r['id']) and r['id'] != 0:
                    c.execute("UPDATE despacho SET devolucion_muestra=%s, saldo_actual=%s, venta_unidades=%s WHERE id=%s", (r['devolucion'], r['saldo_final'], vta, r['id']))
                else:
                    c.execute("INSERT INTO despacho (fecha, id_vendedor, id_producto, saldo_anterior, carga, devolucion_muestra, saldo_actual, venta_unidades) VALUES (%s, %s, %s, %s, 0, %s, %s, %s)", (fecha, id_v, r['id_producto_hidden'], r['saldo_anterior'], r['devolucion'], r['saldo_final'], vta))
        conn.commit()
        st.toast("Inventario guardado")
    finally:
        conn.close()

def get_finanzas(fecha, id_v):
    conn = get_conn()
    try:
        with conn.cursor() as c:
            c.execute("SELECT * FROM finanzas WHERE fecha=%s AND id_vendedor=%s", (fecha, id_v))
            row = c.fetchone()
        
        if row:
            return {"cc":safe_int(row[3]), "co":safe_int(row[4]), "ds":safe_int(row[5]), "bn":safe_int(row[6]), "su":safe_int(row[7]), "om":safe_int(row[8]), "od":row[9] or "Varios", "ef":safe_int(row[10]), "tr":safe_int(row[11]), "pc":safe_int(row[12]), "cc_det":row[13] or "Varios"}
        return {"cc":0, "co":0, "ds":0, "bn":0, "su":0, "om":0, "od":"Varios", "ef":0, "tr":0, "pc":0, "cc_det":"Varios"}
    finally:
        conn.close()

def save_finanzas(fecha, id_v, d):
    conn = get_conn()
    try:
        with conn.cursor() as c:
            c.execute("SELECT id FROM finanzas WHERE fecha=%s AND id_vendedor=%s", (fecha, id_v))
            ex = c.fetchone()
            
            # EL PURIFICADOR: Forzamos la conversión de numpy.int64 a int nativo de Python
            cc = int(d['cc'])
            co = int(d['co'])
            ds = int(d['ds'])
            bn = int(d['bn'])
            su = int(d['su'])
            om = int(d['om'])
            ef = int(d['ef'])
            tr = int(d['tr'])
            pc = int(d['pc'])
            
            if ex:
                c.execute("UPDATE finanzas SET creditos_cobrados=%s, creditos_otorgados=%s, descuentos_total=%s, bencina=%s, sueldo=%s, otros_gastos_monto=%s, otros_gastos_detalle=%s, efectivo_rendido=%s, transferencia_rendida=%s, pago_centralizado=%s, creditos_cobrados_detalle=%s WHERE id=%s", 
                          (cc, co, ds, bn, su, om, d['od'], ef, tr, pc, d['cc_det'], ex[0]))
            else:
                c.execute("INSERT INTO finanzas (fecha, id_vendedor, creditos_cobrados, creditos_otorgados, descuentos_total, bencina, sueldo, otros_gastos_monto, otros_gastos_detalle, efectivo_rendido, transferencia_rendida, pago_centralizado, creditos_cobrados_detalle) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)", 
                          (fecha, id_v, cc, co, ds, bn, su, om, d['od'], ef, tr, pc, d['cc_det']))
        conn.commit()
    finally:
        conn.close()

def obtener_resumen_global(fecha_inicio, fecha_fin):
    conn = get_conn()
    try:
        # 1. Extraemos la base cruda sin cálculos económicos ciegos
        df_despacho = pd.read_sql("""
            SELECT d.fecha, d.id_vendedor, v.nombre as vendedor_nombre, v.dscto_esp, 
                   d.id_producto, p.nombre as producto_nombre, p.precio_estandar, d.venta_unidades 
            FROM despacho d 
            JOIN productos p ON d.id_producto = p.id
            JOIN vendedores v ON d.id_vendedor = v.id
            WHERE d.fecha BETWEEN %s AND %s
        """, conn, params=(fecha_inicio, fecha_fin))
        
        # 2. Replicamos la alquimia de precios para que el reporte cuadre con el inventario
        if not df_despacho.empty:
            precios_hugo = {
                "Lengua": 1150, "Lengua 6": 1000, "Frica": 1250, "Lengua XL (25)": 1350,
                "Molde XL": 1400, "Molde": 1300, "Pizza Individual": 1050, "Frica XL": 1350,
                "Hallulla": 1050, "Tapadito": 1400, "Pan Rallado": 550, "Pizza Familiar": 1300,
                "Lengua XXL (30)": 1950, "Lengua XXXL (35)": 2350, "Lengua XXXXL (40)": 2650
            }
            
            def calcular_monto_real(row):
                vend = str(row['vendedor_nombre']).strip()
                if vend == "Hugo Palacios":
                    prod = str(row['producto_nombre']).strip()
                    precio = precios_hugo.get(prod, row['precio_estandar'])
                else:
                    desc = float(row['dscto_esp']) if pd.notna(row['dscto_esp']) else 0.0
                    precio = row['precio_estandar'] * (1 - desc)
                return row['venta_unidades'] * precio

            df_despacho['venta_pan'] = df_despacho.apply(calcular_monto_real, axis=1)
            df_ventas = df_despacho.groupby(['fecha', 'id_vendedor'])['venta_pan'].sum().reset_index()
        else:
            df_ventas = pd.DataFrame(columns=['fecha', 'id_vendedor', 'venta_pan'])

        # 3. Cruzamos con finanzas (Ya con la doble contabilización corregida)
        df_fin = pd.read_sql("SELECT * FROM finanzas WHERE fecha BETWEEN %s AND %s", conn, params=(fecha_inicio, fecha_fin))
        
        df_full = pd.merge(df_fin, df_ventas, on=['fecha', 'id_vendedor'], how='outer')
        
        for c in ['venta_pan', 'creditos_cobrados', 'bencina', 'sueldo', 'otros_gastos_monto', 'creditos_otorgados', 'descuentos_total', 'efectivo_rendido', 'transferencia_rendida', 'pago_centralizado']:
            if c in df_full.columns: df_full[c] = df_full[c].fillna(0).astype(float)
            else: df_full[c] = 0.0

        vends = pd.read_sql("""
            SELECT id as id_vendedor, nombre as "Vendedor" 
            FROM vendedores 
            WHERE area IN ('especial', 'ambos') 
            AND nombre NOT IN ('Kilaco Venta', 'Vendedor 1', 'Vendedor 2', 'nan')
        """, conn)
        
        df_full = pd.merge(df_full, vends, on='id_vendedor', how='inner')
        
        # 4. Matemáticas de cierre
        df_full['Total Ingresos'] = df_full['venta_pan'] + df_full['creditos_cobrados']
        df_full['Total Gastos'] = df_full['bencina'] + df_full['sueldo'] + df_full['otros_gastos_monto'] + df_full['creditos_otorgados'] + df_full['descuentos_total']
        df_full['Deuda Neta'] = df_full['Total Ingresos'] - df_full['Total Gastos']
        df_full['Pagado'] = df_full['efectivo_rendido'] + df_full['transferencia_rendida'] + df_full['pago_centralizado']
        df_full['Saldo'] = df_full['Deuda Neta'] - df_full['Pagado']
        
        return df_full[['fecha', 'Vendedor', 'Total Ingresos', 'Total Gastos', 'Deuda Neta', 'Pagado', 'Saldo']]
    finally:
        conn.close()

@st.cache_data(ttl=300)
def obtener_clientes_df():
    conn = get_conn()
    try:
        # Se agregan comillas a "Repartidor" para asegurar que Pandas encuentre la columna aunque Postgres devuelva minúsculas
        query = """
            SELECT c.id, c.nombre, c.direccion, c.comuna, c.telefono, 
                   v.nombre as "Repartidor", c.tipo_cliente, c.limite_credito, c.permite_credito
            FROM clientes c
            LEFT JOIN vendedores v ON c.id_vendedor_asignado = v.id
            WHERE c.activo = 1
            ORDER BY c.nombre
        """
        df = pd.read_sql(query, conn)
        return df
    finally:
        conn.close()

def crud_cliente(accion, datos=None):
    conn = get_conn()
    try:
        with conn.cursor() as c:
            limites = {"Nuevo": 50000, "Minorista": 100000, "Mayorista": 500000}
            cupo_auto = limites.get(datos.get('tipo', 'Nuevo'), 50000)
            
            # Capturamos la orden de crédito (por defecto False si no viene)
            permite = datos.get('permite_credito', False)
            
            if accion == "crear":
                c.execute("INSERT INTO clientes (nombre, direccion, comuna, telefono, id_vendedor_asignado, tipo_cliente, limite_credito, permite_credito) VALUES (%s,%s,%s,%s,%s,%s,%s,%s)",
                          (datos['nombre'], datos['dir'], datos['com'], datos['tel'], datos['id_vend'], datos['tipo'], cupo_auto, permite))
            elif accion == "editar":
                c.execute("UPDATE clientes SET nombre=%s, direccion=%s, comuna=%s, telefono=%s, id_vendedor_asignado=%s, tipo_cliente=%s, limite_credito=%s, permite_credito=%s WHERE id=%s",
                          (datos['nombre'], datos['dir'], datos['com'], datos['tel'], datos['id_vend'], datos['tipo'], cupo_auto, permite, datos['id']))
        conn.commit()
        st.cache_data.clear()
    finally:
        conn.close()

def crud_sugerencia(accion, datos=None, id_sug=None):
    """Maneja la creación, aprobación y rechazo de sugerencias de clientes."""
    conn = get_conn()
    try:
        # Purgamos el tipo numpy forzando a int nativo
        if id_sug is not None:
            id_sug = int(id_sug)

        with conn.cursor() as c:
            if accion == "crear":
                id_ref = int(datos['id_ref']) if datos.get('id_ref') else None
                c.execute("""INSERT INTO clientes_sugerencias (tipo_solicitud, id_cliente_ref, nombre, direccion, comuna, telefono, tipo_cliente, id_vendedor, comentario)
                             VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
                          (datos['tipo'], id_ref, datos['nombre'], datos.get('dir'), datos.get('com'), datos.get('tel'), datos.get('tipo_cli'), int(datos['id_vend']), datos.get('comentario')))
            elif accion in ["APROBADA", "RECHAZADA"]:
                c.execute("UPDATE clientes_sugerencias SET estado=%s WHERE id=%s", (accion, id_sug))
                if accion == "APROBADA" and datos:
                    limites = {"Nuevo": 50000, "Minorista": 100000, "Mayorista": 500000}
                    cupo = limites.get(datos.get('tipo_cliente', 'Nuevo'), 50000)
                    id_vendedor = int(datos['id_vendedor'])
                    
                    if datos['tipo_solicitud'] == 'NUEVO':
                        c.execute("INSERT INTO clientes (nombre, direccion, comuna, telefono, id_vendedor_asignado, tipo_cliente, limite_credito) VALUES (%s,%s,%s,%s,%s,%s,%s)",
                                  (datos['nombre'], datos['direccion'], datos['comuna'], datos['telefono'], id_vendedor, datos['tipo_cliente'], cupo))
                    elif datos['tipo_solicitud'] == 'EDICION':
                        id_ref = int(datos['id_cliente_ref'])
                        c.execute("UPDATE clientes SET nombre=%s, direccion=%s, comuna=%s, telefono=%s, id_vendedor_asignado=%s, tipo_cliente=%s, limite_credito=%s WHERE id=%s",
                                  (datos['nombre'], datos['direccion'], datos['comuna'], datos['telefono'], id_vendedor, datos['tipo_cliente'], cupo, id_ref))
        conn.commit()
    finally:
        conn.close()

def get_sugerencias(solo_pendientes=True, id_vend=None, modulo="ESPECIAL"):
    """Obtiene sugerencias filtradas por repartidor y por módulo (Especial/Corriente)."""
    conn = get_conn()
    try:
        q = """SELECT s.id, s.tipo_solicitud, s.id_cliente_ref, s.nombre, s.direccion, s.comuna, s.telefono, s.tipo_cliente, s.comentario, s.estado, TO_CHAR(s.fecha_creacion, 'DD/MM/YYYY') as fecha, v.nombre as "Repartidor", s.id_vendedor
               FROM clientes_sugerencias s LEFT JOIN vendedores v ON s.id_vendedor = v.id"""
        
        # Filtro inteligente: Usamos 'tipo_cliente' como bandera para saber de qué panadería viene
        filtro_mod = "s.tipo_cliente = 'CORRIENTE'" if modulo == "CORRIENTE" else "s.tipo_cliente != 'CORRIENTE'"
        
        if id_vend:
            q += f" WHERE {filtro_mod} AND s.id_vendedor = %s ORDER BY s.id DESC"
            return pd.read_sql(q, conn, params=(id_vend,))
        elif solo_pendientes:
            q += f" WHERE {filtro_mod} AND s.estado = 'PENDIENTE' ORDER BY s.id ASC"
        else:
            q += f" WHERE {filtro_mod} ORDER BY s.id DESC"
        
        return pd.read_sql(q, conn)
    finally:
        conn.close()

def crud_sugerencia_corriente(accion, datos=None, id_sug=None):
    """Maneja las sugerencias exclusivas del Pan Corriente."""
    conn = get_conn()
    try:
        if id_sug is not None: id_sug = int(id_sug)

        with conn.cursor() as c:
            if accion == "crear":
                id_ref = int(datos['id_ref']) if datos.get('id_ref') else None
                # Empaquetamos el precio dentro del comentario para no alterar la BD
                comentario_ext = f"{datos.get('comentario', '')} | PRECIO_SUG: {datos.get('precio', 0)}"
                
                c.execute("""INSERT INTO clientes_sugerencias (tipo_solicitud, id_cliente_ref, nombre, direccion, comuna, telefono, tipo_cliente, id_vendedor, comentario)
                             VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
                          (datos['tipo'], id_ref, datos['nombre'], datos.get('dir'), datos.get('com'), datos.get('tel'), 'CORRIENTE', int(datos['id_vendedor']), comentario_ext))
            
            elif accion in ["APROBADA", "RECHAZADA"]:
                c.execute("UPDATE clientes_sugerencias SET estado=%s WHERE id=%s", (accion, id_sug))
                
                if accion == "APROBADA" and datos:
                    id_vendedor = int(datos['id_vendedor'])
                    
                    # Desempaquetamos el precio
                    precio = 1600 
                    if "PRECIO_SUG:" in str(datos.get('comentario', '')):
                        try: precio = int(str(datos['comentario']).split("PRECIO_SUG:")[1].strip())
                        except: pass
                    
                    if datos['tipo_solicitud'] == 'NUEVO':
                        c.execute("INSERT INTO clientes_corriente (nombre, direccion, comuna, telefono, id_vendedor, precio_pactado, activo) VALUES (%s,%s,%s,%s,%s,%s,1)",
                                  (datos['nombre'], datos['direccion'], datos['comuna'], datos['telefono'], id_vendedor, precio))
                    elif datos['tipo_solicitud'] == 'EDICION':
                        id_ref = int(datos['id_cliente_ref'])
                        c.execute("UPDATE clientes_corriente SET nombre=%s, direccion=%s, comuna=%s, telefono=%s, id_vendedor=%s, precio_pactado=%s WHERE id=%s",
                                  (datos['nombre'], datos['direccion'], datos['comuna'], datos['telefono'], id_vendedor, precio, id_ref))
        conn.commit()
        st.cache_data.clear()
    finally:
        conn.close()

def registrar_movimiento_credito(fecha, id_cli, id_vend, tipo, monto, detalle, usuario="Sistema"):
    conn = get_conn()
    try:
        with conn.cursor() as c:
            c.execute("""INSERT INTO movimientos_credito 
                         (fecha, id_cliente, id_vendedor, tipo_movimiento, monto, detalle, creado_por) 
                         VALUES (%s,%s,%s,%s,%s,%s,%s)""",
                      (fecha, id_cli, id_vend, tipo, monto, detalle, usuario))
        conn.commit()
    finally:
        conn.close()

def editar_movimiento_credito(id_mov, fecha, id_cli, id_vend, tipo, monto, detalle, usuario):
    """Actualiza una operación de crédito dejando huella forense."""
    conn = get_conn()
    try:
        with conn.cursor() as c:
            c.execute("""UPDATE movimientos_credito 
                         SET fecha=%s, id_cliente=%s, id_vendedor=%s, tipo_movimiento=%s, monto=%s, detalle=%s, 
                             modificado_por=%s, ultima_modificacion=NOW() 
                         WHERE id=%s""",
                      (fecha, id_cli, id_vend, tipo, monto, detalle, usuario, id_mov))
        conn.commit()
    finally:
        conn.close()

def obtener_creditos_editables(es_jefatura, id_vendedor_restrictivo=None):
    """Filtro de Candado Temporal Laxo (2 días) para el formulario de corrección de créditos."""
    conn = get_conn()
    try:
        limite_temporal = "" if es_jefatura else "WHERE m.fecha >= CURRENT_DATE - INTERVAL '2 days'"
        
        filtro_vendedor = ""
        if id_vendedor_restrictivo:
            prefijo = "AND" if not es_jefatura else "WHERE"
            filtro_vendedor = f"{prefijo} m.id_vendedor = {id_vendedor_restrictivo}"
            
        query = f"""
            SELECT m.id, m.fecha, m.monto, m.tipo_movimiento, 
                   v.nombre as "Repartidor", c.nombre as "Cliente", 
                   m.id_cliente, m.id_vendedor
            FROM movimientos_credito m
            JOIN vendedores v ON m.id_vendedor = v.id
            JOIN clientes c ON m.id_cliente = c.id
            {limite_temporal}
            {filtro_vendedor}
            ORDER BY m.fecha DESC, m.id DESC
        """
        return pd.read_sql(query, conn)
    finally:
        conn.close()

def registrar_transferencia(fecha, id_v, monto, metodo, banco, tipo, verif, comentario, usuario="Sistema"):
    conn = get_conn()
    try:
        with conn.cursor() as c:
            verif_int = 1 if verif else 0
            c.execute("""INSERT INTO transferencias 
                         (fecha, id_vendedor, monto, metodo_pago, banco_emisor, verificado, tipo_transferencia, comentario, creado_por) 
                         VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
                      (fecha, id_v, monto, metodo, banco, verif_int, tipo, comentario, usuario))
        conn.commit()
    finally:
        conn.close()

def editar_transferencia(id_transf, fecha, id_v, monto, metodo, banco, tipo, usuario):
    """Actualiza una transferencia existente dejando huella forense."""
    conn = get_conn()
    try:
        with conn.cursor() as c:
            c.execute("""UPDATE transferencias 
                         SET fecha=%s, id_vendedor=%s, monto=%s, metodo_pago=%s, banco_emisor=%s, tipo_transferencia=%s, 
                             modificado_por=%s, ultima_modificacion=NOW() 
                         WHERE id=%s""",
                      (fecha, id_v, monto, metodo, banco, tipo, usuario, id_transf))
        conn.commit()
    finally:
        conn.close()

def obtener_transferencias_editables(es_jefatura, id_vendedor_restrictivo=None):
    """Filtro de Candado Temporal Laxo (2 días) para el formulario de corrección."""
    conn = get_conn()
    try:
        # Lógica: Si es jefatura, ve todo. Si es empleado, ve solo los últimos 2 días.
        limite_temporal = "" if es_jefatura else "WHERE t.fecha >= CURRENT_DATE - INTERVAL '2 days'"
        
        # Lógica: Si es repartidor, solo ve las suyas.
        filtro_vendedor = ""
        if id_vendedor_restrictivo:
            prefijo = "AND" if not es_jefatura else "WHERE"
            filtro_vendedor = f"{prefijo} t.id_vendedor = {id_vendedor_restrictivo}"
            
        query = f"""
            SELECT t.id, t.fecha, t.monto, v.nombre as "Repartidor", t.metodo_pago, t.banco_emisor, t.tipo_transferencia
            FROM transferencias t
            JOIN vendedores v ON t.id_vendedor = v.id
            {limite_temporal}
            {filtro_vendedor}
            ORDER BY t.fecha DESC, t.id DESC
        """
        return pd.read_sql(query, conn)
    finally:
        conn.close()

def actualizar_verificacion_masiva(df_cambios):
    conn = get_conn()
    try:
        with conn.cursor() as c:
            count = 0
            for i, row in df_cambios.iterrows():
                estado = 1 if row['verificado'] else 0
                c.execute("UPDATE transferencias SET verificado=%s WHERE id=%s", (estado, row['id']))
                count += 1
        conn.commit()
        st.cache_data.clear() # Limpiamos la caché para que la tabla se actualice al instante
        if count > 0: st.toast(f"{count} transferencias actualizadas.")
    finally:
        conn.close()

def guardar_movimiento_caja(fecha, area, desc, item, ing_ef, ing_tr, eg_ef, eg_tr):
    conn = get_conn()
    try:
        with conn.cursor() as c:
            c.execute("INSERT INTO caja_movimientos (fecha, area, descripcion, item, ingreso_efectivo, ingreso_transferencia, egreso, egreso_transferencia) VALUES (%s,%s,%s,%s,%s,%s,%s,%s)",
                      (fecha, area, desc, item, ing_ef, ing_tr, eg_ef, eg_tr))
        conn.commit()
        st.cache_data.clear() # Limpiamos la caché para que la tabla se actualice al instante
    finally:
        conn.close()

def editar_movimiento_caja_mill(id_mov, fecha, id_cat, id_subcat, id_entidad, nombre_entidad, detalle, ing_ef, ing_tr, eg_ef, eg_tr, usuario):
    """Actualiza una operación de caja chica dejando huella forense."""
    conn = get_conn()
    try:
        with conn.cursor() as c:
            c.execute("""
                UPDATE caja_movimientos 
                SET fecha=%s, id_categoria=%s, id_subcategoria=%s, id_entidad=%s, descripcion=%s, item=%s, 
                    ingreso_efectivo=%s, ingreso_transferencia=%s, egreso=%s, egreso_transferencia=%s, 
                    modificado_por=%s, ultima_modificacion=NOW()
                WHERE id=%s
            """, (fecha, id_cat, id_subcat, id_entidad, nombre_entidad, detalle, ing_ef, ing_tr, eg_ef, eg_tr, usuario, id_mov))
        conn.commit()
        st.cache_data.clear()
    finally:
        conn.close()

def obtener_caja_editables(area, es_jefatura):
    """Filtro de Candado Temporal Laxo (2 días) para corrección de caja."""
    conn = get_conn()
    try:
        limite_temporal = "" if es_jefatura else "AND fecha >= CURRENT_DATE - INTERVAL '2 days'"
        query = f"""
            SELECT * FROM caja_movimientos 
            WHERE area=%s {limite_temporal}
            ORDER BY fecha DESC, id DESC
        """
        return pd.read_sql(query, conn, params=(area,))
    finally:
        conn.close()

# ==========================================
# NUEVAS FUNCIONES DE CAJA (MILL)
# ==========================================

@st.cache_data(ttl=60)
def obtener_categorias_caja():
    """Trae las categorías maestras."""
    conn = get_conn()
    try:
        df = pd.read_sql("SELECT id, nombre FROM caja_categorias WHERE estado='ACTIVO' ORDER BY id ASC", conn)
        return df
    finally:
        conn.close()

@st.cache_data(ttl=60)
def obtener_subcategorias_caja(id_categoria):
    """Filtra las subcategorías dependiendo de la categoría seleccionada."""
    conn = get_conn()
    try:
        df = pd.read_sql("SELECT id, nombre FROM caja_subcategorias WHERE id_categoria=%s AND estado='ACTIVO' ORDER BY id ASC", conn, params=(id_categoria,))
        return df
    finally:
        conn.close()

@st.cache_data(ttl=60)
def obtener_entidades_caja(modulo="Especial"):
    """
    Trae las entidades. Forzamos UPPER en SQL para evitar conflictos de mayúsculas/minúsculas.
    """
    conn = get_conn()
    try:
        df = pd.read_sql("SELECT id, nombre, tipo, alcance FROM entidades WHERE estado='ACTIVO' AND (UPPER(alcance)='GLOBAL' OR UPPER(alcance)=%s) ORDER BY nombre ASC", conn, params=(modulo.upper(),))
        return df
    finally:
        conn.close()

def guardar_movimiento_caja_mill(fecha, area, id_cat, id_subcat, id_entidad, nombre_entidad, detalle, ing_ef, ing_tr, eg_ef, eg_tr):
    """
    Guarda el movimiento en la caja, firmándolo silenciosamente con el nombre y rol 
    del usuario actual para permitir la separación Caja Chica vs Libro Mayor.
    """
    conn = get_conn()
    # Capturamos la firma del operador desde la sesión
    usuario = st.session_state.get("user_name", "Desconocido")
    rol = st.session_state.get("user_role", "sin_rol")
    
    try:
        with conn.cursor() as c:
            # Guardamos tanto los nuevos IDs como el 'nombre_entidad' en el campo 'descripcion' antiguo 
            # para no romper el historial hacia atrás.
            c.execute("""
                INSERT INTO caja_movimientos 
                (fecha, area, id_categoria, id_subcategoria, id_entidad, descripcion, item, ingreso_efectivo, ingreso_transferencia, egreso, egreso_transferencia, usuario_creador, rol_creador)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """, (fecha, area, id_cat, id_subcat, id_entidad, nombre_entidad, detalle, ing_ef, ing_tr, eg_ef, eg_tr, usuario, rol))
        conn.commit()
        st.cache_data.clear() # Limpiamos la caché para que la tabla se actualice al instante
    finally:
        conn.close()

def obtener_caja_mayor_global(fecha):
    """Ignora el área y extrae la consolidación total de la panadería para el día seleccionado."""
    conn = get_conn()
    try:
        query = """
            SELECT id, fecha, area, descripcion as entidad, item as detalle, 
                   ingreso_efectivo, ingreso_transferencia, 
                   egreso as egreso_efectivo, egreso_transferencia, rol_creador 
            FROM caja_movimientos 
            WHERE fecha=%s
            ORDER BY id DESC
        """
        return pd.read_sql(query, conn, params=(fecha,))
    finally:
        conn.close()

# --- Funciones Pan Corriente ---
def get_despacho_corriente(fecha, id_vendedor):
    """Obtiene la ruta diaria, filtrando estrictamente a los clientes activos."""
    conn = get_conn()
    try:
        # El Candado SQL: Agregamos 'AND c.activo = 1'
        query_base = """
            SELECT d.*, c.nombre as "cliente"
            FROM despacho_corriente d
            JOIN clientes_corriente c ON d.id_cliente = c.id
            WHERE d.fecha = %s AND d.id_vendedor = %s AND c.activo = 1
            ORDER BY c.nombre
        """
        df = pd.read_sql(query_base, conn, params=(fecha, id_vendedor))
        
        if df.empty:
            clientes = pd.read_sql("SELECT id, nombre, precio_pactado FROM clientes_corriente WHERE id_vendedor = %s AND activo = 1 AND nombre NOT IN ('Racion', 'Adicional')", conn, params=(id_vendedor,))
            if not clientes.empty:
                with conn.cursor() as c:
                    for i, cli in clientes.iterrows():
                        c.execute("SELECT deuda_final FROM despacho_corriente WHERE id_cliente = %s AND fecha < %s ORDER BY fecha DESC LIMIT 1", (int(cli['id']), fecha))
                        deuda_ant = c.fetchone()
                        saldo_ayer = deuda_ant[0] if deuda_ant else 0
                        c.execute("INSERT INTO despacho_corriente (fecha, id_cliente, id_vendedor, precio_aplicado, saldo_anterior, deuda_final) VALUES (%s, %s, %s, %s, %s, %s)", (fecha, int(cli['id']), id_vendedor, int(cli['precio_pactado']), saldo_ayer, saldo_ayer))
                conn.commit()
                df = pd.read_sql(query_base, conn, params=(fecha, id_vendedor))
        return df
    finally:
        conn.close()

def save_despacho_corriente(df_cambios):
    """Guarda los 8 turnos de despacho y recalcula la deuda."""
    conn = get_conn()
    try:
        with conn.cursor() as c:
            for i, row in df_cambios.iterrows():
                # Ahora sumamos hasta la carga_8
                total_kg = sum([row[f'carga_{k}'] for k in range(1,9)])
                ventas = int(total_kg * row['precio_aplicado'])
                total_pagar = ventas + row['saldo_anterior']
                deuda = total_pagar - row['paga'] - row['pago_centralizado']
                c.execute("""
                    UPDATE despacho_corriente SET 
                    carga_1=%s, carga_2=%s, carga_3=%s, carga_4=%s, carga_5=%s, carga_6=%s, carga_7=%s, carga_8=%s,
                    total_carga=%s, ventas_monto=%s, total_pagar=%s, paga=%s, pago_centralizado=%s, deuda_final=%s
                    WHERE id=%s
                """, (row['carga_1'], row['carga_2'], row['carga_3'], row['carga_4'], row['carga_5'], row['carga_6'], row['carga_7'], row['carga_8'],
                      total_kg, ventas, total_pagar, row['paga'], row['pago_centralizado'], deuda, row['id']))
        conn.commit()
    finally:
        conn.close()

def get_produccion_corriente_unificada(fecha):
    conn = get_conn()
    try:
        # Reemplazamos SUM por MAX para evitar multiplicar la carga por la cantidad de clientes
        df_repartidores = pd.read_sql("""
            SELECT v.nombre as "Concepto",
            COALESCE(MAX(d.carga_1), 0) + COALESCE(MAX(d.carga_2), 0) as rinde_noche,
            COALESCE(MAX(d.carga_3), 0) + COALESCE(MAX(d.carga_4), 0) + COALESCE(MAX(d.carga_5), 0) + COALESCE(MAX(d.carga_6), 0) + COALESCE(MAX(d.carga_7), 0) + COALESCE(MAX(d.carga_8), 0) as rinde_dia
            FROM despacho_corriente d
            JOIN vendedores v ON d.id_vendedor = v.id
            WHERE d.fecha = %s
            GROUP BY v.nombre
        """, conn, params=(fecha,))
        
        ex = None
        with conn.cursor() as c:
            c.execute("SELECT * FROM produccion_extras WHERE fecha=%s", (fecha,))
            ex = c.fetchone()
            if not ex:
                c.execute("INSERT INTO produccion_extras (fecha) VALUES (%s)", (fecha,))
                conn.commit()
                c.execute("SELECT * FROM produccion_extras WHERE fecha=%s", (fecha,))
                ex = c.fetchone()
        
        extras_data = [
            {"Concepto": "Ración", "rinde_noche": ex[3], "rinde_dia": ex[2]},
            {"Concepto": "Adicional", "rinde_noche": ex[5], "rinde_dia": ex[4]},
            {"Concepto": "Kilaco", "rinde_noche": ex[7], "rinde_dia": ex[6]}
        ]
        df_extras = pd.DataFrame(extras_data)
        df_final = pd.concat([df_repartidores, df_extras], ignore_index=True)
        extras_dict = {"rd": ex[2], "rn": ex[3], "ad": ex[4], "an": ex[5], "kd": ex[6], "kn": ex[7]}
        return df_final, extras_dict
    finally:
        conn.close()

def save_extras_produccion(fecha, d):
    conn = get_conn()
    try:
        with conn.cursor() as c:
            c.execute("""
                UPDATE produccion_extras SET 
                racion_dia=%s, racion_noche=%s, 
                adicional_dia=%s, adicional_noche=%s, 
                kilaco_dia=%s, kilaco_noche=%s
                WHERE fecha=%s
            """, (val_db(d['rd']), val_db(d['rn']), val_db(d['ad']), val_db(d['an']), val_db(d['kd']), val_db(d['kn']), fecha))
        conn.commit()
    finally:
        conn.close()

def get_resumen_visor_corriente(fi, ff):
    conn = get_conn()
    try:
        df_fin = pd.read_sql("SELECT f.*, v.nombre as \"Vendedor\" FROM finanzas_corriente f JOIN vendedores v ON f.id_vendedor = v.id WHERE f.fecha BETWEEN %s AND %s", conn, params=(fi, ff))
        df_calle = pd.read_sql("""
            SELECT d.fecha, v.nombre as "Vendedor", SUM(d.deuda_final) as saldo_clientes
            FROM despacho_corriente d
            JOIN vendedores v ON d.id_vendedor = v.id
            WHERE d.fecha BETWEEN %s AND %s
            GROUP BY d.fecha, d.id_vendedor, v.nombre
        """, conn, params=(fi, ff))
        
        df_full = pd.merge(df_fin, df_calle, on=['fecha', 'Vendedor'], how='outer').fillna(0)
        return df_full[['fecha', 'Vendedor', 'total_gastos', 'saldo_final', 'saldo_clientes']]
    finally:
        conn.close()

@st.cache_data(ttl=300)
def obtener_clientes_corriente():
    """Trae TODOS los clientes del corriente, ahora con su simetría espacial y de contacto."""
    conn = get_conn()
    try:
        df = pd.read_sql("""
            SELECT cc.id, cc.nombre, cc.direccion, cc.comuna, cc.telefono, cc.precio_pactado, 
                   v.nombre as "Repartidor", cc.id_vendedor, cc.activo
            FROM clientes_corriente cc
            LEFT JOIN vendedores v ON cc.id_vendedor = v.id
            ORDER BY cc.nombre
        """, conn)
        return df
    finally:
        conn.close()

def crud_cliente_corriente(accion, datos):
    """Crea o edita clientes del corriente con sus nuevos metadatos."""
    conn = get_conn()
    try:
        with conn.cursor() as c:
            if accion == "crear":
                c.execute("""INSERT INTO clientes_corriente 
                             (nombre, direccion, comuna, telefono, id_vendedor, precio_pactado, activo) 
                             VALUES (%s,%s,%s,%s,%s,%s,1)""", 
                          (datos['nombre'], datos.get('dir','-'), datos.get('com','-'), datos.get('tel','-'), datos['id_vendedor'], datos['precio']))
            elif accion == "editar":
                activo_val = 1 if datos.get('activo', True) else 0
                c.execute("""UPDATE clientes_corriente 
                             SET nombre=%s, direccion=%s, comuna=%s, telefono=%s, id_vendedor=%s, precio_pactado=%s, activo=%s 
                             WHERE id=%s""", 
                          (datos['nombre'], datos.get('dir','-'), datos.get('com','-'), datos.get('tel','-'), datos['id_vendedor'], datos['precio'], activo_val, datos['id']))
        conn.commit()
        st.cache_data.clear()
    finally:
        conn.close()

def get_finanzas_corriente(fecha, id_v):
    conn = get_conn()
    try:
        with conn.cursor() as c:
            c.execute("SELECT * FROM finanzas_corriente WHERE fecha=%s AND id_vendedor=%s", (fecha, id_v))
            row = c.fetchone()
            
            c.execute("SELECT SUM(paga + pago_centralizado) FROM despacho_corriente WHERE fecha=%s AND id_vendedor=%s", (fecha, id_v))
            res_recaudo = c.fetchone()
            recaudo_auto = int(res_recaudo[0]) if res_recaudo and res_recaudo[0] else 0
            
            c.execute("SELECT SUM(ventas_monto) FROM despacho_corriente WHERE fecha=%s AND id_vendedor=%s", (fecha, id_v))
            res_ventas = c.fetchone()
            ventas_auto = int(res_ventas[0]) if res_ventas and res_ventas[0] else 0
            
            # Traemos la bencina y comisión directo del perfil del vendedor
            c.execute("SELECT bencina, comision FROM vendedores WHERE id=%s", (id_v,))
            perfil = c.fetchone()
            bencina_auto = int(perfil[0]) if perfil and perfil[0] else 0
            factor_comision = float(perfil[1]) if perfil and perfil[1] else 0.04

        comision_auto = int(ventas_auto * factor_comision)
        
        if row:
            return {
                "id": row[0], "venta": row[3], "recaudo": recaudo_auto, "comision": comision_auto, 
                "bencina": row[6], "sueldo": row[7], "otros": row[8], "det": row[9], 
                "efec": row[11], "trans": row[12], "venta_real": ventas_auto
            }
        return {
            "id": None, "venta": ventas_auto, "recaudo": recaudo_auto, "comision": comision_auto, 
            "bencina": bencina_auto, "sueldo": 0, "otros": 0, "det": "", "efec": 0, "trans": 0, "venta_real": ventas_auto
        }
    finally:
        conn.close()

def save_finanzas_corriente(fecha, id_v, d):
    """Guarda la rendición del corriente, purgando tipos NumPy para evitar crasheos de psycopg2."""
    conn = get_conn()
    
    # EL PURIFICADOR
    venta = int(d['venta'])
    recaudo = int(d['recaudo_total'])
    comision = int(d['comision'])
    bencina = int(d['bencina'])
    sueldo = int(d['sueldo'])
    otros = int(d['otros'])
    efec = int(d['efec'])
    trans = int(d['trans'])
    
    gastos = bencina + sueldo + otros + comision
    saldo = recaudo - gastos - efec - trans
    
    try:
        with conn.cursor() as c:
            c.execute("SELECT id FROM finanzas_corriente WHERE fecha=%s AND id_vendedor=%s", (fecha, id_v))
            ex = c.fetchone()
            if ex:
                c.execute("""UPDATE finanzas_corriente SET venta_diaria=%s, recaudo_diario=%s, comision=%s, bencina=%s, sueldo=%s, otros_gastos=%s, detalle_gastos=%s, total_gastos=%s, pago_efectivo=%s, pago_transferencia=%s, saldo_final=%s WHERE id=%s""", 
                          (venta, recaudo, comision, bencina, sueldo, otros, d['det'], gastos, efec, trans, saldo, ex[0]))
            else:
                c.execute("""INSERT INTO finanzas_corriente (fecha, id_vendedor, venta_diaria, recaudo_diario, comision, bencina, sueldo, otros_gastos, detalle_gastos, total_gastos, pago_efectivo, pago_transferencia, saldo_final) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""", 
                          (fecha, id_v, venta, recaudo, comision, bencina, sueldo, otros, d['det'], gastos, efec, trans, saldo))
        conn.commit()
    finally:
        conn.close()

def obtener_deuda_especial_repartidor(fecha, id_vendedor):
    """Calcula el dinero exacto que un repartidor debe por el pan especial que se llevó (aplicando dscto dinámico)."""
    conn = get_conn()
    try:
        with conn.cursor() as c:
            c.execute("SELECT dscto_esp FROM vendedores WHERE id=%s", (id_vendedor,))
            res = c.fetchone()
            descuento = float(res[0]) if res and res[0] else 0.0
            
        query = """
            SELECT d.saldo_anterior, d.carga, d.devolucion_muestra, d.saldo_actual, p.precio_estandar
            FROM despacho d
            JOIN productos p ON d.id_producto = p.id
            WHERE d.fecha = %s AND d.id_vendedor = %s
        """
        df = pd.read_sql(query, conn, params=(fecha, id_vendedor))
        if df.empty: return 0
        
        df['disp'] = df['saldo_anterior'] + df['carga']
        df['venta'] = (df['disp'] - df['devolucion_muestra'] - df['saldo_actual']).clip(lower=0)
        df['precio_final'] = df['precio_estandar'] * (1 - descuento)
        df['total'] = df['venta'] * df['precio_final']
        return int(df['total'].sum())
    finally:
        conn.close()

def get_admin_table(tabla, fecha):
    conn = get_conn()
    if tabla == "despacho_corriente":
        q = f"SELECT d.*, c.nombre as \"Cliente\", v.nombre as \"Vendedor\" FROM {tabla} d JOIN clientes_corriente c ON d.id_cliente=c.id JOIN vendedores v ON d.id_vendedor=v.id WHERE d.fecha=%s"
    elif tabla == "finanzas_corriente":
        q = f"SELECT f.*, v.nombre as \"Vendedor\" FROM {tabla} f JOIN vendedores v ON f.id_vendedor=v.id WHERE f.fecha=%s"
    else:
        q = f"SELECT * FROM {tabla} WHERE fecha=%s"
    
    try:
        df = pd.read_sql(q, conn, params=(fecha,))
        return df
    except:
        return pd.DataFrame()
    finally:
        conn.close()

def obtener_estado_creditos_vendedor(id_vendedor, fecha_corte):
    """Genera el reporte consolidado de estado de deuda para un vendedor hasta una fecha específica."""
    conn = get_conn()
    try:
        # 1. Obtener clientes del vendedor
        df_cli = pd.read_sql("SELECT id, nombre, limite_credito FROM clientes WHERE id_vendedor_asignado=%s AND activo=1", conn, params=(id_vendedor,))
        
        if df_cli.empty:
            return pd.DataFrame()

        # 2. Calcular deudas (Sumar Creditos - Sumar Abonos)
        # Nota: Hacemos una query agregada para ser eficientes
        ids_clientes = tuple(df_cli['id'].tolist())
        if not ids_clientes: return pd.DataFrame() # Doble check
        
        # EL CANDADO TEMPORAL: Sumamos solo los movimientos ocurridos hasta la fecha de corte
        query_movs = f"""
            SELECT id_cliente, 
            SUM(CASE WHEN tipo_movimiento='CREDITO' THEN monto ELSE 0 END) as total_otorgado,
            SUM(CASE WHEN tipo_movimiento='ABONO' THEN monto ELSE 0 END) as total_pagado
            FROM movimientos_credito 
            WHERE id_cliente IN {ids_clientes} AND fecha <= %s
            GROUP BY id_cliente
        """
        # Fix para tupla de un solo elemento en Python (x,)
        if len(ids_clientes) == 1: query_movs = query_movs.replace(f"({ids_clientes[0]})", f"({ids_clientes[0]})") 
        
        df_movs = pd.read_sql(query_movs, conn, params=(fecha_corte,))
        
        # 3. Fusionar datos
        df_final = pd.merge(df_cli, df_movs, left_on='id', right_on='id_cliente', how='left').fillna(0)
        
        # 4. Calcular métricas finales
        df_final['deuda_actual'] = df_final['total_otorgado'] - df_final['total_pagado']
        
        def calcular_estado(row):
            limite = row['limite_credito']
            deuda = row['deuda_actual']
            if limite == 0: return "Sin Límite"
            if deuda > limite: return "Excedido 🔴"
            if deuda >= (limite * 0.8): return "Cercano al Límite 🟡"
            return "Al día 🟢"

        df_final['estado'] = df_final.apply(calcular_estado, axis=1)
        
        return df_final[['nombre', 'limite_credito', 'total_otorgado', 'total_pagado', 'deuda_actual', 'estado']]
    finally:
        conn.close()

def obtener_historial_movimientos_credito(f_ini, f_fin, id_vendedor=None):
    """Extrae el historial de movimientos de crédito (fiados y abonos) para la nueva auditoría visual."""
    conn = get_conn()
    try:
        filtro_vendedor = f"AND m.id_vendedor = {id_vendedor}" if id_vendedor else ""
        query = f"""
            SELECT TO_CHAR(m.fecha, 'DD/MM/YYYY') as "Fecha", 
                   v.nombre as "Repartidor", c.nombre as "Cliente", 
                   m.tipo_movimiento as "Tipo", m.monto as "Monto", m.detalle as "Detalle",
                   m.fecha as fecha_real
            FROM movimientos_credito m
            JOIN vendedores v ON m.id_vendedor = v.id
            JOIN clientes c ON m.id_cliente = c.id
            WHERE m.fecha BETWEEN %s AND %s
            {filtro_vendedor}
            ORDER BY m.fecha DESC, m.id DESC
        """
        df = pd.read_sql(query, conn, params=(f_ini, f_fin))
        if not df.empty:
            df = df.drop(columns=['fecha_real'])
        return df
    finally:
        conn.close()

def obtener_reporte_transferencias_filtrado(fecha_inicio, fecha_fin, id_vendedor):
    """Reporte específico de transferencias para un rango de fechas."""
    conn = get_conn()
    try:
        query = """
            SELECT TO_CHAR(fecha, 'DD/MM/YYYY') as "Fecha", monto, banco_emisor, metodo_pago as banco_receptor_info, tipo_transferencia as tipo, verificado
            FROM transferencias 
            WHERE fecha BETWEEN %s AND %s AND id_vendedor = %s
            ORDER BY fecha DESC, id DESC
        """
        df = pd.read_sql(query, conn, params=(fecha_inicio, fecha_fin, id_vendedor))
        
        # Mapeo estético
        if not df.empty:
            df['verificado'] = df['verificado'].map({1: "Recibido ✅", 0: "Pendiente ⏳"})
        return df
    finally:
        conn.close()

def obtener_resumen_bandejas_especial(fecha):
    """Calcula el saldo final de bandejas para los repartidores de Pan Especial de forma dinámica."""
    conn = get_conn()
    try:
        # LA MAGIA SQL: Filtramos directamente por el área 'especial' o 'ambos'
        query_vendedores = """
            SELECT id, nombre 
            FROM vendedores 
            WHERE area IN ('especial', 'ambos')
            AND nombre NOT IN ('Kilaco Venta', 'Vendedor 1', 'Vendedor 2', 'nan') 
            AND nombre IS NOT NULL
        """
        vends_especial = pd.read_sql(query_vendedores, conn)
        
        resultados = []
        with conn.cursor() as c:
            for _, row in vends_especial.iterrows():
                id_v = row['id']
                nombre = row['nombre']
                
                # Lógica idéntica a obtener_bandejas pero optimizada para resumen
                c.execute("SELECT saldo_final FROM control_bandejas WHERE fecha=%s AND id_vendedor=%s", (fecha, id_v))
                hoy = c.fetchone()
                
                if hoy:
                    saldo = hoy[0]
                else:
                    c.execute("SELECT saldo_final FROM control_bandejas WHERE id_vendedor=%s AND fecha < %s ORDER BY fecha DESC LIMIT 1", (id_v, fecha))
                    hist = c.fetchone()
                    saldo = hist[0] if hist else 0
                
                resultados.append({"Repartidor": nombre, "Saldo Final Bandejas": saldo})
        
        return pd.DataFrame(resultados)
    finally:
        conn.close()

def obtener_deudas_corriente_vendedor(id_vendedor):
    """Consulta la deuda final más reciente de los clientes activos de un vendedor."""
    conn = get_conn()
    try:
        # DISTINCT ON extrae solo el último registro (el más reciente) por cliente
        query = """
            SELECT DISTINCT ON (d.id_cliente) c.nombre as "Cliente", d.deuda_final as "Deuda Actual"
            FROM despacho_corriente d
            JOIN clientes_corriente c ON d.id_cliente = c.id
            WHERE d.id_vendedor = %s AND c.activo = 1
            ORDER BY d.id_cliente, d.fecha DESC
        """
        df = pd.read_sql(query, conn, params=(id_vendedor,))
        return df
    finally:
        conn.close()

def obtener_cobranzas_pendientes_especial():
    """Cruza Despacho vs Finanzas para los últimos 15 días, filtrando nativamente por área."""
    conn = get_conn()
    try:
        # Añadimos AND v.area IN ('especial', 'ambos') directo en SQL
        query = """
            SELECT TO_CHAR(d.fecha, 'DD/MM/YYYY') as "Fecha Pendiente", v.nombre as "Repartidor"
            FROM (
                SELECT fecha, id_vendedor 
                FROM despacho 
                WHERE fecha >= CURRENT_DATE - INTERVAL '15 days' AND carga > 0 
                GROUP BY fecha, id_vendedor
            ) d
            JOIN vendedores v ON d.id_vendedor = v.id
            LEFT JOIN finanzas f ON d.fecha = f.fecha AND d.id_vendedor = f.id_vendedor
            WHERE f.id IS NULL AND d.fecha < CURRENT_DATE
            AND v.area IN ('especial', 'ambos') 
            AND v.nombre NOT IN ('Kilaco Venta', 'Hugo Palacios') 
            ORDER BY d.fecha DESC, v.nombre ASC
        """
        df = pd.read_sql(query, conn)
        return df
    finally:
        conn.close()

# ==========================================
# REGIÓN 4: VISTAS Y COMPONENTES
# ==========================================

# inicialización de la sesión
init_session()

def login_view():
    """Vista de inicio de sesión conectada a Supabase Auth."""
    
    st.write("")
    st.write("")
    
    # Proporción calibrada para la tarjeta central en PC
    c1, c2, c3 = st.columns([1, 1.2, 1])
    
    with c2:
        # EL TRUCO DEFINITIVO: HTML puro con Base64 para forzar el centrado absoluto
        try:
            import base64
            with open("logo.png", "rb") as image_file:
                encoded_string = base64.b64encode(image_file.read()).decode()
            
            # Se inyecta como HTML inmutable
            st.markdown(f"""
                <div style="display: flex; justify-content: center; margin-bottom: 10px;">
                    <img src="data:image/png;base64,{encoded_string}" width="140" style="max-width: 100%;">
                </div>
            """, unsafe_allow_html=True)
        except Exception as e:
            # Fallback en caso de que logo.png no se encuentre
            st.markdown("<h1 style='text-align: center; font-size: 5rem; margin:0;'>🥖</h1>", unsafe_allow_html=True)
        
        st.markdown("<h3 style='text-align: center; margin-top: 0px;'>Inicio de Sesión</h3>", unsafe_allow_html=True)
        
        with st.form("login_form"):
            email = st.text_input("Correo Electrónico") 
            pw = st.text_input("Contraseña", type="password")
            mantener = st.checkbox("Mantener sesión iniciada")
            
            st.write("")
            if st.form_submit_button("Entrar", type="primary", use_container_width=True):
                if not email or not pw:
                    st.warning("Por favor ingresa correo y contraseña.")
                else:
                    user_data = check_login(email, pw) 
                    
                    if user_data:
                        st.session_state.logged_in = True
                        st.session_state.user_id = user_data["id"]
                        st.session_state.user_name = user_data["nombre"]
                        st.session_state.user_role = user_data["rol"]
                        st.session_state.id_vendedor = user_data["id_vendedor"]
                        
                        if mantener: 
                            st.query_params["session"] = "active"
                        st.rerun()
                    else: 
                        st.error("Credenciales inválidas o error de conexión.")

def menu_view():
    """Pantalla de selección de módulos bajo doctrina MILL."""
    contenedor_menu = st.empty()
    
    with contenedor_menu.container():
        rol = st.session_state.user_role
        nombre = st.session_state.user_name

        # Encabezado jerárquico y limpio
        c1, c2, c3 = st.columns([1, 4, 1], vertical_alignment="center")
        with c1:
            try: st.image("logo.png", width=60)
            except: st.markdown("### 🥖")
        with c2:
            st.markdown(f"<h3 style='margin: 0; color: #2C3E50;'>{nombre}</h3>", unsafe_allow_html=True)
            st.caption(f"{str(rol).upper()}")
        with c3:
            if st.button("Cerrar Sesión", use_container_width=True):
                contenedor_menu.empty() 
                st.session_state.logged_in = False
                st.rerun()

        st.divider()

        ver_especial = rol in ["admin", "pan_especial", "supervisor", "repartidor_esp", "repartidor_corr", "cajero_integral", "despacho_especial"]
        ver_corriente = rol in ["admin", "pan_corriente", "supervisor", "repartidor_corr", "cajero_integral"]

        st.write("") 

        col_l, col_1, col_2, col_r = st.columns([1, 2, 2, 1])

        # SVGs intocables
        svg_esp = '''<svg xmlns="http://www.w3.org/2000/svg" width="55" height="55" fill="#556B2F" class="bi bi-box-seam" viewBox="0 0 16 16"> <path d="M8.186 1.113a.5.5 0 0 0-.372 0L1.846 3.5l2.404.961L10.404 2zm3.564 1.426L5.596 5 8 5.961 14.154 3.5zm3.25 1.7-6.5 2.6v7.922l6.5-2.6V4.24zM7.5 14.762V6.84L1 4.239v7.923zM7.443.184a1.5 1.5 0 0 1 1.114 0l7.129 2.852A.5.5 0 0 1 16 3.5v8.662a1 1 0 0 1-.629.958l-7.185 2.872a1.5 1.5 0 0 1-1.114 0l-7.185-2.872A1 1 0 0 1 0 12.162V3.5a.5.5 0 0 1 .314-.464z"/> </svg>'''
        svg_corr = '''<svg xmlns="http://www.w3.org/2000/svg" width="55" height="55" fill="#556B2F" class="bi bi-truck" viewBox="0 0 16 16"> <path d="M0 3.5A1.5 1.5 0 0 1 1.5 2h9A1.5 1.5 0 0 1 12 3.5V5h1.02a1.5 1.5 0 0 1 1.17.563l1.481 1.85a1.5 1.5 0 0 1 .329.938V10.5a1.5 1.5 0 0 1-1.5 1.5H14a2 2 0 1 1-4 0H5a2 2 0 1 1-3.998-.085A1.5 1.5 0 0 1 0 10.5v-7zm1.294 7.456A1.999 1.999 0 0 1 4.732 11h5.536a2.01 2.01 0 0 1 .732-.732V3.5a.5.5 0 0 0-.5-.5h-9a.5.5 0 0 0-.5.5v7a.5.5 0 0 0 .294.456zM12 10a2 2 0 0 1 1.732 1h.768a.5.5 0 0 0 .5-.5V8.35a.5.5 0 0 0-.11-.312l-1.48-1.85A.5.5 0 0 0 13.02 6H12v4zm-9 1a1 1 0 1 0 0 2 1 1 0 0 0 0-2zm9 0a1 1 0 1 0 0 2 1 1 0 0 0 0-2z"/> </svg>'''

        with col_1:
            if ver_especial:
                with st.container(border=True):
                    st.markdown(f"""
                        <div style="text-align: center; padding: 10px 0;">
                            <div style="display: flex; justify-content: center; margin-bottom: 15px;">{svg_esp}</div>
                            <h3 style="margin: 0 0 5px 0; color: #333;">Pan Especial</h3>
                            <p style="color: #666; font-size: 14px; margin: 0 0 20px 0;">Insumos, Despachos y Créditos</p>
                        </div>
                    """, unsafe_allow_html=True)
                    if st.button("Abrir", key="btn_esp_entry", use_container_width=True, type="primary"):
                        contenedor_menu.empty() 
                        st.session_state.current_module = "especial"
                        st.rerun()
            else:
                st.info("Módulo restringido")

        with col_2:
            if ver_corriente:
                with st.container(border=True):
                    st.markdown(f"""
                        <div style="text-align: center; padding: 10px 0;">
                            <div style="display: flex; justify-content: center; margin-bottom: 15px;">{svg_corr}</div>
                            <h3 style="margin: 0 0 5px 0; color: #333;">Pan Corriente</h3>
                            <p style="color: #666; font-size: 14px; margin: 0 0 20px 0;">Producción y Logística</p>
                        </div>
                    """, unsafe_allow_html=True)
                    if st.button("Abrir", key="btn_corr_entry", use_container_width=True, type="primary"):
                        contenedor_menu.empty() 
                        st.session_state.current_module = "corriente"
                        st.rerun()
            else:
                st.info("Módulo restringido")

        st.markdown("<div style='margin-top: 60px; text-align: center; color: #aaa; font-size: 12px;'>Kilaco ERP v2.0</div>", unsafe_allow_html=True)
        
# ----------------------------------------------------
# APLICACIÓN PAN ESPECIAL (CORREGIDA: FILTRO VENDEDOR)
# ----------------------------------------------------
from streamlit_option_menu import option_menu

def app_pan_especial():
    rol = st.session_state.user_role
    mi_id_vendedor = st.session_state.id_vendedor
    es_repartidor = rol in ["repartidor_esp", "repartidor_corr"]
    
    # --- 1. MENU Y PERMISOS ---
    opciones_full = ["Insumos", "Producción", "Despacho", "Cobranza", "Clientes", "Créditos", "Transferencias", "Caja"]
    iconos_full   = ["box-seam", "tools",       "truck",     "currency-dollar", "people",   "credit-card", "bank",             "cash-stack"]
    
    if rol in ["admin", "pan_especial", "supervisor"]:
        menu_options = opciones_full; menu_icons = iconos_full; permiso_editar = True
    elif rol == "cajero_integral":
        indices = [3, 5, 6, 7] # Cobranza, Créditos, Transferencias, Caja
        menu_options = [opciones_full[i] for i in indices]; menu_icons = [iconos_full[i] for i in indices]; permiso_editar = True
    elif rol == "despacho_especial":
        indices = [0, 1, 2, 3] # Insumos, Producción, Despacho, Cobranza
        menu_options = [opciones_full[i] for i in indices]; menu_icons = [iconos_full[i] for i in indices]; permiso_editar = True
    elif rol == "repartidor_esp":
        indices = [2, 3, 4, 5, 6] 
        menu_options = [opciones_full[i] for i in indices]; menu_icons = [iconos_full[i] for i in indices]; permiso_editar = False
    elif rol == "repartidor_corr":
        menu_options = ["Despacho"]; menu_icons = ["truck"]; permiso_editar = False 
    else:
        menu_options = []; menu_icons = []; permiso_editar = False

    if rol == "admin": menu_options.append("Admin"); menu_icons.append("gear")

    # --- 2. SIDEBAR ---
    with st.sidebar:
        c_logo1, c_logo2, c_logo3 = st.columns([1, 1.5, 1])
        with c_logo2:
            try: st.image("logo.png", use_container_width=True)
            except: st.write("🥖")
        st.markdown("<h3 style='text-align: center; margin-top:0px;'>KILACO ERP</h3>", unsafe_allow_html=True)
        if not menu_options: st.error("Sin acceso."); st.stop()
        
        seleccion = option_menu(
            None, menu_options, icons=menu_icons, menu_icon="cast", default_index=0, 
            styles={"container": {"padding": "0!important", "background-color": "#ffffff"}, "icon": {"color": "#556B2F", "font-size": "14px"}, "nav-link": {"font-size": "14px", "text-align": "left", "margin":"0px"}, "nav-link-selected": {"background-color": "#556B2F", "font-weight": "600"}}
        )
        st.markdown("---")
        if st.button("Volver al Menú", use_container_width=True): st.session_state.current_module = "menu"; st.rerun()

    # --- 3. REFERENCIAS DINÁMICAS DESDE DB ---
    l_comunas, l_bancos, df_vend = get_referencias() 
    dict_vend = dict(zip(df_vend['nombre'], df_vend['id']))
    
    lista_repartidores_todos = df_vend['nombre'].tolist()
    lista_repartidores_esp = df_vend[df_vend['area'].isin(['especial', 'ambos'])]['nombre'].tolist()
    repartidores_corriente = df_vend[df_vend['area'].isin(['corriente', 'ambos'])]['nombre'].tolist()
    
    # LISTA PURA: Extraemos a Kilaco Venta para usarla en Créditos y Transferencias
    lista_repartidores_esp_puro = [r for r in lista_repartidores_esp if r != "Kilaco Venta"]
    
    nombre_vendedor_actual = None
    if es_repartidor and mi_id_vendedor:
        match = df_vend[df_vend['id'] == mi_id_vendedor]
        if not match.empty: nombre_vendedor_actual = match['nombre'].iloc[0]

    idx_defecto = lista_repartidores_todos.index(nombre_vendedor_actual) if es_repartidor and nombre_vendedor_actual in lista_repartidores_todos else None

    # --- 4. VISTAS ---

    if seleccion == "Insumos":
        st.title("Control de Insumos")
        c1, c2 = st.columns([1, 4])
        f_ins = c1.date_input("Fecha", date.today(), format="DD/MM/YYYY")
        
        tab_bandejas, tab_bolsas = st.tabs(["Bandejas", "Bolsas"])

        svg_tray = '''<svg xmlns="http://www.w3.org/2000/svg" width="22" height="22" fill="#556B2F" class="bi bi-inboxes" viewBox="0 0 16 16" style="margin-right: 8px; margin-bottom: 4px;"><path d="M4.98 1a.5.5 0 0 0-.39.188L1.54 5H6a.5.5 0 0 1 .5.5 1.5 1.5 0 0 0 3 0A.5.5 0 0 1 10 5h4.46l-3.05-3.812A.5.5 0 0 0 11.02 1H4.98zm9.954 5H10.45a2.5 2.5 0 0 1-4.9 0H1.066l.32 2.562A.5.5 0 0 0 1.884 9h12.234a.5.5 0 0 0 .496-.438L14.933 6zM3.809.563A1.5 1.5 0 0 1 4.981 0h6.038a1.5 1.5 0 0 1 1.172.563l3.7 4.625a.5.5 0 0 1 .109.273l.94 7.514A1.5 1.5 0 0 1 15.446 14H.554a1.5 1.5 0 0 1-1.493-1.025l.94-7.514a.5.5 0 0 1 .108-.273l3.7-4.625zM2.013 10l-.415 3.315A.5.5 0 0 0 2.094 14h11.812a.5.5 0 0 0 .495-.685L13.987 10h-2.53a3.5 3.5 0 0 1-6.914 0H2.013z"/></svg>'''
        svg_bag = '''<svg xmlns="http://www.w3.org/2000/svg" width="22" height="22" fill="#556B2F" class="bi bi-bag" viewBox="0 0 16 16" style="margin-right: 8px; margin-bottom: 4px;"><path d="M8 1a2.5 2.5 0 0 1 2.5 2.5V4h-5v-.5A2.5 2.5 0 0 1 8 1zm3.5 3v-.5a3.5 3.5 0 1 0-7 0V4H1v10a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V4h-3.5zM2 5h12v9a1 1 0 0 1-1 1H3a1 1 0 0 1-1-1V5z"/></svg>'''
        svg_chart = '''<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" fill="#556B2F" class="bi bi-bar-chart" viewBox="0 0 16 16" style="margin-right: 8px; margin-bottom: 4px;"><path d="M4 11H2v3h2v-3zm5-4H7v7h2V7zm5-5v12h-2V2h2zm-2-1a1 1 0 0 0-1 1v12a1 1 0 0 0 1 1h2a1 1 0 0 0 1-1V2a1 1 0 0 0-1-1h-2zM6 7a1 1 0 0 1 1-1h2a1 1 0 0 1 1 1v7a1 1 0 0 1-1 1H7a1 1 0 0 1-1-1V7zm-5 4a1 1 0 0 1 1-1h2a1 1 0 0 1 1 1v3a1 1 0 0 1-1 1H2a1 1 0 0 1-1-1v-3z"/></svg>'''

        with tab_bandejas:
            st.write("") 
            with st.container(border=True):
                st.markdown(f"<h4 style='display:flex; align-items:center;'>{svg_tray} Control por Repartidor</h4>", unsafe_allow_html=True)
                c_sel, _ = st.columns([1, 1])
                
                v_ban = c_sel.selectbox("Repartidor", lista_repartidores_todos, index=idx_defecto, placeholder="Seleccione repartidor...", label_visibility="collapsed")
                
                if v_ban:
                    id_vban = dict_vend[v_ban]
                    data_ban = obtener_bandejas(f_ins, id_vban)
                    
                    st.divider()
                    col_b1, col_b2, col_b3, col_b4 = st.columns(4)
                    ant_b = col_b1.number_input("Saldo Inicial", value=val_gui(data_ban['ant']), disabled=True)
                    sal_b = col_b2.number_input("Egreso", value=val_gui(data_ban['sal'] if data_ban.get('existe') else None), placeholder="0", min_value=0, disabled=not permiso_editar)
                    ret_b = col_b3.number_input("Retorno", value=val_gui(data_ban['ret'] if data_ban.get('existe') else None), placeholder="0", min_value=0, disabled=not permiso_editar)
                    
                    fin_b = (data_ban['ant'] or 0) - (sal_b or 0) + (ret_b or 0)
                    delta_hoy = (ret_b or 0) - (sal_b or 0)
                    col_b4.metric("Saldo Final", f"{fin_b} un.", delta=f"{delta_hoy} flujo hoy" if sal_b or ret_b else None, delta_color="normal" if delta_hoy >= 0 else "inverse")
                    
                    st.write("")
                    # GRILLA IZQUIERDA [1, 4]
                    c_btn, _ = st.columns([1, 4])
                    if permiso_editar and c_btn.button("Guardar", type="primary", use_container_width=True):
                        guardar_bandejas(f_ins, id_vban, data_ban['ant'], sal_b, ret_b)
                        st.success(f"Guardado.")
                        time.sleep(0.5)
                        st.rerun()
            
            st.write("")
            st.markdown(f"<h5 style='display:flex; align-items:center;'>{svg_chart} Resumen Global</h5>", unsafe_allow_html=True)
            df_resumen_bandejas = obtener_resumen_bandejas_especial(f_ins)
            if not df_resumen_bandejas.empty:
                df_resumen_bandejas = df_resumen_bandejas[~df_resumen_bandejas['Repartidor'].isin(repartidores_corriente)]
                max_bandejas = max(50, int(df_resumen_bandejas['Saldo Final Bandejas'].max()) + 10)
                st.dataframe(
                    df_resumen_bandejas, 
                    use_container_width=True, 
                    hide_index=True,
                    column_config={
                        "Repartidor": st.column_config.TextColumn("Repartidor", width="medium"),
                        "Saldo Final Bandejas": st.column_config.ProgressColumn("Volumen Retenido", format="%d un.", min_value=0, max_value=max_bandejas)
                    }
                )
            else:
                st.info("No hay datos.")

        with tab_bolsas:
            st.write("")
            with st.container(border=True):
                st.markdown(f"<h4 style='display:flex; align-items:center;'>{svg_bag} Stock de Empaques</h4>", unsafe_allow_html=True)
                st.write("")
                
                df_bolsas = obtener_bolsas_manual(f_ins)
                
                cols_bol = {
                    "id": None, "id_cb": None, "factor": None, "gasto_cajas": None, "stock_inicial_bolsas": None,
                    "nombre": st.column_config.TextColumn("Producto", disabled=True, width="medium"), 
                    "stock_inicial_cajas": st.column_config.NumberColumn("Ini (Cajas)", disabled=True, format="%.2f", width="small"), 
                    "ingreso_cajas": st.column_config.NumberColumn("Ingreso (Cajas)", step=0.01, required=True, format="%.2f", width="small"), 
                    "produccion_hoy_unidades": st.column_config.NumberColumn("Prod. (Unidades)", step=1, required=True, width="medium"), 
                    "stock_cajas_final": st.column_config.NumberColumn("Fin (Cajas)", disabled=True, format="%.2f", width="small"), 
                    "stock_bolsas_final": st.column_config.NumberColumn("Fin (Bolsas)", disabled=True, format="%d", width="small")
                }
                orden = ["id", "id_cb", "factor", "gasto_cajas", "nombre", "stock_inicial_cajas", "stock_inicial_bolsas", "ingreso_cajas", "produccion_hoy_unidades", "stock_cajas_final", "stock_bolsas_final"]
                
                if not df_bolsas.empty:
                    df_visual = df_bolsas[orden]
                    def estilizar_bolsas(row):
                        return ['background-color: #FFFFFF; font-weight: 600; color: #1F2937;' if c in ['ingreso_cajas', 'produccion_hoy_unidades'] else 'background-color: #F8F9FA; color: #6C757D;' for c in row.index]
                    df_styled = df_visual.style.apply(estilizar_bolsas, axis=1)
                else:
                    df_styled = pd.DataFrame()

                df_ed_bol = st.data_editor(df_styled, column_config=cols_bol, use_container_width=True, hide_index=True, key="ed_bolsas", disabled=not permiso_editar, height=350)
                
                st.write("")
                c_btn_b, _ = st.columns([1, 4])
                if permiso_editar and c_btn_b.button("Actualizar Stock", type="primary", use_container_width=True):
                    gasto = df_ed_bol['produccion_hoy_unidades'] / df_ed_bol['factor']
                    df_ed_bol['stock_cajas_final'] = df_ed_bol['stock_inicial_cajas'] + df_ed_bol['ingreso_cajas'] - gasto
                    guardar_bolsas_manual(f_ins, df_ed_bol)
                    st.success("Actualizado.")
                    time.sleep(0.5)
                    st.rerun()

    elif seleccion == "Producción":
        st.title("Control de Producción")
        c1, _ = st.columns([1, 4])
        f_st = c1.date_input("Fecha", date.today(), format="DD/MM/YYYY")
        
        svg_prod = '''<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" fill="#556B2F" class="bi bi-clipboard" viewBox="0 0 16 16" style="margin-right: 8px; margin-bottom: 4px;"><path d="M4 1.5H3a2 2 0 0 0-2 2V14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V3.5a2 2 0 0 0-2-2h-1v1h1a1 1 0 0 1 1 1V14a1 1 0 0 1-1 1H3a1 1 0 0 1-1-1V3.5a1 1 0 0 1 1-1h1v-1z"/><path d="M9.5 1a.5.5 0 0 1 .5.5v1a.5.5 0 0 1-.5.5h-3a.5.5 0 0 1-.5-.5v-1a.5.5 0 0 1 .5-.5h3zm-3-1A1.5 1.5 0 0 0 5 1.5v1A1.5 1.5 0 0 0 6.5 4h3A1.5 1.5 0 0 0 11 2.5v-1A1.5 1.5 0 0 0 9.5 0h-3z"/></svg>'''
        
        st.write("")
        st.markdown(f"<h4 style='display:flex; align-items:center;'>{svg_prod} Fabricación</h4>", unsafe_allow_html=True)
        
        df_st = obtener_datos_stock(f_st)
        
        orden_cultural_prod = [
            "Lengua", "Lengua 6", "Lengua XL (25)", "Lengua XXL (30)", 
            "Lengua XXXL (35)", "Lengua XXXXL (40)", "Frica", "Frica XL", 
            "Molde", "Molde XL", "Pizza Individual", "Pizza Familiar", 
            "Hallulla", "Pan Rallado", "Tapadito"
        ]
        
        if not df_st.empty:
            df_st['nombre_cat'] = pd.Categorical(df_st['nombre'].str.strip(), categories=orden_cultural_prod, ordered=True)
            df_st = df_st.sort_values('nombre_cat').drop(columns=['nombre_cat'])
        
        cols_prod = {
            "id": None, 
            "bolsas_por_saco": None, 
            "nombre": st.column_config.TextColumn("🔒 Producto", disabled=True, width="medium"), 
            "stock_inicial": st.column_config.NumberColumn("🔒 Stock Inicial", disabled=True, format="%d", width="small"), 
            "fabricacion": st.column_config.NumberColumn("Fabricación", required=True, width="small"), 
            "salida_calculada": st.column_config.NumberColumn("🔒 Despacho", disabled=True, width="small"), 
            "stock_final": st.column_config.NumberColumn("🔒 Stock Final", disabled=True, width="small"), 
            "produccion_dia_siguiente": st.column_config.NumberColumn("🔒 Meta Mañana", disabled=True, format="%d", width="small"), 
            "bolsas_necesarias": st.column_config.NumberColumn("🔒 Bolsas Faltantes", disabled=True, format="%d", width="small"), 
            "cant_sacos": st.column_config.NumberColumn("🔒 Harina (Sacos)", disabled=True, format="%.2f", width="small")
        }
        orden = ["id", "nombre", "stock_inicial", "fabricacion", "salida_calculada", "stock_final", "produccion_dia_siguiente", "bolsas_necesarias", "cant_sacos"]
        
        with st.container(border=True):
            altura_dinamica_prod = (len(df_st) * 36) + 42 if not df_st.empty else 400
            
            def estilizar_produccion(row):
                return ['background-color: #FFFFFF; font-weight: 600; color: #1F2937;' if c == 'fabricacion' else 'background-color: #F8F9FA; color: #6C757D;' for c in row.index]

            if not df_st.empty:
                df_visual = df_st[[c for c in orden if c in df_st.columns]]
                df_styled = df_visual.style.apply(estilizar_produccion, axis=1)
            else:
                df_styled = pd.DataFrame()

            df_ed = st.data_editor(df_styled, column_config=cols_prod, use_container_width=True, hide_index=True, key="st_editor", height=altura_dinamica_prod, disabled=not permiso_editar)
            
            st.write("")
            c_btn, _ = st.columns([1, 4])
            if permiso_editar and c_btn.button("Guardar", type="primary", use_container_width=True):
                for i, r in df_ed.iterrows():
                    fab_val = int(r['fabricacion'])
                    st_fin = r['stock_inicial'] + fab_val - r['salida_calculada']
                    nec = max(0, r['produccion_dia_siguiente'] - st_fin)
                    registrar_produccion(f_st, r['id'], r['stock_inicial'], fab_val, st_fin, nec)
                st.success("Guardado.")
                time.sleep(0.5)
                st.rerun()

    elif seleccion == "Despacho":
        st.title("Despachos")
        
        df_prod = obtener_productos_activos()
        
        c1, c2 = st.columns([1, 2])
        f_bo = c1.date_input("Fecha", date.today(), format="DD/MM/YYYY")
        
        v_bo = c2.selectbox("Repartidor", lista_repartidores_todos, index=idx_defecto, placeholder="Seleccione un repartidor...", disabled=es_repartidor, key="sel_desp_rep")
        
        svg_truck = '''<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" fill="#556B2F" class="bi bi-truck" viewBox="0 0 16 16" style="margin-right: 8px; margin-bottom: 4px;"><path d="M0 3.5A1.5 1.5 0 0 1 1.5 2h9A1.5 1.5 0 0 1 12 3.5V5h1.02a1.5 1.5 0 0 1 1.17.563l1.481 1.85a1.5 1.5 0 0 1 .329.938V10.5a1.5 1.5 0 0 1-1.5 1.5H14a2 2 0 1 1-4 0H5a2 2 0 1 1-3.998-.085A1.5 1.5 0 0 1 0 10.5v-7zm1.294 7.456A1.999 1.999 0 0 1 4.732 11h5.536a2.01 2.01 0 0 1 .732-.732V3.5a.5.5 0 0 0-.5-.5h-9a.5.5 0 0 0-.5.5v7a.5.5 0 0 0 .294.456zM12 10a2 2 0 0 1 1.732 1h.768a.5.5 0 0 0 .5-.5V8.35a.5.5 0 0 0-.11-.312l-1.48-1.85A.5.5 0 0 0 13.02 6H12v4zm-9 1a1 1 0 1 0 0 2 1 1 0 0 0 0-2zm9 0a1 1 0 1 0 0 2 1 1 0 0 0 0-2z"/></svg>'''
        svg_box = '''<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" fill="#556B2F" class="bi bi-box-seam" viewBox="0 0 16 16" style="margin-right: 8px; margin-bottom: 4px;"><path d="M8.186 1.113a.5.5 0 0 0-.372 0L1.846 3.5l2.404.961L10.404 2zm3.564 1.426L5.596 5 8 5.961 14.154 3.5zm3.25 1.7-6.5 2.6v7.922l6.5-2.6V4.24zM7.5 14.762V6.84L1 4.239v7.923zM7.443.184a1.5 1.5 0 0 1 1.114 0l7.129 2.852A.5.5 0 0 1 16 3.5v8.662a1 1 0 0 1-.629.958l-7.185 2.872a1.5 1.5 0 0 1-1.114 0l-7.185-2.872A1 1 0 0 1 0 12.162V3.5a.5.5 0 0 1 .314-.464z"/></svg>'''
        
        st.write("")
        
        if v_bo:
            id_vb = dict_vend[v_bo]
            
            # --- MURO VISUAL: La Asignación solo existe si NO eres repartidor ---
            if not es_repartidor:
                st.markdown(f"<h4 style='display:flex; align-items:center;'>{svg_truck} Asignación</h4>", unsafe_allow_html=True)
                
                estado_carga = not permiso_editar
                
                with st.form("add_carga", clear_on_submit=True, border=False):
                    cf1, cf2, cf3 = st.columns([2, 1, 1], vertical_alignment="bottom")
                    p_b = cf1.selectbox("Producto", df_prod['nombre'], disabled=estado_carga)
                    q_b = cf2.number_input("Cantidad", min_value=1, value=None, placeholder="0", disabled=estado_carga)
                    
                    if cf3.form_submit_button("Cargar", disabled=estado_carga, type="primary", use_container_width=True):
                        if q_b:
                            id_pb = int(df_prod[df_prod['nombre']==p_b]['id'].values[0])
                            registrar_carga(f_bo, id_vb, id_pb, q_b)
                            st.rerun()
                
                st.divider()
            
            # --- EL MANIFIESTO DE CARGA: Visible para todos ---
            st.markdown(f"<h5 style='display:flex; align-items:center;'>{svg_box} En Vehículo</h5>", unsafe_allow_html=True)
            
            df_c = obtener_despacho_vehiculo(f_bo, id_vb)
            
            if not df_c.empty:
                cols = {
                    "id": None, 
                    "nombre": st.column_config.TextColumn("🔒 Producto", disabled=True, width="medium"), 
                    "saldo_anterior": st.column_config.NumberColumn("🔒 Ant.", disabled=True, width="small"), 
                    "carga": st.column_config.NumberColumn("Carga", min_value=0, required=True, disabled=not permiso_editar, width="small")
                }
                
                def estilizar_despacho(row):
                    return ['background-color: #FFFFFF; font-weight: 600; color: #1F2937;' if c == 'carga' else 'background-color: #F8F9FA; color: #6C757D;' for c in row.index]
                
                df_styled_c = df_c.style.apply(estilizar_despacho, axis=1)
                
                df_c_ed = st.data_editor(df_styled_c, column_config=cols, hide_index=True, use_container_width=True, key="edit_despacho", disabled=not permiso_editar)
                
                st.write("")
                c_btn, _ = st.columns([1, 4])
                if permiso_editar and not es_repartidor and c_btn.button("Guardar", type="primary", use_container_width=True):
                    actualizar_carga_masiva(df_c_ed)
                    time.sleep(0.5)
                    st.rerun()
            else: 
                st.info("Vehículo vacío.")

    elif seleccion == "Cobranza":
        st.title("Cobranza")
        
        svg_inv = '''<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" fill="#556B2F" class="bi bi-box-seam" viewBox="0 0 16 16" style="margin-right: 8px; margin-bottom: 2px;"><path d="M8.186 1.113a.5.5 0 0 0-.372 0L1.846 3.5l2.404.961L10.404 2zm3.564 1.426L5.596 5 8 5.961 14.154 3.5zm3.25 1.7-6.5 2.6v7.922l6.5-2.6V4.24zM7.5 14.762V6.84L1 4.239v7.923zM7.443.184a1.5 1.5 0 0 1 1.114 0l7.129 2.852A.5.5 0 0 1 16 3.5v8.662a1 1 0 0 1-.629.958l-7.185 2.872a1.5 1.5 0 0 1-1.114 0l-7.185-2.872A1 1 0 0 1 0 12.162V3.5a.5.5 0 0 1 .314-.464z"/></svg>'''
        svg_ren = '''<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" fill="#556B2F" class="bi bi-wallet2" viewBox="0 0 16 16" style="margin-right: 8px; margin-bottom: 2px;"><path d="M12.136.326A1.5 1.5 0 0 1 14 1.78V3h.5A1.5 1.5 0 0 1 16 4.5v9a1.5 1.5 0 0 1-1.5 1.5h-13A1.5 1.5 0 0 1 0 13.5v-9a1.5 1.5 0 0 1 1.432-1.499L12.136.326zM5.562 3H13V1.78a.5.5 0 0 0-.621-.484L5.562 3zM1.5 4a.5.5 0 0 0-.5.5v9a.5.5 0 0 0 .5.5h13a.5.5 0 0 0 .5-.5v-9a.5.5 0 0 0-.5-.5h-13z"/></svg>'''
        svg_rep = '''<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" fill="#556B2F" class="bi bi-bar-chart" viewBox="0 0 16 16" style="margin-right: 8px; margin-bottom: 2px;"><path d="M4 11H2v3h2v-3zm5-4H7v7h2V7zm5-5v12h-2V2h2zm-2-1a1 1 0 0 0-1 1v12a1 1 0 0 0 1 1h2a1 1 0 0 0 1-1V2a1 1 0 0 0-1-1h-2zM6 7a1 1 0 0 1 1-1h2a1 1 0 0 1 1 1v7a1 1 0 0 1-1 1H7a1 1 0 0 1-1-1V7zm-5 4a1 1 0 0 1 1-1h2a1 1 0 0 1 1 1v3a1 1 0 0 1-1 1H2a1 1 0 0 1-1-1v-3z"/></svg>'''
        svg_pay = '''<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="#556B2F" class="bi bi-cash-coin" viewBox="0 0 16 16" style="margin-right: 6px; margin-bottom: 2px;"><path fill-rule="evenodd" d="M11 15a4 4 0 1 0 0-8 4 4 0 0 0 0 8zm5-4a5 5 0 1 1-10 0 5 5 0 0 1 10 0z"/><path d="M9.438 11.944c.047.596.518 1.06 1.363 1.116v.44h.375v-.443c.875-.061 1.386-.529 1.386-1.207 0-.618-.39-.936-1.09-1.1l-.296-.07v-1.2c.376.043.614.248.671.532h.658c-.047-.575-.54-1.024-1.329-1.073V8.5h-.375v.45c-.747.073-1.255.522-1.255 1.158 0 .562.378.92 1.007 1.066l.248.061v1.272c-.384-.058-.639-.27-.696-.563h-.668zm1.36-1.354c-.369-.085-.569-.26-.569-.522 0-.294.216-.514.572-.578v1.1h-.003zm.432.746c.449.104.655.272.655.569 0 .339-.257.571-.709.614v-1.195l.054.012z"/><path d="M1 0a1 1 0 0 0-1 1v8a1 1 0 0 0 1 1h4.083c.058-.344.145-.678.258-1H3a2 2 0 0 0-2-2V3a2 2 0 0 0 2-2h10a2 2 0 0 0 2 2v3.528c.38.34.717.728 1 1.154V1a1 1 0 0 0-1-1H1z"/><path d="M9.998 5.083 10 5a2 2 0 1 0-3.132 1.65 5.982 5.982 0 0 1 3.13-1.567z"/></svg>'''

        c1, c2, c3 = st.columns([2, 3, 1], vertical_alignment="bottom")
        f_of = c1.date_input("Fecha", date.today()-timedelta(days=1), format="DD/MM/YYYY")
        
        # --- FILTRO DUAL ---
        if rol == "despacho_especial":
            lista_rep_cob = [r for r in lista_repartidores_todos if r in repartidores_corriente or r == "Kilaco Venta"]
        else:
            lista_rep_cob = lista_repartidores_todos
            
        idx_def_cob = lista_rep_cob.index(nombre_vendedor_actual) if es_repartidor and nombre_vendedor_actual in lista_rep_cob else None
        
        v_of = c2.selectbox("Repartidor", lista_rep_cob, index=idx_def_cob, placeholder="Seleccione un repartidor...", disabled=es_repartidor, key="sel_cob_rep")
        
        if not es_repartidor:
            texto_boton = "Pendientes ▲" if st.session_state.get("mostrar_pendientes", False) else "Pendientes ▼"
            if c3.button(texto_boton, use_container_width=True):
                st.session_state.mostrar_pendientes = not st.session_state.get("mostrar_pendientes", False)
                st.rerun()

            if st.session_state.get("mostrar_pendientes", False):
                df_pendientes = obtener_cobranzas_pendientes_especial()
                if not df_pendientes.empty:
                    df_pendientes = df_pendientes[df_pendientes['Repartidor'] != 'Kilaco Venta']
                    st.dataframe(df_pendientes, hide_index=True, use_container_width=True)
                else:
                    st.success("Al día.")
                st.write("") 
        
        # --- MUTACIÓN ARQUITECTÓNICA DE PESTAÑAS ---
        if rol == "despacho_especial":
            tabs_cob = st.tabs(["Inventario"])
            tab_inv = tabs_cob[0]; tab_fin = None; tab_rep = None
        else:
            tabs_cob = st.tabs(["Inventario", "Rendición", "Reporte"])
            tab_inv = tabs_cob[0]; tab_fin = tabs_cob[1]; tab_rep = tabs_cob[2]
        
        # --- PESTAÑA 1: INVENTARIO ---
        with tab_inv:
            st.write("")
            with st.container(border=True):
                st.markdown(f"<h5 style='display:flex; align-items:center;'>{svg_inv} Retorno Físico</h5>", unsafe_allow_html=True)
                
                if v_of:
                    id_vo = dict_vend[v_of]
                    df_inv = obtener_planilla(f_of, id_vo)
                    total_venta = 0 
                    
                    orden_cultural_inv = [
                        "Lengua", "Lengua 6", "Frica", "Lengua XL (25)", "Pizza Individual", 
                        "Pizza Familiar", "Hallulla", "Molde", "Frica XL", "Tapadito", 
                        "Molde XL", "Pan Rallado", "Lengua XXL (30)", "Lengua XXXL (35)", "Lengua XXXXL (40)"
                    ]
                    
                    if not df_inv.empty:
                        df_inv['nombre_cat'] = pd.Categorical(df_inv['nombre'].str.strip(), categories=orden_cultural_inv, ordered=True)
                        df_inv = df_inv.sort_values('nombre_cat').drop(columns=['nombre_cat'])
                        
                        df_inv['id_producto_hidden'] = df_inv['id_producto']
                        df_d = df_inv.rename(columns={'devolucion_muestra':'devolucion','saldo_actual':'saldo_final'})
                        
                        cols = {"id": None, "id_producto": None, "id_producto_hidden": None, "precio_estandar": None, "disp": None, "orden_visual": None, 
                                "nombre": st.column_config.TextColumn("Producto", disabled=True, width="medium"), 
                                "saldo_anterior": st.column_config.NumberColumn("🔒 Ant.", disabled=True, width="small"), 
                                "carga": st.column_config.NumberColumn("🔒 Carga", disabled=True, width="small"), 
                                "devolucion": st.column_config.NumberColumn("Devolución", required=True, min_value=0, width="small"), 
                                "saldo_final": st.column_config.NumberColumn("Saldo Fin.", required=True, min_value=0, width="small"), 
                                "venta": st.column_config.NumberColumn("🔒 Venta", disabled=True, width="small"), 
                                "total": st.column_config.NumberColumn("🔒 Total", disabled=True, format="$%d", width="small")}
                        
                        altura_dinamica_inv = (len(df_d) * 36) + 42
                        
                        def estilizar_inventario(row):
                            return ['background-color: #FFFFFF; font-weight: 600; color: #1F2937;' if c in ['devolucion', 'saldo_final'] else 'background-color: #F8F9FA; color: #6C757D;' for c in row.index]
                        
                        df_styled_inv = df_d.style.apply(estilizar_inventario, axis=1)

                        df_ed_inv = st.data_editor(df_styled_inv, column_config=cols, hide_index=True, use_container_width=True, key="inv_ed", height=altura_dinamica_inv, disabled=not permiso_editar)
                        df_ed_inv['v_real'] = (df_ed_inv['saldo_anterior']+df_ed_inv['carga']-df_ed_inv['devolucion']-df_ed_inv['saldo_final']).clip(lower=0)
                        total_venta = (df_ed_inv['v_real']*df_ed_inv['precio_estandar']).sum()
                        
                        st.write("")
                        ci, cs = st.columns([3, 1], vertical_alignment="center")
                        ci.metric("Venta Bruta", fmt_clp(total_venta))
                        if permiso_editar and cs.button("Guardar", type="primary", use_container_width=True, key="b_inv"): 
                            guardar_oficina(df_ed_inv, f_of, id_vo); st.success("Guardado"); time.sleep(0.5); st.rerun()
              
        
        # --- PESTAÑA 2: RENDICIÓN ---
        if tab_fin:
            with tab_fin:
                st.write("")
                if v_of:
                    id_vo = dict_vend[v_of]
                    es_repartidor_de_corriente = v_of in repartidores_corriente
                    es_kilaco_venta = v_of == "Kilaco Venta"
                    es_solo_fisico = es_repartidor_de_corriente or es_kilaco_venta
                    
                    if es_solo_fisico:
                        st.info("No requiere rendición de caja.")
                    else:
                        df_inv_fin = obtener_planilla(f_of, id_vo)
                        total_venta_fin = 0
                        if not df_inv_fin.empty: total_venta_fin = df_inv_fin['total'].sum()

                        fin = get_finanzas(f_of, id_vo)
                        with st.container(border=True):
                            st.markdown(f"<h5 style='display:flex; align-items:center;'>{svg_ren} Flujo de Caja</h5>", unsafe_allow_html=True)
                            st.write("")
                            
                            ph_step1 = st.container()
                            
                            st.divider()
                            st.markdown("**2. Gastos y Movimientos**")
                            
                            c_gr1, c_gr2, c_gr3 = st.columns(3)
                            with c_gr1:
                                st.caption("Cobro de Créditos")
                                df_cc = st.data_editor(pd.DataFrame([{"Detalle": fin.get('cc_det', "Varios"), "Monto": int(fin.get('cc', 0))}]), num_rows="dynamic", key="grid_cc", hide_index=True, column_config={"Monto": st.column_config.NumberColumn(format="$%d", required=True)}, disabled=not permiso_editar, use_container_width=True)
                                t_cc = df_cc['Monto'].sum(); txt_cc = ", ".join(df_cc['Detalle'].astype(str).tolist())
                            with c_gr2:
                                st.caption("Créditos (Fiado)")
                                df_co = st.data_editor(pd.DataFrame([{"Detalle": "Varios", "Monto": int(fin.get('co', 0))}]), num_rows="dynamic", key="grid_co", hide_index=True, column_config={"Monto": st.column_config.NumberColumn(format="$%d", required=True)}, disabled=not permiso_editar, use_container_width=True)
                                t_co = df_co['Monto'].sum()
                            with c_gr3:
                                st.caption("Otros Gastos")
                                df_om = st.data_editor(pd.DataFrame([{"Detalle": fin.get('od', "Varios"), "Monto": int(fin.get('om', 0))}]), num_rows="dynamic", key="grid_om", hide_index=True, column_config={"Monto": st.column_config.NumberColumn(format="$%d", required=True)}, disabled=not permiso_editar, use_container_width=True)
                                t_om = df_om['Monto'].sum(); txt_om = ", ".join(df_om['Detalle'].astype(str).tolist())
                            
                            st.write("")
                            g1, g2, g3 = st.columns(3)
                            bn = safe_int(g1.number_input("Bencina", value=val_gui(fin.get('bn',0)), step=1000, placeholder="0", disabled=not permiso_editar))
                            su = safe_int(g2.number_input("Sueldo", value=val_gui(fin.get('su',0)), step=1000, placeholder="0", disabled=not permiso_editar))
                            ds = safe_int(g3.number_input("Descuentos", value=val_gui(fin.get('ds',0)), step=500, placeholder="0", disabled=not permiso_editar))
                            
                            tot_ing = total_venta_fin + t_cc
                            tot_gas = t_co + t_om + bn + su + ds
                            deuda_neta = tot_ing - tot_gas
                            
                            with ph_step1:
                                st.markdown("**1. Dinero Exigible**")
                                ci1, ci2, ci3 = st.columns(3)
                                ci1.metric("Venta Bruta (Física)", fmt_clp(total_venta_fin), "Retorno Inventario", delta_color="off")
                                ci2.metric("Cobro de Créditos", fmt_clp(t_cc), "Fiados recuperados", delta_color="off")
                                ci3.metric("Total a Rendir", fmt_clp(tot_ing), "Base de cálculo", delta_color="normal")
                            
                            st.divider()
                            st.markdown(f"<h6 style='display:flex; align-items:center; color:#556B2F;'>{svg_pay} Cierre y Pagos</h6>", unsafe_allow_html=True)
                            st.write("")
                            
                            st.markdown(f"<h4 style='color:#2C3E50;'>Total: {fmt_clp(deuda_neta)}</h4>", unsafe_allow_html=True)
                            st.write("")
                            
                            r1, r2, r3 = st.columns(3)
                            ef = safe_int(r1.number_input("Efectivo", value=val_gui(fin.get('ef',0)), step=1000, placeholder="0", disabled=not permiso_editar))
                            tr = safe_int(r2.number_input("Transferencias", value=val_gui(fin.get('tr',0)), step=1000, placeholder="0", disabled=not permiso_editar))
                            pc = safe_int(r3.number_input("Centralizado", value=val_gui(fin.get('pc',0)), step=1000, placeholder="0", disabled=not permiso_editar))
                            
                            saldo = deuda_neta - (ef + tr + pc)
                            st.write("")
                            
                            c_sem, c_btn = st.columns([3, 1], vertical_alignment="center")
                            estilo_base = "margin: 0; height: 42px; display: flex; align-items: center; justify-content: center; border-radius: 6px; font-weight: 600; font-size: 15px;"
                            if saldo == 0: c_sem.markdown(f"<div style='{estilo_base} background-color:#F0FDF4; color:#166534; border: 1px solid #BBF7D0;'>Cuadratura Exacta</div>", unsafe_allow_html=True)
                            elif saldo > 0: c_sem.markdown(f"<div style='{estilo_base} background-color:#FEF2F2; color:#991B1B; border: 1px solid #FECACA;'>Faltan {fmt_clp(saldo)}</div>", unsafe_allow_html=True)
                            else: c_sem.markdown(f"<div style='{estilo_base} background-color:#FFFBEB; color:#92400E; border: 1px solid #FDE68A;'>Sobran {fmt_clp(abs(saldo))}</div>", unsafe_allow_html=True)
                            
                            if permiso_editar and c_btn.button("Guardar", type="primary", use_container_width=True, key="b_ren"):
                                save_finanzas(f_of, id_vo, {"cc":t_cc, "cc_det":txt_cc, "co":t_co, "ds":ds, "bn":bn, "su":su, "om":t_om, "od":txt_om, "ef":ef, "tr":tr, "pc":pc})
                                st.success("Guardado"); time.sleep(0.5); st.rerun()

        # --- PESTAÑA 3: REPORTE ---
        if tab_rep:
            with tab_rep:
                st.write("")
                st.markdown(f"<h5 style='display:flex; align-items:center;'>{svg_rep} Estado General</h5>", unsafe_allow_html=True)
                
                col_r1, col_r2, col_r3, col_r4 = st.columns([2, 2, 3, 2], vertical_alignment="bottom")
                fi_e = col_r1.date_input("Desde", date.today() - timedelta(days=7), key="r_ini_e", format="DD/MM/YYYY")
                ff_e = col_r2.date_input("Hasta", date.today(), key="r_fin_e", format="DD/MM/YYYY")
                
                # --- SELECTOR INTELIGENTE DE AISLAMIENTO ---
                if not es_repartidor:
                    filtro_rep_e = col_r3.selectbox("Filtrar Repartidor", ["Todos"] + lista_repartidores_esp, index=0, key="filt_rep_rep_e")
                else:
                    filtro_rep_e = col_r3.selectbox("Repartidor", [nombre_vendedor_actual], disabled=True, key="filt_rep_rep_e")
                
                if col_r4.button("Generar Reporte", type="primary", use_container_width=True, key="btn_rep_e"):
                    df_res = obtener_resumen_global(fi_e, ff_e)
                    
                    if not df_res.empty:
                        # --- MOTOR DE FILTRADO ---
                        if es_repartidor:
                            df_res = df_res[df_res['Vendedor'] == nombre_vendedor_actual]
                        elif filtro_rep_e != "Todos":
                            df_res = df_res[df_res['Vendedor'] == filtro_rep_e]
                            
                        if not df_res.empty:
                            # ORDEN CRONOLÓGICO: Más reciente primero
                            if 'fecha' in df_res.columns.str.lower():
                                col_fecha = 'Fecha' if 'Fecha' in df_res.columns else 'fecha'
                                df_res[col_fecha] = pd.to_datetime(df_res[col_fecha], dayfirst=True, errors='ignore')
                                df_res = df_res.sort_values(by=[col_fecha, 'Vendedor'], ascending=[False, True])
                                df_res[col_fecha] = df_res[col_fecha].dt.strftime('%d/%m/%Y')
                            else:
                                df_res = df_res.sort_values(by='Vendedor', ascending=True)
                                
                            st.dataframe(df_res.style.map(lambda v: 'background-color: #F0FDF4; color: #166534' if v==0 else ('background-color: #FEF2F2; color: #991B1B' if v>0 else 'background-color: #FFFBEB; color: #92400E'), subset=['Saldo']).format({"Total Ingresos": "$ {:,.0f}", "Total Gastos": "$ {:,.0f}", "Deuda Neta": "$ {:,.0f}", "Pagado": "$ {:,.0f}", "Saldo": "$ {:,.0f}"}), hide_index=True, use_container_width=True)
                            st.write("")
                            st.metric("Saldo Pendiente Acumulado", fmt_clp(df_res['Saldo'].sum()))
                        else:
                            st.info("No hay registros para la selección en este rango de fechas.")
                    else:
                        st.info("Sin registros financieros en este período.")
                        
    elif seleccion == "Clientes":
        st.title("Clientes")
        
        svg_people = '''<svg xmlns="http://www.w3.org/2000/svg" width="22" height="22" fill="#556B2F" class="bi bi-people" viewBox="0 0 16 16" style="margin-right: 8px; margin-bottom: 4px;"><path d="M15 14s1 0 1-1-1-4-5-4-5 3-5 4 1 1 1 1h8Zm-7.978-1A.261.261 0 0 1 7 12.996c.001-.264.167-1.03.76-1.72C8.312 10.629 9.282 10 11 10c1.717 0 2.687.63 3.24 1.276.593.69.758 1.457.76 1.72l-.008.002a.274.274 0 0 1-.014.002H7.022ZM11 7a2 2 0 1 0 0-4 2 2 0 0 0 0 4Zm3-2a3 3 0 1 1-6 0 3 3 0 0 1 6 0ZM6.936 9.28a5.88 5.88 0 0 0-1.23-.247A7.35 7.35 0 0 0 5 9c-4 0-5 3-5 4 0 .667.333 1 1 1h4.216A2.238 2.238 0 0 1 5 13c0-1.01.377-2.042 1.09-2.904.243-.294.526-.569.846-.816ZM4.92 10A5.493 5.493 0 0 0 4 13H1c0-.26.164-1.03.76-1.724.545-.636 1.492-1.256 3.16-1.275ZM1.5 5.5a3 3 0 1 1 6 0 3 3 0 0 1-6 0Zm3-2a2 2 0 1 0 0 4 2 2 0 0 0 0-4Z"/></svg>'''
        svg_inbox = '''<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" fill="#556B2F" class="bi bi-inbox" viewBox="0 0 16 16" style="margin-right: 8px; margin-bottom: 2px;"><path d="M4.98 4a.5.5 0 0 0-.39.188L1.54 8H6a.5.5 0 0 1 .5.5 1.5 1.5 0 1 0 3 0A.5.5 0 0 1 10 8h4.46l-3.05-3.812A.5.5 0 0 0 11.02 4H4.98zm9.954 5H10.45a2.5 2.5 0 0 1-4.9 0H1.066l.32 2.562a.5.5 0 0 0 .497.438h12.234a.5.5 0 0 0 .496-.438L14.933 9zM3.809 3.563A1.5 1.5 0 0 1 4.981 3h6.038a1.5 1.5 0 0 1 1.172.563l3.7 4.625a.5.5 0 0 1 .109.273l.94 7.514A1.5 1.5 0 0 1 15.446 15H.554a1.5 1.5 0 0 1-1.493-1.025l.94-7.514a.5.5 0 0 1 .108-.273l3.7-4.625z"/></svg>'''

        es_aprobador = rol in ["admin", "supervisor"]
        tab_dir, tab_sug = st.tabs(["Directorio", "Sugerencias"])
        df_base = obtener_clientes_df()

        with tab_dir:
            st.write("")
            c_head, c_tog = st.columns([3, 1], vertical_alignment="bottom")
            c_head.markdown(f"<h5 style='display:flex; align-items:center; margin-bottom: 0;'>{svg_people} Gestión de Cartera</h5>", unsafe_allow_html=True)

            texto_toggle = "Nuevo Cliente" if permiso_editar else "Sugerir Nuevo"
            modo_crear = c_tog.toggle(texto_toggle)

            if modo_crear:
                st.divider()
                with st.form("new_cli", border=False):
                    n1, n2, n3 = st.columns(3)
                    nom = n1.text_input("Nombre")
                    dire = n2.text_input("Dirección", value="-")
                    tel = n3.text_input("Teléfono", value="-")
                    
                    n4, n5, n6 = st.columns(3)
                    com = n4.selectbox("Comuna", l_comunas)
                    vend = n5.selectbox("Repartidor", lista_repartidores_todos, index=idx_defecto, placeholder="Seleccione un repartidor...", disabled=es_repartidor)
                    tip = n6.selectbox("Tipo", ["Nuevo", "Minorista", "Mayorista"])
                    
                    n_cred = False
                    if permiso_editar:
                        n_cred = st.toggle("Habilitar línea de crédito", value=False)
                    
                    comentario = ""
                    if not permiso_editar:
                        comentario = st.text_input("Justificación (Breve)", placeholder="Ej: Es un almacén nuevo en la ruta...")
                    
                    st.write("")
                    texto_btn = "Guardar Cliente" if permiso_editar else "Enviar Sugerencia"
                    if st.form_submit_button(texto_btn, type="primary"):
                        if nom and vend: 
                            if permiso_editar:
                                crud_cliente("crear", {"nombre":nom, "dir":dire, "com":com, "tel":tel, "id_vend":dict_vend[vend], "tipo":tip, "permite_credito": n_cred})
                                st.success("Creado"); time.sleep(0.5); st.rerun()
                            else:
                                crud_sugerencia("crear", {"tipo":"NUEVO", "nombre":nom, "dir":dire, "com":com, "tel":tel, "id_vend":dict_vend[vend], "tipo_cli":tip, "comentario":comentario})
                                st.success("Sugerencia enviada a revisión."); time.sleep(1.5); st.rerun()
                        else: 
                            st.error("El nombre y el repartidor son obligatorios")
            else:
                st.write("")
                with st.container(border=True):
                    fc1, fc2, fc3 = st.columns(3)
                    filtro_nombre = fc1.text_input("Buscar Cliente", placeholder="Buscar por nombre...")
                    
                    if es_repartidor:
                        df_filtrada_rep = df_base[df_base['Repartidor'] == nombre_vendedor_actual]
                        comunas_posibles = sorted(df_filtrada_rep['comuna'].dropna().unique().tolist())
                        filtro_comuna = fc2.multiselect("Comuna", comunas_posibles, placeholder="Filtrar comuna...")
                        filtro_rep = []
                    else:
                        sel_com = st.session_state.get('filt_com_esp', [])
                        sel_rep = st.session_state.get('filt_rep_esp', [])
                        
                        if sel_rep:
                            com_pos = df_base[df_base['Repartidor'].isin(sel_rep)]['comuna'].dropna().unique().tolist()
                            comunas_posibles = sorted(list(set(com_pos + sel_com))) 
                        else:
                            comunas_posibles = sorted(df_base['comuna'].dropna().unique().tolist())
                            
                        if sel_com:
                            rep_pos = df_base[df_base['comuna'].isin(sel_com)]['Repartidor'].dropna().unique().tolist()
                            repartidores_posibles = sorted(list(set(rep_pos + sel_rep)))
                        else:
                            repartidores_posibles = sorted(df_base['Repartidor'].dropna().unique().tolist())
                            
                        filtro_comuna = fc2.multiselect("Comuna", comunas_posibles, key="filt_com_esp", placeholder="Todas las comunas...")
                        filtro_rep = fc3.multiselect("Repartidor", repartidores_posibles, key="filt_rep_esp", placeholder="Todos los repartidores...")

                    df_show = df_base.copy()
                    if es_repartidor: df_show = df_show[df_show['Repartidor'] == nombre_vendedor_actual]
                    elif filtro_rep: df_show = df_show[df_show['Repartidor'].isin(filtro_rep)]
                    
                    if filtro_nombre: df_show = df_show[df_show['nombre'].str.contains(filtro_nombre, case=False, na=False)]
                    if filtro_comuna: df_show = df_show[df_show['comuna'].isin(filtro_comuna)]
                    
                    st.write("")
                    event = st.dataframe(
                        df_show, use_container_width=True, hide_index=True, selection_mode="single-row", on_select="rerun", 
                        column_config={
                            "id": None, "id_vendedor_asignado": None, 
                            "nombre": st.column_config.TextColumn("Cliente", width="medium"), 
                            "direccion": "Dirección", "comuna": "Comuna", "telefono": "Teléfono", 
                            "Repartidor": st.column_config.TextColumn("Repartidor", width="medium"), 
                            "tipo_cliente": "Tipo", 
                            "limite_credito": st.column_config.NumberColumn("Cupo", format="$ %d"),
                            "permite_credito": st.column_config.CheckboxColumn("¿Crédito?")
                        }
                    )
                    
                if len(event.selection.rows) > 0:
                    row = df_show.iloc[event.selection.rows[0]]
                    st.divider()
                    st.markdown(f"##### Editar: {row['nombre']}")
                    with st.form("edit_cli", border=False):
                        e1, e2, e3 = st.columns(3)
                        en = e1.text_input("Nombre", row['nombre'])
                        ed = e2.text_input("Dirección", row['direccion'])
                        et = e3.text_input("Teléfono", row['telefono'])
                        
                        e4, e5, e6 = st.columns(3)
                        ec = e4.selectbox("Comuna", l_comunas, index=l_comunas.index(row['comuna']) if row['comuna'] in l_comunas else 0)
                        ev = e5.selectbox("Repartidor", lista_repartidores_todos, index=lista_repartidores_todos.index(row['Repartidor']) if row['Repartidor'] in lista_repartidores_todos else None, disabled=es_repartidor)
                        
                        tipo_actual = row['tipo_cliente'] if row['tipo_cliente'] in ["Nuevo", "Minorista", "Mayorista"] else "Nuevo"
                        etip = e6.selectbox("Tipo", ["Nuevo", "Minorista", "Mayorista"], index=["Nuevo", "Minorista", "Mayorista"].index(tipo_actual))
                        
                        e_cred = False
                        if permiso_editar:
                            estado_credito_actual = bool(row.get('permite_credito', False))
                            e_cred = st.toggle("Habilitar línea de crédito", value=estado_credito_actual)
                        
                        comentario_ed = ""
                        if not permiso_editar:
                            comentario_ed = st.text_input("Justificación del Cambio", placeholder="Ej: Cambió el teléfono, Cambió de dueño...")
                        
                        st.write("")
                        texto_btn_ed = "Guardar Cambios" if permiso_editar else "Sugerir Cambio"
                        if st.form_submit_button(texto_btn_ed, type="primary"):
                            if ev:
                                if permiso_editar:
                                    crud_cliente("editar", {"id": int(row['id']), "nombre":en, "dir":ed, "com":ec, "tel":et, "id_vend":dict_vend[ev], "tipo":etip, "permite_credito": e_cred})
                                    st.success("Editado"); time.sleep(0.5); st.rerun()
                                else:
                                    crud_sugerencia("crear", {"tipo":"EDICION", "id_ref": int(row['id']), "nombre":en, "dir":ed, "com":ec, "tel":et, "id_vend":dict_vend[ev], "tipo_cli":etip, "comentario":comentario_ed})
                                    st.success("Sugerencia enviada a revisión."); time.sleep(1.5); st.rerun()
                            else:
                                st.error("Debe seleccionar un repartidor válido.")

        with tab_sug:
            st.write("")
            st.markdown(f"<h5 style='display:flex; align-items:center;'>{svg_inbox} Bandeja de Aprobación</h5>", unsafe_allow_html=True)
            
            if es_aprobador:
                df_sug = get_sugerencias(solo_pendientes=True)
                if not df_sug.empty:
                    ev_sug = st.dataframe(
                        df_sug, use_container_width=True, hide_index=True, selection_mode="single-row", on_select="rerun",
                        column_config={"id":None, "id_cliente_ref":None, "id_vendedor":None, "estado":None, "comentario": "Justificación", "tipo_solicitud": "Solicitud", "nombre": "Cliente", "direccion":"Dir", "comuna":"Comuna", "telefono":"Tel", "Repartidor":"Repartidor", "tipo_cliente":"Tipo", "fecha":"Fecha"}
                    )
                    
                    if len(ev_sug.selection.rows) > 0:
                        s_row = df_sug.iloc[ev_sug.selection.rows[0]]
                        st.divider()
                        tipo_txt = "nuevo cliente" if s_row['tipo_solicitud'] == 'NUEVO' else "edición de cliente"
                        st.markdown(f"**Evaluar {tipo_txt}** | Propuesto por: {s_row['Repartidor']}")
                        st.info(f"Justificación: {s_row['comentario']}")
                        
                        ca, cr = st.columns(2)
                        if ca.button("Aprobar e integrar", type="primary", use_container_width=True):
                            crud_sugerencia("APROBADA", s_row.to_dict(), id_sug=int(s_row['id']))
                            st.success("Integrado."); time.sleep(0.5); st.rerun()
                        if cr.button("Rechazar", use_container_width=True):
                            crud_sugerencia("RECHAZADA", id_sug=int(s_row['id']))
                            st.error("Rechazado."); time.sleep(0.5); st.rerun()
                else:
                    st.success("Bandeja limpia.")
            else:
                df_mis_sug = get_sugerencias(solo_pendientes=False, id_vend=mi_id_vendedor)
                if not df_mis_sug.empty:
                    def color_sug(val):
                        if val == 'PENDIENTE': return 'color: #92400E; font-weight: bold' 
                        if val == 'APROBADA': return 'color: #166534; font-weight: bold'  
                        if val == 'RECHAZADA': return 'color: #991B1B; font-weight: bold' 
                        return ''
                    st.dataframe(
                        df_mis_sug.style.map(color_sug, subset=['estado']),
                        use_container_width=True, hide_index=True,
                        column_config={"id":None, "id_cliente_ref":None, "id_vendedor":None, "Repartidor":None, "comentario": "Justificación", "tipo_solicitud": "Solicitud", "nombre": "Cliente", "estado": "Estado", "fecha": "Fecha"}
                    )
                else:
                    st.info("Sin sugerencias enviadas.")

    elif seleccion == "Créditos":
        st.title("Gestión de Créditos")
        
        svg_credit = '''<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" fill="#556B2F" class="bi bi-credit-card" viewBox="0 0 16 16" style="margin-right: 8px; margin-bottom: 2px;"><path d="M0 4a2 2 0 0 1 2-2h12a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H2a2 2 0 0 1-2-2V4zm2-1a1 1 0 0 0-1 1v1h14V4a1 1 0 0 0-1-1H2zm13 4H1v5a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V7z"/><path d="M2 10a1 1 0 0 1 1-1h1a1 1 0 0 1 1 1v1a1 1 0 0 1-1 1H3a1 1 0 0 1-1-1v-1z"/></svg>'''
        svg_list = '''<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" fill="#556B2F" class="bi bi-card-list" viewBox="0 0 16 16" style="margin-right: 8px; margin-bottom: 2px;"><path d="M14.5 3a.5.5 0 0 1 .5.5v9a.5.5 0 0 1-.5.5h-13a.5.5 0 0 1-.5-.5v-9a.5.5 0 0 1 .5-.5h13zm-13-1A1.5 1.5 0 0 0 0 3.5v9A1.5 1.5 0 0 0 1.5 14h13a1.5 1.5 0 0 0 1.5-1.5v-9A1.5 1.5 0 0 0 14.5 2h-13z"/><path d="M5 8a.5.5 0 0 1 .5-.5h7a.5.5 0 0 1 0 1h-7A.5.5 0 0 1 5 8zm0-2.5a.5.5 0 0 1 .5-.5h7a.5.5 0 0 1 0 1h-7a.5.5 0 0 1-.5-.5zm0 5a.5.5 0 0 1 .5-.5h7a.5.5 0 0 1 0 1h-7a.5.5 0 0 1-.5-.5zm-1-5a.5.5 0 1 1-1 0 .5.5 0 0 1 1 0zM4 8a.5.5 0 1 1-1 0 .5.5 0 0 1 1 0zm0 2.5a.5.5 0 1 1-1 0 .5.5 0 0 1 1 0z"/></svg>'''
        svg_hist = '''<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" fill="#556B2F" class="bi bi-clock-history" viewBox="0 0 16 16" style="margin-right: 8px; margin-bottom: 2px;"><path d="M8.515 1.019A7 7 0 0 0 8 1V0a8 8 0 0 1 .589.022zm2.004.45a7 7 0 0 0-.985-.299l.219-.976q.576.129 1.126.342zm1.37.71a7 7 0 0 0-.439-.27l.493-.87a8 8 0 0 1 .979.654l-.615.789a7 7 0 0 0-.418-.302zm1.834 1.79a7 7 0 0 0-.653-.796l.724-.69q.406.429.747.91zm.744 1.352a7 7 0 0 0-.214-.468l.893-.45a8 8 0 0 1 .45 1.088l-.95.313a7 7 0 0 0-.179-.483m.53 2.507a7 7 0 0 0-.1-1.025l.985-.17q.1.58.116 1.17zm-.131 1.538q.05-.254.081-.51l.993.123a8 8 0 0 1-.23 1.155l-.964-.267q.069-.247.12-.501m-.952 2.379q.276-.436.486-.908l.914.405q-.24.54-.555 1.038zm-.964 1.205q.183-.183.35-.378l.758.653a8 8 0 0 1-.401.432z"/><path d="M8 1a7 7 0 1 0 4.95 11.95l.707.707A8.001 8.001 0 1 1 8 0z"/><path d="M7.5 3a.5.5 0 0 1 .5.5v5.21l3.248 1.856a.5.5 0 0 1-.496.868l-3.5-2A.5.5 0 0 1 7 9V3.5a.5.5 0 0 1 .5-.5"/></svg>'''

        es_jefatura = rol in ["admin", "supervisor"]
        
        # --- MUTACIÓN: 3 Pestañas para Admin, 2 para Repartidores ---
        tabs_creditos = st.tabs(["Operaciones", "Movimientos", "Cartera"]) if not es_repartidor else st.tabs(["Movimientos", "Cartera"])
        
        idx_def_cred = lista_repartidores_esp_puro.index(nombre_vendedor_actual) if es_repartidor and nombre_vendedor_actual in lista_repartidores_esp_puro else None

        if not es_repartidor:
            t_op, t_mov, t_car = tabs_creditos[0], tabs_creditos[1], tabs_creditos[2]
        else:
            t_op, t_mov, t_car = None, tabs_creditos[0], tabs_creditos[1]

        if t_op:
            with t_op:
                st.write("")
                df_clientes = obtener_clientes_df()
                
                c_head, c_tog = st.columns([3, 1], vertical_alignment="bottom")
                c_head.markdown(f"<h5 style='display:flex; align-items:center; margin-bottom: 0;'>{svg_credit} Operaciones</h5>", unsafe_allow_html=True)
                
                modo_correccion = False
                if permiso_editar:
                    modo_correccion = c_tog.toggle("Corregir operación")

                with st.container(border=True):
                    if modo_correccion:
                        df_editables = obtener_creditos_editables(es_jefatura)
                        
                        if not df_editables.empty:
                            st.info("Mostrando registros editables según su nivel de credencial (Candado Temporal).")
                            
                            opciones_edicion = {}
                            for _, r in df_editables.iterrows():
                                fecha_str = r['fecha'].strftime("%d/%m/%Y")
                                tipo_visual = "Fiado" if r['tipo_movimiento'] == "CREDITO" else "Abono"
                                eq = f"{fecha_str} | {r['Repartidor']} -> {r['Cliente']} | $ {int(r['monto'])} | {tipo_visual}"
                                opciones_edicion[eq] = r
                                
                            sel_reg = st.selectbox("Seleccione la operación a corregir", list(opciones_edicion.keys()), index=None, placeholder="Buscar registro...")
                            
                            if sel_reg:
                                row_ed = opciones_edicion[sel_reg]
                                st.divider()
                                
                                f_mov = st.date_input("Fecha", row_ed['fecha'], format="DD/MM/YYYY", key="ed_f_mov")
                                
                                c_op1, c_op2 = st.columns(2)
                                idx_rep_ed = lista_repartidores_esp_puro.index(row_ed['Repartidor']) if row_ed['Repartidor'] in lista_repartidores_esp_puro else 0
                                vend_sel = c_op1.selectbox("Repartidor", lista_repartidores_esp_puro, index=idx_rep_ed, key="ed_vend")
                                
                                if not df_clientes.empty:
                                    df_clientes['permite_credito'] = df_clientes['permite_credito'].fillna(False).astype(bool)
                                    df_cli_filt = df_clientes[(df_clientes['Repartidor'] == vend_sel) & (df_clientes['permite_credito'])]
                                else:
                                    df_cli_filt = pd.DataFrame()
                                    
                                if not df_cli_filt.empty:
                                    idx_cli = df_cli_filt['nombre'].tolist().index(row_ed['Cliente']) if row_ed['Cliente'] in df_cli_filt['nombre'].tolist() else 0
                                    cli_sel = c_op2.selectbox("Cliente", df_cli_filt['nombre'], index=idx_cli, key="ed_cli")
                                else:
                                    cli_sel = c_op2.selectbox("Cliente", ["Sin clientes con crédito habilitado"], disabled=True, key="ed_cli_null")
                                    
                                st.write("")
                                c_opt1, c_opt2, c_opt3 = st.columns([2, 2, 1], vertical_alignment="bottom")
                                idx_tipo = 0 if row_ed['tipo_movimiento'] == 'CREDITO' else 1
                                tipo_sel = c_opt1.radio("Tipo de Movimiento", ["Crédito (Fiado)", "Abono (Pago)"], index=idx_tipo, horizontal=True, key="ed_tipo")
                                monto = c_opt2.number_input("Monto ($)", min_value=0, value=int(row_ed['monto']), step=1000, key="ed_monto")
                                
                                st.write("")
                                c_btn, _ = st.columns([1, 4])
                                if c_btn.button("Actualizar Registro", type="primary", use_container_width=True, key="ed_btn_save"):
                                    if cli_sel and cli_sel != "Sin clientes con crédito habilitado" and monto > 0:
                                        id_cli = int(df_cli_filt[df_cli_filt['nombre']==cli_sel]['id'].values[0])
                                        tipo_db = "CREDITO" if "Crédito" in tipo_sel else "ABONO"
                                        editar_movimiento_credito(row_ed['id'], f_mov, id_cli, dict_vend[vend_sel], tipo_db, monto, "Operación App (Editada)", st.session_state.user_name)
                                        st.success("Registro corregido."); time.sleep(0.5); st.rerun()
                                    else:
                                        st.error("Campos inválidos.")
                        else:
                            st.warning("No hay registros recientes habilitados para su edición.")

                    else:
                        f_mov = st.date_input("Fecha", date.today(), format="DD/MM/YYYY")
                        
                        c_op1, c_op2 = st.columns(2)
                        vend_sel = c_op1.selectbox("Repartidor", lista_repartidores_esp_puro, index=idx_def_cred, placeholder="Seleccione un repartidor...", key="rep_op_cred")
                        
                        if vend_sel:
                            if not df_clientes.empty:
                                df_clientes['permite_credito'] = df_clientes['permite_credito'].fillna(False).astype(bool)
                                df_cli_filt = df_clientes[(df_clientes['Repartidor'] == vend_sel) & (df_clientes['permite_credito'])]
                            else:
                                df_cli_filt = pd.DataFrame()
                            
                            if not df_cli_filt.empty:
                                cli_sel = c_op2.selectbox("Cliente", df_cli_filt['nombre'], index=None, placeholder="Seleccione un cliente autorizado...")
                            else:
                                cli_sel = c_op2.selectbox("Cliente", ["Sin clientes con crédito habilitado"], disabled=True)
                            
                            st.write("")
                            c_opt1, c_opt2, c_opt3 = st.columns([2, 2, 1], vertical_alignment="bottom")
                            tipo_sel = c_opt1.radio("Tipo de Movimiento", ["Crédito (Fiado)", "Abono (Pago)"], horizontal=True)
                            monto = c_opt2.number_input("Monto ($)", min_value=0, value=None, step=1000, placeholder="Vacío")
                            
                            st.write("")
                            c_btn, _ = st.columns([1, 4])
                            if c_btn.button("Guardar", type="primary", use_container_width=True):
                                if cli_sel and cli_sel != "Sin clientes con crédito habilitado":
                                    id_cli = int(df_cli_filt[df_cli_filt['nombre']==cli_sel]['id'].values[0])
                                    tipo_db = "CREDITO" if "Crédito" in tipo_sel else "ABONO"
                                    if monto and monto > 0:
                                        registrar_movimiento_credito(f_mov, id_cli, dict_vend[vend_sel], tipo_db, monto, "Operación App", st.session_state.user_name)
                                        st.success("Registrado"); time.sleep(0.5); st.rerun()
                                    else: 
                                        st.error("Monto inválido")
                                else:
                                    st.warning("Debe seleccionar un cliente válido y autorizado para crédito.")
                        else:
                            c_op2.selectbox("Cliente", ["-"], disabled=True)

        # --- NUEVA PESTAÑA: MOVIMIENTOS ---
        with t_mov:
            st.write("")
            st.markdown(f"<h5 style='display:flex; align-items:center;'>{svg_hist} Historial de Transacciones</h5>", unsafe_allow_html=True)
            c_m1, c_m2, c_m3, c_m4 = st.columns([2, 2, 3, 2], vertical_alignment="bottom")
            
            fm_ini = c_m1.date_input("Desde", date.today() - timedelta(days=7), format="DD/MM/YYYY", key="fm_ini_cred")
            fm_fin = c_m2.date_input("Hasta", date.today(), format="DD/MM/YYYY", key="fm_fin_cred")
            
            if es_repartidor:
                v_mov = c_m3.selectbox("Repartidor", [nombre_vendedor_actual], disabled=True, key="sel_mov_rep_dis")
            else:
                v_mov = c_m3.selectbox("Repartidor", ["Todos"] + lista_repartidores_esp_puro, index=0, key="sel_mov_rep")
                
            if c_m4.button("Consultar", type="primary", use_container_width=True, key="btn_mov_cred"):
                id_v_filtro = dict_vend[v_mov] if v_mov != "Todos" and v_mov in dict_vend else None
                df_movs = obtener_historial_movimientos_credito(fm_ini, fm_fin, id_v_filtro)
                
                if not df_movs.empty:
                    st.write("")
                    def color_tipo(val):
                        if val == 'CREDITO': return 'color: #991B1B; font-weight: bold' # Rojo fiado
                        elif val == 'ABONO': return 'color: #166534; font-weight: bold' # Verde abono
                        return ''
                    
                    st.dataframe(
                        df_movs.style.map(color_tipo, subset=['Tipo']).format({"Monto": "$ {:,.0f}"}), 
                        use_container_width=True, 
                        hide_index=True
                    )
                else:
                    st.info("Sin registros en este período.")

        # --- PESTAÑA ORIGINAL: CARTERA DE ESTADO DE DEUDAS ---
        with t_car:
            st.write("")
            st.markdown(f"<h5 style='display:flex; align-items:center;'>{svg_list} Estado de Deudas</h5>", unsafe_allow_html=True)
            
            cr1, cr2, cr3 = st.columns([1, 2, 1], vertical_alignment="bottom")
            f_rep = cr1.date_input("Fecha de Corte", date.today(), format="DD/MM/YYYY")
            
            v_rep = cr2.selectbox("Filtrar por Repartidor", lista_repartidores_esp_puro, index=idx_def_cred, placeholder="Seleccione repartidor...", disabled=es_repartidor, key="sel_rep_cred_rep")
            
            if cr3.button("Consultar", type="primary", use_container_width=True):
                if v_rep:
                    df_estado = obtener_estado_creditos_vendedor(dict_vend[v_rep], f_rep)
                    
                    if not df_estado.empty:
                        st.write("")
                        k1, k2, k3 = st.columns(3)
                        k1.metric("Total Otorgado", fmt_clp(df_estado['total_otorgado'].sum()))
                        k2.metric("Total Pagado", fmt_clp(df_estado['total_pagado'].sum()))
                        k3.metric("Deuda Activa", fmt_clp(df_estado['deuda_actual'].sum()), delta_color="inverse")
                        
                        st.divider()
                        
                        def color_estado(val):
                            if "🔴" in val: return 'background-color: #FEF2F2; color: #991B1B; font-weight:600'
                            if "🟡" in val: return 'background-color: #FFFBEB; color: #92400E; font-weight:600'
                            if "🟢" in val: return 'color: #166534'
                            return ''
                            
                        st.dataframe(
                            df_estado.style.map(color_estado, subset=['estado']).format({"limite_credito": "$ {:,.0f}", "total_otorgado": "$ {:,.0f}", "total_pagado": "$ {:,.0f}", "deuda_actual": "$ {:,.0f}"}), 
                            use_container_width=True, 
                            hide_index=True, 
                            column_config={
                                "nombre": st.column_config.TextColumn("Cliente", width="medium"), 
                                "limite_credito": "Cupo", 
                                "total_otorgado": "Otorgado",
                                "total_pagado": "Pagado",
                                "deuda_actual": "Deuda Activa", 
                                "estado": "Estado"
                            }
                        )
                    else: 
                        st.info("No hay movimientos registrados para la cartera de este repartidor hasta esa fecha.")
                else:
                    st.warning("Debe seleccionar un repartidor.")

    elif seleccion == "Transferencias":
        st.title("Transferencias")
        
        svg_bank = '''<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" fill="#556B2F" class="bi bi-bank" viewBox="0 0 16 16" style="margin-right: 8px; margin-bottom: 2px;"><path d="m8 0 6.61 3h.89a.5.5 0 0 1 .5.5v2a.5.5 0 0 1-.5.5H15v7a.5.5 0 0 1 .485.38l.5 2a.498.498 0 0 1-.485.62H.5a.498.498 0 0 1-.485-.62l.5-2A.501.501 0 0 1 1 13V6H.5a.5.5 0 0 1-.5-.5v-2A.5.5 0 0 1 .5 3h.89zM3.777 3h8.447L8 1zM2 6v7h1V6zm2 0v7h2.5V6zm3.5 0v7h1V6zm2 0v7H12V6zM13 6v7h1V6zm2-1V4H1v1zm-.39 9H1.39l-.25 1h13.72z"/></svg>'''
        svg_check = '''<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" fill="#556B2F" class="bi bi-check2-square" viewBox="0 0 16 16" style="margin-right: 8px; margin-bottom: 2px;"><path d="M3 14.5A1.5 1.5 0 0 1 1.5 13V3A1.5 1.5 0 0 1 3 1.5h8a.5.5 0 0 1 0 1H3a.5.5 0 0 0-.5.5v10a.5.5 0 0 0 .5.5h10a.5.5 0 0 0 .5-.5V8a.5.5 0 0 1 1 0v5a1.5 1.5 0 0 1-1.5 1.5z"/><path d="m8.354 10.354 7-7a.5.5 0 0 0-.708-.708L8 9.293 5.354 6.646a.5.5 0 1 0-.708.708l3 3a.5.5 0 0 0 .708 0z"/></svg>'''
        svg_hist = '''<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" fill="#556B2F" class="bi bi-clock-history" viewBox="0 0 16 16" style="margin-right: 8px; margin-bottom: 2px;"><path d="M8.515 1.019A7 7 0 0 0 8 1V0a8 8 0 0 1 .589.022zm2.004.45a7 7 0 0 0-.985-.299l.219-.976q.576.129 1.126.342zm1.37.71a7 7 0 0 0-.439-.27l.493-.87a8 8 0 0 1 .979.654l-.615.789a7 7 0 0 0-.418-.302zm1.834 1.79a7 7 0 0 0-.653-.796l.724-.69q.406.429.747.91zm.744 1.352a7 7 0 0 0-.214-.468l.893-.45a8 8 0 0 1 .45 1.088l-.95.313a7 7 0 0 0-.179-.483m.53 2.507a7 7 0 0 0-.1-1.025l.985-.17q.1.58.116 1.17zm-.131 1.538q.05-.254.081-.51l.993.123a8 8 0 0 1-.23 1.155l-.964-.267q.069-.247.12-.501m-.952 2.379q.276-.436.486-.908l.914.405q-.24.54-.555 1.038zm-.964 1.205q.183-.183.35-.378l.758.653a8 8 0 0 1-.401.432z"/><path d="M8 1a7 7 0 1 0 4.95 11.95l.707.707A8.001 8.001 0 1 1 8 0z"/><path d="M7.5 3a.5.5 0 0 1 .5.5v5.21l3.248 1.856a.5.5 0 0 1-.496.868l-3.5-2A.5.5 0 0 1 7 9V3.5a.5.5 0 0 1 .5-.5"/></svg>'''

        es_jefatura = rol in ["admin", "supervisor"]
        idx_def_tr = lista_repartidores_esp_puro.index(nombre_vendedor_actual) if es_repartidor and nombre_vendedor_actual in lista_repartidores_esp_puro else None

        if es_repartidor:
            st.write("")
            st.markdown(f"<h5 style='display:flex; align-items:center;'>{svg_hist} Historial</h5>", unsafe_allow_html=True)
            c_r1, c_r2, c_r3, c_r4 = st.columns([2, 2, 3, 2], vertical_alignment="bottom")
            fr_ini = c_r1.date_input("Desde", date.today() - timedelta(days=7), format="DD/MM/YYYY")
            fr_fin = c_r2.date_input("Hasta", date.today(), format="DD/MM/YYYY")
            
            v_rep_tr = c_r3.selectbox("Repartidor", lista_repartidores_esp_puro, index=idx_def_tr, disabled=True, key="sel_rep_trans_rep_view")
            
            if c_r4.button("Consultar", type="primary", use_container_width=True):
                id_v_tr = dict_vend[v_rep_tr]
                df_tr_rep = obtener_reporte_transferencias_filtrado(fr_ini, fr_fin, id_v_tr)
                if not df_tr_rep.empty:
                    st.write("")
                    st.dataframe(df_tr_rep, use_container_width=True, hide_index=True, column_config={"monto": st.column_config.NumberColumn("Monto", format="$ %d"), "banco_emisor": "Banco Emisor", "banco_receptor_info": "Receptor", "tipo": "Concepto", "verificado": "Estado", "Fecha": "Fecha"})
                else: 
                    st.info("Sin registros en esta fecha.")
                
        else:
            tabs_tr = st.tabs(["Operaciones", "Conciliación", "Historial"])
            
            with tabs_tr[0]:
                st.write("")
                c_head, c_tog = st.columns([3, 1], vertical_alignment="bottom")
                c_head.markdown(f"<h5 style='display:flex; align-items:center; margin-bottom: 0;'>{svg_bank} Registro de Pagos</h5>", unsafe_allow_html=True)
                
                modo_correccion = False
                if permiso_editar:
                    modo_correccion = c_tog.toggle("Corregir un registro")

                with st.container(border=True):
                    if modo_correccion:
                        df_editables = obtener_transferencias_editables(es_jefatura)
                        
                        if not df_editables.empty:
                            st.info("Mostrando registros editables según su nivel de credencial (Candado Temporal).")
                            
                            opciones_edicion = {}
                            for _, r in df_editables.iterrows():
                                fecha_str = r['fecha'].strftime("%d/%m/%Y")
                                eq = f"{fecha_str} | {r['Repartidor']} | $ {int(r['monto'])} | {r['tipo_transferencia']}"
                                opciones_edicion[eq] = r
                                
                            sel_reg = st.selectbox("Seleccione el registro a corregir", list(opciones_edicion.keys()), index=None, placeholder="Buscar registro...")
                            
                            if sel_reg:
                                row_ed = opciones_edicion[sel_reg]
                                st.divider()
                                
                                with st.form("edit_transf", border=False):
                                    t1, t2, t3 = st.columns([1, 2, 1])
                                    ft = t1.date_input("Fecha", row_ed['fecha'], format="DD/MM/YYYY")
                                    idx_rep_ed = lista_repartidores_esp_puro.index(row_ed['Repartidor']) if row_ed['Repartidor'] in lista_repartidores_esp_puro else 0
                                    rt = t2.selectbox("Repartidor", lista_repartidores_esp_puro, index=idx_rep_ed)
                                    mt = t3.number_input("Monto ($)", min_value=0, value=int(row_ed['monto']), step=1000)
                                    
                                    t4, t5, t6, t7 = st.columns([1, 1, 1, 1])
                                    idx_tipo = ["Pago Diario", "Abono Crédito"].index(row_ed['tipo_transferencia']) if row_ed['tipo_transferencia'] in ["Pago Diario", "Abono Crédito"] else 0
                                    tt = t4.selectbox("Concepto", ["Pago Diario", "Abono Crédito"], index=idx_tipo)
                                    
                                    met_actual = "Depósito" if "Depósito" in str(row_ed['metodo_pago']) else "Transferencia"
                                    dest_actual = "Banco Chile" if "Chile" in str(row_ed['metodo_pago']) else "Banco Estado"
                                    
                                    met = t5.selectbox("Método", ["Transferencia", "Depósito"], index=["Transferencia", "Depósito"].index(met_actual))
                                    b_dest = t6.selectbox("Cta. Destino", ["Banco Estado", "Banco Chile"], index=["Banco Estado", "Banco Chile"].index(dest_actual))
                                    
                                    idx_orig = l_bancos.index(row_ed['banco_emisor']) if row_ed['banco_emisor'] in l_bancos else 0
                                    b_orig = t7.selectbox("Banco Origen", l_bancos, index=idx_orig)
                                    
                                    st.write("")
                                    c_btn, _ = st.columns([1, 4])
                                    if c_btn.form_submit_button("Actualizar Registro", type="primary", use_container_width=True):
                                        if rt and mt > 0:
                                            editar_transferencia(row_ed['id'], ft, dict_vend[rt], mt, f"{met} a {b_dest}", b_orig, tt, st.session_state.user_name)
                                            st.success("Registro corregido."); time.sleep(0.5); st.rerun()
                                        else:
                                            st.error("Campos inválidos.")
                        else:
                            st.warning("No hay registros recientes habilitados para su edición.")

                    else:
                        with st.form("new_transf", border=False, clear_on_submit=True):
                            t1, t2, t3 = st.columns([1, 2, 1])
                            ft = t1.date_input("Fecha", date.today(), format="DD/MM/YYYY")
                            rt = t2.selectbox("Repartidor", lista_repartidores_esp_puro, index=None, placeholder="Seleccione un repartidor...", key="rep_tr")
                            mt = t3.number_input("Monto ($)", min_value=0, value=None, step=1000, placeholder="Vacío")
                            
                            t4, t5, t6, t7 = st.columns([1, 1, 1, 1])
                            tt = t4.selectbox("Concepto", ["Pago Diario", "Abono Crédito"])
                            met = t5.selectbox("Método", ["Transferencia", "Depósito"])
                            b_dest = t6.selectbox("Cta. Destino", ["Banco Estado", "Banco Chile"])
                            b_orig = t7.selectbox("Banco Origen", l_bancos)
                            
                            st.write("")
                            c_btn, _ = st.columns([1, 4])
                            if c_btn.form_submit_button("Guardar Registro", type="primary", use_container_width=True):
                                if rt and mt and mt > 0:
                                    registrar_transferencia(ft, dict_vend[rt], mt, f"{met} a {b_dest}", b_orig, tt, False, "", st.session_state.user_name)
                                    st.success("Guardado"); time.sleep(0.5); st.rerun()
                                else: 
                                    st.error("Debe seleccionar un repartidor e ingresar un monto mayor a 0.")

            with tabs_tr[1]:
                st.write("")
                st.markdown(f"<h5 style='display:flex; align-items:center;'>{svg_check} Validación de Pagos</h5>", unsafe_allow_html=True)
                df_p = obtener_transferencias_recientes()
                if not df_p.empty:
                    df_p['verificado'] = df_p['verificado'].astype(bool)
                    cols_tr = {
                        "id": None, "fecha": st.column_config.TextColumn("Fecha", disabled=True, width="small"),
                        "Repartidor": st.column_config.TextColumn("Repartidor", disabled=True, width="medium"),
                        "monto": st.column_config.NumberColumn("Monto", format="$ %d", disabled=True, width="small"),
                        "banco_emisor": st.column_config.TextColumn("Banco Emisor", disabled=True, width="medium"),
                        "tipo_transferencia": st.column_config.TextColumn("Concepto", disabled=True, width="medium"),
                        "verificado": st.column_config.CheckboxColumn("Recibido")
                    }
                    
                    def estilizar_transf(row):
                        return ['background-color: #FFFFFF; font-weight: 600; color: #1F2937;' if c == 'verificado' else 'background-color: #F8F9FA; color: #6C757D;' for c in row.index]
                        
                    df_styled_p = df_p.style.apply(estilizar_transf, axis=1)

                    df_ed = st.data_editor(df_styled_p, key="ed_tr", use_container_width=True, hide_index=True, column_config=cols_tr, disabled=not permiso_editar)
                    
                    st.write("")
                    c_btn, _ = st.columns([1, 4])
                    if permiso_editar and c_btn.button("Guardar Cambios", type="primary", use_container_width=True):
                        actualizar_verificacion_masiva(df_ed)
                        st.success("Actualizado"); time.sleep(0.5); st.rerun()
                else: 
                    st.info("Sin registros recientes.")

            with tabs_tr[2]:
                st.write("")
                st.markdown(f"<h5 style='display:flex; align-items:center;'>{svg_hist} Consulta</h5>", unsafe_allow_html=True)
                tr1, tr2, tr3, tr4 = st.columns([2, 2, 3, 2], vertical_alignment="bottom")
                fr_ini = tr1.date_input("Desde", date.today() - timedelta(days=7), key="tr_d_ini", format="DD/MM/YYYY")
                fr_fin = tr2.date_input("Hasta", date.today(), key="tr_d_fin", format="DD/MM/YYYY")
                vr = tr3.selectbox("Repartidor", lista_repartidores_esp_puro, index=None, placeholder="Seleccione un repartidor...", key="rep_tr_rep")
                
                if tr4.button("Consultar", type="primary", use_container_width=True, key="btn_hist_tr"):
                    if vr:
                        df_rep = obtener_reporte_transferencias_filtrado(fr_ini, fr_fin, dict_vend[vr])
                        if not df_rep.empty:
                            st.write("")
                            st.dataframe(df_rep, use_container_width=True, hide_index=True, column_config={"monto": st.column_config.NumberColumn("Monto", format="$ %d"), "banco_emisor": "Banco Emisor", "banco_receptor_info": "Receptor", "tipo": "Concepto", "verificado": "Estado", "Fecha": "Fecha"})
                        else: 
                            st.info("Sin registros.")
                    else:
                        st.warning("Debe seleccionar un repartidor.")

    elif seleccion == "Caja":
        st.title("Libro de Caja")
        
        svg_cash = '''<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" fill="#556B2F" class="bi bi-cash-stack" viewBox="0 0 16 16" style="margin-right: 8px; margin-bottom: 4px;"><path d="M1 3a1 1 0 0 1 1-1h12a1 1 0 0 1 1 1H1zm7 8a2 2 0 1 0 0-4 2 2 0 0 0 0 4z"/><path d="M0 5a1 1 0 0 1 1-1h14a1 1 0 0 1 1 1v8a1 1 0 0 1-1 1H1a1 1 0 0 1-1-1V5zm3 0a2 2 0 0 1-2 2v4a2 2 0 0 1 2 2h10a2 2 0 0 1 2-2V7a2 2 0 0 1-2-2H3z"/></svg>'''
        svg_book = '''<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" fill="#556B2F" class="bi bi-journal-text" viewBox="0 0 16 16" style="margin-right: 8px; margin-bottom: 4px;"><path d="M5 10.5a.5.5 0 0 1 .5-.5h2a.5.5 0 0 1 0 1h-2a.5.5 0 0 1-.5-.5zm0-2a.5.5 0 0 1 .5-.5h5a.5.5 0 0 1 0 1h-5a.5.5 0 0 1-.5-.5zm0-2a.5.5 0 0 1 .5-.5h5a.5.5 0 0 1 0 1h-5a.5.5 0 0 1-.5-.5zm0-2a.5.5 0 0 1 .5-.5h5a.5.5 0 0 1 0 1h-5a.5.5 0 0 1-.5-.5z"/><path d="M3 0h10a2 2 0 0 1 2 2v12a2 2 0 0 1-2 2H3a2 2 0 0 1-2-2v-1h1v1a1 1 0 0 0 1 1h10a1 1 0 0 0 1-1V2a1 1 0 0 0-1-1H3a1 1 0 0 0-1 1v1H1V2a2 2 0 0 1 2-2z"/><path d="M1 5v-.5a.5.5 0 0 1 1 0V5h.5a.5.5 0 0 1 0 1h-2a.5.5 0 0 1 0-1H1zm0 3v-.5a.5.5 0 0 1 1 0V8h.5a.5.5 0 0 1 0 1h-2a.5.5 0 0 1 0-1H1zm0 3v-.5a.5.5 0 0 1 1 0v.5h.5a.5.5 0 0 1 0 1h-2a.5.5 0 0 1 0-1H1z"/></svg>'''
        svg_gear = '''<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" fill="#556B2F" class="bi bi-gear" viewBox="0 0 16 16" style="margin-right: 8px; margin-bottom: 2px;"><path d="M8 4.754a3.246 3.246 0 1 0 0 6.492 3.246 3.246 0 0 0 0-6.492zM5.754 8a2.246 2.246 0 1 1 4.492 0 2.246 2.246 0 0 1-4.492 0z"/><path d="M9.796 1.343c-.527-1.79-3.065-1.79-3.592 0l-.094.319a.873.873 0 0 1-1.255.52l-.292-.16c-1.64-.892-3.433.902-2.54 2.541l.159.292a.873.873 0 0 1-.52 1.255l-.319.094c-1.79.527-1.79 3.065 0 3.592l.319.094a.873.873 0 0 1 .52 1.255l-.16.292c-.892 1.64.901 3.434 2.541 2.54l.292-.159a.873.873 0 0 1 1.255.52l.094.319c.527 1.79 3.065 1.79 3.592 0l.094-.319a.873.873 0 0 1 1.255-.52l.292.16c1.64.893 3.434-.902 2.54-2.541l-.159-.292a.873.873 0 0 1 .52-1.255l.319-.094c1.79-.527 1.79-3.065 0-3.592l-.319-.094a.873.873 0 0 1-.52-1.255l.16-.292c.893-1.64-.902-3.433-2.541-2.54l-.292.159a.873.873 0 0 1-1.255-.52l-.094-.319zm-2.633.283c.246-.835 1.428-.835 1.674 0l.094.319a1.873 1.873 0 0 0 2.693 1.115l.291-.16c.764-.415 1.6.42 1.184 1.185l-.159.292a1.873 1.873 0 0 0 1.116 2.692l.318.094c.835.246.835 1.428 0 1.674l-.319.094a1.873 1.873 0 0 0-1.115 2.693l.16.291c.415.764-.42 1.6-1.185 1.184l-.291-.159a1.873 1.873 0 0 0-2.693 1.116l-.094.318c-.246.835-1.428.835-1.674 0l-.094-.319a1.873 1.873 0 0 0-2.692-1.115l-.292.16c-.764.415-1.6-.42-1.184-1.185l.159-.291A1.873 1.873 0 0 0 1.945 8.93l-.319-.094c-.835-.246-.835-1.428 0-1.674l.319-.094A1.873 1.873 0 0 0 3.06 4.377l-.16-.292c-.415-.764.42-1.6 1.185-1.184l.292.159a1.873 1.873 0 0 0 2.692-1.115l.094-.319z"/></svg>'''

        es_jefatura = rol in ["admin", "supervisor"]
        
        c_f1, _ = st.columns([1, 4])
        fecha_caja = c_f1.date_input("Fecha", date.today(), key="f_caja", format="DD/MM/YYYY")
            
        tabs_caja = st.tabs(["Operación", "Registro", "Ajustes"]) if not es_repartidor else st.tabs(["Consulta"])
        
        # --- PESTAÑA 1: OPERACIÓN ---
        with tabs_caja[0]:
            if permiso_editar:
                st.write("")
                
                # El interruptor estético
                c_head, c_tog = st.columns([3, 1], vertical_alignment="bottom")
                c_head.markdown(f"<h5 style='display:flex; align-items:center; margin-bottom: 0;'>{svg_cash} Movimientos de Caja</h5>", unsafe_allow_html=True)
                
                modo_correccion = c_tog.toggle("Corregir un registro")
                
                df_cat = obtener_categorias_caja()
                df_entidades = obtener_entidades_caja("Especial")

                with st.container(border=True):
                    # ESTADO 1: FORMULARIO DE CORRECCIÓN
                    if modo_correccion:
                        df_editables = obtener_caja_editables("Pan Especial", es_jefatura)
                        
                        if not df_editables.empty:
                            st.info("Mostrando registros editables (Candado Temporal).")
                            
                            opciones_edicion = {}
                            for _, r in df_editables.iterrows():
                                fecha_str = r['fecha'].strftime("%d/%m/%Y")
                                tipo_monto = ""
                                if r['ingreso_efectivo'] > 0: tipo_monto = f"+ $ {int(r['ingreso_efectivo'])} (Efec)"
                                elif r['ingreso_transferencia'] > 0: tipo_monto = f"+ $ {int(r['ingreso_transferencia'])} (Trans)"
                                elif r['egreso'] > 0: tipo_monto = f"- $ {int(r['egreso'])} (Efec)"
                                elif r['egreso_transferencia'] > 0: tipo_monto = f"- $ {int(r['egreso_transferencia'])} (Trans)"
                                
                                eq = f"{fecha_str} | {r['descripcion']} | {r['item']} | {tipo_monto}"
                                opciones_edicion[eq] = r
                                
                            sel_reg = st.selectbox("Seleccione movimiento a corregir", list(opciones_edicion.keys()), index=None, placeholder="Buscar registro...")
                            
                            if sel_reg:
                                row_ed = opciones_edicion[sel_reg]
                                st.divider()
                                
                                c1, c2 = st.columns(2)
                                
                                idx_cat = df_cat[df_cat['id'] == row_ed['id_categoria']].index[0] if pd.notna(row_ed['id_categoria']) and not df_cat.empty else 0
                                cat_sel = c1.selectbox("Categoría", df_cat['nombre'].tolist() if not df_cat.empty else ["-"], index=int(idx_cat))
                                
                                if cat_sel:
                                    id_cat = int(df_cat[df_cat['nombre'] == cat_sel]['id'].values[0])
                                    df_sub = obtener_subcategorias_caja(id_cat)
                                    lista_sub = df_sub['nombre'].tolist() if not df_sub.empty else ["-"]
                                    
                                    idx_sub = 0
                                    if pd.notna(row_ed['id_subcategoria']) and not df_sub.empty:
                                        match_sub = df_sub[df_sub['id'] == row_ed['id_subcategoria']]
                                        if not match_sub.empty: idx_sub = match_sub.index[0]
                                        
                                    sub_sel = c2.selectbox("Subcategoría", lista_sub, index=int(idx_sub))
                                    
                                    c3, c4 = st.columns(2)
                                    lista_ent = df_entidades['nombre'].tolist() if not df_entidades.empty else ["-"]
                                    
                                    idx_ent = 0
                                    if pd.notna(row_ed['id_entidad']) and not df_entidades.empty:
                                        match_ent = df_entidades[df_entidades['id'] == row_ed['id_entidad']]
                                        if not match_ent.empty: idx_ent = match_ent.index[0]
                                        
                                    ent_sel = c3.selectbox("Entidad", lista_ent, index=int(idx_ent))
                                    det = c4.text_input("Descripción", value=str(row_ed['item']) if pd.notna(row_ed['item']) else "")
                                    
                                    c5, c6, c7 = st.columns(3, vertical_alignment="bottom")
                                    
                                    es_ingreso = (row_ed['ingreso_efectivo'] > 0) or (row_ed['ingreso_transferencia'] > 0)
                                    mov_dir_idx = 0 if es_ingreso else 1
                                    mov_dir = c5.selectbox("Movimiento", ["Ingreso", "Egreso"], index=mov_dir_idx, key="ed_mov_dir")
                                    
                                    es_efectivo = (row_ed['ingreso_efectivo'] > 0) or (row_ed['egreso'] > 0)
                                    mov_met_idx = 0 if es_efectivo else 1 
                                    mov_met = c6.selectbox("Método", ["Efectivo", "Transferencia", "Cheque", "Depósito", "Otro"], index=mov_met_idx, key="ed_mov_met")
                                    
                                    monto_actual = max(row_ed['ingreso_efectivo'], row_ed['ingreso_transferencia'], row_ed['egreso'], row_ed['egreso_transferencia'])
                                    monto = c7.number_input("Monto ($)", min_value=0, value=int(monto_actual), step=1000, key="ed_monto_caja")
                                    
                                    st.write("")
                                    col_izq, _ = st.columns([1, 4])
                                    if col_izq.button("Actualizar Registro", type="primary", use_container_width=True):
                                        if sub_sel and ent_sel and monto and monto > 0:
                                            id_sub_val = int(df_sub[df_sub['nombre']==sub_sel]['id'].values[0]) if not df_sub.empty else None
                                            id_ent_val = int(df_entidades[df_entidades['nombre']==ent_sel]['id'].values[0])
                                            
                                            ie = it = ee = et = 0
                                            if mov_dir == "Ingreso":
                                                if mov_met == "Efectivo": ie = monto
                                                else: it = monto 
                                            else:
                                                if mov_met == "Efectivo": ee = monto
                                                else: et = monto 
                                            
                                            editar_movimiento_caja_mill(row_ed['id'], fecha_caja, id_cat, id_sub_val, id_ent_val, ent_sel, det, ie, it, ee, et, st.session_state.user_name)
                                            st.success("Registro corregido.")
                                            time.sleep(0.5)
                                            st.rerun()
                                        else:
                                            st.error("Campos inválidos.")
                        else:
                            st.warning("No hay registros recientes habilitados para su edición.")

                    # ESTADO 2: FORMULARIO DE INGRESO NORMAL
                    else:
                        c1, c2 = st.columns(2)
                        lista_cat = df_cat['nombre'].tolist() if not df_cat.empty else ["-"]
                        cat_sel = c1.selectbox("Categoría", lista_cat, index=None, placeholder="Seleccione una categoría...", key="sel_cat")
                        
                        if cat_sel:
                            id_cat = int(df_cat[df_cat['nombre'] == cat_sel]['id'].values[0])
                            df_sub = obtener_subcategorias_caja(id_cat)
                            lista_sub = df_sub['nombre'].tolist() if not df_sub.empty else ["-"]
                            sub_sel = c2.selectbox("Subcategoría", lista_sub, index=None, placeholder="Seleccione una subcategoría...", key="sel_subcat")
                            
                            c3, c4 = st.columns(2)
                            lista_ent = df_entidades['nombre'].tolist() if not df_entidades.empty else ["-"]
                            ent_sel = c3.selectbox("Entidad", lista_ent, index=None, placeholder="Buscar entidad...", key="sel_ent")
                            det = c4.text_input("Descripción", placeholder="Ej: Pago factura, bono extra...", key="txt_det")
                            
                            c5, c6, c7 = st.columns(3, vertical_alignment="bottom")
                            mov_dir = c5.selectbox("Movimiento", ["Ingreso", "Egreso"], key="sel_mov_dir")
                            mov_met = c6.selectbox("Método", ["Efectivo", "Transferencia", "Cheque", "Depósito", "Otro"], key="sel_mov_met")
                            monto = c7.number_input("Monto ($)", min_value=0, value=None, step=1000, placeholder="Vacío", key="num_monto_caja")
                            
                            st.write("")
                            col_izq, _ = st.columns([1, 4])
                            if col_izq.button("Guardar", type="primary", use_container_width=True, key="btn_guardar_caja"):
                                if sub_sel and ent_sel and monto and monto > 0:
                                    id_sub_val = int(df_sub[df_sub['nombre']==sub_sel]['id'].values[0]) if not df_sub.empty else None
                                    id_ent_val = int(df_entidades[df_entidades['nombre']==ent_sel]['id'].values[0])
                                    
                                    ie = it = ee = et = 0
                                    if mov_dir == "Ingreso":
                                        if mov_met == "Efectivo": ie = monto
                                        else: it = monto 
                                    else:
                                        if mov_met == "Efectivo": ee = monto
                                        else: et = monto 
                                    
                                    guardar_movimiento_caja_mill(fecha_caja, "Pan Especial", id_cat, id_sub_val, id_ent_val, ent_sel, det, ie, it, ee, et)
                                    st.success("Registrado.")
                                    time.sleep(0.5)
                                    st.rerun()
                                else: 
                                    st.error("Complete todos los campos requeridos (Subcategoría, Entidad y Monto válido).")
                        
                    

        with tabs_caja[1]:
            st.write("")
            c_reg1, c_reg2, c_reg3 = st.columns([2, 1, 1], vertical_alignment="bottom")
            c_reg1.markdown(f"<h5 style='display:flex; align-items:center; margin: 0;'>{svg_book} Auditoría Diaria</h5>", unsafe_allow_html=True)
            
            ver_libro_mayor = False
            modo_global = False
            if es_jefatura:
                # 1. Leemos primero el Global para que domine la jerarquía
                modo_global = c_reg3.toggle("Caja Mayor (Kilaco)", value=False, key="tgl_glb_e")
                
                # 2. El switch del módulo se bloquea (disabled) si el global está encendido
                ver_libro_mayor = c_reg2.toggle("Caja Mayor (Especial)", value=True, disabled=modo_global, key="tgl_lm_e")
                
                # 3. Si vemos el global, forzamos internamente a que se vea todo sin filtros
                if modo_global: ver_libro_mayor = True
            
            if modo_global:
                df_caja = obtener_caja_mayor_global(fecha_caja)
            else:
                df_caja = obtener_caja_del_dia(fecha_caja, 'Pan Especial')
            
            if not df_caja.empty:
                df_caja.fillna(0, inplace=True)
                
                if not es_jefatura or not ver_libro_mayor:
                    df_caja = df_caja[~df_caja['rol_creador'].isin(['admin', 'supervisor'])]
                
                if not df_caja.empty:
                    tie=df_caja['ingreso_efectivo'].sum(); tit=df_caja['ingreso_transferencia'].sum()
                    tee=df_caja['egreso_efectivo'].sum(); tet=df_caja['egreso_transferencia'].sum()
                    
                    cols_visuales = {
                        "id": None, "rol_creador": None,
                        "area": st.column_config.TextColumn("Origen", width="small") if modo_global else None,
                        "fecha": None,
                        "entidad": st.column_config.TextColumn("Entidad", width="medium"), 
                        "detalle": st.column_config.TextColumn("Descripción", width="medium"), 
                        "ingreso_efectivo": st.column_config.NumberColumn("Ingreso Efec.", format="$ %d"), 
                        "ingreso_transferencia": st.column_config.NumberColumn("Ingreso Banco", format="$ %d"), 
                        "egreso_efectivo": st.column_config.NumberColumn("Egreso Efec.", format="$ %d"),
                        "egreso_transferencia": st.column_config.NumberColumn("Egreso Banco", format="$ %d")
                    }

                    st.dataframe(df_caja, use_container_width=True, hide_index=True, column_config=cols_visuales)
                    
                    st.write("")
                    k1, k2, k3 = st.columns(3)
                    k1.metric("Ingresos (Efectivo)", fmt_clp(tie))
                    k2.metric("Egresos (Caja Chica)", fmt_clp(tee))
                    saldo_fisico = tie - tee
                    k3.metric("Saldo Físico en Cajón", fmt_clp(saldo_fisico), delta="A favor" if saldo_fisico >= 0 else "Faltante", delta_color="normal" if saldo_fisico >= 0 else "inverse")
                    
                    st.write("")
                    balance_global = (tie+tit)-(tee+tet)
                    color_bg = "#F0FDF4" if balance_global >= 0 else "#FEF2F2"
                    color_tx = "#166534" if balance_global >= 0 else "#991B1B"
                    borde = "#BBF7D0" if balance_global >= 0 else "#FECACA"
                    
                    titulo_balance = "Balance Consolidado Panadería" if modo_global else ("Balance Financiero Total" if ver_libro_mayor else "Balance Operativo")
                    st.markdown(f"<div style='margin: 0; padding: 16px; background-color: {color_bg}; border: 1px solid {borde}; border-radius: 6px; text-align: center; color: {color_tx}; font-size: 16px;'><b>{titulo_balance}:</b> {fmt_clp(balance_global)}</div>", unsafe_allow_html=True)
                else:
                    st.info("No hay movimientos en la vista actual tras aplicar filtros.")
            else: 
                st.info("Sin movimientos registrados.")

        if not es_repartidor:
            with tabs_caja[2]:
                st.write("")
                st.markdown(f"<h5 style='display:flex; align-items:center;'>{svg_gear} Entidades</h5>", unsafe_allow_html=True)
                
                modo_mantenedor = st.radio("Acción", ["Crear Nueva Entidad", "Gestionar Existente"], horizontal=True, label_visibility="collapsed")
                st.write("")
                df_mant = obtener_todas_entidades()
                
                if modo_mantenedor == "Crear Nueva Entidad":
                    with st.form("new_ent_form", clear_on_submit=True, border=False):
                        e1, e2, e3 = st.columns(3)
                        n_ent = e1.text_input("Nombre")
                        t_ent = e2.selectbox("Tipo", ["Empleado", "Proveedor", "Servicio", "Otro"])
                        a_ent = e3.selectbox("Alcance", ["Global", "Especial", "Corriente"])
                        
                        st.write("")
                        col_izq, _ = st.columns([1, 4])
                        if col_izq.form_submit_button("Guardar", type="primary", use_container_width=True):
                            if n_ent:
                                conn_ent = get_conn()
                                try:
                                    with conn_ent.cursor() as cur:
                                        cur.execute("INSERT INTO entidades (nombre, tipo, alcance) VALUES (%s, %s, %s)", (n_ent, t_ent, a_ent))
                                    conn_ent.commit()
                                    st.cache_data.clear()
                                    st.success("Guardado.")
                                    time.sleep(0.5); st.rerun()
                                except Exception as e:
                                    st.error("Error al guardar. Posible duplicado.")
                                finally:
                                    conn_ent.close()
                            else:
                                st.error("El nombre es obligatorio.")
                else:
                    if not df_mant.empty:
                        op_gest = st.selectbox("Seleccione entidad a gestionar", df_mant['nombre'].tolist(), index=None, placeholder="Buscar entidad...")
                        if op_gest:
                            row = df_mant[df_mant['nombre'] == op_gest].iloc[0]
                            with st.form("edit_ent_form", border=False):
                                e1, e2, e3 = st.columns(3)
                                en_nom = e1.text_input("Nombre", row['nombre'])
                                
                                idx_tipo = ["Empleado", "Proveedor", "Servicio", "Otro"].index(row['tipo']) if row['tipo'] in ["Empleado", "Proveedor", "Servicio", "Otro"] else 0
                                en_tip = e2.selectbox("Tipo", ["Empleado", "Proveedor", "Servicio", "Otro"], index=idx_tipo)
                                
                                alcance_db = row['alcance'].capitalize()
                                idx_alcance = ["Global", "Especial", "Corriente"].index(alcance_db) if alcance_db in ["Global", "Especial", "Corriente"] else 0
                                en_alc = e3.selectbox("Alcance", ["Global", "Especial", "Corriente"], index=idx_alcance)
                                
                                st.write("")
                                c_btn1, c_btn2, _ = st.columns([1, 1, 3])
                                if c_btn1.form_submit_button("Guardar Cambios", type="primary"):
                                    conn_ent = get_conn()
                                    try:
                                        with conn_ent.cursor() as cur:
                                            cur.execute("UPDATE entidades SET nombre=%s, tipo=%s, alcance=%s WHERE id=%s", (en_nom, en_tip, en_alc, int(row['id'])))
                                        conn_ent.commit()
                                        st.cache_data.clear()
                                        st.success("Editado.")
                                        time.sleep(0.5); st.rerun()
                                    finally:
                                        conn_ent.close()
                                        
                                if c_btn2.form_submit_button("Dar de Baja"):
                                    conn_ent = get_conn()
                                    try:
                                        with conn_ent.cursor() as cur:
                                            cur.execute("UPDATE entidades SET estado='INACTIVO' WHERE id=%s", (int(row['id']),))
                                        conn_ent.commit()
                                        st.cache_data.clear()
                                        st.error("Entidad desactivada.")
                                        time.sleep(0.5); st.rerun()
                                    finally:
                                        conn_ent.close()
                        
                    else:
                        st.info("No hay entidades registradas para gestionar.")

    elif seleccion == "Admin" and rol == "admin":
        st.title("Administración de Sistema")
        
        tab_bck, tab_bd, tab_mig = st.tabs(["Respaldos", "Base de Datos", "Migración"])
        
        with tab_bck:
            st.markdown("Descarga del sistema completo en formato Excel.")
            if st.button("Generar Backup", type="primary"):
                with st.spinner("Procesando..."):
                    b = descargar_respaldo_completo()
                    st.download_button("Descargar Archivo", b, f"Backup_{date.today()}.xlsx", "application/vnd.ms-excel")
        
        with tab_bd:
            tablas_esp = [
                "control_bandejas", "control_bolsas", "stock", "despacho", "finanzas", 
                "movimientos_credito", "transferencias", "caja_movimientos",
                "caja_categorias", "caja_subcategorias", "entidades", "clientes", 
                "clientes_sugerencias", "productos", "productos_bolsas", "vendedores"
            ]
            
            c1, c2 = st.columns([1, 2])
            ts = c1.selectbox("Tabla", tablas_esp, index=None, placeholder="Seleccione una tabla para visualizar...")
            
            if ts:
                tablas_sin_fecha = ["caja_categorias", "caja_subcategorias", "entidades", "clientes", "clientes_sugerencias", "productos", "productos_bolsas", "vendedores"]
                
                if ts not in tablas_sin_fecha:
                    f_del = c2.date_input("Fecha", date.today(), format="DD/MM/YYYY")
                else:
                    c2.info("Tabla de catálogo. Mostrando todos los registros.")
                
                conn = get_conn()
                try:
                    if ts in tablas_sin_fecha:
                        df = pd.read_sql(f"SELECT * FROM {ts} ORDER BY id DESC", conn)
                    else:
                        df = pd.read_sql(f"SELECT * FROM {ts} WHERE fecha = %s", conn, params=(f_del,))
                    
                    if not df.empty:
                        df_visual = df.copy()
                        
                        # --- CORRECCIÓN ESPACIO-TEMPORAL (UTC a Santiago) ---
                        for col in df_visual.columns:
                            if pd.api.types.is_datetime64_any_dtype(df_visual[col]):
                                try:
                                    if df_visual[col].dt.tz is None:
                                        df_visual[col] = df_visual[col].dt.tz_localize('UTC').dt.tz_convert('America/Santiago')
                                    else:
                                        df_visual[col] = df_visual[col].dt.tz_convert('America/Santiago')
                                    df_visual[col] = df_visual[col].dt.strftime('%d/%m/%Y %H:%M:%S')
                                except:
                                    pass

                        inv_vend = {v: k for k, v in dict_vend.items()}
                        if 'id_vendedor' in df_visual.columns:
                            df_visual['Repartidor'] = df_visual['id_vendedor'].map(inv_vend)
                        if 'id_producto' in df_visual.columns:
                            tabla_prod = "productos_bolsas" if ts == "control_bolsas" else "productos"
                            df_prod_temp = pd.read_sql(f"SELECT id, nombre FROM {tabla_prod}", conn)
                            df_visual['Producto'] = df_visual['id_producto'].map(dict(zip(df_prod_temp['id'], df_prod_temp['nombre'])))
                        if 'id_cliente' in df_visual.columns:
                            df_cli_temp = pd.read_sql("SELECT id, nombre FROM clientes", conn)
                            df_visual['Cliente'] = df_visual['id_cliente'].map(dict(zip(df_cli_temp['id'], df_cli_temp['nombre'])))

                        st.dataframe(df_visual, use_container_width=True)
                    else:
                        st.info("No hay registros en esta tabla para los parámetros seleccionados.")
                    
                    st.divider()
                    col_id, col_dia, col_purge = st.columns(3)
                    
                    with col_id:
                        if not df.empty:
                            st.markdown("**Borrado Específico**")
                            opciones_borrado = {}
                            for _, r in df_visual.iterrows():
                                etiqueta = f"ID: {r['id']}"
                                if 'nombre' in r and pd.notna(r['nombre']): etiqueta += f" | {r['nombre']}"
                                elif 'Repartidor' in r and pd.notna(r['Repartidor']): etiqueta += f" | Rep: {r['Repartidor']}"
                                if 'estado' in r and pd.notna(r['estado']): etiqueta += f" | Est: {r['estado']}"
                                if 'monto' in r and pd.notna(r['monto']): etiqueta += f" | $: {r['monto']}"
                                opciones_borrado[etiqueta] = r['id']
                            
                            seleccion_humana = st.selectbox("Seleccione el registro", list(opciones_borrado.keys()), index=None, placeholder="Buscar...")
                            
                            if seleccion_humana:
                                id_del = opciones_borrado[seleccion_humana]
                                conf_id = st.toggle("Habilitar acción", key=f"tgl_id_esp_{ts}")
                                tablas_catalogo = ["caja_categorias", "caja_subcategorias", "entidades"]
                                texto_btn = "Desactivar Registro" if ts in tablas_catalogo else "Eliminar Registro"
                                
                                if st.button(texto_btn, type="primary"):
                                    if conf_id:
                                        try:
                                            with conn.cursor() as c: 
                                                if ts in tablas_catalogo:
                                                    c.execute(f"UPDATE {ts} SET estado='INACTIVO' WHERE id=%s", (id_del,))
                                                    msg = "Desactivado (Borrado lógico para proteger historial contable)."
                                                else:
                                                    c.execute(f"DELETE FROM {ts} WHERE id=%s", (id_del,))
                                                    msg = "Eliminado físicamente."
                                            conn.commit()
                                            st.cache_data.clear()
                                            st.success(msg)
                                            time.sleep(1.5); st.rerun()
                                        except Exception as e:
                                            st.error(f"La base de datos impidió la acción. Detalle: {e}")
                                    else: 
                                        st.warning("Requiere habilitación.")
                        else:
                            st.write("Sin datos para borrar individualmente.")
                    
                    with col_dia:
                        if ts not in tablas_sin_fecha:
                            st.markdown("**Borrado de Día**")
                            st.write("Elimina todo lo del día seleccionado arriba.")
                            conf_dia = st.toggle("Habilitar borrado por día", key=f"tgl_dia_esp_{ts}")
                            if st.button("Eliminar Día Completo", type="primary"):
                                if conf_dia:
                                    with conn.cursor() as c: c.execute(f"DELETE FROM {ts} WHERE fecha=%s", (f_del,))
                                    conn.commit(); st.cache_data.clear(); st.success("Día Eliminado"); time.sleep(0.5); st.rerun()
                                else: st.warning("Requiere habilitación.")
                        else:
                            st.info("Borrado por fecha no aplica a catálogos.")
                            
                    with col_purge:
                        st.markdown("**Zona de Peligro: Purga Total**")
                        st.write("⚠️ Vacía la tabla por completo y **reinicia los IDs a 1**.")
                        txt_confirm = st.text_input(f"Escriba '{ts}' para confirmar", key=f"purge_txt_{ts}")
                        if st.button("Vaciar y Reiniciar IDs", type="primary"):
                            if txt_confirm == ts:
                                try:
                                    with conn.cursor() as c:
                                        c.execute(f"TRUNCATE TABLE {ts} RESTART IDENTITY CASCADE")
                                    conn.commit()
                                    st.cache_data.clear()
                                    st.success(f"La tabla {ts} ha sido vaciada y sus IDs reiniciados.")
                                    time.sleep(1.5); st.rerun()
                                except Exception as e:
                                    st.error(f"Error al vaciar la tabla: {e}")
                            else:
                                st.warning("Confirmación incorrecta. Escriba el nombre exacto.")
                                
                except Exception as e: st.error(f"Error en consulta: {e}")
                finally: conn.close()
            

        with tab_mig:
            st.write("")
            st.markdown("##### 📥 Inyector Maestro de Historial (Excel)")
            
            tipo_migracion = st.selectbox(
                "Seleccione el módulo a importar:", 
                ["Bandejas (Pan Especial)", "Bolsas (Pan Especial)", "Producción (Pan Especial)", "Despacho (Pan Especial)", "Cobranza (Pan Especial)", "Transferencias (Pan Especial)", "Créditos Históricos"],
                index=None, placeholder="Seleccione el tipo de archivo..."
            )
            
            if tipo_migracion:
                st.caption("Asegúrese de que las columnas del Excel coincidan exactamente con las variables históricas del sistema.")
                archivo = st.file_uploader("Sube el archivo Excel", type=['xlsx'])
                
                if archivo:
                    df_up = pd.read_excel(archivo)
                    st.dataframe(df_up.head(5), use_container_width=True)
                    
                    if st.button("Procesar e Inyectar", type="primary"):
                        conn = get_conn()
                        try:
                            exitos = 0
                            errores = []
                            
                            if tipo_migracion == "Créditos Históricos":
                                df_vend = pd.read_sql("SELECT id, nombre FROM vendedores", conn)
                                dict_vend = dict(zip(df_vend['nombre'].str.lower().str.strip(), df_vend['id']))
                                df_cli = pd.read_sql("SELECT id, nombre FROM clientes", conn)
                                dict_cli = dict(zip(df_cli['nombre'].str.lower().str.strip(), df_cli['id']))

                                with conn.cursor() as c:
                                    for i, row in df_up.iterrows():
                                        v_nom = str(row['Repartidor']).lower().strip()
                                        c_nom = str(row['Cliente']).lower().strip()
                                        mov = str(row['Movimiento']).upper().strip() 
                                        
                                        if mov in ["FIADO", "CRÉDITO", "CREDITO"]: mov = "CREDITO"
                                        elif mov in ["PAGO", "ABONO"]: mov = "ABONO"

                                        id_v = dict_vend.get(v_nom)
                                        id_c = dict_cli.get(c_nom)

                                        if id_v and id_c:
                                            c.execute("INSERT INTO movimientos_credito (fecha, id_cliente, id_vendedor, tipo_movimiento, monto, detalle) VALUES (%s,%s,%s,%s,%s,%s)",
                                                      (row['Fecha'], id_c, id_v, mov, row['Monto'], "Carga Histórica Automática"))
                                            exitos += 1
                                        else:
                                            errores.append(f"Fila {i+2}: No hay match para Repartidor '{row['Repartidor']}' o Cliente '{row['Cliente']}'.")
                            
                            elif tipo_migracion == "Bandejas (Pan Especial)":
                                df_vend = pd.read_sql("SELECT id, nombre FROM vendedores", conn)
                                dict_vend = dict(zip(df_vend['nombre'].str.lower().str.strip(), df_vend['id']))

                                with conn.cursor() as c:
                                    for i, row in df_up.iterrows():
                                        v_nom = str(row['Repartidor']).lower().strip()
                                        id_v = dict_vend.get(v_nom)

                                        if id_v:
                                            c.execute("INSERT INTO control_bandejas (fecha, id_vendedor, saldo_anterior, salida, retorno, saldo_final) VALUES (%s,%s,%s,%s,%s,%s)",
                                                      (row['Fecha'], id_v, safe_int(row['Saldo Inicial']), safe_int(row['Egreso']), safe_int(row['Retorno']), safe_int(row['Saldo Final'])))
                                            exitos += 1
                                        else:
                                            errores.append(f"Fila {i+2}: No hay match exacto para el Repartidor '{row['Repartidor']}'.")

                            elif tipo_migracion == "Bolsas (Pan Especial)":
                                df_prod = pd.read_sql("SELECT id, nombre FROM productos_bolsas", conn)
                                dict_prod = dict(zip(df_prod['nombre'].str.lower().str.strip(), df_prod['id']))

                                with conn.cursor() as c:
                                    for i, row in df_up.iterrows():
                                        p_nom = str(row['Tipo de Producto']).lower().strip()
                                        id_p = dict_prod.get(p_nom)

                                        if id_p:
                                            c.execute("INSERT INTO control_bolsas (fecha, id_producto, stock_cajas_ayer, ingreso_cajas, produccion_ayer, stock_cajas_final) VALUES (%s,%s,%s,%s,%s,%s)",
                                                      (row['Fecha'], id_p, safe_float(row['Stock Disponible (Cajas)']), safe_float(row['Ingreso de Cajas Nuevas']), safe_int(row['Producción Día Anterior']), safe_float(row['Saldo en Cajas'])))
                                            exitos += 1
                                        else:
                                            errores.append(f"Fila {i+2}: No hay match exacto para el Producto '{row['Tipo de Producto']}'.")

                            elif tipo_migracion == "Producción (Pan Especial)":
                                df_prod = pd.read_sql("SELECT id, nombre FROM productos", conn)
                                dict_prod = dict(zip(df_prod['nombre'].str.lower().str.strip(), df_prod['id']))

                                with conn.cursor() as c:
                                    for i, row in df_up.iterrows():
                                        p_nom = str(row['Producto']).lower().strip()
                                        id_p = dict_prod.get(p_nom)

                                        if id_p:
                                            c.execute("INSERT INTO stock (fecha, id_producto, stock_inicial, fabricacion, stock_final, bolsas_necesarias) VALUES (%s,%s,%s,%s,%s,%s)",
                                                      (row['Fecha'], id_p, safe_int(row['Stock Inicial']), safe_int(row['Fabricacion']), safe_int(row['Stock Final']), safe_int(row['Bolsas Necesarias'])))
                                            exitos += 1
                                        else:
                                            errores.append(f"Fila {i+2}: No hay match exacto para el Producto '{row['Producto']}'.")

                            elif tipo_migracion == "Transferencias (Pan Especial)":
                                df_vend = pd.read_sql("SELECT id, nombre FROM vendedores", conn)
                                dict_vend = dict(zip(df_vend['nombre'].str.lower().str.strip(), df_vend['id']))

                                with conn.cursor() as c:
                                    for i, row in df_up.iterrows():
                                        v_nom = str(row['Repartidor']).lower().strip()
                                        id_v = dict_vend.get(v_nom)

                                        if id_v:
                                            val_estado = str(row['Estado']).strip().lower()
                                            es_verificado = 1 if val_estado == 'recibido' else 0
                                            
                                            c.execute("INSERT INTO transferencias (fecha, id_vendedor, monto, metodo_pago, banco_emisor, verificado, tipo_transferencia, comentario) VALUES (%s,%s,%s,%s,%s,%s,%s,%s)",
                                                      (row['Fecha'], id_v, safe_int(row['Monto']), str(row['Método']), str(row['Banco Emisor']), es_verificado, str(row['Tipo Transferencia']), "Carga Histórica"))
                                            exitos += 1
                                        else:
                                            errores.append(f"Fila {i+2}: No hay match exacto para el Repartidor '{row['Repartidor']}'.")

                            elif tipo_migracion == "Despacho (Pan Especial)":
                                df_vend = pd.read_sql("SELECT id, nombre FROM vendedores", conn)
                                dict_vend = dict(zip(df_vend['nombre'].str.lower().str.strip(), df_vend['id']))
                                
                                df_prod = pd.read_sql("SELECT id, nombre FROM productos", conn)
                                dict_prod = dict(zip(df_prod['nombre'].str.lower().str.strip(), df_prod['id']))

                                with conn.cursor() as c:
                                    for i, row in df_up.iterrows():
                                        v_nom = str(row['Repartidor']).lower().strip()
                                        p_nom = str(row['Producto']).lower().strip()
                                        
                                        id_v = dict_vend.get(v_nom)
                                        id_p = dict_prod.get(p_nom)

                                        if id_v and id_p:
                                            # Extraemos los valores de forma segura tolerando variaciones en los nombres de columnas del Excel
                                            s_ant = safe_int(row.get('Saldo Anterior', 0))
                                            carga = safe_int(row.get('Carga', 0))
                                            dev = safe_int(row.get('Devolución', row.get('Devolucion', 0))) 
                                            s_act = safe_int(row.get('Saldo Actual', row.get('Saldo Final', 0)))
                                            
                                            # Recalculamos la venta real por seguridad algorítmica
                                            vta = max(0, (s_ant + carga) - dev - s_act)
                                            
                                            c.execute("""INSERT INTO despacho 
                                                         (fecha, id_vendedor, id_producto, saldo_anterior, carga, devolucion_muestra, saldo_actual, venta_unidades) 
                                                         VALUES (%s,%s,%s,%s,%s,%s,%s,%s)""",
                                                      (row['Fecha'], id_v, id_p, s_ant, carga, dev, s_act, vta))
                                            exitos += 1
                                        else:
                                            errores.append(f"Fila {i+2}: No hay match para Repartidor '{row.get('Repartidor')}' o Producto '{row.get('Producto')}'.")

                            elif tipo_migracion == "Cobranza (Pan Especial)":
                                df_vend = pd.read_sql("SELECT id, nombre FROM vendedores", conn)
                                dict_vend = dict(zip(df_vend['nombre'].str.lower().str.strip(), df_vend['id']))

                                with conn.cursor() as c:
                                    for i, row in df_up.iterrows():
                                        v_nom = str(row.get('Repartidor', '')).lower().strip()
                                        id_v = dict_vend.get(v_nom)

                                        if id_v:
                                            # Extracción segura con valores por defecto si la columna viene vacía
                                            f_val = row['Fecha']
                                            cc = safe_int(row.get('Cobro Créditos', 0))
                                            co = safe_int(row.get('Créditos (Fiado)', 0))
                                            ds = safe_int(row.get('Descuentos', 0))
                                            bn = safe_int(row.get('Bencina', 0))
                                            su = safe_int(row.get('Sueldo', 0))
                                            om = safe_int(row.get('Otros Gastos', 0))
                                            od = str(row.get('Detalle Otros Gastos', 'Varios'))
                                            ef = safe_int(row.get('Efectivo', 0))
                                            tr = safe_int(row.get('Transferencia', 0))
                                            pc = safe_int(row.get('Centralizado', 0))
                                            cc_det = str(row.get('Detalle Cobro Créditos', 'Varios'))

                                            c.execute("""INSERT INTO finanzas 
                                                         (fecha, id_vendedor, creditos_cobrados, creditos_otorgados, 
                                                          descuentos_total, bencina, sueldo, otros_gastos_monto, 
                                                          otros_gastos_detalle, efectivo_rendido, transferencia_rendida, 
                                                          pago_centralizado, creditos_cobrados_detalle) 
                                                         VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
                                                      (f_val, id_v, cc, co, ds, bn, su, om, od, ef, tr, pc, cc_det))
                                            exitos += 1
                                        else:
                                            errores.append(f"Fila {i+2}: No hay match para Repartidor '{row.get('Repartidor')}'.")

                            conn.commit()
                            st.cache_data.clear()
                            
                            if exitos > 0: 
                                st.success(f"✅ {exitos} registros inyectados en la base de datos de {tipo_migracion}.")
                            if errores:
                                st.warning("⚠️ Los siguientes registros fueron ignorados (revisa la ortografía en el Excel frente a tus catálogos de sistema):")
                                for e in errores: st.write(e)
                                
                        except Exception as e:
                            st.error(f"Error procesando el archivo: {e}")
                        finally:
                            conn.close()
           
            

# ----------------------------------------------------
# APLICACIÓN PAN CORRIENTE (CORREGIDA: FILTRO VENDEDOR)
# ----------------------------------------------------
from streamlit_option_menu import option_menu

def app_pan_corriente():
    rol = st.session_state.user_role
    mi_id_vendedor = st.session_state.id_vendedor
    es_repartidor = rol == "repartidor_corr"
    
    # --- 1. CONFIGURACIÓN DE MENÚ Y PERMISOS ---
    opciones_full = ["Producción", "Despacho", "Cobranza", "Clientes", "Caja"]
    iconos_full   = ["tools",      "truck",    "currency-dollar", "people", "cash-stack"]
    
    if rol in ["admin", "pan_corriente", "supervisor"]:
        menu_options = opciones_full; menu_icons = iconos_full; permiso_editar = True
    elif rol == "cajero_integral":
        indices = [2, 3, 4] # Cobranza, Clientes, Caja
        menu_options = [opciones_full[i] for i in indices]; menu_icons = [iconos_full[i] for i in indices]; permiso_editar = True
    elif es_repartidor:
        indices_permitidos = [1, 2, 3]
        menu_options = [opciones_full[i] for i in indices_permitidos]
        menu_icons = [iconos_full[i] for i in indices_permitidos]
        permiso_editar = False
    else:
        menu_options = []; menu_icons = []; permiso_editar = False

    if rol == "admin":
        menu_options.append("Admin"); menu_icons.append("gear")

    # --- 2. SIDEBAR ---
    with st.sidebar:
        c_logo1, c_logo2, c_logo3 = st.columns([1, 1.5, 1])
        with c_logo2:
            try: st.image("logo.png", use_container_width=True)
            except: st.write("🥖")
        
        st.markdown("<h3 style='text-align: center; margin-top:0px;'>KILACO ERP</h3>", unsafe_allow_html=True)
        if not menu_options: st.error("Sin acceso."); st.stop()
            
        seleccion = option_menu(
            menu_title=None, 
            options=menu_options,
            icons=menu_icons,
            menu_icon="cast", 
            default_index=0,
            styles={"container": {"padding": "0!important", "background-color": "#ffffff"}, "icon": {"color": "#556B2F", "font-size": "14px"}, "nav-link": {"font-size": "14px", "text-align": "left", "margin":"0px", "--hover-color": "#f0f2f6"}, "nav-link-selected": {"background-color": "#556B2F", "font-weight": "600"}}
        )
        st.markdown("---")
        if st.button("Volver al Menú", use_container_width=True): st.session_state.current_module = "menu"; st.rerun()

    # --- 3. REFERENCIAS DINÁMICAS ---
    l_comunas, l_bancos, df_vend = get_referencias()

    if es_repartidor and mi_id_vendedor:
        df_vend = df_vend[df_vend['id'] == mi_id_vendedor]
        if df_vend.empty: st.error("ID inválido."); st.stop()

    dict_vend = dict(zip(df_vend['nombre'], df_vend['id']))
    
    if es_repartidor: 
        vendedores_corriente = df_vend['nombre'].tolist() 
    else: 
        vendedores_corriente = df_vend[df_vend['area'].isin(['corriente', 'ambos'])]['nombre'].tolist()

    nombre_vendedor_actual = df_vend['nombre'].iloc[0] if es_repartidor else None
    
    # Índice neutro: Vacío para jefaturas, anclado para repartidor
    idx_defecto = vendedores_corriente.index(nombre_vendedor_actual) if es_repartidor and nombre_vendedor_actual in vendedores_corriente else None

    # --- 4. ENRUTAMIENTO DE VISTAS ---

    if seleccion == "Producción":
        st.title("Planificación Diaria")
        
        svg_prod = '''<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" fill="#556B2F" class="bi bi-clipboard-data" viewBox="0 0 16 16" style="margin-right: 8px; margin-bottom: 4px;"><path d="M4 11a1 1 0 1 1 2 0v1a1 1 0 1 1-2 0v-1zm6-4a1 1 0 1 1 2 0v5a1 1 0 1 1-2 0V7zM7 9a1 1 0 0 1 2 0v3a1 1 0 1 1-2 0V9z"/><path d="M4 1.5H3a2 2 0 0 0-2 2V14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V3.5a2 2 0 0 0-2-2h-1v1h1a1 1 0 0 1 1 1V14a1 1 0 0 1-1 1H3a1 1 0 0 1-1-1V3.5a1 1 0 0 1 1-1h1v-1z"/><path d="M9.5 1a.5.5 0 0 1 .5.5v1a.5.5 0 0 1-.5.5h-3a.5.5 0 0 1-.5-.5v-1a.5.5 0 0 1 .5-.5h3zm-3-1A1.5 1.5 0 0 0 5 1.5v1A1.5 1.5 0 0 0 6.5 4h3A1.5 1.5 0 0 0 11 2.5v-1A1.5 1.5 0 0 0 9.5 0h-3z"/></svg>'''
        svg_adj = '''<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" fill="#556B2F" class="bi bi-sliders" viewBox="0 0 16 16" style="margin-right: 8px; margin-bottom: 4px;"><path fill-rule="evenodd" d="M11.5 2a1.5 1.5 0 1 0 0 3 1.5 1.5 0 0 0 0-3zM9.05 3a2.5 2.5 0 0 1 4.9 0H16v1h-2.05a2.5 2.5 0 0 1-4.9 0H0V3h9.05zM4.5 7a1.5 1.5 0 1 0 0 3 1.5 1.5 0 0 0 0-3zM2.05 8a2.5 2.5 0 0 1 4.9 0H16v1H6.95a2.5 2.5 0 0 1-4.9 0H0V8h2.05zm9.45 4a1.5 1.5 0 1 0 0 3 1.5 1.5 0 0 0 0-3zm-2.45 1a2.5 2.5 0 0 1 4.9 0H16v1h-2.05a2.5 2.5 0 0 1-4.9 0H0v-1h9.05z"/></svg>'''

        c_p1, _ = st.columns([1, 4])
        f_prod = c_p1.date_input("Fecha", date.today(), format="DD/MM/YYYY")
        
        df_unificado, extras_dict = get_produccion_corriente_unificada(f_prod)
        
        tab_tabla, tab_extras = st.tabs(["Resumen Diario", "Extras"])
        
        with tab_tabla:
            st.write("")
            st.markdown(f"<h5 style='display:flex; align-items:center;'>{svg_prod} Caritas</h5>", unsafe_allow_html=True)
            with st.container(border=True):
                st.dataframe(
                    df_unificado, hide_index=True, use_container_width=True,
                    column_config={
                        "Concepto": st.column_config.TextColumn("Responsable / Destino", width="medium"), 
                        "rinde_noche": st.column_config.NumberColumn("Turno Noche", format="%d kg"), 
                        "rinde_dia": st.column_config.NumberColumn("Turno Día", format="%d kg")
                    }
                )
                
                st.divider()
                k1, k2, k3 = st.columns(3)
                k1.metric("Total Noche", f"{int(df_unificado['rinde_noche'].sum())} kg")
                k2.metric("Total Día", f"{int(df_unificado['rinde_dia'].sum())} kg")
                k3.metric("Producción Total", f"{int(df_unificado['rinde_noche'].sum() + df_unificado['rinde_dia'].sum())} kg")

        if permiso_editar:
            with tab_extras:
                st.write("")
                st.markdown(f"<h5 style='display:flex; align-items:center;'>{svg_adj} Ajustes de Extras</h5>", unsafe_allow_html=True)
                with st.container(border=True):
                    c1, c2, c3 = st.columns(3)
                    
                    with c1:
                        st.markdown("**Ración**")
                        rd = st.number_input("Turno Día (kg)", value=val_gui(extras_dict['rd']), step=1, key="rd")
                        rn = st.number_input("Turno Noche (kg)", value=val_gui(extras_dict['rn']), step=1, key="rn")
                    with c2:
                        st.markdown("**Adicional**")
                        ad = st.number_input("Turno Día (kg)", value=val_gui(extras_dict['ad']), step=1, key="ad")
                        an = st.number_input("Turno Noche (kg)", value=val_gui(extras_dict['an']), step=1, key="an")
                    with c3:
                        st.markdown("**Kilaco**")
                        kd = st.number_input("Turno Día (kg)", value=val_gui(extras_dict['kd']), step=1, key="kd")
                        kn = st.number_input("Turno Noche (kg)", value=val_gui(extras_dict['kn']), step=1, key="kn")
                        
                    st.write("")
                    col_izq, _ = st.columns([1, 2])
                    if col_izq.button("Guardar Cambios", type="primary", use_container_width=True):
                        save_extras_produccion(f_prod, {"rd": rd, "rn": rn, "ad": ad, "an": an, "kd": kd, "kn": kn})
                        st.success("Guardado"); time.sleep(0.5); st.rerun()

    elif seleccion == "Despacho":
        st.title("Registro de Salidas")
        
        svg_truck = '''<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" fill="#556B2F" class="bi bi-truck" viewBox="0 0 16 16" style="margin-right: 8px; margin-bottom: 4px;"><path d="M0 3.5A1.5 1.5 0 0 1 1.5 2h9A1.5 1.5 0 0 1 12 3.5V5h1.02a1.5 1.5 0 0 1 1.17.563l1.481 1.85a1.5 1.5 0 0 1 .329.938V10.5a1.5 1.5 0 0 1-1.5 1.5H14a2 2 0 1 1-4 0H5a2 2 0 1 1-3.998-.085A1.5 1.5 0 0 1 0 10.5v-7zm1.294 7.456A1.999 1.999 0 0 1 4.732 11h5.536a2.01 2.01 0 0 1 .732-.732V3.5a.5.5 0 0 0-.5-.5h-9a.5.5 0 0 0-.5.5v7a.5.5 0 0 0 .294.456zM12 10a2 2 0 0 1 1.732 1h.768a.5.5 0 0 0 .5-.5V8.35a.5.5 0 0 0-.11-.312l-1.48-1.85A.5.5 0 0 0 13.02 6H12v4zm-9 1a1 1 0 1 0 0 2 1 1 0 0 0 0-2zm9 0a1 1 0 1 0 0 2 1 1 0 0 0 0-2z"/></svg>'''
        
        c1, c2, c3 = st.columns([1, 2, 2], vertical_alignment="bottom")
        f_desp = c1.date_input("Fecha", date.today(), format="DD/MM/YYYY")
        
        v_desp = c2.selectbox("Repartidor", vendedores_corriente, index=idx_defecto, placeholder="Seleccione un repartidor...", disabled=es_repartidor, key="sel_rep_corr")
        
        if v_desp in dict_vend:
            id_v_desp = dict_vend.get(v_desp)
            df_ruta = get_despacho_corriente(f_desp, id_v_desp)
            
            if not df_ruta.empty:
                st.write("")
                st.markdown(f"<h5 style='display:flex; align-items:center;'>{svg_truck} Planilla de Cargas Diarias</h5>", unsafe_allow_html=True)
                
                with st.container(border=True):
                    cols_cfg = {
                        "id": None, "fecha": None, "id_cliente": None, "id_vendedor": None,
                        "cliente": st.column_config.TextColumn("Cliente", disabled=True, width="medium"),
                        "precio_aplicado": st.column_config.NumberColumn("🔒 Precio", format="$ %d", disabled=True, width="small"),
                        "saldo_anterior": None, "deuda_final": None, "ventas_monto": None, "total_pagar": None, "paga": None, "pago_centralizado": None,
                        "total_carga": st.column_config.NumberColumn("🔒 Total Kg", disabled=True, width="small")
                    }
                    for i in range(1, 9): cols_cfg[f"carga_{i}"] = st.column_config.NumberColumn(f"T{i}", width="small")

                    column_order = ["cliente", "precio_aplicado"] + [f"carga_{i}" for i in range(1, 9)] + ["total_carga"]
                    cols_finales = column_order + ["id", "fecha", "id_cliente", "id_vendedor"]
                    for col in cols_finales:
                        if col not in df_ruta.columns: df_ruta[col] = 0
                    
                    df_visual = df_ruta[cols_finales]
                    altura_dinamica = (len(df_visual) * 36) + 42 
                    
                    # EL PINCEL: Si la columna empieza con "carga_", es blanca y editable. El resto gris.
                    def estilizar_despacho_corr(row):
                        return ['background-color: #FFFFFF; font-weight: 600; color: #1F2937;' if str(c).startswith('carga_') else 'background-color: #F8F9FA; color: #6C757D;' for c in row.index]
                    
                    df_styled_ruta = df_visual.style.apply(estilizar_despacho_corr, axis=1)
                    
                    edited_ruta = st.data_editor(df_styled_ruta, column_config=cols_cfg, hide_index=True, use_container_width=True, height=altura_dinamica, disabled=not permiso_editar)
                    
                    tot_kg = int(edited_ruta['total_carga'].sum())
                    
                    st.write("")
                    cd1, cd2 = st.columns([1, 3], vertical_alignment="center")
                    if permiso_editar and cd1.button("Guardar Cargas", type="primary", use_container_width=True):
                        save_despacho_corriente(edited_ruta)
                        st.success("Salidas guardadas."); time.sleep(0.5); st.rerun()
                        
                    cd2.info(f"Total Asignado a Ruta: **{tot_kg} kg**")

    elif seleccion == "Cobranza":
        st.title("Cobranza y Rendición")
        
        svg_pay = '''<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" fill="#556B2F" class="bi bi-wallet2" viewBox="0 0 16 16" style="margin-right: 8px; margin-bottom: 2px;"><path d="M12.136.326A1.5 1.5 0 0 1 14 1.78V3h.5A1.5 1.5 0 0 1 16 4.5v9a1.5 1.5 0 0 1-1.5 1.5h-13A1.5 1.5 0 0 1 0 13.5v-9a1.5 1.5 0 0 1 1.432-1.499L12.136.326zM5.562 3H13V1.78a.5.5 0 0 0-.621-.484L5.562 3zM1.5 4a.5.5 0 0 0-.5.5v9a.5.5 0 0 0 .5.5h13a.5.5 0 0 0 .5-.5v-9a.5.5 0 0 0-.5-.5h-13z"/></svg>'''
        svg_coin = '''<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="#556B2F" class="bi bi-cash-coin" viewBox="0 0 16 16" style="margin-right: 6px; margin-bottom: 2px;"><path fill-rule="evenodd" d="M11 15a4 4 0 1 0 0-8 4 4 0 0 0 0 8zm5-4a5 5 0 1 1-10 0 5 5 0 0 1 10 0z"/><path d="M9.438 11.944c.047.596.518 1.06 1.363 1.116v.44h.375v-.443c.875-.061 1.386-.529 1.386-1.207 0-.618-.39-.936-1.09-1.1l-.296-.07v-1.2c.376.043.614.248.671.532h.658c-.047-.575-.54-1.024-1.329-1.073V8.5h-.375v.45c-.747.073-1.255.522-1.255 1.158 0 .562.378.92 1.007 1.066l.248.061v1.272c-.384-.058-.639-.27-.696-.563h-.668zm1.36-1.354c-.369-.085-.569-.26-.569-.522 0-.294.216-.514.572-.578v1.1h-.003zm.432.746c.449.104.655.272.655.569 0 .339-.257.571-.709.614v-1.195l.054.012z"/><path d="M1 0a1 1 0 0 0-1 1v8a1 1 0 0 0 1 1h4.083c.058-.344.145-.678.258-1H3a2 2 0 0 0-2-2V3a2 2 0 0 0 2-2h10a2 2 0 0 0 2 2v3.528c.38.34.717.728 1 1.154V1a1 1 0 0 0-1-1H1z"/><path d="M9.998 5.083 10 5a2 2 0 1 0-3.132 1.65 5.982 5.982 0 0 1 3.13-1.567z"/></svg>'''
        svg_chart = '''<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" fill="#556B2F" class="bi bi-bar-chart" viewBox="0 0 16 16" style="margin-right: 8px; margin-bottom: 2px;"><path d="M4 11H2v3h2v-3zm5-4H7v7h2V7zm5-5v12h-2V2h2zm-2-1a1 1 0 0 0-1 1v12a1 1 0 0 0 1 1h2a1 1 0 0 0 1-1V2a1 1 0 0 0-1-1h-2zM6 7a1 1 0 0 1 1-1h2a1 1 0 0 1 1 1v7a1 1 0 0 1-1 1H7a1 1 0 0 1-1-1V7zm-5 4a1 1 0 0 1 1-1h2a1 1 0 0 1 1 1v3a1 1 0 0 1-1 1H2a1 1 0 0 1-1-1v-3z"/></svg>'''

        c1, c2, c3 = st.columns([1, 2, 2], vertical_alignment="bottom")
        f_cob = c1.date_input("Fecha", date.today(), format="DD/MM/YYYY")
        
        v_cob = c2.selectbox("Repartidor", vendedores_corriente, index=idx_defecto, placeholder="Seleccione un repartidor...", disabled=es_repartidor, key="sel_rep_cob_corr")
        id_vc = dict_vend.get(v_cob) if v_cob in dict_vend else None
        
        tab_pagos, tab_ren, tab_rep = st.tabs(["Pagos", "Rendición", "Reporte"])
        
                # --- PESTAÑA 1: PAGOS ---
        with tab_pagos:
            if id_vc:
                df_pagos = get_despacho_corriente(f_cob, id_vc)
                if not df_pagos.empty:
                    st.write("")
                    st.markdown(f"<h5 style='display:flex; align-items:center;'>{svg_coin} Recaudación por Cliente</h5>", unsafe_allow_html=True)
                    with st.container(border=True):
                        cols_pagos = {
                            "id": None, "fecha": None, "id_cliente": None, "id_vendedor": None, "precio_aplicado": None,
                            "cliente": st.column_config.TextColumn("Cliente", disabled=True, width="medium"),
                            "saldo_anterior": st.column_config.NumberColumn("🔒 Saldo Ant.", format="$ %d", disabled=True, width="small"),
                            "total_carga": st.column_config.NumberColumn("🔒 Total Kg", disabled=True, format="%d", width="small"),
                            "ventas_monto": st.column_config.NumberColumn("🔒 Venta ($)", format="$ %d", disabled=True, width="small"),
                            "total_pagar": st.column_config.NumberColumn("🔒 A Pagar", format="$ %d", disabled=True, width="small"),
                            "paga": st.column_config.NumberColumn("Efectivo/Transf", format="$ %d", width="small"),
                            "pago_centralizado": st.column_config.NumberColumn("Centralizado", format="$ %d", width="small"),
                            "deuda_final": st.column_config.NumberColumn("🔒 Nueva Deuda", format="$ %d", disabled=True, width="small")
                        }
                        for i in range(1, 9): cols_pagos[f"carga_{i}"] = None

                        column_order = ["cliente", "saldo_anterior", "ventas_monto", "total_carga", "total_pagar", "paga", "pago_centralizado", "deuda_final"]
                        cols_finales = column_order + ["id", "fecha", "id_cliente", "id_vendedor", "precio_aplicado"] + [f"carga_{i}" for i in range(1, 9)]
                        for col in cols_finales:
                            if col not in df_pagos.columns: df_pagos[col] = 0
                        
                        df_visual_p = df_pagos[cols_finales]
                        altura_dinamica_p = (len(df_visual_p) * 36) + 42
                        
                        # EL PINCEL: Paga y Pago Centralizado van en blanco, el resto gris
                        def estilizar_pagos_corr(row):
                            return ['background-color: #FFFFFF; font-weight: 600; color: #1F2937;' if c in ['paga', 'pago_centralizado'] else 'background-color: #F8F9FA; color: #6C757D;' for c in row.index]
                        
                        df_styled_p = df_visual_p.style.apply(estilizar_pagos_corr, axis=1)
                        
                        edited_pagos = st.data_editor(df_styled_p, column_config=cols_pagos, hide_index=True, use_container_width=True, height=altura_dinamica_p, disabled=not permiso_editar)
                        
                        st.write("")
                        if permiso_editar and st.button("Guardar Pagos", type="primary"):
                            save_despacho_corriente(edited_pagos)
                            st.success("Pagos guardados."); time.sleep(0.5); st.rerun()

        # --- PESTAÑA 2: RENDICIÓN ---
        with tab_ren:
            if id_vc:
                fin = get_finanzas_corriente(f_cob, id_vc)
                deuda_especial = obtener_deuda_especial_repartidor(f_cob, id_vc)
                recaudo_corr = fin['recaudo']
                tot_ing = recaudo_corr + deuda_especial
                
                df_desp_cent = get_despacho_corriente(f_cob, id_vc)
                tot_cent_c = int(df_desp_cent['pago_centralizado'].sum()) if not df_desp_cent.empty else 0

                st.write("")
                with st.container(border=True):
                    st.markdown(f"<h5 style='display:flex; align-items:center;'>{svg_pay} Flujo de Caja</h5>", unsafe_allow_html=True)
                    st.write("")
                    
                    st.markdown("**1. Dinero Exigible**")
                    ci1, ci2, ci3 = st.columns(3)
                    ci1.metric("Pan Corriente", fmt_clp(recaudo_corr), "Ventas en ruta", delta_color="off")
                    ci2.metric("Pan Especial", fmt_clp(deuda_especial), "Total con Dscto", delta_color="off")
                    ci3.metric("Total a Rendir", fmt_clp(tot_ing), "Base de cálculo", delta_color="normal")
                    
                    st.divider()
                    st.markdown("**2. Gastos y Comisiones**")
                    
                    c_gr1, c_gr2, c_gr3 = st.columns(3)
                    comision = c_gr1.number_input("🔒 Comisión (Dinámica)", value=fin['comision'], step=100, disabled=True)
                    bencina = c_gr2.number_input("🔒 Bencina (Fija)", value=fin['bencina'], step=1000, disabled=True)
                    sueldo = c_gr3.number_input("Sueldo", value=val_gui(fin['sueldo']), step=1000, disabled=not permiso_editar)
                    
                    st.write("")
                    st.caption("Otros Gastos")
                    detalle_defecto = fin['det'] if fin.get('det') else "Varios"
                    df_om_c = st.data_editor(pd.DataFrame([{"Detalle": detalle_defecto, "Monto": int(fin.get('otros', 0))}]), num_rows="dynamic", key="grid_om_c", hide_index=True, column_config={"Monto": st.column_config.NumberColumn(format="$%d", required=True)}, disabled=not permiso_editar, use_container_width=True)
                    t_om_c = df_om_c['Monto'].sum()
                    txt_om_c = ", ".join(df_om_c['Detalle'].astype(str).tolist())
                    
                    tot_gas = comision + bencina + val_db(sueldo) + t_om_c
                    deuda_neta = tot_ing - tot_gas
                    
                    st.divider()
                    st.markdown(f"<h6 style='display:flex; align-items:center; color:#556B2F;'>{svg_coin} Cierre y Pagos</h6>", unsafe_allow_html=True)
                    st.write("")
                    
                    st.markdown(f"<h4 style='color:#2C3E50;'>Total: {fmt_clp(deuda_neta)}</h4>", unsafe_allow_html=True)
                    st.write("")
                    
                    r1, r2, r3 = st.columns(3)
                    ef = safe_int(r1.number_input("Efectivo", value=val_gui(fin['efec']), step=1000, disabled=not permiso_editar))
                    tr = safe_int(r2.number_input("Transferencias", value=val_gui(fin['trans']), step=1000, disabled=not permiso_editar))
                    pc = safe_int(r3.number_input("Centralizado", value=val_gui(tot_cent_c), disabled=True, help="Este valor proviene de la pestaña Pagos"))
                    
                    saldo = deuda_neta - val_db(ef) - val_db(tr) - pc
                    
                    st.write("")
                    c_sem, c_btn = st.columns([3, 1], vertical_alignment="center")
                    
                    estilo_base = "margin: 0; height: 42px; display: flex; align-items: center; justify-content: center; border-radius: 6px; font-weight: 600; font-size: 15px;"
                    if saldo == 0: 
                        c_sem.markdown(f"<div style='{estilo_base} background-color:#F0FDF4; color:#166534; border: 1px solid #BBF7D0;'>Cuadratura Exacta</div>", unsafe_allow_html=True)
                    elif saldo > 0: 
                        c_sem.markdown(f"<div style='{estilo_base} background-color:#FEF2F2; color:#991B1B; border: 1px solid #FECACA;'>Faltan {fmt_clp(saldo)}</div>", unsafe_allow_html=True)
                    else: 
                        c_sem.markdown(f"<div style='{estilo_base} background-color:#FFFBEB; color:#92400E; border: 1px solid #FDE68A;'>Sobran {fmt_clp(abs(saldo))}</div>", unsafe_allow_html=True)
                    
                    if permiso_editar and c_btn.button("Guardar", type="primary", use_container_width=True, key="btn_save_ren_c"):
                        d = {"venta": fin['venta_real'], "recaudo_total": tot_ing, "comision": comision, "bencina": bencina, "sueldo": val_db(sueldo), "otros": t_om_c, "det": txt_om_c, "efec": val_db(ef), "trans": val_db(tr)}
                        save_finanzas_corriente(f_cob, id_vc, d)
                        st.success("Guardado"); time.sleep(0.5); st.rerun()

        # --- PESTAÑA 3: REPORTE ---
        with tab_rep:
            st.write("")
            st.markdown(f"<h5 style='display:flex; align-items:center;'>{svg_chart} Estado General</h5>", unsafe_allow_html=True)
            
            col_r1, col_r2, col_r3, col_r4 = st.columns([2, 2, 3, 2], vertical_alignment="bottom")
            fi = col_r1.date_input("Desde", date.today() - timedelta(days=7), key="r_ini_c", format="DD/MM/YYYY")
            ff = col_r2.date_input("Hasta", date.today(), key="r_fin_c", format="DD/MM/YYYY")
            
            # --- SELECTOR INTELIGENTE DE AISLAMIENTO ---
            if not es_repartidor:
                filtro_rep_c = col_r3.selectbox("Filtrar Repartidor", ["Todos"] + vendedores_corriente, index=0, key="filt_rep_rep_c")
            else:
                filtro_rep_c = col_r3.selectbox("Repartidor", [nombre_vendedor_actual], disabled=True, key="filt_rep_rep_c")
            
            if col_r4.button("Generar Reporte", type="primary", use_container_width=True):
                df_res = get_resumen_visor_corriente(fi, ff)
                
                if not df_res.empty:
                    # --- MOTOR DE FILTRADO ---
                    if es_repartidor:
                        df_res = df_res[df_res['Vendedor'] == nombre_vendedor_actual]
                    elif filtro_rep_c != "Todos":
                        df_res = df_res[df_res['Vendedor'] == filtro_rep_c]
                        
                    if not df_res.empty:
                        # ORDEN CRONOLÓGICO
                        if 'fecha' in df_res.columns.str.lower():
                            col_fecha = 'Fecha' if 'Fecha' in df_res.columns else 'fecha'
                            df_res[col_fecha] = pd.to_datetime(df_res[col_fecha], dayfirst=True, errors='ignore')
                            df_res = df_res.sort_values(by=[col_fecha, 'Vendedor'], ascending=[False, True])
                            df_res[col_fecha] = df_res[col_fecha].dt.strftime('%d/%m/%Y')
                            
                        def color_saldo(val):
                            if val == 0: return 'background-color: #F0FDF4; color: #166534' 
                            elif val > 0: return 'background-color: #FEF2F2; color: #991B1B'
                            return 'background-color: #FFFBEB; color: #92400E'
                        
                        st.write("")
                        st.dataframe(
                            df_res.style.map(color_saldo, subset=['saldo_final']).format({"total_gastos": "$ {:,.0f}", "saldo_final": "$ {:,.0f}", "saldo_clientes": "$ {:,.0f}"}), 
                            use_container_width=True, hide_index=True, 
                            column_config={"saldo_final": "Deuda Neta Repartidor", "saldo_clientes": "Deuda de Clientes (Calle)"}
                        )
                        
                        st.divider()
                        m1, m2 = st.columns(2)
                        m1.metric("Deuda Acumulada Repartidor", fmt_clp(df_res['saldo_final'].sum()))
                        m2.metric("Deuda Acumulada Clientes (Calle)", fmt_clp(df_res['saldo_clientes'].sum()))
                    else:
                        st.info("No hay registros para la selección en este rango de fechas.")
                else:
                    st.info("Sin registros financieros en este período.")

    elif seleccion == "Clientes":
        st.title("Gestión de Clientes")
        
        svg_people = '''<svg xmlns="http://www.w3.org/2000/svg" width="22" height="22" fill="#556B2F" class="bi bi-people" viewBox="0 0 16 16" style="margin-right: 8px; margin-bottom: 4px;"><path d="M15 14s1 0 1-1-1-4-5-4-5 3-5 4 1 1 1 1h8Zm-7.978-1A.261.261 0 0 1 7 12.996c.001-.264.167-1.03.76-1.72C8.312 10.629 9.282 10 11 10c1.717 0 2.687.63 3.24 1.276.593.69.758 1.457.76 1.72l-.008.002a.274.274 0 0 1-.014.002H7.022ZM11 7a2 2 0 1 0 0-4 2 2 0 0 0 0 4Zm3-2a3 3 0 1 1-6 0 3 3 0 0 1 6 0ZM6.936 9.28a5.88 5.88 0 0 0-1.23-.247A7.35 7.35 0 0 0 5 9c-4 0-5 3-5 4 0 .667.333 1 1 1h4.216A2.238 2.238 0 0 1 5 13c0-1.01.377-2.042 1.09-2.904.243-.294.526-.569.846-.816ZM4.92 10A5.493 5.493 0 0 0 4 13H1c0-.26.164-1.03.76-1.724.545-.636 1.492-1.256 3.16-1.275ZM1.5 5.5a3 3 0 1 1 6 0 3 3 0 0 1-6 0Zm3-2a2 2 0 1 0 0 4 2 2 0 0 0 0-4Z"/></svg>'''
        svg_inbox = '''<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" fill="#556B2F" class="bi bi-inbox" viewBox="0 0 16 16" style="margin-right: 8px; margin-bottom: 2px;"><path d="M4.98 4a.5.5 0 0 0-.39.188L1.54 8H6a.5.5 0 0 1 .5.5 1.5 1.5 0 1 0 3 0A.5.5 0 0 1 10 8h4.46l-3.05-3.812A.5.5 0 0 0 11.02 4H4.98zm9.954 5H10.45a2.5 2.5 0 0 1-4.9 0H1.066l.32 2.562a.5.5 0 0 0 .497.438h12.234a.5.5 0 0 0 .496-.438L14.933 9zM3.809 3.563A1.5 1.5 0 0 1 4.981 3h6.038a1.5 1.5 0 0 1 1.172.563l3.7 4.625a.5.5 0 0 1 .109.273l.94 7.514A1.5 1.5 0 0 1 15.446 15H.554a1.5 1.5 0 0 1-1.493-1.025l.94-7.514a.5.5 0 0 1 .108-.273l3.7-4.625z"/></svg>'''
        svg_list = '''<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" fill="#556B2F" class="bi bi-card-list" viewBox="0 0 16 16" style="margin-right: 8px; margin-bottom: 2px;"><path d="M14.5 3a.5.5 0 0 1 .5.5v9a.5.5 0 0 1-.5.5h-13a.5.5 0 0 1-.5-.5v-9a.5.5 0 0 1 .5-.5h13zm-13-1A1.5 1.5 0 0 0 0 3.5v9A1.5 1.5 0 0 0 1.5 14h13a1.5 1.5 0 0 0 1.5-1.5v-9A1.5 1.5 0 0 0 14.5 2h-13z"/><path d="M5 8a.5.5 0 0 1 .5-.5h7a.5.5 0 0 1 0 1h-7A.5.5 0 0 1 5 8zm0-2.5a.5.5 0 0 1 .5-.5h7a.5.5 0 0 1 0 1h-7a.5.5 0 0 1-.5-.5zm0 5a.5.5 0 0 1 .5-.5h7a.5.5 0 0 1 0 1h-7a.5.5 0 0 1-.5-.5zm-1-5a.5.5 0 1 1-1 0 .5.5 0 0 1 1 0zM4 8a.5.5 0 1 1-1 0 .5.5 0 0 1 1 0zm0 2.5a.5.5 0 1 1-1 0 .5.5 0 0 1 1 0z"/></svg>'''

        es_aprobador = rol in ["admin", "supervisor"]
        
        tab_dir, tab_deudas, tab_sug = st.tabs(["Directorio", "Deudas", "Sugerencias"])
        
        df_cli = obtener_clientes_corriente()
        
        # --- PESTAÑA 1: DIRECTORIO ---
        with tab_dir:
            st.write("")
            c_head, c_tog = st.columns([3, 1], vertical_alignment="bottom")
            c_head.markdown(f"<h5 style='display:flex; align-items:center; margin-bottom: 0;'>{svg_people} Gestión de Cartera</h5>", unsafe_allow_html=True)
            
            texto_toggle = "Nuevo Cliente" if permiso_editar else "Sugerir Nuevo"
            modo_crear = c_tog.toggle(texto_toggle, key="tgl_crear_cli_c")

            if modo_crear:
                st.divider()
                with st.form("new_cli_corr", border=False):
                    n1, n2, n3 = st.columns(3)
                    n_nom = n1.text_input("Nombre Cliente")
                    n_dir = n2.text_input("Dirección", value="-")
                    n_tel = n3.text_input("Teléfono", value="-")
                    
                    n4, n5, n6 = st.columns(3)
                    n_com = n4.selectbox("Comuna", l_comunas)
                    n_rep = n5.selectbox("Repartidor", vendedores_corriente, index=idx_defecto, placeholder="Seleccione un repartidor...", disabled=es_repartidor)
                    n_pre = n6.number_input("Precio Pactado", value=1600, step=50)
                    
                    comentario = ""
                    if not permiso_editar:
                        comentario = st.text_input("Justificación (Breve)", placeholder="Ej: Nuevo almacén en la ruta...")
                    
                    st.write("")
                    texto_btn = "Guardar Cliente" if permiso_editar else "Enviar Sugerencia"
                    if st.form_submit_button(texto_btn, type="primary"):
                        if n_nom and n_rep:
                            if permiso_editar:
                                crud_cliente_corriente("crear", {"nombre": n_nom, "dir": n_dir, "com": n_com, "tel": n_tel, "id_vendedor": dict_vend[n_rep], "precio": n_pre})
                                st.success("Creado"); time.sleep(0.5); st.rerun()
                            else:
                                crud_sugerencia_corriente("crear", {"tipo":"NUEVO", "nombre":n_nom, "dir":n_dir, "com":n_com, "tel":n_tel, "id_vendedor":dict_vend[n_rep], "precio":n_pre, "comentario":comentario})
                                st.success("Sugerencia enviada a revisión."); time.sleep(1.5); st.rerun()
                        else:
                            st.error("El nombre y el repartidor son obligatorios.")
            else:
                st.write("")
                with st.container(border=True):
                    fc1, fc2, fc3 = st.columns(3)
                    filtro_cli = fc1.text_input("Buscar Cliente", placeholder="Buscar por nombre...")
                    
                    idx_filtro_rep = (vendedores_corriente.index(nombre_vendedor_actual) + 1) if es_repartidor and nombre_vendedor_actual in vendedores_corriente else 0
                    filtro_rep = fc2.selectbox("Repartidor", ["Todos"] + vendedores_corriente, index=idx_filtro_rep, disabled=es_repartidor)
                    
                    mostrar_inactivos = fc3.toggle("Mostrar Inactivos")
                    
                    df_show = df_cli.copy()
                    if not mostrar_inactivos: df_show = df_show[df_show['activo'] == 1]
                    if filtro_cli: df_show = df_show[df_show['nombre'].str.contains(filtro_cli, case=False, na=False)]
                    if filtro_rep != "Todos": df_show = df_show[df_show['Repartidor'] == filtro_rep]

                    if not df_show.empty:
                        df_show['Estado'] = df_show['activo'].map({1: '🟢 Activo', 0: '🔴 Inactivo'})
                    else:
                        df_show['Estado'] = []

                    st.write("")
                    def row_style(row):
                        if row['activo'] == 0: return ['color: #999; font-style: italic'] * len(row)
                        return [''] * len(row)
                    
                    event = st.dataframe(df_show.style.apply(row_style, axis=1), 
                                         column_config={
                                             "id": None, "id_vendedor": None, "activo": None, 
                                             "nombre": st.column_config.TextColumn("Cliente", width="medium"), 
                                             "direccion": st.column_config.TextColumn("Dirección"),
                                             "comuna": st.column_config.TextColumn("Comuna"),
                                             "Repartidor": st.column_config.TextColumn("Repartidor", width="medium"), 
                                             "precio_pactado": st.column_config.NumberColumn("Precio", format="$ %d"),
                                             "Estado": st.column_config.TextColumn("Estado")
                                         }, 
                                         hide_index=True, use_container_width=True, selection_mode="single-row", on_select="rerun")
                    
                    if len(event.selection.rows) > 0:
                        idx_sel = event.selection.rows[0]; row = df_show.iloc[idx_sel]
                        st.divider()
                        st.markdown(f"##### Editar: {row['nombre']}")
                        with st.form("edit_cli_corr", border=False):
                            e1, e2, e3 = st.columns(3)
                            e_nom = e1.text_input("Nombre", value=row['nombre'])
                            e_dir = e2.text_input("Dirección", value=row['direccion'] if row['direccion'] else "-")
                            e_tel = e3.text_input("Teléfono", value=row['telefono'] if row['telefono'] else "-")
                            
                            e4, e5, e6 = st.columns(3)
                            idx_c = l_comunas.index(row['comuna']) if row['comuna'] in l_comunas else 0
                            e_com = e4.selectbox("Comuna", l_comunas, index=idx_c)
                            
                            idx_r = vendedores_corriente.index(row['Repartidor']) if row['Repartidor'] in vendedores_corriente else None
                            e_rep = e5.selectbox("Repartidor", vendedores_corriente, index=idx_r, placeholder="Seleccione...", disabled=es_repartidor)
                            e_pre = e6.number_input("Precio", value=row['precio_pactado'], step=50)
                            
                            e_act = st.toggle("Cliente Activo", value=bool(row['activo']), help="Desmarcar para dar de baja.", disabled=not permiso_editar)
                            
                            comentario_ed = ""
                            if not permiso_editar:
                                comentario_ed = st.text_input("Justificación del Cambio", placeholder="Ej: Cambió el precio, cambió de dueño...")
                            
                            st.write("")
                            texto_btn_ed = "Guardar Cambios" if permiso_editar else "Sugerir Cambio"
                            if st.form_submit_button(texto_btn_ed, type="primary"):
                                if e_rep:
                                    if permiso_editar:
                                        crud_cliente_corriente("editar", {"id": int(row['id']), "nombre": e_nom, "dir": e_dir, "com": e_com, "tel": e_tel, "id_vendedor": dict_vend[e_rep], "precio": e_pre, "activo": e_act})
                                        st.success("Actualizado"); time.sleep(0.5); st.rerun()
                                    else:
                                        crud_sugerencia_corriente("crear", {"tipo":"EDICION", "id_ref": int(row['id']), "nombre":e_nom, "dir":e_dir, "com":e_com, "tel":e_tel, "id_vendedor":dict_vend[e_rep], "precio":e_pre, "comentario":comentario_ed})
                                        st.success("Sugerencia enviada a revisión."); time.sleep(1.5); st.rerun()
                                else:
                                    st.error("Debe seleccionar un repartidor válido.")

        # --- PESTAÑA 2: DEUDAS (CALLE) ---
        with tab_deudas:
            st.write("")
            st.markdown(f"<h5 style='display:flex; align-items:center;'>{svg_list} Cartera de Deudores (Activos)</h5>", unsafe_allow_html=True)
            c_d1, c_d2, c_d3 = st.columns([2, 1, 2], vertical_alignment="bottom")
            
            v_deuda = c_d1.selectbox("Seleccionar Repartidor", vendedores_corriente, index=idx_defecto, placeholder="Seleccione un repartidor...", disabled=es_repartidor, key="sel_rep_deuda_c")
            
            if c_d2.button("Calcular Deudas", type="primary", use_container_width=True):
                if v_deuda:
                    df_deudas = obtener_deudas_corriente_vendedor(dict_vend[v_deuda])
                    if not df_deudas.empty:
                        st.write("")
                        deuda_total = df_deudas['Deuda Actual'].sum()
                        st.metric("Deuda Total en Calle", fmt_clp(deuda_total), delta_color="inverse")
                        
                        def color_deuda(val):
                            if val > 0: return 'color: #991B1B; font-weight: bold'
                            elif val < 0: return 'color: #166534'
                            return ''
                            
                        st.dataframe(df_deudas.style.map(color_deuda, subset=['Deuda Actual']).format({"Deuda Actual": "$ {:,.0f}"}), use_container_width=True, hide_index=True)
                    else:
                        st.info("Todos los clientes activos están al día o no tienen historial.")

        # --- PESTAÑA 3: BANDEJA DE SUGERENCIAS ---
        with tab_sug:
            st.write("")
            st.markdown(f"<h5 style='display:flex; align-items:center;'>{svg_inbox} Bandeja de Aprobación</h5>", unsafe_allow_html=True)
            
            if es_aprobador:
                df_sug = get_sugerencias(solo_pendientes=True, modulo="CORRIENTE")
                if not df_sug.empty:
                    ev_sug = st.dataframe(
                        df_sug, use_container_width=True, hide_index=True, selection_mode="single-row", on_select="rerun",
                        column_config={"id":None, "id_cliente_ref":None, "id_vendedor":None, "estado":None, "tipo_cliente":None, "comentario": "Justificación", "tipo_solicitud": "Solicitud", "nombre": "Cliente", "direccion":"Dir", "comuna":"Comuna", "telefono":"Tel", "Repartidor":"Repartidor", "fecha":"Fecha"}
                    )
                    
                    if len(ev_sug.selection.rows) > 0:
                        s_row = df_sug.iloc[ev_sug.selection.rows[0]]
                        st.divider()
                        
                        tipo_txt = "nuevo cliente" if s_row['tipo_solicitud'] == 'NUEVO' else "edición de cliente"
                        st.markdown(f"**Evaluar {tipo_txt}** | Propuesto por: {s_row['Repartidor']}")
                        st.info(f"Justificación y Detalles: {s_row['comentario']}")
                        
                        ca, cr = st.columns(2)
                        if ca.button("Aprobar e integrar", type="primary", use_container_width=True, key="btn_aprov_c"):
                            crud_sugerencia_corriente("APROBADA", s_row.to_dict(), id_sug=int(s_row['id']))
                            st.success("Integrado."); time.sleep(0.5); st.rerun()
                            
                        if cr.button("Rechazar", use_container_width=True, key="btn_rej_c"):
                            crud_sugerencia_corriente("RECHAZADA", id_sug=int(s_row['id']))
                            st.error("Rechazado."); time.sleep(0.5); st.rerun()
                else:
                    st.success("Bandeja limpia.")
            
            else:
                df_mis_sug = get_sugerencias(solo_pendientes=False, id_vend=mi_id_vendedor, modulo="CORRIENTE")
                if not df_mis_sug.empty:
                    def color_sug(val):
                        if val == 'PENDIENTE': return 'color: #92400E; font-weight: bold' 
                        if val == 'APROBADA': return 'color: #166534; font-weight: bold'  
                        if val == 'RECHAZADA': return 'color: #991B1B; font-weight: bold' 
                        return ''
                        
                    st.dataframe(
                        df_mis_sug.style.map(color_sug, subset=['estado']),
                        use_container_width=True, hide_index=True,
                        column_config={"id":None, "id_cliente_ref":None, "id_vendedor":None, "Repartidor":None, "tipo_cliente":None, "comentario": "Justificación", "tipo_solicitud": "Solicitud", "nombre": "Cliente", "estado": "Estado", "fecha": "Fecha"}
                    )
                else:
                    st.info("Sin sugerencias enviadas.")

    elif seleccion == "Caja":
        st.title("Libro de Caja")
        
        svg_cash = '''<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" fill="#556B2F" class="bi bi-cash-stack" viewBox="0 0 16 16" style="margin-right: 8px; margin-bottom: 4px;"><path d="M1 3a1 1 0 0 1 1-1h12a1 1 0 0 1 1 1H1zm7 8a2 2 0 1 0 0-4 2 2 0 0 0 0 4z"/><path d="M0 5a1 1 0 0 1 1-1h14a1 1 0 0 1 1 1v8a1 1 0 0 1-1 1H1a1 1 0 0 1-1-1V5zm3 0a2 2 0 0 1-2 2v4a2 2 0 0 1 2 2h10a2 2 0 0 1 2-2V7a2 2 0 0 1-2-2H3z"/></svg>'''
        svg_book = '''<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" fill="#556B2F" class="bi bi-journal-text" viewBox="0 0 16 16" style="margin-right: 8px; margin-bottom: 4px;"><path d="M5 10.5a.5.5 0 0 1 .5-.5h2a.5.5 0 0 1 0 1h-2a.5.5 0 0 1-.5-.5zm0-2a.5.5 0 0 1 .5-.5h5a.5.5 0 0 1 0 1h-5a.5.5 0 0 1-.5-.5zm0-2a.5.5 0 0 1 .5-.5h5a.5.5 0 0 1 0 1h-5a.5.5 0 0 1-.5-.5zm0-2a.5.5 0 0 1 .5-.5h5a.5.5 0 0 1 0 1h-5a.5.5 0 0 1-.5-.5z"/><path d="M3 0h10a2 2 0 0 1 2 2v12a2 2 0 0 1-2 2H3a2 2 0 0 1-2-2v-1h1v1a1 1 0 0 0 1 1h10a1 1 0 0 0 1-1V2a1 1 0 0 0-1-1H3a1 1 0 0 0-1 1v1H1V2a2 2 0 0 1 2-2z"/><path d="M1 5v-.5a.5.5 0 0 1 1 0V5h.5a.5.5 0 0 1 0 1h-2a.5.5 0 0 1 0-1H1zm0 3v-.5a.5.5 0 0 1 1 0V8h.5a.5.5 0 0 1 0 1h-2a.5.5 0 0 1 0-1H1zm0 3v-.5a.5.5 0 0 1 1 0v.5h.5a.5.5 0 0 1 0 1h-2a.5.5 0 0 1 0-1H1z"/></svg>'''
        svg_gear = '''<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" fill="#556B2F" class="bi bi-gear" viewBox="0 0 16 16" style="margin-right: 8px; margin-bottom: 2px;"><path d="M8 4.754a3.246 3.246 0 1 0 0 6.492 3.246 3.246 0 0 0 0-6.492zM5.754 8a2.246 2.246 0 1 1 4.492 0 2.246 2.246 0 0 1-4.492 0z"/><path d="M9.796 1.343c-.527-1.79-3.065-1.79-3.592 0l-.094.319a.873.873 0 0 1-1.255.52l-.292-.16c-1.64-.892-3.433.902-2.54 2.541l.159.292a.873.873 0 0 1-.52 1.255l-.319.094c-1.79.527-1.79 3.065 0 3.592l.319.094a.873.873 0 0 1 .52 1.255l-.16.292c-.892 1.64.901 3.434 2.541 2.54l.292-.159a.873.873 0 0 1 1.255.52l.094.319c.527 1.79 3.065 1.79 3.592 0l.094-.319a.873.873 0 0 1 1.255-.52l.292.16c1.64.893 3.434-.902 2.54-2.541l-.159-.292a.873.873 0 0 1 .52-1.255l.319-.094c1.79-.527 1.79-3.065 0-3.592l-.319-.094a.873.873 0 0 1-.52-1.255l.16-.292c.893-1.64-.902-3.433-2.541-2.54l-.292.159a.873.873 0 0 1-1.255-.52l-.094-.319zm-2.633.283c.246-.835 1.428-.835 1.674 0l.094.319a1.873 1.873 0 0 0 2.693 1.115l.291-.16c.764-.415 1.6.42 1.184 1.185l-.159.292a1.873 1.873 0 0 0 1.116 2.692l.318.094c.835.246.835 1.428 0 1.674l-.319.094a1.873 1.873 0 0 0-1.115 2.693l.16.291c.415.764-.42 1.6-1.185 1.184l-.291-.159a1.873 1.873 0 0 0-2.693 1.116l-.094.318c-.246.835-1.428.835-1.674 0l-.094-.319a1.873 1.873 0 0 0-2.692-1.115l-.292.16c-.764.415-1.6-.42-1.184-1.185l.159-.291A1.873 1.873 0 0 0 1.945 8.93l-.319-.094c-.835-.246-.835-1.428 0-1.674l.319-.094A1.873 1.873 0 0 0 3.06 4.377l-.16-.292c-.415-.764.42-1.6 1.185-1.184l.292.159a1.873 1.873 0 0 0 2.692-1.115l.094-.319z"/></svg>'''

        es_jefatura = rol in ["admin", "supervisor"]
        
        c_f1, _ = st.columns([1, 4])
        fecha_caja = c_f1.date_input("Fecha", date.today(), key="f_caja_corr_global", format="DD/MM/YYYY")
            
        tabs_caja = st.tabs(["Operación", "Registro", "Ajustes"]) if not es_repartidor else st.tabs(["Consulta"])
        
        # --- PESTAÑA 1: OPERACIÓN ---
        with tabs_caja[0]:
            if permiso_editar:
                st.write("")
                
                # El interruptor estético
                c_head, c_tog = st.columns([3, 1], vertical_alignment="bottom")
                c_head.markdown(f"<h5 style='display:flex; align-items:center; margin-bottom: 0;'>{svg_cash} Nuevo Movimiento</h5>", unsafe_allow_html=True)
                
                modo_correccion = c_tog.toggle("Corregir un registro", key="tgl_corr_caja_c")
                
                df_cat = obtener_categorias_caja()
                df_entidades = obtener_entidades_caja("Corriente") 

                with st.container(border=True):
                    # ESTADO 1: FORMULARIO DE CORRECCIÓN
                    if modo_correccion:
                        df_editables = obtener_caja_editables("Pan Corriente", es_jefatura)
                        
                        if not df_editables.empty:
                            st.info("Mostrando registros editables (Candado Temporal).")
                            
                            opciones_edicion = {}
                            for _, r in df_editables.iterrows():
                                fecha_str = r['fecha'].strftime("%d/%m/%Y")
                                tipo_monto = ""
                                if r['ingreso_efectivo'] > 0: tipo_monto = f"+ $ {int(r['ingreso_efectivo'])} (Efec)"
                                elif r['ingreso_transferencia'] > 0: tipo_monto = f"+ $ {int(r['ingreso_transferencia'])} (Trans)"
                                elif r['egreso'] > 0: tipo_monto = f"- $ {int(r['egreso'])} (Efec)"
                                elif r['egreso_transferencia'] > 0: tipo_monto = f"- $ {int(r['egreso_transferencia'])} (Trans)"
                                
                                eq = f"{fecha_str} | {r['descripcion']} | {r['item']} | {tipo_monto}"
                                opciones_edicion[eq] = r
                                
                            sel_reg = st.selectbox("Seleccione movimiento a corregir", list(opciones_edicion.keys()), index=None, placeholder="Buscar registro...", key="sel_ed_caja_c")
                            
                            if sel_reg:
                                row_ed = opciones_edicion[sel_reg]
                                st.divider()
                                
                                c1, c2 = st.columns(2)
                                
                                idx_cat = df_cat[df_cat['id'] == row_ed['id_categoria']].index[0] if pd.notna(row_ed['id_categoria']) and not df_cat.empty else 0
                                cat_sel = c1.selectbox("Categoría", df_cat['nombre'].tolist() if not df_cat.empty else ["-"], index=int(idx_cat), key="ed_cat_c")
                                
                                if cat_sel:
                                    id_cat = int(df_cat[df_cat['nombre'] == cat_sel]['id'].values[0])
                                    df_sub = obtener_subcategorias_caja(id_cat)
                                    lista_sub = df_sub['nombre'].tolist() if not df_sub.empty else ["-"]
                                    
                                    idx_sub = 0
                                    if pd.notna(row_ed['id_subcategoria']) and not df_sub.empty:
                                        match_sub = df_sub[df_sub['id'] == row_ed['id_subcategoria']]
                                        if not match_sub.empty: idx_sub = match_sub.index[0]
                                        
                                    sub_sel = c2.selectbox("Subcategoría", lista_sub, index=int(idx_sub), key="ed_subcat_c")
                                    
                                    c3, c4 = st.columns(2)
                                    lista_ent = df_entidades['nombre'].tolist() if not df_entidades.empty else ["-"]
                                    
                                    idx_ent = 0
                                    if pd.notna(row_ed['id_entidad']) and not df_entidades.empty:
                                        match_ent = df_entidades[df_entidades['id'] == row_ed['id_entidad']]
                                        if not match_ent.empty: idx_ent = match_ent.index[0]
                                        
                                    ent_sel = c3.selectbox("Entidad", lista_ent, index=int(idx_ent), key="ed_ent_c")
                                    det = c4.text_input("Descripción", value=str(row_ed['item']) if pd.notna(row_ed['item']) else "", key="ed_det_c")
                                    
                                    c5, c6, c7 = st.columns(3, vertical_alignment="bottom")
                                    
                                    es_ingreso = (row_ed['ingreso_efectivo'] > 0) or (row_ed['ingreso_transferencia'] > 0)
                                    mov_dir_idx = 0 if es_ingreso else 1
                                    mov_dir = c5.selectbox("Movimiento", ["Ingreso", "Egreso"], index=mov_dir_idx, key="ed_mov_dir_c")
                                    
                                    es_efectivo = (row_ed['ingreso_efectivo'] > 0) or (row_ed['egreso'] > 0)
                                    mov_met_idx = 0 if es_efectivo else 1 
                                    mov_met = c6.selectbox("Método", ["Efectivo", "Transferencia", "Cheque", "Depósito", "Otro"], index=mov_met_idx, key="ed_mov_met_c")
                                    
                                    monto_actual = max(row_ed['ingreso_efectivo'], row_ed['ingreso_transferencia'], row_ed['egreso'], row_ed['egreso_transferencia'])
                                    monto = c7.number_input("Monto ($)", min_value=0, value=int(monto_actual), step=1000, key="ed_num_monto_c")
                                    
                                    st.write("")
                                    col_izq, _ = st.columns([1, 4])
                                    if col_izq.button("Actualizar Registro", type="primary", use_container_width=True, key="ed_btn_upd_c"):
                                        if sub_sel and ent_sel and monto and monto > 0:
                                            id_sub_val = int(df_sub[df_sub['nombre']==sub_sel]['id'].values[0]) if not df_sub.empty else None
                                            id_ent_val = int(df_entidades[df_entidades['nombre']==ent_sel]['id'].values[0])
                                            
                                            ie = it = ee = et = 0
                                            if mov_dir == "Ingreso":
                                                if mov_met == "Efectivo": ie = monto
                                                else: it = monto 
                                            else:
                                                if mov_met == "Efectivo": ee = monto
                                                else: et = monto 
                                            
                                            editar_movimiento_caja_mill(row_ed['id'], fecha_caja, id_cat, id_sub_val, id_ent_val, ent_sel, det, ie, it, ee, et, st.session_state.user_name)
                                            st.success("Registro corregido.")
                                            time.sleep(0.5)
                                            st.rerun()
                                        else:
                                            st.error("Campos inválidos.")
                        else:
                            st.warning("No hay registros recientes habilitados para su edición.")

                    # ESTADO 2: FORMULARIO DE INGRESO NORMAL
                    else:
                        c1, c2 = st.columns(2)
                        lista_cat = df_cat['nombre'].tolist() if not df_cat.empty else ["-"]
                        cat_sel = c1.selectbox("Categoría", lista_cat, index=None, placeholder="Seleccione una categoría...", key="sel_cat_c")
                        
                        if cat_sel:
                            id_cat = int(df_cat[df_cat['nombre'] == cat_sel]['id'].values[0])
                            df_sub = obtener_subcategorias_caja(id_cat)
                            lista_sub = df_sub['nombre'].tolist() if not df_sub.empty else ["-"]
                            sub_sel = c2.selectbox("Subcategoría", lista_sub, index=None, placeholder="Seleccione una subcategoría...", key="sel_subcat_c")
                            
                            c3, c4 = st.columns(2)
                            lista_ent = df_entidades['nombre'].tolist() if not df_entidades.empty else ["-"]
                            ent_sel = c3.selectbox("Entidad", lista_ent, index=None, placeholder="Buscar entidad...", key="sel_ent_c")
                            det = c4.text_input("Descripción", placeholder="Ej: Pago factura, bono extra...", key="txt_det_c")
                            
                            c5, c6, c7 = st.columns(3, vertical_alignment="bottom")
                            mov_dir = c5.selectbox("Movimiento", ["Ingreso", "Egreso"], key="sel_mov_dir_c")
                            mov_met = c6.selectbox("Método", ["Efectivo", "Transferencia", "Cheque", "Depósito", "Otro"], key="sel_mov_met_c")
                            monto = c7.number_input("Monto ($)", min_value=0, value=None, step=1000, placeholder="Vacío", key="num_monto_caja_c")
                            
                            st.write("")
                            col_izq, _ = st.columns([1, 4])
                            if col_izq.button("Guardar", type="primary", use_container_width=True, key="btn_guardar_caja_c"):
                                if sub_sel and ent_sel and monto and monto > 0:
                                    id_sub_val = int(df_sub[df_sub['nombre']==sub_sel]['id'].values[0]) if not df_sub.empty else None
                                    id_ent_val = int(df_entidades[df_entidades['nombre']==ent_sel]['id'].values[0])
                                    
                                    ie = it = ee = et = 0
                                    if mov_dir == "Ingreso":
                                        if mov_met == "Efectivo": ie = monto
                                        else: it = monto 
                                    else:
                                        if mov_met == "Efectivo": ee = monto
                                        else: et = monto 
                                    
                                    guardar_movimiento_caja_mill(fecha_caja, "Pan Corriente", id_cat, id_sub_val, id_ent_val, ent_sel, det, ie, it, ee, et)
                                    st.success("Registrado.")
                                    time.sleep(0.5)
                                    st.rerun()
                                else: 
                                    st.error("Complete todos los campos requeridos (Subcategoría, Entidad y Monto válido).")

        with tabs_caja[1]:
            st.write("")
            c_reg1, c_reg2, c_reg3 = st.columns([2, 1, 1], vertical_alignment="bottom")
            c_reg1.markdown(f"<h5 style='display:flex; align-items:center; margin: 0;'>{svg_book} Auditoría Diaria</h5>", unsafe_allow_html=True)
            
            ver_libro_mayor = False
            modo_global = False
            if es_jefatura:
                # 1. Leemos primero el Global para que domine la jerarquía
                modo_global = c_reg3.toggle("Caja Mayor (Kilaco)", value=False, key="tgl_glb_e")
                
                # 2. El switch del módulo se bloquea (disabled) si el global está encendido
                ver_libro_mayor = c_reg2.toggle("Caja Mayor (Corriente)", value=True, disabled=modo_global, key="tgl_lm_e")
                
                # 3. Si vemos el global, forzamos internamente a que se vea todo sin filtros
                if modo_global: ver_libro_mayor = True
            
            if modo_global:
                df_caja = obtener_caja_mayor_global(fecha_caja)
            else:
                df_caja = obtener_caja_del_dia(fecha_caja, 'Pan Corriente')
            
            if not df_caja.empty:
                df_caja.fillna(0, inplace=True)
                
                if not es_jefatura or not ver_libro_mayor:
                    df_caja = df_caja[~df_caja['rol_creador'].isin(['admin', 'supervisor'])]
                
                if not df_caja.empty:
                    tie=df_caja['ingreso_efectivo'].sum(); tit=df_caja['ingreso_transferencia'].sum()
                    tee=df_caja['egreso_efectivo'].sum(); tet=df_caja['egreso_transferencia'].sum()
                    
                    cols_visuales = {
                        "id": None, "rol_creador": None,
                        "area": st.column_config.TextColumn("Origen", width="small") if modo_global else None,
                        "fecha": None,
                        "entidad": st.column_config.TextColumn("Entidad", width="medium"), 
                        "detalle": st.column_config.TextColumn("Descripción", width="medium"), 
                        "ingreso_efectivo": st.column_config.NumberColumn("Ingreso Efec.", format="$ %d"), 
                        "ingreso_transferencia": st.column_config.NumberColumn("Ingreso Banco", format="$ %d"), 
                        "egreso_efectivo": st.column_config.NumberColumn("Egreso Efec.", format="$ %d"),
                        "egreso_transferencia": st.column_config.NumberColumn("Egreso Banco", format="$ %d")
                    }

                    st.dataframe(df_caja, use_container_width=True, hide_index=True, column_config=cols_visuales)
                    
                    st.write("")
                    k1, k2, k3 = st.columns(3)
                    k1.metric("Ingresos (Efectivo)", fmt_clp(tie))
                    k2.metric("Egresos (Caja Chica)", fmt_clp(tee))
                    saldo_fisico = tie - tee
                    k3.metric("Saldo Físico en Cajón", fmt_clp(saldo_fisico), delta="A favor" if saldo_fisico >= 0 else "Faltante", delta_color="normal" if saldo_fisico >= 0 else "inverse")
                    
                    st.write("")
                    balance_global = (tie+tit)-(tee+tet)
                    color_bg = "#F0FDF4" if balance_global >= 0 else "#FEF2F2"
                    color_tx = "#166534" if balance_global >= 0 else "#991B1B"
                    borde = "#BBF7D0" if balance_global >= 0 else "#FECACA"
                    
                    titulo_balance = "Balance Consolidado Panadería" if modo_global else ("Balance Financiero Total" if ver_libro_mayor else "Balance Operativo")
                    st.markdown(f"<div style='margin: 0; padding: 16px; background-color: {color_bg}; border: 1px solid {borde}; border-radius: 6px; text-align: center; color: {color_tx}; font-size: 16px;'><b>{titulo_balance}:</b> {fmt_clp(balance_global)}</div>", unsafe_allow_html=True)
                else:
                    st.info("No hay movimientos en la vista actual tras aplicar filtros.")
            else: 
                st.info("Sin movimientos registrados.")

        if not es_repartidor:
            with tabs_caja[2]:
                st.write("")
                st.markdown(f"<h5 style='display:flex; align-items:center;'>{svg_gear} Entidades</h5>", unsafe_allow_html=True)
                
                modo_mantenedor = st.radio("Acción", ["Crear Nueva Entidad", "Gestionar Existente"], horizontal=True, label_visibility="collapsed", key="rad_mant_c")
                st.write("")
                df_mant = obtener_todas_entidades()
                
                if modo_mantenedor == "Crear Nueva Entidad":
                    with st.form("new_ent_form_c", clear_on_submit=True, border=False):
                        e1, e2, e3 = st.columns(3)
                        n_ent = e1.text_input("Nombre")
                        t_ent = e2.selectbox("Tipo", ["Empleado", "Proveedor", "Servicio", "Otro"])
                        a_ent = e3.selectbox("Alcance", ["Global", "Especial", "Corriente"])
                        
                        st.write("")
                        col_izq, _ = st.columns([1, 4])
                        if col_izq.form_submit_button("Guardar", type="primary", use_container_width=True):
                            if n_ent:
                                conn_ent = get_conn()
                                try:
                                    with conn_ent.cursor() as cur:
                                        cur.execute("INSERT INTO entidades (nombre, tipo, alcance) VALUES (%s, %s, %s)", (n_ent, t_ent, a_ent))
                                    conn_ent.commit()
                                    st.cache_data.clear()
                                    st.success("Guardado.")
                                    time.sleep(0.5); st.rerun()
                                except Exception as e:
                                    st.error("Error al guardar. Posible duplicado.")
                                finally:
                                    conn_ent.close()
                            else:
                                st.error("El nombre es obligatorio.")
                else:
                    if not df_mant.empty:
                        op_gest = st.selectbox("Seleccione entidad a gestionar", df_mant['nombre'].tolist(), index=None, placeholder="Buscar entidad...", key="sel_gest_c")
                        if op_gest:
                            row = df_mant[df_mant['nombre'] == op_gest].iloc[0]
                            with st.form("edit_ent_form_c", border=False):
                                e1, e2, e3 = st.columns(3)
                                en_nom = e1.text_input("Nombre", row['nombre'])
                                
                                idx_tipo = ["Empleado", "Proveedor", "Servicio", "Otro"].index(row['tipo']) if row['tipo'] in ["Empleado", "Proveedor", "Servicio", "Otro"] else 0
                                en_tip = e2.selectbox("Tipo", ["Empleado", "Proveedor", "Servicio", "Otro"], index=idx_tipo)
                                
                                alcance_db = row['alcance'].capitalize()
                                idx_alcance = ["Global", "Especial", "Corriente"].index(alcance_db) if alcance_db in ["Global", "Especial", "Corriente"] else 0
                                en_alc = e3.selectbox("Alcance", ["Global", "Especial", "Corriente"], index=idx_alcance)
                                
                                st.write("")
                                c_btn1, c_btn2, _ = st.columns([1, 1, 3])
                                if c_btn1.form_submit_button("Guardar Cambios", type="primary"):
                                    conn_ent = get_conn()
                                    try:
                                        with conn_ent.cursor() as cur:
                                            cur.execute("UPDATE entidades SET nombre=%s, tipo=%s, alcance=%s WHERE id=%s", (en_nom, en_tip, en_alc, int(row['id'])))
                                        conn_ent.commit()
                                        st.cache_data.clear()
                                        st.success("Editado.")
                                        time.sleep(0.5); st.rerun()
                                    finally:
                                        conn_ent.close()
                                        
                                if c_btn2.form_submit_button("Dar de Baja"):
                                    conn_ent = get_conn()
                                    try:
                                        with conn_ent.cursor() as cur:
                                            cur.execute("UPDATE entidades SET estado='INACTIVO' WHERE id=%s", (int(row['id']),))
                                        conn_ent.commit()
                                        st.cache_data.clear()
                                        st.error("Entidad desactivada.")
                                        time.sleep(0.5); st.rerun()
                                    finally:
                                        conn_ent.close()

    elif seleccion == "Admin" and rol == "admin":
        st.title("Administración de Sistema")
        tab_bck, tab_bd = st.tabs(["Respaldos", "Base de Datos"])
        
        with tab_bck:
            st.markdown("Descarga del sistema completo en formato Excel.")
            if st.button("Generar Backup", type="primary"):
                with st.spinner("Procesando..."):
                    b = descargar_respaldo_completo()
                    st.download_button("Descargar Archivo", b, f"Backup_{date.today()}.xlsx", "application/vnd.ms-excel")
        
        with tab_bd:
            tablas_corr = ["despacho_corriente", "clientes_corriente", "caja_movimientos", "produccion_corriente", "finanzas_corriente", "produccion_extras"]
            c1, c2 = st.columns([1, 2])
            ts = c1.selectbox("Tabla", tablas_corr, index=None, placeholder="Seleccione una tabla para visualizar...")
            
            if ts:
                has_date = ts != "clientes_corriente"
                f_del = c2.date_input("Fecha", date.today(), format="DD/MM/YYYY") if has_date else None
                
                conn = get_conn()
                try:
                    q = f"SELECT * FROM {ts}" + (" WHERE fecha = %s" if has_date else "")
                    df = pd.read_sql(q, conn, params=(f_del,) if has_date else ())
                    
                    if not df.empty:
                        df_visual = df.copy()
                        
                        # --- CORRECCIÓN ESPACIO-TEMPORAL (UTC a Santiago) ---
                        for col in df_visual.columns:
                            if pd.api.types.is_datetime64_any_dtype(df_visual[col]):
                                try:
                                    if df_visual[col].dt.tz is None:
                                        df_visual[col] = df_visual[col].dt.tz_localize('UTC').dt.tz_convert('America/Santiago')
                                    else:
                                        df_visual[col] = df_visual[col].dt.tz_convert('America/Santiago')
                                    df_visual[col] = df_visual[col].dt.strftime('%d/%m/%Y %H:%M:%S')
                                except:
                                    pass

                        inv_vend = {v: k for k, v in dict_vend.items()}
                        
                        if 'id_vendedor' in df_visual.columns:
                            df_visual['Repartidor'] = df_visual['id_vendedor'].map(inv_vend)
                        if 'id_cliente' in df_visual.columns:
                            df_cli_temp = pd.read_sql("SELECT id, nombre FROM clientes_corriente", conn)
                            df_visual['Cliente'] = df_visual['id_cliente'].map(dict(zip(df_cli_temp['id'], df_cli_temp['nombre'])))

                        st.dataframe(df_visual, use_container_width=True)
                    else:
                        st.info("No hay registros en esta tabla para los parámetros seleccionados.")
                        
                        st.divider()
                        col_id, col_dia, col_purge = st.columns(3)
                        
                        with col_id:
                            st.markdown("**Borrado Específico**")
                            opciones_borrado = {}
                            for _, r in df_visual.iterrows():
                                etiqueta = f"ID: {r['id']}"
                                if 'Repartidor' in r and pd.notna(r['Repartidor']): etiqueta += f" | Rep: {r['Repartidor']}"
                                if 'Cliente' in r and pd.notna(r['Cliente']): etiqueta += f" | Cli: {r['Cliente']}"
                                if 'nombre' in r and pd.notna(r['nombre']): etiqueta += f" | Nom: {r['nombre']}"
                                if 'venta_diaria' in r and pd.notna(r['venta_diaria']): etiqueta += f" | Vta: {r['venta_diaria']}"
                                
                                opciones_borrado[etiqueta] = r['id']
                            
                            seleccion_humana = st.selectbox("Seleccione el registro a eliminar", list(opciones_borrado.keys()), index=None, placeholder="Buscar registro...")
                            
                            if seleccion_humana:
                                id_del = opciones_borrado[seleccion_humana]
                                conf_id = st.toggle("Habilitar borrado", key=f"tgl_id_corr_{ts}")
                                if st.button("Eliminar Registro", type="primary"):
                                    if conf_id:
                                        with conn.cursor() as c: c.execute(f"DELETE FROM {ts} WHERE id=%s", (id_del,))
                                        conn.commit(); st.cache_data.clear(); st.success("Eliminado"); time.sleep(0.5); st.rerun()
                                    else: st.warning("Requiere habilitación previa.")
                        
                        with col_dia:
                            if has_date:
                                st.markdown("**Borrado Masivo**")
                                st.write(f"Afecta a todos los registros de la fecha indicada.")
                                conf_dia = st.toggle("Habilitar borrado por fecha", key=f"tgl_dia_corr_{ts}")
                                if st.button("Eliminar Día Completo", type="primary"):
                                    if conf_dia:
                                        with conn.cursor() as c: c.execute(f"DELETE FROM {ts} WHERE fecha=%s", (f_del,))
                                        conn.commit(); st.cache_data.clear(); st.success("Eliminado"); time.sleep(0.5); st.rerun()
                                    else: st.warning("Requiere habilitación previa.")
                            else:
                                st.info("El borrado por fecha no aplica a esta tabla.")
                                
                        with col_purge:
                            st.markdown("**Zona de Peligro: Purga Total**")
                            st.write("⚠️ Vacía la tabla por completo y **reinicia los IDs a 1**.")
                            txt_confirm = st.text_input(f"Escriba '{ts}' para confirmar", key=f"purge_txt_corr_{ts}")
                            if st.button("Vaciar y Reiniciar IDs", type="primary"):
                                if txt_confirm == ts:
                                    try:
                                        with conn.cursor() as c:
                                            c.execute(f"TRUNCATE TABLE {ts} RESTART IDENTITY CASCADE")
                                        conn.commit()
                                        st.cache_data.clear()
                                        st.success(f"La tabla {ts} ha sido vaciada y sus IDs reiniciados.")
                                        time.sleep(1.5); st.rerun()
                                    except Exception as e:
                                        st.error(f"Error al vaciar la tabla: {e}")
                                else:
                                    st.warning("Confirmación incorrecta. Escriba el nombre exacto.")
                except Exception as e: st.error(f"Error en consulta: {e}")
                finally: conn.close()
            

# ----------------------------------------------------
# EJECUCIÓN PRINCIPAL
# ----------------------------------------------------
# ¡La llave maestra que enciende el CSS en toda la app!
aplicar_estilos_kilaco() 

if not st.session_state.logged_in: 
    login_view()
else:
    if st.session_state.current_module == "menu": menu_view()
    elif st.session_state.current_module == "especial": app_pan_especial()
    elif st.session_state.current_module == "corriente": app_pan_corriente()

