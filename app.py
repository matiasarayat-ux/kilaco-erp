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
    """Inyecta CSS para unificar el diseño con los colores nativos de Streamlit."""
    st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600&display=swap');
        .stApp {background-color: #f4f6f9; color: #333333; font-family: 'Inter', sans-serif;}
        
        /* UI General */
        .css-card {background-color: white; padding: 20px; border-radius: 12px; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.05); margin-bottom: 20px; border: 1px solid #e0e0e0;}
        
        /* Encabezados de Tabla en Rojo Streamlit */
        div[data-testid="stDataFrame"] div[data-testid="stVerticalBlock"] div[role="columnheader"] {
            background-color: #ff4b4b !important; 
            color: white !important; 
            font-weight: 600;
        }
        th {background-color: #ff4b4b !important; color: white !important;}
        
        /* Botones en Rojo Nativo (Pasivo y Hover) */
        .stButton>button {
            background-color: #ff4b4b; 
            color: white; 
            border-radius: 8px; 
            border: none; 
            height: 3em; 
            width: 100%; 
            font-weight: 600; 
            transition: all 0.2s;
        }
        .stButton>button:hover {
            background-color: #d93333; /* Rojo más intenso/oscuro para el hover */
            color: white;
            transform: translateY(-1px);
        }
        
        .menu-card {
            background: white; padding: 30px; border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); 
            text-align: center; border: 1px solid #e0e0e0; transition: transform 0.2s; cursor: pointer; height: 100%;
        }
        .menu-card:hover {transform: translateY(-5px); border-color: #ff4b4b;}
        </style>
    """, unsafe_allow_html=True)

aplicar_estilos_kilaco()

# ==========================================
# REGIÓN 2: CONSTANTES Y UTILIDADES
# ==========================================

DB_URI = st.secrets["DB_URI"]

# Configuración de gastos fijos por repartidor
CONFIG_REPARTIDORES = {
    "Hector Silva": {"bencina": 15000, "comision": 0.04, "dscto_esp": 0.2},
    "Byron Navarro": {"bencina": 23000, "comision": 0.04, "dscto_esp": 0.2},
    "Jose Albarracin": {"bencina": 14000, "comision": 0.04, "dscto_esp": 0.2},
    "Tomas Mendez": {"bencina": 0, "comision": 0.04, "dscto_esp": 0.2}
}

def get_conn():
    """Establece la conexión a la base de datos Supabase (PostgreSQL)."""
    try:
        return psycopg2.connect(DB_URI)
    except Exception as e:
        st.error(f"Error crítico de conexión a Base de Datos: {e}")
        st.stop()

def safe_float(val):
    """Convierte un valor a float de forma segura, retornando 0.0 si falla."""
    if val is None: return 0.0
    try: return float(val)
    except: return 0.0

def safe_int(val):
    """Convierte un valor a int de forma segura, retornando 0 si falla."""
    if val is None: return 0
    try: return int(float(val))
    except: return 0

def fmt_clp(valor):
    """Formatea un número como moneda chilena ($ 1.000)."""
    return "$ 0" if valor is None else "$ " + "{:,.0f}".format(valor).replace(",", ".")

def val_gui(val):
    """Helper para inputs: devuelve None si es 0 o vacío (para UX limpia)."""
    return val if val and val > 0 else None

def val_db(val):
    """Helper para DB: convierte None a 0 para guardar en base de datos."""
    return int(val) if val else 0

def get_orden_cultural(nombre_db):
    """Devuelve el índice de ordenamiento visual para los productos."""
    prioridad = {
        "Lengua": 1, "Lengua 6": 2, "Frica": 3, "Lengua XL (25)": 4, 
        "Pizza Individual": 5, "Pizza Familiar": 6, "Hallulla": 7, "Molde": 8, 
        "Frica XL": 9, "Tapadito": 10, "Molde XL": 11, "Pan Rallado": 12, 
        "Lengua XXL (30)": 13, "Lengua XXXL (35)": 14, "Lengua XXXXL (40)": 15
    }
    for k, v in prioridad.items():
        if k.lower() == nombre_db.strip().lower(): return v
    return 99

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
                        # Excel limita nombres de hoja a 31 chars
                        df.to_excel(writer, sheet_name=tabla[:31], index=False)
                except Exception as e:
                    print(f"Advertencia respaldando {tabla}: {e}")
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
            comunas = [r[0] for r in c.fetchall()]
            c.execute("SELECT nombre FROM bancos ORDER BY nombre")
            bancos = [r[0] for r in c.fetchall()]
        
        # Filtramos vendedores "basura" o internos que no deben salir en listas generales
        query_vend = """
            SELECT id, nombre FROM vendedores 
            WHERE nombre NOT IN ('Vendedor 1', 'Vendedor 2', 'nan', 'NaN') 
            AND nombre IS NOT NULL 
            ORDER BY nombre
        """
        vendedores = pd.read_sql(query_vend, conn)
        return comunas, bancos, vendedores
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
        st.toast("Carga corregida")
    finally:
        conn.close()

def obtener_planilla(fecha, id_v):
    conn = get_conn()
    try:
        df_prod = pd.read_sql("SELECT id, nombre, precio_estandar FROM productos", conn)
        df_hoy = pd.read_sql("SELECT id as id_despacho, id_producto, saldo_anterior, carga, devolucion_muestra, saldo_actual FROM despacho WHERE fecha=%s AND id_vendedor=%s", conn, params=(fecha, id_v))
        
        # Lógica de descuento 20% si el vendedor es del 'corriente'
        with conn.cursor() as c:
            c.execute("SELECT nombre FROM vendedores WHERE id=%s", (id_v,))
            vendedor_nombre = c.fetchone()[0]
        
        config_vend = CONFIG_REPARTIDORES.get(vendedor_nombre, {})
        descuento = config_vend.get('dscto_esp', 0.0)
        if descuento > 0:
            df_prod['precio_estandar'] = df_prod['precio_estandar'] * (1 - descuento)
        
        filas = []
        with conn.cursor() as c:
            for i, prod in df_prod.iterrows():
                id_p = prod['id']
                reg_hoy = df_hoy[df_hoy['id_producto'] == id_p]
                if not reg_hoy.empty:
                    r = reg_hoy.iloc[0]
                    filas.append({"id": r['id_despacho'], "id_producto": id_p, "nombre": prod['nombre'], "precio_estandar": prod['precio_estandar'], "saldo_anterior": r['saldo_anterior'], "carga": r['carga'], "devolucion_muestra": r['devolucion_muestra'], "saldo_actual": r['saldo_actual']})
                else:
                    c.execute("SELECT saldo_actual FROM despacho WHERE id_vendedor=%s AND id_producto=%s AND fecha < %s ORDER BY fecha DESC LIMIT 1", (id_v, id_p, fecha))
                    last = c.fetchone()
                    saldo_ayer = last[0] if last else 0
                    if saldo_ayer > 0:
                        filas.append({"id": None, "id_producto": id_p, "nombre": prod['nombre'], "precio_estandar": prod['precio_estandar'], "saldo_anterior": saldo_ayer, "carga": 0, "devolucion_muestra": 0, "saldo_actual": saldo_ayer})
        
        if not filas: return pd.DataFrame()
        df = pd.DataFrame(filas)
        df['orden_sort'] = df['nombre'].apply(get_orden_cultural)
        df = df.sort_values('orden_sort')
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
            if ex:
                c.execute("UPDATE finanzas SET creditos_cobrados=%s, creditos_otorgados=%s, descuentos_total=%s, bencina=%s, sueldo=%s, otros_gastos_monto=%s, otros_gastos_detalle=%s, efectivo_rendido=%s, transferencia_rendida=%s, pago_centralizado=%s, creditos_cobrados_detalle=%s WHERE id=%s", 
                          (d['cc'],d['co'],d['ds'],d['bn'],d['su'],d['om'],d['od'],d['ef'],d['tr'],d['pc'],d['cc_det'],ex[0]))
            else:
                c.execute("INSERT INTO finanzas (fecha, id_vendedor, creditos_cobrados, creditos_otorgados, descuentos_total, bencina, sueldo, otros_gastos_monto, otros_gastos_detalle, efectivo_rendido, transferencia_rendida, pago_centralizado, creditos_cobrados_detalle) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)", 
                          (fecha, id_v, d['cc'],d['co'],d['ds'],d['bn'],d['su'],d['om'],d['od'],d['ef'],d['tr'],d['pc'],d['cc_det']))
        conn.commit()
    finally:
        conn.close()

def obtener_resumen_global(fecha_inicio, fecha_fin):
    conn = get_conn()
    try:
        df_ventas = pd.read_sql("SELECT fecha, id_vendedor, SUM(venta_unidades * p.precio_estandar) as venta_pan FROM despacho d JOIN productos p ON d.id_producto = p.id WHERE fecha BETWEEN %s AND %s GROUP BY fecha, id_vendedor", conn, params=(fecha_inicio, fecha_fin))
        df_fin = pd.read_sql("SELECT * FROM finanzas WHERE fecha BETWEEN %s AND %s", conn, params=(fecha_inicio, fecha_fin))
        df_movs = pd.read_sql("SELECT fecha, id_vendedor, SUM(CASE WHEN tipo_movimiento='ABONO' THEN monto ELSE 0 END) as abonos_creditos FROM movimientos_credito WHERE fecha BETWEEN %s AND %s GROUP BY fecha, id_vendedor", conn, params=(fecha_inicio, fecha_fin))
        
        df_full = pd.merge(df_fin, df_ventas, on=['fecha', 'id_vendedor'], how='outer')
        df_full = pd.merge(df_full, df_movs, on=['fecha', 'id_vendedor'], how='outer')
        
        for c in ['venta_pan', 'creditos_cobrados', 'abonos_creditos', 'bencina', 'sueldo', 'otros_gastos_monto', 'creditos_otorgados', 'descuentos_total', 'efectivo_rendido', 'transferencia_rendida', 'pago_centralizado']:
            if c in df_full.columns: df_full[c] = df_full[c].fillna(0).astype(float)
            else: df_full[c] = 0.0

        vends = pd.read_sql("SELECT id, nombre FROM vendedores", conn)
        v_map = dict(zip(vends['id'], vends['nombre']))
        if 'id_vendedor' in df_full.columns: df_full['Vendedor'] = df_full['id_vendedor'].map(v_map)
        
        df_full['Total Ingresos'] = df_full['venta_pan'] + df_full['creditos_cobrados'] + df_full['abonos_creditos']
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
                   v.nombre as "Repartidor", c.tipo_cliente, c.limite_credito
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
            if accion == "crear":
                c.execute("INSERT INTO clientes (nombre, direccion, comuna, telefono, id_vendedor_asignado, tipo_cliente, limite_credito) VALUES (%s,%s,%s,%s,%s,%s,%s)",
                          (datos['nombre'], datos['dir'], datos['com'], datos['tel'], datos['id_vend'], datos['tipo'], cupo_auto))
            elif accion == "editar":
                c.execute("UPDATE clientes SET nombre=%s, direccion=%s, comuna=%s, telefono=%s, id_vendedor_asignado=%s, tipo_cliente=%s, limite_credito=%s WHERE id=%s",
                          (datos['nombre'], datos['dir'], datos['com'], datos['tel'], datos['id_vend'], datos['tipo'], cupo_auto, datos['id']))
        conn.commit()
    finally:
        conn.close()

def registrar_movimiento_credito(fecha, id_cli, id_vend, tipo, monto, detalle):
    conn = get_conn()
    try:
        with conn.cursor() as c:
            c.execute("INSERT INTO movimientos_credito (fecha, id_cliente, id_vendedor, tipo_movimiento, monto, detalle) VALUES (%s,%s,%s,%s,%s,%s)",
                      (fecha, id_cli, id_vend, tipo, monto, detalle))
        conn.commit()
    finally:
        conn.close()

def registrar_transferencia(fecha, id_v, monto, metodo, banco, tipo, verif, comentario):
    conn = get_conn()
    try:
        with conn.cursor() as c:
            verif_int = 1 if verif else 0
            c.execute("INSERT INTO transferencias (fecha, id_vendedor, monto, metodo_pago, banco_emisor, verificado, tipo_transferencia, comentario) VALUES (%s,%s,%s,%s,%s,%s,%s,%s)",
                      (fecha, id_v, monto, metodo, banco, verif_int, tipo, comentario))
        conn.commit()
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
    finally:
        conn.close()

# --- Funciones Pan Corriente ---
def get_despacho_corriente(fecha, id_vendedor):
    conn = get_conn()
    try:
        query_base = """
            SELECT d.*, c.nombre as "cliente"
            FROM despacho_corriente d
            JOIN clientes_corriente c ON d.id_cliente = c.id
            WHERE d.fecha = %s AND d.id_vendedor = %s
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
    conn = get_conn()
    try:
        with conn.cursor() as c:
            for i, row in df_cambios.iterrows():
                total_kg = sum([row[f'carga_{k}'] for k in range(1,8)])
                ventas = int(total_kg * row['precio_aplicado'])
                total_pagar = ventas + row['saldo_anterior']
                deuda = total_pagar - row['paga'] - row['pago_centralizado']
                c.execute("""
                    UPDATE despacho_corriente SET 
                    carga_1=%s, carga_2=%s, carga_3=%s, carga_4=%s, carga_5=%s, carga_6=%s, carga_7=%s,
                    total_carga=%s, ventas_monto=%s, total_pagar=%s, paga=%s, pago_centralizado=%s, deuda_final=%s
                    WHERE id=%s
                """, (row['carga_1'], row['carga_2'], row['carga_3'], row['carga_4'], row['carga_5'], row['carga_6'], row['carga_7'],
                      total_kg, ventas, total_pagar, row['paga'], row['pago_centralizado'], deuda, row['id']))
        conn.commit()
    finally:
        conn.close()

def get_produccion_corriente_unificada(fecha):
    conn = get_conn()
    try:
        df_repartidores = pd.read_sql("""
            SELECT v.nombre as "Concepto",
            SUM(d.carga_1 + d.carga_2) as rinde_noche,
            SUM(d.carga_3 + d.carga_4 + d.carga_5 + d.carga_6 + d.carga_7) as rinde_dia
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
            GROUP BY d.fecha, d.id_vendedor
        """, conn, params=(fi, ff))
        
        df_full = pd.merge(df_fin, df_calle, on=['fecha', 'Vendedor'], how='outer').fillna(0)
        return df_full[['fecha', 'Vendedor', 'total_gastos', 'saldo_final', 'saldo_clientes']]
    finally:
        conn.close()

@st.cache_data(ttl=300)
def obtener_clientes_corriente():
    conn = get_conn()
    try:
        df = pd.read_sql("""
            SELECT cc.id, cc.nombre, cc.precio_pactado, v.nombre as "Repartidor", cc.id_vendedor
            FROM clientes_corriente cc
            LEFT JOIN vendedores v ON cc.id_vendedor = v.id
            WHERE cc.activo = 1 ORDER BY cc.nombre
        """, conn)
        return df
    finally:
        conn.close()

def crud_cliente_corriente(accion, datos):
    conn = get_conn()
    try:
        with conn.cursor() as c:
            if accion == "crear":
                c.execute("INSERT INTO clientes_corriente (nombre, id_vendedor, precio_pactado) VALUES (%s,%s,%s)", (datos['nombre'], datos['id_vendedor'], datos['precio']))
            elif accion == "editar":
                c.execute("UPDATE clientes_corriente SET nombre=%s, id_vendedor=%s, precio_pactado=%s WHERE id=%s", (datos['nombre'], datos['id_vendedor'], datos['precio'], datos['id']))
        conn.commit()
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
            
            c.execute("SELECT nombre FROM vendedores WHERE id=%s", (id_v,))
            res_vend = c.fetchone()
            nom_vend = res_vend[0] if res_vend else ""

        comision_auto = int(ventas_auto * 0.04)
        bencina_auto = CONFIG_REPARTIDORES.get(nom_vend, {}).get("bencina", 0)
        
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
    conn = get_conn()
    gastos = d['bencina'] + d['sueldo'] + d['otros'] + d['comision']
    saldo = d['recaudo'] - gastos - d['efec'] - d['trans']
    
    try:
        with conn.cursor() as c:
            c.execute("SELECT id FROM finanzas_corriente WHERE fecha=%s AND id_vendedor=%s", (fecha, id_v))
            ex = c.fetchone()
            if ex:
                c.execute("""UPDATE finanzas_corriente SET venta_diaria=%s, recaudo_diario=%s, comision=%s, bencina=%s, sueldo=%s, otros_gastos=%s, detalle_gastos=%s, total_gastos=%s, pago_efectivo=%s, pago_transferencia=%s, saldo_final=%s WHERE id=%s""", 
                          (d['venta'], d['recaudo'], d['comision'], d['bencina'], d['sueldo'], d['otros'], d['det'], gastos, d['efec'], d['trans'], saldo, ex[0]))
            else:
                c.execute("""INSERT INTO finanzas_corriente (fecha, id_vendedor, venta_diaria, recaudo_diario, comision, bencina, sueldo, otros_gastos, detalle_gastos, total_gastos, pago_efectivo, pago_transferencia, saldo_final) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""", 
                          (fecha, id_v, d['venta'], d['recaudo'], d['comision'], d['bencina'], d['sueldo'], d['otros'], d['det'], gastos, d['efec'], d['trans'], saldo))
        conn.commit()
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

# ==========================================
# REGIÓN 4: VISTAS Y COMPONENTES
# ==========================================

# inicialización de la sesión
init_session()

def login_view():
    """Vista de inicio de sesión conectada a Supabase Auth."""
    c1, c2, c3 = st.columns([1, 1, 1])
    with c2:
        lc1, lc2, lc3 = st.columns([1.5, 1, 1.5])
        with lc2:
            try: st.image("logo.png", use_container_width=True) 
            except: st.markdown("<h1 style='text-align: center;'>🥖</h1>", unsafe_allow_html=True)
        
        st.markdown("<h3 style='text-align: center; margin-top: 0px;'>Inicio de Sesión</h3>", unsafe_allow_html=True)
        
        with st.form("login_form"):
            # IMPORTANTE: Supabase usa Email para autenticar
            email = st.text_input("Correo Electrónico") 
            pw = st.text_input("Contraseña", type="password")
            mantener = st.checkbox("Mantener sesión iniciada")
            
            if st.form_submit_button("Entrar", type="primary"):
                if not email or not pw:
                    st.warning("Por favor ingresa correo y contraseña.")
                else:
                    # Llamamos a la nueva lógica de Supabase
                    user_data = check_login(email, pw) 
                    
                    if user_data:
                        # Guardamos TODOS los datos críticos en la sesión
                        st.session_state.logged_in = True
                        st.session_state.user_id = user_data["id"]
                        st.session_state.user_name = user_data["nombre"]
                        st.session_state.user_role = user_data["rol"]
                        st.session_state.id_vendedor = user_data["id_vendedor"] # Vital para repartidores
                        
                        if mantener: 
                            st.query_params["session"] = "active"
                        st.rerun()
                    else: 
                        st.error("Credenciales inválidas o error de conexión.")

def menu_view():
    """Menú principal con filtrado de roles."""
    contenedor_menu = st.empty()
    
    # Recuperamos el rol actual de la sesión (ej: 'admin', 'repartidor')
    rol = st.session_state.user_role
    nombre = st.session_state.user_name

    with contenedor_menu.container():
        # Encabezado personalizado
        with st.container():
            c1, c2, c3 = st.columns([1, 6, 1], vertical_alignment="center")
            with c1:
                try: st.image("logo.png", width=60)
                except: st.markdown("### 🥖")
            with c2:
                st.markdown(f"<h2 style='margin: 0; padding: 0;'>Hola, {nombre}</h2>", unsafe_allow_html=True)
                # Mostramos el rol para que sepas cómo te ve el sistema
                st.caption(f"Perfil: {str(rol).upper()}") 
            with c3:
                if st.button("Salir", key="btn_logout"):
                    # Limpieza de sesión al salir
                    contenedor_menu.empty()
                    st.session_state.logged_in = False
                    st.session_state.user_role = None
                    st.session_state.current_module = "menu"
                    st.query_params.clear()
                    st.rerun()

        st.divider()

        # --- LÓGICA DE ROLES (El Cerebro de la Seguridad) ---
        # Definimos quién puede ver qué
        ver_especial = rol in ["admin", "pan_especial", "supervisor", "repartidor_esp", "repartidor_corr"]
        ver_corriente = rol in ["admin", "pan_corriente", "supervisor", "repartidor_corr"]

        col_l, col_card1, col_card2, col_r = st.columns([1, 2, 2, 1])

        # TARJETA 1: PAN ESPECIAL (Solo Admin y Producción Especial)
        with col_card1:
            if ver_especial: 
                with st.container(border=True):
                    st.markdown("<div style='height: 10px'></div>", unsafe_allow_html=True)
                    # Icono de caja (puedes cambiarlo si no usas bootstrap icons)
                    st.markdown("""<div style="text-align: center;">📦<h3 style="margin-top: 5px;">Pan Especial</h3></div>""", unsafe_allow_html=True)
                    st.markdown("<div style='height: 15px'></div>", unsafe_allow_html=True)
                    if st.button("Ingresar", key="btn_esp", use_container_width=True, type="primary"):
                        st.session_state.current_module = "especial"
                        st.rerun()
                    st.markdown("<div style='height: 10px'></div>", unsafe_allow_html=True)
            else:
                # Si no tiene permiso, mostramos un espacio vacío o bloqueado
                st.info("🔒 Módulo restringido")

        # TARJETA 2: PAN CORRIENTE (Admin, Producción Corriente y Repartidores)
        with col_card2:
            if ver_corriente:
                with st.container(border=True):
                    st.markdown("<div style='height: 10px'></div>", unsafe_allow_html=True)
                    st.markdown("""<div style="text-align: center;">🚚<h3 style="margin-top: 5px;">Pan Corriente</h3></div>""", unsafe_allow_html=True)
                    st.markdown("<div style='height: 15px'></div>", unsafe_allow_html=True)
                    if st.button("Ingresar", key="btn_corr", use_container_width=True, type="primary"):
                        st.session_state.current_module = "corriente"
                        st.rerun()
                    st.markdown("<div style='height: 10px'></div>", unsafe_allow_html=True)
            else:
                st.info("🔒 Módulo restringido")

# ----------------------------------------------------
# APLICACIÓN PAN ESPECIAL (CORREGIDA: FILTRO VENDEDOR)
# ----------------------------------------------------
from streamlit_option_menu import option_menu

def app_pan_especial():
    # Recuperamos rol e ID de vendedor de la sesión
    rol = st.session_state.user_role
    mi_id_vendedor = st.session_state.id_vendedor
    
    # --- 1. DEFINICIÓN DE PERMISOS Y MENÚ ---
    opciones_full = ["Insumos", "Producción", "Despacho", "Cobranza", "Clientes", "Créditos", "Transferencias", "Caja"]
    iconos_full   = ["box-seam", "tools",      "truck",    "currency-dollar", "people",   "credit-card", "bank",           "cash-stack"]
    
    if rol in ["admin", "pan_especial", "supervisor"]:
        menu_options = opciones_full
        menu_icons = iconos_full
        permiso_editar = True
    
    elif rol == "repartidor_esp":
        # Repartidor Especial: Ve Despacho, Cobranza, Clientes, Créditos, Transferencias
        indices_permitidos = [2, 3, 4, 5, 6] 
        menu_options = [opciones_full[i] for i in indices_permitidos]
        menu_icons = [iconos_full[i] for i in indices_permitidos]
        permiso_editar = False # Solo lectura / Operativa limitada
        
    elif rol == "repartidor_corr":
        # Repartidor Corriente: SOLO ve Despacho en este módulo
        menu_options = ["Despacho"]
        menu_icons = ["truck"]
        permiso_editar = False 
        
    else:
        menu_options = []
        menu_icons = []
        permiso_editar = False

    if rol == "admin":
        menu_options.append("Admin")
        menu_icons.append("gear")

    # --- 2. BARRA LATERAL ---
    with st.sidebar:
        c_logo1, c_logo2, c_logo3 = st.columns([1, 1.5, 1])
        with c_logo2:
            try: st.image("logo.png", use_container_width=True)
            except: st.write("🥖")
        
        st.markdown("""
            <h3 style='text-align: center; margin-top: 0px; margin-bottom: 25px; font-weight: 600; color: #333;'>
                KILACO ERP
            </h3>
        """, unsafe_allow_html=True)
        
        if not menu_options:
            st.error("No tienes acceso a este módulo.")
            st.stop()

        seleccion = option_menu(
            menu_title=None, 
            options=menu_options,
            icons=menu_icons,
            menu_icon="cast", 
            default_index=0,
            styles={
                "container": {"padding": "0!important", "background-color": "#ffffff"},
                "icon": {"color": "#ff4b4b", "font-size": "14px"}, 
                "nav-link": {"font-size": "14px", "text-align": "left", "margin":"0px", "--hover-color": "#f0f2f6"},
                "nav-link-selected": {"background-color": "#ff4b4b", "font-weight": "600"},
            }
        )
        
        st.markdown("---")
        if st.button("Volver al Menú", use_container_width=True):
            st.session_state.current_module = "menu"
            st.rerun()

    # --- 3. CARGA DE REFERENCIAS Y FILTRO DE SEGURIDAD ---
    l_comunas, l_bancos, df_vend = get_referencias()
    
    # --- BLOQUE NUEVO: OBTENER NOMBRE DEL VENDEDOR ACTUAL ---
    nombre_vendedor_actual = None
    if mi_id_vendedor:
        # Buscamos el nombre que corresponde al ID logueado (ej: 336 -> Franco)
        match = df_vend[df_vend['id'] == mi_id_vendedor]
        if not match.empty:
            nombre_vendedor_actual = match['nombre'].iloc[0]
    # ============================================

    dict_vend = dict(zip(df_vend['nombre'], df_vend['id']))
    
    lista_descripciones = sorted([
        "Alejandro Valenzuela", "Carlos Alvarez", "Randy Galvis", "Carlos Jara", "Eduardo Sanchez",
        "Edward Baillont", "Flavio Meza", "Franco De La Puente", "Hugo Palacios", "Jose Perozo",
        "Marcelo Jara", "Robert Cordova", "Robert Mogollon", "Hector Silva", "Byron Navarro",
        "Jose Albarracin", "Tomas Mendez", "Junior Negron", "Kilaco Venta", "Otros", "Petroleo",
        "Anderson", "Aseo", "Alberto", "Juanito", "Marmix", "Osvaldo Torres", "Fabiola Tessini",
        "Ines Cordova", "Mitzy Panes", "David", "Aceite", "Alex (Carne Pato)", "Provision",
        "Yves", "Kervis", "Carlos Hornero", "Mila", "Yordys", "Gene", "Yeilid", "Michel",
        "Maestro Guillermo", "Rahab", "Mary", "Imposiciones", "Caro", "Yeimy", "Leo",
        "Jonathan", "Luis", "Jorge", "Genesis", "Guillermo", "Paola"
    ])

    # --- 4. ENRUTAMIENTO DE VISTAS ---

    # === MÓDULO: INSUMOS ===
    if seleccion == "Insumos":
        st.title("Control de Insumos")
        c1, c2 = st.columns([1, 4])
        f_ins = c1.date_input("Fecha Insumos", date.today(), format="DD/MM/YYYY")
        
        tab_bandejas, tab_bolsas = st.tabs(["Bandejas", "Bolsas"])

        with tab_bandejas:
            st.markdown("#### Control de Bandejas")
            with st.container(border=True):
                # Al filtrar df_vend arriba, este selectbox ahora solo muestra 1 opción (el usuario)
                v_ban = st.selectbox("Repartidor", df_vend['nombre'])
                id_vban = dict_vend[v_ban]
                
                data_ban = obtener_bandejas(f_ins, id_vban)
                col_b1, col_b2, col_b3, col_b4 = st.columns(4)
                ant_b = col_b1.number_input("🔒 Saldo Inicial", value=val_gui(data_ban['ant']), disabled=True)
                
                key_s = f"s_{id_vban}_{f_ins}"; key_r = f"r_{id_vban}_{f_ins}"
                val_s = data_ban['sal'] if data_ban.get('existe') else None
                val_r = data_ban['ret'] if data_ban.get('existe') else None
                
                sal_b = col_b2.number_input("Egreso", value=val_gui(val_s), placeholder="0", min_value=0, key=key_s, disabled=not permiso_editar)
                ret_b = col_b3.number_input("Retorno", value=val_gui(val_r), placeholder="0", min_value=0, key=key_r, disabled=not permiso_editar)
                
                fin_b = (data_ban['ant'] or 0) - (sal_b or 0) + (ret_b or 0)
                col_b4.metric("Saldo Final", fin_b)
                
                if permiso_editar:
                    if st.button("Guardar Bandejas"):
                        guardar_bandejas(f_ins, id_vban, data_ban['ant'], sal_b, ret_b)
                        st.toast(f"Bandejas de {v_ban} guardadas.")

        with tab_bolsas:
            st.markdown("#### Inventario de Bolsas")
            with st.container(border=True):
                df_bolsas = obtener_bolsas_manual(f_ins)
                cols_bol = {
                    "id": None, "id_cb": None, "factor": None, "gasto_cajas": None,
                    "nombre": st.column_config.TextColumn("Producto", disabled=True, width="medium"),
                    "stock_inicial_cajas": st.column_config.NumberColumn("🔒 Stock Ini (Cajas)", disabled=True, format="%.2f", width="small"),
                    "stock_inicial_bolsas": st.column_config.NumberColumn("🔒 Stock Ini (Bolsas)", disabled=True, format="%d", width="small"),
                    "ingreso_cajas": st.column_config.NumberColumn("Ingreso Cajas", step=0.01, required=True, format="%.2f", width="small"),
                    "produccion_hoy_unidades": st.column_config.NumberColumn("Producción (hoy)", step=1, required=True, width="small"),
                    "stock_cajas_final": st.column_config.NumberColumn("🔒 Total Cajas", disabled=True, format="%.2f", width="small"),
                    "stock_bolsas_final": st.column_config.NumberColumn("🔒 Total Bolsas", disabled=True, format="%d", width="small")
                }
                orden_visual = ["id", "id_cb", "factor", "gasto_cajas", "nombre", "stock_inicial_cajas", "stock_inicial_bolsas", "ingreso_cajas", "produccion_hoy_unidades", "stock_cajas_final", "stock_bolsas_final"]
                df_bolsas = df_bolsas[orden_visual]
                df_ed_bol = st.data_editor(df_bolsas, column_config=cols_bol, use_container_width=True, hide_index=True, key="ed_bolsas", height=500, disabled=not permiso_editar)
                
                if permiso_editar:
                    if st.button("Guardar Stock Bolsas"):
                        gasto = df_ed_bol['produccion_hoy_unidades'] / df_ed_bol['factor']
                        df_ed_bol['stock_cajas_final'] = df_ed_bol['stock_inicial_cajas'] + df_ed_bol['ingreso_cajas'] - gasto
                        guardar_bolsas_manual(f_ins, df_ed_bol)
                        st.success("Stock actualizado.")

    # === MÓDULO: PRODUCCIÓN ===
    elif seleccion == "Producción":
        st.title("Control de Producción")
        col_f, _ = st.columns([1, 4])
        f_st = col_f.date_input("Fecha Producción", date.today(), format="DD/MM/YYYY")
        df_st = obtener_datos_stock(f_st)
        
        orden_prod = ["id", "nombre", "stock_inicial", "fabricacion", "salida_calculada", "stock_final", "produccion_dia_siguiente", "bolsas_necesarias", "bolsas_por_saco", "cant_sacos"]
        df_st = df_st[[c for c in orden_prod if c in df_st.columns]]
        
        cols_prod = {
            "id": None, 
            "nombre": st.column_config.TextColumn("Producto", disabled=True, width="medium"),
            "stock_inicial": st.column_config.NumberColumn("🔒 Stock Inicial", disabled=True, format="%d", width="small"),
            "fabricacion": st.column_config.NumberColumn("Fabricación", required=True, width="small"),
            "salida_calculada": st.column_config.NumberColumn("🔒 Salida", disabled=True, width="small"),
            "stock_final": st.column_config.NumberColumn("🔒 Stock Final", disabled=True, width="small"),
            "produccion_dia_siguiente": st.column_config.NumberColumn("🔒 Producción Día Siguiente", disabled=True, format="%d", width="small"),
            "bolsas_necesarias": st.column_config.NumberColumn("🔒 Bolsas Necesarias", disabled=True, format="%d", width="small"),
            "bolsas_por_saco": st.column_config.NumberColumn("🔒 Bolsas por Saco", disabled=True, format="%d", width="small"),
            "cant_sacos": st.column_config.NumberColumn("🔒 Cant. Sacos Dia Siguiente", disabled=True, format="%.2f", width="small")
        }
        df_ed = st.data_editor(df_st, column_config=cols_prod, use_container_width=True, hide_index=True, key="st_editor", height=600, disabled=not permiso_editar)
        
        if permiso_editar:
            if st.button("Guardar Producción"):
                for i,r in df_ed.iterrows():
                    st_fin = r['stock_inicial'] + r['fabricacion'] - r['salida_calculada']
                    nec = max(0, r['produccion_dia_siguiente'] - st_fin)
                    registrar_produccion(f_st, r['id'], r['stock_inicial'], r['fabricacion'], st_fin, nec)
                st.success("Producción actualizada.")

    # === MÓDULO: DESPACHO ===
    elif seleccion == "Despacho":
        st.title("Despacho Repartidores")
        conn = get_conn(); df_prod = pd.read_sql("SELECT * FROM productos ORDER BY orden_visual", conn); conn.close()
        c1, c2 = st.columns([1, 2])
        f_bo = c1.date_input("Fecha Despacho", date.today(), format="DD/MM/YYYY")
        # El selectbox ahora está bloqueado en tu nombre
        v_bo = c2.selectbox("Seleccionar Repartidor", df_vend['nombre'], key="sel_desp_rep")
        id_vb = dict_vend[v_bo]
        
        with st.container(border=True):
            st.markdown("**Agregar Carga**")
            if permiso_editar:
                with st.form("add_carga", clear_on_submit=True):
                    col_f1, col_f2, col_f3 = st.columns([2, 1, 1])
                    p_b = col_f1.selectbox("Producto", df_prod['nombre'])
                    q_b = col_f2.number_input("Cantidad", min_value=1, value=None, placeholder="0")
                    if col_f3.form_submit_button("Cargar"):
                        if q_b is not None:
                            id_pb = int(df_prod[df_prod['nombre']==p_b]['id'].values[0])
                            registrar_carga(f_bo, id_vb, id_pb, q_b); st.rerun()
                        else: st.warning("Cantidad inválida")
            else:
                st.info("🔒 Modo solo lectura: No puedes agregar carga.")
        
        with st.container(border=True):
            st.markdown(f"**Detalle Carga: {v_bo}**")
            conn = get_conn()
            df_c = pd.read_sql("SELECT d.id, p.nombre, d.saldo_anterior, d.carga FROM despacho d JOIN productos p ON d.id_producto=p.id WHERE d.fecha=%s AND d.id_vendedor=%s AND d.carga > 0 ORDER BY p.orden_visual ASC", conn, params=(f_bo, id_vb))
            conn.close()
            
            if not df_c.empty:
                cols_desp = {
                    "id": None, 
                    "nombre": st.column_config.TextColumn("Producto", disabled=True, width="medium"), 
                    "saldo_anterior": st.column_config.NumberColumn("🔒 Saldo Anterior", disabled=True, width="small"), 
                    "carga": st.column_config.NumberColumn("Carga", min_value=0, required=True, width="small")
                }
                df_c_ed = st.data_editor(df_c, column_config=cols_desp, hide_index=True, use_container_width=True, key="edit_despacho", disabled=not permiso_editar)
                if permiso_editar:
                    if st.button("Corregir Carga"): actualizar_carga_masiva(df_c_ed)
            else: st.info("No hay carga registrada.")

    # === MÓDULO: COBRANZA ===
    elif seleccion == "Cobranza":
        st.title("Cobranza y Rendición")
        c1, c2 = st.columns([1, 2])
        f_of = c1.date_input("Fecha Cierre", date.today()-timedelta(days=1), format="DD/MM/YYYY")
        # El selectbox ahora está bloqueado en tu nombre
        v_of = c2.selectbox("Repartidor a Rendir", df_vend['nombre'], key="sel_cob_rep")
        id_vo = dict_vend[v_of]
        
        df_inv = obtener_planilla(f_of, id_vo)
        total_venta = 0 
        
        tab_inv, tab_fin, tab_rep = st.tabs(["Inventario", "Rendición", "Visor Reporte"])
        
        with tab_inv:
            with st.container(border=True):
                st.markdown("**Cuadre Inventario**")
                if not df_inv.empty:
                    df_inv['id_producto_hidden'] = df_inv['id_producto']
                    df_display = df_inv.rename(columns={'devolucion_muestra':'devolucion','saldo_actual':'saldo_final'})
                    
                    cols_cuadre = {
                        "id": None, "id_producto": None, "id_producto_hidden": None, "precio_estandar": None, "disp": None, "orden_sort": None, 
                        "nombre": st.column_config.TextColumn("Producto", disabled=True, width="medium"), 
                        "saldo_anterior": st.column_config.NumberColumn("🔒 Saldo Ant.", disabled=True, width="small"), 
                        "carga": st.column_config.NumberColumn("🔒 Carga", disabled=True, width="small"), 
                        "devolucion": st.column_config.NumberColumn("Devolución", required=True, min_value=0, width="small"), 
                        "saldo_final": st.column_config.NumberColumn("Saldo Final", required=True, min_value=0, width="small"), 
                        "venta": st.column_config.NumberColumn("🔒 Venta", disabled=True, width="small"), 
                        "total": st.column_config.NumberColumn("🔒 Total $", disabled=True, format="$%d", width="small")
                    }
                    
                    df_ed_inv = st.data_editor(df_display, column_config=cols_cuadre, hide_index=True, use_container_width=True, key="inv_ed", height=500, disabled=not permiso_editar)
                    df_ed_inv['v_real'] = (df_ed_inv['saldo_anterior']+df_ed_inv['carga']-df_ed_inv['devolucion']-df_ed_inv['saldo_final']).clip(lower=0)
                    total_venta = (df_ed_inv['v_real']*df_ed_inv['precio_estandar']).sum()
                    
                    c_inf, c_sv = st.columns([2,1])
                    c_inf.markdown(f"### Venta Estimada: {fmt_clp(total_venta)}")
                    if permiso_editar:
                        if c_sv.button("Guardar Inventario"): 
                            guardar_oficina(df_ed_inv, f_of, id_vo)
                            st.toast("Inventario actualizado")
                else: 
                    st.warning("No hay carga ni saldo anterior para este día.")
        
        with tab_fin:
            fin = get_finanzas(f_of, id_vo)
            with st.container(border=True):
                st.markdown("#### Rendición de Cuentas")
                st.info(f"💰 Venta Productos (Calculada): **{fmt_clp(total_venta)}**")
                
                col_izq, col_der = st.columns(2)
                with col_izq:
                    st.caption("Créditos Cobrados (Ingreso)")
                    df_cc_ed = st.data_editor(pd.DataFrame([{"Cliente": fin.get('cc_det', "Varios"), "Monto": int(fin.get('cc', 0))}]), num_rows="dynamic", key="grid_cc", hide_index=True, column_config={"Monto": st.column_config.NumberColumn(format="$%d", required=True)}, disabled=not permiso_editar)
                    total_cc = df_cc_ed['Monto'].sum(); cc_detalle_txt = ", ".join(df_cc_ed['Cliente'].astype(str).tolist())
                with col_der:
                    st.caption("Créditos Otorgados (Gasto)")
                    df_co_ed = st.data_editor(pd.DataFrame([{"Cliente": "Varios", "Monto": int(fin.get('co', 0))}]), num_rows="dynamic", key="grid_co", hide_index=True, column_config={"Monto": st.column_config.NumberColumn(format="$%d", required=True)}, disabled=not permiso_editar)
                    total_co = df_co_ed['Monto'].sum()
                    st.caption("Otros Gastos")
                    df_om_ed = st.data_editor(pd.DataFrame([{"Item": fin.get('od', "Varios"), "Monto": int(fin.get('om', 0))}]), num_rows="dynamic", key="grid_om", hide_index=True, column_config={"Monto": st.column_config.NumberColumn(format="$%d", required=True)}, disabled=not permiso_editar)
                    total_om = df_om_ed['Monto'].sum(); om_detalle_txt = ", ".join(df_om_ed['Item'].astype(str).tolist())

                st.divider()
                c_g1, c_g2, c_g3 = st.columns(3)
                bn = safe_int(c_g1.number_input("Bencina", value=val_gui(fin.get('bn',0)), step=1000, placeholder="0", disabled=not permiso_editar))
                su = safe_int(c_g2.number_input("Sueldo", value=val_gui(fin.get('su',0)), step=1000, placeholder="0", disabled=not permiso_editar))
                ds = safe_int(c_g3.number_input("Descuentos", value=val_gui(fin.get('ds',0)), step=500, placeholder="0", disabled=not permiso_editar))
                
                total_ingresos = total_venta + total_cc
                total_gastos = total_co + total_om + bn + su + ds
                deuda_bruta = total_ingresos - total_gastos
                
                st.markdown(f"**Total Ingresos: {fmt_clp(total_ingresos)}** | **Total Gastos: {fmt_clp(total_gastos)}**")
                st.markdown(f"### Total a Pagar: {fmt_clp(deuda_bruta)}")
                st.divider()
                
                st.markdown("**Entrega de Dinero**")
                c_r1, c_r2, c_r3 = st.columns(3)
                ef = safe_int(c_r1.number_input("Efectivo", value=val_gui(fin.get('ef',0)), step=1000, placeholder="0", disabled=not permiso_editar))
                tr = safe_int(c_r2.number_input("Transferencia", value=val_gui(fin.get('tr',0)), step=1000, placeholder="0", disabled=not permiso_editar))
                pc = safe_int(c_r3.number_input("Pago Centralizado", value=val_gui(fin.get('pc',0)), step=1000, placeholder="0", disabled=not permiso_editar))
                
                total_dinero = ef + tr + pc
                saldo_final = deuda_bruta - total_dinero
                
                if permiso_editar:
                    if st.button("Cerrar Cobro / Guardar Rendición", type="primary", use_container_width=True):
                        save_finanzas(f_of, id_vo, {"cc":total_cc, "cc_det":cc_detalle_txt, "co":total_co, "ds":ds, "bn":bn, "su":su, "om":total_om, "od":om_detalle_txt, "ef":ef, "tr":tr, "pc":pc})
                        st.toast("Cobranza cerrada correctamente")
                else:
                    st.info("🔒 Modo solo lectura")
                
                if saldo_final == 0: st.success(f"✅ CUADRE PERFECTO")
                elif saldo_final > 0: st.error(f"❌ DEBE: {fmt_clp(saldo_final)}")
                else: st.warning(f"⚠️ SOBRA: {fmt_clp(abs(saldo_final))}")

        with tab_rep:
            st.markdown("#### Resumen Histórico")
            cr1, cr2 = st.columns(2)
            fi = cr1.date_input("Desde", date.today() - timedelta(days=7), format="DD/MM/YYYY")
            ff = cr2.date_input("Hasta", date.today(), format="DD/MM/YYYY")
            
            if st.button("Generar Reporte Global"):
                df_res = obtener_resumen_global(fi, ff)
                if not df_res.empty:
                    def color_saldo(val):
                        if val == 0: return 'background-color: #d4edda; color: #155724' 
                        elif val > 0: return 'background-color: #f8d7da; color: #721c24'
                        return 'background-color: #fff3cd; color: #856404'
                    
                    st.dataframe(df_res.style.map(color_saldo, subset=['Saldo']).format({"Total Ingresos": "$ {:,.0f}", "Total Gastos": "$ {:,.0f}", "Deuda Neta": "$ {:,.0f}", "Pagado": "$ {:,.0f}", "Saldo": "$ {:,.0f}"}), use_container_width=True, height=500)
                    deuda_total = df_res['Saldo'].sum()
                    st.markdown(f"### Saldo Pendiente Total: {fmt_clp(deuda_total)}")
                else: st.info("No hay registros en este rango.")

    # === MÓDULO: CLIENTES ===
    elif seleccion == "Clientes":
        st.title("Directorio de Clientes")
        col_tools1, col_tools2 = st.columns([3, 1])
        # A. LÓGICA DE FILTROS VISUALES
        # Definimos si es un rol restringido
        es_repartidor = rol in ["repartidor_esp", "repartidor_corr"]
        
        filtro_rep = [] # Inicializamos vacío por defecto

        with col_tools1:
            with st.expander("🔍 Filtros de Búsqueda", expanded=True):
                if es_repartidor:
                    # CASO FRANCO: Solo ve 2 filtros (Nombre y Comuna). 
                    # El filtro de repartidor se oculta porque es redundante (solo se ve a sí mismo).
                    fc1, fc2 = st.columns(2)
                    filtro_nombre = fc1.text_input("Nombre", placeholder="Ej: Panadería...")
                    filtro_comuna = fc2.multiselect("Comuna", l_comunas)
                else:
                    # CASO ADMIN: Ve los 3 filtros incluyendo Repartidor
                    fc1, fc2, fc3 = st.columns(3)
                    filtro_nombre = fc1.text_input("Nombre", placeholder="Ej: Panadería...")
                    filtro_comuna = fc2.multiselect("Comuna", l_comunas)
                    filtro_rep = fc3.multiselect("Repartidor", df_vend['nombre'])
        
        modo_crear = False
        if permiso_editar:
            with col_tools2:
                modo_crear = st.toggle("➕ Crear Nuevo")

        # B. CARGA Y FILTRADO DE DATOS (El "Candado")
        df_show = obtener_clientes_df()
        
        # AQUÍ ESTÁ LA MAGIA: Si es repartidor, filtramos la data ANTES de mostrarla
        if es_repartidor and nombre_vendedor_actual:
            df_show = df_show[df_show['Repartidor'] == nombre_vendedor_actual]

        # C. APLICAR LOS OTROS FILTROS (Nombre y Comuna)
        if filtro_nombre: df_show = df_show[df_show['nombre'].str.contains(filtro_nombre, case=False, na=False)]
        if filtro_comuna: df_show = df_show[df_show['comuna'].isin(filtro_comuna)]
        
        # Si es Admin y seleccionó repartidores específicos en el filtro visual:
        if not es_repartidor and filtro_rep: 
            df_show = df_show[df_show['Repartidor'].isin(filtro_rep)]

        if modo_crear and permiso_editar:
            st.divider()
            st.markdown("#### Ficha de Nuevo Cliente")
            with st.form("form_cliente_crear", border=False):
                c1, c2, c3 = st.columns(3)
                nombre = c1.text_input("Nombre / Razón Social")
                direccion = c2.text_input("Dirección")
                telefono = c3.text_input("Teléfono")
                c4, c5, c6 = st.columns(3)
                comuna = c4.selectbox("Comuna", l_comunas)
                vendedor = c5.selectbox("Repartidor Asignado", df_vend['nombre'])
                tipo = c6.selectbox("Tipo Cliente", ["Nuevo", "Minorista", "Mayorista"])
                
                if st.form_submit_button("Guardar Cliente en Base de Datos", type="primary"):
                    if nombre:
                        crud_cliente("crear", {"nombre":nombre, "dir":direccion, "com":comuna, "tel":telefono, "id_vend":dict_vend[vendedor], "tipo":tipo})
                        st.success(f"Cliente {nombre} creado exitosamente."); time.sleep(1); st.rerun()
                    else: st.error("El nombre es obligatorio.")
        else:
            st.markdown(f"**Resultados:** {len(df_show)} clientes")
            event = st.dataframe(
                df_show, 
                use_container_width=True, 
                hide_index=True,
                selection_mode="single-row", 
                on_select="rerun",            
                column_config={
                    "id": None, "id_vendedor_asignado": None,
                    "nombre": st.column_config.TextColumn("Cliente", width="large", required=True),
                    "direccion": st.column_config.TextColumn("Dirección", width="medium"),
                    "comuna": st.column_config.TextColumn("Comuna", width="small"),
                    "telefono": st.column_config.TextColumn("Contacto", width="small"),
                    "Repartidor": st.column_config.TextColumn("Repartidor", width="medium"),
                    "limite_credito": st.column_config.NumberColumn("Cupo Crédito", format="$ %d", width="small"),
                    "tipo_cliente": st.column_config.TextColumn("Categoría", width="small")
                },
                height=500
            )
            
            if len(event.selection.rows) > 0 and permiso_editar:
                idx_sel = event.selection.rows[0]
                row = df_show.iloc[idx_sel]
                st.divider()
                st.markdown(f"#### ✏️ Editando: {row['nombre']}")
                with st.form("form_edit_context"):
                    ec1, ec2, ec3 = st.columns(3)
                    nom_e = ec1.text_input("Nombre", value=row['nombre'])
                    dir_e = ec2.text_input("Dirección", value=row['direccion'] if row['direccion'] != '-' else "")
                    tel_e = ec3.text_input("Teléfono", value=row['telefono'] if row['telefono'] != '-' else "")
                    ec4, ec5, ec6 = st.columns(3)
                    idx_c = l_comunas.index(row['comuna']) if row['comuna'] in l_comunas else 0
                    com_e = ec4.selectbox("Comuna", l_comunas, index=idx_c)
                    idx_v = list(dict_vend.keys()).index(row['Repartidor']) if row['Repartidor'] in dict_vend else 0
                    ven_e = ec5.selectbox("Repartidor", df_vend['nombre'], index=idx_v)
                    tipos = ["Nuevo", "Minorista", "Mayorista"]
                    idx_t = tipos.index(row['tipo_cliente']) if row['tipo_cliente'] in tipos else 0
                    tip_e = ec6.selectbox("Tipo", tipos, index=idx_t)
                    
                    if st.form_submit_button("Guardar Cambios"):
                        crud_cliente("editar", {"id": int(row['id']), "nombre":nom_e, "dir":dir_e, "com":com_e, "tel":tel_e, "id_vend":dict_vend[ven_e], "tipo":tip_e})
                        st.success("Datos actualizados."); time.sleep(1); st.rerun()

    # === MÓDULO: CRÉDITOS ===
    elif seleccion == "Créditos":
        st.title("Cuentas Corrientes Clientes")
        col_izq, col_der = st.columns([1, 1.5])
        df_clientes = obtener_clientes_df()
        
        # ---------------------------------------------------------
        # BLOQUE DE SEGURIDAD NUEVO:
        # Si es repartidor (Franco), eliminamos de la lista a los clientes que no son suyos.
        if rol in ["repartidor_esp", "repartidor_corr"] and nombre_vendedor_actual:
            df_clientes = df_clientes[df_clientes['Repartidor'] == nombre_vendedor_actual]
        # ---------------------------------------------------------
    

        with col_izq:
            st.markdown("##### Registrar Operación")
            with st.container(border=True):
                f_mov = st.date_input("Fecha", date.today())
                if not df_clientes.empty:
                    cli_sel = st.selectbox("Cliente", df_clientes['nombre'])
                    if "last_client_sel" not in st.session_state: st.session_state.last_client_sel = None
                    if cli_sel != st.session_state.last_client_sel:
                        row_c = df_clientes[df_clientes['nombre'] == cli_sel].iloc[0]
                        rep_def = row_c.get('Repartidor', '')
                        if rep_def in list(df_vend['nombre']):
                            st.session_state["rep_mov_key"] = rep_def 
                        st.session_state.last_client_sel = cli_sel
                    
                    # Selectbox bloqueado si es repartidor
                    vend_sel = st.selectbox("Repartidor", df_vend['nombre'], key="rep_mov_key")
                    id_vend_sel = dict_vend.get(vend_sel)
                    
                    tipo_sel = st.radio("Acción", ["Crédito (Fiado)", "Abono (Pago)"], horizontal=True)
                    tipo_db = "CREDITO" if "Crédito" in tipo_sel else "ABONO"
                    monto = st.number_input("Monto ($)", min_value=0, value=None, placeholder="Ingrese monto...")
                    detalle = st.text_input("Detalle", value=("Crédito Pan" if tipo_db == "CREDITO" else "Abono Cuenta"))
                    
                    # Permiso especial: Repartidores SI pueden crear créditos/abonos
                    # (Aunque permiso_editar sea False globalmente, en este modulo específico sí pueden)
                    # Ajuste: Si el usuario es repartidor, habilitamos el botón
                    puede_registrar = permiso_editar or rol == "repartidor_esp"
                    
                    if puede_registrar:
                        if st.button("Guardar Operación", type="primary", use_container_width=True):
                            id_cli_sel = int(df_clientes[df_clientes['nombre']==cli_sel]['id'].values[0])
                            if monto and monto > 0:
                                registrar_movimiento_credito(f_mov, id_cli_sel, id_vend_sel, tipo_db, monto, detalle)
                                st.success("Registrado correctamente")
                                time.sleep(0.5); st.rerun()
                            else: st.warning("Debe ingresar un monto válido.")
                    else:
                        st.info("🔒 Solo lectura")
                else: st.warning("Cree clientes primero.")
        
        with col_der:
            if not df_clientes.empty and cli_sel:
                id_cli_sel = int(df_clientes[df_clientes['nombre']==cli_sel]['id'].values[0])
                conn = get_conn()
                with conn.cursor() as c:
                    c.execute("SELECT SUM(CASE WHEN tipo_movimiento='CREDITO' THEN monto ELSE -monto END) FROM movimientos_credito WHERE id_cliente=%s", (id_cli_sel,))
                    saldo = c.fetchone()[0]
                saldo = saldo if saldo else 0
                conn.close()
                
                kpi_col1, kpi_col2 = st.columns([2,1])
                kpi_col1.markdown(f"### Historial: {cli_sel}")
                color_delta = "inverse" if saldo > 0 else "normal"
                kpi_col2.metric("Deuda Total", fmt_clp(saldo), delta_color=color_delta)
                
                conn = get_conn()
                hist = pd.read_sql("SELECT fecha, tipo_movimiento as \"Tipo\", monto, detalle FROM movimientos_credito WHERE id_cliente=%s ORDER BY fecha DESC LIMIT 15", conn, params=(id_cli_sel,))
                conn.close()
                
                if not hist.empty:
                    hist['Tipo'] = hist['Tipo'].replace({'CREDITO': 'Crédito', 'ABONO': 'Pago'})
                    st.dataframe(hist, use_container_width=True, hide_index=True, column_config={"fecha": st.column_config.DateColumn("Fecha", format="DD/MM/YYYY"), "Tipo": st.column_config.TextColumn("Tipo"), "monto": st.column_config.NumberColumn("Monto", format="$ %d"), "detalle": st.column_config.TextColumn("Detalle")})
                else: st.info("Este cliente no tiene movimientos registrados.")

    # === MÓDULO: TRANSFERENCIAS ===
    elif seleccion == "Transferencias":
        st.title("Gestión de Transferencias Bancarias")
        col_reg, col_rev = st.columns([1, 1])
        
        with col_reg:
            st.markdown("##### 1. Ingreso Rápido")
            with st.container(border=True):
                # Repartidor puede registrar transferencias
                puede_transf = permiso_editar or rol == "repartidor_esp"
                if puede_transf:
                    with st.form("form_transf_fast", clear_on_submit=True):
                        c1, c2 = st.columns(2)
                        f_tr = c1.date_input("Fecha", date.today())
                        monto_tr = c2.number_input("Monto", min_value=0, value=None, placeholder="Monto...")
                        c3, c4 = st.columns(2)
                        # Bloqueado en su nombre
                        rep_tr = c3.selectbox("Repartidor", df_vend['nombre'], key="rep_tr")
                        tipo_tr = c4.selectbox("Tipo", ["Pago Diario", "Abono Crédito"])
                        c5, c6, c7 = st.columns(3)
                        metodo = c5.selectbox("Método", ["Transferencia", "Depósito"])
                        banco_dest = c6.selectbox("Destino", ["Banco Estado", "Banco Chile"])
                        banco_emis = c7.selectbox("Origen", l_bancos)
                        
                        if st.form_submit_button("Agregar a la Cola"):
                            if monto_tr:
                                metodo_full = f"{metodo} a {banco_dest}"
                                registrar_transferencia(f_tr, dict_vend[rep_tr], monto_tr, metodo_full, banco_emis, tipo_tr, False, "")
                                st.toast("Transferencia agregada a la cola de verificación.")
                            else: st.warning("Falta el monto.")
                else:
                    st.info("🔒 Solo lectura")
        
        with col_rev:
            st.markdown("##### 2. Conciliación (Admin/Jefa)")
            conn = get_conn()
            df_pendientes = pd.read_sql("""SELECT t.id, t.fecha, v.nombre as "Repartidor", t.monto, t.banco_emisor, t.tipo_transferencia, t.verificado FROM transferencias t JOIN vendedores v ON t.id_vendedor=v.id ORDER BY t.fecha DESC, t.id DESC LIMIT 50""", conn)
            conn.close()
            
            if not df_pendientes.empty:
                df_pendientes['verificado'] = df_pendientes['verificado'].astype(bool)
                # Solo Admin/Jefe pueden verificar (permiso_editar=True)
                edited_df = st.data_editor(df_pendientes, key="editor_transf", use_container_width=True, hide_index=True, column_config={"id": None, "fecha": st.column_config.DateColumn("Fecha", format="DD/MM/YYYY", disabled=True), "Repartidor": st.column_config.TextColumn("Repartidor", disabled=True), "monto": st.column_config.NumberColumn("Monto", format="$ %d", disabled=True), "banco_emisor": st.column_config.TextColumn("Banco Emisor", disabled=True), "tipo_transferencia": st.column_config.TextColumn("Tipo", disabled=True), "verificado": st.column_config.CheckboxColumn("Ok (En Banco)", help="Marcar si el dinero llegó")}, height=600, disabled=not permiso_editar)
                if permiso_editar:
                    if st.button("💾 Guardar Verificaciones"):
                        actualizar_verificacion_masiva(edited_df)
                        time.sleep(1); st.rerun()
            else: st.info("No hay transferencias recientes.")

    # === MÓDULO: CAJA ===
    elif seleccion == "Caja":
        st.title("Libro de Caja Diaria")
        col_c1, col_c2, col_c3 = st.columns(3)
        fecha_caja = col_c1.date_input("Fecha", date.today(), key="f_caja")
        filtro_area = col_c2.selectbox("Área", ["Todas", "Pan Especial", "Pan Corriente"])
        
        if permiso_editar:
            with st.expander("Agregar Movimiento Manual", expanded=True):
                with st.form("form_caja_manual", clear_on_submit=True):
                    mc1, mc2, mc3, mc4 = st.columns(4)
                    desc_m = mc1.selectbox("Descripción", lista_descripciones, index=None, placeholder="Buscar...")
                    item_m = mc2.text_input("Item")
                    area_m = mc3.selectbox("Area", ["Pan Especial", "Pan Corriente"])
                    tipo_opciones = ["Ingreso (Efectivo)", "Ingreso (Transferencia)", "Egreso (Efectivo - Caja Chica)", "Egreso (Bancos/Cheque - Jefa)"]
                    tipo_m = mc4.selectbox("Tipo de Movimiento", tipo_opciones)
                    monto_m = st.number_input("Monto ($)", min_value=0, step=1000)
                    if st.form_submit_button("Registrar Movimiento", type="primary"):
                        if desc_m and monto_m > 0:
                            ie, it, eg_ef, eg_tr = 0, 0, 0, 0
                            if tipo_m == "Ingreso (Efectivo)": ie = monto_m
                            elif tipo_m == "Ingreso (Transferencia)": it = monto_m
                            elif tipo_m == "Egreso (Efectivo - Caja Chica)": eg_ef = monto_m 
                            else: eg_tr = monto_m 
                            guardar_movimiento_caja(fecha_caja, area_m, desc_m, item_m, ie, it, eg_ef, eg_tr)
                            st.success("Guardado correctamente"); time.sleep(0.5); st.rerun()
                        else: st.error("Debe seleccionar una descripción y un monto válido.")

        conn = get_conn()
        query = "SELECT id, area, descripcion, item, ingreso_efectivo, ingreso_transferencia, egreso as egreso_efectivo, egreso_transferencia FROM caja_movimientos WHERE fecha = %s"
        params = [fecha_caja]
        if filtro_area != "Todas": query += " AND area = %s"; params.append(filtro_area)
        df_caja = pd.read_sql(query, conn, params=tuple(params))
        conn.close()
        
        if not df_caja.empty:
            df_caja.fillna(0, inplace=True)
            tot_ing_ef = df_caja['ingreso_efectivo'].sum()
            tot_ing_tr = df_caja['ingreso_transferencia'].sum()
            tot_eg_ef = df_caja['egreso_efectivo'].sum()
            tot_eg_tr = df_caja['egreso_transferencia'].sum()
            saldo_caja_fisica = tot_ing_ef - tot_eg_ef 
            saldo_final_global = (tot_ing_ef + tot_ing_tr) - (tot_eg_ef + tot_eg_tr)
            
            st.divider()
            k1, k2, k3, k4, k5 = st.columns(5)
            k1.metric("Ingreso Efec.", fmt_clp(tot_ing_ef))
            k2.metric("Egresos Efec.", fmt_clp(tot_eg_ef))
            k3.metric("TOTAL CAJA EFECTIVO", fmt_clp(saldo_caja_fisica), delta="En Cajón", delta_color="off")
            k4.metric("Mov. Bancarios (Neto)", fmt_clp(tot_ing_tr - tot_eg_tr))
            k5.metric("Balance Final Día", fmt_clp(saldo_final_global), delta_color="normal")
            st.dataframe(df_caja, use_container_width=True, hide_index=True, column_config={"id": None, "fecha": None, "area": st.column_config.TextColumn("Área", width="small"), "descripcion": st.column_config.TextColumn("Descripción", width="medium"), "item": st.column_config.TextColumn("Item", width="medium"), "ingreso_efectivo": st.column_config.NumberColumn("Ingreso Efec.", format="$ %d", width="small"), "ingreso_transferencia": st.column_config.NumberColumn("Ingreso Transf.", format="$ %d", width="small"), "egreso_efectivo": st.column_config.NumberColumn("Egreso Efec.", format="$ %d", width="small"), "egreso_transferencia": st.column_config.NumberColumn("Egreso Banco", format="$ %d", width="small")})
        else: st.info("No hay movimientos registrados para esta fecha.")

    # === ADMIN ===
    elif seleccion == "Admin" and rol == "admin":
        st.error("ZONA DE ADMINISTRACIÓN DE BASE DE DATOS")
        with st.container(border=True):
            col_b1, col_b2 = st.columns([3, 1])
            col_b1.markdown("### 💾 Copia de Seguridad")
            col_b1.caption("Descarga toda la base de datos en formato Excel local.")
            if col_b2.button("Generar Backup"):
                with st.spinner("Compilando todas las tablas..."):
                    buffer = descargar_respaldo_completo()
                    val_fecha = date.today().strftime("%Y-%m-%d")
                    st.download_button(label="⬇️ Descargar .xlsx", data=buffer, file_name=f"Respaldo_Kilaco_{val_fecha}.xlsx", mime="application/vnd.ms-excel", key="btn_download_backup")
        st.divider()

        tablas_disponibles = ["control_bandejas", "control_bolsas", "stock", "despacho", "movimientos_credito", "transferencias", "caja_movimientos"]
        tabla_sel = st.selectbox("Seleccionar Tabla a Corregir", tablas_disponibles)
        col1, col2 = st.columns(2)
        fecha_filtro = col1.date_input("Filtrar por Fecha", date.today())
        conn = get_conn()
        try:
            query = f"SELECT * FROM {tabla_sel} WHERE fecha = %s"
            df = pd.read_sql(query, conn, params=(fecha_filtro,))
            st.divider()
            st.subheader(f"Registros en '{tabla_sel}' del {fecha_filtro}")
            if not df.empty:
                st.dataframe(df, use_container_width=True)
                col_del1, col_del2 = st.columns([1, 1])
                ids_validos = df['id'].tolist()
                id_a_borrar = col_del1.selectbox("Borrar un registro (ID):", ids_validos)
                if col_del1.button("Eliminar ID Seleccionado"):
                    with conn.cursor() as c:
                        c.execute(f"DELETE FROM {tabla_sel} WHERE id=%s", (id_a_borrar,))
                    conn.commit()
                    st.success(f"ID {id_a_borrar} eliminado."); time.sleep(1); st.rerun()
                if col_del2.button("BORRAR TODO EL DÍA", type="primary"):
                    with conn.cursor() as c:
                        c.execute(f"DELETE FROM {tabla_sel} WHERE fecha=%s", (fecha_filtro,))
                    conn.commit()
                    st.error(f"Todos los registros del {fecha_filtro} fueron eliminados."); time.sleep(1); st.rerun()
            else: st.info("No hay registros para esta fecha.")
        except Exception as e: st.error(f"Error: {e}")
        finally: conn.close()

# ----------------------------------------------------
# APLICACIÓN PAN CORRIENTE (CORREGIDA: FILTRO VENDEDOR)
# ----------------------------------------------------
from streamlit_option_menu import option_menu

def app_pan_corriente():
    rol = st.session_state.user_role
    mi_id_vendedor = st.session_state.id_vendedor
    
    # --- 1. CONFIGURACIÓN DE MENÚ Y PERMISOS ---
    opciones_full = ["Producción", "Despacho", "Cobranza", "Clientes", "Caja"]
    iconos_full   = ["tools",      "truck",    "currency-dollar", "people", "cash-stack"]
    
    if rol in ["admin", "pan_corriente", "supervisor"]:
        menu_options = opciones_full
        menu_icons = iconos_full
        permiso_editar = True
        
    elif rol == "repartidor_corr":
        # Repartidor Corriente: Ve Despacho, Cobranza, Clientes
        indices_permitidos = [1, 2, 3]
        menu_options = [opciones_full[i] for i in indices_permitidos]
        menu_icons = [iconos_full[i] for i in indices_permitidos]
        permiso_editar = False # Solo lectura
        
    else:
        menu_options = []
        menu_icons = []
        permiso_editar = False

    if rol == "admin":
        menu_options.append("Admin")
        menu_icons.append("gear")

    # --- 2. SIDEBAR ---
    with st.sidebar:
        c_logo1, c_logo2, c_logo3 = st.columns([1, 1.5, 1])
        with c_logo2:
            st.image("logo.png", use_container_width=True)
        
        st.markdown("""
            <h3 style='text-align: center; margin-top: 0px; margin-bottom: 25px; font-weight: 600; color: #333;'>
                KILACO ERP
            </h3>
        """, unsafe_allow_html=True)
        
        if not menu_options:
            st.error("No tienes acceso a este módulo.")
            st.stop()
            
        seleccion = option_menu(
            menu_title=None, 
            options=menu_options,
            icons=menu_icons,
            menu_icon="cast", 
            default_index=0,
            styles={
                "container": {"padding": "0!important", "background-color": "#ffffff"},
                "icon": {"color": "#ff4b4b", "font-size": "14px"}, 
                "nav-link": {"font-size": "14px", "text-align": "left", "margin":"0px", "--hover-color": "#f0f2f6"},
                "nav-link-selected": {"background-color": "#ff4b4b", "font-weight": "600"},
            }
        )
        
        st.markdown("---")
        if st.button("Volver al Menú", use_container_width=True):
            st.session_state.current_module = "menu"
            st.rerun()

    # --- 3. REFERENCIAS Y FILTRO DE SEGURIDAD ---
    l_comunas, l_bancos, df_vend = get_referencias()

    # === 🔒 FIX DE SEGURIDAD ===
    # Si es repartidor, filtramos para que solo exista SU vendedor en las listas.
    if rol == "repartidor_corr" and mi_id_vendedor:
        df_vend = df_vend[df_vend['id'] == mi_id_vendedor]
        if df_vend.empty:
            st.error(f"Error crítico: Tu ID de vendedor {mi_id_vendedor} no existe en la base de datos.")
            st.stop()
    # ============================

    dict_vend = dict(zip(df_vend['nombre'], df_vend['id']))
    
    lista_descripciones = sorted([
        "Alejandro Valenzuela", "Carlos Alvarez", "Randy Galvis", "Carlos Jara", "Eduardo Sanchez",
        "Edward Baillont", "Flavio Meza", "Franco De La Puente", "Hugo Palacios", "Jose Perozo",
        "Marcelo Jara", "Robert Cordova", "Robert Mogollon", "Hector Silva", "Byron Navarro",
        "Jose Albarracin", "Tomas Mendez", "Junior Negron", "Kilaco Venta", "Otros", "Petroleo",
        "Anderson", "Aseo", "Alberto", "Juanito", "Marmix", "Osvaldo Torres", "Fabiola Tessini",
        "Ines Cordova", "Mitzy Panes", "David", "Aceite", "Alex (Carne Pato)", "Provision",
        "Yves", "Kervis", "Carlos Hornero", "Mila", "Yordys", "Gene", "Yeilid", "Michel",
        "Maestro Guillermo", "Rahab", "Mary", "Imposiciones", "Caro", "Yeimy", "Leo",
        "Jonathan", "Luis", "Jorge", "Genesis", "Guillermo", "Paola"
    ])
    
    # Lista original de vendedores corrientes
    vendedores_corriente_full = ["Hector Silva", "Byron Navarro", "Jose Albarracin", "Tomas Mendez"]
    
    # 🔒 Aplicamos el filtro también a la lista específica de corriente
    if rol == "repartidor_corr":
        # Sobrescribimos la lista para que solo tenga el nombre del usuario
        # (df_vend ya fue filtrado arriba, así que tomamos el único nombre que queda)
        vendedores_corriente = df_vend['nombre'].tolist() 
    else:
        vendedores_corriente = vendedores_corriente_full

    # --- 4. ENRUTAMIENTO DE VISTAS ---

    # === PRODUCCIÓN (Planificación Diaria) ===
    if seleccion == "Producción":
        st.title("Planificación Diaria")
        f_prod = st.date_input("Fecha Producción", date.today())
        
        df_unificado, extras_dict = get_produccion_corriente_unificada(f_prod)
        
        st.dataframe(df_unificado, hide_index=True, use_container_width=True,
                     column_config={
                         "Concepto": st.column_config.TextColumn("Responsable"),
                         "rinde_noche": st.column_config.NumberColumn("Rinde Noche", format="%d kg"),
                         "rinde_dia": st.column_config.NumberColumn("Rinde Día", format="%d kg")
                     })

        st.divider()
        st.markdown("##### Editar Extras")
        
        c1, c2, c3 = st.columns(3)
        with c1:
            st.caption("Ración")
            rd = st.number_input("Día", value=val_gui(extras_dict['rd']), step=1, placeholder="0", key="rd", disabled=not permiso_editar)
            rn = st.number_input("Noche", value=val_gui(extras_dict['rn']), step=1, placeholder="0", key="rn", disabled=not permiso_editar)
        with c2:
            st.caption("Adicional")
            ad = st.number_input("Día", value=val_gui(extras_dict['ad']), step=1, placeholder="0", key="ad", disabled=not permiso_editar)
            an = st.number_input("Noche", value=val_gui(extras_dict['an']), step=1, placeholder="0", key="an", disabled=not permiso_editar)
        with c3:
            st.caption("Kilaco (Comodín)")
            kd = st.number_input("Día", value=val_gui(extras_dict['kd']), step=1, placeholder="0", key="kd", disabled=not permiso_editar)
            kn = st.number_input("Noche", value=val_gui(extras_dict['kn']), step=1, placeholder="0", key="kn", disabled=not permiso_editar)

        total_noche = df_unificado['rinde_noche'].sum()
        total_dia = df_unificado['rinde_dia'].sum()
        
        st.divider()
        k1, k2, k3 = st.columns(3)
        k1.metric("Total Noche", f"{int(total_noche)} kg")
        k2.metric("Total Día", f"{int(total_dia)} kg")
        k3.metric("GRAN TOTAL", f"{int(total_noche + total_dia)} kg")
        
        if permiso_editar:
            if st.button("Guardar Cambios en Extras"):
                save_extras_produccion(f_prod, {"rd": rd, "rn": rn, "ad": ad, "an": an, "kd": kd, "kn": kn})
                st.success("Actualizado"); time.sleep(0.5); st.rerun()

    # === DESPACHO (Hoja de Ruta) ===
    elif seleccion == "Despacho":
        st.title("Hoja de Ruta")
        c1, c2, c3 = st.columns(3)
        f_desp = c1.date_input("Fecha Despacho", date.today())
        # Bloqueado en su nombre por el filtro inicial
        v_desp = c2.selectbox("Seleccionar Repartidor", vendedores_corriente, key="sel_rep_corr")
        vista_detalle = c3.radio("Modo Visualización", ["Detalle Turnos", "Pagos"], horizontal=True, label_visibility="collapsed")
        
        if v_desp in dict_vend:
            id_v_desp = dict_vend.get(v_desp)
            df_ruta = get_despacho_corriente(f_desp, id_v_desp)
            
            if not df_ruta.empty:
                cols_cfg = {
                    "id": None, "fecha": None, "id_cliente": None, "id_vendedor": None,
                    "cliente": st.column_config.TextColumn("Cliente", disabled=True, width="medium"),
                    "precio_aplicado": st.column_config.NumberColumn("🔒 Precio", format="$ %d", disabled=True, width="small"),
                    "saldo_anterior": st.column_config.NumberColumn("🔒 Saldo Ant.", format="$ %d", disabled=True, width="small"),
                    "deuda_final": st.column_config.NumberColumn("🔒 Deuda Final", format="$ %d", disabled=True, width="small"),
                    "total_carga": st.column_config.NumberColumn("🔒 Total Kg", disabled=True, format="%d", width="small"),
                    "ventas_monto": st.column_config.NumberColumn("🔒 Venta ($)", format="$ %d", disabled=True, width="small"),
                    "total_pagar": st.column_config.NumberColumn("🔒 Total", format="$ %d", disabled=True, width="small"),
                    "paga": st.column_config.NumberColumn("Paga", format="$ %d", width="small"),
                    "pago_centralizado": st.column_config.NumberColumn("Centralizado", format="$ %d", width="small")
                }
                
                for i in range(1,8): cols_cfg[f"carga_{i}"] = st.column_config.NumberColumn(f"{i}", width="small")

                if vista_detalle == "Pagos":
                    column_order = ["cliente", "saldo_anterior", "ventas_monto", "total_carga", "total_pagar", "paga", "pago_centralizado", "deuda_final"]
                    for k in ["precio_aplicado"] + [f"carga_{i}" for i in range(1,8)]: cols_cfg[k] = None
                else:
                    column_order = ["cliente", "precio_aplicado"] + [f"carga_{i}" for i in range(1,8)] + ["total_carga"]
                    for k in ["saldo_anterior", "ventas_monto", "total_pagar", "paga", "pago_centralizado", "deuda_final"]: cols_cfg[k] = None

                cols_finales = column_order + ["id", "fecha", "id_cliente", "id_vendedor"]
                for col in cols_finales:
                    if col not in df_ruta.columns: df_ruta[col] = 0
                
                df_visual = df_ruta[cols_finales]
                edited_ruta = st.data_editor(df_visual, column_config=cols_cfg, hide_index=True, use_container_width=True, height=600, disabled=not permiso_editar)
                
                tot_kg = int(edited_ruta['total_carga'].sum())
                if vista_detalle == "Detalle Turnos": st.info(f"Total Carga del día: **{tot_kg} kg**")
                
                if permiso_editar:
                    if st.button("Guardar Despacho"):
                        save_despacho_corriente(edited_ruta)
                        st.success("Ruta guardada.")
            else: st.warning("No hay clientes asignados.")

    # === COBRANZA ===
    elif seleccion == "Cobranza":
        st.title("Cobranza y Rendición")
        
        tab_cob, tab_visor = st.tabs(["Rendición Diaria", "Visor Reporte"])
        
        # Pestaña 1: Rendición (Input de dinero)
        with tab_cob:
            cc1, cc2 = st.columns(2)
            f_cob = cc1.date_input("Fecha", date.today())
            # Bloqueado en su nombre
            v_cob = cc2.selectbox("Repartidor", vendedores_corriente, key="sel_rep_cob_corr")
            
            if v_cob in dict_vend:
                id_vc = dict_vend.get(v_cob)
                fin = get_finanzas_corriente(f_cob, id_vc)
                
                with st.form("form_cobranza_corr"):
                    st.markdown(f"**Venta del día (Total Despachado): {fmt_clp(fin['venta_real'])}**")
                    c1, c2 = st.columns(2)
                    recaudo = c1.number_input("🔒 Recaudo Total", value=fin['recaudo'], step=1000, disabled=True)
                    
                    comision_sugerida = int(fin['venta_real'] * 0.04)
                    val_com = fin['comision'] if fin['comision'] > 0 else comision_sugerida
                    comision = c2.number_input("🔒 Comisión (4%)", value=val_com, step=100, disabled=True)
                    
                    st.divider()
                    st.markdown("**Gastos y Descuentos**")
                    g1, g2, g3 = st.columns(3)
                    bencina = g1.number_input("🔒 Bencina (Fijo)", value=fin['bencina'], step=1000, disabled=True)
                    sueldo = g2.number_input("Sueldo", value=val_gui(fin['sueldo']), step=1000, placeholder="0", disabled=not permiso_editar)
                    otros = g3.number_input("Otros Gastos", value=val_gui(fin['otros']), step=1000, placeholder="0", disabled=not permiso_editar)
                    detalles = st.text_input("Detalle Otros Gastos", value=fin['det'], disabled=not permiso_editar)
                    
                    st.divider()
                    st.markdown("**Entrega Final**")
                    p1, p2 = st.columns(2)
                    efectivo = p1.number_input("Pago Efectivo", value=val_gui(fin['efec']), step=1000, placeholder="0", disabled=not permiso_editar)
                    transf = p2.number_input("Pago Transferencia", value=val_gui(fin['trans']), step=1000, placeholder="0", disabled=not permiso_editar)
                    
                    total_gastos = comision + bencina + val_db(sueldo) + val_db(otros)
                    saldo_final = recaudo - total_gastos - val_db(efectivo) - val_db(transf)
                    
                    if saldo_final == 0: st.success("✅ CUADRADO")
                    elif saldo_final > 0: st.error(f"❌ FALTA: {fmt_clp(saldo_final)}")
                    else: st.warning(f"⚠️ SOBRA: {fmt_clp(abs(saldo_final))}")
                    
                    if permiso_editar:
                        if st.form_submit_button("Guardar Rendición"):
                            datos = {"venta": fin['venta_real'], "recaudo": recaudo, "comision": comision, "bencina": bencina, "sueldo": val_db(sueldo), "otros": val_db(otros), "det": detalles, "efec": val_db(efectivo), "trans": val_db(transf)}
                            save_finanzas_corriente(f_cob, id_vc, datos)
                            st.success("Guardado")
                    else:
                        st.info("🔒 Solo lectura")

        # Pestaña 2: Visor (Reporte Global)
        with tab_visor:
            st.markdown("#### Estado General Pan Corriente")
            col_r1, col_r2 = st.columns(2)
            fi = col_r1.date_input("Desde", date.today() - timedelta(days=7))
            ff = col_r2.date_input("Hasta", date.today())
            
            if st.button("Generar Reporte"):
                df_res = get_resumen_visor_corriente(fi, ff)
                if not df_res.empty:
                    def color_saldo(val):
                        if val == 0: return 'background-color: #d4edda; color: #155724' 
                        elif val > 0: return 'background-color: #f8d7da; color: #721c24'
                        return 'background-color: #fff3cd; color: #856404'
                    
                    st.dataframe(df_res.style.map(color_saldo, subset=['saldo_final']).format({"total_gastos": "$ {:,.0f}", "saldo_final": "$ {:,.0f}", "saldo_clientes": "$ {:,.0f}"}), use_container_width=True, height=500, column_config={"saldo_final": "Deuda Neta Repartidor", "saldo_clientes": "Deuda Calle (Clientes)"})
                    m1, m2 = st.columns(2)
                    m1.metric("Total Deuda Repartidores", fmt_clp(df_res['saldo_final'].sum()))
                    m2.metric("Total Deuda Clientes (Calle)", fmt_clp(df_res['saldo_clientes'].sum()))
                else: st.info("No hay registros.")

    # === CLIENTES ===
    elif seleccion == "Clientes":
        st.title("Gestión de Clientes")
        col_tools1, col_tools2 = st.columns([3, 1])
        with col_tools1:
            with st.expander("🔍 Filtros de Búsqueda", expanded=False):
                fc1, fc2 = st.columns(2)
                filtro_cli = fc1.text_input("Cliente", placeholder="Buscar...")
                # Solo ve su nombre en el multiselect
                filtro_rep = fc2.multiselect("Repartidor", vendedores_corriente)
        
        modo_crear = False
        if permiso_editar:
            with col_tools2:
                modo_crear = st.toggle("➕ Crear Nuevo")

        df_cli = obtener_clientes_corriente()
        if filtro_cli: df_cli = df_cli[df_cli['nombre'].str.contains(filtro_cli, case=False, na=False)]
        if filtro_rep: df_cli = df_cli[df_cli['Repartidor'].isin(filtro_rep)]

        if modo_crear and permiso_editar:
            st.divider()
            with st.form("new_cli_corr", border=False):
                nc1, nc2, nc3 = st.columns(3)
                n_nom = nc1.text_input("Nombre Cliente")
                n_rep = nc2.selectbox("Repartidor", vendedores_corriente)
                n_pre = nc3.number_input("Precio Pactado", value=1600, step=50)
                if st.form_submit_button("Crear Cliente", type="primary"):
                    if n_nom:
                        crud_cliente_corriente("crear", {"nombre": n_nom, "id_vendedor": dict_vend[n_rep], "precio": n_pre})
                        st.success("Creado"); time.sleep(0.5); st.rerun()
        else:
            st.markdown(f"**Resultados:** {len(df_cli)} clientes")
            event = st.dataframe(
                df_cli,
                column_config={
                    "id": None, "id_vendedor": None,
                    "nombre": st.column_config.TextColumn("Cliente", width="large"),
                    "Repartidor": st.column_config.TextColumn("Repartidor", width="medium"),
                    "precio_pactado": st.column_config.NumberColumn("Precio ($)", format="$ %d")
                },
                hide_index=True,
                use_container_width=True,
                selection_mode="single-row",
                on_select="rerun",
                height=500
            )
            if len(event.selection.rows) > 0 and permiso_editar:
                idx_sel = event.selection.rows[0]
                row = df_cli.iloc[idx_sel]
                st.divider()
                st.markdown(f"#### ✏️ Editando: {row['nombre']}")
                with st.form("edit_cli_corr"):
                    e1, e2, e3 = st.columns(3)
                    e_nom = e1.text_input("Nombre", value=row['nombre'])
                    idx_r = vendedores_corriente.index(row['Repartidor']) if row['Repartidor'] in vendedores_corriente else 0
                    e_rep = e2.selectbox("Repartidor", vendedores_corriente, index=idx_r)
                    e_pre = e3.number_input("Precio", value=row['precio_pactado'], step=50)
                    if st.form_submit_button("Guardar Cambios"):
                        crud_cliente_corriente("editar", {"id": int(row['id']), "nombre": e_nom, "id_vendedor": dict_vend[e_rep], "precio": e_pre})
                        st.success("Actualizado"); time.sleep(0.5); st.rerun()

    # === CAJA ===
    elif seleccion == "Caja":
        st.title("Caja Pan Corriente")
        col_c1, col_c2, col_c3 = st.columns(3)
        fecha_caja = col_c1.date_input("Fecha", date.today(), key="f_caja_corr")
        
        if permiso_editar:
            with st.expander("Agregar Movimiento Manual", expanded=True):
                with st.form("form_caja_corr", clear_on_submit=True):
                    mc1, mc2, mc3, mc4 = st.columns(4)
                    desc_m = mc1.selectbox("Descripción", lista_descripciones, index=None, placeholder="Buscar...")
                    item_m = mc2.text_input("Item")
                    tipo_opciones = ["Ingreso (Efectivo)", "Ingreso (Transferencia)", "Egreso (Efectivo - Caja Chica)", "Egreso (Bancos/Cheque - Jefa)"]
                    tipo_m = mc4.selectbox("Tipo de Movimiento", tipo_opciones)
                    monto_m = st.number_input("Monto ($)", min_value=0, step=1000)
                    if st.form_submit_button("Registrar Movimiento", type="primary"):
                        if desc_m and monto_m > 0:
                            ie, it, eg_ef, eg_tr = 0, 0, 0, 0
                            if tipo_m == "Ingreso (Efectivo)": ie = monto_m
                            elif tipo_m == "Ingreso (Transferencia)": it = monto_m
                            elif tipo_m == "Egreso (Efectivo - Caja Chica)": eg_ef = monto_m 
                            else: eg_tr = monto_m 
                            guardar_movimiento_caja(fecha_caja, "Pan Corriente", desc_m, item_m, ie, it, eg_ef, eg_tr)
                            st.success("Guardado"); time.sleep(0.5); st.rerun()
                        else: st.error("Faltan datos")

        conn = get_conn()
        query = "SELECT id, area, descripcion, item, ingreso_efectivo, ingreso_transferencia, egreso as egreso_efectivo, egreso_transferencia FROM caja_movimientos WHERE fecha = %s AND area='Pan Corriente'"
        params = [fecha_caja]
        df_caja = pd.read_sql(query, conn, params=tuple(params))
        conn.close()
        
        if not df_caja.empty:
            df_caja.fillna(0, inplace=True)
            tot_ing_ef = df_caja['ingreso_efectivo'].sum()
            tot_ing_tr = df_caja['ingreso_transferencia'].sum()
            tot_eg_ef = df_caja['egreso_efectivo'].sum()
            tot_eg_tr = df_caja['egreso_transferencia'].sum()
            saldo_caja_fisica = tot_ing_ef - tot_eg_ef 
            saldo_final_global = (tot_ing_ef + tot_ing_tr) - (tot_eg_ef + tot_eg_tr)
            st.divider()
            k1, k2, k3, k4, k5 = st.columns(5)
            k1.metric("Ingreso Efec.", fmt_clp(tot_ing_ef))
            k2.metric("Egresos Efec.", fmt_clp(tot_eg_ef))
            k3.metric("TOTAL CAJA EFECTIVO", fmt_clp(saldo_caja_fisica), delta="En Cajón", delta_color="off")
            k4.metric("Mov. Bancarios (Neto)", fmt_clp(tot_ing_tr - tot_eg_tr))
            k5.metric("Balance Final Día", fmt_clp(saldo_final_global), delta_color="normal")
            st.dataframe(df_caja, use_container_width=True, hide_index=True, column_config={"id": None, "fecha": None, "area": st.column_config.TextColumn("Área", width="small"), "descripcion": st.column_config.TextColumn("Descripción", width="medium"), "item": st.column_config.TextColumn("Item", width="medium"), "ingreso_efectivo": st.column_config.NumberColumn("Ingreso Efec.", format="$ %d", width="small"), "ingreso_transferencia": st.column_config.NumberColumn("Ingreso Transf.", format="$ %d", width="small"), "egreso_efectivo": st.column_config.NumberColumn("Egreso Efec.", format="$ %d", width="small"), "egreso_transferencia": st.column_config.NumberColumn("Egreso Banco", format="$ %d", width="small")})
        else: st.info("No hay movimientos registrados para esta fecha.")

    # === ADMIN ===
    elif seleccion == "Admin" and rol == "admin":
        st.error("ZONA DE ADMINISTRACIÓN DE BASE DE DATOS")
        with st.container(border=True):
            col_b1, col_b2 = st.columns([3, 1])
            col_b1.markdown("### 💾 Copia de Seguridad")
            col_b1.caption("Descarga toda la base de datos en formato Excel local.")
            if col_b2.button("Generar Backup"):
                with st.spinner("Compilando..."):
                    buffer = descargar_respaldo_completo()
                    val_fecha = date.today().strftime("%Y-%m-%d")
                    st.download_button(label="⬇️ Descargar", data=buffer, file_name=f"Respaldo_Kilaco_{val_fecha}.xlsx", mime="application/vnd.ms-excel")
        st.divider()
        tablas_disponibles = ["despacho_corriente", "clientes_corriente", "caja_movimientos", "produccion_corriente", "finanzas_corriente", "produccion_extras"]
        tabla_sel = st.selectbox("Seleccionar Tabla", tablas_disponibles)
        fecha_filtro = st.date_input("Filtrar por Fecha", date.today())
        conn = get_conn()
        try:
            query = f"SELECT * FROM {tabla_sel}"
            if tabla_sel != "clientes_corriente": query += " WHERE fecha = %s"
            params = (fecha_filtro,) if tabla_sel != "clientes_corriente" else ()
            df = pd.read_sql(query, conn, params=params)
            st.dataframe(df, use_container_width=True)
            if not df.empty:
                col_del1, col_del2 = st.columns([1, 1])
                ids_validos = df['id'].tolist()
                id_a_borrar = col_del1.selectbox("ID a borrar:", ids_validos)
                if col_del1.button("Eliminar ID"):
                    with conn.cursor() as c: c.execute(f"DELETE FROM {tabla_sel} WHERE id=%s", (id_a_borrar,))
                    conn.commit(); st.success("Eliminado"); time.sleep(1); st.rerun()
                if tabla_sel != "clientes_corriente" and col_del2.button("BORRAR TODO EL DÍA", type="primary"):
                    with conn.cursor() as c: c.execute(f"DELETE FROM {tabla_sel} WHERE fecha=%s", (fecha_filtro,))
                    conn.commit(); st.error("Día eliminado."); time.sleep(1); st.rerun()
        except Exception as e: st.error(e)
        finally: conn.close()

# ----------------------------------------------------
# EJECUCIÓN PRINCIPAL
# ----------------------------------------------------
if not st.session_state.logged_in: login_view()
else:
    if st.session_state.current_module == "menu": menu_view()
    elif st.session_state.current_module == "especial": app_pan_especial()
    elif st.session_state.current_module == "corriente": app_pan_corriente()