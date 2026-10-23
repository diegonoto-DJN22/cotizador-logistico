[app.py](https://github.com/user-attachments/files/33884554/app.py)
import streamlit as st
import pandas as pd
import os
import glob

st.set_page_config(page_title="Cotizador y Comparador Logístico", layout="wide")

st.markdown("""
<style>
    .stApp { background-color: #6a1b29; }
    .stApp, .stApp p, .stApp label, .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp div { color: #fdfbf7 !important; }
    .resultado-card { background-color: #8b2b3d; padding: 15px; border-radius: 12px; box-shadow: 2px 2px 10px rgba(0,0,0,0.3); text-align: center; margin-bottom: 10px; border: 1px solid #fdfbf7; }
    .resultado-card h2, .resultado-card h3 { color: #ffffff !important; margin: 5px 0;}
    /* Separador vertical para dividir la pantalla */
    .css-1r6slb0 { border-right: 2px solid #8b2b3d; } 
</style>
""", unsafe_allow_html=True)

# --- ENCABEZADO ---
col_logo1, col_titulo, col_logo2 = st.columns([1, 2, 1])
with col_logo1:
    if os.path.exists("toso.png"): st.image("toso.png", width=120)
    else: st.markdown("### 🍷 Pascual Toso")
with col_titulo:
    st.markdown("<h1 style='text-align: center;'>⚖️ Comparador Logístico</h1>", unsafe_allow_html=True)
with col_logo2:
    if os.path.exists("llorente.png"): st.image("llorente.png", width=120)
    else: st.markdown("### 🏢 Llorente & Cia SA")
st.markdown("---")

archivos = glob.glob("*.xls*")
archivo_correcto = "Base Tarifario_2.xlsx"

if archivo_correcto not in archivos and "Base Tarifario.xlsx" not in archivos:
    st.error("🚨 No hay ningún archivo Excel en la carpeta.")
    st.stop()

@st.cache_data
def cargar_datos():
    archivo = archivo_correcto if archivo_correcto in archivos else "Base Tarifario.xlsx"
    try:
        df = pd.read_excel(archivo)
        df.columns = df.columns.str.strip().str.replace('\xa0', ' ')
        for col in df.select_dtypes(include=['object']).columns:
            df[col] = df[col].astype(str).str.strip().str.replace('\xa0', ' ').replace({'nan': '', 'None': ''})
        return df
    except Exception as e:
        st.error(f"Error al leer Excel: {e}")
        return pd.DataFrame()

df = cargar_datos()

if df.empty:
    st.stop()

cols = df.columns.tolist()

# ESCÁNER GLOBAL DE COLUMNAS
col_transporte = None
for col in cols:
    if df[col].astype(str).str.contains("calicó|calico", case=False, na=False).any():
        col_transporte = col
        break
if not col_transporte: col_transporte = next((c for c in cols if c.strip().lower() in ['transporte', 'operador', 'descripción']), cols[0])

col_tipo = next((c for c in cols if 'tipo tarifa' in c.lower() or 'tipo entrega' in c.lower()), None)
col_loc_ori = next((c for c in cols if c.strip().lower() == 'localidad origen'), None)
if not col_loc_ori: col_loc_ori = next((c for c in cols if 'localidad origen' in c.lower() or 'subzona desde' in c.lower()), None)
col_loc_des = next((c for c in cols if c.strip().lower() == 'localidad destino'), None)
if not col_loc_des: col_loc_des = next((c for c in cols if 'localidad destino' in c.lower() or ('subzona' in c.lower() and 'desde' not in c.lower())), None)
col_directo = next((c for c in cols if 'viaje' in c.lower() or 'directo' in c.lower()), None)
col_base = next((c for c in cols if 'base' in c.lower()), None)
col_interior = next((c for c in cols if 'interior' in c.lower() or 'distribucion' in c.lower()), None)
col_manipuleo = next((c for c in cols if 'bulto' in c.lower() or 'manipuleo' in c.lower()), None)

# =====================================================================
# FUNCIÓN DEL COTIZADOR INDIVIDUAL (Para poder duplicarlo)
# =====================================================================
def crear_panel_cotizador(prefijo_key, color_titulo):
    st.markdown(f"<h2 style='text-align: center; color: {color_titulo};'>{prefijo_key}</h2>", unsafe_allow_html=True)
    
    # 1. MODALIDAD
    modalidad = st.radio("Modalidad:", ["Distribución", "Directo"], horizontal=True, key=f"mod_{prefijo_key}")
    
    df_filtrado = df
    if col_tipo:
        if modalidad == "Distribución":
            filtro = df[df[col_tipo].str.contains("distribuc", case=False, na=False)]
        else:
            filtro = df[df[col_tipo].str.contains("direct", case=False, na=False)]
        if not filtro.empty:
            df_filtrado = filtro

    # 2. OPERADOR
    transportes = sorted([t for t in df_filtrado[col_transporte].unique() if t])
    transporte = st.selectbox("Transporte:", transportes if transportes else ["Sin datos"], key=f"trans_{prefijo_key}")
    df_transporte = df_filtrado[df_filtrado[col_transporte] == transporte] if transporte != "Sin datos" else pd.DataFrame()
    es_calico = "calicó" in str(transporte).lower() or "calico" in str(transporte).lower()

    # 3. RUTAS
    c_ori, c_des = st.columns(2)
    with c_ori:
        origenes = sorted([o for o in df_transporte[col_loc_ori].unique() if o]) if col_loc_ori and not df_transporte.empty else []
        origen = st.selectbox("Desde:", origenes if origenes else ["-"], key=f"ori_{prefijo_key}")
        df_rutas = df_transporte[df_transporte[col_loc_ori] == origen] if origen and origen != "-" else pd.DataFrame()

    with c_des:
        destinos = sorted([d for d in df_rutas[col_loc_des].unique() if d]) if col_loc_des and not df_rutas.empty else []
        destino = st.selectbox("Hacia:", destinos if destinos else ["-"], key=f"des_{prefijo_key}")

    # 4. CARGA
    c_kil, c_bul = st.columns(2)
    with c_kil:
        kilos = st.number_input("Kilos", min_value=0.0, value=10.0, key=f"kg_{prefijo_key}")
        monto = st.number_input("Facturación ($)", min_value=0.0, value=50000.0, key=f"fac_{prefijo_key}")
    with c_bul:
        bultos = st.number_input("Bultos (Calicó)", min_value=0, value=1 if es_calico else 0, disabled=not es_calico, key=f"bul_{prefijo_key}")

    # 5. CÁLCULO
    if not df_rutas.empty and destino != "-":
        fila = df_rutas[df_rutas[col_loc_des] == destino].iloc[0]
        
        if modalidad == "Directo":
            precio = float(fila.get(col_directo, 0)) if pd.notna(fila.get(col_directo)) else 0.0
            flete = precio
            detalle = f"✔️ Flete Directo: ${flete:,.2f}"
        else:
            p_base = float(fila.get(col_base, 0)) if pd.notna(fila.get(col_base)) else 0.0
            p_int = float(fila.get(col_interior, 0)) if pd.notna(fila.get(col_interior)) else 0.0
            
            p_kilo = p_base if p_base > 0 else (p_int if p_int > 0 else 0.0)
            tipo_tarifa = "Base" if p_base > 0 else ("Interior" if p_int > 0 else "Sin tarifa")
            
            flete = kilos * p_kilo
            detalle = f"✔️ Flete ({tipo_tarifa}): {kilos} kg × ${p_kilo:,.2f} = ${flete:,.2f}"
            
        manipuleo = (bultos * float(fila.get(col_manipuleo, 0))) if es_calico and pd.notna(fila.get(col_manipuleo)) else 0.0
        total = flete + manipuleo
        incidencia = (total / monto * 100) if monto > 0 else 0
        
        st.markdown(f"""
        <div class="resultado-card">
            <h2 style="color: white;">Total: ${total:,.2f}</h2>
            <h3 style="color: {color_titulo};">Incidencia: {incidencia:.2f}%</h3>
        </div>
        """, unsafe_allow_html=True)
        st.write(detalle)
        if es_calico: st.write(f"✔️ Manipuleo: {bultos} bultos = ${manipuleo:,.2f}")
    else:
        st.info("Configura la ruta para ver el cálculo.")

# =====================================================================
# MAQUETADO DE LA PANTALLA DIVIDIDA
# =====================================================================
col_izq, col_der = st.columns(2)

with col_izq:
    crear_panel_cotizador("Opción 1", "#a8e6cf") # Títulos en verde agua para diferenciar

with col_der:
    crear_panel_cotizador("Opción 2", "#ff8b94") # Títulos en rojo suave para diferenciar
