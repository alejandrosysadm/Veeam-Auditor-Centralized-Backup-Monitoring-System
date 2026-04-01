import streamlit as st
import pandas as pd
import psycopg2
import plotly.express as px
from datetime import datetime, timedelta
import subprocess
import sys
import warnings
import io  # Necesario para la generación de archivos en memoria

# 1. Silenciar avisos de Pandas
warnings.filterwarnings("ignore", category=UserWarning)

st.set_page_config(
    page_title="Veeam Auditor | Laberit",
    layout="wide",
    page_icon="🛡️"
)

# --- ESTILOS CSS AVANZADOS ---
st.markdown("""
    <style>
    .block-container { padding-top: 1rem; padding-bottom: 1rem; }
    
    /* KPIs con estilo limpio */
    div[data-testid="stMetric"] { 
        background-color: #ffffff; 
        border-left: 5px solid #007bff; 
        border-radius: 5px; 
        box-shadow: 0 2px 4px rgba(0,0,0,0.05); 
    }
    
    /* Contenedor del Log tipo Terminal */
    .log-container {
        background-color: #1a1a1a;
        color: #d4d4d4;
        padding: 15px;
        border-radius: 0 0 8px 8px;
        font-family: 'Consolas', 'Monaco', monospace;
        font-size: 14px;
        line-height: 1.6;
        height: 450px;
        overflow-y: auto;
        border: 1px solid #333;
    }
    
    /* Cabecera del Log */
    .log-header {
        background-color: #333;
        color: white;
        padding: 8px 15px;
        border-radius: 8px 8px 0 0;
        font-weight: bold;
        display: flex;
        justify-content: space-between;
        border: 1px solid #333;
    }

    /* Resaltados específicos dentro del Log */
    .line-error { background-color: #451a1a; color: #ff4b4b; font-weight: bold; padding: 2px 4px; border-radius: 2px; }
    .line-warning { background-color: #423315; color: #ffa500; font-weight: bold; padding: 2px 4px; border-radius: 2px; }
    .line-info { color: #55ccff; font-weight: bold; }
    .line-success-muted { color: #555555; } 
    </style>
    """, unsafe_allow_html=True)

# --- FUNCIONES DE BASE DE DATOS ---
def get_connection():
    return psycopg2.connect(
        dbname="veeam_monitor", 
        user="postgres", 
        password="CONTRASEÑA BBDD", 
        host="127.0.0.1"
    )

def cargar_datos():
    conn = get_connection()
    query = "SELECT id, cliente, job_name, status, fecha, revisado, log_cuerpo FROM backups ORDER BY fecha DESC"
    df = pd.read_sql(query, conn)
    df['fecha'] = pd.to_datetime(df['fecha'])
    conn.close()
    return df

def actualizar_revisado(id_tarea):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("UPDATE backups SET revisado = True WHERE id = %s", (id_tarea,))
    conn.commit()
    conn.close()

# --- LÓGICA DE DATOS ---
df_master = cargar_datos()
lista_clientes = sorted(df_master['cliente'].unique().tolist())

if 'cliente_idx' not in st.session_state:
    st.session_state.cliente_idx = 0

# --- BARRA LATERAL ---
with st.sidebar:
    st.title("🛡️ Veeam Auditor")
    
    if st.button("🔄 Sincronizar Outlook", use_container_width=True):
        with st.spinner("Actualizando registros..."):
            # Nota: En Linux este script debe usar IMAP
            subprocess.run([sys.executable, "sync_veeam.py"])
            st.rerun()
    
    st.markdown("---")
    opcion_fecha = st.radio("Periodo de vista:", ["Hoy y Ayer", "Hoy", "Ayer", "Últimos 7 días", "Todo"], index=0)
    
    st.markdown("---")
    st.subheader("🚀 Navegación de Clientes")
    nav_col1, nav_col2 = st.columns(2)
    if nav_col1.button("⬅️ Anterior", use_container_width=True):
        st.session_state.cliente_idx = (st.session_state.cliente_idx - 1) % len(lista_clientes)
        st.rerun()
    if nav_col2.button("Siguiente ➡️", use_container_width=True):
        st.session_state.cliente_idx = (st.session_state.cliente_idx + 1) % len(lista_clientes)
        st.rerun()

    if st.session_state.cliente_idx >= len(lista_clientes): st.session_state.cliente_idx = 0
    cliente_actual = st.selectbox("Seleccionar Cliente:", options=lista_clientes, index=st.session_state.cliente_idx)
    st.session_state.cliente_idx = lista_clientes.index(cliente_actual)
    
    mostrar_revisados = st.toggle("Ver históricos revisados", value=False)

    # --- SECCIÓN DE EXPORTACIÓN ---
    st.markdown("---")
    st.subheader("📊 Exportar Reporte")
    
    # Preparamos el filtrado de fecha para el DataFrame actual
    hoy = datetime.now().date()
    df_f = df_master[df_master['cliente'] == cliente_actual].copy()

    if not mostrar_revisados: 
        df_f = df_f[df_f['revisado'] == False]

    if opcion_fecha == "Hoy y Ayer":
        df_f = df_f[df_f['fecha'].dt.date.isin([hoy, hoy - timedelta(days=1)])]
    elif opcion_fecha == "Hoy":
        df_f = df_f[df_f['fecha'].dt.date == hoy]
    elif opcion_fecha == "Ayer":
        df_f = df_f[df_f['fecha'].dt.date == (hoy - timedelta(days=1))]
    elif opcion_fecha == "Últimos 7 días":
        df_f = df_f[df_f['fecha'].dt.date >= (hoy - timedelta(days=7))]

    if not df_f.empty:
        # Exportar a CSV
        csv_data = df_f[['cliente', 'job_name', 'status', 'fecha', 'revisado']].to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Descargar CSV",
            data=csv_data,
            file_name=f"Reporte_{cliente_actual}_{hoy}.csv",
            mime="text/csv",
            use_container_width=True
        )

        # Exportar a Excel (requiere openpyxl)
        try:
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_f[['cliente', 'job_name', 'status', 'fecha', 'revisado']].to_excel(writer, index=False)
            
            st.download_button(
                label="📈 Descargar Excel",
                data=buffer.getvalue(),
                file_name=f"Reporte_{cliente_actual}_{hoy}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        except Exception:
            st.warning("Instala openpyxl para exportar a Excel.")

# --- CUERPO PRINCIPAL ---
st.title(f"💼 Cliente: {cliente_actual}")

st.subheader("📋 Resumen de Tareas")

def style_status(val):
    if val == 'Success': return 'background-color: #d4edda; color: #155724;'
    if val == 'Failed': return 'background-color: #f8d7da; color: #721c24;'
    return 'background-color: #fff3cd; color: #856404;'

if not df_f.empty:
    # Tabla de tareas
    seleccion = st.dataframe(
        df_f[['job_name', 'status', 'fecha']].style.applymap(style_status, subset=['status']),
        use_container_width=True, 
        height=250,
        column_config={
            "job_name": st.column_config.TextColumn("Asunto / Job", width="max"),
            "status": st.column_config.TextColumn("Estado", width="small"),
            "fecha": st.column_config.DatetimeColumn("Recibido", format="DD/MM/YY HH:mm")
        },
        on_select="rerun",
        selection_mode="single-row",
        hide_index=True
    )

    st.markdown("---")
    st.subheader("🔍 Inspección de Detalle (Log)")
    
    indices = seleccion.selection.rows
    if indices:
        fila = df_f.iloc[indices[0]]
        log_raw = fila['log_cuerpo'] if fila['log_cuerpo'] else "Sin detalles disponibles."
        
        # --- PROCESADOR VISUAL DE LOGS ---
        lineas = log_raw.split('\n')
        log_hl = ""
        
        criticos = ["ERROR", "FAILED", "EXCEPTION", "TIMEOUT", "COULD NOT", "LOW ON FREE DISK SPACE"]
        avisos = ["WARNING", "GETTING LOW", "DEGRADED", "RETRY"]
        info = ["DESCRIPTION:", "DETAILS:", "TOTAL SIZE", "REPOSITORY", "BACKUP SIZE"]

        for l in lineas:
            l_up = l.upper()
            if any(x in l_up for x in criticos):
                log_hl += f"<span class='line-error'>🚨 {l}</span>\n"
            elif any(x in l_up for x in avisos):
                log_hl += f"<span class='line-warning'>⚠️ {l}</span>\n"
            elif any(x in l_up for x in info):
                log_hl += f"<span class='line-info'>ℹ️ {l}</span>\n"
            elif "SUCCESS" in l_up and ("VM" in l_up or "\t" in l):
                log_hl += f"<span class='line-success-muted'>{l}</span>\n"
            else:
                log_hl += f"{l}\n"

        st.markdown(f'<div class="log-header"><span>📄 {fila["job_name"][:80]}</span><span>ID: {fila["id"]}</span></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="log-container"><pre style="white-space: pre-wrap; color: inherit; background: none; border: none; font-size: 14px;">{log_hl}</pre></div>', unsafe_allow_html=True)
        
        if not fila['revisado']:
            if st.button(f"✅ Marcar como Revisado", use_container_width=True):
                actualizar_revisado(fila['id'])
                st.rerun()
    else:
        st.info("💡 Selecciona una fila arriba para ver el log detallado.")
else:
    st.success("✅ Todo al día. No hay tareas pendientes para los filtros seleccionados.")

# --- MÉTRICAS DE PIE ---
st.markdown("---")
m1, m2, m3, m4 = st.columns([1, 1, 1, 2])
with m1: st.metric("Pendientes", len(df_f[df_f['revisado'] == False]))
with m2: st.metric("Fails", len(df_f[df_f['status'] == 'Failed']))
with m3: st.metric("Cliente actual", f"{st.session_state.cliente_idx + 1} de {len(lista_clientes)}")
with m4:
    if not df_f.empty:
        fig = px.bar(df_f, x='status', color='status', 
                     color_discrete_map={'Success':'#28a745','Failed':'#dc3545','Warning':'#ffc107'}, height=150)
        fig.update_layout(showlegend=False, margin=dict(t=0, b=0, l=0, r=0), paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
        st.plotly_chart(fig, use_container_width=True)