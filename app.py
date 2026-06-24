import streamlit as st
import pandas as pd
from database import (
    SessionLocal, engine, Base,
    Usuario, Proyecto, Proveedor, PartidaPresupuesto, Gasto, Pago, Auditoria,
    CobroCliente
)
import io
import os
import datetime
import hashlib
import json
from sqlalchemy import func, desc
import plotly.express as px
import plotly.graph_objects as go

# ============================================
# CONFIGURACIÓN INICIAL Y FUNCIONES AUXILIARES
# ============================================

Base.metadata.create_all(bind=engine)

def encriptar_password(password: str) -> str:
    return hashlib.sha256(password.encode()).hexdigest()

# Crear usuarios por defecto
db_init = SessionLocal()
if db_init.query(Usuario).count() == 0:
    db_init.add(Usuario(nombre="admin", password_hash=encriptar_password("mara2026"), rol="Gerencia"))
    db_init.add(Usuario(nombre="operario", password_hash=encriptar_password("mara123"), rol="Operario"))
    db_init.add(Usuario(nombre="auxiliar", password_hash=encriptar_password("aux123"), rol="Auxiliar Contable"))
    db_init.commit()
db_init.close()

st.set_page_config(page_title="MARA INGENIEROS - Control de Costos", layout="wide")

# Estilos personalizados
st.markdown("""
    <style>
    /* Sidebar con fondo azul cielo */
    section[data-testid="stSidebar"] {
        background-color: #E3F2FD;  /* Azul cielo claro */
    }
    section[data-testid="stSidebar"] * {
        color: #0C2340;  /* Texto oscuro para contraste */
    }
    /* Título de usuario y rol */
    section[data-testid="stSidebar"] .css-1d391kg {
        color: #0C2340 !important;
        font-weight: bold;
    }
    /* Botón de cerrar sesión */
    section[data-testid="stSidebar"] .stButton button {
        background-color: #FFFFFF;  /* Fondo blanco */
        color: #0C2340;             /* Texto oscuro */
        border: 2px solid #0C2340;  /* Borde azul oscuro */
        border-radius: 8px;
        font-weight: bold;
        padding: 0.5rem 1rem;
        width: 100%;
        transition: all 0.2s;
    }
    section[data-testid="stSidebar"] .stButton button:hover {
        background-color: #0C2340;   /* Fondo azul oscuro al pasar mouse */
        color: #FFFFFF;              /* Texto blanco */
        border-color: #0C2340;
    }
    /* Radio buttons (opciones del menú) */
    section[data-testid="stSidebar"] .stRadio label {
        font-weight: 500;
        padding: 6px 8px;
        border-radius: 6px;
        transition: background 0.2s;
    }
    section[data-testid="stSidebar"] .stRadio label:hover {
        background-color: #BBDEFB;   /* Azul más intenso al pasar mouse */
    }
    /* Ocultar el texto "Navegación" si se desea (opcional) */
    /* section[data-testid="stSidebar"] .stRadio > div:first-child { display: none; } */
    </style>
""", unsafe_allow_html=True)

st.markdown("""
    <style>
    .main-title { font-size:38px !important; font-weight: bold; color: #0C2340; text-align: center; margin-bottom: 5px; }
    .subtitle { font-size:18px !important; color: #5C768D; text-align: center; margin-bottom: 25px; }
    .card-metrica { background-color: #F0F4F8; padding: 18px; border-radius: 10px; border-left: 5px solid #0056B3; margin-bottom: 15px; }
    .card-utilidad { background-color: #E2F0D9; padding: 18px; border-radius: 10px; border-left: 5px solid #385723; margin-bottom: 15px; }
    .block-container { padding-top: 1rem; padding-bottom: 0rem; }
    @media (max-width: 768px) {
        .main-title { font-size: 28px !important; }
        .stColumns { gap: 0.5rem; }
    }
    </style>
""", unsafe_allow_html=True)

# Estado de sesión
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
    st.session_state.carrito = []
    st.session_state.partidas_temp = []
    st.session_state.pago_abierto = None

def cerrar_sesion():
    st.session_state.autenticado = False
    st.session_state.carrito = []
    st.session_state.partidas_temp = []
    st.rerun()

db = SessionLocal()

# --- Funciones de auditoría y eliminación ---
def registrar_auditoria(tabla, registro_id, accion, usuario_id, datos_anteriores=None, datos_nuevos=None):
    audit = Auditoria(
        tabla=tabla,
        registro_id=registro_id,
        accion=accion,
        usuario_id=usuario_id,
        datos_anteriores=json.dumps(datos_anteriores) if datos_anteriores else None,
        datos_nuevos=json.dumps(datos_nuevos) if datos_nuevos else None
    )
    db.add(audit)
    db.commit()

def eliminar_registro(modelo, registro_id, tabla_nombre, usuario_id):
    registro = db.query(modelo).filter(modelo.id == registro_id).first()
    if registro:
        db.delete(registro)
        db.commit()
        registrar_auditoria(tabla_nombre, registro_id, "delete", usuario_id)
        return True
    return False

# --- Carga de LPU ---
@st.cache_data
def cargar_lpu():
    posibles_nombres = ["LPU Lista de Precios.xlsx", "LPU Lista de Precios.xlsx - Hoja1.csv"]
    for nombre in posibles_nombres:
        if os.path.exists(nombre):
            try:
                df = pd.read_csv(nombre) if nombre.endswith('.csv') else pd.read_excel(nombre)
                df = df.rename(columns={"Bomcode": "CÓDIGO", "Descripción WOM": "DESCRIPCIÓN", "UM": "UNIDAD", "Price": "VALOR UNITARIO"})
                df["CÓDIGO"] = df["CÓDIGO"].astype(str).str.replace(".0", "", regex=False)
                df["DESCRIPCIÓN"] = df["DESCRIPCIÓN"].astype(str)
                return df
            except:
                pass
    # Datos de ejemplo
    return pd.DataFrame({
        "CÓDIGO": ["8828119377", "8828119094", "8828119411", "8828120200"],
        "DESCRIPCIÓN": ["Traslado de Sondas. Incluye Transporte entre Estaciones", "Excavación para Instalación de Tubería EMT/PVC", "Instalación de nuevo sector, incluye hasta 3 elementos", "Apertura y Cierre de Cañuelas incluye Fusible"],
        "UNIDAD": ["U", "ML", "Por sector", "U"],
        "VALOR UNITARIO": [127500.0, 45000.0, 406997.0, 51840.0]
    })

df_lpu = cargar_lpu()
df_lpu["Selector"] = df_lpu["CÓDIGO"] + " - " + df_lpu["DESCRIPCIÓN"]

def get_usuario_actual():
    if "usuario_actual" in st.session_state:
        return db.query(Usuario).filter(Usuario.nombre == st.session_state.usuario_actual).first()
    return None

def calcular_kpi_proyecto(proyecto_id):
    ingresos = db.query(func.sum(PartidaPresupuesto.total)).filter(PartidaPresupuesto.proyecto_id == proyecto_id).scalar() or 0
    costos = db.query(func.sum(Gasto.valor_total)).filter(Gasto.proyecto_id == proyecto_id).scalar() or 0
    pagos = db.query(func.sum(Pago.monto)).join(Gasto).filter(Gasto.proyecto_id == proyecto_id).scalar() or 0
    saldo_pendiente = costos - pagos
    avance = (costos / ingresos * 100) if ingresos > 0 else 0
    rentabilidad = ((ingresos - costos) / ingresos * 100) if ingresos > 0 else 0
    return {
        "ingresos": ingresos,
        "costos": costos,
        "pagos": pagos,
        "saldo_pendiente": saldo_pendiente,
        "avance": avance,
        "rentabilidad": rentabilidad
    }

@st.cache_data
def cargar_sitios_excel():
    ruta_excel = os.path.join(os.getcwd(), "sitios.xlsx")
    if os.path.exists(ruta_excel):
        try:
            df = pd.read_excel(ruta_excel)
            df.rename(columns={
                "Wom_Site_Code": "codigo",
                "Site_Name": "nombre",
                "torrero": "torrero",
                "codigo_torrero": "codigo_torrero",
                "Latitude": "latitud",
                "Longitude": "longitud",
                "Address": "direccion",
                "Department": "departamento",
                "Municipality": "municipio",
                "Regional FM&R": "regional",
                "Tipo de Energía": "tipo_energia"
            }, inplace=True)
            df["codigo"] = df["codigo"].astype(str)
            df["nombre"] = df["nombre"].astype(str)
            df["torrero"] = df["torrero"].astype(str)
            df["codigo_torrero"] = df["codigo_torrero"].astype(str)
            df["latitud"] = pd.to_numeric(df["latitud"], errors="coerce")
            df["longitud"] = pd.to_numeric(df["longitud"], errors="coerce")
            df.fillna("", inplace=True)
            return df
        except Exception as e:
            st.error(f"Error al cargar sitios.xlsx: {e}")
            return pd.DataFrame()
    else:
        st.warning("Archivo sitios.xlsx no encontrado. El autocompletado no estará disponible.")
        return pd.DataFrame()

# ============================================
# LOGIN
# ============================================

st.markdown("<div class='main-title'> MARA INGENIEROS S.A.S.</div>", unsafe_allow_html=True)
st.markdown("<div class='subtitle'> INGENIERÍA DE CALIDAD </div>", unsafe_allow_html=True)

if not st.session_state.autenticado:
    st.markdown("---")
    col_login_1, col_login_2, col_login_3 = st.columns([1, 1, 1])
    with col_login_2:
        st.markdown("### 🔑 Ingreso al Sistema")
        u = st.text_input("Usuario")
        p = st.text_input("Contraseña", type="password")
        if st.button(" Iniciar Sesión", use_container_width=True):
            p_hash = encriptar_password(p)
            user = db.query(Usuario).filter(Usuario.nombre == u).first()
            if user and (user.password_hash == p_hash or user.password_hash == p):
                st.session_state.autenticado = True
                st.session_state.usuario_actual = user.nombre
                st.session_state.rol_actual = user.rol
                st.rerun()
            else:
                st.error("Credenciales de acceso incorrectas")
    st.stop()

# ============================================
# SIDEBAR CON ÍCONOS Y PERMISOS
# ============================================

st.sidebar.title(f"👤 {st.session_state.usuario_actual.upper()}")
st.sidebar.markdown(f"**Rol:** `{st.session_state.rol_actual}`")

# Botón de cerrar sesión (con estilo mejorado)
if st.sidebar.button("🔒 Cerrar Sesión", use_container_width=True):
    cerrar_sesion()

st.sidebar.markdown("---")  # Separador visual

# Definir opciones según rol
rol = st.session_state.rol_actual
if rol == "Operario":
    opciones_base = ["Proyectos", "Partidas Presupuesto", "Gastos", "Pagos", "Proveedores", "Reportes"]
elif rol == "Auxiliar Contable":
    opciones_base = ["Proyectos", "Partidas Presupuesto", "Gastos", "Pagos", "Proveedores", "Conciliación", "Reportes"]
else:  # Gerencia
    opciones_base = ["Proyectos", "Partidas Presupuesto", "Gastos", "Pagos", "Proveedores", "Conciliación", "Reportes", "Dashboard", "Usuarios"]

# Diccionario de íconos
iconos = {
    "Proyectos": "🏗️",
    "Partidas Presupuesto": "📋",
    "Gastos": "💸",
    "Pagos": "💳",
    "Proveedores": "🏢",
    "Conciliación": "✅",
    "Reportes": "📊",
    "Dashboard": "📈",
    "Usuarios": "👥"
}

# Función para mostrar opción con ícono
def opcion_con_icono(opcion):
    return f"{iconos.get(opcion, '')} {opcion}"

# Menú de navegación
menu = st.sidebar.radio(
    "Navegación",
    opciones_base,
    format_func=opcion_con_icono
)
# ============================================
# FUNCIONES DE CADA MÓDULO
# ============================================

def pagina_dashboard():
    st.markdown("## 📊 Panel de Control Inteligente")
    st.markdown("---")
    
    if st.session_state.rol_actual not in ["Gerencia", "Auxiliar Contable"]:
        st.warning("No tienes permisos para ver el Dashboard.")
        return

    with st.expander("🔍 Filtros del Dashboard", expanded=True):
        col_f1, col_f2, col_f3 = st.columns(3)
        
        with col_f1:
            proyectos = db.query(Proyecto).all()
            opciones_proy = ["Todos"] + [p.nombre for p in proyectos]
            proyecto_seleccionado = st.selectbox("Proyecto", opciones_proy, key="dash_proyecto")
        
        with col_f2:
            fecha_desde = st.date_input("Fecha desde", 
                                        datetime.date.today() - datetime.timedelta(days=90),
                                        key="dash_desde")
            fecha_hasta = st.date_input("Fecha hasta", 
                                        datetime.date.today(),
                                        key="dash_hasta")
        
        with col_f3:
            agrupacion = st.selectbox("Agrupar por", ["Mes", "Trimestre", "Año"], key="dash_agrupacion")
            tipo_analisis = st.multiselect(
                "Mostrar gráficos",
                ["Distribución por Categoría", "Rentabilidad por Proyecto", 
                 "Evolución Temporal", "Ingresos vs Costos"],
                default=["Distribución por Categoría", "Rentabilidad por Proyecto", "Evolución Temporal"]
            )

    if proyecto_seleccionado == "Todos":
        proyectos_ids = [p.id for p in db.query(Proyecto).all()]
    else:
        proyecto_obj = db.query(Proyecto).filter(Proyecto.nombre == proyecto_seleccionado).first()
        proyectos_ids = [proyecto_obj.id] if proyecto_obj else []

    def aplicar_filtros_fecha(query, tabla_fecha):
        if fecha_desde:
            query = query.filter(tabla_fecha >= fecha_desde)
        if fecha_hasta:
            query = query.filter(tabla_fecha <= fecha_hasta + datetime.timedelta(days=1))
        return query

    query_ingresos = db.query(func.sum(PartidaPresupuesto.total))
    if proyectos_ids:
        query_ingresos = query_ingresos.filter(PartidaPresupuesto.proyecto_id.in_(proyectos_ids))
    query_ingresos = aplicar_filtros_fecha(query_ingresos, PartidaPresupuesto.created_at)
    total_ingresos = query_ingresos.scalar() or 0

    query_costos = db.query(func.sum(Gasto.valor_total))
    if proyectos_ids:
        query_costos = query_costos.filter(Gasto.proyecto_id.in_(proyectos_ids))
    query_costos = aplicar_filtros_fecha(query_costos, Gasto.created_at)
    total_costos = query_costos.scalar() or 0

    query_pagos = db.query(func.sum(Pago.monto)).join(Gasto)
    if proyectos_ids:
        query_pagos = query_pagos.filter(Gasto.proyecto_id.in_(proyectos_ids))
    query_pagos = aplicar_filtros_fecha(query_pagos, Pago.fecha)
    total_pagos = query_pagos.scalar() or 0

    rentabilidad_global = ((total_ingresos - total_costos) / total_ingresos * 100) if total_ingresos > 0 else 0
    avance_global = (total_costos / total_ingresos * 100) if total_ingresos > 0 else 0

    mes_actual = datetime.date.today().replace(day=1)
    mes_anterior = (mes_actual - datetime.timedelta(days=1)).replace(day=1)
    
    query_costos_mes_actual = db.query(func.sum(Gasto.valor_total))
    if proyectos_ids:
        query_costos_mes_actual = query_costos_mes_actual.filter(Gasto.proyecto_id.in_(proyectos_ids))
    query_costos_mes_actual = query_costos_mes_actual.filter(Gasto.created_at >= mes_actual)
    query_costos_mes_actual = query_costos_mes_actual.filter(Gasto.created_at <= fecha_hasta + datetime.timedelta(days=1) if fecha_hasta else Gasto.created_at)
    costos_mes_actual = query_costos_mes_actual.scalar() or 0
    
    query_costos_mes_anterior = db.query(func.sum(Gasto.valor_total))
    if proyectos_ids:
        query_costos_mes_anterior = query_costos_mes_anterior.filter(Gasto.proyecto_id.in_(proyectos_ids))
    query_costos_mes_anterior = query_costos_mes_anterior.filter(Gasto.created_at >= mes_anterior)
    query_costos_mes_anterior = query_costos_mes_anterior.filter(Gasto.created_at < mes_actual)
    costos_mes_anterior = query_costos_mes_anterior.scalar() or 0
    variacion_costos = ((costos_mes_actual - costos_mes_anterior) / costos_mes_anterior * 100) if costos_mes_anterior > 0 else 0

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("💰 Ingresos Presupuestados", f"${total_ingresos:,.0f}")
    col2.metric("📉 Costos Reales", f"${total_costos:,.0f}", 
                delta=f"{avance_global:.1f}% ejecutado")
    col3.metric("💵 Pagos Realizados", f"${total_pagos:,.0f}",
                delta=f"{variacion_costos:+.1f}% vs mes anterior")
    col4.metric("📈 Rentabilidad Real", f"{rentabilidad_global:.1f}%", 
                delta_color="normal" if rentabilidad_global > 0 else "inverse")

    st.markdown("### 📈 Análisis Dinámico")
    proyectos_filtrados = db.query(Proyecto).filter(Proyecto.id.in_(proyectos_ids)).all() if proyectos_ids else []

    if "Distribución por Categoría" in tipo_analisis:
        col1g, col2g = st.columns(2)
        
        query_cat = db.query(Gasto.categoria, func.sum(Gasto.valor_total))
        if proyectos_ids:
            query_cat = query_cat.filter(Gasto.proyecto_id.in_(proyectos_ids))
        query_cat = aplicar_filtros_fecha(query_cat, Gasto.created_at)
        gastos_cat = query_cat.group_by(Gasto.categoria).all()
        
        if gastos_cat and len(gastos_cat) > 1:
            df_cat = pd.DataFrame(gastos_cat, columns=["Categoría", "Total"])
            fig = px.pie(df_cat, values="Total", names="Categoría", 
                        title="Distribución de Costos por Categoría",
                        color_discrete_sequence=px.colors.qualitative.Set3)
            fig.update_traces(textposition='inside', textinfo='percent+label')
            col1g.plotly_chart(fig, use_container_width=True)
        elif gastos_cat and len(gastos_cat) == 1:
            col1g.info("📊 Solo hay una categoría de gastos con los filtros actuales.")
            col1g.metric("Categoría única", f"{gastos_cat[0][0]}", f"${gastos_cat[0][1]:,.0f}")
        else:
            col1g.info("No hay datos de costos para los filtros seleccionados.")
        
        if "Rentabilidad por Proyecto" in tipo_analisis:
            proyectos_data = []
            for p in proyectos_filtrados:
                query_ing = db.query(func.sum(PartidaPresupuesto.total)).filter(PartidaPresupuesto.proyecto_id == p.id)
                query_ing = aplicar_filtros_fecha(query_ing, PartidaPresupuesto.created_at)
                ingresos = query_ing.scalar() or 0
                
                query_cost = db.query(func.sum(Gasto.valor_total)).filter(Gasto.proyecto_id == p.id)
                query_cost = aplicar_filtros_fecha(query_cost, Gasto.created_at)
                costos = query_cost.scalar() or 0
                
                rentabilidad = ((ingresos - costos) / ingresos * 100) if ingresos > 0 else 0
                proyectos_data.append({
                    "Proyecto": p.nombre, 
                    "Rentabilidad": rentabilidad,
                    "Ingresos": ingresos,
                    "Costos": costos
                })
            
            if proyectos_data:
                df_rent = pd.DataFrame(proyectos_data)
                df_rent = df_rent.sort_values("Rentabilidad", ascending=True)
                fig2 = px.bar(df_rent, 
                             x="Proyecto", 
                             y="Rentabilidad",
                             title="Rentabilidad Real por Proyecto",
                             text_auto='.1f',
                             color="Rentabilidad",
                             color_continuous_scale=["red", "yellow", "green"],
                             range_color=[-20, 40])
                fig2.update_traces(textposition='outside')
                fig2.update_layout(yaxis_title="Rentabilidad (%)")
                col2g.plotly_chart(fig2, use_container_width=True)
            else:
                col2g.info("No hay datos de proyectos para los filtros seleccionados.")

    if "Evolución Temporal" in tipo_analisis:
        st.markdown("### 📉 Evolución de Costos en el Tiempo")
        
        if agrupacion == "Mes":
            group_by = func.strftime("%Y-%m", Gasto.created_at)
        elif agrupacion == "Trimestre":
            group_by = func.strftime("%Y-Q", Gasto.created_at)
        else:
            group_by = func.strftime("%Y", Gasto.created_at)
        
        query_evol = db.query(group_by.label("periodo"), func.sum(Gasto.valor_total).label("total"))
        if proyectos_ids:
            query_evol = query_evol.filter(Gasto.proyecto_id.in_(proyectos_ids))
        query_evol = aplicar_filtros_fecha(query_evol, Gasto.created_at)
        gastos_evol = query_evol.group_by("periodo").order_by("periodo").all()
        
        if gastos_evol and len(gastos_evol) > 1:
            df_evol = pd.DataFrame(gastos_evol, columns=["Período", "Total"])
            fig3 = px.line(df_evol, x="Período", y="Total", 
                          title=f"Evolución de Costos por {agrupacion}",
                          markers=True)
            fig3.update_traces(line=dict(width=3))
            fig3.update_layout(yaxis_title="Costos ($)")
            st.plotly_chart(fig3, use_container_width=True)
        else:
            st.info("No hay suficientes datos temporales para mostrar evolución con los filtros actuales.")

    if "Ingresos vs Costos" in tipo_analisis:
        st.markdown("### 📊 Comparativa Ingresos vs Costos por Proyecto")
        
        proyectos_comp = []
        for p in proyectos_filtrados:
            query_ing = db.query(func.sum(PartidaPresupuesto.total)).filter(PartidaPresupuesto.proyecto_id == p.id)
            query_ing = aplicar_filtros_fecha(query_ing, PartidaPresupuesto.created_at)
            ingresos = query_ing.scalar() or 0
            
            query_cost = db.query(func.sum(Gasto.valor_total)).filter(Gasto.proyecto_id == p.id)
            query_cost = aplicar_filtros_fecha(query_cost, Gasto.created_at)
            costos = query_cost.scalar() or 0
            
            if ingresos > 0 or costos > 0:
                proyectos_comp.append({
                    "Proyecto": p.nombre,
                    "Ingresos": ingresos,
                    "Costos": costos
                })
        
        if proyectos_comp:
            df_comp = pd.DataFrame(proyectos_comp)
            df_comp = df_comp.sort_values("Ingresos", ascending=True)
            
            fig4 = go.Figure()
            fig4.add_trace(go.Bar(
                x=df_comp["Proyecto"],
                y=df_comp["Ingresos"],
                name="Ingresos",
                marker_color="#2E86C1"
            ))
            fig4.add_trace(go.Bar(
                x=df_comp["Proyecto"],
                y=df_comp["Costos"],
                name="Costos",
                marker_color="#E74C3C"
            ))
            fig4.update_layout(
                title="Comparativa Ingresos vs Costos por Proyecto",
                barmode='group',
                yaxis_title="Monto ($)",
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
            )
            st.plotly_chart(fig4, use_container_width=True)
        else:
            st.info("No hay datos suficientes para comparar ingresos vs costos con los filtros actuales.")

    st.markdown("### 📋 Detalle Financiero por Proyecto")
    
    data = []
    for p in proyectos_filtrados:
        query_ing = db.query(func.sum(PartidaPresupuesto.total)).filter(PartidaPresupuesto.proyecto_id == p.id)
        query_ing = aplicar_filtros_fecha(query_ing, PartidaPresupuesto.created_at)
        ingresos = query_ing.scalar() or 0
        
        query_cost = db.query(func.sum(Gasto.valor_total)).filter(Gasto.proyecto_id == p.id)
        query_cost = aplicar_filtros_fecha(query_cost, Gasto.created_at)
        costos = query_cost.scalar() or 0
        
        rentabilidad = ((ingresos - costos) / ingresos * 100) if ingresos > 0 else 0
        ejecucion = (costos / ingresos * 100) if ingresos > 0 else 0
        
        if rentabilidad > 20:
            estado = "🟢 Saludable"
        elif rentabilidad > 10:
            estado = "🟡 Aceptable"
        elif rentabilidad > 0:
            estado = "🟠 Riesgo"
        else:
            estado = "🔴 Crítico"
        
        data.append({
            "Proyecto": p.nombre,
            "Cliente": p.cliente,
            "Ingresos": ingresos,
            "Costos": costos,
            "Rentabilidad %": rentabilidad,
            "Ejecución %": ejecucion,
            "Estado": estado
        })
    
    if data:
        df_resumen = pd.DataFrame(data)
        st.dataframe(
            df_resumen,
            use_container_width=True,
            column_config={
                "Proyecto": "Proyecto",
                "Cliente": "Cliente",
                "Ingresos": st.column_config.NumberColumn("Ingresos", format="$%,.0f"),
                "Costos": st.column_config.NumberColumn("Costos", format="$%,.0f"),
                "Rentabilidad %": st.column_config.NumberColumn("Rentabilidad", format="%.1f%%"),
                "Ejecución %": st.column_config.NumberColumn("Ejecución", format="%.1f%%"),
                "Estado": "Estado"
            }
        )
    else:
        st.info("No hay datos para mostrar con los filtros seleccionados.")

def pagina_proyectos():
    st.markdown("## 🏗️ Gestión de Proyectos")
    st.markdown("---")
    usuario = get_usuario_actual()
    
    # Formulario nuevo proyecto
    with st.expander("➕ Nuevo Proyecto", expanded=False):
        df_sitios = cargar_sitios_excel()
        if "sitio_seleccionado" not in st.session_state:
            st.session_state.sitio_seleccionado = None

        with st.form("nuevo_proyecto"):
            cliente = st.text_input("Cliente", "WOM")
            
            if "WOM" in cliente.upper():
                st.markdown("#### 🏢 Buscar sitio en base de datos")
                if not df_sitios.empty:
                    busqueda = st.text_input(
                        "🔍 Buscar por código o nombre",
                        placeholder="Ej: JGUA_00024 o Cuatro Vias",
                        key="busqueda_sitio"
                    )
                    
                    if busqueda:
                        filtro = df_sitios[
                            df_sitios["codigo"].str.contains(busqueda, case=False, na=False) |
                            df_sitios["nombre"].str.contains(busqueda, case=False, na=False)
                        ]
                    else:
                        filtro = df_sitios.head(100)
                    
                    if not filtro.empty:
                        opciones = filtro.apply(
                            lambda row: f"{row['codigo']} - {row['nombre']}", axis=1
                        ).tolist()
                        opciones.insert(0, "Seleccione un sitio...")
                        
                        seleccion = st.selectbox(
                            "Seleccione un sitio",
                            options=opciones,
                            index=0,
                            key="selector_sitio"
                        )
                        
                        if seleccion != "Seleccione un sitio...":
                            codigo_seleccionado = seleccion.split(" - ")[0]
                            sitio = df_sitios[df_sitios["codigo"] == codigo_seleccionado].iloc[0]
                            st.session_state.sitio_seleccionado = sitio
                            st.success(f"✅ Sitio seleccionado: **{sitio['nombre']}**")
                        else:
                            st.session_state.sitio_seleccionado = None
                    else:
                        st.info("No se encontraron sitios con esa búsqueda.")
                else:
                    st.warning("No se pudo cargar la base de datos de sitios.")
            else:
                st.session_state.sitio_seleccionado = None

            if st.session_state.sitio_seleccionado is not None:
                sitio = st.session_state.sitio_seleccionado
                nombre_default = sitio.get("nombre", "")
                ubicacion_default = f"{sitio.get('direccion', '')} {sitio.get('municipio', '')} {sitio.get('departamento', '')}".strip()
                latitud_default = float(sitio.get("latitud", 0.0))
                longitud_default = float(sitio.get("longitud", 0.0))
            else:
                nombre_default = ""
                ubicacion_default = "Colombia"
                latitud_default = 0.0
                longitud_default = 0.0

            nombre = st.text_input("Nombre del Proyecto/Sitio", value=nombre_default)
            ubicacion = st.text_input("Ubicación", value=ubicacion_default)
            n_requerimiento = st.text_input("N° de Requerimiento", "Mara_0001")
            acta = st.text_input("Acta de Conciliación", "SIN CONCILIAR")
            estado = st.selectbox("Estado", ["Activo", "Finalizado"])
            latitud = st.number_input("Latitud (opcional)", value=latitud_default, format="%.6f", step=0.000001)
            longitud = st.number_input("Longitud (opcional)", value=longitud_default, format="%.6f", step=0.000001)

            if st.session_state.sitio_seleccionado is not None:
                sitio = st.session_state.sitio_seleccionado
                st.caption(f"**Torrero:** {sitio.get('torrero', 'N/A')} | **Código Torre:** {sitio.get('codigo_torrero', 'N/A')}")

            submitted = st.form_submit_button("Crear Proyecto")
            if submitted and nombre:
                nuevo = Proyecto(
                    nombre=nombre,
                    ubicacion=ubicacion,
                    cliente=cliente,
                    n_requerimiento=n_requerimiento,
                    acta_conciliacion=acta,
                    estado=estado,
                    created_by=usuario.id,
                    latitud=latitud if latitud != 0.0 else None,
                    longitud=longitud if longitud != 0.0 else None
                )
                db.add(nuevo)
                db.commit()
                registrar_auditoria("Proyecto", nuevo.id, "insert", usuario.id)
                st.success("Proyecto creado exitosamente")
                st.session_state.sitio_seleccionado = None
                st.rerun()

    # Filtros y listado
    st.markdown("### 🔍 Buscar Proyectos")
    
    col_search1, col_search2 = st.columns([3, 1])
    with col_search1:
        busqueda_proyecto = st.text_input(
            "Buscar por nombre o cliente",
            placeholder="Ej: CLAMR, WOM, CALDAS...",
            key="busqueda_proyecto"
        )
    with col_search2:
        filtro_estado = st.selectbox(
            "Estado",
            ["Todos", "Activo", "Finalizado"],
            key="filtro_estado_proyecto"
        )
    
    col_fecha1, col_fecha2, col_fecha3 = st.columns([2, 2, 1])
    with col_fecha1:
        fecha_desde = st.date_input(
            "Fecha desde",
            value=None,
            key="fecha_desde_proyecto"
        )
    with col_fecha2:
        fecha_hasta = st.date_input(
            "Fecha hasta",
            value=None,
            key="fecha_hasta_proyecto"
        )
    with col_fecha3:
        st.write("")
        if st.button("🗑️ Limpiar fechas", use_container_width=True):
            st.session_state.fecha_desde_proyecto = None
            st.session_state.fecha_hasta_proyecto = None
            st.rerun()
                
    query = db.query(Proyecto)
    if busqueda_proyecto:
        query = query.filter(
            (Proyecto.nombre.contains(busqueda_proyecto)) |
            (Proyecto.cliente.contains(busqueda_proyecto)) |
            (Proyecto.ubicacion.contains(busqueda_proyecto))
        )
    if filtro_estado != "Todos":
        query = query.filter(Proyecto.estado == filtro_estado)
    
    proyectos = query.order_by(desc(Proyecto.id)).all()
    
    if not proyectos:
        st.info("No hay proyectos que coincidan con los filtros.")
    else:
        st.markdown(f"**{len(proyectos)}** proyectos encontrados.")
        
        for p in proyectos:
            kpi = calcular_kpi_proyecto(p.id)
            
            with st.container(border=True):
                col_info1, col_info2, col_info3, col_info4 = st.columns([3, 1.5, 1.5, 1])
                with col_info1:
                    st.markdown(f"""
                    <div style="font-size:14px; font-weight:bold; color:#0C2340;">
                        {p.nombre}
                    </div>
                    <div style="font-size:13px; color:#5C768D;">
                        {p.ubicacion} · {p.cliente} · {p.estado}
                    </div>
                    """, unsafe_allow_html=True)
                with col_info2:
                    st.markdown(f"<div style='font-size:14px;'><b>Ingresos:</b> ${kpi['ingresos']:,.0f}</div>", unsafe_allow_html=True)
                with col_info3:
                    st.markdown(f"<div style='font-size:14px;'><b>Costos:</b> ${kpi['costos']:,.0f}</div>", unsafe_allow_html=True)
                with col_info4:
                    color = "green" if kpi['rentabilidad'] >= 20 else "orange" if kpi['rentabilidad'] >= 10 else "red"
                    st.markdown(f"<div style='font-size:14px;'><b>Rentabilidad:</b> <span style='color:{color};'>{kpi['rentabilidad']:.1f}%</span></div>", unsafe_allow_html=True)

                col_btn1, col_btn2, col_btn3, col_btn4, col_btn5 = st.columns([1, 1, 1, 1, 2])
                with col_btn1:
                    if st.session_state.rol_actual in ["Gerencia", "Auxiliar Contable"]:
                        if st.button("✏️ Editar", key=f"edit_{p.id}", use_container_width=True):
                            st.session_state.proyecto_editar = p.id
                            st.rerun()
                with col_btn2:
                    if st.session_state.rol_actual in ["Gerencia", "Auxiliar Contable"]:
                        if st.button("🗑️ Eliminar", key=f"delete_{p.id}", use_container_width=True):
                            partidas = db.query(PartidaPresupuesto).filter(PartidaPresupuesto.proyecto_id == p.id).count()
                            gastos = db.query(Gasto).filter(Gasto.proyecto_id == p.id).count()
                            if partidas > 0 or gastos > 0:
                                st.error(f"No se puede eliminar porque tiene {partidas} partidas y {gastos} gastos asociados.")
                            else:
                                if st.button(f"Confirmar", key=f"confirm_delete_{p.id}"):
                                    db.delete(p)
                                    db.commit()
                                    registrar_auditoria("Proyecto", p.id, "delete", usuario.id)
                                    st.success(f"Proyecto {p.nombre} eliminado")
                                    st.rerun()
                with col_btn3:
                    if p.latitud and p.longitud:
                        maps_url = f"https://www.google.com/maps?q={p.latitud},{p.longitud}"
                        st.link_button("🗺️ Mapa", maps_url, use_container_width=True)
                    else:
                        st.write("")
                with col_btn4:
                    if st.button(f"📊 Detalle", key=f"btn_detalle_{p.id}", use_container_width=True):
                        st.session_state[f"show_detail_{p.id}"] = not st.session_state.get(f"show_detail_{p.id}", False)
                        st.rerun()
                with col_btn5:
                    st.write("")
                
                if "proyecto_editar" in st.session_state and st.session_state.proyecto_editar == p.id:
                    with st.form(f"edit_proy_{p.id}"):
                        nuevo_nombre = st.text_input("Nombre", value=p.nombre)
                        nueva_ubicacion = st.text_input("Ubicación", value=p.ubicacion)
                        nuevo_cliente = st.text_input("Cliente", value=p.cliente)
                        nuevo_req = st.text_input("N° Requerimiento", value=p.n_requerimiento)
                        nueva_acta = st.text_input("Acta", value=p.acta_conciliacion)
                        nuevo_estado = st.selectbox("Estado", ["Activo", "Finalizado"], index=["Activo", "Finalizado"].index(p.estado))
                        nueva_latitud = st.number_input("Latitud (opcional)", value=p.latitud or 0.0, format="%.6f", step=0.000001)
                        nueva_longitud = st.number_input("Longitud (opcional)", value=p.longitud or 0.0, format="%.6f", step=0.000001)
                        submit_edit = st.form_submit_button("Guardar Cambios")
                        if submit_edit:
                            cambios = {}
                            if p.nombre != nuevo_nombre: cambios["nombre"] = (p.nombre, nuevo_nombre)
                            if p.ubicacion != nueva_ubicacion: cambios["ubicacion"] = (p.ubicacion, nueva_ubicacion)
                            if p.cliente != nuevo_cliente: cambios["cliente"] = (p.cliente, nuevo_cliente)
                            if p.n_requerimiento != nuevo_req: cambios["n_requerimiento"] = (p.n_requerimiento, nuevo_req)
                            if p.acta_conciliacion != nueva_acta: cambios["acta_conciliacion"] = (p.acta_conciliacion, nueva_acta)
                            if p.estado != nuevo_estado: cambios["estado"] = (p.estado, nuevo_estado)
                            nueva_lat = nueva_latitud if nueva_latitud != 0.0 else None
                            nueva_lon = nueva_longitud if nueva_longitud != 0.0 else None
                            if p.latitud != nueva_lat: cambios["latitud"] = (p.latitud, nueva_lat)
                            if p.longitud != nueva_lon: cambios["longitud"] = (p.longitud, nueva_lon)
                            if cambios:
                                p.nombre = nuevo_nombre
                                p.ubicacion = nueva_ubicacion
                                p.cliente = nuevo_cliente
                                p.n_requerimiento = nuevo_req
                                p.acta_conciliacion = nueva_acta
                                p.estado = nuevo_estado
                                p.latitud = nueva_lat
                                p.longitud = nueva_lon
                                db.commit()
                                registrar_auditoria("Proyecto", p.id, "update", usuario.id,
                                                    datos_anteriores={k: v[0] for k,v in cambios.items()},
                                                    datos_nuevos={k: v[1] for k,v in cambios.items()})
                                st.success("Proyecto actualizado")
                                del st.session_state.proyecto_editar
                                st.rerun()
                
                if st.session_state.get(f"show_detail_{p.id}", False):
                    st.markdown("---")
                    st.markdown(f"### 📊 Detalle Financiero - {p.nombre}")
                    if p.latitud and p.longitud:
                        maps_url = f"https://www.google.com/maps?q={p.latitud},{p.longitud}"
                        st.markdown(f"📍 [Ver ubicación en Google Maps]({maps_url}){{:target='_blank'}}")
                    
                    partidas = db.query(PartidaPresupuesto).filter(PartidaPresupuesto.proyecto_id == p.id).all()
                    gastos = db.query(Gasto).filter(Gasto.proyecto_id == p.id).all()
                    
                    st.markdown("#### 💰 Ingresos (Partidas Presupuestales)")
                    if partidas:
                        df_partidas = pd.DataFrame([{
                            "Categoría": part.categoria,
                            "Descripción": part.descripcion,
                            "Cantidad": part.cantidad,
                            "Valor Unitario": part.valor_unitario,
                            "Total": part.total
                        } for part in partidas])
                        st.dataframe(df_partidas, use_container_width=True)
                        total_ingresos = df_partidas["Total"].sum()
                        st.metric("Total Ingresos", f"${total_ingresos:,.0f}")
                        
                        if len(partidas) > 1:
                            df_ing_cat = df_partidas.groupby("Categoría")["Total"].sum().reset_index()
                            fig_ing = px.pie(df_ing_cat, values="Total", names="Categoría", 
                                            title="Distribución de Ingresos por Categoría",
                                            color_discrete_sequence=px.colors.qualitative.Pastel)
                            st.plotly_chart(fig_ing, use_container_width=True)
                    else:
                        st.info("No hay partidas presupuestales registradas para este proyecto.")
                    
                    st.markdown("---")
                    
                    st.markdown("#### 📉 Costos (Gastos Registrados)")
                    if gastos:
                        datos_gastos = []
                        for g in gastos:
                            pagos = db.query(Pago).filter(Pago.gasto_id == g.id).all()
                            total_pagado = sum(p.monto for p in pagos)
                            proveedor = db.query(Proveedor).filter(Proveedor.id == g.proveedor_id).first()
                            datos_gastos.append({
                                "Concepto": g.concepto,
                                "Categoría": g.categoria,
                                "Cantidad": g.cantidad,
                                "Valor Unitario": g.valor_unitario,
                                "Total": g.valor_total,
                                "Pagado": total_pagado,
                                "Saldo": g.valor_total - total_pagado,
                                "Estado": g.estado_pago,
                                "Proveedor": proveedor.nombre if proveedor else "N/A"
                            })
                        df_gastos = pd.DataFrame(datos_gastos)
                        st.dataframe(df_gastos, use_container_width=True)
                        total_costos = df_gastos["Total"].sum()
                        total_pagado = df_gastos["Pagado"].sum()
                        total_saldo = df_gastos["Saldo"].sum()
                        
                        col_c1, col_c2, col_c3 = st.columns(3)
                        col_c1.metric("Total Costos", f"${total_costos:,.0f}")
                        col_c2.metric("Total Pagado", f"${total_pagado:,.0f}")
                        col_c3.metric("Saldo Pendiente", f"${total_saldo:,.0f}", delta_color="inverse")
                        
                        if len(gastos) > 1:
                            df_cost_cat = df_gastos.groupby("Categoría")["Total"].sum().reset_index()
                            fig_cost = px.pie(df_cost_cat, values="Total", names="Categoría", 
                                             title="Distribución de Costos por Categoría",
                                             color_discrete_sequence=px.colors.qualitative.Set3)
                            st.plotly_chart(fig_cost, use_container_width=True)
                    else:
                        st.info("No hay gastos registrados para este proyecto.")
                    
                    st.markdown("---")
                    st.markdown("#### 📊 Resumen Comparativo")
                    col_r1, col_r2, col_r3 = st.columns(3)
                    col_r1.metric("💰 Ingresos Totales", f"${kpi['ingresos']:,.0f}")
                    col_r2.metric("📉 Costos Totales", f"${kpi['costos']:,.0f}")
                    col_r3.metric("📈 Rentabilidad", f"{kpi['rentabilidad']:.1f}%", 
                                 delta_color="normal" if kpi['rentabilidad'] > 0 else "inverse")
                    
                    if st.button("❌ Cerrar Detalle", key=f"btn_cerrar_{p.id}"):
                        st.session_state[f"show_detail_{p.id}"] = False
                        st.rerun()

def pagina_partidas():
    st.markdown("## 📋 Partidas Presupuestales")
    st.markdown("---")
    usuario = get_usuario_actual()
    proyectos = db.query(Proyecto).all()
    if not proyectos:
        st.warning("Cree un proyecto primero.")
        return

    with st.expander("🔍 Filtros de búsqueda", expanded=True):
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            proyecto_seleccionado = st.selectbox("Seleccione Proyecto", proyectos, format_func=lambda x: x.nombre)
        with col_f2:
            busqueda_partida = st.text_input("Buscar por descripción o categoría", placeholder="Ej: Mano de Obra, Excavación...")
    
    if proyecto_seleccionado:
        st.subheader(f"Presupuesto de: {proyecto_seleccionado.nombre}")
        
        query = db.query(PartidaPresupuesto).filter(PartidaPresupuesto.proyecto_id == proyecto_seleccionado.id)
        if busqueda_partida:
            query = query.filter(
                (PartidaPresupuesto.descripcion.contains(busqueda_partida)) |
                (PartidaPresupuesto.categoria.contains(busqueda_partida))
            )
        partidas = query.all()
        
        if partidas:
            st.markdown("### Lista de Partidas")
            for idx, partida in enumerate(partidas):
                col1, col2, col3, col4, col5, col6 = st.columns([2, 2, 1, 1, 1, 0.5])
                col1.write(partida.categoria)
                col2.write(partida.descripcion)
                col3.write(partida.cantidad)
                col4.write(f"${partida.valor_unitario:,.0f}")
                col5.write(f"${partida.total:,.0f}")
                with col6:
                    if st.session_state.rol_actual in ["Gerencia", "Auxiliar Contable"]:
                        if st.button("🗑️", key=f"del_partida_{partida.id}"):
                            cobros = db.query(CobroCliente).filter(CobroCliente.partida_id == partida.id).count()
                            if cobros > 0:
                                st.error(f"No se puede eliminar porque tiene {cobros} cobros asociados.")
                            else:
                                db.delete(partida)
                                db.commit()
                                registrar_auditoria("PartidaPresupuesto", partida.id, "delete", usuario.id)
                                st.success("Partida eliminada")
                                st.rerun()
            
            total_pres = sum(p.total for p in partidas)
            st.metric("Total Ingresos Presupuestados", f"${total_pres:,.0f}")
        else:
            st.info("No hay partidas que coincidan con los filtros para este proyecto.")

        with st.expander("➕ Agregar Partida", expanded=False):
            tipo = st.radio("Origen", ["Buscar en LPU", "Ingreso Manual"], horizontal=True)
            if tipo == "Buscar en LPU":
                busqueda = st.text_input("🔍 Buscar en LPU", key="busq_part")
                if busqueda:
                    df_filt = df_lpu[df_lpu["Selector"].str.contains(busqueda, case=False, na=False)]
                else:
                    df_filt = df_lpu
                if not df_filt.empty:
                    selec = st.selectbox("Seleccione ítem", df_filt["Selector"].tolist())
                    fila = df_lpu[df_lpu["Selector"] == selec].iloc[0]
                    cod, desc, uni, val = fila["CÓDIGO"], fila["DESCRIPCIÓN"], fila["UNIDAD"], fila["VALOR UNITARIO"]
                    st.write(f"**Código:** {cod}  |  **UM:** {uni}  |  **Precio Unitario:** ${val:,.0f}")
                    precio_congelado = val
                else:
                    st.error("No se encontraron coincidencias")
                    precio_congelado = 0.0
                    desc = ""
            else:
                desc = st.text_input("Descripción de la partida")
                precio_congelado = st.number_input("Valor Unitario ($)", min_value=0.0, step=100.0)

            categoria = st.selectbox("Categoría", ["Mano de Obra", "Materiales", "Subcontratos", "Equipos", "Transporte", "Otros"])
            cantidad = st.number_input("Cantidad", min_value=0.0, value=1.0)
            if st.button("Agregar Partida"):
                if desc and precio_congelado > 0:
                    total = cantidad * precio_congelado
                    nueva_partida = PartidaPresupuesto(
                        proyecto_id=proyecto_seleccionado.id,
                        categoria=categoria,
                        descripcion=desc,
                        cantidad=cantidad,
                        valor_unitario=precio_congelado,
                        total=total,
                        created_by=usuario.id
                    )
                    db.add(nueva_partida)
                    db.commit()
                    registrar_auditoria("PartidaPresupuesto", nueva_partida.id, "insert", usuario.id)
                    st.success("Partida agregada al presupuesto")
                    st.rerun()
                else:
                    st.error("Complete todos los campos")

def pagina_gastos():
    st.markdown("## 💸 Registro y Gestión de Gastos")
    st.markdown("---")
    usuario = get_usuario_actual()
    proyectos = db.query(Proyecto).all()
    if not proyectos:
        st.warning("Cree un proyecto primero.")
        return

    with st.expander("🔍 Filtros de búsqueda", expanded=True):
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            proyecto_sel = st.selectbox("Proyecto", proyectos, format_func=lambda x: x.nombre, key="proy_gastos")
        with col_f2:
            busqueda_texto = st.text_input("Buscar por concepto, proveedor o categoría", placeholder="Ej: material, Johán...")
        with col_f3:
            fecha_desde = st.date_input("Fecha desde", datetime.date.today() - datetime.timedelta(days=30))
            fecha_hasta = st.date_input("Fecha hasta", datetime.date.today())

    if proyecto_sel:
        st.info(f"**Cliente:** {proyecto_sel.cliente} | **Requerimiento:** {proyecto_sel.n_requerimiento} | **Acta:** {proyecto_sel.acta_conciliacion}")

        total_ingresos = db.query(func.sum(PartidaPresupuesto.total)).filter(PartidaPresupuesto.proyecto_id == proyecto_sel.id).scalar() or 0
        total_costos = db.query(func.sum(Gasto.valor_total)).filter(Gasto.proyecto_id == proyecto_sel.id).scalar() or 0
        total_pagado = db.query(func.sum(Pago.monto)).join(Gasto).filter(Gasto.proyecto_id == proyecto_sel.id).scalar() or 0
        total_pendiente = total_costos - total_pagado
        rentabilidad = ((total_ingresos - total_costos) / total_ingresos * 100) if total_ingresos > 0 else 0

        col_res1, col_res2, col_res3, col_res4 = st.columns(4)
        col_res1.metric("💰 Ingresos", f"${total_ingresos:,.0f}")
        col_res2.metric("📉 Costos", f"${total_costos:,.0f}")
        col_res3.metric("💵 Pagado", f"${total_pagado:,.0f}")
        col_res4.metric("📈 Rentabilidad", f"{rentabilidad:.1f}%", delta_color="normal")

        col_f4, col_f5, col_f6 = st.columns(3)
        with col_f4:
            filtro_conc = st.selectbox("Conciliado", ["Todos", "Conciliado", "No Conciliado"])
        with col_f5:
            filtro_estado = st.selectbox("Estado de Pago", ["Todos", "Pendiente", "Parcial", "Pagado"])
        with col_f6:
            filtro_saldo = st.selectbox("Saldo Pendiente", ["Todos", "Con saldo > 0", "Saldo cero"])

        query = db.query(Gasto).filter(Gasto.proyecto_id == proyecto_sel.id)
        if filtro_conc == "Conciliado":
            query = query.filter(Gasto.conciliado == True)
        elif filtro_conc == "No Conciliado":
            query = query.filter(Gasto.conciliado == False)
        if filtro_estado != "Todos":
            query = query.filter(Gasto.estado_pago == filtro_estado)
        if busqueda_texto:
            query = query.join(Proveedor, Gasto.proveedor_id == Proveedor.id, isouter=True)
            query = query.filter(
                (Gasto.concepto.contains(busqueda_texto)) |
                (Gasto.categoria.contains(busqueda_texto)) |
                (Proveedor.nombre.contains(busqueda_texto))
            )
        if fecha_desde and fecha_hasta:
            query = query.filter(Gasto.created_at >= fecha_desde)
            query = query.filter(Gasto.created_at <= fecha_hasta + datetime.timedelta(days=1))

        gastos = query.all()

        gastos_filtrados = []
        for g in gastos:
            pagos = db.query(Pago).filter(Pago.gasto_id == g.id).all()
            total_pagado_g = sum(p.monto for p in pagos)
            saldo = g.valor_total - total_pagado_g
            if filtro_saldo == "Con saldo > 0" and saldo <= 0:
                continue
            elif filtro_saldo == "Saldo cero" and saldo > 0:
                continue
            gastos_filtrados.append((g, total_pagado_g, saldo))

        if not gastos_filtrados:
            st.info("No hay gastos que coincidan con los filtros.")
        else:
            for g, pagado, saldo in gastos_filtrados:
                with st.container(border=True):
                    col1, col2, col3, col4 = st.columns([3, 1, 1, 0.5])
                    proveedor = db.query(Proveedor).filter(Proveedor.id == g.proveedor_id).first()
                    col1.markdown(f"**{g.concepto}**  \n*{g.categoria}*  \nProveedor: {proveedor.nombre if proveedor else 'N/A'}")
                    col2.metric("Total", f"${g.valor_total:,.0f}")
                    col3.metric("Saldo", f"${saldo:,.0f}")
                    with col4:
                        if st.session_state.rol_actual in ["Gerencia", "Auxiliar Contable"]:
                            if st.button("🗑️", key=f"del_gasto_{g.id}"):
                                pagos_asociados = db.query(Pago).filter(Pago.gasto_id == g.id).count()
                                if pagos_asociados > 0:
                                    st.error(f"No se puede eliminar porque tiene {pagos_asociados} pagos asociados.")
                                else:
                                    db.delete(g)
                                    db.commit()
                                    registrar_auditoria("Gasto", g.id, "delete", usuario.id)
                                    st.success("Gasto eliminado")
                                    st.rerun()
                    
                    if saldo > 0:
                        with st.expander(f"💳 Registrar Pago - {g.concepto[:50]} (Saldo: ${saldo:,.0f})"):
                            with st.form(f"pago_inline_{g.id}"):
                                tipo = st.selectbox("Tipo", ["Factura", "Anticipo"], key=f"tipo_{g.id}")
                                n_factura = st.text_input("N° Factura", key=f"fact_{g.id}")
                                fecha = st.date_input("Fecha", datetime.date.today(), key=f"fecha_{g.id}")
                                monto = st.number_input("Monto a pagar", min_value=0.0, max_value=saldo, step=1000.0, key=f"monto_{g.id}")
                                observaciones = st.text_area("Observaciones", key=f"obs_{g.id}")
                                submitted = st.form_submit_button("Registrar Pago")
                                if submitted and monto > 0:
                                    nuevo_pago = Pago(
                                        gasto_id=g.id,
                                        tipo=tipo,
                                        numero_factura=n_factura,
                                        fecha=fecha,
                                        monto=monto,
                                        observaciones=observaciones,
                                        created_by=usuario.id
                                    )
                                    db.add(nuevo_pago)
                                    nuevo_total_pagado = pagado + monto
                                    if nuevo_total_pagado >= g.valor_total:
                                        g.estado_pago = "Pagado"
                                    elif nuevo_total_pagado > 0:
                                        g.estado_pago = "Parcial"
                                    db.commit()
                                    registrar_auditoria("Pago", nuevo_pago.id, "insert", usuario.id)
                                    registrar_auditoria("Gasto", g.id, "update", usuario.id,
                                                        datos_anteriores={"estado_pago": g.estado_pago},
                                                        datos_nuevos={"estado_pago": g.estado_pago})
                                    st.success("✅ Pago registrado exitosamente")
                                    st.rerun()

        with st.expander("➕ Nuevo Gasto", expanded=False):
            tipo_item = st.radio("Origen del ítem", ["LPU", "Manual"], horizontal=True)
            if tipo_item == "LPU":
                busqueda = st.text_input("🔍 Buscar en LPU", key="busq_gasto")
                if busqueda:
                    df_filt = df_lpu[df_lpu["Selector"].str.contains(busqueda, case=False, na=False)]
                else:
                    df_filt = df_lpu
                if not df_filt.empty:
                    selec = st.selectbox("Seleccione", df_filt["Selector"].tolist())
                    fila = df_lpu[df_lpu["Selector"] == selec].iloc[0]
                    cod, desc, uni, val = fila["CÓDIGO"], fila["DESCRIPCIÓN"], fila["UNIDAD"], fila["VALOR UNITARIO"]
                    st.write(f"**Precio Unitario:** ${val:,.0f}")
                else:
                    st.error("Sin coincidencias")
                    desc, uni, val = "", "U", 0.0
            else:
                desc = st.text_input("Descripción del gasto")
                uni = st.text_input("Unidad", "U")
                val = st.number_input("Valor Unitario", min_value=0.0, step=100.0)

            if desc and val > 0:
                cantidad = st.number_input("Cantidad", min_value=0.01, value=1.0)
                categoria = st.selectbox("Categoría", ["Materiales", "Mano de Obra", "Subcontratos", "Equipos", "Transporte", "Otros"])
                proveedores = db.query(Proveedor).all()
                opciones_prov = [p.nombre for p in proveedores] + ["➕ Crear nuevo"]
                proveedor_sel = st.selectbox("Proveedor", opciones_prov)
                if proveedor_sel == "➕ Crear nuevo":
                    with st.form("nuevo_proveedor"):
                        nom = st.text_input("Nombre")
                        nit = st.text_input("NIT")
                        contacto = st.text_input("Contacto")
                        telefono = st.text_input("Teléfono")
                        submit_prov = st.form_submit_button("Crear Proveedor")
                        if submit_prov and nom:
                            nuevo_prov = Proveedor(nombre=nom, nit=nit, contacto=contacto, telefono=telefono, created_by=usuario.id)
                            db.add(nuevo_prov)
                            db.commit()
                            registrar_auditoria("Proveedor", nuevo_prov.id, "insert", usuario.id)
                            st.success("Proveedor creado")
                            st.rerun()
                    proveedor_id = None
                else:
                    proveedor_obj = db.query(Proveedor).filter(Proveedor.nombre == proveedor_sel).first()
                    proveedor_id = proveedor_obj.id if proveedor_obj else None

                if st.button("➕ Agregar al carrito"):
                    total_item = cantidad * val
                    st.session_state.carrito.append({
                        "proyecto_id": proyecto_sel.id,
                        "concepto": desc,
                        "categoria": categoria,
                        "unidad": uni,
                        "cantidad": cantidad,
                        "valor_unitario": val,
                        "valor_total": total_item,
                        "proveedor_id": proveedor_id
                    })
                    st.success("Ítem agregado al carrito")
                    st.rerun()

        if st.session_state.carrito:
            st.markdown("### 🛒 Carrito de Gastos")
            df_cart = pd.DataFrame(st.session_state.carrito)
            st.dataframe(df_cart[["concepto", "categoria", "cantidad", "valor_unitario", "valor_total"]], use_container_width=True)
            st.metric("Total Acumulado", f"${df_cart['valor_total'].sum():,.0f}")
            col1, col2 = st.columns(2)
            if col1.button("💾 Guardar todos los gastos", type="primary"):
                for item in st.session_state.carrito:
                    nuevo_gasto = Gasto(
                        proyecto_id=item["proyecto_id"],
                        concepto=item["concepto"],
                        categoria=item["categoria"],
                        unidad=item["unidad"],
                        cantidad=item["cantidad"],
                        valor_unitario=item["valor_unitario"],
                        valor_total=item["valor_total"],
                        proveedor_id=item["proveedor_id"],
                        estado_pago="Pendiente",
                        conciliado=False,
                        created_by=usuario.id
                    )
                    db.add(nuevo_gasto)
                    db.commit()
                    registrar_auditoria("Gasto", nuevo_gasto.id, "insert", usuario.id)
                st.session_state.carrito = []
                st.success("Gastos guardados exitosamente")
                st.rerun()
            if col2.button("Vaciar carrito"):
                st.session_state.carrito = []
                st.rerun()

def pagina_pagos():
    st.markdown("## 💳 Gestión de Pagos")
    st.markdown("---")
    usuario = get_usuario_actual()
    proyectos = db.query(Proyecto).all()
    if not proyectos:
        st.warning("No hay proyectos")
        return

    with st.expander("🔍 Filtros de búsqueda", expanded=True):
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            proyecto_sel = st.selectbox("Proyecto", proyectos, format_func=lambda x: x.nombre)
        with col_f2:
            busqueda_pago = st.text_input("Buscar por concepto, proveedor o categoría", placeholder="Ej: material, Johán...")
        with col_f3:
            filtro_estado_pago = st.selectbox("Estado de Pago", ["Todos", "Pendiente", "Parcial", "Pagado"])
    
    if proyecto_sel:
        query = db.query(Gasto).filter(Gasto.proyecto_id == proyecto_sel.id)
        if busqueda_pago:
            query = query.join(Proveedor, Gasto.proveedor_id == Proveedor.id, isouter=True)
            query = query.filter(
                (Gasto.concepto.contains(busqueda_pago)) |
                (Gasto.categoria.contains(busqueda_pago)) |
                (Proveedor.nombre.contains(busqueda_pago))
            )
        if filtro_estado_pago != "Todos":
            query = query.filter(Gasto.estado_pago == filtro_estado_pago)
        
        gastos = query.all()
        
        if not gastos:
            st.info("No hay gastos que coincidan con los filtros para este proyecto.")
        else:
            for gasto in gastos:
                pagos = db.query(Pago).filter(Pago.gasto_id == gasto.id).all()
                total_pagado = sum(p.monto for p in pagos)
                saldo = gasto.valor_total - total_pagado
                with st.container(border=True):
                    col1, col2, col3 = st.columns([2, 1, 1])
                    col1.markdown(f"**{gasto.concepto}**  \nProveedor: {db.query(Proveedor).filter(Proveedor.id == gasto.proveedor_id).first().nombre if gasto.proveedor_id else 'N/A'}")
                    col2.metric("Total Gasto", f"${gasto.valor_total:,.0f}")
                    col3.metric("Saldo Pendiente", f"${saldo:,.0f}")
                    
                    # Mostrar pagos existentes
                    if pagos:
                        st.markdown("**Pagos registrados:**")
                        for pago in pagos:
                            col_p1, col_p2, col_p3, col_p4, col_p5, col_p6 = st.columns([1.2, 1.2, 1.2, 1.2, 1.5, 0.5])
                            col_p1.write(pago.fecha.strftime("%Y-%m-%d"))
                            col_p2.write(pago.tipo)
                            col_p3.write(pago.concepto or "")  # Nuevo campo
                            col_p4.write(f"${pago.monto:,.0f}")
                            col_p5.write(pago.numero_factura or "")
                            with col_p6:
                                if st.session_state.rol_actual in ["Gerencia", "Auxiliar Contable"]:
                                    if st.button("🗑️", key=f"del_pago_{pago.id}"):
                                        db.delete(pago)
                                        db.commit()
                                        registrar_auditoria("Pago", pago.id, "delete", usuario.id)
                                        st.success("Pago eliminado")
                                        st.rerun()
                    
                    if saldo > 0:
                        with st.expander("Registrar Pago"):
                            with st.form(f"pago_{gasto.id}"):
                                tipo = st.selectbox("Tipo de documento", ["Factura", "Cotización", "Cuenta de Cobro"])
                                concepto = st.selectbox("Concepto del pago", ["Anticipo", "Avance", "Finiquito"])
                                n_factura = st.text_input("N° Documento")
                                fecha = st.date_input("Fecha", datetime.date.today())
                                monto = st.number_input("Monto a pagar", min_value=0.0, max_value=saldo, step=1000.0)
                                observaciones = st.text_area("Observaciones")
                                submitted = st.form_submit_button("Registrar Pago")
                                if submitted and monto > 0:
                                    nuevo_pago = Pago(
                                        gasto_id=gasto.id,
                                        tipo=tipo,
                                        concepto=concepto,          # Nuevo campo
                                        numero_factura=n_factura,
                                        fecha=fecha,
                                        monto=monto,
                                        observaciones=observaciones,
                                        created_by=usuario.id
                                    )
                                    db.add(nuevo_pago)
                                    nuevo_total_pagado = total_pagado + monto
                                    if nuevo_total_pagado >= gasto.valor_total:
                                        gasto.estado_pago = "Pagado"
                                    elif nuevo_total_pagado > 0:
                                        gasto.estado_pago = "Parcial"
                                    db.commit()
                                    registrar_auditoria("Pago", nuevo_pago.id, "insert", usuario.id)
                                    registrar_auditoria("Gasto", gasto.id, "update", usuario.id,
                                                        datos_anteriores={"estado_pago": gasto.estado_pago},
                                                        datos_nuevos={"estado_pago": gasto.estado_pago})
                                    st.success("Pago registrado")
                                    st.rerun()
                    else:
                        st.success("✅ Gasto totalmente pagado")

def pagina_proveedores():
    st.markdown("## 🏢 Gestión de Proveedores")
    st.markdown("---")
    usuario = get_usuario_actual()
    proveedores = db.query(Proveedor).all()
    
    if proveedores:
        st.markdown("### Lista de Proveedores")
        for idx, proveedor in enumerate(proveedores):
            col1, col2, col3, col4, col5 = st.columns([2, 1.5, 1.5, 1.5, 0.5])
            col1.write(proveedor.nombre)
            col2.write(proveedor.nit or "")
            col3.write(proveedor.contacto or "")
            col4.write(proveedor.telefono or "")
            with col5:
                if st.session_state.rol_actual in ["Gerencia", "Auxiliar Contable"]:
                    if st.button("🗑️", key=f"del_prov_{proveedor.id}"):
                        gastos = db.query(Gasto).filter(Gasto.proveedor_id == proveedor.id).count()
                        if gastos > 0:
                            st.error(f"No se puede eliminar porque tiene {gastos} gastos asociados.")
                        else:
                            db.delete(proveedor)
                            db.commit()
                            registrar_auditoria("Proveedor", proveedor.id, "delete", usuario.id)
                            st.success("Proveedor eliminado")
                            st.rerun()
    else:
        st.info("No hay proveedores registrados")

    with st.expander("➕ Nuevo Proveedor"):
        with st.form("nuevo_prov"):
            nombre = st.text_input("Nombre")
            nit = st.text_input("NIT")
            contacto = st.text_input("Contacto")
            telefono = st.text_input("Teléfono")
            submitted = st.form_submit_button("Guardar")
            if submitted and nombre:
                nuevo = Proveedor(nombre=nombre, nit=nit, contacto=contacto, telefono=telefono, created_by=usuario.id)
                db.add(nuevo)
                db.commit()
                registrar_auditoria("Proveedor", nuevo.id, "insert", usuario.id)
                st.success("Proveedor creado")
                st.rerun()

def pagina_conciliacion():
    if st.session_state.rol_actual not in ["Gerencia", "Auxiliar Contable"]:
        st.error("No tienes permiso para acceder a esta sección.")
        return

    st.markdown("## ✅ Centro de Conciliación")
    st.markdown("Gestiona el estado de pagos y conciliaciones.")
    st.markdown("---")

    tipo_conciliacion = st.radio(
        "Tipo de Gestión",
        ["💳 Gastos (Pagos a Proveedores)", "💰 Ingresos (Cobros a Clientes)"],
        horizontal=True
    )

    col_f1, col_f2, col_f3, col_f4 = st.columns(4)
    with col_f1:
        proyectos = db.query(Proyecto).all()
        opciones_proy = ["Todos"] + [p.nombre for p in proyectos]
        proy_sel = st.selectbox("Proyecto", opciones_proy, key="conc_proyecto")
    with col_f2:
        busqueda = st.text_input("🔍 Buscar", placeholder="Concepto, proyecto, cliente...", key="conc_busqueda")
    with col_f3:
        fecha_desde = st.date_input("Fecha desde", datetime.date.today() - datetime.timedelta(days=30), key="conc_desde")
        fecha_hasta = st.date_input("Fecha hasta", datetime.date.today(), key="conc_hasta")
    with col_f4:
        busqueda_acta = st.text_input("📄 Buscar por Acta", placeholder="Número de acta...", key="conc_acta")

    # ============================================
    # SECCIÓN GASTOS (PAGOS A PROVEEDORES)
    # ============================================
    if tipo_conciliacion == "💳 Gastos (Pagos a Proveedores)":
        st.markdown("### 💳 Gestión de Pagos a Proveedores")
        st.markdown("Selecciona los gastos y marca como Pagado/Pendiente. Asigna acta y evidencia.")

        query = db.query(Gasto)
        if proy_sel != "Todos":
            proyecto_obj = db.query(Proyecto).filter(Proyecto.nombre == proy_sel).first()
            if proyecto_obj:
                query = query.filter(Gasto.proyecto_id == proyecto_obj.id)

        if busqueda:
            query = query.join(Proyecto, Gasto.proyecto_id == Proyecto.id).join(Proveedor, Gasto.proveedor_id == Proveedor.id, isouter=True)
            query = query.filter(
                (Gasto.concepto.contains(busqueda)) |
                (Proyecto.nombre.contains(busqueda)) |
                (Proyecto.cliente.contains(busqueda)) |
                (Proveedor.nombre.contains(busqueda))
            )

        if busqueda_acta:
            query = query.filter(Gasto.acta_conciliacion.contains(busqueda_acta))

        if fecha_desde and fecha_hasta:
            query = query.filter(Gasto.created_at >= fecha_desde)
            query = query.filter(Gasto.created_at <= fecha_hasta + datetime.timedelta(days=1))

        filtro_estado = st.selectbox("Filtrar por Estado de Pago", ["Todos", "Pendiente", "Parcial", "Pagado"], key="filtro_estado_gasto")
        if filtro_estado != "Todos":
            query = query.filter(Gasto.estado_pago == filtro_estado)

        gastos = query.order_by(desc(Gasto.id)).all()

        if not gastos:
            st.info("No hay gastos que coincidan con los filtros.")
        else:
            data = []
            for g in gastos:
                proveedor = db.query(Proveedor).filter(Proveedor.id == g.proveedor_id).first()
                proyecto = db.query(Proyecto).filter(Proyecto.id == g.proyecto_id).first()
                data.append({
                    "ID": g.id,
                    "Proyecto": proyecto.nombre if proyecto else "N/A",
                    "Cliente": proyecto.cliente if proyecto else "N/A",
                    "Concepto": g.concepto,
                    "Categoría": g.categoria,
                    "Total": g.valor_total,
                    "Estado Pago": g.estado_pago,
                    "Acta": g.acta_conciliacion or "",
                    "Evidencia": "📎" if g.archivo_evidencia else "",
                    "Proveedor": proveedor.nombre if proveedor else "N/A",
                    "Fecha": g.created_at.strftime("%Y-%m-%d") if g.created_at else "N/A"
                })
            df = pd.DataFrame(data)

            eventos = st.dataframe(
                df,
                use_container_width=True,
                on_select="rerun",
                selection_mode="multi-row",
                key="tabla_gastos_pagos"
            )

            if hasattr(eventos, 'selection') and hasattr(eventos.selection, 'rows'):
                filas_seleccionadas = eventos.selection.rows
            else:
                filas_seleccionadas = eventos.get("selection", {}).get("rows", [])

            ids_seleccionados = [df.iloc[i]["ID"] for i in filas_seleccionadas] if filas_seleccionadas else []
            st.session_state.gastos_ids = ids_seleccionados

            st.markdown(f"**{len(ids_seleccionados)}** gastos seleccionados.")

            col_acc1, col_acc2, col_acc3 = st.columns(3)
            with col_acc1:
                if st.button("✅ Marcar como Pagados", type="primary", use_container_width=True):
                    if ids_seleccionados:
                        usuario = get_usuario_actual()
                        for g_id in ids_seleccionados:
                            gasto = db.query(Gasto).filter(Gasto.id == g_id).first()
                            if gasto and gasto.estado_pago != "Pagado":
                                gasto.estado_pago = "Pagado"
                                db.commit()
                                registrar_auditoria(
                                    "Gasto", gasto.id, "update", usuario.id,
                                    datos_anteriores={"estado_pago": gasto.estado_pago},
                                    datos_nuevos={"estado_pago": "Pagado"}
                                )
                        st.success(f"✅ {len(ids_seleccionados)} gastos marcados como Pagados.")
                        st.session_state.gastos_ids = []
                        st.rerun()
                    else:
                        st.warning("⚠️ Selecciona al menos un gasto.")

                if st.button("⏳ Marcar como Pendientes", use_container_width=True):
                    if ids_seleccionados:
                        usuario = get_usuario_actual()
                        for g_id in ids_seleccionados:
                            gasto = db.query(Gasto).filter(Gasto.id == g_id).first()
                            if gasto and gasto.estado_pago != "Pendiente":
                                gasto.estado_pago = "Pendiente"
                                db.commit()
                                registrar_auditoria(
                                    "Gasto", gasto.id, "update", usuario.id,
                                    datos_anteriores={"estado_pago": gasto.estado_pago},
                                    datos_nuevos={"estado_pago": "Pendiente"}
                                )
                        st.success(f"✅ {len(ids_seleccionados)} gastos marcados como Pendientes.")
                        st.session_state.gastos_ids = []
                        st.rerun()
                    else:
                        st.warning("⚠️ Selecciona al menos un gasto.")

            with col_acc2:
                acta_masiva = st.text_input("📄 Asignar Acta", placeholder="Número de acta...", key="acta_masiva_gasto")
                if st.button("📄 Asignar Acta a Seleccionados", use_container_width=True):
                    if ids_seleccionados and acta_masiva:
                        usuario = get_usuario_actual()
                        for g_id in ids_seleccionados:
                            gasto = db.query(Gasto).filter(Gasto.id == g_id).first()
                            if gasto:
                                gasto.acta_conciliacion = acta_masiva
                                db.commit()
                                registrar_auditoria(
                                    "Gasto", gasto.id, "update", usuario.id,
                                    datos_anteriores={"acta_conciliacion": gasto.acta_conciliacion},
                                    datos_nuevos={"acta_conciliacion": acta_masiva}
                                )
                        st.success(f"✅ Acta asignada a {len(ids_seleccionados)} gastos.")
                        st.session_state.gastos_ids = []
                        st.rerun()
                    else:
                        st.warning("⚠️ Selecciona gastos y escribe un número de acta.")

            with col_acc3:
                if st.button("🧹 Limpiar Selección", use_container_width=True):
                    st.session_state.gastos_ids = []
                    st.rerun()

            # --- EDICIÓN INDIVIDUAL DE UN GASTO ---
            st.markdown("---")
            st.markdown("#### ✏️ Edición individual de un gasto")
            if gastos:
                opciones_gastos = {f"{g.id} - {g.concepto[:50]}": g.id for g in gastos}
                gasto_seleccionado_editar = st.selectbox(
                    "Selecciona un gasto para editar",
                    options=list(opciones_gastos.keys()),
                    key="selector_gasto_editar"
                )
                gasto_id_editar = opciones_gastos[gasto_seleccionado_editar]
                gasto = db.query(Gasto).filter(Gasto.id == gasto_id_editar).first()

                if gasto:
                    if gasto.archivo_evidencia:
                        st.info(f"📎 Archivo actual: {gasto.archivo_evidencia}")
                        if st.button("🗑️ Eliminar archivo", key=f"del_gasto_{gasto.id}"):
                            ruta_archivo = os.path.join("uploads", gasto.archivo_evidencia)
                            if os.path.exists(ruta_archivo):
                                os.remove(ruta_archivo)
                            gasto.archivo_evidencia = None
                            db.commit()
                            st.success("Archivo eliminado")
                            st.rerun()

                    with st.form(f"edit_gasto_{gasto.id}"):
                        st.markdown(f"**Gasto:** {gasto.concepto}")
                        st.markdown(f"**Proyecto:** {gasto.proyecto.nombre if gasto.proyecto else 'N/A'}")
                        
                        nuevo_estado = st.selectbox(
                            "Estado de Pago",
                            ["Pendiente", "Parcial", "Pagado"],
                            index=["Pendiente", "Parcial", "Pagado"].index(gasto.estado_pago)
                        )
                        nueva_acta = st.text_input("Acta de Conciliación", value=gasto.acta_conciliacion or "")
                        
                        archivo_subido = st.file_uploader(
                            "📎 Subir evidencia (PDF, imagen, etc.)",
                            type=["pdf", "png", "jpg", "jpeg", "doc", "docx"],
                            key=f"evidencia_gasto_{gasto.id}"
                        )
                        
                        submitted = st.form_submit_button("💾 Guardar Cambios")
                        if submitted:
                            cambios = {}
                            if gasto.estado_pago != nuevo_estado:
                                cambios["estado_pago"] = (gasto.estado_pago, nuevo_estado)
                                gasto.estado_pago = nuevo_estado
                            if gasto.acta_conciliacion != nueva_acta:
                                cambios["acta_conciliacion"] = (gasto.acta_conciliacion, nueva_acta)
                                gasto.acta_conciliacion = nueva_acta
                            
                            if archivo_subido:
                                if not os.path.exists("uploads"):
                                    os.makedirs("uploads")
                                nombre_archivo = f"gasto_{gasto.id}_{archivo_subido.name}"
                                ruta_completa = os.path.join("uploads", nombre_archivo)
                                with open(ruta_completa, "wb") as f:
                                    f.write(archivo_subido.getbuffer())
                                gasto.archivo_evidencia = nombre_archivo
                                cambios["archivo_evidencia"] = (None, nombre_archivo)
                            
                            if cambios:
                                db.commit()
                                usuario = get_usuario_actual()
                                registrar_auditoria(
                                    "Gasto", gasto.id, "update", usuario.id,
                                    datos_anteriores={k: v[0] for k, v in cambios.items()},
                                    datos_nuevos={k: v[1] for k, v in cambios.items()}
                                )
                                st.success("✅ Cambios guardados correctamente")
                                st.rerun()

    # ============================================
    # SECCIÓN INGRESOS (COBROS A CLIENTES)
    # ============================================
    else:
        st.markdown("### 💰 Gestión de Ingresos (Cobros a Clientes)")
        st.markdown("Gestiona los cobros de tus clientes. Marca como cobrado, conciliado, asigna acta y sube evidencia.")
        
        tab1, tab2 = st.tabs(["📋 Gestión de Cobros", "📊 Reporte Gerencial"])
        
        with tab1:
            col_f5, col_f6 = st.columns(2)
            with col_f5:
                filtro_cobrado = st.selectbox("Cobrado", ["Todos", "Cobrado", "Pendiente"], key="filtro_cobrado")
            with col_f6:
                filtro_conciliado = st.selectbox("Conciliado", ["Todos", "Conciliado", "No Conciliado"], key="filtro_conciliado_ingreso")
            
            query = db.query(PartidaPresupuesto)
            if proy_sel != "Todos":
                proyecto_obj = db.query(Proyecto).filter(Proyecto.nombre == proy_sel).first()
                if proyecto_obj:
                    query = query.filter(PartidaPresupuesto.proyecto_id == proyecto_obj.id)

            if busqueda:
                query = query.join(Proyecto, PartidaPresupuesto.proyecto_id == Proyecto.id)
                query = query.filter(
                    (PartidaPresupuesto.descripcion.contains(busqueda)) |
                    (Proyecto.nombre.contains(busqueda)) |
                    (Proyecto.cliente.contains(busqueda))
                )

            if busqueda_acta:
                query = query.filter(PartidaPresupuesto.acta_conciliacion_ingreso.contains(busqueda_acta))

            if fecha_desde and fecha_hasta:
                query = query.filter(PartidaPresupuesto.created_at >= fecha_desde)
                query = query.filter(PartidaPresupuesto.created_at <= fecha_hasta + datetime.timedelta(days=1))

            if filtro_cobrado == "Cobrado":
                query = query.filter(PartidaPresupuesto.cobrado == True)
            elif filtro_cobrado == "Pendiente":
                query = query.filter(PartidaPresupuesto.cobrado == False)
            if filtro_conciliado == "Conciliado":
                query = query.filter(PartidaPresupuesto.conciliado_ingreso == True)
            elif filtro_conciliado == "No Conciliado":
                query = query.filter(PartidaPresupuesto.conciliado_ingreso == False)

            partidas = query.order_by(desc(PartidaPresupuesto.id)).all()

            if not partidas:
                st.info("No hay partidas que coincidan con los filtros.")
            else:
                data = []
                for p in partidas:
                    proyecto = db.query(Proyecto).filter(Proyecto.id == p.proyecto_id).first()
                    total_cobrado = db.query(func.sum(CobroCliente.monto)).filter(CobroCliente.partida_id == p.id).scalar() or 0
                    saldo_pendiente = p.total - total_cobrado
                    
                    data.append({
                        "ID": p.id,
                        "Proyecto": proyecto.nombre if proyecto else "N/A",
                        "Cliente": proyecto.cliente if proyecto else "N/A",
                        "Concepto": p.descripcion,
                        "Categoría": p.categoria,
                        "Valor": p.total,
                        "Cobrado": total_cobrado,
                        "Saldo": saldo_pendiente,
                        "Estado Cobro": "✅ Cobrado" if p.cobrado else "⏳ Pendiente",
                        "Conciliado": "✅" if p.conciliado_ingreso else "❌",
                        "Acta": p.acta_conciliacion_ingreso or "",
                        "Evidencia": "📎" if p.archivo_evidencia_ingreso else "",
                        "Fecha": p.created_at.strftime("%Y-%m-%d") if p.created_at else "N/A"
                    })
                df = pd.DataFrame(data)

                st.markdown("#### Selecciona las partidas a modificar")
                
                df_edit = df.copy()
                df_edit["Seleccionar"] = False
                
                edited_df = st.data_editor(
                    df_edit,
                    use_container_width=True,
                    column_config={
                        "Seleccionar": st.column_config.CheckboxColumn("Seleccionar", default=False)
                    },
                    disabled=["ID", "Proyecto", "Cliente", "Concepto", "Categoría", "Valor", "Cobrado", "Saldo", "Estado Cobro", "Conciliado", "Acta", "Evidencia", "Fecha"],
                    key="tabla_ingresos_cobros"
                )
                
                ids_seleccionados = []
                for idx, row in edited_df.iterrows():
                    if row["Seleccionar"]:
                        ids_seleccionados.append(row["ID"])
                
                st.session_state.ingresos_ids = ids_seleccionados
                
                st.markdown(f"**{len(ids_seleccionados)}** partidas seleccionadas.")

                col_acc1, col_acc2, col_acc3 = st.columns(3)
                
                with col_acc1:
                    if st.button("💰 Marcar como Cobrado", type="primary", use_container_width=True):
                        if ids_seleccionados:
                            usuario = get_usuario_actual()
                            for p_id in ids_seleccionados:
                                partida = db.query(PartidaPresupuesto).filter(PartidaPresupuesto.id == p_id).first()
                                if partida:
                                    partida.cobrado = True
                                    db.commit()
                                    registrar_auditoria(
                                        "PartidaPresupuesto", partida.id, "update", usuario.id,
                                        datos_anteriores={"cobrado": False},
                                        datos_nuevos={"cobrado": True}
                                    )
                            st.success(f"✅ {len(ids_seleccionados)} partidas marcadas como Cobrado.")
                            st.session_state.ingresos_ids = []
                            st.rerun()
                        else:
                            st.warning("⚠️ Selecciona al menos una partida.")
                    
                    if st.button("⏳ Marcar como Pendiente", use_container_width=True):
                        if ids_seleccionados:
                            usuario = get_usuario_actual()
                            for p_id in ids_seleccionados:
                                partida = db.query(PartidaPresupuesto).filter(PartidaPresupuesto.id == p_id).first()
                                if partida:
                                    partida.cobrado = False
                                    db.commit()
                                    registrar_auditoria(
                                        "PartidaPresupuesto", partida.id, "update", usuario.id,
                                        datos_anteriores={"cobrado": True},
                                        datos_nuevos={"cobrado": False}
                                    )
                            st.success(f"✅ {len(ids_seleccionados)} partidas marcadas como Pendiente.")
                            st.session_state.ingresos_ids = []
                            st.rerun()
                        else:
                            st.warning("⚠️ Selecciona al menos una partida.")

                with col_acc2:
                    if st.button("✅ Marcar como Conciliado", type="primary", use_container_width=True):
                        if ids_seleccionados:
                            usuario = get_usuario_actual()
                            for p_id in ids_seleccionados:
                                partida = db.query(PartidaPresupuesto).filter(PartidaPresupuesto.id == p_id).first()
                                if partida:
                                    partida.conciliado_ingreso = True
                                    db.commit()
                                    registrar_auditoria(
                                        "PartidaPresupuesto", partida.id, "update", usuario.id,
                                        datos_anteriores={"conciliado_ingreso": False},
                                        datos_nuevos={"conciliado_ingreso": True}
                                    )
                            st.success(f"✅ {len(ids_seleccionados)} partidas marcadas como Conciliado.")
                            st.session_state.ingresos_ids = []
                            st.rerun()
                        else:
                            st.warning("⚠️ Selecciona al menos una partida.")
                    
                    if st.button("❌ Desmarcar Conciliado", use_container_width=True):
                        if ids_seleccionados:
                            usuario = get_usuario_actual()
                            for p_id in ids_seleccionados:
                                partida = db.query(PartidaPresupuesto).filter(PartidaPresupuesto.id == p_id).first()
                                if partida:
                                    partida.conciliado_ingreso = False
                                    db.commit()
                                    registrar_auditoria(
                                        "PartidaPresupuesto", partida.id, "update", usuario.id,
                                        datos_anteriores={"conciliado_ingreso": True},
                                        datos_nuevos={"conciliado_ingreso": False}
                                    )
                            st.success(f"❌ {len(ids_seleccionados)} partidas desmarcadas como Conciliado.")
                            st.session_state.ingresos_ids = []
                            st.rerun()
                        else:
                            st.warning("⚠️ Selecciona al menos una partida.")

                with col_acc3:
                    acta_masiva = st.text_input("📄 Asignar Acta", placeholder="Número de acta...", key="acta_masiva_ingreso")
                    if st.button("📄 Asignar Acta a Seleccionados", use_container_width=True):
                        if ids_seleccionados and acta_masiva:
                            usuario = get_usuario_actual()
                            for p_id in ids_seleccionados:
                                partida = db.query(PartidaPresupuesto).filter(PartidaPresupuesto.id == p_id).first()
                                if partida:
                                    partida.acta_conciliacion_ingreso = acta_masiva
                                    db.commit()
                                    registrar_auditoria(
                                        "PartidaPresupuesto", partida.id, "update", usuario.id,
                                        datos_anteriores={"acta_conciliacion_ingreso": partida.acta_conciliacion_ingreso},
                                        datos_nuevos={"acta_conciliacion_ingreso": acta_masiva}
                                    )
                            st.success(f"✅ Acta asignada a {len(ids_seleccionados)} partidas.")
                            st.session_state.ingresos_ids = []
                            st.rerun()
                        else:
                            st.warning("⚠️ Selecciona partidas y escribe un número de acta.")

                    if st.button("🧹 Limpiar Selección", use_container_width=True):
                        st.session_state.ingresos_ids = []
                        st.rerun()

                st.markdown("---")
                st.markdown("#### ✏️ Edición individual de una partida")
                
                if partidas:
                    opciones_partidas = {f"{p.id} - {p.descripcion[:50]}": p.id for p in partidas}
                    partida_seleccionada_editar = st.selectbox(
                        "Selecciona una partida para editar",
                        options=list(opciones_partidas.keys()),
                        key="selector_partida_editar"
                    )
                    partida_id_editar = opciones_partidas[partida_seleccionada_editar]
                    partida = db.query(PartidaPresupuesto).filter(PartidaPresupuesto.id == partida_id_editar).first()
                    
                    if partida:
                        if partida.archivo_evidencia_ingreso:
                            st.info(f"📎 Archivo actual: {partida.archivo_evidencia_ingreso}")
                            if st.button("🗑️ Eliminar archivo", key=f"del_partida_{partida.id}"):
                                ruta_archivo = os.path.join("uploads", partida.archivo_evidencia_ingreso)
                                if os.path.exists(ruta_archivo):
                                    os.remove(ruta_archivo)
                                partida.archivo_evidencia_ingreso = None
                                db.commit()
                                st.success("Archivo eliminado")
                                st.rerun()

                        with st.form(f"edit_partida_{partida.id}"):
                            st.markdown(f"**Partida:** {partida.descripcion}")
                            st.markdown(f"**Proyecto:** {partida.proyecto.nombre if partida.proyecto else 'N/A'}")
                            
                            nuevo_cobrado = st.checkbox("💰 Cobrado", value=partida.cobrado)
                            nuevo_conciliado = st.checkbox("✅ Conciliado", value=partida.conciliado_ingreso)
                            nueva_acta = st.text_input("📄 Acta de Conciliación", value=partida.acta_conciliacion_ingreso or "")
                            
                            archivo_subido = st.file_uploader(
                                "📎 Subir evidencia (PDF, imagen, etc.)",
                                type=["pdf", "png", "jpg", "jpeg", "doc", "docx"],
                                key=f"evidencia_partida_{partida.id}"
                            )
                            
                            submitted = st.form_submit_button("💾 Guardar Cambios")
                            if submitted:
                                cambios = {}
                                
                                if partida.cobrado != nuevo_cobrado:
                                    cambios["cobrado"] = (partida.cobrado, nuevo_cobrado)
                                    partida.cobrado = nuevo_cobrado
                                
                                if partida.conciliado_ingreso != nuevo_conciliado:
                                    cambios["conciliado_ingreso"] = (partida.conciliado_ingreso, nuevo_conciliado)
                                    partida.conciliado_ingreso = nuevo_conciliado
                                
                                if partida.acta_conciliacion_ingreso != nueva_acta:
                                    cambios["acta_conciliacion_ingreso"] = (partida.acta_conciliacion_ingreso, nueva_acta)
                                    partida.acta_conciliacion_ingreso = nueva_acta
                                
                                if archivo_subido:
                                    if not os.path.exists("uploads"):
                                        os.makedirs("uploads")
                                    nombre_archivo = f"partida_{partida.id}_{archivo_subido.name}"
                                    ruta_completa = os.path.join("uploads", nombre_archivo)
                                    with open(ruta_completa, "wb") as f:
                                        f.write(archivo_subido.getbuffer())
                                    partida.archivo_evidencia_ingreso = nombre_archivo
                                    cambios["archivo_evidencia_ingreso"] = (None, nombre_archivo)
                                
                                if cambios:
                                    db.commit()
                                    usuario = get_usuario_actual()
                                    registrar_auditoria(
                                        "PartidaPresupuesto", partida.id, "update", usuario.id,
                                        datos_anteriores={k: v[0] for k, v in cambios.items()},
                                        datos_nuevos={k: v[1] for k, v in cambios.items()}
                                    )
                                    st.success("✅ Cambios guardados correctamente")
                                    st.rerun()

                if ids_seleccionados:
                    st.markdown("---")
                    st.markdown("#### 💰 Registrar Cobro Manual a las partidas seleccionadas")
                    
                    st.info(f"Partidas seleccionadas: {len(ids_seleccionados)}")
                    for p_id in ids_seleccionados:
                        partida = db.query(PartidaPresupuesto).filter(PartidaPresupuesto.id == p_id).first()
                        if partida:
                            st.write(f"• {partida.descripcion[:50]} (${partida.total:,.0f})")
                    
                    with st.form("registrar_cobro_manual"):
                        monto_cobro = st.number_input("Monto a cobrar por partida", min_value=0.0, step=1000.0, value=1000.0)
                        fecha_cobro = st.date_input("Fecha de Cobro", datetime.date.today())
                        n_factura = st.text_input("N° Factura Emitida")
                        observaciones = st.text_area("Observaciones")
                        submitted = st.form_submit_button("💰 Registrar Cobro")
                        if submitted and monto_cobro > 0:
                            usuario = get_usuario_actual()
                            registrados = 0
                            for p_id in ids_seleccionados:
                                partida = db.query(PartidaPresupuesto).filter(PartidaPresupuesto.id == p_id).first()
                                if partida:
                                    total_cobrado_actual = db.query(func.sum(CobroCliente.monto)).filter(CobroCliente.partida_id == p_id).scalar() or 0
                                    saldo = partida.total - total_cobrado_actual
                                    if monto_cobro <= saldo:
                                        nuevo_cobro = CobroCliente(
                                            partida_id=p_id,
                                            proyecto_id=partida.proyecto_id,
                                            monto=monto_cobro,
                                            fecha=fecha_cobro,
                                            numero_factura=n_factura,
                                            observaciones=observaciones,
                                            creado_por=usuario.id
                                        )
                                        db.add(nuevo_cobro)
                                        db.commit()
                                        registrar_auditoria(
                                            "CobroCliente", nuevo_cobro.id, "insert", usuario.id,
                                            datos_nuevos={"partida_id": p_id, "monto": monto_cobro}
                                        )
                                        registrados += 1
                                    else:
                                        st.warning(f"⚠️ El monto excede el saldo de {partida.descripcion[:30]} (Saldo: ${saldo:,.0f})")
                            if registrados > 0:
                                st.success(f"✅ Cobro registrado para {registrados} partidas.")
                                st.session_state.ingresos_ids = []
                                st.rerun()

        # --- TAB 2: REPORTE GERENCIAL DE INGRESOS ---
        with tab2:
            st.markdown("### 📊 Reporte Gerencial de Cobros a Clientes")
            st.markdown("Resumen del estado de pagos de todos los clientes.")
            
            proyectos_query = db.query(Proyecto)
            if proy_sel != "Todos":
                proyecto_obj = db.query(Proyecto).filter(Proyecto.nombre == proy_sel).first()
                if proyecto_obj:
                    proyectos_query = proyectos_query.filter(Proyecto.id == proyecto_obj.id)
            
            proyectos_lista = proyectos_query.all()
            
            if not proyectos_lista:
                st.info("No hay proyectos para generar el reporte.")
            else:
                reporte_data = []
                total_general = {
                    "ingresos": 0,
                    "cobrado": 0,
                    "pendiente": 0
                }
                
                for proyecto in proyectos_lista:
                    partidas_proyecto = db.query(PartidaPresupuesto).filter(PartidaPresupuesto.proyecto_id == proyecto.id).all()
                    
                    if partidas_proyecto:
                        ingresos_totales = sum(p.total for p in partidas_proyecto)
                        total_cobrado = sum(db.query(func.sum(CobroCliente.monto)).filter(CobroCliente.partida_id == p.id).scalar() or 0 for p in partidas_proyecto)
                        pendiente = ingresos_totales - total_cobrado
                        
                        gastos_proyecto = db.query(func.sum(Gasto.valor_total)).filter(Gasto.proyecto_id == proyecto.id).scalar() or 0
                        rentabilidad = ((ingresos_totales - gastos_proyecto) / ingresos_totales * 100) if ingresos_totales > 0 else 0
                        
                        cobradas = sum(1 for p in partidas_proyecto if p.cobrado)
                        conciliadas = sum(1 for p in partidas_proyecto if p.conciliado_ingreso)
                        
                        reporte_data.append({
                            "Proyecto": proyecto.nombre,
                            "Cliente": proyecto.cliente,
                            "Ingresos Totales": ingresos_totales,
                            "Cobrado": total_cobrado,
                            "Pendiente": pendiente,
                            "Cobradas": cobradas,
                            "Conciliadas": conciliadas,
                            "Total Partidas": len(partidas_proyecto),
                            "Rentabilidad %": rentabilidad
                        })
                        
                        total_general["ingresos"] += ingresos_totales
                        total_general["cobrado"] += total_cobrado
                        total_general["pendiente"] += pendiente
                
                if reporte_data:
                    df_reporte = pd.DataFrame(reporte_data)
                    
                    col_r1, col_r2, col_r3, col_r4 = st.columns(4)
                    col_r1.metric("💰 Ingresos Totales", f"${total_general['ingresos']:,.0f}")
                    col_r2.metric("💵 Cobrado", f"${total_general['cobrado']:,.0f}")
                    col_r3.metric("⏳ Pendiente", f"${total_general['pendiente']:,.0f}")
                    col_r4.metric("📊 Eficiencia de Cobro", f"{(total_general['cobrado']/total_general['ingresos']*100):.1f}%" if total_general['ingresos'] > 0 else "0%")
                    
                    st.dataframe(
                        df_reporte,
                        use_container_width=True,
                        column_config={
                            "Proyecto": "Proyecto",
                            "Cliente": "Cliente",
                            "Ingresos Totales": st.column_config.NumberColumn("Ingresos", format="$%,.0f"),
                            "Cobrado": st.column_config.NumberColumn("Cobrado", format="$%,.0f"),
                            "Pendiente": st.column_config.NumberColumn("Pendiente", format="$%,.0f"),
                            "Cobradas": "Cobradas",
                            "Conciliadas": "Conciliadas",
                            "Total Partidas": "Total Partidas",
                            "Rentabilidad %": st.column_config.NumberColumn("Rentabilidad", format="%.1f%%")
                        }
                    )
                    
                    st.markdown("#### 📊 Comparativa por Proyecto")
                    fig = go.Figure()
                    fig.add_trace(go.Bar(
                        x=df_reporte["Proyecto"],
                        y=df_reporte["Ingresos Totales"],
                        name="Ingresos",
                        marker_color="#2E86C1"
                    ))
                    fig.add_trace(go.Bar(
                        x=df_reporte["Proyecto"],
                        y=df_reporte["Cobrado"],
                        name="Cobrado",
                        marker_color="#28B463"
                    ))
                    fig.add_trace(go.Bar(
                        x=df_reporte["Proyecto"],
                        y=df_reporte["Pendiente"],
                        name="Pendiente",
                        marker_color="#E74C3C"
                    ))
                    fig.update_layout(
                        title="Ingresos, Cobrado y Pendiente por Proyecto",
                        barmode='group',
                        yaxis_title="Monto ($)",
                        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                    )
                    st.plotly_chart(fig, use_container_width=True)
                    
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_reporte.to_excel(writer, sheet_name="Reporte Cobros", index=False)
                    output.seek(0)
                    st.download_button(
                        label="📥 Descargar Reporte Excel",
                        data=output,
                        file_name=f"Reporte_Cobros_Clientes_{datetime.date.today()}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.info("No hay datos para generar el reporte.")

def pagina_reportes():
    st.markdown("## 📊 Generación de Reportes")
    st.markdown("---")
    st.markdown("Genera reportes completos con todos los datos del sistema.")

    with st.expander("🔍 Filtros del Reporte", expanded=True):
        col_f1, col_f2, col_f3 = st.columns(3)
        
        with col_f1:
            fecha_desde = st.date_input(
                "Fecha desde",
                datetime.date.today() - datetime.timedelta(days=90),
                key="reporte_fecha_desde"
            )
            fecha_hasta = st.date_input(
                "Fecha hasta",
                datetime.date.today(),
                key="reporte_fecha_hasta"
            )
        
        with col_f2:
            proyectos = db.query(Proyecto).filter(Proyecto.is_active == True).all()
            opciones_proy = ["Todos"] + [p.nombre for p in proyectos]
            proyectos_seleccionados = st.multiselect(
                "Proyectos",
                options=opciones_proy,
                default=["Todos"],
                key="reporte_proyectos"
            )
        
        with col_f3:
            tipo_transaccion = st.multiselect(
                "Tipo de Transacción",
                options=["Todos", "Partidas", "Gastos", "Pagos", "Cobros"],
                default=["Todos"],
                key="reporte_tipo"
            )

    def aplicar_filtros(query, tabla_fecha, fecha_desde, fecha_hasta):
        if fecha_desde:
            query = query.filter(tabla_fecha >= fecha_desde)
        if fecha_hasta:
            query = query.filter(tabla_fecha <= fecha_hasta + datetime.timedelta(days=1))
        return query

    def obtener_proyectos_filtrados():
        if "Todos" in proyectos_seleccionados or not proyectos_seleccionados:
            return [p.id for p in db.query(Proyecto).filter(Proyecto.is_active == True).all()]
        else:
            return [p.id for p in db.query(Proyecto).filter(Proyecto.nombre.in_(proyectos_seleccionados)).all()]

    def calcular_kpi_proyecto_reporte(proyecto_id):
        ingresos = db.query(func.sum(PartidaPresupuesto.total)).filter(
            PartidaPresupuesto.proyecto_id == proyecto_id,
            PartidaPresupuesto.is_active == True
        ).scalar() or 0
        
        costos = db.query(func.sum(Gasto.valor_total)).filter(
            Gasto.proyecto_id == proyecto_id,
            Gasto.is_active == True
        ).scalar() or 0
        
        cobros = db.query(func.sum(CobroCliente.monto)).join(
            PartidaPresupuesto, CobroCliente.partida_id == PartidaPresupuesto.id
        ).filter(
            PartidaPresupuesto.proyecto_id == proyecto_id
        ).scalar() or 0
        
        pagos = db.query(func.sum(Pago.monto)).join(
            Gasto, Pago.gasto_id == Gasto.id
        ).filter(
            Gasto.proyecto_id == proyecto_id,
            Gasto.is_active == True
        ).scalar() or 0
        
        rentabilidad = ((ingresos - costos) / ingresos * 100) if ingresos > 0 else 0
        avance = (costos / ingresos * 100) if ingresos > 0 else 0
        
        if rentabilidad > 20:
            estado = "🟢 Saludable"
        elif rentabilidad > 10:
            estado = "🟡 Aceptable"
        elif rentabilidad > 0:
            estado = "🟠 Riesgo"
        else:
            estado = "🔴 Crítico"
        
        return {
            "ingresos": ingresos,
            "costos": costos,
            "cobros": cobros,
            "pagos": pagos,
            "rentabilidad": rentabilidad,
            "avance": avance,
            "estado": estado
        }

    def generar_kpis_proyectos(proyectos_ids):
        data = []
        for p_id in proyectos_ids:
            proyecto = db.query(Proyecto).filter(Proyecto.id == p_id).first()
            if proyecto:
                kpi = calcular_kpi_proyecto_reporte(p_id)
                data.append({
                    "ID": p_id,
                    "Proyecto": proyecto.nombre,
                    "Ubicación": proyecto.ubicacion,
                    "Cliente": proyecto.cliente,
                    "Estado": proyecto.estado,
                    "Ingresos": kpi["ingresos"],
                    "Costos": kpi["costos"],
                    "Cobrado": kpi["cobros"],
                    "Pagado": kpi["pagos"],
                    "Rentabilidad %": kpi["rentabilidad"],
                    "Avance %": kpi["avance"],
                    "Estado General": kpi["estado"]
                })
        return pd.DataFrame(data)

    def generar_sabana_unica(proyectos_ids):
        incluir_partidas = "Todos" in tipo_transaccion or "Partidas" in tipo_transaccion
        incluir_gastos = "Todos" in tipo_transaccion or "Gastos" in tipo_transaccion
        incluir_pagos = "Todos" in tipo_transaccion or "Pagos" in tipo_transaccion
        incluir_cobros = "Todos" in tipo_transaccion or "Cobros" in tipo_transaccion
        
        registros = []
        
        if incluir_partidas:
            query = db.query(PartidaPresupuesto).filter(
                PartidaPresupuesto.proyecto_id.in_(proyectos_ids),
                PartidaPresupuesto.is_active == True
            )
            query = aplicar_filtros(query, PartidaPresupuesto.created_at, fecha_desde, fecha_hasta)
            partidas = query.all()
            
            for p in partidas:
                proyecto = db.query(Proyecto).filter(Proyecto.id == p.proyecto_id).first()
                kpi = calcular_kpi_proyecto_reporte(p.proyecto_id) if proyecto else None
                dias_antiguedad = (datetime.date.today() - p.created_at.date()).days if p.created_at else 0
                
                registros.append({
                    "Tipo": "Partida",
                    "ID": p.id,
                    "Proyecto": proyecto.nombre if proyecto else "N/A",
                    "Cliente": proyecto.cliente if proyecto else "N/A",
                    "Concepto/Descripción": p.descripcion,
                    "Categoría": p.categoria,
                    "Cantidad": p.cantidad,
                    "Valor Unitario": p.valor_unitario,
                    "Total": p.total,
                    "Cobrado": "✅" if p.cobrado else "❌",
                    "Conciliado": "✅" if p.conciliado_ingreso else "❌",
                    "Acta": p.acta_conciliacion_ingreso or "",
                    "Evidencia": "📎" if p.archivo_evidencia_ingreso else "",
                    "Estado": "Cobrado" if p.cobrado else "Pendiente",
                    "Fecha Creación": p.created_at.strftime("%Y-%m-%d %H:%M") if p.created_at else "",
                    "Usuario Creador": db.query(Usuario).filter(Usuario.id == p.created_by).first().nombre if p.created_by else "N/A",
                    "Última Modificación": p.updated_at.strftime("%Y-%m-%d %H:%M") if p.updated_at else "",
                    "Usuario Modificador": db.query(Usuario).filter(Usuario.id == p.updated_by).first().nombre if p.updated_by else "N/A",
                    "Activo": "✅" if p.is_active else "❌",
                    "Proveedor": "N/A",
                    "Días de Atraso": dias_antiguedad,
                    "Rentabilidad Proyecto %": kpi["rentabilidad"] if kpi else 0,
                    "Avance Proyecto %": kpi["avance"] if kpi else 0,
                    "Estado Proyecto": kpi["estado"] if kpi else "N/A"
                })
        
        if incluir_gastos:
            query = db.query(Gasto).filter(
                Gasto.proyecto_id.in_(proyectos_ids),
                Gasto.is_active == True
            )
            query = aplicar_filtros(query, Gasto.created_at, fecha_desde, fecha_hasta)
            gastos = query.all()
            
            for g in gastos:
                proyecto = db.query(Proyecto).filter(Proyecto.id == g.proyecto_id).first()
                proveedor = db.query(Proveedor).filter(Proveedor.id == g.proveedor_id).first()
                kpi = calcular_kpi_proyecto_reporte(g.proyecto_id) if proyecto else None
                dias_antiguedad = (datetime.date.today() - g.created_at.date()).days if g.created_at else 0
                
                registros.append({
                    "Tipo": "Gasto",
                    "ID": g.id,
                    "Proyecto": proyecto.nombre if proyecto else "N/A",
                    "Cliente": proyecto.cliente if proyecto else "N/A",
                    "Concepto/Descripción": g.concepto,
                    "Categoría": g.categoria,
                    "Cantidad": g.cantidad,
                    "Valor Unitario": g.valor_unitario,
                    "Total": g.valor_total,
                    "Cobrado": "N/A",
                    "Conciliado": "✅" if g.conciliado else "❌",
                    "Acta": g.acta_conciliacion or "",
                    "Evidencia": "📎" if g.archivo_evidencia else "",
                    "Estado": g.estado_pago,
                    "Fecha Creación": g.created_at.strftime("%Y-%m-%d %H:%M") if g.created_at else "",
                    "Usuario Creador": db.query(Usuario).filter(Usuario.id == g.created_by).first().nombre if g.created_by else "N/A",
                    "Última Modificación": g.updated_at.strftime("%Y-%m-%d %H:%M") if g.updated_at else "",
                    "Usuario Modificador": db.query(Usuario).filter(Usuario.id == g.updated_by).first().nombre if g.updated_by else "N/A",
                    "Activo": "✅" if g.is_active else "❌",
                    "Proveedor": proveedor.nombre if proveedor else "N/A",
                    "Días de Atraso": dias_antiguedad,
                    "Rentabilidad Proyecto %": kpi["rentabilidad"] if kpi else 0,
                    "Avance Proyecto %": kpi["avance"] if kpi else 0,
                    "Estado Proyecto": kpi["estado"] if kpi else "N/A"
                })
        
        if incluir_pagos:
            query = db.query(Pago).join(Gasto).filter(
                Gasto.proyecto_id.in_(proyectos_ids),
                Gasto.is_active == True,
                Pago.is_active == True
            )
            query = aplicar_filtros(query, Pago.fecha, fecha_desde, fecha_hasta)
            pagos = query.all()
            
            for pag in pagos:
                gasto = db.query(Gasto).filter(Gasto.id == pag.gasto_id).first()
                proyecto = db.query(Proyecto).filter(Proyecto.id == gasto.proyecto_id).first() if gasto else None
                kpi = calcular_kpi_proyecto_reporte(gasto.proyecto_id) if proyecto else None
                dias_antiguedad = (datetime.date.today() - pag.fecha).days if pag.fecha else 0
                
                registros.append({
                    "Tipo": "Pago",
                    "ID": pag.id,
                    "Proyecto": proyecto.nombre if proyecto else "N/A",
                    "Cliente": proyecto.cliente if proyecto else "N/A",
                    "Concepto/Descripción": f"Pago a {gasto.concepto if gasto else 'N/A'}",
                    "Categoría": gasto.categoria if gasto else "N/A",
                    "Cantidad": "N/A",
                    "Valor Unitario": "N/A",
                    "Total": pag.monto,
                    "Cobrado": "N/A",
                    "Conciliado": "N/A",
                    "Acta": "N/A",
                    "Evidencia": "N/A",
                    "Estado": "Pagado",
                    "Fecha Creación": pag.created_at.strftime("%Y-%m-%d %H:%M") if pag.created_at else "",
                    "Usuario Creador": db.query(Usuario).filter(Usuario.id == pag.created_by).first().nombre if pag.created_by else "N/A",
                    "Última Modificación": pag.updated_at.strftime("%Y-%m-%d %H:%M") if pag.updated_at else "",
                    "Usuario Modificador": db.query(Usuario).filter(Usuario.id == pag.updated_by).first().nombre if pag.updated_by else "N/A",
                    "Activo": "✅" if pag.is_active else "❌",
                    "Proveedor": db.query(Proveedor).filter(Proveedor.id == gasto.proveedor_id).first().nombre if gasto and gasto.proveedor_id else "N/A",
                    "Días de Atraso": dias_antiguedad,
                    "Rentabilidad Proyecto %": kpi["rentabilidad"] if kpi else 0,
                    "Avance Proyecto %": kpi["avance"] if kpi else 0,
                    "Estado Proyecto": kpi["estado"] if kpi else "N/A"
                })
        
        if incluir_cobros:
            query = db.query(CobroCliente).filter(
                CobroCliente.proyecto_id.in_(proyectos_ids),
                CobroCliente.is_active == True
            )
            query = aplicar_filtros(query, CobroCliente.fecha, fecha_desde, fecha_hasta)
            cobros = query.all()
            
            for cob in cobros:
                partida = db.query(PartidaPresupuesto).filter(PartidaPresupuesto.id == cob.partida_id).first()
                proyecto = db.query(Proyecto).filter(Proyecto.id == cob.proyecto_id).first()
                kpi = calcular_kpi_proyecto_reporte(cob.proyecto_id) if proyecto else None
                dias_antiguedad = (datetime.date.today() - cob.fecha).days if cob.fecha else 0
                
                registros.append({
                    "Tipo": "Cobro",
                    "ID": cob.id,
                    "Proyecto": proyecto.nombre if proyecto else "N/A",
                    "Cliente": proyecto.cliente if proyecto else "N/A",
                    "Concepto/Descripción": f"Cobro de {partida.descripcion if partida else 'N/A'}",
                    "Categoría": partida.categoria if partida else "N/A",
                    "Cantidad": "N/A",
                    "Valor Unitario": "N/A",
                    "Total": cob.monto,
                    "Cobrado": "✅",
                    "Conciliado": "N/A",
                    "Acta": "N/A",
                    "Evidencia": "N/A",
                    "Estado": "Cobrado",
                    "Fecha Creación": cob.created_at.strftime("%Y-%m-%d %H:%M") if cob.created_at else "",
                    "Usuario Creador": db.query(Usuario).filter(Usuario.id == cob.creado_por).first().nombre if cob.creado_por else "N/A",
                    "Última Modificación": cob.updated_at.strftime("%Y-%m-%d %H:%M") if cob.updated_at else "",
                    "Usuario Modificador": db.query(Usuario).filter(Usuario.id == cob.updated_by).first().nombre if cob.updated_by else "N/A",
                    "Activo": "✅" if cob.is_active else "❌",
                    "Proveedor": "N/A",
                    "Días de Atraso": dias_antiguedad,
                    "Rentabilidad Proyecto %": kpi["rentabilidad"] if kpi else 0,
                    "Avance Proyecto %": kpi["avance"] if kpi else 0,
                    "Estado Proyecto": kpi["estado"] if kpi else "N/A"
                })
        
        return pd.DataFrame(registros)

    def generar_excel_multiple(proyectos_ids):
        df_proyectos = generar_kpis_proyectos(proyectos_ids)
        
        query_partidas = db.query(PartidaPresupuesto).filter(
            PartidaPresupuesto.proyecto_id.in_(proyectos_ids),
            PartidaPresupuesto.is_active == True
        )
        query_partidas = aplicar_filtros(query_partidas, PartidaPresupuesto.created_at, fecha_desde, fecha_hasta)
        partidas = query_partidas.all()
        data_partidas = []
        for p in partidas:
            proyecto = db.query(Proyecto).filter(Proyecto.id == p.proyecto_id).first()
            data_partidas.append({
                "ID": p.id,
                "Proyecto": proyecto.nombre if proyecto else "N/A",
                "Cliente": proyecto.cliente if proyecto else "N/A",
                "Categoría": p.categoria,
                "Descripción": p.descripcion,
                "Cantidad": p.cantidad,
                "Valor Unitario": p.valor_unitario,
                "Total": p.total,
                "Cobrado": "Sí" if p.cobrado else "No",
                "Conciliado": "Sí" if p.conciliado_ingreso else "No",
                "Acta": p.acta_conciliacion_ingreso or "",
                "Evidencia": p.archivo_evidencia_ingreso or "",
                "Fecha Creación": p.created_at.strftime("%Y-%m-%d %H:%M") if p.created_at else "",
                "Creado por": db.query(Usuario).filter(Usuario.id == p.created_by).first().nombre if p.created_by else "N/A",
                "Última Modificación": p.updated_at.strftime("%Y-%m-%d %H:%M") if p.updated_at else "",
                "Modificado por": db.query(Usuario).filter(Usuario.id == p.updated_by).first().nombre if p.updated_by else "N/A",
                "Activo": "Sí" if p.is_active else "No"
            })
        df_partidas = pd.DataFrame(data_partidas)
        
        query_gastos = db.query(Gasto).filter(
            Gasto.proyecto_id.in_(proyectos_ids),
            Gasto.is_active == True
        )
        query_gastos = aplicar_filtros(query_gastos, Gasto.created_at, fecha_desde, fecha_hasta)
        gastos = query_gastos.all()
        data_gastos = []
        for g in gastos:
            proyecto = db.query(Proyecto).filter(Proyecto.id == g.proyecto_id).first()
            proveedor = db.query(Proveedor).filter(Proveedor.id == g.proveedor_id).first()
            data_gastos.append({
                "ID": g.id,
                "Proyecto": proyecto.nombre if proyecto else "N/A",
                "Cliente": proyecto.cliente if proyecto else "N/A",
                "Concepto": g.concepto,
                "Categoría": g.categoria,
                "Unidad": g.unidad,
                "Cantidad": g.cantidad,
                "Valor Unitario": g.valor_unitario,
                "Total": g.valor_total,
                "Estado Pago": g.estado_pago,
                "Conciliado": "Sí" if g.conciliado else "No",
                "Acta": g.acta_conciliacion or "",
                "Evidencia": g.archivo_evidencia or "",
                "Proveedor": proveedor.nombre if proveedor else "N/A",
                "Fecha Creación": g.created_at.strftime("%Y-%m-%d %H:%M") if g.created_at else "",
                "Creado por": db.query(Usuario).filter(Usuario.id == g.created_by).first().nombre if g.created_by else "N/A",
                "Última Modificación": g.updated_at.strftime("%Y-%m-%d %H:%M") if g.updated_at else "",
                "Modificado por": db.query(Usuario).filter(Usuario.id == g.updated_by).first().nombre if g.updated_by else "N/A",
                "Activo": "Sí" if g.is_active else "No"
            })
        df_gastos = pd.DataFrame(data_gastos)
        
        query_pagos = db.query(Pago).join(Gasto).filter(
            Gasto.proyecto_id.in_(proyectos_ids),
            Gasto.is_active == True,
            Pago.is_active == True
        )
        query_pagos = aplicar_filtros(query_pagos, Pago.fecha, fecha_desde, fecha_hasta)
        pagos = query_pagos.all()
        data_pagos = []
        for pag in pagos:
            gasto = db.query(Gasto).filter(Gasto.id == pag.gasto_id).first()
            proyecto = db.query(Proyecto).filter(Proyecto.id == gasto.proyecto_id).first() if gasto else None
            proveedor = db.query(Proveedor).filter(Proveedor.id == gasto.proveedor_id).first() if gasto else None
            data_pagos.append({
                "ID": pag.id,
                "Gasto Asociado": gasto.concepto if gasto else "N/A",
                "Proyecto": proyecto.nombre if proyecto else "N/A",
                "Cliente": proyecto.cliente if proyecto else "N/A",
                "Tipo": pag.tipo,
                "N° Factura": pag.numero_factura or "",
                "Fecha": pag.fecha.strftime("%Y-%m-%d") if pag.fecha else "",
                "Monto": pag.monto,
                "Observaciones": pag.observaciones or "",
                "Proveedor": proveedor.nombre if proveedor else "N/A",
                "Fecha Creación": pag.created_at.strftime("%Y-%m-%d %H:%M") if pag.created_at else "",
                "Creado por": db.query(Usuario).filter(Usuario.id == pag.created_by).first().nombre if pag.created_by else "N/A",
                "Última Modificación": pag.updated_at.strftime("%Y-%m-%d %H:%M") if pag.updated_at else "",
                "Modificado por": db.query(Usuario).filter(Usuario.id == pag.updated_by).first().nombre if pag.updated_by else "N/A",
                "Activo": "Sí" if pag.is_active else "No"
            })
        df_pagos = pd.DataFrame(data_pagos)
        
        query_cobros = db.query(CobroCliente).filter(
            CobroCliente.proyecto_id.in_(proyectos_ids),
            CobroCliente.is_active == True
        )
        query_cobros = aplicar_filtros(query_cobros, CobroCliente.fecha, fecha_desde, fecha_hasta)
        cobros = query_cobros.all()
        data_cobros = []
        for cob in cobros:
            partida = db.query(PartidaPresupuesto).filter(PartidaPresupuesto.id == cob.partida_id).first()
            proyecto = db.query(Proyecto).filter(Proyecto.id == cob.proyecto_id).first()
            data_cobros.append({
                "ID": cob.id,
                "Partida Asociada": partida.descripcion if partida else "N/A",
                "Proyecto": proyecto.nombre if proyecto else "N/A",
                "Cliente": proyecto.cliente if proyecto else "N/A",
                "Monto": cob.monto,
                "Fecha": cob.fecha.strftime("%Y-%m-%d") if cob.fecha else "",
                "N° Factura": cob.numero_factura or "",
                "Observaciones": cob.observaciones or "",
                "Fecha Creación": cob.created_at.strftime("%Y-%m-%d %H:%M") if cob.created_at else "",
                "Creado por": db.query(Usuario).filter(Usuario.id == cob.creado_por).first().nombre if cob.creado_por else "N/A",
                "Última Modificación": cob.updated_at.strftime("%Y-%m-%d %H:%M") if cob.updated_at else "",
                "Modificado por": db.query(Usuario).filter(Usuario.id == cob.updated_by).first().nombre if cob.updated_by else "N/A",
                "Activo": "Sí" if cob.is_active else "No"
            })
        df_cobros = pd.DataFrame(data_cobros)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            if not df_proyectos.empty:
                df_proyectos.to_excel(writer, sheet_name="Resumen Proyectos", index=False)
            if not df_partidas.empty:
                df_partidas.to_excel(writer, sheet_name="Partidas", index=False)
            if not df_gastos.empty:
                df_gastos.to_excel(writer, sheet_name="Gastos", index=False)
            if not df_pagos.empty:
                df_pagos.to_excel(writer, sheet_name="Pagos", index=False)
            if not df_cobros.empty:
                df_cobros.to_excel(writer, sheet_name="Cobros", index=False)
            
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width
                for cell in worksheet[1]:
                    cell.font = cell.font.copy(bold=True)
        
        output.seek(0)
        return output

    st.markdown("---")
    
    col_btn1, col_btn2 = st.columns(2)
    
    with col_btn1:
        if st.button("📊 Generar Sábana Única", type="primary", use_container_width=True):
            with st.spinner("Generando sábana única..."):
                proyectos_ids = obtener_proyectos_filtrados()
                df_sabana = generar_sabana_unica(proyectos_ids)
                
                if df_sabana.empty:
                    st.warning("No hay datos para generar el reporte con los filtros seleccionados.")
                else:
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_sabana.to_excel(writer, sheet_name="Sábana Única", index=False)
                        worksheet = writer.sheets["Sábana Única"]
                        for column in worksheet.columns:
                            max_length = 0
                            column_letter = column[0].column_letter
                            for cell in column:
                                try:
                                    if len(str(cell.value)) > max_length:
                                        max_length = len(str(cell.value))
                                except:
                                    pass
                            adjusted_width = min(max_length + 2, 50)
                            worksheet.column_dimensions[column_letter].width = adjusted_width
                        for cell in worksheet[1]:
                            cell.font = cell.font.copy(bold=True)
                    
                    output.seek(0)
                    st.download_button(
                        label="📥 Descargar Sábana Única",
                        data=output,
                        file_name=f"Sabana_Unica_{datetime.date.today()}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_sabana"
                    )
                    st.success("✅ Sábana única generada exitosamente.")
    
    with col_btn2:
        if st.button("📋 Generar Reporte por Módulos", type="primary", use_container_width=True):
            with st.spinner("Generando reporte por módulos..."):
                proyectos_ids = obtener_proyectos_filtrados()
                excel_data = generar_excel_multiple(proyectos_ids)
                
                if excel_data.getvalue():
                    st.download_button(
                        label="📥 Descargar Reporte por Módulos",
                        data=excel_data,
                        file_name=f"Reporte_Modulos_{datetime.date.today()}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_modulos"
                    )
                    st.success("✅ Reporte por módulos generado exitosamente.")
                else:
                    st.warning("No hay datos para generar el reporte con los filtros seleccionados.")

def pagina_usuarios():
    if st.session_state.rol_actual != "Gerencia":
        st.error("No tienes permiso para gestionar usuarios")
        return

    st.markdown("## 👥 Gestión de Usuarios")
    st.markdown("---")
    usuarios = db.query(Usuario).all()
    df_usu = pd.DataFrame([{
        "ID": u.id,
        "Nombre": u.nombre,
        "Rol": u.rol
    } for u in usuarios])
    st.dataframe(df_usu, use_container_width=True)

    with st.expander("➕ Crear Nuevo Usuario"):
        with st.form("nuevo_user"):
            nombre = st.text_input("Nombre de usuario")
            password = st.text_input("Contraseña", type="password")
            rol = st.selectbox("Rol", ["Gerencia", "Operario", "Auxiliar Contable"])
            submitted = st.form_submit_button("Crear")
            if submitted and nombre and password:
                existe = db.query(Usuario).filter(Usuario.nombre == nombre).first()
                if existe:
                    st.error("El usuario ya existe")
                else:
                    nuevo = Usuario(nombre=nombre, password_hash=encriptar_password(password), rol=rol)
                    db.add(nuevo)
                    db.commit()
                    registrar_auditoria("Usuario", nuevo.id, "insert", get_usuario_actual().id)
                    st.success("Usuario creado")
                    st.rerun()

# ============================================
# EJECUCIÓN DE LA PÁGINA SEGÚN MENÚ
# ============================================

if menu == "Dashboard":
    pagina_dashboard()
elif menu == "Proyectos":
    pagina_proyectos()
elif menu == "Partidas Presupuesto":
    pagina_partidas()
elif menu == "Gastos":
    pagina_gastos()
elif menu == "Pagos":
    pagina_pagos()
elif menu == "Proveedores":
    pagina_proveedores()
elif menu == "Conciliación":
    pagina_conciliacion()
elif menu == "Reportes":
    pagina_reportes()
elif menu == "Usuarios":
    pagina_usuarios()

db.close()
