import streamlit as st
import pandas as pd
import os
import json

st.set_page_config(page_title="Gestión de Inventario", layout="wide")

ARCHIVO_DATOS = "inventario_equipos.xlsx"
ARCHIVO_USUARIOS = "usuarios.json"

# Columnas del sistema
COLUMNAS = [
    "MV",
    "Material",
    "Denominación de objeto técnico",
    "Stat.sist.",
    "StatUsu",
    "ESTATUS ACTUAL.",
    "Denomin.",
    "UBICACIÓN ACTUAL",
    "OBSERVACIONES"
]

# --- FUNCIONES DE GESTIÓN DE USUARIOS ---
def cargar_usuarios():
    if os.path.exists(ARCHIVO_USUARIOS):
        try:
            with open(ARCHIVO_USUARIOS, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    # Usuarios por defecto si no existe el archivo JSON
    usuarios_default = {
        "admin": {"clave": "admin123", "rol": "Administrador"},
        "usuario": {"clave": "user123", "rol": "Visualizador"}
    }
    guardar_usuarios(usuarios_default)
    return usuarios_default

def guardar_usuarios(usuarios):
    with open(ARCHIVO_USUARIOS, "w", encoding="utf-8") as f:
        json.dump(usuarios, f, ensure_ascii=False, indent=4)

USUARIOS = cargar_usuarios()

# --- CONTROL DE SESIÓN Y LOGIN ---
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False
    st.session_state["usuario"] = ""
    st.session_state["rol"] = ""

def login():
    st.title("🔒 Acceso al Sistema de Inventario")
    with st.form("form_login"):
        user_input = st.text_input("Usuario")
        pass_input = st.text_input("Contraseña", type="password")
        btn_login = st.form_submit_button("Iniciar Sesión")
        
        if btn_login:
            usuarios_actuales = cargar_usuarios()
            if user_input in usuarios_actuales and usuarios_actuales[user_input]["clave"] == pass_input:
                st.session_state["autenticado"] = True
                st.session_state["usuario"] = user_input
                st.session_state["rol"] = usuarios_actuales[user_input]["rol"]
                st.success(f"Bienvenido {user_input} ({usuarios_actuales[user_input]['rol']})")
                st.rerun()
            else:
                st.error("⚠️ Usuario o contraseña incorrectos.")

def logout():
    st.session_state["autenticado"] = False
    st.session_state["usuario"] = ""
    st.session_state["rol"] = ""
    st.rerun()

# Si no está autenticado, muestra pantalla de login
if not st.session_state["autenticado"]:
    login()
    st.stop()

# --- MOSTRAR BARRA LATERAL DE USUARIO Y ROL ---
st.sidebar.markdown(f"👤 **Usuario:** `{st.session_state['usuario']}`")
st.sidebar.markdown(f"🔑 **Rol:** `{st.session_state['rol']}`")
if st.sidebar.button("🚪 Cerrar Sesión"):
    logout()
st.sidebar.divider()

# --- FUNCIONES DE BASE DE DATOS INVENTARIO ---
def cargar_datos():
    if os.path.exists(ARCHIVO_DATOS):
        try:
            df = pd.read_excel(ARCHIVO_DATOS, dtype=str).fillna("")
            for col in COLUMNAS:
                if col not in df.columns:
                    df[col] = ""
            return df[COLUMNAS].astype(str)
        except Exception:
            return pd.DataFrame(columns=COLUMNAS).astype(str)
    else:
        return pd.DataFrame(columns=COLUMNAS).astype(str)

def guardar_datos(df):
    df = df.astype(str)
    df.to_excel(ARCHIVO_DATOS, index=False)

df = cargar_datos()

st.title("📦 Sistema de Inventario de Equipos")

# Opciones del Menú según el Rol
es_admin = st.session_state["rol"] == "Administrador"

opciones_menu = ["📋 Buscar y Gestionar Inventario"]
if es_admin:
    opciones_menu.extend([
        "➕ Registrar Nuevo Equipo",
        "💾 Importar / Exportar Respaldos",
        "👥 Gestión de Usuarios"
    ])

opcion = st.sidebar.selectbox("Selecciona una opción", opciones_menu)

# 1. BUSCAR Y GESTIONAR (VER, EDITAR Y ELIMINAR DIRECTAMENTE)
if opcion == "📋 Buscar y Gestionar Inventario":
    # Encabezado con métrica del total de equipos
    col_titulo, col_metrica = st.columns([3, 1])
    with col_titulo:
        st.subheader("📋 Consultar y Gestionar Inventario")
    with col_metrica:
        st.metric(label="📊 Total de Equipos", value=len(df))
    
    busqueda = st.text_input("🔍 Buscar por MV, Material, Ubicación, Objeto técnico, etc.:")
    
    if not df.empty:
        if busqueda:
            mascara = df.apply(lambda row: row.astype(str).str.contains(busqueda, case=False).any(), axis=1)
            df_filtrado = df[mascara]
            st.write(f"Resultados encontrados: **{len(df_filtrado)}** de **{len(df)}**")
            st.dataframe(df_filtrado, use_container_width=True)
            df_mostrar = df_filtrado
        else:
            st.dataframe(df, use_container_width=True)
            df_mostrar = df

        # Acciones sobre los resultados de la búsqueda
        if not df_mostrar.empty:
            st.divider()
            st.subheader("⚡ Acciones Rápidas sobre Registro Seleccionado")
            
            lista_mvs = df_mostrar["MV"].astype(str).unique().tolist()
            mv_seleccionado = st.selectbox("Selecciona el código MV para gestionar:", lista_mvs)
            
            # Obtener datos del registro seleccionado
            registro_idx = df[df["MV"].astype(str) == mv_seleccionado].index[0]
            registro = df.loc[registro_idx]
            
            if es_admin:
                tab_editar, tab_eliminar = st.tabs(["✏️ Editar Registro", "🗑️ Eliminar Registro"])
                
                # TAB EDITAR DIRECTO
                with tab_editar:
                    with st.form("form_editar_directo"):
                        col1, col2 = st.columns(2)
                        with col1:
                            mv_edit = st.text_input("MV", value=str(registro["MV"]))
                            material_edit = st.text_input("Material", value=str(registro["Material"]))
                            denominacion_obj_edit = st.text_input("Denominación de objeto técnico", value=str(registro["Denominación de objeto técnico"]))
                            stat_sist_edit = st.text_input("Stat.sist.", value=str(registro["Stat.sist."]))
                            stat_usu_edit = st.text_input("StatUsu", value=str(registro["StatUsu"]))
                        with col2:
                            estatus_actual_edit = st.text_input("ESTATUS ACTUAL.", value=str(registro["ESTATUS ACTUAL."]))
                            denomin_edit = st.text_input("Denomin.", value=str(registro["Denomin."]))
                            ubicacion_edit = st.text_input("UBICACIÓN ACTUAL", value=str(registro["UBICACIÓN ACTUAL"]))
                            observaciones_edit = st.text_area("OBSERVACIONES", value=str(registro["OBSERVACIONES"]))
                            
                        actualizar = st.form_submit_button("💾 Guardar Cambios")
                        
                        if actualizar:
                            df = df.astype(object)
                            df.loc[registro_idx, "MV"] = str(mv_edit)
                            df.loc[registro_idx, "Material"] = str(material_edit)
                            df.loc[registro_idx, "Denominación de objeto técnico"] = str(denominacion_obj_edit)
                            df.loc[registro_idx, "Stat.sist."] = str(stat_sist_edit)
                            df.loc[registro_idx, "StatUsu"] = str(stat_usu_edit)
                            df.loc[registro_idx, "ESTATUS ACTUAL."] = str(estatus_actual_edit)
                            df.loc[registro_idx, "Denomin."] = str(denomin_edit)
                            df.loc[registro_idx, "UBICACIÓN ACTUAL"] = str(ubicacion_edit)
                            df.loc[registro_idx, "OBSERVACIONES"] = str(observaciones_edit)
                            
                            guardar_datos(df)
                            st.success("✅ Registro actualizado con éxito.")
                            st.rerun()

                # TAB ELIMINAR DIRECTO
                with tab_eliminar:
                    st.warning(f"⚠️ ¿Estás seguro de que deseas borrar permanentemente el equipo con MV: **{mv_seleccionado}**?")
                    if st.button("❌ Confirmar Eliminación"):
                        df = df[df["MV"].astype(str) != mv_seleccionado]
                        guardar_datos(df)
                        st.success(f"✅ El equipo {mv_seleccionado} ha sido eliminado.")
                        st.rerun()
            else:
                st.info("🔒 Estás en modo 'Visualizador'. Si necesitas editar o borrar registros, solicita acceso de Administrador.")
    else:
        st.info("El inventario está vacío actualmente.")

# 2. REGISTRAR (SOLO ADMIN)
elif opcion == "➕ Registrar Nuevo Equipo" and es_admin:
    st.subheader("➕ Registrar Nuevo Equipo")
    
    with st.form("form_agregar"):
        col1, col2 = st.columns(2)
        with col1:
            mv = st.text_input("MV / Identificador (Único)")
            material = st.text_input("Material")
            denominacion_obj = st.text_input("Denominación de objeto técnico")
            stat_sist = st.text_input("Stat.sist.")
            stat_usu = st.text_input("StatUsu")
        with col2:
            estatus_actual = st.selectbox("ESTATUS ACTUAL.", ["Bueno", "Regular", "Malo", "En revisión", "De baja", "Otro"])
            denomin = st.text_input("Denomin.")
            ubicacion = st.text_input("UBICACIÓN ACTUAL")
            observaciones = st.text_area("OBSERVACIONES")
            
        guardado = st.form_submit_button("Guardar Equipo")
        
        if guardado:
            if not mv:
                st.error("⚠️ El campo 'MV' es obligatorio ya que identifica al equipo.")
            elif str(mv) in df["MV"].astype(str).values:
                st.error(f"⚠️ El código MV '{mv}' ya existe en el inventario.")
            else:
                nuevo_registro = pd.DataFrame([{
                    "MV": str(mv),
                    "Material": str(material),
                    "Denominación de objeto técnico": str(denominacion_obj),
                    "Stat.sist.": str(stat_sist),
                    "StatUsu": str(stat_usu),
                    "ESTATUS ACTUAL.": str(estatus_actual),
                    "Denomin.": str(denomin),
                    "UBICACIÓN ACTUAL": str(ubicacion),
                    "OBSERVACIONES": str(observaciones)
                }])
                df = pd.concat([df, nuevo_registro], ignore_index=True)
                guardar_datos(df)
                st.success("✅ Equipo registrado con éxito.")
                st.rerun()

# 3. IMPORTAR / EXPORTAR (SOLO ADMIN)
elif opcion == "💾 Importar / Exportar Respaldos" and es_admin:
    st.subheader("📥 Exportar Respaldo (Excel)")
    st.write("Descarga una copia completa de la base de datos en formato Excel (.xlsx).")
    
    if not df.empty:
        import io
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Inventario')
        
        st.download_button(
            label="📥 Descargar Excel de Inventario",
            data=buffer.getvalue(),
            file_name="Inventario_Respaldo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("El inventario está vacío. No hay datos para exportar.")
        
    st.divider()
    
    st.subheader("📤 Importar Datos desde Excel")
    st.write("Carga tu archivo Excel para actualizar o restaurar los datos.")
    
    archivo_subido = st.file_uploader("Cargar archivo Excel (.xlsx)", type=["xlsx", "xls"])
    modo_importacion = st.radio("Modo de importación:", ["Reemplazar todo el inventario", "Anexar nuevos registros sin borrar los existentes"])
    
    if archivo_subido is not None:
        if st.button("Procesar e Importar"):
            try:
                df_nuevo = pd.read_excel(archivo_subido, dtype=str).fillna("")
                
                for col in COLUMNAS:
                    if col not in df_nuevo.columns:
                        df_nuevo[col] = ""
                df_nuevo = df_nuevo[COLUMNAS]
                
                if modo_importacion == "Reemplazar todo el inventario":
                    df_final = df_nuevo
                else:
                    df_final = pd.concat([df, df_nuevo], ignore_index=True).drop_duplicates(subset=["MV"], keep="last")
                
                guardar_datos(df_final)
                st.success("✅ Archivo importado correctamente.")
                st.rerun()
            except Exception as e:
                st.error(f"Error al leer el archivo Excel: {e}")

# 4. GESTIÓN DE USUARIOS (SOLO ADMIN)
elif opcion == "👥 Gestión de Usuarios" and es_admin:
    st.subheader("👥 Administración de Usuarios del Sistema")
    
    usuarios_dict = cargar_usuarios()
    
    col_lista, col_crear = st.columns([1, 1])
    
    with col_lista:
        st.write("### 📋 Usuarios Existentes")
        df_usuarios = pd.DataFrame([
            {"Usuario": u, "Rol": datos["rol"]} for u, datos in usuarios_dict.items()
        ])
        st.dataframe(df_usuarios, use_container_width=True)
        
        st.write("### 🗑️ Eliminar Usuario")
        usuario_a_borrar = st.selectbox("Selecciona un usuario a eliminar:", list(usuarios_dict.keys()))
        if st.button("Eliminar Usuario Seleccionado"):
            if usuario_a_borrar == st.session_state["usuario"]:
                st.error("⚠️ No puedes eliminar el usuario con el que tienes sesión iniciada.")
            elif usuario_a_borrar == "admin":
                st.error("⚠️ El usuario 'admin' principal no se puede eliminar.")
            else:
                del usuarios_dict[usuario_a_borrar]
                guardar_usuarios(usuarios_dict)
                st.success(f"✅ Usuario '{usuario_a_borrar}' eliminado.")
                st.rerun()

    with col_crear:
        st.write("### ➕ Crear Nuevo Usuario")
        with st.form("form_nuevo_usuario"):
            nuevo_user = st.text_input("Nombre de Usuario").strip()
            nueva_clave = st.text_input("Contraseña", type="password")
            nuevo_rol = st.selectbox("Rol del Usuario", ["Visualizador", "Administrador"])
            btn_crear_user = st.form_submit_button("Crear Usuario")
            
            if btn_crear_user:
                if not nuevo_user or not nueva_clave:
                    st.error("⚠️ Debes completar tanto el usuario como la contraseña.")
                elif nuevo_user in usuarios_dict:
                    st.error(f"⚠️ El usuario '{nuevo_user}' ya existe.")
                else:
                    usuarios_dict[nuevo_user] = {"clave": nueva_clave, "rol": nuevo_rol}
                    guardar_usuarios(usuarios_dict)
                    st.success(f"✅ Usuario '{nuevo_user}' creado exitosamente con el rol '{nuevo_rol}'.")
                    st.rerun()