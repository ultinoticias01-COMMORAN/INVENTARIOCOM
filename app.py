import streamlit as st
import pandas as pd
import os
import json
from datetime import datetime
import io

st.set_page_config(page_title="Sistema de Inventario - Master Multioficina", layout="wide")

ARCHIVO_DATOS = "inventario_equipos.xlsx"
ARCHIVO_USUARIOS = "usuarios.json"
ARCHIVO_PERMISOS = "permisos.json"
ARCHIVO_HISTORIAL_TRASLADOS = "historial_traslados.xlsx"

# Columnas estándar del inventario
COLUMNAS = [
    "MV",
    "Material",
    "Denominación de objeto técnico",
    "Stat.sist.",
    "StatUsu",
    "ESTATUS ACTUAL.",
    "Denomin.",
    "UBICACIÓN ACTUAL",
    "OFICINA",
    "OBSERVACIONES"
]

# --- PERMISOS POR DEFECTO ---
PERMISOS_DEFAULT = {
    "Administrador": {
        "crear_equipos": True,
        "editar_equipos": True,
        "eliminar_equipos": True,
        "trasladar_equipos": True,
        "gestion_usuarios": True,
        "renombrar_oficinas": True,
        "exportar_importar": True,
        "ver_todas_oficinas": False
    },
    "Visualizador": {
        "crear_equipos": False,
        "editar_equipos": False,
        "eliminar_equipos": False,
        "trasladar_equipos": False,
        "gestion_usuarios": False,
        "renombrar_oficinas": False,
        "exportar_importar": False,
        "ver_todas_oficinas": False
    }
}

def cargar_permisos():
    if os.path.exists(ARCHIVO_PERMISOS):
        try:
            with open(ARCHIVO_PERMISOS, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    guardar_permisos(PERMISOS_DEFAULT)
    return PERMISOS_DEFAULT

def guardar_permisos(permisos):
    with open(ARCHIVO_PERMISOS, "w", encoding="utf-8") as f:
        json.dump(permisos, f, ensure_ascii=False, indent=4)

# --- GESTIÓN DE USUARIOS ---
def cargar_usuarios():
    usuarios = {}
    if os.path.exists(ARCHIVO_USUARIOS):
        try:
            with open(ARCHIVO_USUARIOS, "r", encoding="utf-8") as f:
                usuarios = json.load(f)
                for u in usuarios:
                    if "oficina" not in usuarios[u]:
                        usuarios[u]["oficina"] = "Oficina Principal"
        except Exception:
            pass
            
    # Forzar actualización/creación de usuario master con clave VPRO21
    usuarios["master"] = {"clave": "VPRO21", "rol": "Master", "oficina": "Sede Central (Master)"}
    
    if "admin" not in usuarios:
        usuarios["admin"] = {"clave": "admin123", "rol": "Administrador", "oficina": "Oficina Principal"}
    if "user1" not in usuarios:
        usuarios["user1"] = {"clave": "user123", "rol": "Visualizador", "oficina": "Oficina Norte"}
        
    guardar_usuarios(usuarios)
    return usuarios

def guardar_usuarios(usuarios):
    with open(ARCHIVO_USUARIOS, "w", encoding="utf-8") as f:
        json.dump(usuarios, f, ensure_ascii=False, indent=4)

# --- INICIALIZAR PERMISOS Y USUARIOS ---
USUARIOS = cargar_usuarios()
PERMISOS = cargar_permisos()

# --- CONTROL DE SESIÓN ---
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False
    st.session_state["usuario"] = ""
    st.session_state["rol"] = ""
    st.session_state["oficina"] = ""

def login():
    st.title("🔒 Acceso al Sistema de Inventario Multioficina")
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
                st.session_state["oficina"] = usuarios_actuales[user_input].get("oficina", "Oficina Principal")
                st.success(f"Bienvenido {user_input} [{usuarios_actuales[user_input]['rol']}]")
                st.rerun()
            else:
                st.error("⚠️ Usuario o contraseña incorrectos.")

def logout():
    st.session_state["autenticado"] = False
    st.session_state["usuario"] = ""
    st.session_state["rol"] = ""
    st.session_state["oficina"] = ""
    st.rerun()

if not st.session_state["autenticado"]:
    login()
    st.stop()

# --- VERIFICACIÓN DE PERMISOS DINÁMICOS ---
rol_actual = st.session_state["rol"]
es_master = (rol_actual == "Master")

def tiene_permiso(accion):
    if es_master:
        return True
    permisos_actuales = cargar_permisos()
    return permisos_actuales.get(rol_actual, {}).get(accion, False)

# Sidebar
st.sidebar.markdown(f"👑 **Rol:** `{st.session_state['rol']}`")
st.sidebar.markdown(f"👤 **Usuario:** `{st.session_state['usuario']}`")
st.sidebar.markdown(f"🏢 **Tu Oficina:** `{st.session_state['oficina']}`")
if st.sidebar.button("🚪 Cerrar Sesión"):
    logout()
st.sidebar.divider()

# --- BASE DE DATOS INVENTARIO ---
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

# Lista de oficinas registradas
usuarios_dict = cargar_usuarios()
oficinas_usuarios = [u["oficina"] for u in usuarios_dict.values() if "oficina" in u]
oficinas_equipos = df["OFICINA"].unique().tolist() if not df.empty else []
lista_oficinas = sorted(list(set(oficinas_usuarios + oficinas_equipos + ["Oficina Principal"])))

st.title("📦 Gestión de Inventario Multioficina")

# --- MENÚ DE NAVEGACIÓN ---
opciones_menu = ["📋 Consultar Inventario"]
if tiene_permiso("crear_equipos"):
    opciones_menu.append("➕ Registrar Nuevo Equipo")
if tiene_permiso("trasladar_equipos"):
    opciones_menu.append("🚚 Traslados entre Oficinas")
if tiene_permiso("eliminar_equipos"):
    opciones_menu.append("🗑️ Eliminación Masiva / Limpieza")
if tiene_permiso("gestion_usuarios"):
    opciones_menu.append("👥 Gestión de Usuarios")
if tiene_permiso("renombrar_oficinas"):
    opciones_menu.append("🏢 Editar Nombres de Oficinas")
if tiene_permiso("exportar_importar"):
    opciones_menu.append("💾 Respaldos (Excel)")
if es_master:
    opciones_menu.append("⚙️ Panel Master (Permisos del Sistema)")

opcion = st.sidebar.selectbox("Selecciona una opción", opciones_menu)

# 1. CONSULTAR Y GESTIONAR INVENTARIO
if opcion == "📋 Consultar Inventario":
    col_t, col_m = st.columns([3, 1])
    
    # Control estricto de visibilidad por Oficina
    if es_master or tiene_permiso("ver_todas_oficinas"):
        oficinas_filtro = ["Todas las Oficinas"] + lista_oficinas
        oficina_sel = st.sidebar.selectbox("🏬 Filtrar por Oficina:", oficinas_filtro)
    else:
        oficina_sel = st.session_state["oficina"]
        st.sidebar.info(f"🔒 Mostrando únicamente tu oficina: **{oficina_sel}**")

    if oficina_sel == "Todas las Oficinas":
        df_view = df.copy()
    else:
        df_view = df[df["OFICINA"].astype(str) == oficina_sel]

    with col_t:
        st.subheader(f"📋 Equipos en: {oficina_sel}")
    with col_m:
        st.metric(label="📊 Total Equipos", value=len(df_view))

    busqueda = st.text_input("🔍 Buscar por MV, Material, Ubicación, Objeto técnico, etc.:")
    
    if not df_view.empty:
        if busqueda:
            mascara = df_view.apply(lambda row: row.astype(str).str.contains(busqueda, case=False).any(), axis=1)
            df_mostrar = df_view[mascara]
        else:
            df_mostrar = df_view
            
        st.dataframe(df_mostrar, use_container_width=True)

        if not df_mostrar.empty and (tiene_permiso("editar_equipos") or tiene_permiso("eliminar_equipos")):
            st.divider()
            st.subheader("⚡ Acciones sobre Equipo Seleccionado")
            
            lista_mvs = df_mostrar["MV"].astype(str).unique().tolist()
            mv_sel = st.selectbox("Selecciona el código MV para gestionar:", lista_mvs)
            
            reg_idx = df[df["MV"].astype(str) == mv_sel].index[0]
            registro = df.loc[reg_idx]

            # Verificación de seguridad de oficina
            if not es_master and not tiene_permiso("ver_todas_oficinas") and str(registro["OFICINA"]) != st.session_state["oficina"]:
                st.error("⚠️ No tienes permisos para modificar equipos pertenecientes a otra oficina.")
            else:
                tabs_acciones = []
                if tiene_permiso("editar_equipos"): tabs_acciones.append("✏️ Editar Equipo")
                if tiene_permiso("eliminar_equipos"): tabs_acciones.append("🗑️ Eliminar Equipo")
                
                tabs = st.tabs(tabs_acciones)
                
                if tiene_permiso("editar_equipos"):
                    with tabs[0]:
                        with st.form("form_edit_eq"):
                            c1, c2 = st.columns(2)
                            with c1:
                                mv_e = st.text_input("MV", value=str(registro["MV"]))
                                mat_e = st.text_input("Material", value=str(registro["Material"]))
                                den_obj_e = st.text_input("Denominación Objeto Técnico", value=str(registro["Denominación de objeto técnico"]))
                                stat_s_e = st.text_input("Stat.sist.", value=str(registro["Stat.sist."]))
                                stat_u_e = st.text_input("StatUsu", value=str(registro["StatUsu"]))
                            with c2:
                                est_e = st.text_input("ESTATUS ACTUAL.", value=str(registro["ESTATUS ACTUAL."]))
                                den_e = st.text_input("Denomin.", value=str(registro["Denomin."]))
                                ubi_e = st.text_input("UBICACIÓN ACTUAL", value=str(registro["UBICACIÓN ACTUAL"]))
                                obs_e = st.text_area("OBSERVACIONES", value=str(registro["OBSERVACIONES"]))
                                
                            if st.form_submit_button("💾 Guardar Cambios"):
                                df.loc[reg_idx, "MV"] = str(mv_e)
                                df.loc[reg_idx, "Material"] = str(mat_e)
                                df.loc[reg_idx, "Denominación de objeto técnico"] = str(den_obj_e)
                                df.loc[reg_idx, "Stat.sist."] = str(stat_s_e)
                                df.loc[reg_idx, "StatUsu"] = str(stat_u_e)
                                df.loc[reg_idx, "ESTATUS ACTUAL."] = str(est_e)
                                df.loc[reg_idx, "Denomin."] = str(den_e)
                                df.loc[reg_idx, "UBICACIÓN ACTUAL"] = str(ubi_e)
                                df.loc[reg_idx, "OBSERVACIONES"] = str(obs_e)
                                guardar_datos(df)
                                st.success("✅ Equipo actualizado correctamente.")
                                st.rerun()

                if tiene_permiso("eliminar_equipos"):
                    idx_tab_del = 1 if tiene_permiso("editar_equipos") else 0
                    with tabs[idx_tab_del]:
                        st.warning(f"⚠️ ¿Deseas eliminar permanentemente el equipo **{mv_sel}**?")
                        if st.button("❌ Confirmar Eliminación"):
                            df = df[df["MV"].astype(str) != mv_sel]
                            guardar_datos(df)
                            st.success(f"✅ Equipo {mv_sel} eliminado.")
                            st.rerun()
    else:
        st.info(f"No hay equipos registrados en {oficina_sel}.")

# 2. REGISTRAR NUEVO EQUIPO
elif opcion == "➕ Registrar Nuevo Equipo" and tiene_permiso("crear_equipos"):
    st.subheader("➕ Registrar Nuevo Equipo")
    
    with st.form("form_agregar"):
        c1, c2 = st.columns(2)
        with c1:
            mv = st.text_input("MV / Identificador (Único)")
            material = st.text_input("Material")
            denominacion_obj = st.text_input("Denominación de objeto técnico")
            stat_sist = st.text_input("Stat.sist.")
            stat_usu = st.text_input("StatUsu")
        with c2:
            estatus_actual = st.selectbox("ESTATUS ACTUAL.", ["Bueno", "Regular", "Malo", "En revisión", "De baja", "Otro"])
            denomin = st.text_input("Denomin.")
            ubicacion = st.text_input("UBICACIÓN ACTUAL")
            
            # Si no es master, asigna automáticamente a su oficina
            if es_master or tiene_permiso("ver_todas_oficinas"):
                oficina_dest = st.selectbox("🏬 Oficina Asignada:", lista_oficinas)
            else:
                oficina_dest = st.session_state["oficina"]
                st.info(f"El equipo se asignará a tu oficina: **{oficina_dest}**")
                
            observaciones = st.text_area("OBSERVACIONES")
            
        if st.form_submit_button("Guardar Equipo"):
            if not mv:
                st.error("⚠️ El campo 'MV' es obligatorio.")
            elif str(mv) in df["MV"].astype(str).values:
                st.error(f"⚠️ El código MV '{mv}' ya existe.")
            else:
                nuevo_reg = pd.DataFrame([{
                    "MV": str(mv),
                    "Material": str(material),
                    "Denominación de objeto técnico": str(denominacion_obj),
                    "Stat.sist.": str(stat_sist),
                    "StatUsu": str(stat_usu),
                    "ESTATUS ACTUAL.": str(estatus_actual),
                    "Denomin.": str(denomin),
                    "UBICACIÓN ACTUAL": str(ubicacion),
                    "OFICINA": str(oficina_dest),
                    "OBSERVACIONES": str(observaciones)
                }])
                df = pd.concat([df, nuevo_reg], ignore_index=True)
                guardar_datos(df)
                st.success(f"✅ Equipo registrado en **{oficina_dest}**.")
                st.rerun()

# 3. TRASLADOS ENTRE OFICINAS
elif opcion == "🚚 Traslados entre Oficinas" and tiene_permiso("trasladar_equipos"):
    st.subheader("🚚 Traslado de Equipos entre Oficinas con Cuadro de Envío")
    
    if es_master or tiene_permiso("ver_todas_oficinas"):
        oficina_origen = st.selectbox("🏢 Oficina Origen del Traslado:", lista_oficinas)
    else:
        oficina_origen = st.session_state["oficina"]
        st.info(f"Oficina Origen: **{oficina_origen}**")

    df_origen = df[df["OFICINA"].astype(str) == oficina_origen]
    
    if df_origen.empty:
        st.warning(f"No hay equipos en **{oficina_origen}** para realizar traslados.")
    else:
        oficinas_destino_opt = [o for o in lista_oficinas if o != oficina_origen]
        oficina_destino = st.selectbox("🎯 Oficina Destino:", oficinas_destino_opt)
        
        equipos_a_trasladar = st.multiselect(
            "📦 Selecciona los equipos a trasladar (por MV / Objeto):",
            options=df_origen["MV"].tolist(),
            format_func=lambda x: f"MV: {x} - {df_origen[df_origen['MV']==x]['Denominación de objeto técnico'].values[0]}"
        )
        
        motivo = st.text_area("📝 Motivo u Observación del Traslado:")
        
        if st.button("🚀 Ejecutar Traslado y Generar Cuadro de Envío (Excel)"):
            if not equipos_a_trasladar:
                st.error("⚠️ Debes seleccionar al menos un equipo para trasladar.")
            elif not oficina_destino:
                st.error("⚠️ Debes seleccionar una oficina destino válida.")
            else:
                df.loc[df["MV"].isin(equipos_a_trasladar), "OFICINA"] = oficina_destino
                guardar_datos(df)
                
                df_trasladados = df_origen[df_origen["MV"].isin(equipos_a_trasladar)].copy()
                fecha_hoy = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_resumen = pd.DataFrame([
                        {"FECHA DE ENVÍO": fecha_hoy,
                         "OFICINA ORIGEN": oficina_origen,
                         "OFICINA DESTINO": oficina_destino,
                         "RESPONSABLE DE ENVÍO": st.session_state["usuario"],
                         "CANTIDAD DE EQUIPOS": len(equipos_a_trasladar),
                         "MOTIVO": motivo}
                    ])
                    df_resumen.to_excel(writer, sheet_name="Guía de Envío", index=False)
                    df_trasladados.to_excel(writer, sheet_name="Equipos Detalle", index=False)
                
                st.success(f"✅ ¡Traslado completado! **{len(equipos_a_trasladar)} equipo(s)** movidos de **{oficina_origen}** a **{oficina_destino}**.")
                
                st.download_button(
                    label="📄 Descargar Cuadro de Envío (Excel)",
                    data=output.getvalue(),
                    file_name=f"Cuadro_Envio_{oficina_origen}_a_{oficina_destino}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

# 4. ELIMINACIÓN MASIVA / LIMPIEZA DE INVENTARIO
elif opcion == "🗑️ Eliminación Masiva / Limpieza" and tiene_permiso("eliminar_equipos"):
    st.subheader("🗑️ Opciones de Eliminación de Inventario")
    
    # Determinar visibilidad de oficinas para borrado
    if es_master or tiene_permiso("ver_todas_oficinas"):
        oficinas_disponibles = lista_oficinas
    else:
        oficinas_disponibles = [st.session_state["oficina"]]
        st.info(f"🔒 Las acciones de borrado aplicarán únicamente a tu oficina: **{st.session_state['oficina']}**")

    t_sel, t_ofi, t_todo = st.tabs([
        "🎯 Borrar Equipos Seleccionados", 
        "🏢 Borrar Inventario por Oficina", 
        "💥 Borrar Todo el Inventario"
    ])

    # Tab 1: Borrar Equipos Seleccionados (Multiselección)
    with t_sel:
        st.markdown("### 🎯 Eliminar Equipos Específicos")
        oficina_sel_del = st.selectbox("Selecciona la Oficina:", oficinas_disponibles, key="del_of_sel")
        df_ofic_sel = df[df["OFICINA"].astype(str) == oficina_sel_del]

        if df_ofic_sel.empty:
            st.info(f"No hay equipos en **{oficina_sel_del}**.")
        else:
            mvs_eliminar = st.multiselect(
                "Selecciona los equipos que deseas borrar:",
                options=df_ofic_sel["MV"].tolist(),
                format_func=lambda x: f"MV: {x} - {df_ofic_sel[df_ofic_sel['MV']==x]['Denominación de objeto técnico'].values[0]}"
            )

            if st.button("🗑️ Eliminar Equipos Seleccionados", type="primary"):
                if not mvs_eliminar:
                    st.error("⚠️ Debes seleccionar al menos un equipo para eliminar.")
                else:
                    df = df[~df["MV"].isin(mvs_eliminar)]
                    guardar_datos(df)
                    st.success(f"✅ Se han eliminado **{len(mvs_eliminar)} equipo(s)** de la oficina **{oficina_sel_del}**.")
                    st.rerun()

    # Tab 2: Borrar Todo el Inventario de una Oficina Especifica
    with t_ofi:
        st.markdown("### 🏢 Borrar Todo el Inventario de una Oficina")
        oficina_a_vaciar = st.selectbox("Selecciona la oficina que deseas VACIAR por completo:", oficinas_disponibles, key="vac_ofi")
        
        cant_equipos_ofi = len(df[df["OFICINA"].astype(str) == oficina_a_vaciar])
        st.warning(f"⚠️ La oficina **{oficina_a_vaciar}** actualmente tiene **{cant_equipos_ofi} equipos**.")
        
        check_confirm_ofi = st.checkbox(f"Confirmo que deseo ELIMINAR TODOS los {cant_equipos_ofi} equipos de {oficina_a_vaciar}")
        
        if st.button(f"🔥 Borrar Todo de {oficina_a_vaciar}", type="primary"):
            if not check_confirm_ofi:
                st.error("⚠️ Por seguridad, debes marcar la casilla de verificación antes de proceder.")
            elif cant_equipos_ofi == 0:
                st.info("La oficina seleccionada no contiene equipos para eliminar.")
            else:
                df = df[df["OFICINA"].astype(str) != oficina_a_vaciar]
                guardar_datos(df)
                st.success(f"✅ Se vació por completo el inventario de la oficina **{oficina_a_vaciar}**.")
                st.rerun()

    # Tab 3: Borrar TODO el Inventario del Sistema
    with t_todo:
        st.markdown("### 💥 Borrar TODO el Inventario General")
        if not (es_master or tiene_permiso("ver_todas_oficinas")):
            st.error("⚠️ Esta acción solo la puede realizar un usuario **Master** o con permisos globales sobre todas las oficinas.")
        else:
            cant_total = len(df)
            st.error(f"🚨 **¡ATENCIÓN!** Esta acción borrará permanentemente los **{cant_total} equipos** registrados en TODAS las oficinas del sistema.")
            
            confirm_texto = st.text_input("Para confirmar, escribe la palabra **ELIMINAR** abajo:")
            check_confirm_todo = st.checkbox("Entiendo que esta acción borrará TODO el inventario y no se puede deshacer.")
            
            if st.button("💥 VACIAR INVENTARIO COMPLETO", type="primary"):
                if confirm_texto.strip().upper() != "ELIMINAR":
                    st.error("⚠️ Debes escribir la palabra 'ELIMINAR' para confirmar.")
                elif not check_confirm_todo:
                    st.error("⚠️ Debes marcar la casilla de verificación de confirmación.")
                else:
                    df = pd.DataFrame(columns=COLUMNAS)
                    guardar_datos(df)
                    st.success("✅ Todo el inventario del sistema ha sido eliminado por completo.")
                    st.rerun()

# 5. GESTIÓN DE USUARIOS
elif opcion == "👥 Gestión de Usuarios" and tiene_permiso("gestion_usuarios"):
    st.subheader("👥 Administración de Usuarios del Sistema")
    
    usuarios_dict = cargar_usuarios()
    
    df_u = pd.DataFrame([
        {"Usuario": u, "Rol": datos["rol"], "Oficina Asignada": datos.get("oficina", "Oficina Principal")}
        for u, datos in usuarios_dict.items()
    ])
    st.dataframe(df_u, use_container_width=True)
    st.divider()
    
    t_crear, t_editar, t_eliminar = st.tabs(["➕ Crear Usuario", "✏️ Editar Usuario", "🗑️ Eliminar Usuario"])
    
    with t_crear:
        with st.form("f_crear_u"):
            u_nom = st.text_input("Nombre de Usuario").strip()
            u_pass = st.text_input("Contraseña", type="password")
            
            roles_disp = ["Administrador", "Visualizador"]
            if es_master: roles_disp.append("Master")
            u_rol = st.selectbox("Rol", roles_disp)
            
            u_of = st.selectbox("Oficina Asignada", lista_oficinas)
            
            if st.form_submit_button("➕ Crear Usuario"):
                if not u_nom or not u_pass:
                    st.error("⚠️ Usuario y contraseña son obligatorios.")
                elif u_nom in usuarios_dict:
                    st.error(f"⚠️ El usuario '{u_nom}' ya existe.")
                else:
                    usuarios_dict[u_nom] = {"clave": u_pass, "rol": u_rol, "oficina": u_of}
                    guardar_usuarios(usuarios_dict)
                    st.success(f"✅ Usuario '{u_nom}' creado exitosamente.")
                    st.rerun()

    with t_editar:
        u_sel = st.selectbox("Selecciona Usuario a Editar:", list(usuarios_dict.keys()))
        d_act = usuarios_dict[u_sel]
        
        with st.form("f_edit_u"):
            n_nom = st.text_input("Nuevo Usuario", value=u_sel).strip()
            n_pass = st.text_input("Nueva Contraseña", value=d_act["clave"])
            
            roles_disp = ["Administrador", "Visualizador"]
            if es_master: roles_disp.append("Master")
            idx_r = roles_disp.index(d_act["rol"]) if d_act["rol"] in roles_disp else 0
            n_rol = st.selectbox("Rol", roles_disp, index=idx_r)
            
            idx_of = lista_oficinas.index(d_act.get("oficina", "Oficina Principal")) if d_act.get("oficina") in lista_oficinas else 0
            n_of = st.selectbox("Oficina Asignada", lista_oficinas, index=idx_of)
            
            if st.form_submit_button("💾 Guardar Cambios"):
                if n_nom != u_sel:
                    del usuarios_dict[u_sel]
                usuarios_dict[n_nom] = {"clave": n_pass, "rol": n_rol, "oficina": n_of}
                guardar_usuarios(usuarios_dict)
                
                if st.session_state["usuario"] == u_sel:
                    st.session_state["usuario"] = n_nom
                    st.session_state["rol"] = n_rol
                    st.session_state["oficina"] = n_of
                    
                st.success(f"✅ Usuario '{n_nom}' actualizado.")
                st.rerun()

    with t_eliminar:
        u_del = st.selectbox("Selecciona Usuario a Eliminar:", list(usuarios_dict.keys()), key="del_u")
        if st.button("❌ Eliminar Usuario"):
            if u_del == st.session_state["usuario"]:
                st.error("⚠️ No puedes eliminar tu propio usuario activo.")
            else:
                del usuarios_dict[u_del]
                guardar_usuarios(usuarios_dict)
                st.success(f"✅ Usuario '{u_del}' eliminado.")
                st.rerun()

# 6. EDITAR NOMBRES DE LAS OFICINAS
elif opcion == "🏢 Editar Nombres de Oficinas" and tiene_permiso("renombrar_oficinas"):
    st.subheader("🏢 Edición y Renombrado de Oficinas")
    st.write("Cambia el nombre de una oficina. Se actualizarán automáticamente los usuarios y los equipos asociados.")
    
    oficina_origen_renombrar = st.selectbox("Selecciona la oficina a renombrar:", lista_oficinas)
    nuevo_nombre_oficina = st.text_input("Escribe el nuevo nombre de la oficina:").strip()
    
    if st.button("💾 Renombrar Oficina"):
        if not nuevo_nombre_oficina:
            st.error("⚠️ El nuevo nombre no puede estar vacío.")
        elif nuevo_nombre_oficina in lista_oficinas:
            st.error("⚠️ Ya existe una oficina con ese nombre.")
        else:
            for u in usuarios_dict:
                if usuarios_dict[u].get("oficina") == oficina_origen_renombrar:
                    usuarios_dict[u]["oficina"] = nuevo_nombre_oficina
            guardar_usuarios(usuarios_dict)
            
            if not df.empty:
                df.loc[df["OFICINA"] == oficina_origen_renombrar, "OFICINA"] = nuevo_nombre_oficina
                guardar_datos(df)
                
            if st.session_state["oficina"] == oficina_origen_renombrar:
                st.session_state["oficina"] = nuevo_nombre_oficina
                
            st.success(f"✅ La oficina **'{oficina_origen_renombrar}'** ha sido renombrada con éxito a **'{nuevo_nombre_oficina}'**.")
            st.rerun()

# 7. IMPORTAR Y EXPORTAR RESPALDOS
elif opcion == "💾 Respaldos (Excel)" and tiene_permiso("exportar_importar"):
    st.subheader("📥 Exportar Respaldo General (Excel)")
    if not df.empty:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Inventario')
        st.download_button(
            label="📥 Descargar Excel de Inventario",
            data=buffer.getvalue(),
            file_name="Inventario_General_Respaldo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    st.divider()
    st.subheader("📤 Importar Respaldo (Excel)")
    up_file = st.file_uploader("Cargar Excel", type=["xlsx", "xls"])
    if up_file and st.button("Procesar e Importar"):
        try:
            df_n = pd.read_excel(up_file, dtype=str).fillna("")
            for c in COLUMNAS:
                if c not in df_n.columns:
                    df_n[c] = "Oficina Principal" if c == "OFICINA" else ""
            df_final = df_n[COLUMNAS]
            guardar_datos(df_final)
            st.success("✅ Datos importados correctamente.")
            st.rerun()
        except Exception as e:
            st.error(f"Error al importar: {e}")

# 8. PANEL MASTER DE CONTROL DE PERMISOS
elif opcion == "⚙️ Panel Master (Permisos del Sistema)" and es_master:
    st.subheader("⚙️ Panel de Control Master - Gestión Dinámica de Permisos")
    st.write("Configura lo que los roles **Administrador** y **Visualizador** tienen permitido realizar en todo el sistema:")
    
    permisos_config = cargar_permisos()
    
    with st.form("form_permisos_master"):
        col_adm, col_vis = st.columns(2)
        
        with col_adm:
            st.markdown("### 🛠️ Permisos: Administrador")
            p_adm_crear = st.checkbox("Crear Equipos", value=permisos_config["Administrador"].get("crear_equipos", True))
            p_adm_editar = st.checkbox("Editar Equipos", value=permisos_config["Administrador"].get("editar_equipos", True))
            p_adm_eliminar = st.checkbox("Eliminar Equipos", value=permisos_config["Administrador"].get("eliminar_equipos", True))
            p_adm_traslado = st.checkbox("Trasladar Equipos entre Oficinas", value=permisos_config["Administrador"].get("trasladar_equipos", True))
            p_adm_users = st.checkbox("Gestionar Usuarios", value=permisos_config["Administrador"].get("gestion_usuarios", True))
            p_adm_renombrar = st.checkbox("Renombrar Oficinas", value=permisos_config["Administrador"].get("renombrar_oficinas", True))
            p_adm_respaldos = st.checkbox("Exportar e Importar Respaldos", value=permisos_config["Administrador"].get("exportar_importar", True))
            p_adm_ver_todo = st.checkbox("Ver Inventario de TODAS las Oficinas", value=permisos_config["Administrador"].get("ver_todas_oficinas", False))
            
        with col_vis:
            st.markdown("### 👁️ Permisos: Visualizador")
            p_vis_crear = st.checkbox("Crear Equipos ", value=permisos_config["Visualizador"].get("crear_equipos", False))
            p_vis_editar = st.checkbox("Editar Equipos ", value=permisos_config["Visualizador"].get("editar_equipos", False))
            p_vis_eliminar = st.checkbox("Eliminar Equipos ", value=permisos_config["Visualizador"].get("eliminar_equipos", False))
            p_vis_traslado = st.checkbox("Trasladar Equipos ", value=permisos_config["Visualizador"].get("trasladar_equipos", False))
            p_vis_users = st.checkbox("Gestionar Usuarios ", value=permisos_config["Visualizador"].get("gestion_usuarios", False))
            p_vis_renombrar = st.checkbox("Renombrar Oficinas ", value=permisos_config["Visualizador"].get("renombrar_oficinas", False))
            p_vis_respaldos = st.checkbox("Exportar e Importar Respaldos ", value=permisos_config["Visualizador"].get("exportar_importar", False))
            p_vis_ver_todo = st.checkbox("Ver Inventario de TODAS las Oficinas ", value=permisos_config["Visualizador"].get("ver_todas_oficinas", False))

        if st.form_submit_button("💾 Guardar Permisos Globales"):
            nuevos_permisos = {
                "Administrador": {
                    "crear_equipos": p_adm_crear,
                    "editar_equipos": p_adm_editar,
                    "eliminar_equipos": p_adm_eliminar,
                    "trasladar_equipos": p_adm_traslado,
                    "gestion_usuarios": p_adm_users,
                    "renombrar_oficinas": p_adm_renombrar,
                    "exportar_importar": p_adm_respaldos,
                    "ver_todas_oficinas": p_adm_ver_todo
                },
                "Visualizador": {
                    "crear_equipos": p_vis_crear,
                    "editar_equipos": p_vis_editar,
                    "eliminar_equipos": p_vis_eliminar,
                    "trasladar_equipos": p_vis_traslado,
                    "gestion_usuarios": p_vis_users,
                    "renombrar_oficinas": p_vis_renombrar,
                    "exportar_importar": p_vis_respaldos,
                    "ver_todas_oficinas": p_vis_ver_todo
                }
            }
            guardar_permisos(nuevos_permisos)
            st.success("✅ Permisos globales actualizados exitosamente.")
            st.rerun()

