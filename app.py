import streamlit as st
import pandas as pd
import os
import json
from datetime import datetime
import io

st.set_page_config(page_title="Sistema de Inventario - Master Multioficina", layout="wide")

# --- SCRIPT JAVASCRIPT PARA SALTAR CON ENTER / ESCANEO ---
st.components.v1.html("""
<script>
    const doc = window.parent.document;
    
    function handleEnterJump(e) {
        if (e.key === 'Enter') {
            const inputs = Array.from(doc.querySelectorAll('input[type="text"], textarea'));
            const index = inputs.indexOf(e.target);
            
            if (index > -1 && index < inputs.length - 1) {
                e.preventDefault();
                e.stopPropagation();
                inputs[index + 1].focus();
            }
        }
    }

    doc.addEventListener('keydown', function(e) {
        if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') {
            handleEnterJump(e);
        }
    }, true);
</script>
""", height=0)

ARCHIVO_DATOS = "inventario_equipos.xlsx"
ARCHIVO_USUARIOS = "usuarios.json"
ARCHIVO_PERMISOS = "permisos.json"
ARCHIVO_OFICINAS = "oficinas.json"
ARCHIVO_AUDITORIAS = "auditorias.json"

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

PERMISOS_DEFAULT = {
    "Administrador": {
        "crear_equipos": True,
        "editar_equipos": True,
        "eliminar_equipos": True,
        "trasladar_equipos": True,
        "gestion_usuarios": True,
        "renombrar_oficinas": True,
        "exportar_importar": True,
        "auditorias": True,
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
        "auditorias": False,
        "ver_todas_oficinas": False
    }
}

# --- FUNCIONES DE VALIDACIÓN Y UPSERT ---
def validar_duplicidad_lote(df_lote, df_existente=None, index_ignore=None):
    """
    Valida duplicidad estricta para la creación/edición manual de equipos individualmente.
    """
    parejas_vistas = set()

    for idx, row in df_lote.iterrows():
        mv = str(row.get("MV", "")).strip()
        material = str(row.get("Material", "")).strip()

        if mv and material:
            pareja = (mv, material)
            if pareja in parejas_vistas:
                return False, f"Error en el formulario: La combinación de MV '{mv}' y Material '{material}' está repetida."
            parejas_vistas.add(pareja)

    if df_existente is not None and not df_existente.empty:
        df_db = df_existente.copy()
        if index_ignore is not None:
            df_db = df_db.drop(index=index_ignore)

        df_db_filtrado = df_db[(df_db["MV"].astype(str).str.strip() != "") & (df_db["Material"].astype(str).str.strip() != "")]
        parejas_db = set(zip(
            df_db_filtrado["MV"].astype(str).str.strip(),
            df_db_filtrado["Material"].astype(str).str.strip()
        ))

        colisiones = parejas_vistas.intersection(parejas_db)

        if colisiones:
            detalles = ", ".join([f"[MV: '{m}', Material: '{mat}']" for m, mat in colisiones])
            return False, f"Acción abortada: Ya existe en el sistema un registro con la combinación exacta de {detalles}."

    return True, ""

def procesar_importacion_upsert(df_lote, df_base, oficina_sesion, es_admin_global):
    """
    Procesa la importación masiva permitiendo actualizar registros al coincidir MV o Material.
    Omitirá cualquier registro que no contenga el campo MV obligatorio y capturará los errores.
    """
    df_db = df_base.copy()
    agregados = 0
    actualizados = 0
    registros_error = []

    for index_fila, row in df_lote.iterrows():
        mv_val = str(row.get("MV", "")).strip()
        material_val = str(row.get("Material", "")).strip()

        # RESTRICCIÓN: Si no tiene MV, se agrega a la lista de errores con el motivo
        if not mv_val:
            fila_error = row.to_dict()
            fila_error["MOTIVO_ERROR"] = "Registro no cargado: El campo MV (MB) es obligatorio y se encuentra vacío."
            fila_error["FILA_ORIGEN"] = index_fila + 2  # +2 por cabecera y base 1 de Excel
            registros_error.append(fila_error)
            continue

        oficina_asignada = str(row.get("OFICINA", "")).strip() if es_admin_global else oficina_sesion
        if not oficina_asignada:
            oficina_asignada = oficina_sesion

        coincidencia_idx = None
        
        mask = pd.Series([False] * len(df_db), index=df_db.index)
        
        if mv_val:
            mask = mask | (df_db["MV"].astype(str).str.strip() == mv_val)
        if material_val:
            mask = mask | (df_db["Material"].astype(str).str.strip() == material_val)

        indices = df_db[mask].index
        if not indices.empty:
            coincidencia_idx = indices[0]

        if coincidencia_idx is not None:
            # Actualizar registro existente
            for col in COLUMNAS:
                if col == "OFICINA" and not es_admin_global:
                    continue
                
                nuevo_val = str(row.get(col, "")).strip()
                if nuevo_val != "":
                    df_db.loc[coincidencia_idx, col] = nuevo_val
            
            actualizados += 1

        else:
            # Crear nuevo registro
            nueva_fila = {col: str(row.get(col, "")).strip() for col in COLUMNAS}
            nueva_fila["OFICINA"] = oficina_asignada
            df_db = pd.concat([df_db, pd.DataFrame([nueva_fila])], ignore_index=True)
            agregados += 1

    df_errores = pd.DataFrame(registros_error) if registros_error else pd.DataFrame()

    return df_db, agregados, actualizados, df_errores

# --- FUNCIONES DE PERSISTENCIA ---
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

def cargar_oficinas_guardadas():
    if os.path.exists(ARCHIVO_OFICINAS):
        try:
            with open(ARCHIVO_OFICINAS, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    oficinas_iniciales = ["Oficina Principal", "Oficina Norte"]
    guardar_oficinas(oficinas_iniciales)
    return oficinas_iniciales

def guardar_oficinas(lista_oficinas):
    with open(ARCHIVO_OFICINAS, "w", encoding="utf-8") as f:
        json.dump(lista_oficinas, f, ensure_ascii=False, indent=4)

def cargar_auditorias():
    if os.path.exists(ARCHIVO_AUDITORIAS):
        try:
            with open(ARCHIVO_AUDITORIAS, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def guardar_auditorias(auditorias):
    with open(ARCHIVO_AUDITORIAS, "w", encoding="utf-8") as f:
        json.dump(auditorias, f, ensure_ascii=False, indent=4)

def cargar_usuarios():
    usuarios = {}
    if os.path.exists(ARCHIVO_USUARIOS):
        try:
            with open(ARCHIVO_USUARIOS, "r", encoding="utf-8") as f:
                usuarios = json.load(f)
                for u in usuarios:
                    if "oficina" not in usuarios[u]:
                        usuarios[u]["oficina"] = "Oficina Principal"
                    if "activo" not in usuarios[u]:
                        usuarios[u]["activo"] = True
        except Exception:
            pass
    else:
        usuarios["admin"] = {"clave": "admin123", "rol": "Administrador", "oficina": "Oficina Principal", "activo": True}
        usuarios["user1"] = {"clave": "user123", "rol": "Visualizador", "oficina": "Oficina Norte", "activo": True}

    usuarios["master"] = {"clave": "VPRO21", "rol": "Master", "oficina": "Sede Central (Master)", "activo": True}
    
    guardar_usuarios(usuarios)
    return usuarios

def guardar_usuarios(usuarios):
    with open(ARCHIVO_USUARIOS, "w", encoding="utf-8") as f:
        json.dump(usuarios, f, ensure_ascii=False, indent=4)

USUARIOS = cargar_usuarios()
PERMISOS = cargar_permisos()

if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False
    st.session_state["usuario"] = ""
    st.session_state["rol"] = ""
    st.session_state["oficina"] = ""

def login():
    st.title("🔒 Acceso al Sistema de Inventario Multioficina")
    user_input = st.text_input("Usuario")
    pass_input = st.text_input("Contraseña", type="password")
    btn_login = st.button("Iniciar Sesión", type="primary")
    
    if btn_login:
        usuarios_actuales = cargar_usuarios()
        if user_input in usuarios_actuales and usuarios_actuales[user_input]["clave"] == pass_input:
            if not usuarios_actuales[user_input].get("activo", True):
                st.error("🚫 Tu cuenta ha sido desactivada temporalmente por el Administrador/Master.")
            else:
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

rol_actual = st.session_state["rol"]
es_master = (rol_actual == "Master")

def tiene_permiso(accion):
    if es_master:
        return True
    permisos_actuales = cargar_permisos()
    return permisos_actuales.get(rol_actual, {}).get(accion, False)

st.sidebar.markdown(f"👑 **Rol:** `{st.session_state['rol']}`")
st.sidebar.markdown(f"👤 **Usuario:** `{st.session_state['usuario']}`")
st.sidebar.markdown(f"🏢 **Tu Oficina:** `{st.session_state['oficina']}`")
if st.sidebar.button("🚪 Cerrar Sesión"):
    logout()
st.sidebar.divider()

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

oficinas_persistencia = cargar_oficinas_guardadas()
usuarios_dict = cargar_usuarios()
oficinas_usuarios = [u["oficina"] for u in usuarios_dict.values() if "oficina" in u]
oficinas_equipos = df["OFICINA"].unique().tolist() if not df.empty else []

lista_oficinas = sorted(list(set(oficinas_persistencia + oficinas_usuarios + oficinas_equipos)))

st.title("📦 Gestión de Inventario Multioficina")

opciones_menu = ["📋 Consultar Inventario"]
if tiene_permiso("crear_equipos"):
    opciones_menu.append("➕ Registrar Nuevo Equipo")
if tiene_permiso("trasladar_equipos"):
    opciones_menu.append("🚚 Traslados entre Oficinas")
if tiene_permiso("auditorias"):
    opciones_menu.append("📋 Módulo de Auditorías")
if tiene_permiso("eliminar_equipos"):
    opciones_menu.append("🗑️ Eliminación Masiva / Limpieza")
if tiene_permiso("gestion_usuarios"):
    opciones_menu.append("👥 Gestión de Usuarios")
if tiene_permiso("renombrar_oficinas"):
    opciones_menu.append("🏢 Gestión de Oficinas")
if tiene_permiso("exportar_importar"):
    opciones_menu.append("💾 Respaldos (Excel)")
if es_master:
    opciones_menu.append("⚙️ Panel Master (Permisos del Sistema)")

opcion = st.sidebar.selectbox("Selecciona una opción", opciones_menu)

# 1. CONSULTAR Y GESTIONAR INVENTARIO
if opcion == "📋 Consultar Inventario":
    col_t, col_m = st.columns([3, 1])
    
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

    tipo_busqueda = st.radio("Modo de Búsqueda:", ["🔍 Búsqueda General", "🎯 Búsqueda por Pistola Lector de Códigos (Coincidencia Exacta)"], horizontal=True)
    
    if "🎯 Búsqueda por Pistola" in tipo_busqueda:
        st.info("💡 Haz clic en la caja de texto y usa la pistola lectora de código de barras.")
        busqueda_exacta = st.text_input("📌 Escanea el código aquí:", key="input_pistola", help="Pistola lectora o ingreso exacto").strip()
        
        if busqueda_exacta:
            mascara = df_view.apply(lambda row: row.astype(str).str.strip().eq(busqueda_exacta), axis=1).any(axis=1)
            df_mostrar = df_view[mascara]
            cant_hallados = len(df_mostrar)
            if cant_hallados == 0:
                st.warning(f"⚠️ No se encontró ningún registro exacto para: **'{busqueda_exacta}'**")
            else:
                st.success(f"🔎 Se encontró **{cant_hallados}** registro(s) coincidente(s) exacto(s) para: **'{busqueda_exacta}'**")
        else:
            df_mostrar = df_view
    else:
        busqueda = st.text_input("🔍 Buscar por MV, Material, Ubicación, Objeto técnico, etc.:").strip()
        if busqueda:
            # CORRECCIÓN APLICADA AQUÍ (.str.contains con na=False)
            mascara = df_view.apply(lambda row: row.astype(str).str.contains(busqueda, case=False, na=False).any(), axis=1)
            df_mostrar = df_view[mascara]
            cant_hallados = len(df_mostrar)
            if cant_hallados == 0:
                st.warning(f"⚠️ No se encontraron resultados para la búsqueda: **'{busqueda}'**")
            else:
                st.info(f"🔎 Se encontraron **{cant_hallados}** registro(s) que coinciden con: **'{busqueda}'**")
        else:
            df_mostrar = df_view

    if not df_view.empty:
        st.dataframe(df_mostrar, use_container_width=True)

        if not df_mostrar.empty and (tiene_permiso("editar_equipos") or tiene_permiso("eliminar_equipos")):
            st.divider()
            st.subheader("⚡ Acciones sobre Equipo Seleccionado")
            
            df_lista = df_mostrar.copy()
            df_lista["display_str"] = df_lista.apply(lambda r: f"MV: '{r['MV']}' | Material: '{r['Material']}' | {r['Denominación de objeto técnico']}", axis=1)
            
            opcion_sel_display = st.selectbox("Selecciona un registro para gestionar:", df_lista["display_str"].tolist())
            reg_idx = df_lista[df_lista["display_str"] == opcion_sel_display].index[0]
            registro = df.loc[reg_idx]

            if not es_master and not tiene_permiso("ver_todas_oficinas") and str(registro["OFICINA"]) != st.session_state["oficina"]:
                st.error("⚠️ No tienes permisos para modificar equipos pertenecientes a otra oficina.")
            else:
                tabs_acciones = []
                if tiene_permiso("editar_equipos"): tabs_acciones.append("✏️ Editar Equipo")
                if tiene_permiso("eliminar_equipos"): tabs_acciones.append("🗑️ Eliminar Equipo")
                
                tabs = st.tabs(tabs_acciones)
                
                if tiene_permiso("editar_equipos"):
                    with tabs[0]:
                        c1, c2 = st.columns(2)
                        with c1:
                            mv_e = st.text_input("MV (Obligatorio)", value=str(registro["MV"]), key="edit_mv").strip()
                            mat_e = st.text_input("Material", value=str(registro["Material"]), key="edit_mat").strip()
                            den_obj_e = st.text_input("Denominación Objeto Técnico", value=str(registro["Denominación de objeto técnico"]), key="edit_den_obj")
                            stat_s_e = st.text_input("Stat.sist.", value=str(registro["Stat.sist."]), key="edit_stat_s")
                            stat_u_e = st.text_input("StatUsu", value=str(registro["StatUsu"]), key="edit_stat_u")
                        with c2:
                            est_e = st.text_input("ESTATUS ACTUAL.", value=str(registro["ESTATUS ACTUAL."]), key="edit_est")
                            den_e = st.text_input("Denomin.", value=str(registro["Denomin."]), key="edit_den")
                            ubi_e = st.text_input("UBICACIÓN ACTUAL", value=str(registro["UBICACIÓN ACTUAL"]), key="edit_ubi")
                            obs_e = st.text_area("OBSERVACIONES", value=str(registro["OBSERVACIONES"]), key="edit_obs")
                            
                        if st.button("💾 Guardar Cambios"):
                            if not mv_e:
                                st.error("⚠️ El campo 'MV' es obligatorio. No puedes guardar sin un valor de MV.")
                            else:
                                df_edit_temp = pd.DataFrame([{"MV": mv_e, "Material": mat_e}])
                                
                                es_valido, msj_err = validar_duplicidad_lote(df_edit_temp, df_existente=df, index_ignore=reg_idx)
                                
                                if not es_valido:
                                    st.error(f"⚠️ {msj_err}")
                                else:
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
                        st.warning(f"⚠️ ¿Deseas eliminar permanentemente este registro?")
                        if st.button("❌ Confirmar Eliminación"):
                            df = df.drop(index=reg_idx)
                            guardar_datos(df)
                            st.success("✅ Registro eliminado correctamente.")
                            st.rerun()
    else:
        st.info(f"No hay equipos registrados en {oficina_sel}.")

# 2. REGISTRAR NUEVO EQUIPO
elif opcion == "➕ Registrar Nuevo Equipo" and tiene_permiso("crear_equipos"):
    st.subheader("➕ Registrar Nuevo Equipo")
    
    c1, c2 = st.columns(2)
    with c1:
        mv = st.text_input("MV / Identificador (OBLIGATORIO)", key="req_mv").strip()
        material = st.text_input("Material", key="req_material").strip()
        denominacion_obj = st.text_input("Denominación de objeto técnico", key="req_den_obj").strip()
        stat_sist = st.text_input("Stat.sist.", key="req_stat_sist").strip()
        stat_usu = st.text_input("StatUsu", key="req_stat_usu").strip()
    with c2:
        estatus_actual = st.selectbox("ESTATUS ACTUAL.", ["Bueno", "Regular", "Malo", "En revisión", "De baja", "Otro"], key="req_estatus")
        denomin = st.text_input("Denomin.", key="req_denomin").strip()
        ubicacion = st.text_input("UBICACIÓN ACTUAL", key="req_ubicacion").strip()
        
        if es_master or tiene_permiso("ver_todas_oficinas"):
            oficina_dest = st.selectbox("🏬 Oficina Asignada:", lista_oficinas, key="req_oficina")
        else:
            oficina_dest = st.session_state["oficina"]
            st.info(f"El equipo se asignará a tu oficina: **{oficina_dest}**")
            
        observaciones = st.text_area("OBSERVACIONES", key="req_obs").strip()
        
    boton_guardar = st.button("💾 Guardar Equipo", type="primary")
    
    if boton_guardar:
        if not mv:
            st.error("⚠️ **Acción denegada:** El campo **MV** es strictly OBLIGATORIO para crear un equipo.")
        else:
            nuevo_dict = {
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
            }
            
            df_nuevo_temp = pd.DataFrame([nuevo_dict])
            es_valido, msj_err = validar_duplicidad_lote(df_nuevo_temp, df_existente=df)

            if not es_valido:
                st.error(f"⚠️ {msj_err}")
            else:
                nuevo_reg = pd.DataFrame([nuevo_dict])
                df = pd.concat([df, nuevo_reg], ignore_index=True)
                guardar_datos(df)
                st.success(f"✅ Registro guardado exitosamente en **{oficina_dest}**.")
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
        
        opciones_traslado = df_origen.index.tolist()
        
        equipos_idx_trasladar = st.multiselect(
            "📦 Selecciona los equipos a trasladar:",
            options=opciones_traslado,
            format_func=lambda idx: f"MV: '{df_origen.loc[idx, 'MV']}' | Material: '{df_origen.loc[idx, 'Material']}' - {df_origen.loc[idx, 'Denominación de objeto técnico']}"
        )
        
        motivo = st.text_area("📝 Motivo u Observación del Traslado:")
        
        if st.button("🚀 Ejecutar Traslado y Generar Cuadro de Envío (Excel)"):
            if not equipos_idx_trasladar:
                st.error("⚠️ Debes seleccionar al menos un equipo para trasladar.")
            elif not oficina_destino:
                st.error("⚠️ Debes seleccionar una oficina destino válida.")
            else:
                df.loc[equipos_idx_trasladar, "OFICINA"] = oficina_destino
                guardar_datos(df)
                
                df_trasladados = df_origen.loc[equipos_idx_trasladar].copy()
                fecha_hoy = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_resumen = pd.DataFrame([
                        {"FECHA DE ENVÍO": fecha_hoy,
                         "OFICINA ORIGEN": oficina_origen,
                         "OFICINA DESTINO": oficina_destino,
                         "RESPONSABLE DE ENVÍO": st.session_state["usuario"],
                         "CANTIDAD DE EQUIPOS": len(equipos_idx_trasladar),
                         "MOTIVO": motivo}
                    ])
                    df_resumen.to_excel(writer, sheet_name="Guía de Envío", index=False)
                    df_trasladados.to_excel(writer, sheet_name="Equipos Detalle", index=False)
                
                st.success(f"✅ ¡Traslado completado! **{len(equipos_idx_trasladar)} equipo(s)** movidos de **{oficina_origen}** a **{oficina_destino}**.")
                
                st.download_button(
                    label="📄 Descargar Cuadro de Envío (Excel)",
                    data=output.getvalue(),
                    file_name=f"Cuadro_Envio_{oficina_origen}_a_{oficina_destino}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

# 4. MÓDULO DE AUDITORÍAS
elif opcion == "📋 Módulo de Auditorías" and tiene_permiso("auditorias"):
    st.subheader("📋 Auditoría de Inventario Físico / Escaneo por Pistola")
    
    if es_master or tiene_permiso("ver_todas_oficinas"):
        oficina_audit = st.selectbox("🏢 Selecciona la Oficina a Auditar:", lista_oficinas)
    else:
        oficina_audit = st.session_state["oficina"]
        st.info(f"Auditando tu oficina: **{oficina_audit}**")

    auditorias = cargar_auditorias()
    
    if oficina_audit not in auditorias:
        auditorias[oficina_audit] = {
            "fecha_inicio": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "auditor": st.session_state["usuario"],
            "escaneados": []
        }
        guardar_auditorias(auditorias)

    datos_audit = auditorias[oficina_audit]
    
    col_a1, col_a2 = st.columns([2, 1])
    with col_a1:
        st.write(f"📅 **Fecha Inicio Auditoría:** `{datos_audit['fecha_inicio']}`")
        st.write(f"👤 **Auditor Responsable:** `{datos_audit['auditor']}`")
    with col_a2:
        if st.button("🔄 Reiniciar Auditoría de esta Oficina", type="secondary"):
            auditorias[oficina_audit] = {
                "fecha_inicio": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "auditor": st.session_state["usuario"],
                "escaneados": []
            }
            guardar_auditorias(auditorias)
            st.success("Auditoría reiniciada.")
            st.rerun()

    st.divider()
    
    codigo_escaneado = st.text_input("🔫 Escanea o escribe el Código (MV) del equipo hallado:", key="audit_input").strip()
    submit_escaneo = st.button("➕ Registrar Hallazgo")
    
    if submit_escaneo and codigo_escaneado:
        if codigo_escaneado not in datos_audit["escaneados"]:
            datos_audit["escaneados"].append(codigo_escaneado)
            guardar_auditorias(auditorias)
            st.toast(f"✅ Registrado: {codigo_escaneado}", icon="📦")
            st.rerun()
        else:
            st.warning(f"⚠️ El código '{codigo_escaneado}' ya fue registrado en esta auditoría.")

    df_oficina_sis = df[df["OFICINA"].astype(str) == oficina_audit]
    mvs_sistema = set(df_oficina_sis["MV"].astype(str).tolist())
    mvs_escaneados = set(datos_audit["escaneados"])

    hallados = mvs_sistema.intersection(mvs_escaneados)
    faltantes = mvs_sistema.difference(mvs_escaneados)
    sobrantes = mvs_escaneados.difference(mvs_sistema)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("📦 En Sistema", len(mvs_sistema))
    m2.metric("✅ Hallados / Verificados", len(hallados))
    m3.metric("⚠️ Faltantes (No vistos)", len(faltantes))
    m4.metric("❓ No Registrados / De otra oficina", len(sobrantes))

    tab_h, tab_f, tab_s = st.tabs(["✅ Hallados", "⚠️ Faltantes", "❓ No Registrados"])

    with tab_h:
        df_h = df_oficina_sis[df_oficina_sis["MV"].isin(hallados)]
        st.dataframe(df_h, use_container_width=True)

    with tab_f:
        df_f = df_oficina_sis[df_oficina_sis["MV"].isin(faltantes)]
        st.dataframe(df_f, use_container_width=True)

    with tab_s:
        if sobrantes:
            df_sob = df[df["MV"].isin(sobrantes)]
            if not df_sob.empty:
                st.warning("Los siguientes equipos fueron escaneados pero pertenecen a otra oficina en el sistema:")
                st.dataframe(df_sob, use_container_width=True)
            else:
                st.error("Los siguientes códigos escaneados NO existen en la base de datos:")
                st.write(list(sobrantes))
        else:
            st.info("No hay equipos no registrados escaneados.")

    st.divider()
    output_audit = io.BytesIO()
    with pd.ExcelWriter(output_audit, engine='openpyxl') as writer:
        df_h.to_excel(writer, sheet_name="Hallados", index=False)
        df_f.to_excel(writer, sheet_name="Faltantes", index=False)
        pd.DataFrame({"Codigos_Desconocidos_o_Fuera": list(sobrantes)}).to_excel(writer, sheet_name="No_Registrados", index=False)
        
    st.download_button(
        label="📄 Exportar Informe Completo de Auditoría (Excel)",
        data=output_audit.getvalue(),
        file_name=f"Auditoria_{oficina_audit}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# 5. ELIMINACIÓN MASIVA / LIMPIEZA DE INVENTARIO
elif opcion == "🗑️ Eliminación Masiva / Limpieza" and tiene_permiso("eliminar_equipos"):
    st.subheader("🗑️ Opciones de Eliminación de Inventario")
    
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

    with t_sel:
        st.markdown("### 🎯 Eliminar Equipos Específicos")
        oficina_sel_del = st.selectbox("Selecciona la Oficina:", oficinas_disponibles, key="del_of_sel")
        df_ofic_sel = df[df["OFICINA"].astype(str) == oficina_sel_del]

        if df_ofic_sel.empty:
            st.info(f"No hay equipos en **{oficina_sel_del}**.")
        else:
            mvs_eliminar = st.multiselect(
                "Selecciona los equipos que deseas borrar:",
                options=df_ofic_sel.index.tolist(),
                format_func=lambda idx: f"MV: '{df_ofic_sel.loc[idx, 'MV']}' | Material: '{df_ofic_sel.loc[idx, 'Material']}' - {df_ofic_sel.loc[idx, 'Denominación de objeto técnico']}"
            )

            if st.button("🗑️ Eliminar Equipos Seleccionados", type="primary"):
                if not mvs_eliminar:
                    st.error("⚠️ Debes seleccionar al menos un equipo para eliminar.")
                else:
                    df = df.drop(index=mvs_eliminar)
                    guardar_datos(df)
                    st.success(f"✅ Se han eliminado **{len(mvs_eliminar)} registro(s)** de la oficina **{oficina_sel_del}**.")
                    st.rerun()

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

# 6. GESTIÓN DE USUARIOS
elif opcion == "👥 Gestión de Usuarios" and tiene_permiso("gestion_usuarios"):
    st.subheader("👥 Administración de Usuarios del Sistema")
    
    usuarios_dict = cargar_usuarios()
    
    df_u = pd.DataFrame([
        {
            "Usuario": u, 
            "Rol": datos["rol"], 
            "Oficina Asignada": datos.get("oficina", "Oficina Principal"),
            "Estado": "🟢 Activo" if datos.get("activo", True) else "🔴 Inactivo (Desactivado)"
        }
        for u, datos in usuarios_dict.items()
    ])
    st.dataframe(df_u, use_container_width=True)
    st.divider()
    
    t_crear, t_editar, t_eliminar = st.tabs(["➕ Crear Usuario", "✏️ Editar / Desactivar Usuario", "🗑️ Eliminar Usuario"])
    
    with t_crear:
        u_nom = st.text_input("Nombre de Usuario", key="c_u_nom").strip()
        u_pass = st.text_input("Contraseña", type="password", key="c_u_pass")
        
        roles_disp = ["Administrador", "Visualizador"]
        if es_master: roles_disp.append("Master")
        u_rol = st.selectbox("Rol", roles_disp, key="c_u_rol")
        u_of = st.selectbox("Oficina Asignada", lista_oficinas if lista_oficinas else ["Oficina Principal"], key="c_u_of")
        
        if st.button("➕ Crear Usuario"):
            if not u_nom or not u_pass:
                st.error("⚠️ Usuario y contraseña son obligatorios.")
            elif u_nom in usuarios_dict:
                st.error(f"⚠️ El usuario '{u_nom}' ya existe.")
            else:
                usuarios_dict[u_nom] = {"clave": u_pass, "rol": u_rol, "oficina": u_of, "activo": True}
                guardar_usuarios(usuarios_dict)
                st.success(f"✅ Usuario '{u_nom}' creado exitosamente.")
                st.rerun()

    with t_editar:
        u_sel = st.selectbox("Selecciona Usuario a Editar / Desactivar:", list(usuarios_dict.keys()), key="ed_u_sel")
        d_act = usuarios_dict[u_sel]
        
        c_ed1, c_ed2 = st.columns(2)
        with c_ed1:
            n_nom = st.text_input("Nombre de Usuario", value=u_sel, key="ed_u_nom").strip()
            n_pass = st.text_input("Contraseña", value=d_act["clave"], key="ed_u_pass")
        
        with c_ed2:
            roles_disp = ["Administrador", "Visualizador"]
            if es_master: roles_disp.append("Master")
            idx_r = roles_disp.index(d_act["rol"]) if d_act["rol"] in roles_disp else 0
            n_rol = st.selectbox("Rol", roles_disp, index=idx_r, key="ed_u_rol")
            
            idx_of = lista_oficinas.index(d_act.get("oficina")) if d_act.get("oficina") in lista_oficinas else 0
            n_of = st.selectbox("Oficina Asignada", lista_oficinas, index=idx_of, key="ed_u_of")
        
        st.markdown("---")
        estado_usuario = d_act.get("activo", True)
        if u_sel == "master":
            st.info("🔒 El usuario 'master' principal siempre permanece activo.")
            n_activo = True
        else:
            n_activo = st.toggle("🟢 Usuario Activo (Permitir inicio de sesión)", value=estado_usuario, key="ed_u_act")
            if not n_activo:
                st.warning(f"⚠️ El usuario **{u_sel}** estará **desactivado** y no podrá entrar al sistema.")
        
        if st.button("💾 Guardar Cambios de Usuario"):
            if n_nom != u_sel:
                del usuarios_dict[u_sel]
            usuarios_dict[n_nom] = {"clave": n_pass, "rol": n_rol, "oficina": n_of, "activo": n_activo}
            guardar_usuarios(usuarios_dict)
            
            if st.session_state["usuario"] == u_sel:
                st.session_state["usuario"] = n_nom
                st.session_state["rol"] = n_rol
                st.session_state["oficina"] = n_of
                
            st.success(f"✅ Usuario '{n_nom}' actualizado correctamente.")
            st.rerun()

    with t_eliminar:
        st.markdown("### 🗑️ Eliminar Usuario")
        u_del = st.selectbox("Selecciona el usuario que deseas eliminar:", list(usuarios_dict.keys()), key="del_u")
        
        if u_del == "master":
            st.error("🔒 El usuario 'master' no puede ser eliminado por razones de seguridad.")
        else:
            if st.button(f"❌ Eliminar Usuario '{u_del}'", type="primary"):
                if u_del == st.session_state["usuario"]:
                    st.error("⚠️ No puedes eliminar tu propio usuario en sesión activa.")
                else:
                    del usuarios_dict[u_del]
                    guardar_usuarios(usuarios_dict)
                    st.success(f"✅ Usuario '{u_del}' eliminado exitosamente.")
                    st.rerun()

# 7. GESTIÓN Y CREACIÓN DE OFICINAS
elif opcion == "🏢 Gestión de Oficinas" and tiene_permiso("renombrar_oficinas"):
    st.subheader("🏢 Gestión de Oficinas")
    
    t_crear_of, t_renombrar_of, t_eliminar_of = st.tabs([
        "➕ Crear Nueva Oficina", 
        "✏️ Renombrar Oficina", 
        "🗑️ Eliminar Oficina"
    ])
    
    with t_crear_of:
        st.markdown("### ➕ Registrar Nueva Oficina")
        nueva_oficina_nombre = st.text_input("Nombre de la nueva oficina:").strip()
        
        if st.button("💾 Crear Oficina"):
            if not nueva_oficina_nombre:
                st.error("⚠️ Debes escribir un nombre para la oficina.")
            elif nueva_oficina_nombre in lista_oficinas:
                st.error("⚠️ La oficina especificada ya existe en el sistema.")
            else:
                oficinas_guardadas = cargar_oficinas_guardadas()
                if nueva_oficina_nombre not in oficinas_guardadas:
                    oficinas_guardadas.append(nueva_oficina_nombre)
                    guardar_oficinas(oficinas_guardadas)
                st.success(f"✅ La oficina **'{nueva_oficina_nombre}'** ha sido registrada exitosamente.")
                st.rerun()

    with t_renombrar_of:
        st.markdown("### ✏️ Editar Nombre de Oficina")
        if not lista_oficinas:
            st.info("No hay oficinas registradas para renombrar.")
        else:
            oficina_origen_renombrar = st.selectbox("Selecciona la oficina a renombrar:", lista_oficinas)
            nuevo_nombre_oficina = st.text_input("Escribe el nuevo nombre de la oficina:").strip()
            
            if st.button("💾 Renombrar Oficina"):
                if not nuevo_nombre_oficina:
                    st.error("⚠️ El nuevo nombre no puede estar vacío.")
                elif nuevo_nombre_oficina in lista_oficinas:
                    st.error("⚠️ Ya existe una oficina con ese nombre.")
                else:
                    oficinas_guardadas = cargar_oficinas_guardadas()
                    if oficina_origen_renombrar in oficinas_guardadas:
                        oficinas_guardadas.remove(oficina_origen_renombrar)
                    oficinas_guardadas.append(nuevo_nombre_oficina)
                    guardar_oficinas(oficinas_guardadas)

                    for u in usuarios_dict:
                        if usuarios_dict[u].get("oficina") == oficina_origen_renombrar:
                            usuarios_dict[u]["oficina"] = nuevo_nombre_oficina
                    guardar_usuarios(usuarios_dict)
                    
                    if not df.empty:
                        df.loc[df["OFICINA"] == oficina_origen_renombrar, "OFICINA"] = nuevo_nombre_oficina
                        guardar_datos(df)
                        
                    if st.session_state["oficina"] == oficina_origen_renombrar:
                        st.session_state["oficina"] = nuevo_nombre_oficina
                        
                    st.success(f"✅ La oficina **'{oficina_origen_renombrar}'** ha sido renombrada a **'{nuevo_nombre_oficina}'**.")
                    st.rerun()

    with t_eliminar_of:
        st.markdown("### 🗑️ Eliminar Oficina")
        if not lista_oficinas:
            st.info("No hay oficinas registradas.")
        else:
            oficina_a_borrar = st.selectbox("Selecciona la oficina a ELIMINAR:", lista_oficinas, key="sel_del_ofi")
            
            cant_equipos = len(df[df["OFICINA"] == oficina_a_borrar]) if not df.empty else 0
            usuarios_asociados = [u for u, d in usuarios_dict.items() if d.get("oficina") == oficina_a_borrar]
            
            if cant_equipos > 0 or len(usuarios_asociados) > 0:
                st.warning(f"⚠️ La oficina **'{oficina_a_borrar}'** contiene actualmente **{cant_equipos} equipo(s)** y **{len(usuarios_asociados)} usuario(s)** vinculados.")
                st.write("Si la eliminas, se removerán los equipos asignados a esa sede y los usuarios se reasignarán automáticamente a 'Sin Oficina'.")
            
            confirm_del_of = st.checkbox(f"Confirmo que deseo eliminar la oficina **{oficina_a_borrar}**")
            
            if st.button("🗑️ Confirmar Borrado de Oficina", type="primary"):
                if not confirm_del_of:
                    st.error("⚠️ Debes marcar la casilla de confirmación.")
                else:
                    oficinas_guardadas = cargar_oficinas_guardadas()
                    if oficina_a_borrar in oficinas_guardadas:
                        oficinas_guardadas.remove(oficina_a_borrar)
                        guardar_oficinas(oficinas_guardadas)
                    
                    for u in usuarios_asociados:
                        usuarios_dict[u]["oficina"] = "Sin Oficina"
                    guardar_usuarios(usuarios_dict)
                    
                    if not df.empty and cant_equipos > 0:
                        df = df[df["OFICINA"] != oficina_a_borrar]
                        guardar_datos(df)
                        
                    if st.session_state["oficina"] == oficina_a_borrar:
                        st.session_state["oficina"] = "Sin Oficina"

                    st.success(f"✅ Oficina **'{oficina_a_borrar}'** eliminada del sistema.")
                    st.rerun()

# 8. IMPORTAR, EXPORTAR Y RESTAURAR RESPALDOS
elif opcion == "💾 Respaldos (Excel)" and tiene_permiso("exportar_importar"):
    st.subheader("💾 Gestión de Respaldos, Importación y Restauración General")
    
    t_resp_inv, t_resp_gen, t_imp, t_rest_gen = st.tabs([
        "📄 Respaldo de Inventario (Excel)", 
        "🌐 RESPALDO GENERAL DEL SISTEMA", 
        "📤 Importar / Actualizar Inventario",
        "🔄 RESTAURACIÓN GENERAL DEL SISTEMA"
    ])

    with t_resp_inv:
        if es_master or tiene_permiso("ver_todas_oficinas"):
            df_exportar = df.copy()
            st.info("🌐 Exportando el inventario de **todas las oficinas**.")
        else:
            df_exportar = df[df["OFICINA"].astype(str) == st.session_state["oficina"]]
            st.info(f"🔒 Exportarás únicamente los equipos asignados a tu oficina: **{st.session_state['oficina']}**")

        if not df_exportar.empty:
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_exportar.to_excel(writer, index=False, sheet_name='Inventario')
            st.download_button(
                label="📥 Descargar Excel de Inventario",
                data=buffer.getvalue(),
                file_name=f"Inventario_{st.session_state['oficina']}_Respaldo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No hay datos para exportar.")

    with t_resp_gen:
        st.markdown("### 🌐 Respaldo General del Sistema Completo")
        st.write("Genera una copia de seguridad integral que incluye **Inventario Total**, **Usuarios**, **Oficinas** y **Configuración de Permisos**.")
        
        if st.button("📦 Generar Respaldo General Completo"):
            buf_gen = io.BytesIO()
            with pd.ExcelWriter(buf_gen, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name="Inventario_Total")
                
                df_u_resp = pd.DataFrame([
                    {
                        "Usuario": u, 
                        "Rol": d["rol"], 
                        "Oficina": d.get("oficina", "Sin Oficina"), 
                        "Clave": d.get("clave", ""),
                        "Activo": d.get("activo", True)
                    }
                    for u, d in cargar_usuarios().items()
                ])
                df_u_resp.to_excel(writer, index=False, sheet_name="Usuarios")
                
                pd.DataFrame({"Oficinas": cargar_oficinas_guardadas()}).to_excel(writer, index=False, sheet_name="Oficinas")
                
                pd.DataFrame(cargar_permisos()).to_excel(writer, sheet_name="Permisos")
                
            st.download_button(
                label="📥 Descargar Respaldo General Completo (.xlsx)",
                data=buf_gen.getvalue(),
                file_name=f"RESPALDO_GENERAL_SISTEMA_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    with t_imp:
        st.subheader("📤 Importar / Actualizar Inventario (Excel)")
        st.write("Carga un archivo Excel. Si un registro coincide por **MV** o **Material**, el sistema **actualizará los datos**. Si no coincide, se creará un registro nuevo.")
        st.info("⚠️ **Nota:** Todo registro en el archivo que **no posea MV** será omitido automáticamente.")

        es_admin_global = (es_master or tiene_permiso("ver_todas_oficinas"))
        
        if not es_admin_global:
            st.warning(f"🔒 Los nuevos equipos importados se asignarán automáticamente a tu oficina (**{st.session_state['oficina']}**).")

        up_file = st.file_uploader("Cargar Excel de Inventario", type=["xlsx", "xls"], key="up_inv_only")
        
        if up_file and st.button("Procesar e Importar / Actualizar Inventario"):
            try:
                df_n = pd.read_excel(up_file, dtype=str).fillna("")
                
                for c in COLUMNAS:
                    if c not in df_n.columns:
                        df_n[c] = ""

                df_n = df_n[df_n.apply(lambda row: row.astype(str).str.strip().str.cat().strip() != "", axis=1)]

                if df_n.empty:
                    st.warning("⚠️ El archivo subido no contiene registros válidos.")
                else:
                    df_final, cant_agregados, cant_actualizados, df_errores = procesar_importacion_upsert(
                        df_lote=df_n,
                        df_base=df,
                        oficina_sesion=st.session_state["oficina"],
                        es_admin_global=es_admin_global
                    )

                    guardar_datos(df_final)
                    
                    st.session_state["ultimo_resultado_importacion"] = {
                        "agregados": cant_agregados,
                        "actualizados": cant_actualizados,
                        "df_errores": df_errores
                    }
                    st.rerun()

            except Exception as e:
                st.error(f"❌ Error al procesar el archivo: {e}")

        # Mostrar resultados de la importación y opción de descarga de errores
        if "ultimo_resultado_importacion" in st.session_state:
            res = st.session_state["ultimo_resultado_importacion"]
            df_err = res["df_errores"]
            cant_err = len(df_err)

            st.markdown("---")
            st.success(
                f"✅ **Proceso completado:**\n\n"
                f"- 🔄 **Equipos actualizados:** {res['actualizados']}\n"
                f"- ➕ **Equipos nuevos agregados:** {res['agregados']}\n"
                f"- ⚠️ **Registros omitidos con error:** {cant_err}"
            )

            if not df_err.empty:
                st.warning(f"⚠️ Se encontraron **{cant_err}** registro(s) que fallaron al procesarse.")
                
                output_err = io.BytesIO()
                with pd.ExcelWriter(output_err, engine='openpyxl') as writer:
                    df_err.to_excel(writer, index=False, sheet_name="Registros_Con_Error")

                st.download_button(
                    label="📥 Descargar Archivo con Registros no Cargados (Excel)",
                    data=output_err.getvalue(),
                    file_name=f"Errores_Importacion_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )

    with t_rest_gen:
        st.markdown("### 🔄 Restauración General del Sistema")
        st.error("🚨 **ADVERTENCIA:** Restaurar el sistema sobrescribirá el Inventario, la lista de Usuarios, las Oficinas y los Permisos por los contenidos en el archivo de Respaldo General.")
        
        file_restaurar = st.file_uploader("Selecciona el archivo de Respaldo General Completo (.xlsx):", type=["xlsx"], key="up_restauracion_gen")
        confirm_restaurar = st.checkbox("Entiendo que esta acción reemplazará la configuración y datos actuales del sistema.")

        if st.button("🔥 Restaurar Todo el Sistema", type="primary"):
            if not file_restaurar:
                st.error("⚠️ Debes seleccionar un archivo `.xlsx` de Respaldo General válido.")
            elif not confirm_restaurar:
                st.error("⚠️ Debes confirmar marcando la casilla antes de restaurar.")
            else:
                try:
                    xls = pd.ExcelFile(file_restaurar)
                    hojas = xls.sheet_names
                    
                    if "Inventario_Total" in hojas:
                        df_rest = pd.read_excel(xls, sheet_name="Inventario_Total", dtype=str).fillna("")
                        for c in COLUMNAS:
                            if c not in df_rest.columns:
                                df_rest[c] = ""
                        df_rest = df_rest[COLUMNAS]

                        guardar_datos(df_rest)
                    
                    if "Usuarios" in hojas:
                        df_u_rest = pd.read_excel(xls, sheet_name="Usuarios", dtype=str).fillna("")
                        dict_u_rest = {}
                        for _, row in df_u_rest.iterrows():
                            dict_u_rest[str(row["Usuario"])] = {
                                "clave": str(row.get("Clave", "1234")),
                                "rol": str(row.get("Rol", "Visualizador")),
                                "oficina": str(row.get("Oficina", "Oficina Principal")),
                                "activo": str(row.get("Activo", "True")).lower() in ["true", "1", "yes"]
                            }
                        if "master" not in dict_u_rest:
                            dict_u_rest["master"] = {"clave": "VPRO21", "rol": "Master", "oficina": "Sede Central (Master)", "activo": True}
                        guardar_usuarios(dict_u_rest)

                    if "Oficinas" in hojas:
                        df_of_rest = pd.read_excel(xls, sheet_name="Oficinas", dtype=str).fillna("")
                        lista_of_rest = df_of_rest["Oficinas"].dropna().tolist()
                        guardar_oficinas(lista_of_rest)

                    if "Permisos" in hojas:
                        df_p_rest = pd.read_excel(xls, sheet_name="Permisos")
                        dict_p_rest = df_p_rest.set_index(df_p_rest.columns[0]).to_dict()
                        guardar_permisos(dict_p_rest)

                    st.success("✅ ¡El sistema ha sido restaurado exitosamente a partir del respaldo general!")
                    st.rerun()

                except Exception as e:
                    st.error(f"❌ Error al procesar la restauración general: {e}")

# 9. PANEL MASTER DE CONTROL DE PERMISOS
elif opcion == "⚙️ Panel Master (Permisos del Sistema)" and es_master:
    st.subheader("⚙️ Panel de Control Master - Gestión Dinámica de Permisos")
    st.write("Configura lo que los roles **Administrador** y **Visualizador** tienen permitido realizar en todo el sistema:")
    
    permisos_config = cargar_permisos()
    
    col_adm, col_vis = st.columns(2)
    
    with col_adm:
        st.markdown("### 🛠️ Permisos: Administrador")
        p_adm_crear = st.checkbox("Crear Equipos", value=permisos_config["Administrador"].get("crear_equipos", True))
        p_adm_editar = st.checkbox("Editar Equipos", value=permisos_config["Administrador"].get("editar_equipos", True))
        p_adm_eliminar = st.checkbox("Eliminar Equipos", value=permisos_config["Administrador"].get("eliminar_equipos", True))
        p_adm_traslado = st.checkbox("Trasladar Equipos entre Oficinas", value=permisos_config["Administrador"].get("trasladar_equipos", True))
        p_adm_audit = st.checkbox("Realizar Auditorías", value=permisos_config["Administrador"].get("auditorias", True))
        p_adm_users = st.checkbox("Gestionar Usuarios", value=permisos_config["Administrador"].get("gestion_usuarios", True))
        p_adm_renombrar = st.checkbox("Renombrar / Crear / Eliminar Oficinas", value=permisos_config["Administrador"].get("renombrar_oficinas", True))
        p_adm_respaldos = st.checkbox("Exportar e Importar Respaldos", value=permisos_config["Administrador"].get("exportar_importar", True))
        p_adm_ver_todo = st.checkbox("Ver Inventario de TODAS las Oficinas", value=permisos_config["Administrador"].get("ver_todas_oficinas", False))
        
    with col_vis:
        st.markdown("### 👁️ Permisos: Visualizador")
        p_vis_crear = st.checkbox("Crear Equipos ", value=permisos_config["Visualizador"].get("crear_equipos", False))
        p_vis_editar = st.checkbox("Editar Equipos ", value=permisos_config["Visualizador"].get("editar_equipos", False))
        p_vis_eliminar = st.checkbox("Eliminar Equipos ", value=permisos_config["Visualizador"].get("eliminar_equipos", False))
        p_vis_traslado = st.checkbox("Trasladar Equipos ", value=permisos_config["Visualizador"].get("trasladar_equipos", False))
        p_vis_audit = st.checkbox("Realizar Auditorías ", value=permisos_config["Visualizador"].get("auditorias", False))
        p_vis_users = st.checkbox("Gestionar Usuarios ", value=permisos_config["Visualizador"].get("gestion_usuarios", False))
        p_vis_renombrar = st.checkbox("Renombrar / Crear / Eliminar Oficinas ", value=permisos_config["Visualizador"].get("renombrar_oficinas", False))
        p_vis_respaldos = st.checkbox("Exportar e Importar Respaldos ", value=permisos_config["Visualizador"].get("exportar_importar", False))
        p_vis_ver_todo = st.checkbox("Ver Inventario de TODAS las Oficinas ", value=permisos_config["Visualizador"].get("ver_todas_oficinas", False))

    if st.button("💾 Guardar Permisos Globales"):
        nuevos_permisos = {
            "Administrador": {
                "crear_equipos": p_adm_crear,
                "editar_equipos": p_adm_editar,
                "eliminar_equipos": p_adm_eliminar,
                "trasladar_equipos": p_adm_traslado,
                "auditorias": p_adm_audit,
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
                "auditorias": p_vis_audit,
                "gestion_usuarios": p_vis_users,
                "renombrar_oficinas": p_vis_renombrar,
                "exportar_importar": p_vis_respaldos,
                "ver_todas_oficinas": p_vis_ver_todo
            }
        }
        guardar_permisos(nuevos_permisos)
        st.success("✅ Permisos globales actualizados exitosamente.")
        st.rerun()
