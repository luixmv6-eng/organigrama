import streamlit as st
import pandas as pd
import os
import re
import unicodedata
import json
from datetime import datetime
from dateutil.relativedelta import relativedelta
import plotly.express as px
import plotly.graph_objects as go
import streamlit.components.v1 as components


# Configuración de página
st.set_page_config(page_title="Portal de Talento Humano", layout="wide", page_icon="👥")

# --- ESTILOS CSS GENERALES ---
st.markdown("""
    <style>
    .main-header { font-size: 2.2rem; font-weight: 700; color: #1E3A8A; text-align: center; margin-bottom: 25px; }
    .search-container { max-width: 800px; margin: 0 auto 10px auto; }
    .leader-header { background: #1E3A8A; color: white; padding: 20px; border-radius: 12px; margin-bottom: 20px; }
    .group-card { background: #F8FAFC; border-radius: 8px; padding: 12px; border-left: 5px solid #1E3A8A; margin-bottom: 8px; }
    .badge-leader { background: #3B82F6; color: white; padding: 2px 8px; border-radius: 10px; font-size: 0.75rem; margin-left: 5px; }
    .area-badge { background: #F1F5F9; color: #475569; padding: 2px 6px; border-radius: 4px; font-size: 0.7rem; margin-left: 10px; border: 1px solid #E2E8F0; }
    .count-badge { background: #1E3A8A; color: white; padding: 2px 8px; border-radius: 12px; font-size: 0.8rem; float: right; }
    </style>
""", unsafe_allow_html=True)

# --- FUNCIONES DE APOYO ---
def format_fecha_v2(f_str):
    if pd.isna(f_str): return "N/A"
    try:
        f = pd.to_datetime(f_str)
        meses = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
        return f"{f.day} de {meses[f.month-1]} de {f.year}"
    except: return str(f_str)

def calcular_tiempo_completo_v2(fecha):
    if pd.isna(fecha): return "N/A"
    try:
        f = pd.to_datetime(fecha)
        hoy = datetime.now()
        diff = relativedelta(hoy, f)
        res = []
        if diff.years > 0: res.append(f"{diff.years} {'año' if diff.years == 1 else 'años'}")
        if diff.months > 0: res.append(f"{diff.months} {'mes' if diff.months == 1 else 'meses'}")
        if diff.days > 0: res.append(f"{diff.days} {'día' if diff.days == 1 else 'días'}")
        if not res: return "0 días"
        return ", ".join(res[:-1]) + " y " + res[-1] if len(res) > 1 else res[0]
    except: return "N/A"

def calcular_edad_v3(fecha):
    if pd.isna(fecha): return "N/A"
    try:
        f = pd.to_datetime(fecha)
        hoy = datetime.now()
        return hoy.year - f.year - ((hoy.month, hoy.day) < (f.month, f.day))
    except: return "N/A"

def categorizar_region(mun):
    if not isinstance(mun, str): return 'Otras Regiones'
    m = mun.upper()
    sur_valle = ['CALI', 'PALMIRA', 'JAMUNDI', 'YUMBO', 'GUADALAJARA DE BUGA', 'CANDELARIA', 'EL CERRITO', 'FLORIDA', 'PRADERA', 'GUACARI', 'VIJES', 'LA CUMBRE', 'SAN PEDRO', 'CERRITO']
    norte_valle = ['CARTAGO', 'TULUA', 'ZARZAL', 'SEVILLA', 'ROLDANILLO', 'BUGALAGRANDE', 'ANDALUCIA', 'LA UNION', 'CAICEDONIA', 'OBANDO', 'ANSERMANUEVO', 'TORO', 'LA VICTORIA', 'RIOFRIO', 'ARGELIA', 'VERSALLES', 'BOLIVAR', 'TRUJILLO']
    cauca = ['PUERTO TEJADA', 'MIRANDA', 'CORINTO', 'SANTANDER DE QUILICHAO', 'GUACHENE', 'CALOTO', 'VILLA RICA', 'PADILLA', 'POPAYAN', 'GUACHENE']
    if m in sur_valle: return 'Sur del Valle'
    if m in norte_valle: return 'Norte del Valle'
    if m in cauca: return 'Cauca'
    return 'Otras Regiones'

# --- CARGA Y UNIFICACIÓN ---
def _get_excel_mtime():
    """Retorna la fecha de modificación del Excel para invalidar caché automáticamente."""
    xlsx = 'HC_ajustado.xlsx'
    if os.path.exists(xlsx):
        return os.path.getmtime(xlsx)
    return None

@st.cache_data(show_spinner=False)
def load_data(_mtime=None):
    # Prioridad al Excel de BIABLE
    if os.path.exists('HC_ajustado.xlsx'):
        df = pd.read_excel('HC_ajustado.xlsx')
        # Mapeo de BIABLE a nombres internos
        mapeo = {
            'Cedula Empleado': 'Cedula',
            'Grupo Ccostos': 'Centro de Costo'
        }
        df = df.rename(columns=mapeo)
    else:
        files = [f for f in os.listdir('.') if f.endswith('.csv') and 'Headcount' in f]
        if not files: return None
        df = pd.read_csv(files[0])
    
    def limpiar_texto(t):
        if not isinstance(t, str): return ""
        t = "".join(c for c in unicodedata.normalize('NFD', t) if unicodedata.category(c) != 'Mn')
        return re.sub(r'\s+', ' ', t).upper().strip()

    # Columnas a limpiar
    cols_limpiar = ['Nombre Empleado', 'Jefe', 'Responsable', 'Area', 'Cargo Empleado', 'Empresa', 'Municipio', 'Centro de Costo']
    for col in cols_limpiar:
        if col in df.columns:
            df[col] = df[col].apply(limpiar_texto)

    # Unificación Ana Maria Mosquera
    todos_nombres = pd.concat([df['Nombre Empleado'], df['Jefe'], df['Responsable']]).dropna().unique()
    mapa_unificacion = {}
    nombre_oficial_ana = "MOSQUERA GOMEZ ANA MARIA"
    for n in todos_nombres:
        huella = " ".join(sorted(n.split()))
        if "ANA" in n and "MARIA" in n and "MOSQUERA" in n:
            mapa_unificacion[huella] = nombre_oficial_ana
        else:
            if huella not in mapa_unificacion or len(n) > len(mapa_unificacion[huella]):
                mapa_unificacion[huella] = n
            
    for col in ['Nombre Empleado', 'Jefe', 'Responsable']:
        if col in df.columns:
            df[col] = df[col].apply(lambda x: mapa_unificacion.get(" ".join(sorted(str(x).split())), x) if x else x)

    # Estandarizar fechas
    for col in ['Fecha Ingreso', 'Fecha Retiro', 'Fecha Nacimiento']:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')

    # Filtro de activos
    if 'Fecha Retiro' in df.columns:
        df['Fecha Retiro_DT'] = df['Fecha Retiro'].fillna(pd.Timestamp('2099-12-31'))
        df = df[df['Fecha Retiro_DT'] >= pd.Timestamp.now().normalize()]
        
    return df

def render_recursivo(lider_nombre, df_full_global):
    lider_nombre = str(lider_nombre).upper().strip()
    df_equipo = df_full_global[df_full_global['Responsable'] == lider_nombre].copy()
    if df_equipo.empty: return
    
    # Pesos para ordenamiento
    pesos = {'JEFE':1, 'COORDINADOR':2, 'ANALISTA I':3, 'ANALISTA II':4, 'TECNICO':5, 'AUXILIAR':6}
    def get_peso(c):
        c_u = str(c).upper()
        for k,v in pesos.items():
            if k in c_u: return v
        return 10

    cargos = sorted(df_equipo['Cargo Empleado'].unique(), key=get_peso)
    for cargo in cargos:
        df_g = df_equipo[df_equipo['Cargo Empleado'] == cargo]
        st.markdown(f'<div class="group-card"><span class="count-badge">{len(df_g)}</span><b>{cargo}</b></div>', unsafe_allow_html=True)
        with st.expander(f"Ver {cargo}"):
            for _, emp in df_g.iterrows():
                n_emp, area_emp = emp['Nombre Empleado'], emp['Area']
                df_sub = df_full_global[df_full_global['Responsable'] == n_emp]
                tiene = len(df_sub) > 0
                c1, c2 = st.columns([3, 1])
                with c1:
                    l_tag = f'<span class="badge-leader">LÍDER ({len(df_sub)})</span>' if tiene else ""
                    st.markdown(f"**{n_emp}** {l_tag}<span class='area-badge'>{area_emp}</span>", unsafe_allow_html=True)
                with c2:
                    if tiene:
                        with st.expander("Ver Equipo"):
                            render_recursivo(n_emp, df_full_global)

# --- ORGANIGRAMA GRÁFICO COLLAPSIBLE D3 ---
def get_first_names(full_name):
    if not isinstance(full_name, str):
        return ""
    parts = full_name.strip().split()
    if len(parts) >= 3:
        return " ".join(parts[2:])
    elif len(parts) == 2:
        return parts[1]
    return full_name

def get_cargo_hierarchy_weight(cargo):
    c_u = str(cargo).upper().strip()
    if 'GERENTE GENERAL' in c_u:
        return 0
    elif 'DIRECCION CAMPO' in c_u or 'DIRECTOR CAMPO' in c_u or 'DIRECCION DE CAMPO' in c_u:
        return 1
    elif 'JEFE' in c_u:
        return 2
    elif 'COORDINADOR' in c_u:
        return 3
    elif 'TECNICO' in c_u:
        return 4
    elif 'ANALISTA I' in c_u:
        return 5
    elif 'ANALISTA II' in c_u:
        return 6
    elif 'ANALISTA' in c_u:
        return 6
    elif 'AUXILIAR DE CAMPO' in c_u:
        return 7
    elif 'AUXILIAR' in c_u or 'ASISTENTE' in c_u:
        return 7.5
    elif 'OFICIOS VARIOS' in c_u:
        return 8
    return 10

def get_section_for_report(row):
    cargo = str(row['Cargo Empleado']).upper()
    area = str(row['Area']).upper()
    
    if 'JEFE DE ZONA' in cargo or 'JEFE ZONA' in cargo:
        if 'ZONA 1' in area or 'ZONA 1' in cargo:
            return 'JEFE ZONA 1'
        elif 'ZONA 2' in area or 'ZONA 2' in cargo:
            return 'JEFE ZONA 2'
        elif 'ZONA 3' in area or 'ZONA 3' in cargo:
            return 'JEFE ZONA 3'
        return 'JEFE ZONA'
    elif 'JEFE' in cargo:
        return 'OTRAS JEFATURAS'
    elif 'COORDINADOR' in cargo:
        return 'COORDINACIONES'
    return 'OTROS COLABORADORES'

def adjust_hierarchy(df):
    df_copy = df.copy()
    managers = df_copy['Responsable'].dropna().unique()
    
    for mgr in managers:
        mgr_clean = str(mgr).strip().upper()
        if not mgr_clean or mgr_clean in ['', 'NAN', 'NONE', 'JEFE DE PROCESO']:
            continue
            
        reports_idx = df_copy[df_copy['Responsable'] == mgr_clean].index
        if len(reports_idx) == 0:
            continue
            
        leaders = []
        subordinates = []
        for idx in reports_idx:
            row = df_copy.loc[idx]
            cargo = str(row['Cargo Empleado']).upper()
            weight = get_cargo_hierarchy_weight(cargo)
            if weight < 4:
                leaders.append(row)
            else:
                subordinates.append((idx, row))
                
        if not leaders or not subordinates:
            continue
            
        leader_mappings = []
        for leader in leaders:
            l_name = str(leader['Nombre Empleado']).upper().strip()
            l_area = str(leader['Area']).upper().strip()
            l_cargo = str(leader['Cargo Empleado']).upper().strip()
            
            keys = set()
            for zone in ['ZONA 1', 'ZONA 2', 'ZONA 3']:
                if zone in l_area or zone in l_cargo:
                    keys.add(zone)
            
            words = [w for w in re.split(r'[\s\-_/]+', l_area) if len(w) > 3 and w not in ['DIRECCION', 'CAMPO', 'OPERACIONES']]
            for w in words:
                keys.add(w)
                
            cargo_words = [w for w in re.split(r'[\s\-_/]+', l_cargo) if len(w) > 3 and w not in ['LIDER', 'UNIDAD', 'JEFE', 'COORDINADOR', 'DIRECTOR']]
            for w in cargo_words:
                keys.add(w)
                
            leader_mappings.append((l_name, keys))
            
        for idx, sub in subordinates:
            sub_area = str(sub['Area']).upper().strip()
            sub_cargo = str(sub['Cargo Empleado']).upper().strip()
            
            matched_leader = None
            for l_name, keys in leader_mappings:
                for key in keys:
                    if key in sub_area or key in sub_cargo:
                        matched_leader = l_name
                        break
                if matched_leader:
                    break
            
            if matched_leader:
                df_copy.at[idx, 'Responsable'] = matched_leader
                if 'Jefe' in df_copy.columns:
                    df_copy.at[idx, 'Jefe'] = matched_leader
                    
    return df_copy

def calculate_hierarchy_counts(df):
    reports = {}
    for _, row in df.iterrows():
        emp = str(row['Nombre Empleado']).strip().upper()
        boss = str(row['Responsable']).strip().upper()
        if boss and boss not in ['', 'NAN', 'NONE', 'JEFE DE PROCESO']:
            reports.setdefault(boss, []).append(emp)
            
    total_counts = {}
    
    def count_recursive(name):
        if name in total_counts:
            return total_counts[name]
        direct = reports.get(name, [])
        total = len(direct)
        for d in direct:
            total += count_recursive(d)
        total_counts[name] = total
        return total
        
    for name in df['Nombre Empleado'].dropna().unique():
        name_clean = str(name).strip().upper()
        count_recursive(name_clean)
        
    return reports, total_counts

def build_json_tree(boss_name, df, reports, total_counts, max_depth=6, current_depth=0):
    boss_name_clean = str(boss_name).strip().upper()
    direct_names = reports.get(boss_name_clean, [])
    
    if not direct_names or current_depth >= max_depth:
        return []
        
    # Buscar el cargo del jefe actual
    boss_rows = df[df['Nombre Empleado'] == boss_name_clean]
    boss_cargo = str(boss_rows.iloc[0]['Cargo Empleado']).upper() if not boss_rows.empty else ''
    boss_weight = get_cargo_hierarchy_weight(boss_cargo)
    
    # Agrupamos por sección si el jefe es Director/Gerente (peso <= 1)
    should_group_by_section = (boss_weight <= 1 and len(direct_names) > 3)
    
    children = []
    
    if should_group_by_section:
        # Clasificar reportes en secciones
        sections = {}
        for name in direct_names:
            name_clean = str(name).strip().upper()
            emp_rows = df[df['Nombre Empleado'] == name_clean]
            row = emp_rows.iloc[0] if not emp_rows.empty else None
            
            if row is not None:
                sec_name = get_section_for_report(row)
                cargo = str(row['Cargo Empleado']).strip().upper()
                area = str(row['Area']).strip().upper()
            else:
                sec_name = 'OTROS COLABORADORES'
                cargo = 'DESCONOCIDO'
                area = 'DESCONOCIDO'
                
            has_reports = name_clean in reports and len(reports[name_clean]) > 0
            sections.setdefault(sec_name, []).append((name_clean, cargo, area, has_reports))
            
        # Crear los nodos virtuales de sección
        for sec_name, items in sections.items():
            sec_children = []
            leaders = [it for it in items if it[3]]
            leaves = [it for it in items if not it[3]]
            
            # Leaders
            for name, cargo, area, _ in leaders:
                sub_children = build_json_tree(name, df, reports, total_counts, max_depth, current_depth + 2)
                sec_children.append({
                    "name": name,
                    "cargo": cargo,
                    "area": area,
                    "is_leader": True,
                    "is_group": False,
                    "total_reports": total_counts.get(name, 0),
                    "direct_reports": len(reports.get(name, [])),
                    "children": sub_children
                })
                
            # Leaves grouped by Cargo
            leaves_by_cargo = {}
            for name, cargo, area, _ in leaves:
                if 'OFICIOS VARIOS' in cargo:
                    cargo_key = 'OFICIOS VARIOS CAMPO'
                elif 'SERVICIOS GENERALES' in cargo:
                    cargo_key = 'AUXILIAR SERVICIOS GENERALES'
                else:
                    cargo_key = cargo
                leaves_by_cargo.setdefault(cargo_key, []).append((name, area))
                
            for cargo, emps in leaves_by_cargo.items():
                if 'OFICIOS VARIOS' in cargo or 'SERVICIOS GENERALES' in cargo:
                    names = [get_first_names(e[0]) for e in emps]
                else:
                    names = [e[0] for e in emps]
                areas = list(set(e[1] for e in emps))
                sec_children.append({
                    "name": f"{cargo} ({len(emps)})",
                    "cargo": cargo,
                    "area": ", ".join(areas),
                    "is_leader": False,
                    "is_group": True,
                    "count": len(emps),
                    "names": names,
                    "children": []
                })
                
            sec_children.sort(key=lambda x: get_cargo_hierarchy_weight(x["cargo"]))
            
            # Calcular total de reportes en la sección
            sec_total = 0
            for child in sec_children:
                if child.get("is_leader"):
                    sec_total += child.get("total_reports", 0) + 1
                elif child.get("is_group"):
                    sec_total += child.get("count", 0)
                else:
                    sec_total += 1
                    
            children.append({
                "name": sec_name,
                "cargo": "SECCIÓN",
                "area": sec_name,
                "is_leader": False,
                "is_group": False,
                "is_section": True,
                "total_reports": sec_total,
                "children": sec_children
            })
            
        # Ordenar las secciones por importancia o alfabéticamente
        def get_section_weight(s_name):
            if 'JEFE ZONA 1' in s_name: return 1
            if 'JEFE ZONA 2' in s_name: return 2
            if 'JEFE ZONA 3' in s_name: return 3
            if 'OTRAS JEFATURAS' in s_name: return 4
            if 'COORDINACIONES' in s_name: return 5
            return 10
            
        children.sort(key=lambda x: get_section_weight(x["name"]))
        
    else:
        # Lógica estándar de construcción del árbol
        leaders = []
        leaves = []
        for name in direct_names:
            name_clean = str(name).strip().upper()
            emp_rows = df[df['Nombre Empleado'] == name_clean]
            if not emp_rows.empty:
                cargo = str(emp_rows.iloc[0]['Cargo Empleado']).strip().upper()
                area = str(emp_rows.iloc[0]['Area']).strip().upper()
            else:
                cargo = 'DESCONOCIDO'
                area = 'DESCONOCIDO'
                
            has_reports = name_clean in reports and len(reports[name_clean]) > 0
            if has_reports:
                leaders.append((name_clean, cargo, area))
            else:
                leaves.append((name_clean, cargo, area))
                
        for name, cargo, area in leaders:
            sub_children = build_json_tree(name, df, reports, total_counts, max_depth, current_depth + 1)
            children.append({
                "name": name,
                "cargo": cargo,
                "area": area,
                "is_leader": True,
                "is_group": False,
                "total_reports": total_counts.get(name, 0),
                "direct_reports": len(reports.get(name, [])),
                "children": sub_children
            })
            
        leaves_by_cargo = {}
        for name, cargo, area in leaves:
            if 'OFICIOS VARIOS' in cargo:
                cargo_key = 'OFICIOS VARIOS CAMPO'
            elif 'SERVICIOS GENERALES' in cargo:
                cargo_key = 'AUXILIAR SERVICIOS GENERALES'
            else:
                cargo_key = cargo
            leaves_by_cargo.setdefault(cargo_key, []).append((name, area))
            
        for cargo, emps in leaves_by_cargo.items():
            if 'OFICIOS VARIOS' in cargo or 'SERVICIOS GENERALES' in cargo:
                names = [get_first_names(e[0]) for e in emps]
            else:
                names = [e[0] for e in emps]
            areas = list(set(e[1] for e in emps))
            children.append({
                "name": f"{cargo} ({len(emps)})",
                "cargo": cargo,
                "area": ", ".join(areas),
                "is_leader": False,
                "is_group": True,
                "count": len(emps),
                "names": names,
                "children": []
            })
            
        children.sort(key=lambda x: get_cargo_hierarchy_weight(x["cargo"]))
        
    return children

def generate_d3_html(root_node):
    json_data = json.dumps(root_node, ensure_ascii=False)
    
    html_code = f"""
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="utf-8">
        <script src="https://d3js.org/d3.v7.min.js"></script>
        <style>
            body {{
                font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
                background-color: #f8fafc;
                margin: 0;
                padding: 0;
                overflow: hidden;
                width: 100vw;
                height: 100vh;
            }}
            .link {{
                fill: none;
                stroke: #cbd5e1;
                stroke-width: 2px;
            }}
            .node {{
                cursor: pointer;
            }}
            .controls {{
                position: absolute;
                top: 12px;
                left: 12px;
                z-index: 100;
                background: rgba(255, 255, 255, 0.95);
                padding: 8px 12px;
                border-radius: 8px;
                box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1), 0 2px 4px -2px rgb(0 0 0 / 0.1);
                display: flex;
                gap: 8px;
                border: 1px solid #e2e8f0;
                align-items: center;
            }}
            .btn {{
                background: #1e3a8a;
                color: white;
                border: none;
                padding: 6px 12px;
                border-radius: 6px;
                cursor: pointer;
                font-size: 11px;
                font-weight: 600;
                transition: all 0.2s;
                display: inline-flex;
                align-items: center;
                gap: 4px;
            }}
            .btn:hover {{
                background: #2563eb;
                transform: translateY(-1px);
            }}
            .btn-secondary {{
                background: #64748b;
            }}
            .btn-secondary:hover {{
                background: #475569;
            }}
            svg {{
                width: 100%;
                height: 100%;
                background-color: #f8fafc;
            }}
            .tooltip {{
                position: absolute;
                background: #0f172a;
                color: white;
                padding: 8px 12px;
                border-radius: 6px;
                font-size: 10px;
                pointer-events: none;
                opacity: 0;
                transition: opacity 0.2s;
                z-index: 1000;
                max-width: 250px;
                line-height: 1.4;
                box-shadow: 0 10px 15px -3px rgb(0 0 0 / 0.3);
            }}
        </style>
    </head>
    <body>
        <div class="controls">
            <button class="btn" id="btn-toggle-names">👤 Ocultar Nombres</button>
            <button class="btn btn-secondary" id="btn-reset">🔄 Reset Vista</button>
            <button class="btn btn-secondary" id="btn-expand">➕ Expandir Todo</button>
            <button class="btn btn-secondary" id="btn-collapse">➖ Contraer Todo</button>
            <button class="btn" style="background:#059669;" id="btn-download-svg">📥 Guardar SVG</button>
            <button class="btn" style="background:#059669;" id="btn-download-png">🖼️ Guardar PNG</button>
        </div>
        <div class="tooltip" id="tooltip"></div>
        <svg id="chart-svg"></svg>

        <script>
            const treeData = {json_data};
            
            let showNames = true;
            
            const svg = d3.select("#chart-svg");
            const width = window.innerWidth;
            const height = window.innerHeight;
            
            const rectWidth = 220;
            const rectHeight = 80;
            const siblingGap = 250;
            const levelGap = 150;
            
            const g = svg.append("g");
            
            const zoom = d3.zoom()
                .scaleExtent([0.1, 3])
                .on("zoom", (event) => {{
                    g.attr("transform", event.transform);
                }});
                
            svg.call(zoom);
            
            function resetView() {{
                svg.transition().duration(750).call(
                    zoom.transform,
                    d3.zoomIdentity.translate(width / 2, 50).scale(0.85)
                );
            }}
            
            const treemap = d3.tree().nodeSize([siblingGap, levelGap]);
            
            let root = d3.hierarchy(treeData, d => d.children);
            root.x0 = 0;
            root.y0 = 0;
            
            if (root.children) {{
                root.children.forEach(collapse);
            }}
            
            let i = 0;
            update(root);
            resetView();
            
            function collapse(d) {{
                if (d.children) {{
                    d._children = d.children;
                    d._children.forEach(collapse);
                    d.children = null;
                }}
            }}
            
            function expand(d) {{
                if (d._children) {{
                    d.children = d._children;
                    d._children = null;
                }}
                if (d.children) {{
                    d.children.forEach(expand);
                }}
            }}
            
            function update(source) {{
                const treeData = treemap(root);
                const nodes = treeData.descendants();
                const links = treeData.links();
                
                nodes.forEach(d => {{ d.y = d.depth * levelGap; }});
                
                const node = g.selectAll("g.node")
                    .data(nodes, d => d.id || (d.id = ++i));
                    
                const nodeEnter = node.enter().append("g")
                    .attr("class", "node")
                    .attr("transform", d => `translate(${{source.x0}},${{source.y0}})`)
                    .on("click", (event, d) => {{
                        if (d.data.is_group) return;
                        if (d.children) {{
                            d._children = d.children;
                            d.children = null;
                        }} else {{
                            d.children = d._children;
                            d._children = null;
                        }}
                        update(d);
                    }});
                    
                const fo = nodeEnter.append("foreignObject")
                    .attr("width", rectWidth)
                    .attr("height", rectHeight)
                    .attr("x", -rectWidth / 2)
                    .attr("y", -rectHeight / 2);
                    
                fo.on("mouseover", (event, d) => {{
                    if (d.data.names && d.data.names.length > 0) {{
                        const tt = d3.select("#tooltip");
                        let namesHtml = d.data.names.map(n => `• ${{n}}`).join("<br>");
                        if (d.data.names.length > 15) {{
                            namesHtml = d.data.names.slice(0, 15).map(n => `• ${{n}}`).join("<br>") + `<br>... y ${{d.data.names.length - 15}} más`;
                        }}
                        tt.html(`<strong>Colaboradores:</strong><br>${{namesHtml}}`)
                          .style("opacity", 0.95);
                    }}
                }})
                .on("mousemove", (event) => {{
                    d3.select("#tooltip")
                      .style("left", (event.pageX + 15) + "px")
                      .style("top", (event.pageY - 20) + "px");
                }})
                .on("mouseout", () => {{
                    d3.select("#tooltip").style("opacity", 0);
                }});
                
                const nodeUpdate = nodeEnter.merge(node);
                
                nodeUpdate.transition().duration(500)
                    .attr("transform", d => `translate(${{d.x}},${{d.y}})`);
                    
                nodeUpdate.select("foreignObject")
                    .html(d => {{
                        let nodeColor = '#3b82f6';
                        if (d.depth === 0) nodeColor = '#0f172a';
                        else if (d.data.is_section) nodeColor = '#10b981';
                        else if (d.data.is_group) nodeColor = '#0284c7';
                        else if (d.data.is_leader) nodeColor = '#1e3a8a';
                        
                        let headerText = d.data.cargo || 'CARGO';
                        if (d.data.is_section) headerText = 'SECCIÓN';
                        
                        let detailText = '';
                        if (d.data.is_group) {{
                            detailText = `(${{d.data.count}} personas)`;
                        }} else if (d.data.is_section) {{
                            detailText = d.data.area;
                        }} else {{
                            detailText = showNames ? d.data.name : '';
                        }}
                        
                        let reportText = '';
                        if (d.data.is_leader) {{
                            let indicator = d.children ? '▼' : (d._children ? '▶' : '•');
                            reportText = `${{indicator}} Equipo: ${{d.data.total_reports}} pers.`;
                        }} else if (d.data.is_section) {{
                            let indicator = d.children ? '▼' : (d._children ? '▶' : '•');
                            reportText = `${{indicator}} Total: ${{d.data.total_reports}} pers.`;
                        }} else if (d.data.is_group) {{
                            reportText = `Grupo de Cargo`;
                        }} else {{
                            reportText = `Colaborador`;
                        }}
                        
                        return `
                            <div style="
                                box-sizing: border-box;
                                width: 100%;
                                height: 100%;
                                padding: 6px;
                                background: white;
                                border-left: 5px solid ${{nodeColor}};
                                border-radius: 6px;
                                box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
                                display: flex;
                                flex-direction: column;
                                justify-content: space-between;
                                border: 1px solid #e2e8f0;
                                overflow: hidden;
                            ">
                                <div style="font-size: 10px; font-weight: 700; color: #1e293b; text-transform: uppercase; white-space: normal; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; overflow: hidden; text-overflow: ellipsis; line-height: 1.15; max-height: 24px;" title="${{headerText}}">
                                    ${{headerText}}
                                </div>
                                <div style="font-size: 9px; color: #475569; font-weight: 500; white-space: normal; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; overflow: hidden; text-overflow: ellipsis; line-height: 1.15; max-height: 22px;" title="${{detailText}}">
                                    ${{detailText}}
                                </div>
                                <div style="font-size: 8.5px; font-weight: 600; color: #2563eb; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; line-height: 1.1;">
                                    ${{reportText}}
                                </div>
                            </div>
                        `;
                    }});
                    
                const nodeExit = node.exit().transition().duration(500)
                    .attr("transform", d => `translate(${{source.x}},${{source.y}})`)
                    .remove();
                    
                const link = g.selectAll("path.link")
                    .data(links, d => d.target.id);
                    
                const linkEnter = link.enter().insert("path", "g")
                    .attr("class", "link")
                    .attr("d", d => {{
                        const o = {{ x: source.x0, y: source.y0 }};
                        return diagonal(o, o);
                    }});
                    
                const linkUpdate = linkEnter.merge(link);
                
                linkUpdate.transition().duration(500)
                    .attr("d", d => diagonal(d.source, d.target));
                    
                link.exit().transition().duration(500)
                    .attr("d", d => {{
                        const o = {{ x: source.x, y: source.y }};
                        return diagonal(o, o);
                    }})
                    .remove();
                    
                nodes.forEach(d => {{
                    d.x0 = d.x;
                    d.y0 = d.y;
                }});
            }}
            
            function diagonal(s, d) {{
                const sy = s.y + rectHeight / 2;
                const sx = s.x;
                const dy = d.y - rectHeight / 2;
                const dx = d.x;
                const midY = sy + (dy - sy) / 2;
                
                return `M ${{sx}} ${{sy}}
                        L ${{sx}} ${{midY}}
                        L ${{dx}} ${{midY}}
                        L ${{dx}} ${{dy}}`;
            }}
            
            document.getElementById("btn-toggle-names").addEventListener("click", () => {{
                showNames = !showNames;
                document.getElementById("btn-toggle-names").innerText = showNames ? "👤 Ocultar Nombres" : "👤 Mostrar Nombres";
                update(root);
            }});
            
            document.getElementById("btn-reset").addEventListener("click", resetView);
            
            document.getElementById("btn-expand").addEventListener("click", () => {{
                root.children.forEach(expand);
                update(root);
            }});
            
            document.getElementById("btn-collapse").addEventListener("click", () => {{
                root.children.forEach(collapse);
                update(root);
            }});
            
            // Download SVG
            document.getElementById("btn-download-svg").addEventListener("click", () => {{
                const bbox = g.node().getBBox();
                const originalTransform = g.attr("transform");
                g.attr("transform", `translate(${{-bbox.x + 20}}, ${{-bbox.y + 20}})`);
                
                const svgString = new XMLSerializer().serializeToString(document.getElementById("chart-svg"));
                g.attr("transform", originalTransform);
                
                const fileBlob = new Blob([
                    `<?xml version="1.0" encoding="utf-8"?>\\n` + 
                    `<svg xmlns="http://www.w3.org/2000/svg" width="${{bbox.width + 40}}" height="${{bbox.height + 40}}">\\n` +
                    `<style>\\n` +
                    `  .link {{ fill: none; stroke: #cbd5e1; stroke-width: 2px; }}\\n` +
                    `</style>\\n` +
                    svgString.substring(svgString.indexOf("<g"))
                ], {{type: "image/svg+xml;charset=utf-8"}});
                
                const url = URL.createObjectURL(fileBlob);
                const a = document.createElement("a");
                a.href = url;
                a.download = `organigrama_${{treeData.name.replace(/\\s+/g, "_")}}.svg`;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
            }});
            
            // Download PNG
            document.getElementById("btn-download-png").addEventListener("click", () => {{
                const bbox = g.node().getBBox();
                const originalTransform = g.attr("transform");
                g.attr("transform", `translate(${{-bbox.x + 20}}, ${{-bbox.y + 20}})`);
                
                const svgString = new XMLSerializer().serializeToString(document.getElementById("chart-svg"));
                g.attr("transform", originalTransform);
                
                const width = bbox.width + 40;
                const height = bbox.height + 40;
                
                const fullSvg = `<?xml version="1.0" encoding="utf-8"?>\\n` + 
                    `<svg xmlns="http://www.w3.org/2000/svg" width="${{width}}" height="${{height}}">\\n` +
                    `<style>\\n` +
                    `  .link {{ fill: none; stroke: #cbd5e1; stroke-width: 2px; }}\\n` +
                    `</style>\\n` +
                    svgString.substring(svgString.indexOf("<g")) + `</svg>`;
                    
                const blob = new Blob([fullSvg], {{type: "image/svg+xml;charset=utf-8"}});
                const url = URL.createObjectURL(blob);
                
                const img = new Image();
                img.onload = function() {{
                    const canvas = document.createElement("canvas");
                    canvas.width = width * 2;
                    canvas.height = height * 2;
                    const ctx = canvas.getContext("2d");
                    ctx.scale(2, 2);
                    ctx.fillStyle = "#f8fafc";
                    ctx.fillRect(0, 0, width, height);
                    ctx.drawImage(img, 0, 0);
                    
                    canvas.toBlob(function(pngBlob) {{
                        const downloadUrl = URL.createObjectURL(pngBlob);
                        const a = document.createElement("a");
                        a.href = downloadUrl;
                        a.download = `organigrama_${{treeData.name.replace(/\\s+/g, "_")}}.png`;
                        document.body.appendChild(a);
                        a.click();
                        document.body.removeChild(a);
                    }}, "image/png");
                }};
                img.src = url;
            }});
        </script>
    </body>
    </html>
    """
    return html_code

def create_org_chart(root_name, df_full, max_depth=6):
    df_adjusted = adjust_hierarchy(df_full)
    reports, total_counts = calculate_hierarchy_counts(df_adjusted)
    root_name_clean = str(root_name).strip().upper()
    
    root_rows = df_adjusted[df_adjusted['Nombre Empleado'] == root_name_clean]
    if not root_rows.empty:
        root_cargo = str(root_rows.iloc[0]['Cargo Empleado']).strip().upper()
        root_area = str(root_rows.iloc[0]['Area']).strip().upper()
    else:
        root_cargo = 'DIRECTOR'
        root_area = 'CAMPO'
        
    children = build_json_tree(root_name_clean, df_adjusted, reports, total_counts, max_depth=max_depth)
    
    root_node = {
        "name": root_name_clean,
        "cargo": root_cargo,
        "area": root_area,
        "is_leader": True,
        "is_group": False,
        "total_reports": total_counts.get(root_name_clean, 0),
        "direct_reports": len(reports.get(root_name_clean, [])),
        "children": children
    }
    
    return generate_d3_html(root_node)

# --- APP PRINCIPAL ---
df = load_data(_mtime=_get_excel_mtime())

if df is not None:
    st.markdown('<div class="main-header">Portal de Gestión de Talento</div>', unsafe_allow_html=True)
    
    id_patterns = ['CEDULA', 'IDENTIFIC', 'DOCUMENTO', 'ID', 'CC', 'IDENT']
    col_id = next((c for c in df.columns if any(p in c.upper() for p in id_patterns)), None)
    col_nac = next((c for c in df.columns if 'NACIMIENTO' in c.upper()), None)
    col_ing = next((c for c in df.columns if 'INGRESO' in c.upper()), None)
    col_cc = next((c for c in df.columns if 'COSTO' in c.upper()), None)

    tab1, tab2, tab3, tab4 = st.tabs(["📊 Estructura Organizacional", "🔍 Ficha del Colaborador", "📈 Estadísticas de Seguridad Social", "🌳 Organigrama Gráfico"])
    
    with tab1:
        if 'reset_key' not in st.session_state: st.session_state.reset_key = 0
        with st.container():
            lideres = sorted(list(set(df['Jefe'].dropna().unique()) | set(df['Responsable'].dropna().unique())))
            lideres_clean = [l for l in lideres if l not in ['', 'NAN', 'NONE', 'JEFE DE PROCESO']]
            c_s, c_c = st.columns([5, 1])
            with c_s:
                j_sel = st.selectbox("Seleccione el líder:", [""] + lideres_clean, index=0, key=f"sb_{st.session_state.reset_key}")
            with c_c:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("🗑️ Limpiar", key="btn_clear"):
                    st.session_state.reset_key += 1
                    st.rerun()
        if j_sel:
            j_clean = str(j_sel).strip().upper()
            df_tramo = df[df['Jefe'] == j_clean]
            if df_tramo.empty: df_tramo = df[df['Responsable'] == j_clean]
            st.markdown(f'<div class="leader-header"><h2 style="margin:0; color:white;">{j_clean}</h2><p style="margin:0; opacity: 0.9;">Reportes directos: {len(df_tramo)} personas</p></div>', unsafe_allow_html=True)
            render_recursivo(j_clean, df)

    with tab2:
        st.subheader("Búsqueda por Cédula")
        cedula_input = st.text_input("Ingrese la cédula del colaborador:", placeholder="Ej: 12345678")
        
        if cedula_input:
            ced_clean = re.sub(r'\D', '', str(cedula_input))
            if col_id:
                persona = df[df[col_id].astype(str).str.replace('.0', '', regex=False).str.contains(ced_clean)]
                if not persona.empty:
                    p = persona.iloc[0]
                    edad = calcular_edad_v3(p.get(col_nac))
                    antiguedad = calcular_tiempo_completo_v2(p.get(col_ing))
                    region = categorizar_region(p.get('Municipio', ''))
                    
                    with st.container(border=True):
                        st.header(f"👤 {p['Nombre Empleado']}")
                        
                        # Banner de información clave
                        st.info(f"📍 **Unidad Orgánica:** {p.get(col_cc, 'N/A')} - {p.get('Desc Ccostos', '')}")
                        
                        c1, c2, c3 = st.columns(3)
                        with c1:
                            st.write("**Identificación:**", p[col_id])
                            st.write("**Cargo:**", p['Cargo Empleado'])
                            st.write("**Centro de Costo (Cód):**", f"`{p.get(col_cc, 'N/A')}`")
                        with c2:
                            st.write("**Área:**", p['Area'])
                            st.write("**Empresa:**", p.get('Empresa', 'N/A'))
                        with c3:
                            st.write("**Líder Directo:**", p['Responsable'])
                            st.write("**Estado:**", "🟢 Activo")
                        
                        st.divider()
                        st.subheader("📈 Métricas de Tiempo")
                        m1, m2 = st.columns(2)
                        with m1:
                            st.metric("Edad Actual", f"{edad} años")
                            st.caption(f"Nacimiento: {format_fecha_v2(p.get(col_nac))}")
                        with m2:
                            st.metric("Antigüedad", antiguedad)
                            st.caption(f"Ingreso: {format_fecha_v2(p.get(col_ing))}")
                        
                        st.divider()
                        st.subheader("📍 Ubicación")
                        u1, u2 = st.columns(2)
                        with u1:
                            st.write("**Municipio:**", p.get('Municipio', 'N/A'))
                        with u2:
                            st.write("**Región Estratégica:**", f"📍 {region}")

                        st.divider()
                        st.subheader("🛡️ Seguridad Social y Prestaciones")
                        s1, s2, s3, s4 = st.columns(4)
                        
                        def get_val(col):
                            v = p.get(col)
                            return "No Registra" if pd.isna(v) else str(v)

                        with s1:
                            st.write("**EPS:**")
                            st.write(get_val('EPS'))
                        with s2:
                            st.write("**Pensiones:**")
                            st.write(get_val('Entidad Pension'))
                        with s3:
                            st.write("**Cesantías:**")
                            st.write(get_val('Cesantias'))
                        with s4:
                            st.write("**Riesgo ARL:**")
                            st.write(get_val('Tipo Riesgo ARL'))
                else:
                    st.error(f"No se encontró el colaborador con ID: {ced_clean}")
            else:
                st.error("No se detectó columna de cédula.")

    with tab3:
        st.subheader("📈 Estadísticas de Seguridad Social y Prestaciones")
        st.markdown("Distribución general del personal por EPS, Fondos de Pensión, Cesantías y Niveles de Riesgo ARL.")
        
        c1, c2 = st.columns(2)
        
        with c1:
            st.markdown("#### Distribución por EPS")
            if 'EPS' in df.columns:
                eps_df = df['EPS'].fillna('No Registra').value_counts().reset_index()
                eps_df.columns = ['Entidad', 'Cantidad']
                st.plotly_chart(px.pie(eps_df, values='Cantidad', names='Entidad', hole=0.4), use_container_width=True)
            else:
                st.info("No hay datos de EPS disponibles.")
            
            st.markdown("#### Distribución por Fondo de Cesantías")
            if 'Cesantias' in df.columns:
                ces_df = df['Cesantias'].fillna('No Registra').value_counts().reset_index()
                ces_df.columns = ['Entidad', 'Cantidad']
                st.plotly_chart(px.bar(ces_df, x='Entidad', y='Cantidad', text='Cantidad', color='Entidad'), use_container_width=True)
            else:
                st.info("No hay datos de Cesantías disponibles.")

        with c2:
            st.markdown("#### Distribución por Fondo de Pensión")
            if 'Entidad Pension' in df.columns:
                pen_df = df['Entidad Pension'].fillna('No Registra').value_counts().reset_index()
                pen_df.columns = ['Entidad', 'Cantidad']
                st.plotly_chart(px.bar(pen_df, x='Cantidad', y='Entidad', orientation='h', text='Cantidad', color='Entidad'), use_container_width=True)
            else:
                st.info("No hay datos de Pensión disponibles.")
            
            st.markdown("#### Nivel de Riesgo ARL")
            if 'Tipo Riesgo ARL' in df.columns:
                arl_df = df['Tipo Riesgo ARL'].fillna('No Registra').value_counts().reset_index()
                arl_df.columns = ['Nivel', 'Cantidad']
                st.plotly_chart(px.pie(arl_df, values='Cantidad', names='Nivel', hole=0.4), use_container_width=True)
            else:
                st.info("No hay datos de ARL disponibles.")

    with tab4:
        st.subheader("🌳 Organigrama Gráfico por Jerarquía de Cargos")
        st.markdown(
            "Selecciona un líder raíz para ver su estructura. "
            "Los colaboradores con el **mismo cargo** bajo el mismo responsable "
            "se agrupan en un solo nodo. Haz **zoom** y **arrastra** para explorar."
        )

        col_root, col_depth = st.columns([3, 1])
        with col_root:
            lideres_chart = sorted(list(
                set(df['Jefe'].dropna().unique()) | set(df['Responsable'].dropna().unique())
            ))
            lideres_chart = [l for l in lideres_chart if l not in ['', 'NAN', 'NONE', 'JEFE DE PROCESO']]
            root_chart = st.selectbox(
                "Seleccionar líder raíz:", [''] + lideres_chart, key="org_chart_root"
            )
        with col_depth:
            st.markdown("<br>", unsafe_allow_html=True)
            max_d = st.number_input("Niveles máx.:", min_value=1, max_value=10, value=6, step=1)

        if root_chart:
            with st.spinner("Generando organigrama..."):
                fig_org = create_org_chart(root_chart, df, max_depth=int(max_d))

            if fig_org:
                # Leyenda
                lg1, lg2, lg3, lg4 = st.columns(4)
                lg1.markdown(
                    '<div style="background:#0F172A;color:white;padding:6px 12px;'
                    'border-radius:6px;text-align:center;font-size:0.8rem">'
                    '🔲 Raíz / Director</div>', unsafe_allow_html=True)
                lg2.markdown(
                    '<div style="background:#10B981;color:white;padding:6px 12px;'
                    'border-radius:6px;text-align:center;font-size:0.8rem">'
                    '🔲 Sección / Zona</div>', unsafe_allow_html=True)
                lg3.markdown(
                    '<div style="background:#1E3A8A;color:white;padding:6px 12px;'
                    'border-radius:6px;text-align:center;font-size:0.8rem">'
                    '🔲 Líder Individual</div>', unsafe_allow_html=True)
                lg4.markdown(
                    '<div style="background:#0284C7;color:white;padding:6px 12px;'
                    'border-radius:6px;text-align:center;font-size:0.8rem">'
                    '🔲 Grupo de Cargo (N)</div>', unsafe_allow_html=True)

                components.html(fig_org, height=800, scrolling=True)
                st.caption(
                    "💡 Scroll para zoom · Arrastra para desplazarte · "
                    "Pasa el cursor sobre un nodo para ver todos los nombres del equipo."
                )
            else:
                st.info(f"'{root_chart}' no tiene subordinados directos registrados en los datos.")

    st.divider()
else:
    st.error("No se pudo cargar el archivo CSV.")
