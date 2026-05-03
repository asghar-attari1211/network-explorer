import streamlit as st
import pandas as pd
import folium
from streamlit_folium import st_folium
import os
import re

# --- Configuration ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def get_latest_file(pattern, directory=BASE_DIR):
    if not os.path.exists(directory):
        return None
    try:
        files = [f for f in os.listdir(directory) if re.search(pattern, f) and f.endswith('.xlsx')]
        if not files: return None
        files.sort(key=lambda x: os.path.getmtime(os.path.join(directory, x)), reverse=True)
        return os.path.join(directory, files[0])
    except Exception:
        return None

ROUTES_FILE = os.path.join(BASE_DIR, "Routes.xlsx")
DEPS_FILE = os.path.join(BASE_DIR, "Dependencies.xlsx")
VLAN_FILE = get_latest_file(r"Consolidated_VLAN_Report_")
MASTER_FILE = os.path.join(BASE_DIR, "VLAN List -- AI Seekho -- Sample.xlsx")

st.set_page_config(layout="wide", page_title="Network Explorer Dashboard")

# --- Custom CSS for Layout ---
st.markdown("""
    <style>
    [data-testid="stSidebar"] { width: 25vw !important; max-width: 25vw !important; }
    .site-label { font-weight: bold; color: black; text-shadow: 1px 1px white; white-space: nowrap; }
    </style>
""", unsafe_allow_html=True)

# --- Data Loading ---
@st.cache_data
def load_all_data():
    xl_master = pd.ExcelFile(MASTER_FILE)
    sites_df = pd.DataFrame()
    site_id_col = None
    lat_col = None
    long_col = None

    if "Sites" in xl_master.sheet_names:
        sites_df = pd.read_excel(xl_master, sheet_name="Sites")
        site_id_col = find_column_like(sites_df, ['Site ID', 'ID', 'SITE', 'Site Name'])
        lat_col = find_column_like(sites_df, ['Lat', 'Latitude'])
        long_col = find_column_like(sites_df, ['Long', 'Longitude', 'Lon'])
    
    if sites_df.empty or not (site_id_col and lat_col and long_col):
        for sn in xl_master.sheet_names:
            try:
                tmp = pd.read_excel(xl_master, sheet_name=sn, nrows=5)
                lcols = [str(c).lower() for c in tmp.columns]
                if any('lat' in c for c in lcols) and any('lon' in c or 'long' in c for c in lcols):
                    sites_df = pd.read_excel(xl_master, sheet_name=sn)
                    site_id_col = find_column_like(sites_df, ['Site ID', 'ID', 'SITE', 'Site Name'])
                    lat_col = find_column_like(sites_df, ['Lat', 'Latitude'])
                    long_col = find_column_like(sites_df, ['Long', 'Longitude', 'Lon'])
                    if site_id_col and lat_col and long_col:
                        break
            except Exception:
                continue

    if sites_df.empty or not (site_id_col and lat_col and long_col):
        raise ValueError("Could not find a suitable 'Sites' sheet with 'Site ID', 'Lat', and 'Long' columns.")

    coord_map = {}
    if not sites_df.empty:
        for _, r in sites_df.iterrows():
            sid = str(r.get(site_id_col, r.get('ID', ''))).strip().upper()
            if sid and pd.notna(r.get(lat_col)) and pd.notna(r.get(long_col)):
                coord_map[sid] = [float(r[lat_col]), float(r[long_col])]
    
    lr_df = pd.read_excel(xl_master, sheet_name="LR")
    for _, r in lr_df.iterrows():
        s1 = str(r.get('Site name S1', '')).strip().upper()
        if s1 and s1 not in coord_map and pd.notna(r.get('S1 Lat')) and pd.notna(r.get('S1 Long')):
            coord_map[s1] = [float(r['S1 Lat']), float(r['S1 Long'])]
            
        s2 = str(r.get('Site name S2', '')).strip().upper()
        if s2 and s2 not in coord_map and pd.notna(r.get('S2 Lat')) and pd.notna(r.get('S2 Long')):
            coord_map[s2] = [float(r['S2 Lat']), float(r['S2 Long'])]
    
    routes_df = pd.read_excel(ROUTES_FILE, sheet_name="Final Routes") if os.path.exists(ROUTES_FILE) else pd.DataFrame()
    vlan_df = pd.read_excel(VLAN_FILE) if VLAN_FILE and os.path.exists(VLAN_FILE) else pd.DataFrame()
    
    deps_data = {}
    if os.path.exists(DEPS_FILE):
        xl_deps = pd.ExcelFile(DEPS_FILE)
        for sn in xl_deps.sheet_names:
            deps_data[sn] = pd.read_excel(xl_deps, sheet_name=sn)
            
    ofn_sites = set(pd.read_excel(xl_master, sheet_name="OFN").iloc[:, 0].dropna().astype(str).str.upper())
    fttt_sites = set(pd.read_excel(xl_master, sheet_name="FTTT").iloc[:, 0].dropna().astype(str).str.upper())
    vsat_sites = set(pd.read_excel(xl_master, sheet_name="VSAT").iloc[:, 0].dropna().astype(str).str.upper())

    return coord_map, lr_df, routes_df, vlan_df, ofn_sites, fttt_sites, vsat_sites, deps_data

def find_column_like(df, keywords):
    if df.empty: return None
    for col in df.columns:
        lcol = str(col).lower()
        for kw in keywords:
            if kw.lower() in lcol:
                return col
    return None

try:
    coords, lr_df, routes_df, vlan_df, ofn_set, fttt_set, vsat_set, deps_data = load_all_data()
except Exception as e:
    st.error(f"Error loading files: {e}. Ensure you have run Route.py first.")
    st.stop()

# --- Sidebar logic ---
st.sidebar.title("📡 Network Explorer")

csv_data = "Site ID\n" + "\n".join(sorted(coords.keys()))
st.sidebar.download_button("📥 Export All Sites", data=csv_data, file_name="all_sites.csv", mime="text/csv")

main_menu = st.sidebar.selectbox("Main Menu", ["Search (Site/OFN)", "View (Route and Dependency)"])

sites_to_highlight = {}
links_to_highlight = set()
center_loc = [24.8607, 67.0011]
zoom_lvl = 11

if main_menu == "Search (Site/OFN)":
    search_cat = st.sidebar.selectbox("Sub-menu", ["Site", "OFN"])
    
    if search_cat == "Site":
        target_sites_input = st.sidebar.text_area("Site IDs", height=100, help="Multiple Site with Comma Separation").upper()
        st.sidebar.caption("Multiple Site with Comma Separation")
        
        st.sidebar.markdown("---")
        st.sidebar.subheader("Details")
        
        if target_sites_input:
            target_sites = [s.strip() for s in target_sites_input.split(',') if s.strip()]
            
            all_site_routes = []
            all_deps = []
            
            for target_site in target_sites:
                sites_to_highlight[target_site] = {'is_route': True, 'is_dep': False, 'show_label': True}
                if target_site in coords: 
                    center_loc = coords[target_site]
                    zoom_lvl = 14
                
                if not routes_df.empty:
                    site_routes = routes_df[routes_df['Site ID'] == target_site]
                    if not site_routes.empty:
                        drop_nodes = site_routes['Drop Node'].dropna().unique()
                        for dn in drop_nodes:
                            surrounding = routes_df[routes_df['Drop Node'] == dn]['Site ID'].unique()
                            for s_site in surrounding:
                                if s_site not in sites_to_highlight:
                                    sites_to_highlight[s_site] = {'is_route': False, 'is_dep': False, 'show_label': True}
                        all_site_routes.append(site_routes)

                if deps_data:
                    for sheet_name, df_dep in deps_data.items():
                        if not df_dep.empty and 'Site Name' in df_dep.columns and 'Dependent Sites' in df_dep.columns:
                            deps_for_site = df_dep[df_dep['Site Name'] == target_site]
                            if not deps_for_site.empty:
                                temp_dep = deps_for_site[['Site Name', 'Dependent Sites']].copy()
                                temp_dep['Source'] = sheet_name
                                all_deps.append(temp_dep)

            if all_site_routes:
                combined_routes = pd.concat(all_site_routes, ignore_index=True)
                st.sidebar.subheader("Routes Info")
                display_df = combined_routes[['Site ID', 'Service', 'VLAN', 'Drop Node', 'Status', 'Comment']].copy()
                route_paths = []
                for _, r in combined_routes.iterrows():
                    path = [r['Site ID']]
                    for i in range(1, 20):
                        hop = r.get(f'FE{i}')
                        if pd.notna(hop) and str(hop).strip(): path.append(str(hop).upper())
                        else: break
                    route_paths.append(" -> ".join(path))
                display_df['Route Path'] = route_paths
                st.sidebar.dataframe(display_df, use_container_width=True)
            else:
                st.sidebar.info("No route information available for selected site(s).")
                
            if all_deps:
                st.sidebar.subheader("Dependencies Info")
                combined_deps = pd.concat(all_deps, ignore_index=True)
                st.sidebar.dataframe(combined_deps, use_container_width=True)
            else:
                st.sidebar.info("No dependency information available for selected site(s).")

    elif search_cat == "OFN":
        target_ofn = st.sidebar.selectbox("OFN ID", options=[""] + sorted(list(ofn_set)))
        
        st.sidebar.markdown("---")
        st.sidebar.subheader("Details")
        
        if target_ofn and not routes_df.empty:
            dependent_sites_on_ofn = routes_df[routes_df['Drop Node'] == target_ofn]['Site ID'].unique()
            for site_id in dependent_sites_on_ofn:
                sites_to_highlight[site_id] = {'is_route': False, 'is_dep': True, 'show_label': True}
            
            if target_ofn in coords:
                sites_to_highlight[target_ofn] = {'is_route': True, 'is_dep': False, 'show_label': True}
                center_loc = coords[target_ofn]
                zoom_lvl = 12
            
            st.sidebar.subheader(f"Sites Dependent on OFN: {target_ofn}")
            if len(dependent_sites_on_ofn) > 0:
                st.sidebar.write(f"The following sites drop to {target_ofn}:")
                st.sidebar.dataframe(pd.DataFrame(dependent_sites_on_ofn, columns=['Site ID']), use_container_width=True)
            else:
                st.sidebar.info(f"No sites found dependent on OFN {target_ofn}.")

else: # View Mode
    show_route = st.sidebar.checkbox("Route View", value=True)
    show_dep = st.sidebar.checkbox("Dependency View")
    
    target_v_input = st.sidebar.text_area("Enter Site IDs (comma-separated)", height=100).upper()
    
    service_layers = []
    if show_route:
        st.sidebar.write("Service Layers:")
        col1, col2 = st.sidebar.columns(2)
        if col1.checkbox("2G", value=True): service_layers.append("2G")
        if col1.checkbox("3G"): service_layers.append("3G")
        if col2.checkbox("4G"): service_layers.append("4G")
        if col2.checkbox("5G"): service_layers.append("5G")

    st.sidebar.markdown("---")
    st.sidebar.subheader("Details")

    if target_v_input:
        target_vs = [s.strip() for s in target_v_input.split(',') if s.strip()]
        
        all_routes_view = []
        all_deps_view = []
        
        for target_v in target_vs:
            if target_v in coords:
                center_loc = coords[target_v]
                zoom_lvl = 12
            
            if show_route and service_layers and not routes_df.empty:
                site_routes = routes_df[routes_df['Site ID'] == target_v]
                if not site_routes.empty:
                    # Find Drop Node(s) and highlight surrounding sites
                    drop_nodes = site_routes['Drop Node'].dropna().unique()
                    for dn in drop_nodes:
                        surrounding = routes_df[routes_df['Drop Node'] == dn]['Site ID'].unique()
                        for s_site in surrounding:
                            if s_site not in sites_to_highlight:
                                sites_to_highlight[s_site] = {'is_route': False, 'is_dep': False, 'show_label': True}

                    for svc in service_layers:
                        rows = site_routes[site_routes['Service'] == svc]
                        if not rows.empty:
                            all_routes_view.append(rows)
                            for _, r in rows.iterrows():
                                path = [target_v]
                                for i in range(1, 20): # Max 19 hops
                                    hop = r.get(f'FE{i}')
                                    if pd.notna(hop) and str(hop).strip(): path.append(str(hop).upper())
                                    else: break
                                
                                # Mark sites in the path for highlighting and labeling
                                for site_id in path:
                                    sites_to_highlight[site_id] = sites_to_highlight.get(site_id, {'is_route': False, 'is_dep': False, 'show_label': False})
                                    sites_to_highlight[site_id]['is_route'] = True
                                    sites_to_highlight[site_id]['show_label'] = True 
                                
                                # Mark links in the path for highlighting
                                for j in range(len(path)-1):
                                    links_to_highlight.add((path[j], path[j+1], svc))

            if show_dep and deps_data:
                for sheet_name, df_dep in deps_data.items():
                    if not df_dep.empty and 'Site Name' in df_dep.columns and 'Dependent Sites' in df_dep.columns:
                        deps_for_site = df_dep[df_dep['Site Name'] == target_v]
                        if not deps_for_site.empty:
                            temp_dep = deps_for_site[['Site Name', 'Dependent Sites']].copy()
                            temp_dep['Source'] = sheet_name
                            all_deps_view.append(temp_dep)
                            # Mark dependent sites for highlighting and labeling
                            for _, r_dep in deps_for_site.iterrows():
                                dep_site = str(r_dep['Dependent Sites']).upper()
                                if dep_site in coords:
                                    sites_to_highlight[dep_site] = sites_to_highlight.get(dep_site, {'is_route': False, 'is_dep': False, 'show_label': False})
                                    sites_to_highlight[dep_site]['is_dep'] = True
                                    sites_to_highlight[dep_site]['show_label'] = True 
                                    
                                    # Also highlight the link from target_v to dep_site if it exists in LR
                                    if target_v in coords and dep_site in coords:
                                        if ((lr_df['Site name S1'].str.upper() == target_v) & (lr_df['Site name S2'].str.upper() == dep_site)).any() or \
                                           ((lr_df['Site name S1'].str.upper() == dep_site) & (lr_df['Site name S2'].str.upper() == target_v)).any():
                                            links_to_highlight.add((target_v, dep_site, 'DEP'))

        if show_route and service_layers:
            if all_routes_view:
                combined_routes = pd.concat(all_routes_view, ignore_index=True)
                st.sidebar.subheader("Routes Info")
                display_df = combined_routes[['Site ID', 'Service', 'VLAN', 'Drop Node', 'Status', 'Comment']].copy()
                route_paths = []
                for _, r in combined_routes.iterrows():
                    path = [r['Site ID']]
                    for i in range(1, 20):
                        hop = r.get(f'FE{i}')
                        if pd.notna(hop) and str(hop).strip(): path.append(str(hop).upper())
                        else: break
                    route_paths.append(" -> ".join(path))
                display_df['Route Path'] = route_paths
                st.sidebar.dataframe(display_df, use_container_width=True)
            else:
                st.sidebar.info("No routes found.")

        if show_dep:
            if all_deps_view:
                st.sidebar.subheader("Dependencies Info")
                combined_deps = pd.concat(all_deps_view, ignore_index=True)
                st.sidebar.dataframe(combined_deps, use_container_width=True)
            else:
                st.sidebar.info("No dependencies found.")

# --- Map Engine ---
m = folium.Map(location=center_loc, zoom_start=zoom_lvl, tiles="cartodbpositron")

def get_site_type(sid):
    if sid in ofn_set: return 'OFN'
    if sid in fttt_set: return 'FTTT'
    if sid in vsat_set: return 'VSAT'
    return 'MW'

def add_site_marker(map_obj, site_id, site_type, is_highlighted_route=False, is_highlighted_dependency=False, show_label=False, is_background_site=False):
    if site_id not in coords: return
    loc = coords[site_id]
    
    icon_html = ""
    radius = 4
    border_color = "black"
    fill_color = "gray"
    border_weight = 1
    
    if site_type == 'OFN':
        icon_shape = "triangle"
        radius = 12
        border_weight = 2
        fill_color = "blue"
    elif site_type == 'FTTT':
        icon_color = "pink"
        icon_shape = "circle"
        icon_html = f'<i class="fa fa-flag" style="color:{icon_color};"></i>'
        radius = 8
        border_weight = 2
        fill_color = "pink"
    elif site_type == 'VSAT':
        icon_color = "black"
        icon_shape = "circle"
        icon_html = f'<i class="fa fa-satellite" style="color:{icon_color};"></i>'
        radius = 8
        border_weight = 2
        fill_color = "black"
    else:
        pass

    if is_background_site and not (is_highlighted_route or is_highlighted_dependency):
        border_color = "gray"
        fill_color = "gray"
        radius = 4
        border_weight = 1

    if is_highlighted_route:
        border_color = "red"
        border_weight = 3
        radius = max(radius, 6)
        fill_color = "red"
        
    if is_highlighted_dependency and not is_highlighted_route:
        border_color = "orange"
        border_weight = 3
        radius = max(radius, 6)
        fill_color = "orange"

    if icon_html:
        folium.Marker(
            location=loc,
            icon=folium.DivIcon(
                icon_size=(24, 24),
                icon_anchor=(12, 12),
                html=f'<div style="font-size: 12pt;">{icon_html}</div>'
            ),
            popup=site_id
        ).add_to(map_obj)
    else:
        if site_type == "OFN":
            folium.RegularPolygonMarker(
                location=loc,
                number_of_sides=3,
                radius=radius,
                color=border_color,
                fill_color=fill_color,
                fill_opacity=0.7,
                weight=border_weight,
                popup=site_id
            ).add_to(map_obj)
        else:
            folium.CircleMarker(
                location=loc,
                radius=radius,
                color=border_color,
                fill_color=fill_color,
                fill_opacity=0.7,
                weight=border_weight,
                popup=site_id
            ).add_to(map_obj)

    if show_label:
        label_offset_lat = 0.0005
        label_offset_lon = 0 
        if radius > 8:
            label_offset_lat = 0.001

        folium.Marker(
            location=[loc[0] + label_offset_lat, loc[1] + label_offset_lon],
            icon=folium.DivIcon(
                icon_size=(150, 36),
                icon_anchor=(0, 0),
                html=f'<div class="site-label">{site_id}</div>'
            )
        ).add_to(map_obj)

all_unique_sites_on_map = set(coords.keys())
for site_id in sites_to_highlight.keys():
    all_unique_sites_on_map.add(site_id)
for s1, s2, _ in links_to_highlight:
    all_unique_sites_on_map.add(s1)
    all_unique_sites_on_map.add(s2)

for _, row in lr_df.head(1000).iterrows(): # Limit background links for performance
    s1, s2 = str(row['Site name S1']).upper(), str(row['Site name S2']).upper()
    if s1 in coords and s2 in coords:
        if (s1, s2, 'LR') not in links_to_highlight and (s2, s1, 'LR') not in links_to_highlight and \
           (s1, s2, 'DEP') not in links_to_highlight and (s2, s1, 'DEP') not in links_to_highlight:
            folium.PolyLine([coords[s1], coords[s2]], color="blue", weight=1, opacity=0.2).add_to(m)

drawn = 0
for site_id in sites_to_highlight.keys():
    site_info = sites_to_highlight[site_id]
    add_site_marker(m, site_id, get_site_type(site_id), 
                    is_highlighted_route=site_info['is_route'], 
                    is_highlighted_dependency=site_info['is_dep'], 
                    show_label=site_info['show_label'],
                    is_background_site=False)
    drawn += 1

for site_id in all_unique_sites_on_map:
    if site_id not in sites_to_highlight:
        add_site_marker(m, site_id, get_site_type(site_id), 
                        is_highlighted_route=False, 
                        is_highlighted_dependency=False, 
                        show_label=False,
                        is_background_site=True)
        drawn += 1
        if drawn >= 1000: break

for s1, s2, link_type in links_to_highlight:
    if s1 in coords and s2 in coords:
        color = "red"
        if link_type == 'DEP': color = "purple"
        elif link_type == 'LR': color = "green"
        folium.PolyLine([coords[s1], coords[s2]], color=color, weight=5, opacity=0.8).add_to(m)

st.title("Interactive Network Map")
st_folium(m, width="100%", height=700)