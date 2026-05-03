import streamlit as st
import pandas as pd
import folium
from streamlit_folium import st_folium
import os
import re
from datetime import datetime

# --- Configuration ---
BASE_DIR = "/home/asghar_attari1211/RnD"

def get_latest_file(pattern):
    files = [f for f in os.listdir(BASE_DIR) if re.search(pattern, f) and f.endswith('.xlsx')]
    if not files: return None
    files.sort(key=lambda x: os.path.getmtime(os.path.join(BASE_DIR, x)), reverse=True)
    return os.path.join(BASE_DIR, files[0])

ROUTES_FILE = os.path.join(BASE_DIR, "Routes.xlsx")
DEPS_FILE = os.path.join(BASE_DIR, "Dependencies.xlsx")
VLAN_FILE = get_latest_file(r"Consolidated_VLAN_Report_")
MASTER_FILE = os.path.join(BASE_DIR, "VLAN List -- AI Seekho -- Sample.xlsx")

st.set_page_config(layout="wide", page_title="Network Explorer Dashboard")

# --- Custom CSS for Layout and Blinking ---
st.markdown("""
    <style>
    [data-testid="stSidebar"] { width: 20vw !important; max-width: 20vw !important; }
    @keyframes blinker {
        0% { opacity: 1.0; stroke-width: 8; }
        50% { opacity: 0.1; stroke-width: 2; }
        100% { opacity: 1.0; stroke-width: 8; }
    }
    .blink-highlight { animation: blinker 0.6s linear 3; }
    </style>
""", unsafe_allow_html=True)

# --- Data Loading ---
@st.cache_data
def load_all_data():
    # Load Coordinates
    xl_master = pd.ExcelFile(MASTER_FILE)

    # --- Flexible sheet detection for "Sites" data ---
    sites_df = pd.DataFrame()
    site_id_col = None
    lat_col = None
    long_col = None

    # First, try to find a sheet explicitly named "Sites"
    if "Sites" in xl_master.sheet_names:
        sites_df = pd.read_excel(xl_master, sheet_name="Sites")
        site_id_col = find_column_like(sites_df, ['Site ID', 'ID', 'SITE', 'Site Name'])
        lat_col = find_column_like(sites_df, ['Lat', 'Latitude'])
        long_col = find_column_like(sites_df, ['Long', 'Longitude', 'Lon'])
    
    # If "Sites" sheet not found or essential columns are missing, try other sheets
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
                        break # Found a suitable sheet
            except Exception:
                continue

    if sites_df.empty or not (site_id_col and lat_col and long_col):
        raise ValueError("Could not find a suitable 'Sites' sheet with 'Site ID', 'Lat', and 'Long' columns in the master Excel file.")

    coord_map = {}
    if not sites_df.empty:
        for _, r in sites_df.iterrows():
            sid = str(r.get(site_id_col, r.get('ID', ''))).strip().upper()
            if sid and pd.notna(r.get(lat_col)) and pd.notna(r.get(long_col)):
                coord_map[sid] = [float(r[lat_col]), float(r[long_col])]
    
    # Load Link Relations
    lr_df = pd.read_excel(xl_master, sheet_name="LR")
    
    # Scrape coordinates from LR sheet for sites not in master Sites list
    for _, r in lr_df.iterrows():
        s1 = str(r.get('Site name S1', '')).strip().upper()
        if s1 and s1 not in coord_map and pd.notna(r.get('S1 Lat')) and pd.notna(r.get('S1 Long')):
            coord_map[s1] = [float(r['S1 Lat']), float(r['S1 Long'])]
            
        s2 = str(r.get('Site name S2', '')).strip().upper()
        if s2 and s2 not in coord_map and pd.notna(r.get('S2 Lat')) and pd.notna(r.get('S2 Long')):
            coord_map[s2] = [float(r['S2 Lat']), float(r['S2 Long'])]
    
    # Load Generated Outputs
    routes_df = pd.read_excel(ROUTES_FILE, sheet_name="Final Routes") if ROUTES_FILE else pd.DataFrame()
    vlan_df = pd.read_excel(VLAN_FILE) if VLAN_FILE else pd.DataFrame()
    
    # Legend Sets
    ofn_sites = set(pd.read_excel(xl_master, sheet_name="OFN").iloc[:, 0].dropna().astype(str).str.upper())
    fttt_sites = set(pd.read_excel(xl_master, sheet_name="FTTT").iloc[:, 0].dropna().astype(str).str.upper())
    vsat_sites = set(pd.read_excel(xl_master, sheet_name="VSAT").iloc[:, 0].dropna().astype(str).str.upper())

    return coord_map, lr_df, routes_df, vlan_df, ofn_sites, fttt_sites, vsat_sites

# Helper function to find columns (copied from Route.py)
def find_column_like(df, keywords):
    """Find first column in df whose name contains any of the keywords (case-insensitive).
    Returns the original column name or None if not found."""
    if df.empty: return None
    for col in df.columns:
        lcol = str(col).lower()
        for kw in keywords:
            if kw.lower() in lcol:
                return col
    
    return None

try:
    coords, lr_df, routes_df, vlan_df, ofn_set, fttt_set, vsat_set = load_all_data()
except Exception as e:
    st.error(f"Error loading files: {e}. Ensure you have run Route.py first.")
    st.stop()

# --- Sidebar logic ---
st.sidebar.title("📡 Network Explorer")
main_menu = st.sidebar.selectbox("Main Menu", ["Search (Site/Link/OFN/Group)", "View (Route and Dependency)"])

highlight_sites = []
highlight_links = []
center_loc = [24.8607, 67.0011] # Default center
zoom_lvl = 11

if main_menu == "Search (Site/Link/OFN/Group)":
    search_cat = st.sidebar.selectbox("Sub-menu", ["Site", "Link", "OFN", "Area/Cluster"])
    
    if search_cat == "Site":
        target_site = st.sidebar.selectbox("Site ID", options=[""] + sorted(list(coords.keys())))
        if target_site:
            highlight_sites = [target_site]
            center_loc = coords[target_site]
            zoom_lvl = 15
            
    elif search_cat == "Link":
        search_str = st.sidebar.text_input("Site name for Link filtering").upper()
        relevant_lr = lr_df[(lr_df['Site name S1'].str.contains(search_str, na=False)) | 
                            (lr_df['Site name S2'].str.contains(search_str, na=False))]
        
        link_choice = st.sidebar.selectbox("Select Mod A-B", ["All Links"] + list(relevant_lr['Mod A-B'].unique()))
        if search_str:
            if link_choice == "All Links":
                highlight_links = relevant_lr.to_dict('records')
            else:
                highlight_links = relevant_lr[relevant_lr['Mod A-B'] == link_choice].to_dict('records')
            
            if highlight_links:
                s1 = str(highlight_links[0]['Site name S1']).upper()
                if s1 in coords: center_loc, zoom_lvl = coords[s1], 13

    elif search_cat == "OFN":
        target_ofn = st.sidebar.selectbox("OFN ID", options=[""] + sorted(list(ofn_set)))
        if target_ofn:
            highlight_sites = list(routes_df[routes_df['Drop Node'] == target_ofn]['Site ID'].unique())
            if target_ofn in coords: center_loc, zoom_lvl = coords[target_ofn], 12

    elif search_cat == "Area/Cluster":
        if 'UNMS Group Structure' in vlan_df.columns:
            group = st.sidebar.selectbox("Cluster Group", options=[""] + sorted(list(vlan_df['UNMS Group Structure'].dropna().unique())))
            if group:
                highlight_sites = list(vlan_df[vlan_df['UNMS Group Structure'] == group]['NE NAME'].dropna().unique())
                for s in highlight_sites:
                    if s in coords: center_loc, zoom_lvl = coords[s], 11; break

else: # View Mode
    show_route = st.sidebar.checkbox("Route View", value=True)
    show_dep = st.sidebar.checkbox("Dependency View")
    
    target_v = st.sidebar.text_input("Enter Site ID").upper()
    
    if show_route:
        layers = []
        st.sidebar.write("Service Layers:")
        col1, col2 = st.sidebar.columns(2)
        if col1.checkbox("2G"): layers.append("2G")
        if col1.checkbox("3G"): layers.append("3G")
        if col2.checkbox("4G"): layers.append("4G")
        if col2.checkbox("5G"): layers.append("5G")
        
        if target_v and layers:
            for l in layers:
                rows = routes_df[(routes_df['Site ID'] == target_v) & (routes_df['Service'] == l)]
                for _, r in rows.iterrows():
                    path = [target_v]
                    for i in range(1, 20):
                        hop = r.get(f'FE{i}')
                        if pd.notna(hop) and str(hop).strip(): path.append(str(hop).upper())
                        else: break
                    highlight_sites.extend(path)
                    for j in range(len(path)-1):
                        highlight_links.append({'Site name S1': path[j], 'Site name S2': path[j+1]})

    if show_dep and target_v:
        try:
            xl_dep = pd.ExcelFile(DEPS_FILE)
            for sn in xl_dep.sheet_names:
                df_d = pd.read_excel(xl_dep, sheet_name=sn)
                found = df_d[df_d['Site Name'] == target_v]['Dependent Sites'].unique()
                highlight_sites.extend(list(found))
        except: pass

    if target_v in coords: center_loc = coords[target_v]

# --- Map Engine ---
m = folium.Map(location=center_loc, zoom_start=zoom_lvl, tiles="cartodbpositron")

def add_site_marker(sid, is_highlight=False):
    if sid not in coords: return
    loc = coords[sid]
    p = f"Site: {sid}"
    cname = "blink-highlight" if is_highlight else ""
    
    if sid in ofn_set:
        folium.RegularPolygonMarker(loc, number_of_sides=3, radius=12, color="blue", fill=True, popup=p, class_name=cname).add_to(m)
    elif sid in fttt_set:
        folium.Marker(loc, icon=folium.Icon(color='pink', icon='flag', prefix='fa'), popup=p).add_to(m)
        if is_highlight: folium.CircleMarker(loc, radius=15, color="red", class_name=cname).add_to(m)
    elif sid in vsat_set:
        folium.Marker(loc, icon=folium.Icon(color='black', icon='satellite', prefix='fa'), popup=p).add_to(m)
        if is_highlight: folium.CircleMarker(loc, radius=15, color="red", class_name=cname).add_to(m)
    else:
        folium.CircleMarker(loc, radius=4, color="gray", fill=True, popup=p, class_name=cname).add_to(m)

# 1. Draw Background Links (faded)
for _, row in lr_df.head(500).iterrows(): # Performance cap
    s1, s2 = str(row['Site name S1']).upper(), str(row['Site name S2']).upper()
    if s1 in coords and s2 in coords:
        folium.PolyLine([coords[s1], coords[s2]], color="blue", weight=1, opacity=0.2).add_to(m)

# 2. Draw Highlighted Links
for link in highlight_links:
    s1, s2 = str(link.get('Site name S1', link.get('Near End'))).upper(), str(link.get('Site name S2', link.get('Far End'))).upper()
    if s1 in coords and s2 in coords:
        folium.PolyLine([coords[s1], coords[s2]], color="red", weight=5, opacity=0.8, class_name="blink-highlight").add_to(m)

# 3. Draw Markers
if highlight_sites:
    for s in highlight_sites: add_site_marker(s, is_highlight=True)
else:
    # Standard view: show limited sites for performance unless searched
    for s in list(coords.keys())[:1000]: add_site_marker(s)

st.title("Interactive Network Map")
st_folium(m, width="100%", height=700)