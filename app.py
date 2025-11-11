"""
Excel Duplicate Finder - Clean & Intuitive Review Interface
Streamlined grouping with direct visibility and efficient rendering
"""

import streamlit as st
import pandas as pd
import numpy as np
import hashlib
from rapidfuzz import fuzz
from io import BytesIO
from typing import Dict, List, Tuple, Set
import time
from collections import defaultdict
import gc

# Try optional performance libraries
try:
    import modin.pandas as mpd
    MODIN_AVAILABLE = True
except ImportError:
    MODIN_AVAILABLE = False

try:
    import polars as pl
    POLARS_AVAILABLE = True
except ImportError:
    POLARS_AVAILABLE = False

# Configure Streamlit page
st.set_page_config(
    page_title="Excel Duplicate Finder",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Minimal, clean CSS
st.markdown("""
<style>
    .main-header {font-size: 2.2rem; font-weight: 600; color: #1a73e8; margin-bottom: 0.3rem;}
    .sub-header {font-size: 1rem; color: #5f6368; margin-bottom: 1.5rem;}
    
    .group-section {
        background: #f8f9fa;
        border-radius: 8px;
        padding: 1.2rem;
        margin-bottom: 1.5rem;
    }
    
    .group-title {
        font-size: 1.1rem;
        font-weight: 600;
        color: #202124;
        margin-bottom: 0.8rem;
        padding-bottom: 0.5rem;
        border-bottom: 2px solid #e8eaed;
    }
    
    .dup-row {
        background: white;
        border-radius: 6px;
        padding: 0.8rem;
        margin-bottom: 0.6rem;
        border: 1px solid #e8eaed;
        transition: all 0.15s;
    }
    
    .dup-row:hover {
        border-color: #1a73e8;
        box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    }
    
    .dup-row-selected {
        border-color: #d93025;
        background: #fce8e6;
    }
    
    .badge {
        display: inline-block;
        padding: 0.2rem 0.6rem;
        border-radius: 4px;
        font-size: 0.8rem;
        font-weight: 500;
        margin-right: 0.5rem;
    }
    
    .badge-exact {background: #e6f4ea; color: #137333;}
    .badge-fuzzy {background: #fef7e0; color: #b06000;}
    
    .meta-text {
        color: #5f6368;
        font-size: 0.85rem;
        line-height: 1.4;
    }
    
    /* Clean up Streamlit defaults */
    .element-container {margin-bottom: 0 !important;}
    div[data-testid="stHorizontalBlock"] {gap: 0.5rem;}
    div[data-testid="column"] {padding: 0.2rem;}
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'workbooks_data' not in st.session_state:
    st.session_state.workbooks_data = {}
if 'duplicate_groups' not in st.session_state:
    st.session_state.duplicate_groups = {}
if 'selected_for_removal' not in st.session_state:
    st.session_state.selected_for_removal = set()
if 'preview_mode' not in st.session_state:
    st.session_state.preview_mode = True

# ============================================================================
# CORE FUNCTIONS
# ============================================================================

def normalize_value(val) -> str:
    """Normalize a cell value for comparison"""
    if pd.isna(val):
        return "__NAN__"
    if isinstance(val, (int, float, np.integer, np.floating)):
        return f"NUM:{float(val)}"
    if isinstance(val, (pd.Timestamp, np.datetime64)):
        return f"DATE:{str(val)}"
    if isinstance(val, (bool, np.bool_)):
        return f"BOOL:{bool(val)}"
    return f"STR:{str(val).strip().lower()}"

def row_to_hash(row: pd.Series) -> str:
    """Create a stable MD5 hash for a row"""
    normalized = [normalize_value(val) for val in row.values]
    row_str = "||".join(normalized)
    return hashlib.md5(row_str.encode('utf-8')).hexdigest()

def is_empty_row(row: pd.Series) -> bool:
    """Check if a row is completely empty"""
    return row.isna().all() or (row.astype(str).str.strip() == '').all()

@st.cache_data(show_spinner=False, ttl=3600)
def load_excel_files(uploaded_files, preview_mode: bool = True, preview_rows: int = 10000) -> Dict:
    """Load multiple Excel files and extract all sheets"""
    workbooks = {}
    
    for file in uploaded_files:
        try:
            file_bytes = file.read()
            file.seek(0)
            
            xls = pd.ExcelFile(BytesIO(file_bytes))
            sheets = {}
            
            for sheet_name in xls.sheet_names:
                if preview_mode:
                    df = xls.parse(sheet_name, nrows=preview_rows)
                else:
                    df = xls.parse(sheet_name)
                
                if not df.empty:
                    sheets[sheet_name] = df
            
            if sheets:
                workbooks[file.name] = sheets
            
            gc.collect()
                
        except Exception as e:
            st.error(f"Error loading {file.name}: {str(e)}")
            continue
    
    return workbooks

def detect_exact_duplicates(workbooks_data: Dict) -> Dict[str, List[Dict]]:
    """Detect exact duplicates using optimized hashing"""
    duplicate_groups = defaultdict(list)
    hash_index = defaultdict(list)
    
    for workbook, sheets in workbooks_data.items():
        for sheet, df in sheets.items():
            for idx in range(len(df)):
                if not is_empty_row(df.iloc[idx]):
                    row_hash = row_to_hash(df.iloc[idx])
                    hash_index[row_hash].append({
                        'workbook': workbook,
                        'sheet': sheet,
                        'row_index': idx,
                        'excel_row': idx + 2,
                        'data': df.iloc[idx].to_dict()
                    })
    
    for row_hash, locations in hash_index.items():
        if len(locations) > 1:
            unique_sources = set((loc['workbook'], loc['sheet']) for loc in locations)
            
            if len(unique_sources) == 1:
                workbook = locations[0]['workbook']
                sheet = locations[0]['sheet']
                group_key = f"within_{workbook}_{sheet}"
                group_label = f'Within "{sheet}" (Workbook: {workbook})'
                
                for loc in locations[1:]:
                    duplicate_groups[group_key].append({
                        'group_label': group_label,
                        'duplicate_type': 'Exact',
                        'similarity': 100.0,
                        'workbook': loc['workbook'],
                        'sheet': loc['sheet'],
                        'row_index': loc['row_index'],
                        'excel_row': loc['excel_row'],
                        'data': loc['data'],
                        'original_location': f"{locations[0]['workbook']}|{locations[0]['sheet']}|Row {locations[0]['excel_row']}"
                    })
            else:
                first_loc = locations[0]
                for loc in locations[1:]:
                    if first_loc['workbook'] == loc['workbook']:
                        group_key = f"cross_{first_loc['workbook']}_{first_loc['sheet']}_{loc['sheet']}"
                        group_label = f'Between "{first_loc["sheet"]}" and "{loc["sheet"]}" (Workbook: {first_loc["workbook"]})'
                    else:
                        group_key = f"cross_{first_loc['workbook']}_{first_loc['sheet']}_{loc['workbook']}_{loc['sheet']}"
                        group_label = f'Between "{first_loc["sheet"]}" ({first_loc["workbook"]}) and "{loc["sheet"]}" ({loc["workbook"]})'
                    
                    duplicate_groups[group_key].append({
                        'group_label': group_label,
                        'duplicate_type': 'Exact',
                        'similarity': 100.0,
                        'workbook': loc['workbook'],
                        'sheet': loc['sheet'],
                        'row_index': loc['row_index'],
                        'excel_row': loc['excel_row'],
                        'data': loc['data'],
                        'original_location': f"{first_loc['workbook']}|{first_loc['sheet']}|Row {first_loc['excel_row']}"
                    })
    
    return dict(duplicate_groups)

def detect_fuzzy_duplicates(workbooks_data: Dict, threshold: int = 90, max_comparisons: int = 50000) -> Dict[str, List[Dict]]:
    """Detect fuzzy duplicates using rapidfuzz"""
    duplicate_groups = defaultdict(list)
    all_rows = []
    
    for workbook, sheets in workbooks_data.items():
        for sheet, df in sheets.items():
            for idx in range(len(df)):
                if not is_empty_row(df.iloc[idx]):
                    row_str = " ".join([str(val) for val in df.iloc[idx].values if not pd.isna(val)])
                    all_rows.append({
                        'workbook': workbook,
                        'sheet': sheet,
                        'row_index': idx,
                        'excel_row': idx + 2,
                        'row_string': row_str.lower().strip(),
                        'data': df.iloc[idx].to_dict()
                    })
    
    total_possible_comparisons = len(all_rows) * (len(all_rows) - 1) // 2
    
    if total_possible_comparisons > max_comparisons:
        st.warning(f"⚠️ Large dataset: limiting to {max_comparisons:,} comparisons for performance")
        import random
        sample_size = int(np.sqrt(max_comparisons * 2))
        if sample_size < len(all_rows):
            all_rows = random.sample(all_rows, min(sample_size, len(all_rows)))
    
    compared_pairs = set()
    comparison_count = 0
    
    for i in range(len(all_rows)):
        if comparison_count >= max_comparisons:
            break
            
        for j in range(i + 1, len(all_rows)):
            if comparison_count >= max_comparisons:
                break
            
            if (all_rows[i]['workbook'] == all_rows[j]['workbook'] and 
                all_rows[i]['sheet'] == all_rows[j]['sheet'] and 
                all_rows[i]['row_index'] == all_rows[j]['row_index']):
                continue
            
            len_diff = abs(len(all_rows[i]['row_string']) - len(all_rows[j]['row_string']))
            max_len = max(len(all_rows[i]['row_string']), len(all_rows[j]['row_string']))
            if max_len > 0 and (len_diff / max_len) > (100 - threshold) / 100:
                continue
            
            similarity = fuzz.ratio(all_rows[i]['row_string'], all_rows[j]['row_string'])
            comparison_count += 1
            
            if similarity >= threshold:
                pair_key = tuple(sorted([
                    (all_rows[i]['workbook'], all_rows[i]['sheet'], all_rows[i]['row_index']),
                    (all_rows[j]['workbook'], all_rows[j]['sheet'], all_rows[j]['row_index'])
                ]))
                
                if pair_key not in compared_pairs:
                    compared_pairs.add(pair_key)
                    
                    if all_rows[i]['workbook'] == all_rows[j]['workbook']:
                        if all_rows[i]['sheet'] == all_rows[j]['sheet']:
                            group_key = f"fuzzy_within_{all_rows[i]['workbook']}_{all_rows[i]['sheet']}"
                            group_label = f'Fuzzy within "{all_rows[i]["sheet"]}" (Workbook: {all_rows[i]["workbook"]})'
                        else:
                            group_key = f"fuzzy_cross_{all_rows[i]['workbook']}_{all_rows[i]['sheet']}_{all_rows[j]['sheet']}"
                            group_label = f'Fuzzy between "{all_rows[i]["sheet"]}" and "{all_rows[j]["sheet"]}" (Workbook: {all_rows[i]["workbook"]})'
                    else:
                        group_key = f"fuzzy_cross_{all_rows[i]['workbook']}_{all_rows[i]['sheet']}_{all_rows[j]['workbook']}_{all_rows[j]['sheet']}"
                        group_label = f'Fuzzy between "{all_rows[i]["sheet"]}" ({all_rows[i]["workbook"]}) and "{all_rows[j]["sheet"]}" ({all_rows[j]["workbook"]})'
                    
                    duplicate_groups[group_key].append({
                        'group_label': group_label,
                        'duplicate_type': 'Fuzzy',
                        'similarity': round(similarity, 2),
                        'workbook': all_rows[j]['workbook'],
                        'sheet': all_rows[j]['sheet'],
                        'row_index': all_rows[j]['row_index'],
                        'excel_row': all_rows[j]['excel_row'],
                        'data': all_rows[j]['data'],
                        'original_location': f"{all_rows[i]['workbook']}|{all_rows[i]['sheet']}|Row {all_rows[i]['excel_row']}"
                    })
    
    return dict(duplicate_groups)

def clean_selected_rows(workbooks_data: Dict, selected_ids: Set[str]) -> Dict:
    """Remove selected duplicate rows from the data"""
    cleaned = {}
    
    for workbook, sheets in workbooks_data.items():
        cleaned[workbook] = {}
        for sheet, df in sheets.items():
            rows_to_remove = set()
            for unique_id in selected_ids:
                parts = unique_id.split('|')
                if len(parts) == 3:
                    wb, sh, idx = parts
                    if wb == workbook and sh == sheet:
                        rows_to_remove.add(int(idx))
            
            if rows_to_remove:
                mask = [i not in rows_to_remove for i in range(len(df))]
                cleaned[workbook][sheet] = df[mask].reset_index(drop=True)
            else:
                cleaned[workbook][sheet] = df.copy()
    
    return cleaned

def dataframe_to_excel_bytes(workbooks_data: Dict) -> Dict[str, BytesIO]:
    """Convert workbooks data to Excel bytes for download"""
    excel_files = {}
    
    for workbook, sheets in workbooks_data.items():
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for sheet_name, df in sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        output.seek(0)
        excel_files[workbook] = output
    
    return excel_files

# ============================================================================
# STREAMLIT UI
# ============================================================================

st.markdown('<p class="main-header">🔍 Excel Duplicate Finder</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Fast duplicate detection with clean, intuitive interface</p>', unsafe_allow_html=True)

# Sidebar settings
with st.sidebar:
    st.header("⚙️ Settings")
    
    st.subheader("Detection Options")
    detect_within_sheet = st.checkbox("Within Each Sheet", value=True)
    detect_across_sheets = st.checkbox("Across Sheets/Files", value=True)
    detect_fuzzy = st.checkbox("Fuzzy Matching", value=False)
    
    if detect_fuzzy:
        fuzzy_threshold = st.slider("Similarity Threshold %", 70, 100, 90)
        max_fuzzy_comparisons = st.number_input("Max Comparisons", 1000, 100000, 50000, 5000)
    
    st.divider()
    
    st.subheader("Performance")
    preview_mode = st.toggle("Preview Mode (10K rows)", value=True)
    st.session_state.preview_mode = preview_mode
    
    if MODIN_AVAILABLE or POLARS_AVAILABLE:
        st.success("✅ Performance libs detected")

# Main tabs
tab1, tab2, tab3, tab4 = st.tabs(["📤 Upload", "🔍 Analyze", "📋 Review", "💾 Download"])

# TAB 1: Upload Files
with tab1:
    st.header("Upload Excel Files")
    
    uploaded_files = st.file_uploader(
        "Choose Excel files",
        type=['xlsx', 'xls'],
        accept_multiple_files=True
    )
    
    if uploaded_files:
        st.success(f"✅ {len(uploaded_files)} file(s) uploaded")
        
        with st.spinner("Loading..."):
            workbooks_data = load_excel_files(uploaded_files, st.session_state.preview_mode, 10000)
            st.session_state.workbooks_data = workbooks_data
        
        total_sheets = sum(len(sheets) for sheets in workbooks_data.values())
        total_rows = sum(len(df) for sheets in workbooks_data.values() for df in sheets.values())
        
        col1, col2, col3 = st.columns(3)
        col1.metric("Workbooks", len(workbooks_data))
        col2.metric("Sheets", total_sheets)
        col3.metric("Rows", f"{total_rows:,}")

# TAB 2: Analyze
with tab2:
    st.header("Analyze for Duplicates")
    
    if not st.session_state.workbooks_data:
        st.warning("⚠️ Upload files first")
    else:
        if st.button("🔍 Start Analysis", type="primary", use_container_width=True):
            all_duplicate_groups = {}
            progress_bar = st.progress(0)
            
            if detect_within_sheet or detect_across_sheets:
                exact_groups = detect_exact_duplicates(st.session_state.workbooks_data)
                for group_key, duplicates in exact_groups.items():
                    if duplicates:
                        dup_type = duplicates[0]['duplicate_type']
                        if ('Within' in duplicates[0]['group_label'] and detect_within_sheet) or \
                           ('Between' in duplicates[0]['group_label'] and detect_across_sheets):
                            all_duplicate_groups[group_key] = duplicates
                progress_bar.progress(0.5)
            
            if detect_fuzzy:
                fuzzy_groups = detect_fuzzy_duplicates(
                    st.session_state.workbooks_data, 
                    threshold=fuzzy_threshold,
                    max_comparisons=max_fuzzy_comparisons
                )
                all_duplicate_groups.update(fuzzy_groups)
                progress_bar.progress(0.9)
            
            st.session_state.duplicate_groups = all_duplicate_groups
            progress_bar.progress(1.0)
            time.sleep(0.3)
            progress_bar.empty()
            
            total_dups = sum(len(dups) for dups in all_duplicate_groups.values())
            
            if total_dups > 0:
                st.success(f"✅ Found {total_dups:,} duplicates in {len(all_duplicate_groups)} groups")
                st.balloons()
            else:
                st.success("🎉 No duplicates found!")

# TAB 3: Clean Review Interface
with tab3:
    st.header("Review & Clean Duplicates")
    
    if not st.session_state.duplicate_groups:
        st.info("ℹ️ Run analysis first")
    else:
        duplicate_groups = st.session_state.duplicate_groups
        total_dups = sum(len(dups) for dups in duplicate_groups.values())
        
        # Summary
        col1, col2 = st.columns([3, 1])
        with col1:
            st.markdown(f"**Found {total_dups:,} duplicates in {len(duplicate_groups)} comparison sets**")
        with col2:
            st.metric("Selected", len(st.session_state.selected_for_removal))
        
        st.divider()
        
        # Quick actions
        st.markdown("**Quick Actions:**")
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            if st.button("✅ Select All", use_container_width=True):
                for dups in duplicate_groups.values():
                    for dup in dups:
                        st.session_state.selected_for_removal.add(f"{dup['workbook']}|{dup['sheet']}|{dup['row_index']}")
                st.rerun()
        
        with col2:
            if st.button("❌ Clear", use_container_width=True):
                st.session_state.selected_for_removal = set()
                st.rerun()
        
        with col3:
            if st.button("🔼 Keep First", use_container_width=True, help="Keep first occurrence, delete rest"):
                for dups in duplicate_groups.values():
                    for dup in dups:
                        st.session_state.selected_for_removal.add(f"{dup['workbook']}|{dup['sheet']}|{dup['row_index']}")
                st.rerun()
        
        with col4:
            if st.button("🔽 Keep Last", use_container_width=True, help="Keep last occurrence, delete rest"):
                for dups in duplicate_groups.values():
                    for dup in dups[:-1]:
                        st.session_state.selected_for_removal.add(f"{dup['workbook']}|{dup['sheet']}|{dup['row_index']}")
                st.rerun()
        
        st.divider()
        
        # Pagination
        items_per_page = st.selectbox("Items per page", [10, 25, 50, 100], index=1)
        
        # Flatten and group duplicates
        all_dups = []
        for group_key, dups in duplicate_groups.items():
            for dup in dups:
                dup['group_key'] = group_key
                all_dups.append(dup)
        
        all_dups.sort(key=lambda x: (x['group_label'], x['workbook'], x['sheet'], x['row_index']))
        
        total_items = len(all_dups)
        total_pages = (total_items + items_per_page - 1) // items_per_page
        
        if 'current_page' not in st.session_state:
            st.session_state.current_page = 0
        
        # Page navigation
        if total_pages > 1:
            col1, col2, col3 = st.columns([1, 2, 1])
            with col1:
                if st.button("← Prev", disabled=st.session_state.current_page == 0):
                    st.session_state.current_page -= 1
                    st.rerun()
            with col2:
                st.markdown(f"<center>Page {st.session_state.current_page + 1} / {total_pages}</center>", unsafe_allow_html=True)
            with col3:
                if st.button("Next →", disabled=st.session_state.current_page >= total_pages - 1):
                    st.session_state.current_page += 1
                    st.rerun()
        
        # Display duplicates by group
        start = st.session_state.current_page * items_per_page
        end = min(start + items_per_page, total_items)
        page_dups = all_dups[start:end]
        
        current_group = None
        
        for idx, dup in enumerate(page_dups):
            unique_id = f"{dup['workbook']}|{dup['sheet']}|{dup['row_index']}"
            is_selected = unique_id in st.session_state.selected_for_removal
            
            # Show group header when it changes
            if current_group != dup['group_label']:
                current_group = dup['group_label']
                st.markdown(f'<div class="group-section"><div class="group-title">📊 {current_group}</div>', unsafe_allow_html=True)
            
            # Duplicate row container
            row_class = "dup-row dup-row-selected" if is_selected else "dup-row"
            st.markdown(f'<div class="{row_class}">', unsafe_allow_html=True)
            
            col1, col2 = st.columns([0.5, 11.5])
            
            with col1:
                # Checkbox
                key = f"cb_{start + idx}_{unique_id}"
                new_val = st.checkbox("", value=is_selected, key=key, label_visibility="collapsed")
                if new_val != is_selected:
                    if new_val:
                        st.session_state.selected_for_removal.add(unique_id)
                    else:
                        st.session_state.selected_for_removal.discard(unique_id)
                    st.rerun()
            
            with col2:
                # Badge and metadata
                badge_class = "badge-exact" if dup['duplicate_type'] == 'Exact' else "badge-fuzzy"
                st.markdown(
                    f'<span class="badge {badge_class}">{dup["duplicate_type"]}</span>'
                    f'<span class="badge {badge_class}">{dup["similarity"]}%</span>'
                    f'<span class="meta-text">📁 {dup["workbook"]} → 📄 {dup["sheet"]} → Row {dup["excel_row"]}</span>',
                    unsafe_allow_html=True
                )
                
                # Row data
                st.dataframe(pd.DataFrame([dup['data']]), use_container_width=True, hide_index=True)
            
            st.markdown('</div>', unsafe_allow_html=True)
        
        # Close last group section
        if current_group:
            st.markdown('</div>', unsafe_allow_html=True)

# TAB 4: Download
with tab4:
    st.header("Download Cleaned Files")
    
    if not st.session_state.workbooks_data:
        st.warning("⚠️ Upload files first")
    else:
        removal_count = len(st.session_state.selected_for_removal)
        
        if removal_count > 0:
            st.success(f"✅ Ready to remove {removal_count:,} rows")
        else:
            st.info("ℹ️ No rows selected")
        
        with st.spinner("Preparing files..."):
            cleaned_data = clean_selected_rows(st.session_state.workbooks_data, st.session_state.selected_for_removal)
            excel_files = dataframe_to_excel_bytes(cleaned_data)
        
        st.divider()
        
        for workbook_name, excel_bytes in excel_files.items():
            col1, col2 = st.columns([3, 1])
            
            with col1:
                st.markdown(f"**📁 {workbook_name}**")
                original = sum(len(df) for df in st.session_state.workbooks_data[workbook_name].values())
                cleaned = sum(len(df) for df in cleaned_data[workbook_name].values())
                if original != cleaned:
                    st.caption(f"{original:,} rows → {cleaned:,} rows (removed {original - cleaned:,})")
            
            with col2:
                st.download_button(
                    "⬇️ Download",
                    excel_bytes,
                    f"CLEANED_{workbook_name}",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )

st.divider()
st.caption("🔒 All processing is done locally. Your data never leaves your machine.")

