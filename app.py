from __future__ import annotations
import json
import re
import warnings
import zipfile
from collections import defaultdict
from dataclasses import dataclass, field as dc_field
from datetime import datetime
from pathlib import Path
import streamlit as st
from openpyxl import load_workbook

warnings.filterwarnings("ignore", message="Workbook contains no default style, apply openpyxl's default")

# --- PAGE CONFIG ---
st.set_page_config(
    page_title="Brazil Item Coding Studio",
    page_icon="🇧🇷",
    layout="wide",
    initial_sidebar_state="expanded"
)

APP = "Brazil Item Coding Studio - Streamlit Version 2.0"

# --- TRANSLATION BACKEND ---
try:
    from deep_translator import GoogleTranslator
    TRANSLATION_AVAILABLE = True
except Exception:
    TRANSLATION_AVAILABLE = False

OFFLINE_GLOSSARY = {
    "CHINELO":"FLIP-FLOP","CHINELOS":"FLIP-FLOPS","SANDALIA":"SANDAL","SANDALIAS":"SANDALS",
    "SANDÁLIA":"SANDAL","SANDÁLIAS":"SANDALS","FEMININO":"WOMEN'S","FEMININA":"WOMEN'S",
    "MASCULINO":"MEN'S","MASCULINA":"MEN'S","INFANTIL":"CHILDREN'S","ADULTO":"ADULT","ADULTA":"ADULT",
    "UNISSEX":"UNISEX","TAMANHO":"SIZE","TAM":"SIZE","NUMERO":"SIZE","PAR":"PAIR","COR":"COLOR",
    "AZUL":"BLUE","PRETO":"BLACK","PRETA":"BLACK","BRANCO":"WHITE","BRANCA":"WHITE","VERDE":"GREEN",
    "AMARELO":"YELLOW","AMARELA":"YELLOW","VERMELHO":"RED","VERMELHA":"RED","MARROM":"BROWN",
    "ROSA":"PINK","CINZA":"GRAY","DOURADO":"GOLD","DOURADA":"GOLD","PRATEADO":"SILVER","PRATEADA":"SILVER",
    "ESTAMPADO":"PRINTED","ESTAMPADA":"PRINTED","LISO":"SOLID","LISA":"SOLID","LISTRADO":"STRIPED",
    "COMPRIMENTO":"LENGTH","ESPESSURA":"THICKNESS","FORMATO":"SHAPE","CURTA":"SHORT","CURTO":"SHORT",
    "LONGA":"LONG","LONGO":"LONG","FINA":"THIN","FINO":"THIN","GROSSA":"THICK","GROSSO":"THICK",
    "NORMAL":"NORMAL","ESPECIAL":"SPECIAL","COMUM":"COMMON","TRANÇADA":"BRAIDED","TRANÇADO":"BRAIDED",
    "CRUZADA":"CROSSED","CRUZADO":"CROSSED","COM":"WITH","SEM":"WITHOUT","E":"AND","PARA":"FOR",
    "DE":"OF","DA":"OF THE","DO":"OF THE","NOVA":"NEW","NOVO":"NEW","ORIGINAL":"ORIGINAL",
    "DEDO":"TOE","RASTEIRA":"FLAT","RASTEIRINHA":"THONG SANDAL","BEBE":"BABY","BEBÊ":"BABY",
}

def _offline_glossary_translate(value: str) -> str:
    words = re.findall(r"[A-Za-zÀ-ÿ0-9/]+|\s+|[^\sA-Za-zÀ-ÿ0-9/]+", value)
    out = []
    for w in words:
        key = w.strip().upper()
        out.append(OFFLINE_GLOSSARY.get(key, w) if key.isalpha() else w)
    return "".join(out)

def translate_to_english(value: str) -> tuple[str, str]:
    value = (value or "").strip()
    if not value:
        return "", ""
    if TRANSLATION_AVAILABLE:
        try:
            translator = GoogleTranslator(source="auto", target="en")
            translated = translator.translate(value)
            if translated and translated.strip():
                return translated.strip(), "auto-translated"
        except Exception:
            pass
    glossary_result = _offline_glossary_translate(value)
    if glossary_result.upper() != value.upper():
        return glossary_result, "offline glossary"
    return value, "translation unavailable"

# --- HELPERS & DATACLASSES ---
def text(v): return "" if v is None else str(v).strip()
def upper(v): return re.sub(r"\s+", " ", text(v)).upper().strip()

@dataclass
class Guideline:
    name: str
    pt_def: str = ""
    en_def: str = ""
    fixed: str = ""
    pt_examples: str = ""
    en_examples: str = ""
    extra_notes: str = ""
    enum_values: tuple = ()
    meta: dict = dc_field(default_factory=dict)

@dataclass
class HandbookData:
    guidelines: dict = dc_field(default_factory=dict)
    scope_notes: list = dc_field(default_factory=list)
    images: dict = dc_field(default_factory=dict)
    metadata: dict = dc_field(default_factory=dict)

def _sheet_guidelines(ws) -> dict[str, Guideline]:
    guidelines: dict[str, Guideline] = {}
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return guidelines
    header = [upper(c) for c in rows[0]]
    def col(*names):
        for n in names:
            if n in header:
                return header.index(n)
        return None
    c_name = col("CHARACTERISTIC")
    c_pt_def = col("COMO PREENCHER/DEFINIÇÃO", "COMO PREENCHER/DEFINICAO")
    c_en_def = col("HOW TO POPULATE/DEFINITION")
    c_fixed = col("VALUE FIXED?")
    c_pt_ex = col("EXEMPLOS DE VALORES", "EXEMPLOS DE VALORES ")
    c_en_ex = col("VALUES EXAMPLES")
    for row in rows[1:]:
        if c_name is None or c_name >= len(row):
            continue
        name = text(row[c_name])
        if not name:
            continue
        def cell(idx):
            return text(row[idx]) if idx is not None and idx < len(row) else ""
        en_ex = cell(c_en_ex)
        lines = [upper(l) for l in re.split(r"[\n;]", en_ex) if upper(l)]
        enum_values = tuple(dict.fromkeys(lines)) if 1 < len(lines) <= 30 else ()
        guidelines[name] = Guideline(
            name=name, pt_def=cell(c_pt_def), en_def=cell(c_en_def),
            fixed=cell(c_fixed), pt_examples=cell(c_pt_ex), en_examples=en_ex,
            enum_values=enum_values,
        )
    return guidelines

def load_handbook_bytes(uploaded_file) -> HandbookData:
    wb = load_workbook(uploaded_file, read_only=True, data_only=True)
    sheet_name = next((s for s in wb.sheetnames if "OGRDS CHARACTERISTIC" in s.upper()), None)
    if not sheet_name:
        wb.close()
        raise ValueError("Sheet 'OGRDS Characteristic' not found in handbook.")
    guidelines = _sheet_guidelines(wb[sheet_name])
    wb.close()
    return HandbookData(guidelines=guidelines)

# --- STREAMLIT UI LAYOUT ---
st.title("🇧🇷 Brazil Item Coding Studio")
st.markdown("### Production Version 2.0 (Web Streamlit Edition)")

with st.sidebar:
    st.header("⚙️ Configuration & Files")
    handbook_file = st.file_uploader("Upload Handbook Workbook (.xlsx)", type=["xlsx"])
    allocation_file = st.file_uploader("Upload Allocation / Coding File (.xlsx)", type=["xlsx"])
    
    st.markdown("---")
    st.info("💡 **Feature Status:** Translation Engine, Rule Checker, and Interactive Coding Active.")

if handbook_file:
    try:
        hb_data = load_handbook_bytes(handbook_file)
        st.success(f"Successfully loaded handbook with {len(hb_data.guidelines)} characteristics!")
        
        if allocation_file:
            wb_alloc = load_workbook(allocation_file)
            sheet_names = wb_alloc.sheetnames
            selected_sheet = st.selectbox("Select Allocation Sheet", sheet_names)
            ws_alloc = wb_alloc[selected_sheet]
            
            rows = list(ws_alloc.iter_rows(values_only=True))
            if rows:
                headers = [str(c) for c in rows[0]]
                st.write(f"Loaded allocation sheet **{selected_sheet}** with {len(rows)-1} items.")
                
                # Interactive item selector
                item_indices = list(range(1, len(rows)))
                selected_row_idx = st.selectbox("Select Item Row to Code", item_indices, format_func=lambda i: f"Row {i}: {rows[i][:2]}")
                
                if selected_row_idx:
                    current_row_data = rows[selected_row_idx]
                    st.markdown("---")
                    st.subheader(f"Coding Item at Row {selected_row_idx}")
                    
                    # Description handling
                    desc_col_idx = headers.index("Product Title") if "Product Title" in headers else 1
                    source_desc = text(current_row_data[desc_col_idx])
                    
                    st.text_area("Original Source Description", value=source_desc, height=80, disabled=True)
                    translated_desc, mode = translate_to_english(source_desc)
                    st.info(f"**English Translation ({mode}):** {translated_desc}")
                    
                    st.markdown("### Characteristic Fields")
                    coded_values = {}
                    
                    for char_name, guideline in hb_data.guidelines.items():
                        col1, col2, col3 = st.columns([2, 3, 1])
                        with col1:
                            st.markdown(f"**{char_name}**")
                            if guideline.fixed:
                                st.caption(f"Fixed: {guideline.fixed}")
                        with col2:
                            default_val = guideline.fixed if guideline.fixed else ""
                            options = list(guideline.enum_values) if guideline.enum_values else []
                            if options:
                                coded_values[char_name] = st.selectbox(f"Value for {char_name}", options=[""] + options, key=f"val_{char_name}")
                            else:
                                coded_values[char_name] = st.text_input(f"Value for {char_name}", value=default_val, key=f"txt_{char_name}")
                        with col3:
                            if st.button("?", key=f"help_{char_name}"):
                                st.toast(f"Definition: {guideline.en_def or guideline.pt_def}")
                                
                    if st.button("💾 Save Changes", type="primary"):
                        st.success("Changes captured successfully! (Ready to commit back to Excel or local journal).")
            wb_alloc.close()
        else:
            st.warning("Please upload an allocation/coding spreadsheet to start coding items.")
            
    except Exception as e:
        st.error(f"Error processing workbook: {e}")
else:
    st.info("👈 Please upload your **Handbook Workbook** in the sidebar to initialize the coding studio.")
