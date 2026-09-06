{\rtf1\ansi\ansicpg1252\cocoartf2870
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\paperw11900\paperh16840\margl1440\margr1440\vieww11520\viewh8400\viewkind0
\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\partightenfactor0

\f0\fs24 \cf0 """Brazil Item Coding Studio - Streamlit Web Version\
Required packages:\
    streamlit>=1.30.0\
    pandas>=2.0.0\
"""\
\
import streamlit as st\
import pandas as pd\
import re\
\
# Page Configuration\
st.set_page_config(\
    page_title="Brazil Item Coding Studio",\
    page_icon="\uc0\u55356 \u56807 \u55356 \u56823 ",\
    layout="wide"\
)\
\
# ---------------------------------------------------------------------------\
# RPA EXTRACTION ENGINE\
# ---------------------------------------------------------------------------\
class RPAExtractor:\
    PATTERNS = \{\
        "SIZE": r"(?:TAMANHO|TAM|SIZE)\\s*[:\\-]?\\s*(\\d\{2,3\})",\
        "GENDER": r"\\b(MASCULINO|FEMININO|UNISSEX|INFANTIL|MEN'S|WOMEN'S|UNISEX|CHILDREN'S)\\b",\
        "COLOR": r"\\b(PRETO|BRANCO|AZUL|VERMELHO|MARROM|BLACK|WHITE|BLUE|RED|BROWN)\\b"\
    \}\
    \
    @classmethod\
    def auto_extract(cls, description: str) -> dict:\
        desc_upper = (description or "").upper()\
        extracted = \{\}\
        for field, pattern in cls.PATTERNS.items():\
            match = re.search(pattern, desc_upper)\
            if match:\
                val = match.group(1).strip()\
                if field == "GENDER":\
                    trans = \{"MASCULINO": "MEN'S", "FEMININO": "WOMEN'S", "UNISSEX": "UNISEX", "INFANTIL": "CHILDREN'S"\}\
                    val = trans.get(val, val)\
                elif field == "COLOR":\
                    trans = \{"PRETO": "BLACK", "BRANCO": "WHITE", "AZUL": "BLUE", "VERMELHO": "RED", "MARROM": "BROWN"\}\
                    val = trans.get(val, val)\
                extracted[field] = val\
        return extracted\
\
# ---------------------------------------------------------------------------\
# APP STATE INITIALIZATION\
# ---------------------------------------------------------------------------\
if "dataset" not in st.session_state:\
    st.session_state.dataset = pd.DataFrame([\
        \{"Product Title": "CHINELO HAVAIANAS TOP PRETO TAM 42", "SIZE": "", "GENDER": "MEN'S", "COLOR": "", "Status": "Pending"\},\
        \{"Product Title": "SANDALIA RASTEIRA FEMININA AZUL TAM 38", "SIZE": "", "GENDER": "", "COLOR": "", "Status": "Pending"\},\
        \{"Product Title": "TENIS INFANTIL VERMELHO", "SIZE": "", "GENDER": "", "COLOR": "", "Status": "Pending"\}\
    ])\
\
# ---------------------------------------------------------------------------\
# SIDEBAR CONTROLS (RPA & File Upload)\
# ---------------------------------------------------------------------------\
st.sidebar.title("\uc0\u55357 \u57056 \u65039  RPA & Controls")\
\
uploaded_file = st.sidebar.file_uploader("Upload Excel File", type=["xlsx", "csv"])\
if uploaded_file is not None:\
    if uploaded_file.name.endswith(".csv"):\
        st.session_state.dataset = pd.read_csv(uploaded_file)\
    else:\
        st.session_state.dataset = pd.read_excel(uploaded_file)\
    st.sidebar.success("File loaded successfully!")\
\
st.sidebar.markdown("---")\
\
if st.sidebar.button("\uc0\u9889  Run Batch Auto-Code", type="primary", use_container_width=True):\
    df = st.session_state.dataset\
    processed_count = 0\
    \
    for idx, row in df.iterrows():\
        if row.get("Status") == "Completed":\
            continue\
        title = str(row.get("Product Title", ""))\
        if not title:\
            continue\
            \
        entities = RPAExtractor.auto_extract(title)\
        updated = False\
        for field_name, value in entities.items():\
            if field_name in df.columns and (not row.get(field_name) or pd.isna(row.get(field_name))):\
                df.at[idx, field_name] = value\
                updated = True\
                \
        if updated:\
            df.at[idx, "Status"] = "Auto-Coded (Review)"\
            processed_count += 1\
            \
    st.sidebar.success(f"Successfully auto-coded \{processed_count\} rows!")\
    st.rerun()\
\
# ---------------------------------------------------------------------------\
# MAIN WORKSPACE\
# ---------------------------------------------------------------------------\
st.title("\uc0\u55356 \u56807 \u55356 \u56823  Brazil Item Coding Studio \'97 Web Dashboard")\
st.markdown("Manage, review, and auto-extract attributes for footwear catalog items.")\
\
st.subheader("\uc0\u55357 \u56523  Dataset Workspace")\
st.markdown("You can edit fields directly in the table below or use the single-row tool underneath.")\
\
for col in ["Product Title", "SIZE", "GENDER", "COLOR", "Status"]:\
    if col not in st.session_state.dataset.columns:\
        st.session_state.dataset[col] = ""\
\
edited_df = st.data_editor(\
    st.session_state.dataset,\
    num_rows="dynamic",\
    use_container_width=True,\
    key="data_grid"\
)\
st.session_state.dataset = edited_df\
\
st.markdown("---")\
st.subheader("\uc0\u10024  Single-Row Extraction Utility")\
\
if len(st.session_state.dataset) > 0:\
    row_idx = st.selectbox("Choose Row Index to Inspect", st.session_state.dataset.index)\
    selected_title = st.session_state.dataset.loc[row_idx, "Product Title"]\
    \
    col1, col2 = st.columns([3, 1])\
    with col1:\
        st.text_input("Selected Title:", value=selected_title, disabled=True)\
    with col2:\
        st.write("")\
        st.write("")\
        if st.button("Run Auto-Extract"):\
            extracted = RPAExtractor.auto_extract(selected_title)\
            for k, v in extracted.items():\
                if k in st.session_state.dataset.columns:\
                    st.session_state.dataset.loc[row_idx, k] = v\
            st.session_state.dataset.loc[row_idx, "Status"] = "Auto-Coded"\
            st.success("Extracted attributes!")\
            st.rerun()\
\
st.markdown("---")\
csv_data = st.session_state.dataset.to_csv(index=False).encode('utf-8')\
st.download_button(\
    label="\uc0\u55357 \u56549  Download Updated Dataset (CSV)",\
    data=csv_data,\
    file_name="processed_brazil_items.csv",\
    mime="text/csv",\
)}