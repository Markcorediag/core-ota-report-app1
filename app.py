import io
import datetime as dt
from typing import Optional

import pandas as pd
import streamlit as st
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from PIL import Image as PILImage

# =============================
# Streamlit Page Config
# =============================
st.set_page_config(
    page_title="Core Diagnostics – OTA Calibration Report Generator",
    page_icon="📄",
    layout="centered",
)

st.title("Core Diagnostics – OTA Calibration Report Generator")
st.caption("Upload a Cary Excel sheet to generate a branded PDF with a Brand + Model summary.")

# -----------------------------
# Languages (i18n)
# -----------------------------
LANGUAGES = {
    'en': 'English',
    'fr': 'Français',
    'es': 'Español',
    'da': 'Dansk',
    'sv': 'Svenska',
    'pt': 'Português',
}

I18N = {
    'en': {
        'default_title': 'Core Diagnostics — Cary UK OTA Calibrations Report',
        'intro_default': 'This report summarises the top vehicle brands based on OTA calibration data collected by Core Diagnostics. The analysis reflects brand frequency from the provided dataset and highlights the most common vehicle brands calibrated through OTA processes.',
        'bm_heading': 'Brand + Model Summary',
        'rank_hdr': 'Rank', 'brand_hdr': 'Brand', 'model_hdr': 'Model', 'count_hdr': 'Count',
        'footer': 'Confidential — Core Diagnostics © {year}',
    },
    'fr': {
        'default_title': 'Rapport des calibrations OTA — Core Diagnostics',
        'intro_default': 'Ce rapport résume les principales marques de véhicules sur la base des données de calibration OTA collectées par Core Diagnostics. L’analyse reflète la fréquence des marques dans l’ensemble de données fourni et met en évidence les marques les plus courantes calibrées via des processus OTA.',
        'bm_heading': 'Résumé Marque + Modèle',
        'rank_hdr': 'Rang', 'brand_hdr': 'Marque', 'model_hdr': 'Modèle', 'count_hdr': 'Nombre',
        'footer': 'Confidentiel — Core Diagnostics © {year}',
    },
    'es': {
        'default_title': 'Informe de calibraciones OTA — Core Diagnostics',
        'intro_default': 'Este informe resume las principales marcas de vehículos basándose en los datos de calibraciones OTA recopilados por Core Diagnostics. El análisis refleja la frecuencia de marcas en el conjunto de datos proporcionado y destaca las marcas más comunes calibradas mediante procesos OTA.',
        'bm_heading': 'Resumen Marca + Modelo',
        'rank_hdr': 'Puesto', 'brand_hdr': 'Marca', 'model_hdr': 'Modelo', 'count_hdr': 'Cantidad',
        'footer': 'Confidencial — Core Diagnostics © {year}',
    },
    'da': {
        'default_title': 'OTA-kalibreringsrapport — Core Diagnostics',
        'intro_default': 'Denne rapport sammenfatter de mest almindelige bilmærker baseret på OTA-kalibreringsdata indsamlet af Core Diagnostics. Analysen afspejler mærkefrekvens i det leverede datasæt og fremhæver de mest almindelige bilmærker kalibreret via OTA-processer.',
        'bm_heading': 'Resume Mærke + Model',
        'rank_hdr': 'Rang', 'brand_hdr': 'Mærke', 'model_hdr': 'Model', 'count_hdr': 'Antal',
        'footer': 'Fortroligt — Core Diagnostics © {year}',
    },
    'sv': {
        'default_title': 'OTA-kalibreringsrapport — Core Diagnostics',
        'intro_default': 'Denna rapport sammanfattar de vanligaste bilmärkena baserat på OTA-kalibreringsdata insamlade av Core Diagnostics. Analysen återspeglar varumärkesfrekvens i det tillhandahållna datasetet och lyfter fram de vanligaste märkena som kalibrerats via OTA-processer.',
        'bm_heading': 'Sammanfattning: Märke + Modell',
        'rank_hdr': 'Placering', 'brand_hdr': 'Märke', 'model_hdr': 'Modell', 'count_hdr': 'Antal',
        'footer': 'Konfidentiellt — Core Diagnostics © {year}',
    },
    'pt': {
        'default_title': 'Relatório de calibrações OTA — Core Diagnostics',
        'intro_default': 'Este relatório resume as principais marcas de veículos com base nos dados de calibrações OTA coletados pela Core Diagnostics. A análise reflete a frequência das marcas no conjunto de dados fornecido e destaca as marcas mais comuns calibradas por meio de processos OTA.',
        'bm_heading': 'Resumo Marca + Modelo',
        'rank_hdr': 'Classificação', 'brand_hdr': 'Marca', 'model_hdr': 'Modelo', 'count_hdr': 'Quantidade',
        'footer': 'Confidencial — Core Diagnostics © {year}',
    },
}

# =============================
# Sidebar Controls
# =============================
st.sidebar.header("Report Settings")

# Language selector
lang_code = st.sidebar.selectbox(
    "Language / Langue / Idioma / Sprog / Svenska / Língua",
    options=list(LANGUAGES.keys()),
    format_func=lambda x: LANGUAGES[x],
    index=0,
)

# Language-aware default title (editable per language)
rt_key = f"report_title_{lang_code}"
if rt_key not in st.session_state:
    st.session_state[rt_key] = I18N[lang_code]['default_title']
report_title = st.sidebar.text_input("Report title", value=st.session_state[rt_key], key=rt_key)

# Language-aware default intro (editable per language)
intro_key = f"intro_text_{lang_code}"
if intro_key not in st.session_state:
    st.session_state[intro_key] = I18N[lang_code]['intro_default']
intro_text = st.sidebar.text_area("Introduction paragraph", value=st.session_state[intro_key], height=140, key=intro_key)

max_items = st.sidebar.slider("How many top Brand + Model rows?", min_value=5, max_value=50, value=20, step=1)

st.sidebar.subheader("Branding")
logo_top = st.sidebar.file_uploader("Top logo (PNG/JPG)", type=["png", "jpg", "jpeg"])
logo_top_w = st.sidebar.slider("Top logo width (px)", 60, 300, 140)
logo_top_h = st.sidebar.slider("Top logo height (px)", 20, 150, 45)

logo_bottom = st.sidebar.file_uploader("Bottom 'Powered by Core' image (PNG/JPG)", type=["png", "jpg", "jpeg"], key="bottom")
logo_bottom_w = st.sidebar.slider("Bottom logo width (px)", 40, 260, 100)
logo_bottom_h = st.sidebar.slider("Bottom logo height (px)", 20, 200, 50)

st.sidebar.subheader("Typography")
header_fs = st.sidebar.slider("Header font size", 12, 28, 18)
body_fs = st.sidebar.slider("Body font size", 8, 16, 10)
header_pad = st.sidebar.slider("Table header padding", 2, 12, 6)

# Keep to a single A4 page by trimming rows intelligently
one_page_mode = st.sidebar.checkbox("One-page A4 auto-fit (trim rows if needed)", value=True)

# Spacing controls
title_space = st.sidebar.slider("Space below title (px)", 0, 32, 8)
intro_space = st.sidebar.slider("Space below intro (px)", 0, 48, 12)
table_top_space = st.sidebar.slider("Extra space above table (px)", 0, 64, 12)
body_row_pad = st.sidebar.slider("Table row padding (px)", 1, 12, 4)

# =============================
# File Upload
# =============================
excel_file = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])

# =============================
# Helpers
# =============================
def image_flowable_from_bytes(img_bytes: bytes, max_w: int, max_h: int) -> RLImage:
    """Scale image proportionally to fit within (max_w, max_h)."""
    bio = io.BytesIO(img_bytes)
    img = PILImage.open(bio)
    w, h = img.size
    bio.seek(0)
    max_w = max_w or w
    max_h = max_h or h
    scale = min(max_w / float(w), max_h / float(h))
    new_w = max(1, int(w * scale))
    new_h = max(1, int(h * scale))
    return RLImage(bio, width=new_w, height=new_h)

def detect_brand_column(df: pd.DataFrame) -> Optional[str]:
    """Identify a column that likely contains brand (or brand|model) data."""
    candidates = ["Vehicle Path", "Brand", "Make", "Manufacturer", "VEHICLE PATH", "MAKE", "BRAND"]
    for c in candidates:
        if c in df.columns:
            return c
    # Fallback heuristic: look for a column containing many '|' delimiters
    for col in df.columns:
        if df[col].dtype == object:
            sample = df[col].dropna().astype(str).head(50)
            if not sample.empty and (sample.str.contains("\\|").mean() > 0.2):
                return col
    return None

def extract_brand_series(df: pd.DataFrame, brand_col: str) -> pd.Series:
    s = df[brand_col].dropna().astype(str)
    # If looks like "MAZDA | CX-30 - (DM)" take part before pipe as Brand
    if s.str.contains(r"\\|").mean() > 0.2:
        s = s.str.split("|").str[0]
    s = s.str.strip()
    # Normalise a few common variants
    s = s.replace({
        "VOLKSWAGEN": "VW",
        "MERCEDES": "MERCEDES-BENZ",
        "MERCEDES BENZ": "MERCEDES-BENZ",
        "FORD USA": "FORD",
    }, regex=False)
    return s

def extract_model_series(df: pd.DataFrame, brand_col: str) -> pd.Series:
    """Extract Model from combined column (after '|', before ' - '), or use a dedicated 'Model' column if present."""
    # Prefer explicit model column when available
    for c in ["Model", "MODEL", "Vehicle Model", "VEHICLE MODEL"]:
        if c in df.columns:
            return df[c].dropna().astype(str).str.strip()
    s = df[brand_col].dropna().astype(str)
    if s.str.contains(r"\\|").mean() > 0.2:
        model = s.str.split("|").str[1].str.strip()
        model = model.str.split(" - ").str[0].str.strip()
        # Remove trailing parenthetical codes like " (DM)"
        model = model.str.replace(r"\\s+\\([^)]*\\)$", "", regex=True).str.strip()
        return model
    return pd.Series([], dtype=str)

def compute_brand_model_table(df: pd.DataFrame, brand_col: str, n: int) -> pd.DataFrame:
    brands = extract_brand_series(df, brand_col)
    models = extract_model_series(df, brand_col)
    if models.empty:
        return pd.DataFrame(columns=["Brand", "Model", "Count"])  # no model info
    bm = (
        pd.DataFrame({"Brand": brands, "Model": models})
        .dropna()
        .groupby(["Brand", "Model"]).size()
        .reset_index(name="Count")
        .sort_values(["Count", "Brand", "Model"], ascending=[False, True, True])
        .head(n)
    )
    return bm

def build_pdf(
    df_brand_models: pd.DataFrame,
    title: str,
    intro: str,
    lang: str,
    top_logo_bytes: Optional[bytes],
    bottom_logo_bytes: Optional[bytes],
    top_w: int,
    top_h: int,
    bottom_w: int,
    bottom_h: int,
    header_fs: int,
    body_fs: int,
    header_pad: int,
    one_page_mode: bool,
    title_space: int,
    intro_space: int,
    table_top_space: int,
    body_row_pad: int,
) -> bytes:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()

    # Styles
    styles["Title"].fontSize = header_fs
    styles["Title"].leading = header_fs + 2
    styles["Normal"].fontSize = body_fs
    styles["Normal"].leading = body_fs + 2

    # Looser intro paragraph for readability
    intro_style = ParagraphStyle(
        name="Intro",
        parent=styles["Normal"],
        leading=body_fs + 4,
        spaceAfter=0,
    )

    elements = []

    # Localized strings
    t = I18N.get(lang, I18N['en'])

    # Top Logo
    if top_logo_bytes:
        elements.append(image_flowable_from_bytes(top_logo_bytes, top_w, top_h))
        elements.append(Spacer(1, 8))

    # Title & Intro (with adjustable spacing)
    elements.append(Paragraph(f"<b>{title}</b>", styles["Title"]))
    elements.append(Spacer(1, title_space))
    elements.append(Paragraph(intro, intro_style))
    elements.append(Spacer(1, intro_space))

    # Brand + Model table only (lowered by extra spacer)
    elements.append(Paragraph(f"<b>{t['bm_heading']}</b>", styles["Normal"]))
    elements.append(Spacer(1, table_top_space))

    if df_brand_models.empty:
        elements.append(Paragraph("No model information detected in the selected column.", styles["Normal"]))
    else:
        # Auto-fit rows to one A4 page if enabled (rough estimate based on intro length and logo heights)
        if one_page_mode:
            intro_penalty = max(0, len(intro) // 180)   # every ~180 chars ≈ 1 row
            logo_penalty = int((top_h + bottom_h) / 18) # every ~18px of logos ≈ 1 row
            rows_allowed = max(8, 26 - intro_penalty - logo_penalty)
            df_brand_models = df_brand_models.head(rows_allowed)

        data = [[t['rank_hdr'], t['brand_hdr'], t['model_hdr'], t['count_hdr']]]
        for i, row in enumerate(df_brand_models.itertuples(index=False), start=1):
            data.append([i, getattr(row, "Brand"), getattr(row, "Model"), int(getattr(row, "Count"))])

        table = Table(data, colWidths=[40, 165, 165, 45], repeatRows=1)
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#003366")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("ALIGN", (0, 0), (-1, 0), "CENTER"),
            ("ALIGN", (1, 1), (2, -1), "LEFT"),
            ("ALIGN", (3, 1), (3, -1), "CENTER"),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, 0), body_fs),
            ("BOTTOMPADDING", (0, 0), (-1, 0), header_pad),
            ("TOPPADDING", (0, 0), (-1, 0), header_pad),
            ("BACKGROUND", (0, 1), (-1, -1), colors.whitesmoke),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("FONTSIZE", (0, 1), (-1, -1), body_fs),
            ("BOTTOMPADDING", (0, 1), (-1, -1), body_row_pad),
            ("TOPPADDING", (0, 1), (-1, -1), body_row_pad),
        ]))
        elements.append(table)

    # Bottom branding
    if bottom_logo_bytes:
        elements.append(Spacer(1, 8))
        elements.append(image_flowable_from_bytes(bottom_logo_bytes, bottom_w, bottom_h))

    # Footer
    elements.append(Spacer(1, 6))
    elements.append(Paragraph(
        f"<font size={max(6, body_fs-1)} color=grey>{t['footer'].format(year=dt.date.today().year)}</font>",
        styles["Normal"],
    ))

    doc.build(elements)
    buffer.seek(0)
    return buffer.read()

# =============================
# Main App Flow
# =============================
if excel_file is not None:
    try:
        xl = pd.ExcelFile(excel_file)
        sheet = st.selectbox("Select worksheet", options=xl.sheet_names, index=0)
        df = xl.parse(sheet)
        st.success(f"Loaded sheet: {sheet} — {df.shape[0]} rows × {df.shape[1]} columns")
        st.dataframe(df.head(10))

        # Column selection (auto-detected default)
        detected = detect_brand_column(df)
        brand_col = st.selectbox(
            "Brand/Vehicle Path column",
            options=list(df.columns),
            index=(list(df.columns).index(detected) if detected in df.columns else 0),
            help=("Choose the column that stores brand or 'Brand | Model - (...)' information. "
                  "If using 'Vehicle Path', the brand is parsed before the '|' and model before ' - '."),
        )

        # Compute Brand + Model summary only
        bm_df = compute_brand_model_table(df, brand_col, max_items)
        st.subheader(f"{I18N[lang_code]['bm_heading']} (Preview)")
        if bm_df.empty:
            st.warning("No model information detected. Try selecting 'Vehicle Path' or another column that contains Brand | Model.")
        else:
            st.dataframe(bm_df)

        # === Layout Preview ===
        st.markdown("---")
        st.subheader("Layout Preview")
        if bm_df.empty:
            st.info("Upload data and select a column that includes Model details (e.g., 'Vehicle Path').")
        else:
            preview_df = bm_df.copy()
            if one_page_mode:
                intro_penalty = max(0, len(intro_text) // 180)
                logo_penalty = int((logo_top_h + logo_bottom_h) / 18)
                rows_allowed = max(8, 26 - intro_penalty - logo_penalty)
                preview_df = preview_df.head(rows_allowed)

            if logo_top is not None:
                st.image(logo_top, width=logo_top_w)
            st.markdown(f"### {report_title}")
            st.write(intro_text)
            st.markdown(f"<div style='height:{table_top_space}px'></div>", unsafe_allow_html=True)
            st.dataframe(preview_df)
            if logo_bottom is not None:
                st.image(logo_bottom, width=logo_bottom_w)

        st.markdown("---")
        st.subheader("Generate PDF")
        if st.button("Build Branded PDF"):
            pdf_bytes = build_pdf(
                df_brand_models=bm_df,
                title=report_title,
                intro=intro_text,
                lang=lang_code,
                top_logo_bytes=(logo_top.read() if logo_top is not None else None),
                bottom_logo_bytes=(logo_bottom.read() if logo_bottom is not None else None),
                top_w=logo_top_w,
                top_h=logo_top_h,
                bottom_w=logo_bottom_w,
                bottom_h=logo_bottom_h,
                header_fs=header_fs,
                body_fs=body_fs,
                header_pad=header_pad,
                one_page_mode=one_page_mode,
                title_space=title_space,
                intro_space=intro_space,
                table_top_space=table_top_space,
                body_row_pad=body_row_pad,
            )

            st.download_button(
                label="Download PDF Report",
                data=pdf_bytes,
                file_name="OTA_Calibrations_Report.pdf",
                mime="application/pdf",
            )

        # CSV download
        if not bm_df.empty:
            st.download_button(
                "Download Brand+Model CSV",
                data=bm_df.to_csv(index=False).encode("utf-8"),
                file_name="brand_model_summary.csv",
                mime="text/csv",
            )

    except Exception as e:
        st.error(f"Failed to process the Excel file: {e}")
else:
    st.info("👆 Upload an Excel file to begin. Optionally add your top and bottom logos from the sidebar.")

st.markdown("---")
st.caption("Run locally with `streamlit run app.py`. If needed, install deps: `pip install streamlit pandas reportlab pillow openpyxl`.")
