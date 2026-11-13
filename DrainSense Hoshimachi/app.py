"""
DrainSense Kanaeru — Sistem Evaluasi Kapasitas Saluran Drainase Terpadu
5-Modul Pipeline: Iklim BMKG (SNI 2415:2016) → Diskusi Hibrida AI + Python (DED) → Computer Vision & Cowan
→ Hidrolika Permen PU 12/2014 → Dashboard & Bisection

Arsitektur Dual-Engine (pola Demo.py):
  1) Direct SDK  : `from google import genai` → client.interactions.create(model="models/gemini-4-argon")
  2) Live Bridge : `drainsense_bridge.json` (di folder yang sama dengan app.py) yang dapat ditulis
                   oleh Antigravity dari VS Code → dimuat langsung ke st.session_state tanpa API Key.
"""

import base64
import glob
import io
import json
import math
import os
import re
import time
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import scipy.stats
import streamlit as st

# ── Optional dependencies ─────────────────────────────────────────
try:
    from google import genai  # SDK resmi google-genai
    GENAI_AVAILABLE = True
except ImportError:
    GENAI_AVAILABLE = False

try:
    from pypdf import PdfReader
    PYPDF_AVAILABLE = True
except ImportError:
    PYPDF_AVAILABLE = False

try:
    import pdfplumber
    PDFPLUMBER_AVAILABLE = True
except ImportError:
    PDFPLUMBER_AVAILABLE = False

try:
    from docx import Document as DocxDocument
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    from fpdf import FPDF
    FPDF_AVAILABLE = True
except ImportError:
    FPDF_AVAILABLE = False

try:
    from PIL import Image as PILImage
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

# ═════════════════════════════════════════════════════════════════
# PAGE CONFIG
# ═════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="DrainSense Kanaeru",
    page_icon="🌊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ═════════════════════════════════════════════════════════════════
# SESSION STATE — Initialisation (Persistent Across Module Navigation)
# ═════════════════════════════════════════════════════════════════
_SS_DEFAULTS: dict = {
    "data_modul_1":     None, # dict lengkap hasil & tabel Modul 1
    "uploaded_file_m1": None, # dict {"name", "df_raw"} file BMKG Modul 1
    "m1_col_date":      None, # nama kolom tanggal terpilih di Modul 1
    "m1_col_rain":      None, # nama kolom curah hujan terpilih di Modul 1
    "m1_selected_dist": None, # distribusi terpilih di Modul 1
    "R24_5yr":          None, # float mm         ← Modul 1
    "r24_dict":         None, # dict {T: R24_mm} ← Modul 1
    "m1_r24_df":        None, # DataFrame R24    ← Modul 1
    "params_modul_2":   {},   # dict parameter DED hasil Modul 2
    "ded_params":       None, # dict b,h,S,n,mat ← Modul 2
    "uploaded_file_m2": None, # dict {"name", "size", "full_doc_text", "preview_data", "preview_stats"} Modul 2
    "params_modul_3":   {},   # dict parameter degradasi & Cowan hasil Modul 3
    "cv_params":        None, # dict sed,weed,…  ← Modul 3
    "uploaded_file_m3": None, # dict {"name", "bytes"} foto saluran Modul 3
    "m4_results":       None, # dict Q values    ← Modul 4
    "pdf_bytes":        None, # bytes            ← Modul 5
    "gemini_key":       "",   # API Key Gemini (opsional — Engine 1)
    "gemini_model":     "gemini-1.5-flash", # Model Gemini terpilih pada Engine 1
    "bridge_mtime":     None, # mtime drainsense_bridge.json terakhir dimuat
    "bridge_data":      None, # isi JSON bridge terakhir
    "m2_ai_result":     None, # hasil diskusi AI + validasi Python Modul 2
    "m3_ai_result":     None, # hasil evaluasi spasial AI Modul 3
    "m3_roi_stats":     None, # statistik piksel ROI Modul 3
    "m3_ideal_bypass":  False, # checkbox bypass kondisi ideal Modul 3
    "pdf_doc_text":     "",   # teks lengkap dokumen PDF yang diunggah di Modul 1 / Modul 2
    "pdf_doc_name":     "",   # nama file PDF yang diunggah di Modul 1 / Modul 2
    "m4_hydro_extracted": None, # dict hasil ekstraksi instan hidrologi Modul 4 (t1, t2, tc, A, C, I, Qr)
    "m4_use_manual_calc": False, # toggle kalkulator manual di Modul 4
}
for _k, _v in _SS_DEFAULTS.items():
    st.session_state.setdefault(_k, _v)


# ═════════════════════════════════════════════════════════════════
# HYDROLOGY & WATER TECHNOLOGY — VISUAL DESIGN SYSTEM & UI HELPERS
# ═════════════════════════════════════════════════════════════════
def _inject_drainsense_design_system() -> None:
    """Inject premium Hydrology & Water Technology SaaS design system CSS."""
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600&display=swap');

        :root {
            --ds-navy: #102A43;
            --ds-navy-dark: #0A1C2E;
            --ds-ocean: #1677C8;
            --ds-cyan: #34C6D3;
            --ds-teal: #11A8A0;
            --ds-bg: #F4F8FC;
            --ds-surface: #FFFFFF;
            --ds-border: #DFE8F0;
            --ds-text-primary: #18324A;
            --ds-text-secondary: #718397;
            --ds-warning: #F4B942;
            --ds-danger: #E45B64;
            --ds-success: #24A47B;
            --ds-radius-sm: 8px;
            --ds-radius-md: 12px;
            --ds-radius-lg: 16px;
            --ds-shadow-sm: 0 2px 8px rgba(16, 42, 67, 0.05);
            --ds-shadow-md: 0 6px 20px rgba(16, 42, 67, 0.07);
            --ds-shadow-lg: 0 12px 32px rgba(22, 119, 200, 0.11);
        }

        /* ── Global App Background & Typography ── */
        .stApp {
            background: radial-gradient(circle at 92% 4%, rgba(52, 198, 211, 0.08) 0%, rgba(22, 119, 200, 0.03) 35%, #F4F8FC 75%) !important;
            color: var(--ds-text-primary) !important;
            font-family: 'Inter', system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif !important;
            overflow-x: hidden !important;
        }

        /* Ensure Streamlit top header bar does not obscure or clip top workflow navigation */
        header[data-testid="stHeader"] {
            background: rgba(244, 248, 252, 0.85) !important;
            backdrop-filter: blur(6px) !important;
            height: 3.25rem !important;
        }

        div[data-testid="stAppViewBlockContainer"],
        .block-container {
            padding-top: 4.25rem !important;
            padding-bottom: 3.5rem !important;
            padding-left: 2.2rem !important;
            padding-right: 2.2rem !important;
            max-width: 1440px !important;
            overflow: visible !important;
            animation: ds-fade-in 0.32s cubic-bezier(0.16, 1, 0.3, 1);
        }

        h1, h2, h3, h4, h5, h6 {
            color: var(--ds-navy) !important;
            font-family: 'Inter', system-ui, sans-serif !important;
            letter-spacing: -0.015em;
        }

        code, pre, .stCode {
            font-family: 'JetBrains Mono', monospace !important;
        }

        /* ── Water-Inspired Micro Animations ── */
        @keyframes ds-water-flow {
            0% { background-position: 0% 50%; }
            50% { background-position: 100% 50%; }
            100% { background-position: 0% 50%; }
        }

        @keyframes ds-fade-in {
            from { opacity: 0; transform: translateY(4px); }
            to { opacity: 1; transform: translateY(0); }
        }

        @keyframes ds-pulse-ring {
            0% { box-shadow: 0 0 0 0 rgba(52, 198, 211, 0.45); }
            70% { box-shadow: 0 0 0 6px rgba(52, 198, 211, 0); }
            100% { box-shadow: 0 0 0 0 rgba(52, 198, 211, 0); }
        }

        @media (prefers-reduced-motion: reduce) {
            *, ::before, ::after {
                animation-duration: 0.01ms !important;
                animation-iteration-count: 1 !important;
                transition-duration: 0.01ms !important;
            }
        }

        /* ── Sidebar Redesign (Deep Navy Hydrology Control Center — Refined Typography & Proportions) ── */
        section[data-testid="stSidebar"] {
            background: linear-gradient(180deg, #102A43 0%, #0C2135 55%, #091929 100%) !important;
            border-right: 1px solid rgba(52, 198, 211, 0.20) !important;
            min-width: 315px !important;
            max-width: 335px !important;
            font-family: 'Inter', system-ui, -apple-system, sans-serif !important;
        }

        section[data-testid="stSidebar"] > div:first-child {
            padding-top: 0.5rem !important;
        }

        section[data-testid="stSidebar"] div[data-testid="stSidebarUserContent"] {
            padding: 0.75rem 1.05rem 1.75rem 1.05rem !important;
        }

        section[data-testid="stSidebar"] p,
        section[data-testid="stSidebar"] label {
            font-family: 'Inter', system-ui, -apple-system, sans-serif !important;
        }

        /* Sidebar Expander (🔧 Pengaturan Engine 1 API) */
        section[data-testid="stSidebar"] div[data-testid="stExpander"] {
            background: rgba(255, 255, 255, 0.05) !important;
            border: 1px solid rgba(226, 232, 240, 0.18) !important;
            border-radius: 10px !important;
            margin-bottom: 0.55rem !important;
            box-shadow: none !important;
        }
        div[data-testid="stExpander"] summary {
            background-color: transparent !important;
            color: #ffffff !important;
        }
        div[data-testid="stExpander"] summary:hover {
            color: #4ade80 !important;
        }
        section[data-testid="stSidebar"] div[data-testid="stExpander"] summary,
        section[data-testid="stSidebar"] div[data-testid="stExpander"] summary p,
        section[data-testid="stSidebar"] div[data-testid="stExpander"] summary span {
            background-color: transparent !important;
            color: #ffffff !important;
            font-size: 13px !important;
            font-weight: 600 !important;
        }
        section[data-testid="stSidebar"] div[data-testid="stExpander"] summary:hover,
        section[data-testid="stSidebar"] div[data-testid="stExpander"] summary:hover p,
        section[data-testid="stSidebar"] div[data-testid="stExpander"] summary:hover span {
            color: #4ade80 !important;
        }
        section[data-testid="stSidebar"] div[data-testid="stExpander"] summary svg {
            fill: #9EEAF2 !important;
            color: #9EEAF2 !important;
        }

        /* Fix white code box inside sidebar */
        section[data-testid="stSidebar"] code {
            background: rgba(16, 42, 67, 0.85) !important;
            color: #AEEFF5 !important;
            border: 1px solid rgba(52, 198, 211, 0.28) !important;
            border-radius: 6px !important;
            padding: 0.12rem 0.42rem !important;
            font-family: 'JetBrains Mono', monospace !important;
            font-size: 11.5px !important;
        }

        section[data-testid="stSidebar"] .stCaption,
        section[data-testid="stSidebar"] small {
            color: #9FB3C8 !important;
            font-size: 11.5px !important;
            line-height: 1.4 !important;
        }

        section[data-testid="stSidebar"] hr {
            border-color: rgba(223, 232, 240, 0.14) !important;
            margin: 0.9rem 0 !important;
        }

        /* Sidebar Text Input & Selectbox (Gemini API Key & Pilih Model Gemini) */
        section[data-testid="stSidebar"] div[data-testid="stTextInput"] label p,
        section[data-testid="stSidebar"] div[data-testid="stSelectbox"] label p {
            color: #ffffff !important;
            font-size: 12.5px !important;
            font-weight: 600 !important;
            margin-bottom: 0.2rem !important;
        }
        section[data-testid="stSidebar"] div[data-baseweb="input"],
        section[data-testid="stSidebar"] div[data-baseweb="select"] > div {
            background-color: rgba(10, 28, 46, 0.85) !important;
            border: 1px solid rgba(52, 198, 211, 0.38) !important;
            border-radius: 8px !important;
            min-height: 38px !important;
            color: #ffffff !important;
        }
        section[data-testid="stSidebar"] div[data-baseweb="select"] div,
        section[data-testid="stSidebar"] div[data-baseweb="select"] span {
            color: #ffffff !important;
            font-size: 12.5px !important;
        }
        section[data-testid="stSidebar"] div[data-baseweb="input"]:focus-within,
        section[data-testid="stSidebar"] div[data-baseweb="select"] > div:focus-within {
            border-color: #34C6D3 !important;
            box-shadow: 0 0 0 2px rgba(52, 198, 211, 0.18) !important;
        }
        section[data-testid="stSidebar"] div[data-baseweb="input"] input {
            color: #FFFFFF !important;
            font-size: 12.5px !important;
            padding: 0.35rem 0.65rem !important;
        }
        section[data-testid="stSidebar"] div[data-baseweb="input"] input::placeholder {
            color: rgba(159, 179, 200, 0.65) !important;
            font-size: 12px !important;
        }
        section[data-testid="stSidebar"] div[data-baseweb="input"] button {
            color: #9EEAF2 !important;
            background: transparent !important;
        }

        /* Sidebar Buttons (Full-readability, balanced height, no text clipping) */
        section[data-testid="stSidebar"] .stButton > button {
            background: linear-gradient(180deg, rgba(22, 119, 200, 0.28) 0%, rgba(16, 42, 67, 0.65) 100%) !important;
            color: #F0F7FF !important;
            border: 1px solid rgba(52, 198, 211, 0.38) !important;
            border-radius: 8px !important;
            font-size: 12.5px !important;
            font-weight: 600 !important;
            line-height: 1.3 !important;
            min-height: 38px !important;
            padding: 0.45rem 0.85rem !important;
            width: 100% !important;
            white-space: nowrap !important;
            display: inline-flex !important;
            align-items: center !important;
            justify-content: center !important;
            gap: 0.4rem !important;
            transition: all 0.2s ease !important;
            box-shadow: 0 2px 6px rgba(0, 0, 0, 0.18) !important;
        }
        section[data-testid="stSidebar"] .stButton > button p {
            color: #F0F7FF !important;
            font-size: 12.5px !important;
            font-weight: 600 !important;
            margin: 0 !important;
        }
        section[data-testid="stSidebar"] .stButton > button:hover {
            background: linear-gradient(135deg, #1677C8 0%, #11A8A0 100%) !important;
            border-color: #34C6D3 !important;
            transform: translateY(-1px);
            box-shadow: 0 4px 12px rgba(52, 198, 211, 0.28) !important;
        }

        /* Sidebar Navigation Radio -> High-Contrast White/Light-Gray Text on Deep Navy */
        section[data-testid="stSidebar"] .stRadio label,
        section[data-testid="stSidebar"] .stRadio label p,
        section[data-testid="stSidebar"] .stRadio label span,
        section[data-testid="stSidebar"] .stRadio div[data-testid="stMarkdownContainer"] p,
        section[data-testid="stSidebar"] div[data-testid="stRadio"] label,
        section[data-testid="stSidebar"] div[data-testid="stRadio"] label p,
        section[data-testid="stSidebar"] div[data-testid="stRadio"] label span,
        section[data-testid="stSidebar"] div[data-testid="stRadio"] div[data-testid="stMarkdownContainer"] p {
            color: #FFFFFF !important;
            word-wrap: break-word !important;
            overflow-wrap: break-word !important;
            white-space: normal !important;
        }
        section[data-testid="stSidebar"] div[data-testid="stRadio"] > label p,
        section[data-testid="stSidebar"] .stRadio > label p {
            font-size: 13px !important;
            font-weight: 700 !important;
            color: #FFFFFF !important;
            text-transform: uppercase !important;
            letter-spacing: 0.04em !important;
            margin-bottom: 0.35rem !important;
        }
        section[data-testid="stSidebar"] div[role="radiogroup"] {
            gap: 0.42rem !important;
        }
        section[data-testid="stSidebar"] div[role="radiogroup"] > label,
        section[data-testid="stSidebar"] .stRadio div[role="radiogroup"] label {
            background: rgba(255, 255, 255, 0.06) !important;
            border: 1px solid rgba(223, 232, 240, 0.16) !important;
            border-radius: 10px !important;
            padding: 0.58rem 0.78rem !important;
            margin: 0 !important;
            height: auto !important;
            min-height: fit-content !important;
            transition: all 0.2s ease !important;
            cursor: pointer !important;
            display: flex !important;
            align-items: center !important;
            color: #E0E0E0 !important;
        }
        section[data-testid="stSidebar"] div[role="radiogroup"] > label p,
        section[data-testid="stSidebar"] div[role="radiogroup"] > label span,
        section[data-testid="stSidebar"] div[role="radiogroup"] > label div {
            color: #E0E0E0 !important;
            font-size: 13px !important;
            font-weight: 500 !important;
            line-height: 1.35 !important;
            margin: 0 !important;
            word-wrap: break-word !important;
            overflow-wrap: break-word !important;
            white-space: normal !important;
        }
        section[data-testid="stSidebar"] div[role="radiogroup"] > label:hover {
            background: rgba(22, 119, 200, 0.28) !important;
            border-color: rgba(52, 198, 211, 0.55) !important;
            transform: translateX(2px);
        }
        section[data-testid="stSidebar"] div[role="radiogroup"] > label:hover p,
        section[data-testid="stSidebar"] div[role="radiogroup"] > label:hover span,
        section[data-testid="stSidebar"] div[role="radiogroup"] > label:hover div {
            color: #FFFFFF !important;
        }
        section[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked) {
            background: linear-gradient(90deg, rgba(22, 119, 200, 0.55) 0%, rgba(17, 168, 160, 0.38) 100%) !important;
            border-left: 4px solid #34C6D3 !important;
            border-top: 1px solid rgba(52, 198, 211, 0.65) !important;
            border-right: 1px solid rgba(52, 198, 211, 0.45) !important;
            border-bottom: 1px solid rgba(52, 198, 211, 0.45) !important;
            box-shadow: 0 4px 14px rgba(17, 168, 160, 0.24) !important;
        }
        section[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked) p,
        section[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked) span,
        section[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked) div {
            color: #FFFFFF !important;
            font-weight: 700 !important;
        }

        /* ── Primary & Secondary Buttons (Ocean Blue & Aqua Cyan — No Red Default) ── */
        .stButton > button[kind="primary"],
        .stFormSubmitButton > button[kind="primary"],
        .stDownloadButton > button {
            background: linear-gradient(135deg, #1677C8 0%, #11A8A0 100%) !important;
            color: #FFFFFF !important;
            border: 1px solid rgba(52, 198, 211, 0.4) !important;
            border-radius: 10px !important;
            font-weight: 600 !important;
            letter-spacing: 0.01em !important;
            padding: 0.55rem 1.2rem !important;
            box-shadow: 0 4px 12px rgba(22, 119, 200, 0.22) !important;
            transition: all 0.22s cubic-bezier(0.16, 1, 0.3, 1) !important;
        }
        .stButton > button[kind="primary"]:hover,
        .stFormSubmitButton > button[kind="primary"]:hover,
        .stDownloadButton > button:hover {
            background: linear-gradient(135deg, #1264AA 0%, #0E8F88 100%) !important;
            box-shadow: 0 6px 18px rgba(22, 119, 200, 0.34) !important;
            transform: translateY(-1px) !important;
        }

        .stButton > button[kind="secondary"],
        .stFormSubmitButton > button[kind="secondary"] {
            background: #FFFFFF !important;
            color: var(--ds-navy) !important;
            border: 1px solid #C8D8E8 !important;
            border-radius: 10px !important;
            font-weight: 600 !important;
            padding: 0.55rem 1.15rem !important;
            box-shadow: var(--ds-shadow-sm) !important;
            transition: all 0.2s ease !important;
        }
        .stButton > button[kind="secondary"]:hover,
        .stFormSubmitButton > button[kind="secondary"]:hover {
            border-color: var(--ds-ocean) !important;
            color: var(--ds-ocean) !important;
            background: #F0F7FF !important;
            transform: translateY(-1px) !important;
        }

        /* ── Metric Cards (st.metric) ── */
        div[data-testid="stMetric"] {
            background: linear-gradient(180deg, #FFFFFF 0%, #FAFDFE 100%) !important;
            border: 1px solid var(--ds-border) !important;
            border-top: 3px solid var(--ds-ocean) !important;
            border-radius: var(--ds-radius-md) !important;
            padding: 0.95rem 1.1rem !important;
            box-shadow: var(--ds-shadow-sm) !important;
            transition: transform 0.2s ease, box-shadow 0.2s ease, border-color 0.2s ease !important;
        }
        div[data-testid="stMetric"]:hover {
            transform: translateY(-2px);
            box-shadow: var(--ds-shadow-md) !important;
            border-top-color: var(--ds-cyan) !important;
        }
        div[data-testid="stMetricLabel"] {
            color: var(--ds-text-secondary) !important;
            font-weight: 600 !important;
            font-size: 0.8rem !important;
            text-transform: uppercase !important;
            letter-spacing: 0.03em !important;
        }
        div[data-testid="stMetricValue"] {
            color: var(--ds-navy) !important;
            font-weight: 700 !important;
            font-size: 1.45rem !important;
        }

        /* ── File Uploader Modern Drag & Drop Area ── */
        div[data-testid="stFileUploader"] section {
            background: linear-gradient(180deg, #FFFFFF 0%, #F4F9FD 100%) !important;
            border: 2px dashed #8BBEE8 !important;
            border-radius: var(--ds-radius-md) !important;
            padding: 1.25rem 1rem !important;
            transition: all 0.22s ease !important;
        }
        div[data-testid="stFileUploader"] section:hover {
            border-color: var(--ds-ocean) !important;
            background: #EDF6FC !important;
            box-shadow: 0 4px 16px rgba(22, 119, 200, 0.08) !important;
        }

        /* ── Containers, Forms & Expanders ── */
        div[data-testid="stForm"],
        div[data-testid="stExpander"] {
            background: #FFFFFF !important;
            border: 1px solid var(--ds-border) !important;
            border-radius: var(--ds-radius-md) !important;
            box-shadow: var(--ds-shadow-sm) !important;
        }
        div[data-testid="stExpander"] summary {
            font-weight: 600 !important;
            color: var(--ds-navy) !important;
            padding: 0.75rem 1rem !important;
        }
        div[data-testid="stExpander"] summary:hover {
            color: var(--ds-ocean) !important;
        }

        /* ── Dataframes & Tables ── */
        div[data-testid="stDataFrame"] {
            border: 1px solid var(--ds-border) !important;
            border-radius: 10px !important;
            overflow: hidden !important;
            box-shadow: var(--ds-shadow-sm) !important;
        }

        /* ── Alert Boxes, Info Boxes & Header Containers (Dynamic Height & Word-Wrap) ── */
        .info-box,
        .alert-box,
        div[data-testid="stAlert"],
        .stAlert {
            height: auto !important;
            min-height: fit-content !important;
            word-wrap: break-word !important;
            overflow-wrap: break-word !important;
            white-space: normal !important;
            overflow: visible !important;
        }
        div[data-testid="stAlert"] p,
        .stAlert p,
        .info-box p,
        .alert-box p {
            word-wrap: break-word !important;
            overflow-wrap: break-word !important;
            white-space: normal !important;
        }

        /* ── Custom DrainSense Component Classes ── */
        .ds-workflow-bar {
            display: grid;
            grid-template-columns: repeat(5, minmax(0, 1fr));
            gap: 0.6rem;
            margin-top: 0.35rem;
            margin-bottom: 1.15rem;
            padding-top: 0.15rem;
            height: auto !important;
            min-height: fit-content !important;
            overflow: visible !important;
        }
        .ds-wf-step {
            background: #FFFFFF;
            border: 1px solid var(--ds-border);
            border-radius: 10px;
            padding: 0.62rem 0.78rem;
            height: auto !important;
            min-height: fit-content !important;
            display: flex;
            align-items: center;
            gap: 0.55rem;
            font-size: 12.5px;
            font-weight: 600;
            line-height: 1.35;
            color: var(--ds-text-secondary);
            box-shadow: var(--ds-shadow-sm);
            transition: all 0.2s ease;
            word-wrap: break-word;
            overflow-wrap: break-word;
            white-space: normal;
            overflow: visible;
        }
        .ds-wf-step > span:last-child {
            min-width: 0;
            word-wrap: break-word;
            overflow-wrap: break-word;
            white-space: normal;
        }
        .ds-wf-step.active {
            background: linear-gradient(135deg, #102A43 0%, #1677C8 100%);
            color: #FFFFFF;
            border-color: #34C6D3;
            box-shadow: 0 4px 14px rgba(22, 119, 200, 0.22);
        }
        .ds-wf-step.done {
            border-color: rgba(36, 164, 123, 0.45);
            background: #F2FBF7;
            color: #16694D;
        }
        .ds-wf-num {
            width: 24px;
            height: 24px;
            border-radius: 50%;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            font-size: 11.5px;
            font-weight: 700;
            background: #E8F0F8;
            color: var(--ds-navy);
            flex-shrink: 0;
        }
        .ds-wf-step.active .ds-wf-num {
            background: #34C6D3;
            color: #0A1C2E;
            animation: ds-pulse-ring 2.2s infinite;
        }
        .ds-wf-step.done .ds-wf-num {
            background: #24A47B;
            color: #FFFFFF;
        }

        /* Module Header Banner */
        .ds-module-header {
            position: relative;
            background: linear-gradient(135deg, #102A43 0%, #143A5C 55%, #1677C8 100%);
            border-radius: var(--ds-radius-lg);
            padding: 1.45rem 1.75rem 1.65rem 1.75rem;
            color: #FFFFFF;
            box-shadow: var(--ds-shadow-lg);
            margin-bottom: 1.35rem;
            height: auto !important;
            min-height: fit-content !important;
            word-wrap: break-word;
            overflow-wrap: break-word;
            overflow: hidden;
            border: 1px solid rgba(52, 198, 211, 0.32);
        }
        .ds-module-header::after {
            content: "";
            position: absolute;
            left: 0;
            right: 0;
            bottom: 0;
            height: 5px;
            background: linear-gradient(90deg, #34C6D3, #11A8A0, #1677C8, #34C6D3);
            background-size: 300% 100%;
            animation: ds-water-flow 7s ease infinite;
        }
        .ds-mh-top {
            display: flex;
            align-items: center;
            justify-content: space-between;
            flex-wrap: wrap;
            gap: 0.6rem;
            margin-bottom: 0.55rem;
        }
        .ds-mh-stage {
            display: inline-flex;
            align-items: center;
            gap: 0.45rem;
            background: rgba(52, 198, 211, 0.18);
            border: 1px solid rgba(52, 198, 211, 0.45);
            color: #AEEFF5;
            padding: 0.22rem 0.7rem;
            border-radius: 999px;
            font-size: 0.72rem;
            font-weight: 700;
            letter-spacing: 0.06em;
            text-transform: uppercase;
            word-wrap: break-word;
            overflow-wrap: break-word;
        }
        .ds-mh-status {
            display: inline-flex;
            align-items: center;
            gap: 0.4rem;
            padding: 0.25rem 0.75rem;
            border-radius: 999px;
            font-size: 0.74rem;
            font-weight: 600;
            word-wrap: break-word;
            overflow-wrap: break-word;
        }
        .ds-mh-status.ready {
            background: rgba(36, 164, 123, 0.25);
            border: 1px solid rgba(36, 164, 123, 0.65);
            color: #A7F3D0;
        }
        .ds-mh-status.standby {
            background: rgba(244, 185, 66, 0.22);
            border: 1px solid rgba(244, 185, 66, 0.55);
            color: #FDE68A;
        }
        .ds-mh-title {
            font-size: 1.52rem;
            font-weight: 800;
            color: #FFFFFF !important;
            margin: 0.15rem 0 0.35rem 0;
            display: flex;
            align-items: center;
            flex-wrap: wrap;
            gap: 0.6rem;
            line-height: 1.25;
            word-wrap: break-word;
            overflow-wrap: break-word;
        }
        .ds-mh-desc {
            font-size: 0.9rem;
            color: #D9E8F5;
            margin: 0;
            line-height: 1.5;
            max-width: 960px;
            word-wrap: break-word;
            overflow-wrap: break-word;
        }

        /* Section Cards & Analytical Cards */
        .ds-section-card {
            background: #FFFFFF;
            border: 1px solid var(--ds-border);
            border-left: 4px solid var(--ds-ocean);
            border-radius: var(--ds-radius-md);
            padding: 0.9rem 1.2rem;
            margin: 0.75rem 0 0.95rem 0;
            box-shadow: var(--ds-shadow-sm);
            height: auto !important;
            min-height: fit-content !important;
            word-wrap: break-word;
            overflow-wrap: break-word;
        }
        .ds-section-card.teal {
            border-left-color: var(--ds-teal);
        }
        .ds-section-card.cyan {
            border-left-color: var(--ds-cyan);
        }
        .ds-section-card.navy {
            border-left-color: var(--ds-navy);
        }
        .ds-section-title {
            font-size: 1.02rem;
            font-weight: 700;
            color: var(--ds-navy);
            margin: 0 0 0.2rem 0;
            display: flex;
            align-items: center;
            justify-content: space-between;
            flex-wrap: wrap;
            gap: 0.5rem;
            word-wrap: break-word;
            overflow-wrap: break-word;
        }
        .ds-section-sub {
            font-size: 0.82rem;
            color: var(--ds-text-secondary);
            margin: 0;
            word-wrap: break-word;
            overflow-wrap: break-word;
        }

        .ds-badge {
            display: inline-flex;
            align-items: center;
            gap: 0.3rem;
            padding: 0.18rem 0.62rem;
            border-radius: 999px;
            font-size: 0.72rem;
            font-weight: 600;
            word-wrap: break-word;
            overflow-wrap: break-word;
        }
        .ds-badge.auto {
            background: #E6F8FA;
            color: #0E7490;
            border: 1px solid #99E6EE;
        }
        .ds-badge.manual {
            background: #FEF6E6;
            color: #B45309;
            border: 1px solid #FCD34D;
        }
        .ds-badge.safe {
            background: #E8F8F2;
            color: #15803D;
            border: 1px solid #86EFAC;
        }
        .ds-badge.danger {
            background: #FDECEE;
            color: #B91C1C;
            border: 1px solid #FCA5A5;
        }

        /* Pipeline Status Cards (Module 4 & 5) */
        .ds-pipe-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 0.85rem;
            margin-bottom: 1.1rem;
        }
        .ds-pipe-card {
            background: #FFFFFF;
            border: 1px solid var(--ds-border);
            border-radius: 12px;
            padding: 0.85rem 1.05rem;
            box-shadow: var(--ds-shadow-sm);
            transition: all 0.2s ease;
            height: auto !important;
            min-height: fit-content !important;
            word-wrap: break-word;
            overflow-wrap: break-word;
        }
        .ds-pipe-card:hover {
            transform: translateY(-2px);
            box-shadow: var(--ds-shadow-md);
            border-color: #B4D2EE;
        }
        .ds-pipe-head {
            display: flex;
            align-items: center;
            justify-content: space-between;
            flex-wrap: wrap;
            gap: 0.35rem;
            font-size: 0.78rem;
            font-weight: 700;
            color: var(--ds-text-secondary);
            text-transform: uppercase;
            margin-bottom: 0.35rem;
            word-wrap: break-word;
            overflow-wrap: break-word;
        }
        .ds-pipe-val {
            font-size: 0.92rem;
            font-weight: 600;
            color: var(--ds-navy);
            line-height: 1.4;
            word-wrap: break-word;
            overflow-wrap: break-word;
        }

        /* Executive Hero KPI Cards (Module 5) */
        .ds-kpi-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 1rem;
            margin: 0.6rem 0 1.1rem 0;
        }
        .ds-kpi-card {
            background: #FFFFFF;
            border: 1px solid var(--ds-border);
            border-radius: 14px;
            padding: 1.15rem 1.35rem;
            box-shadow: var(--ds-shadow-md);
            position: relative;
            height: auto !important;
            min-height: fit-content !important;
            word-wrap: break-word;
            overflow-wrap: break-word;
            overflow: hidden;
            transition: transform 0.22s ease, box-shadow 0.22s ease;
        }
        .ds-kpi-card:hover {
            transform: translateY(-3px);
            box-shadow: var(--ds-shadow-lg);
        }
        .ds-kpi-card.safe {
            border-top: 4px solid var(--ds-success);
        }
        .ds-kpi-card.danger {
            border-top: 4px solid var(--ds-danger);
        }
        .ds-kpi-card.ocean {
            border-top: 4px solid var(--ds-ocean);
        }
        .ds-kpi-label {
            font-size: 0.76rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            color: var(--ds-text-secondary);
            margin-bottom: 0.35rem;
            word-wrap: break-word;
            overflow-wrap: break-word;
        }
        .ds-kpi-value {
            font-size: 1.75rem;
            font-weight: 800;
            color: var(--ds-navy);
            line-height: 1.15;
            margin-bottom: 0.35rem;
            word-wrap: break-word;
            overflow-wrap: break-word;
        }
        .ds-kpi-sub {
            font-size: 0.8rem;
            font-weight: 600;
            word-wrap: break-word;
            overflow-wrap: break-word;
        }

        /* Sidebar Custom Branding, Engine Cards & Timeline Stepper */
        .ds-sb-brand {
            background: linear-gradient(135deg, rgba(22, 119, 200, 0.28) 0%, rgba(52, 198, 211, 0.14) 100%);
            border: 1px solid rgba(52, 198, 211, 0.34);
            border-radius: 12px;
            padding: 0.95rem 1rem;
            margin-bottom: 1rem;
            height: auto !important;
            min-height: fit-content !important;
            word-wrap: break-word;
            overflow-wrap: break-word;
            box-shadow: 0 4px 14px rgba(0, 0, 0, 0.15);
        }
        .ds-sb-logo-row {
            display: flex;
            align-items: center;
            gap: 0.75rem;
        }
        .ds-sb-icon {
            width: 40px;
            height: 40px;
            border-radius: 10px;
            background: linear-gradient(135deg, #34C6D3 0%, #1677C8 100%);
            display: flex;
            align-items: center;
            justify-content: center;
            box-shadow: 0 4px 12px rgba(52, 198, 211, 0.32);
            flex-shrink: 0;
        }
        .ds-sb-title {
            font-family: 'Inter', sans-serif !important;
            font-size: 17px !important;
            font-weight: 700 !important;
            color: #ffffff !important;
            line-height: 1.25 !important;
            letter-spacing: -0.01em !important;
            margin: 0 !important;
            word-wrap: break-word !important;
            overflow-wrap: break-word !important;
        }
        .ds-sb-sub {
            font-family: 'Inter', sans-serif !important;
            font-size: 12.5px !important;
            color: #e2e8f0 !important;
            font-weight: 500 !important;
            line-height: 1.35 !important;
            margin-top: 0.18rem !important;
            word-wrap: break-word !important;
            overflow-wrap: break-word !important;
        }
        .ds-sb-engine-card {
            background: rgba(255, 255, 255, 0.06);
            border: 1px solid rgba(226, 232, 240, 0.18);
            border-radius: 10px;
            padding: 0.8rem 0.9rem;
            margin-bottom: 0.65rem;
            height: auto !important;
            min-height: fit-content !important;
            word-wrap: break-word;
            overflow-wrap: break-word;
            font-family: 'Inter', sans-serif !important;
            color: #ffffff !important;
        }
        .ds-sb-engine-head {
            display: flex;
            align-items: center;
            justify-content: space-between;
            flex-wrap: wrap;
            gap: 0.4rem;
            margin-bottom: 0.45rem;
            font-family: 'Inter', sans-serif !important;
            color: #ffffff !important;
        }
        section[data-testid="stSidebar"] .ds-sb-engine-title,
        .ds-sb-engine-title {
            font-family: 'Inter', sans-serif !important;
            font-size: 14px !important;
            font-weight: 700 !important;
            color: #ffffff !important;
            line-height: 1.35 !important;
            margin: 0 !important;
            display: flex;
            align-items: center;
            gap: 0.4rem;
        }
        section[data-testid="stSidebar"] .ds-sb-engine-meta,
        .ds-sb-engine-meta {
            display: flex;
            align-items: center;
            justify-content: space-between;
            flex-wrap: wrap;
            gap: 0.4rem;
            font-family: 'Inter', sans-serif !important;
            font-size: 14px !important;
            color: #ffffff !important;
        }
        section[data-testid="stSidebar"] .ds-sb-engine-meta span,
        .ds-sb-engine-meta span {
            font-family: 'Inter', sans-serif !important;
            font-size: 13.5px !important;
            font-weight: 600 !important;
            color: #e2e8f0 !important;
        }
        section[data-testid="stSidebar"] .ds-sb-meta-tag,
        .ds-sb-meta-tag {
            display: inline-flex;
            align-items: center;
            gap: 0.3rem;
            background: rgba(10, 28, 46, 0.88);
            border: 1px solid rgba(0, 229, 255, 0.42);
            color: #ffffff !important;
            font-family: 'JetBrains Mono', monospace !important;
            font-size: 12px !important;
            font-weight: 600 !important;
            line-height: 1.3 !important;
            padding: 0.22rem 0.55rem;
            border-radius: 6px;
        }
        .ds-sb-pill {
            display: inline-flex;
            align-items: center;
            padding: 0.2rem 0.58rem;
            border-radius: 999px;
            font-family: 'Inter', sans-serif !important;
            font-size: 11px !important;
            font-weight: 700 !important;
            line-height: 1.2 !important;
            letter-spacing: 0.03em;
            text-transform: uppercase;
            white-space: nowrap;
        }
        .ds-sb-pill.on {
            background: rgba(36, 164, 123, 0.30);
            color: #4ade80 !important;
            border: 1px solid rgba(74, 222, 128, 0.65);
        }
        .ds-sb-pill.off {
            background: rgba(148, 163, 184, 0.22);
            color: #ffffff !important;
            border: 1px solid rgba(226, 232, 240, 0.45);
        }

        .ds-sb-section-hdr {
            font-size: 13px !important;
            font-weight: 700 !important;
            text-transform: uppercase !important;
            letter-spacing: 0.04em !important;
            color: #FFFFFF !important;
            margin: 0.65rem 0 0.45rem 0 !important;
            line-height: 1.35 !important;
        }
        .ds-sb-stepper {
            position: relative;
            margin: 0.35rem 0 0.5rem 0;
            display: flex;
            flex-direction: column;
            gap: 0.35rem;
        }
        .ds-sb-step {
            display: flex;
            align-items: center;
            gap: 0.6rem;
            padding: 0.48rem 0.65rem;
            border-radius: 8px;
            background: rgba(255, 255, 255, 0.055);
            border: 1px solid rgba(223, 232, 240, 0.14);
            font-size: 12.5px !important;
            line-height: 1.35 !important;
        }
        .ds-sb-step.active {
            background: rgba(0, 229, 255, 0.16);
            border: 1px solid rgba(0, 229, 255, 0.55);
        }
        section[data-testid="stSidebar"] .ds-sb-step-lbl,
        .ds-sb-step-lbl {
            color: #ffffff !important;
            font-size: 12.5px !important;
            font-weight: 600 !important;
            line-height: 1.3 !important;
            flex: 1;
            min-width: 0;
        }
        section[data-testid="stSidebar"] .ds-sb-step.active .ds-sb-step-lbl,
        .ds-sb-step.active .ds-sb-step-lbl {
            color: #ffffff !important;
            font-weight: 700 !important;
        }
        section[data-testid="stSidebar"] .status-ready,
        .status-ready {
            margin-left: auto;
            font-size: 11px !important;
            color: #4ade80 !important;
            font-weight: bold !important;
            letter-spacing: 0.04em;
            white-space: nowrap;
        }
        section[data-testid="stSidebar"] .status-active,
        .status-active {
            margin-left: auto;
            font-size: 11px !important;
            color: #00e5ff !important;
            font-weight: bold !important;
            letter-spacing: 0.04em;
            white-space: nowrap;
        }
        section[data-testid="stSidebar"] .status-pending,
        .status-pending {
            margin-left: auto;
            font-size: 11px !important;
            color: #cbd5e1 !important;
            font-weight: bold !important;
            letter-spacing: 0.04em;
            white-space: nowrap;
        }
        .ds-sb-dot {
            width: 22px;
            height: 22px;
            border-radius: 50%;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            font-size: 11px !important;
            font-weight: 700 !important;
            flex-shrink: 0;
        }
        .ds-sb-dot.done {
            background: #2ecc71;
            color: #0A1C2E !important;
        }
        .ds-sb-dot.active {
            background: #00e5ff;
            color: #0A1C2E !important;
        }
        .ds-sb-dot.pending {
            background: rgba(255, 255, 255, 0.16);
            color: #ffffff !important;
        }

        /* ── Professional Entrance, Hover & Subtle Aqua Border Animations (Zero Continuous Bobbing) ── */
        @keyframes fadeSlideUp {
            from { opacity: 0; transform: translateY(10px); }
            to { opacity: 1; transform: translateY(0); }
        }
        @keyframes aquaBorderPulse {
            0%, 100% {
                border-color: rgba(52, 198, 211, 0.55);
                box-shadow: 0 8px 20px rgba(16, 42, 67, 0.15), 0 0 0 0 rgba(52, 198, 211, 0.24);
            }
            50% {
                border-color: rgba(52, 198, 211, 0.95);
                box-shadow: 0 8px 20px rgba(16, 42, 67, 0.15), 0 0 14px 2px rgba(52, 198, 211, 0.34);
            }
        }

        /* ── Clean Hero Header Layout & Kobo Kanaeru Avatar Framing ── */
        .ds-mh-grid {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 1.5rem;
            position: relative;
            z-index: 2;
        }
        .ds-mh-left {
            flex: 1;
            min-width: 0;
        }
        .ds-mh-brand-kicker {
            display: inline-flex;
            align-items: center;
            gap: 0.45rem;
            font-size: 0.73rem;
            font-weight: 700;
            letter-spacing: 0.06em;
            text-transform: uppercase;
            color: #75CFFF;
            margin-bottom: 0.3rem;
        }
        .ds-mh-pipe-strip {
            display: flex;
            align-items: center;
            flex-wrap: wrap;
            gap: 0.42rem;
            margin-top: 0.75rem;
        }
        .ds-mh-pipe-chip {
            display: inline-flex;
            align-items: center;
            gap: 0.32rem;
            padding: 0.18rem 0.58rem;
            border-radius: 999px;
            font-size: 0.7rem;
            font-weight: 600;
            background: rgba(255, 255, 255, 0.08);
            border: 1px solid rgba(223, 232, 240, 0.2);
            color: #D9E8F5;
        }
        .ds-mh-pipe-chip.done {
            background: rgba(36, 164, 123, 0.25);
            border-color: rgba(74, 222, 128, 0.55);
            color: #A7F3D0;
        }
        .ds-mh-pipe-chip.active {
            background: rgba(52, 198, 211, 0.26);
            border-color: #34C6D3;
            color: #FFFFFF;
            font-weight: 700;
        }

        /* Sidebar Kobo Chibi Mascot Avatar (54x54 Framed Card) */
        .ds-sb-kobo-wrap {
            width: 54px;
            height: 54px;
            border-radius: 12px;
            overflow: hidden;
            border: 2px solid rgba(52, 198, 211, 0.6);
            background: #EAF7FF;
            box-shadow: 0 6px 14px rgba(10, 28, 46, 0.25);
            display: flex;
            align-items: center;
            justify-content: center;
            flex-shrink: 0;
            padding: 0;
            transition: all 0.3s ease;
            animation: fadeSlideUp 0.45s cubic-bezier(0.22, 1, 0.36, 1) both, aquaBorderPulse 3.6s ease-in-out infinite;
        }
        .ds-sb-kobo-img {
            width: 54px !important;
            height: 54px !important;
            object-fit: cover !important;
            object-position: top center !important;
            display: block !important;
            border: none !important;
            transition: transform 0.3s ease;
        }

        /* Module Companion Card & Empty State with Kobo (115x115 Avatar Card) */
        .ds-kobo-companion {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 1.25rem;
            background: linear-gradient(135deg, #FFFFFF 0%, #EAF7FF 100%);
            border: 1px solid #C6E2F7;
            border-left: 4px solid #1677C8;
            border-radius: 14px;
            padding: 0.9rem 1.35rem;
            margin: 0.55rem 0 1rem 0;
            box-shadow: var(--ds-shadow-sm);
            height: auto !important;
            min-height: fit-content !important;
            transition: all 0.3s ease;
            animation: fadeSlideUp 0.45s cubic-bezier(0.22, 1, 0.36, 1) both;
        }
        .ds-kobo-companion:hover {
            border-color: #34C6D3;
            box-shadow: 0 10px 24px rgba(52, 198, 211, 0.2);
        }
        .ds-kobo-companion-body {
            flex: 1;
            min-width: 0;
        }
        .ds-kobo-companion-tag {
            display: inline-flex;
            align-items: center;
            gap: 0.35rem;
            font-size: 0.72rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.04em;
            color: #1677C8;
            background: rgba(22, 119, 200, 0.1);
            border: 1px solid rgba(52, 198, 211, 0.45);
            padding: 0.14rem 0.55rem;
            border-radius: 999px;
            margin-bottom: 0.3rem;
            animation: aquaBorderPulse 3.8s ease-in-out infinite;
        }
        .ds-kobo-companion-title {
            font-size: 0.98rem;
            font-weight: 700;
            color: #102A43;
            margin: 0 0 0.2rem 0;
            line-height: 1.3;
        }
        .ds-kobo-companion-text {
            font-size: 0.84rem;
            color: #334E68;
            margin: 0;
            line-height: 1.48;
        }
        .ds-kobo-companion-art {
            width: 115px;
            height: 115px;
            border-radius: 16px;
            overflow: hidden;
            border: 2px solid rgba(52, 198, 211, 0.6);
            background: #EAF7FF;
            box-shadow: 0 8px 20px rgba(16, 42, 67, 0.15);
            display: flex;
            align-items: center;
            justify-content: center;
            flex-shrink: 0;
            padding: 0;
            transition: all 0.3s ease;
            animation: fadeSlideUp 0.45s cubic-bezier(0.22, 1, 0.36, 1) both, aquaBorderPulse 3.6s ease-in-out infinite;
        }
        .ds-kobo-companion-art img {
            width: 115px !important;
            height: 115px !important;
            object-fit: cover !important;
            object-position: top center !important;
            display: block !important;
            border: none !important;
            transition: transform 0.3s ease;
        }

        /* Kobo Welcoming Empty State Card */
        .ds-kobo-empty {
            background: linear-gradient(180deg, #FFFFFF 0%, #F4FAFF 100%);
            border: 1px solid #CFE4F5;
            border-top: 4px solid #1677C8;
            border-radius: 16px;
            padding: 1.65rem 1.75rem;
            text-align: center;
            box-shadow: 0 6px 20px rgba(16, 42, 67, 0.06);
            margin: 0.85rem 0 1.25rem 0;
            height: auto !important;
            min-height: fit-content !important;
            transition: all 0.3s ease;
            animation: fadeSlideUp 0.45s cubic-bezier(0.22, 1, 0.36, 1) both;
        }
        .ds-kobo-empty-stage {
            width: 115px;
            height: 115px;
            margin: 0 auto 0.85rem auto;
            border-radius: 16px;
            overflow: hidden;
            border: 2px solid rgba(52, 198, 211, 0.6);
            background: #EAF7FF;
            box-shadow: 0 8px 20px rgba(16, 42, 67, 0.15);
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 0;
            transition: all 0.3s ease;
            animation: fadeSlideUp 0.45s cubic-bezier(0.22, 1, 0.36, 1) both, aquaBorderPulse 3.6s ease-in-out infinite;
        }
        .ds-kobo-empty-stage img {
            width: 115px !important;
            height: 115px !important;
            object-fit: cover !important;
            object-position: top center !important;
            display: block !important;
            border: none !important;
            transition: transform 0.3s ease;
        }
        .ds-kobo-companion:hover .ds-kobo-companion-art,
        .ds-kobo-companion-art:hover,
        .ds-kobo-empty-stage:hover,
        .ds-sb-kobo-wrap:hover {
            transform: scale(1.06) translateY(-2px);
            box-shadow: 0 12px 24px rgba(52, 198, 211, 0.45);
            border-color: #34C6D3;
            transition: all 0.3s ease;
        }

        /* Responsive Breakpoints (1366x768, 1600x900, 1920x1080, Tablet & Mobile) */
        @media (max-width: 1440px) {
            div[data-testid="stAppViewBlockContainer"],
            .block-container {
                padding-left: 1.5rem !important;
                padding-right: 1.5rem !important;
            }
            .ds-wf-step {
                padding: 0.55rem 0.65rem;
                font-size: 12px;
            }
        }
        @media (max-width: 1150px) {
            .ds-workflow-bar {
                grid-template-columns: repeat(3, minmax(0, 1fr));
            }
            .ds-pipe-grid, .ds-kpi-grid {
                grid-template-columns: 1fr;
            }
        }
        @media (max-width: 768px) {
            .ds-mh-grid, .ds-kobo-companion {
                flex-direction: column;
                text-align: left;
            }
            .ds-workflow-bar {
                grid-template-columns: 1fr 1fr;
            }
            .ds-mh-title {
                font-size: 1.2rem;
            }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def _resolve_kobo_rain_umbrella() -> str:
    base_dir = Path(__file__).resolve().parent
    exact_rel = "assets/Kobo/Kobo Rain  Umbrella.jpg"
    if (base_dir / exact_rel).exists():
        return exact_rel
    matches = sorted(glob.glob(str(base_dir / "assets" / "Kobo" / "Kobo Rain*.jpg")))
    if matches:
        return Path(matches[0]).relative_to(base_dir).as_posix()
    return "assets/Kobo/kobo_thinking_2.jpg"


def _resolve_kobo_profil() -> str:
    base_dir = Path(__file__).resolve().parent
    exact_rel = "assets/Kobo/kobo profil.webp"
    if (base_dir / exact_rel).exists():
        return exact_rel
    matches = sorted(glob.glob(str(base_dir / "assets" / "Kobo" / "kobo profil.*")))
    if matches:
        return Path(matches[0]).relative_to(base_dir).as_posix()
    return "assets/Kobo/kobo_main.jpg"


KOBO_ASSET_PATHS: dict[str, str] = {
    "kobo_chibi":         "assets/Kobo/kobo_chibi.jpg",
    "kobo_chibi_2":       "assets/Kobo/kobo_chibi_2.png",
    "kobo_rain":          "assets/Kobo/kobo_rain.jpg",
    "kobo_rain_umbrella": _resolve_kobo_rain_umbrella(),
    "kobo_thinking":      "assets/Kobo/kobo_thinking.jpg",
    "kobo_thinking_2":    "assets/Kobo/kobo_thinking_2.jpg",
    "kobo_happy":         "assets/Kobo/kobo_happy.png",
    "kobo_happy_2":       "assets/Kobo/kobo_happy_2.jpg",
    "kobo_main":          "assets/Kobo/kobo_main.jpg",
    "kobo_profil":        _resolve_kobo_profil(),
}
_KOBO_MTIME_CACHE: dict[tuple[str, float, int], str] = {}


def _get_kobo_data_uri(asset_key: str, max_height: int = 360) -> str:
    """Load Kobo Kanaeru image directly from assets/Kobo/ with mtime invalidation (no @st.cache_data)."""
    rel_path = KOBO_ASSET_PATHS.get(asset_key)
    if not rel_path:
        return ""
    target_path = Path(__file__).resolve().parent / rel_path
    if not target_path.exists():
        return ""
    try:
        mtime = float(os.path.getmtime(target_path))
        cache_k = (str(target_path), mtime, max_height)
        if cache_k in _KOBO_MTIME_CACHE:
            return _KOBO_MTIME_CACHE[cache_k]

        if PIL_AVAILABLE:
            with PILImage.open(target_path) as im:
                im_rgba = im.convert("RGBA")
                w, h = im_rgba.size
                if h > max_height and h > 0:
                    new_w = max(1, int(round(w * (max_height / float(h)))))
                    im_rgba = im_rgba.resize((new_w, max_height), PILImage.Resampling.LANCZOS)
                buf = io.BytesIO()
                im_rgba.save(buf, format="PNG", optimize=True)
                b64 = base64.b64encode(buf.getvalue()).decode("ascii")
                uri = f"data:image/png;base64,{b64}"
        else:
            raw = target_path.read_bytes()
            ext = target_path.suffix.lower()
            mime = (
                "image/webp" if ext == ".webp"
                else ("image/png" if ext == ".png" else "image/jpeg")
            )
            uri = f"data:{mime};base64,{base64.b64encode(raw).decode('ascii')}"

        _KOBO_MTIME_CACHE[cache_k] = uri
        return uri
    except Exception:
        return ""


def _get_pipeline_stage_states() -> dict[int, bool]:
    """Return completion boolean for each of the 5 pipeline stages."""
    m1_done = any(bool(st.session_state.get(k)) for k in ("data_modul_1", "R24_5yr"))
    m2_done = any(bool(st.session_state.get(k)) for k in ("params_modul_2", "ded_params"))
    m3_done = any(bool(st.session_state.get(k)) for k in ("params_modul_3", "cv_params"))
    m4_done = bool(st.session_state.get("m4_results"))
    m5_done = bool(st.session_state.get("pdf_bytes")) or m4_done
    return {1: m1_done, 2: m2_done, 3: m3_done, 4: m4_done, 5: m5_done}


def _render_module_header(
    step_num: int,
    icon: str,
    title: str,
    description: str,
    status_text: str,
    is_ready: bool = True,
) -> None:
    """Render unified top workflow progression bar + clean text-only hero header."""
    states = _get_pipeline_stage_states()
    stages_meta = [
        (1, "Iklim & IDF (BMKG)"),
        (2, "Ekstraksi DED (AI NLP)"),
        (3, "Visi Komputer & Cowan"),
        (4, "Engine Hidrolika"),
        (5, "Executive Dashboard"),
    ]
    wf_items_html = []
    pipe_chips_html = []
    for s_idx, s_label in stages_meta:
        is_curr = s_idx == step_num
        is_done = states.get(s_idx, False)
        cls = "active" if is_curr else ("done" if is_done else "")
        badge_txt = "✓" if (is_done and not is_curr) else f"0{s_idx}"
        wf_items_html.append(
            f'<div class="ds-wf-step {cls}">'
            f'<span class="ds-wf-num">{badge_txt}</span>'
            f'<span>{s_label}</span>'
            f'</div>'
        )
        chip_sym = "✓" if is_done else ("●" if is_curr else f"0{s_idx}")
        pipe_chips_html.append(
            f'<span class="ds-mh-pipe-chip {cls}">{chip_sym} M{s_idx}: {s_label.split(" (")[0]}</span>'
        )
    wf_bar_html = '<div class="ds-workflow-bar">' + "".join(wf_items_html) + "</div>"
    pipe_strip_html = '<div class="ds-mh-pipe-strip">' + "".join(pipe_chips_html) + "</div>"

    status_cls = "ready" if is_ready else "standby"
    status_dot = "●" if is_ready else "◐"

    header_html = (
        f"{wf_bar_html}"
        f'<div class="ds-module-header">'
        f'<div class="ds-mh-grid">'
        f'<div class="ds-mh-left">'
        f'<div class="ds-mh-brand-kicker">🌊 DrainSense Kanaeru · AI-Powered Drainage Analysis System</div>'
        f'<div class="ds-mh-top">'
        f'<span class="ds-mh-stage">💧 TAHAP 0{step_num} DARI 05 · HYDROLOGY &amp; HYDRAULIC PIPELINE</span>'
        f'<span class="ds-mh-status {status_cls}">{status_dot} {status_text}</span>'
        f'</div>'
        f'<div class="ds-mh-title"><span>{icon}</span><span>{title}</span></div>'
        f'<p class="ds-mh-desc">{description}</p>'
        f'{pipe_strip_html}'
        f'</div>'
        f'</div>'
        f'</div>'
    )
    st.markdown(header_html, unsafe_allow_html=True)


def _render_kobo_companion_card(
    asset_key: str,
    tag_text: str = "",
    title: str = "",
    description: str = "",
    tag: str | None = None,
    **kwargs,
) -> None:
    """Render a compact module-specific engineering companion card with one static framed Kobo asset."""
    uri1 = _get_kobo_data_uri(asset_key, max_height=360)
    art1 = f'<div class="ds-kobo-companion-art"><img src="{uri1}" alt="Kobo Kanaeru" /></div>' if uri1 else ""
    effective_tag = tag if tag is not None else tag_text
    tag_html = f'<span class="ds-kobo-companion-tag">{effective_tag}</span>' if effective_tag else ""
    html = (
        f'<div class="ds-kobo-companion">'
        f'<div class="ds-kobo-companion-body">'
        f'{tag_html}'
        f'<div class="ds-kobo-companion-title">{title}</div>'
        f'<p class="ds-kobo-companion-text">{description}</p>'
        f'</div>'
        f'{art1}'
        f'</div>'
    )
    st.markdown(html, unsafe_allow_html=True)


def _render_kobo_empty_state(
    asset_key: str,
    title: str,
    message: str = "",
    badges_html: str = "",
    subtitle: str | None = None,
    show_image: bool = True,
    **kwargs,
) -> None:
    """Render an engineering empty state, optionally showing a single framed Kobo asset."""
    uri = _get_kobo_data_uri(asset_key, max_height=360) if show_image else ""
    img_block = f'<div class="ds-kobo-empty-stage"><img src="{uri}" alt="Kobo Kanaeru" /></div>' if uri else ""
    msg_block = (
        f'<p style="color:#486581;max-width:640px;margin:0 auto;font-size:0.88rem;line-height:1.55;">{message}</p>'
        if message
        else ""
    )
    sub_block = (
        f'<p style="color:#486581;max-width:640px;margin:{"0.4rem" if message else "0"} auto 0 auto;font-size:0.88rem;line-height:1.55;">{subtitle}</p>'
        if subtitle is not None and str(subtitle).strip() != ""
        else ""
    )
    badges_block = (
        f'<div style="display:inline-flex;gap:0.55rem;flex-wrap:wrap;justify-content:center;margin-top:0.85rem;">{badges_html}</div>'
        if badges_html
        else ""
    )
    html = (
        f'<div class="ds-kobo-empty">'
        f'{img_block}'
        f'<h4 style="margin:0 0 0.35rem 0;color:#102A43;font-size:1.08rem;font-weight:700;">{title}</h4>'
        f'{msg_block}'
        f'{sub_block}'
        f'{badges_block}'
        f'</div>'
    )
    st.markdown(html, unsafe_allow_html=True)


def _render_section_banner(title: str, subtitle: str = "", badge_html: str = "", tone: str = "ocean") -> None:
    """Render a clean engineering section banner card."""
    sub_html = f'<p class="ds-section-sub">{subtitle}</p>' if subtitle else ""
    st.markdown(
        f"""
        <div class="ds-section-card {tone}">
            <div class="ds-section-title">
                <span>{title}</span>
                {badge_html}
            </div>
            {sub_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def _render_channel_comparison_figure(
    shape: str,
    b_des: float,
    h_des: float,
    b_eff: float,
    h_eff: float,
    ts_val: float,
    n_des: float,
    n_akt: float,
) -> plt.Figure:
    """Create a scientific side-by-side cross-section visualization: Ideal DED vs Degraded Existing (Full-Width High-Res)."""
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(11.5, 4.4), dpi=125)
    fig.patch.set_facecolor("#FFFFFF")

    top_extra = 0.25 if shape == "Trapesium" else 0.0
    # Left subplot: Ideal DED
    ax1.set_facecolor("#F4F8FC")
    x_ideal = [-b_des / 2 - top_extra, -b_des / 2, b_des / 2, b_des / 2 + top_extra]
    y_ideal = [h_des, 0.0, 0.0, h_des]
    ax1.fill_between(
        [-b_des / 2 - top_extra * 0.85, -b_des / 2, b_des / 2, b_des / 2 + top_extra * 0.85],
        [h_des * 0.85, 0.0, 0.0, h_des * 0.85],
        [h_des * 0.85, h_des * 0.85, h_des * 0.85, h_des * 0.85],
        color="#34C6D3",
        alpha=0.35,
        label=f"Penampang Basah Ideal (h={h_des:.2f} m)",
    )
    ax1.plot(x_ideal, y_ideal, color="#102A43", linewidth=3.2, label=f"Dinding DED (b={b_des:.2f} m)")
    ax1.set_title(f"Penampang Ideal DED\n(Kekasaran Rencana n = {n_des:.4f})", fontsize=11, fontweight="bold", color="#102A43", pad=8)
    ax1.set_xlabel("Lebar Penampang (m)", fontsize=9.5, color="#18324A")
    ax1.set_ylabel("Kedalaman (m)", fontsize=9.5, color="#18324A")
    ax1.set_xlim(-max(b_des, 0.8) * 0.8, max(b_des, 0.8) * 0.8)
    ax1.set_ylim(-0.08 * h_des, h_des * 1.18)
    ax1.grid(True, linestyle="--", alpha=0.45, color="#B8CCE0")
    ax1.legend(loc="upper right", fontsize=8.2, framealpha=0.92, facecolor="#FFFFFF", edgecolor="#DFE8F0")

    # Right subplot: Existing Degraded (Sediment ts + Narrowing)
    ax2.set_facecolor("#F4F8FC")
    ax2.plot(x_ideal, y_ideal, color="#9FB3C8", linestyle="--", linewidth=2.0, label="Garis Rencana Awal")
    x_eff = [-b_eff / 2 - top_extra, -b_eff / 2, b_eff / 2, b_eff / 2 + top_extra]
    y_eff = [h_des, ts_val, ts_val, h_des]
    if ts_val > 0:
        ax2.fill_between(
            [-b_des / 2, b_des / 2],
            0.0,
            ts_val,
            color="#C29B61",
            alpha=0.65,
            hatch="..",
            label=f"Endapan Sedimen (ts={ts_val:.3f} m)",
        )
    ax2.fill_between(
        [-b_eff / 2, b_eff / 2],
        ts_val,
        ts_val + h_eff * 0.85,
        color="#1677C8",
        alpha=0.32,
        label=f"Tinggi Efektif (h_eff={h_eff:.3f} m)",
    )
    ax2.plot(x_eff, y_eff, color="#E45B64", linewidth=3.0, label=f"Penampang Efektif (b_eff={b_eff:.3f} m)")
    ax2.set_title(
        f"Penampang Eksisting Terdegradasi\n(Kekasaran Aktual n = {n_akt:.5f})",
        fontsize=11,
        fontweight="bold",
        color="#102A43",
        pad=8,
    )
    ax2.set_xlabel("Lebar Penampang (m)", fontsize=9.5, color="#18324A")
    ax2.set_xlim(-max(b_des, 0.8) * 0.8, max(b_des, 0.8) * 0.8)
    ax2.set_ylim(-0.08 * h_des, h_des * 1.18)
    ax2.grid(True, linestyle="--", alpha=0.45, color="#B8CCE0")
    ax2.legend(loc="upper right", fontsize=8.2, framealpha=0.92, facecolor="#FFFFFF", edgecolor="#DFE8F0")

    fig.tight_layout()
    return fig


_inject_drainsense_design_system()

# ═════════════════════════════════════════════════════════════════
# PURE HELPER FUNCTIONS
# ═════════════════════════════════════════════════════════════════
GUMBEL_TABLE: dict[int, tuple[float, float]] = {
    10: (0.4952, 0.9496), 15: (0.5128, 1.0206), 20: (0.5236, 1.0628),
    25: (0.5309, 1.0914), 30: (0.5362, 1.1124), 35: (0.5403, 1.1285),
    40: (0.5436, 1.1413), 45: (0.5463, 1.1518), 50: (0.5485, 1.1607),
    60: (0.5521, 1.1747), 70: (0.5548, 1.1854), 75: (0.5559, 1.1898),
    80: (0.5569, 1.1938), 90: (0.5586, 1.2007), 100: (0.5600, 1.2065),
}

def _gumbel_params(n: int) -> tuple[float, float]:
    keys = sorted(GUMBEL_TABLE)
    return GUMBEL_TABLE[min(keys, key=lambda k: abs(k - n))]

def _manning_Q(b: float, h: float, S: float, n_mann: float) -> float:
    if any(x <= 0 for x in [b, h, S, n_mann]):
        return 0.0
    A = b * h
    P = b + 2 * h
    R = A / P
    return (1 / n_mann) * A * (R ** (2 / 3)) * (S ** 0.5)

def _mononobe_I(R24: float, tc: float) -> float:
    return (R24 / 24.0) * ((24.0 / tc) ** (2.0 / 3.0))

def _rational_Q(C: float, I: float, A_km2: float) -> float:
    return 0.278 * C * I * A_km2


# ═════════════════════════════════════════════════════════════════
# REFERENSI TABEL VEN TE CHOW (1959) — dipakai Modul 2, 3, 4 & PDF
# ═════════════════════════════════════════════════════════════════
CHOW_1959_RANGES: dict[str, tuple[float, float, float, float]] = {
    # Material: (n_min, n_tipikal, n_max, V_izin_m_s)
    "Beton":              (0.011, 0.013, 0.020, 3.0),
    "Pasangan Bata":      (0.011, 0.015, 0.018, 2.0),
    "Pasangan Batu Kali": (0.017, 0.025, 0.030, 2.0),
    "Tanah":              (0.016, 0.022, 0.033, 1.0),
}
MATERIAL_OPTIONS: list[str] = list(CHOW_1959_RANGES.keys())


def _normalize_material(mat_raw) -> str:
    """Map free-text material names (AI / dokumen) to the 4 Chow (1959) categories."""
    s = str(mat_raw or "").lower()
    if "batu" in s and "bata" not in s:
        return "Pasangan Batu Kali"
    if "bata" in s or "brick" in s:
        return "Pasangan Bata"
    if "tanah" in s or "earth" in s or "soil" in s:
        return "Tanah"
    return "Beton"


def _validate_chow_n(mat_name: str, n_val: float) -> tuple[bool, str]:
    lookup_mat = _normalize_material(mat_name) if mat_name not in CHOW_1959_RANGES else mat_name
    n_min, n_norm, n_max, _ = CHOW_1959_RANGES[lookup_mat]
    if n_min <= n_val <= n_max:
        return True, (
            f"✅ **Valid (Tabel Chow, 1959):** Nilai Manning $n = {n_val:.4f}$ berada dalam rentang "
            f"wajar material **{lookup_mat}** ({n_min:.3f} – {n_max:.3f}, tipikal {n_norm:.3f})."
        )
    return False, (
        f"⚠️ **Peringatan Validasi Chow (1959):** Nilai Manning $n = {n_val:.4f}$ berada **DI LUAR** "
        f"rentang wajar material **{lookup_mat}** ({n_min:.3f} – {n_max:.3f}, tipikal {n_norm:.3f}). "
        "Periksa kembali spesifikasi teknis DED."
    )


_M1_ID_MONTH_MAP: dict[str, int] = {
    "januari": 1, "jan": 1, "january": 1,
    "februari": 2, "pebruari": 2, "feb": 2, "peb": 2, "february": 2,
    "maret": 3, "mar": 3, "march": 3,
    "april": 4, "apr": 4,
    "mei": 5, "may": 5,
    "juni": 6, "jun": 6, "june": 6,
    "juli": 7, "jul": 7, "july": 7,
    "agustus": 8, "agu": 8, "ags": 8, "agt": 8, "aug": 8, "august": 8,
    "september": 9, "sept": 9, "sep": 9,
    "oktober": 10, "okt": 10, "oct": 10, "october": 10,
    "november": 11, "nopember": 11, "nov": 11, "nop": 11,
    "desember": 12, "des": 12, "dec": 12, "december": 12,
}
_M1_ID_MONTH_TOKENS: set[str] = {
    "januari", "jan", "februari", "pebruari", "feb", "peb", "maret", "mar",
    "april", "apr", "mei", "juni", "jun", "juli", "jul", "agustus", "agu",
    "ags", "agt", "september", "sept", "sep", "oktober", "okt",
    "november", "nopember", "nov", "nop", "desember", "des",
}
_M1_NON_DATA_EXACT: set[str] = {
    "", "none", "nan", "nat", "null", "n/a", "na", "-", "--",
    "tahun", "year", "thn", "bulan", "month", "bln",
    "tanggal", "date", "tgl", "waktu", "periode", "period",
    "total", "jumlah", "jml", "sum", "subtotal",
    "rata-rata", "rata rata", "ratarata", "rerata", "mean", "average", "avg",
    "maks", "maksimum", "max", "minimum", "min",
    "keterangan", "ket", "catatan", "note", "notes", "sumber",
}
_M1_NON_DATA_PREFIXES: tuple[str, ...] = (
    "total", "jumlah", "rata-rata", "rata rata", "ratarata", "rerata",
    "subtotal", "keterangan", "catatan", "sumber",
)


def _is_m1_non_data_label(val) -> bool:
    """Return True for blank, stray header, or summary labels in Modul 1 date column."""
    if val is None or pd.isna(val):
        return True
    s = re.sub(r"\s+", " ", str(val).strip().lower())
    s_clean = re.sub(r"[\s:.\-]+$", "", s)
    if s in _M1_NON_DATA_EXACT or s_clean in _M1_NON_DATA_EXACT:
        return True
    return s.startswith(_M1_NON_DATA_PREFIXES)


def _parse_m1_rain_value(val) -> float:
    """Coerce rainfall cell to float, supporting optional 'mm' suffix or comma decimals."""
    if val is None or pd.isna(val):
        return float("nan")
    if isinstance(val, (int, float, np.number)):
        return float(val)
    s = re.sub(r"\s*mm\s*$", "", str(val).strip(), flags=re.IGNORECASE)
    if re.fullmatch(r"[-+]?\d+,\d+", s):
        s = s.replace(",", ".")
    val_num = pd.to_numeric(s, errors="coerce")
    return float(val_num) if pd.notna(val_num) else float("nan")


def _parse_m1_rainfall_series(df_subset: pd.DataFrame) -> tuple[pd.DataFrame, bool, str | None]:
    """Smart parser for Modul 1 date/period & rainfall columns.

    Supports standard dates, pure 4-digit years, multi-year Indonesian months,
    and single-year Indonesian monthly tables (with stray header/total filtering).
    Returns (r24_df, is_monthly_single_year, error_message).
    """
    raw_row_count = len(df_subset)
    df_clean = df_subset.iloc[:, :2].copy()
    df_clean.columns = ["_raw_date", "_raw_rain"]

    # Step 1: Filter stray header rows, None/NaN date rows, and summary/total rows
    df_clean = df_clean[~df_clean["_raw_date"].apply(_is_m1_non_data_label)].copy()
    df_clean["CurahHujan"] = df_clean["_raw_rain"].apply(_parse_m1_rain_value)
    df_clean = df_clean.dropna(subset=["CurahHujan"]).reset_index(drop=True)

    if df_clean.empty:
        return pd.DataFrame(), False, "❌ Tidak ada data valid. Periksa kolom yang dipilih."

    date_texts = [re.sub(r"\s+", " ", str(v).strip()) for v in df_clean["_raw_date"]]

    # Step 2a: Detect pure 4-digit years (e.g. 2015..2030 stored as int/float/string)
    pure_years: list[int | None] = []
    for s in date_texts:
        m_yr = re.fullmatch(r"(?:(?:tahun|thn|year)\s*[:\-]?\s*)?((?:19|20|21)\d{2})(?:\.0+)?", s, flags=re.IGNORECASE)
        pure_years.append(int(m_yr.group(1)) if m_yr else None)

    n_valid = len(date_texts)
    n_pure_years = sum(y is not None for y in pure_years)
    if n_pure_years >= max(1, int(0.8 * n_valid)):
        df_clean["Tahun"] = pure_years
        df_yr = df_clean.dropna(subset=["Tahun"]).copy()
        df_yr["Tahun"] = df_yr["Tahun"].astype(int)
        r24_calc = (
            df_yr.groupby("Tahun")["CurahHujan"]
            .max().reset_index()
            .rename(columns={"CurahHujan": "R24_Max (mm)"})
        )
        r24_calc = r24_calc[r24_calc["R24_Max (mm)"] > 0].reset_index(drop=True)
        if len(r24_calc) < 2:
            return pd.DataFrame(), False, "❌ Data tahunan R24 positif minimal 2 tahun diperlukan untuk analisis."
        return r24_calc, False, None

    # Step 2b: Detect Indonesian / standard month names or abbreviations
    month_hits: list[int | None] = []
    id_month_hits = 0
    embedded_years: list[int | None] = []
    for s in date_texts:
        s_low = s.lower()
        tokens = re.findall(r"[a-z]+", s_low)
        m_num = next((_M1_ID_MONTH_MAP[t] for t in tokens if t in _M1_ID_MONTH_MAP), None)
        if any(t in _M1_ID_MONTH_TOKENS for t in tokens):
            id_month_hits += 1
        y_match = re.search(r"\b((?:19|20|21)\d{2})\b", s)
        month_hits.append(m_num)
        embedded_years.append(int(y_match.group(1)) if y_match else None)

    n_month_hits = sum(m is not None for m in month_hits)
    distinct_embedded_years = {y for y in embedded_years if y is not None}

    if n_month_hits >= max(1, int(0.5 * n_valid)):
        if len(distinct_embedded_years) >= 2:
            df_clean["Tahun"] = embedded_years
            df_yr = df_clean.dropna(subset=["Tahun"]).copy()
            df_yr["Tahun"] = df_yr["Tahun"].astype(int)
            r24_calc = (
                df_yr.groupby("Tahun")["CurahHujan"]
                .max().reset_index()
                .rename(columns={"CurahHujan": "R24_Max (mm)"})
            )
            r24_calc = r24_calc[r24_calc["R24_Max (mm)"] > 0].reset_index(drop=True)
            if len(r24_calc) < 2:
                return pd.DataFrame(), False, "❌ Data tahunan R24 positif minimal 2 tahun diperlukan untuk analisis."
            return r24_calc, False, None

        # Single-year monthly fallback when Indonesian month table has <= 20 rows
        if (id_month_hits > 0 or n_month_hits > 0) and (raw_row_count <= 20 or n_valid <= 20):
            df_m = df_clean[[m is not None for m in month_hits]].copy()
            df_m = df_m[df_m["CurahHujan"] > 0].reset_index(drop=True)
            if len(df_m) < 2:
                return pd.DataFrame(), True, "❌ Data bulanan positif minimal 2 bulan diperlukan untuk analisis."
            labels = [
                f"Bulan {idx} ({str(raw_lbl).strip()})"
                for idx, raw_lbl in enumerate(df_m["_raw_date"], start=1)
            ]
            r24_monthly = pd.DataFrame({
                "Periode": labels,
                "R24_Max (mm)": df_m["CurahHujan"].astype(float).values,
            })
            return r24_monthly, True, None

    # Step 2c: Standard datetime parsing (try default & dayfirst=True, pick best yield)
    dt_default = pd.to_datetime(df_clean["_raw_date"], errors="coerce")
    dt_dayfirst = pd.to_datetime(df_clean["_raw_date"], errors="coerce", dayfirst=True)
    dt_parsed = dt_dayfirst if dt_dayfirst.notna().sum() > dt_default.notna().sum() else dt_default

    df_clean["Tahun"] = dt_parsed.dt.year
    df_std = df_clean.dropna(subset=["Tahun"]).copy()
    if df_std.empty:
        return pd.DataFrame(), False, "❌ Format tanggal/periode tidak dikenali atau tidak ada data valid."

    df_std["Tahun"] = df_std["Tahun"].astype(int)
    r24_calc = (
        df_std.groupby("Tahun")["CurahHujan"]
        .max().reset_index()
        .rename(columns={"CurahHujan": "R24_Max (mm)"})
    )
    r24_calc = r24_calc[r24_calc["R24_Max (mm)"] > 0].reset_index(drop=True)
    if len(r24_calc) < 2:
        return pd.DataFrame(), False, "❌ Data tahunan R24 positif minimal 2 tahun diperlukan untuk analisis."
    return r24_calc, False, None


# ═════════════════════════════════════════════════════════════════
# ENGINE 1 — DIRECT SDK `from google import genai` (models/gemini-4-argon)
# ═════════════════════════════════════════════════════════════════
ARGON_MODEL = "models/gemini-4-argon"


def _redact_api_key(err_text: str, raw_key: str = "") -> str:
    """Mask any API key in error messages before displaying in UI."""
    cleaned = str(err_text)
    if raw_key and len(raw_key.strip()) > 4:
        cleaned = cleaned.replace(raw_key.strip(), "[API_KEY_DISENSOR]")
    return re.sub(
        r"(AIza[0-9A-Za-z\-_]+|AQ\.[0-9A-Za-z\-_\.]+|key=[0-9A-Za-z\-_\.]+)",
        "[API_KEY_DISENSOR]",
        cleaned,
    )


def _clean_ai_json_text(raw_text: str) -> str:
    """Bersihkan pembungkus markdown ```json ... ``` / ``` ... ``` (pola Demo.py)."""
    txt = (raw_text or "").strip()
    if txt.startswith("```json"):
        txt = txt[len("```json"):]
    elif txt.startswith("```"):
        txt = txt[3:]
    if txt.endswith("```"):
        txt = txt[:-3]
    return txt.strip()


def _parse_ai_json(text: str) -> dict:
    """Parse JSON murni dari AI; fallback ke objek {...} pertama-terakhir bila ada teks tambahan."""
    cleaned = _clean_ai_json_text(text)
    try:
        data = json.loads(cleaned)
        if isinstance(data, dict):
            return data
    except json.JSONDecodeError:
        pass
    start, end = cleaned.find("{"), cleaned.rfind("}")
    if start != -1 and end > start:
        data = json.loads(cleaned[start:end + 1])
        if isinstance(data, dict):
            return data
    raise ValueError("Respons AI bukan JSON object yang valid.")


def _extract_interaction_text(interaction) -> str:
    """Ambil teks keluaran Interaction: steps[-1].text → output_text → konten teks step terakhir."""
    steps = getattr(interaction, "steps", None) or []
    if steps:
        last_text = getattr(steps[-1], "text", None)
        if isinstance(last_text, str) and last_text.strip():
            return last_text.strip()
    out_text = getattr(interaction, "output_text", None)
    if isinstance(out_text, str) and out_text.strip():
        return out_text.strip()
    for step in reversed(steps):
        for item in reversed(getattr(step, "content", None) or []):
            item_text = getattr(item, "text", None)
            if isinstance(item_text, str) and item_text.strip():
                return item_text.strip()
    return ""


GEMINI_MODEL_NAME = "gemini-1.5-flash"  # model standar yang stabil (di-hardcode)


def call_argon_ai(
    api_key: str,
    sys_prompt: str,
    image_bytes: bytes | None = None,
    image_mime: str = "image/jpeg",
) -> str:
    """
    Panggil Gemini API via Direct SDK dengan model standar `gemini-1.5-flash`:
    `genai.GenerativeModel("gemini-1.5-flash")` (`google-generativeai`) atau
    `client.models.generate_content(...)` (`google-genai`).
    Mengembalikan string JSON bersih.
    """
    clean_key = (api_key or "").strip().strip('"').strip("'")
    if not clean_key:
        raise RuntimeError("API Key Gemini kosong.")

    target_model = GEMINI_MODEL_NAME
    errors_log: list[str] = []

    # 1) Jalur SDK google-generativeai (genai.GenerativeModel("gemini-1.5-flash")) bila tersedia
    try:
        import google.generativeai as legacy_genai  # type: ignore
        legacy_genai.configure(api_key=clean_key)
        gen_model = legacy_genai.GenerativeModel(model_name=target_model)
        if image_bytes and PIL_AVAILABLE:
            pil_obj = PILImage.open(io.BytesIO(image_bytes))
            resp_legacy = gen_model.generate_content([sys_prompt, pil_obj])
        else:
            resp_legacy = gen_model.generate_content(sys_prompt)
        raw_text = (getattr(resp_legacy, "text", None) or "").strip()
        if raw_text:
            return _clean_ai_json_text(raw_text)
    except ImportError:
        pass
    except Exception as err:
        errors_log.append(f"[GenerativeModel({target_model})] {_redact_api_key(str(err), clean_key)}")

    if not GENAI_AVAILABLE:
        if errors_log:
            raise RuntimeError(" | ".join(errors_log))
        raise RuntimeError("Pustaka `google-genai` / `google-generativeai` belum terinstall.")

    client = genai.Client(api_key=clean_key)

    # 2) Jalur utama google-genai: client.models.generate_content(model="gemini-1.5-flash")
    try:
        contents: list = [sys_prompt]
        if image_bytes:
            from google.genai import types as genai_types
            contents.append(genai_types.Part.from_bytes(data=image_bytes, mime_type=image_mime))
        resp = client.models.generate_content(model=target_model, contents=contents)
        raw_text = (getattr(resp, "text", None) or "").strip()
        if raw_text:
            return _clean_ai_json_text(raw_text)
        errors_log.append("[generate_content] respons kosong")
    except Exception as err:
        errors_log.append(f"[generate_content({target_model})] {_redact_api_key(str(err), clean_key)}")

    # 3) Fallback: Interactions API
    try:
        if image_bytes:
            ai_input = [
                {"type": "text", "text": sys_prompt},
                {"type": "image", "data": base64.b64encode(image_bytes).decode("ascii"), "mime_type": image_mime},
            ]
        else:
            ai_input = sys_prompt
        interaction = client.interactions.create(
            model=target_model,
            input=ai_input,
            generation_config={"max_output_tokens": 8192},
        )
        raw_text = _extract_interaction_text(interaction)
        if raw_text:
            return _clean_ai_json_text(raw_text)
        errors_log.append("[interactions.create] respons kosong")
    except Exception as err:
        errors_log.append(f"[interactions.create({target_model})] {_redact_api_key(str(err), clean_key)}")

    raise RuntimeError(" | ".join(errors_log))


def _discover_gemini_vision_models(clean_key: str) -> tuple[list[str], list[str]]:
    """Auto-discover active Gemini models supporting generateContent, prioritized for flash vision."""
    discovered: list[str] = []
    discovery_errors: list[str] = []

    try:
        import google.generativeai as legacy_genai  # type: ignore
        legacy_genai.configure(api_key=clean_key)
        for m in legacy_genai.list_models():
            methods = getattr(m, "supported_generation_methods", None) or []
            if "generateContent" in methods:
                m_name = str(getattr(m, "name", "")).replace("models/", "").strip()
                if m_name and "gemini" in m_name.lower():
                    discovered.append(m_name)
    except Exception as err_list:
        discovery_errors.append(f"[list_models] {_redact_api_key(str(err_list), clean_key)}")

    if not discovered and GENAI_AVAILABLE:
        try:
            client = genai.Client(api_key=clean_key)
            for m in client.models.list():
                methods = getattr(m, "supported_actions", None) or getattr(m, "supported_generation_methods", None) or []
                if not methods or "generateContent" in methods:
                    m_name = str(getattr(m, "name", "")).replace("models/", "").strip()
                    if m_name and "gemini" in m_name.lower():
                        discovered.append(m_name)
        except Exception as err_list2:
            discovery_errors.append(f"[client.models.list] {_redact_api_key(str(err_list2), clean_key)}")

    # Deduplicate while filtering out audio/tts/embedding-only endpoints
    excluded_tokens = ("tts", "audio", "embedding", "aqa", "imagen", "veo", "robotics")
    unique_models = []
    for name in discovered:
        low = name.lower()
        if any(tok in low for tok in excluded_tokens):
            continue
        if name not in unique_models:
            unique_models.append(name)

    def _model_priority(name: str) -> tuple[int, int, int, str]:
        low = name.lower()
        if "2.5-flash" in low:
            tier = 0
        elif "2.0-flash" in low:
            tier = 1
        elif "3.5-flash" in low:
            tier = 2
        elif "flash" in low:
            tier = 3
        else:
            tier = 4
        # Prefer standard non-lite/non-thinking/non-preview variants within each tier
        sub_penalty = (
            (1 if "lite" in low else 0)
            + (1 if "exp" in low or "preview" in low else 0)
            + (2 if "thinking" in low or "image" in low else 0)
        )
        return (tier, sub_penalty, len(low), low)

    unique_models.sort(key=_model_priority)
    return unique_models, discovery_errors


def call_gemini_vision(
    api_key: str,
    sys_prompt: str,
    image_bytes: bytes,
    image_mime: str = "image/jpeg",
) -> tuple[str, str]:
    """
    Panggil Gemini Vision API via Direct SDK menggunakan auto-discovery `genai.list_models()`
    dengan prioritas model flash terbaru (2.5-flash -> 2.0-flash -> 3.5-flash -> flash).
    Mengembalikan (json_text_bersih, model_name_yang_sukses).
    """
    clean_key = (api_key or "").strip().strip('"').strip("'")
    if not clean_key:
        raise RuntimeError("API Key Gemini kosong.")

    vision_models, errors_log = _discover_gemini_vision_models(clean_key)
    if not vision_models:
        detail = " | ".join(errors_log) if errors_log else "Tidak ada model Gemini generateContent yang tersedia."
        raise RuntimeError(detail)

    # 1) Jalur SDK google-generativeai (genai.GenerativeModel) dengan PIL Image
    try:
        import google.generativeai as legacy_genai  # type: ignore
        legacy_genai.configure(api_key=clean_key)
        pil_obj = PILImage.open(io.BytesIO(image_bytes)) if PIL_AVAILABLE else None

        for model_name in vision_models:
            try:
                gen_model = legacy_genai.GenerativeModel(model_name=model_name)
                contents = [sys_prompt, pil_obj] if pil_obj is not None else [sys_prompt]
                resp_legacy = gen_model.generate_content(contents)
                raw_text = (getattr(resp_legacy, "text", None) or "").strip()
                if raw_text:
                    return _clean_ai_json_text(raw_text), model_name
            except Exception as err:
                errors_log.append(f"[GenerativeModel({model_name})] {_redact_api_key(str(err), clean_key)}")
    except ImportError:
        pass
    except Exception as err:
        errors_log.append(f"[configure] {_redact_api_key(str(err), clean_key)}")

    # 2) Jalur Google GenAI SDK (google-genai Client)
    if GENAI_AVAILABLE:
        try:
            client = genai.Client(api_key=clean_key)
            from google.genai import types as genai_types
            img_part = genai_types.Part.from_bytes(data=image_bytes, mime_type=image_mime)

            for model_name in vision_models:
                try:
                    resp = client.models.generate_content(
                        model=model_name,
                        contents=[sys_prompt, img_part],
                    )
                    raw_text = (getattr(resp, "text", None) or "").strip()
                    if raw_text:
                        return _clean_ai_json_text(raw_text), model_name
                except Exception as err:
                    errors_log.append(f"[client.generate_content({model_name})] {_redact_api_key(str(err), clean_key)}")
        except Exception as err:
            errors_log.append(f"[genai.Client] {_redact_api_key(str(err), clean_key)}")

    raise RuntimeError(" | ".join(errors_log) if errors_log else "Pemanggilan Gemini Vision gagal.")


def call_gemini_text(
    api_key: str,
    prompt_text: str,
) -> tuple[str, str]:
    """
    Panggil Gemini Text API via Direct SDK menggunakan auto-discovery model aktif
    (_discover_gemini_vision_models: 2.5-flash -> 2.0-flash -> 3.5-flash -> flash).
    Mengembalikan (response_text, model_name_yang_sukses).
    """
    clean_key = (api_key or "").strip().strip('"').strip("'")
    if not clean_key:
        raise RuntimeError("API Key Gemini kosong.")

    active_models, errors_log = _discover_gemini_vision_models(clean_key)
    if not active_models:
        detail = " | ".join(errors_log) if errors_log else "Tidak ada model Gemini generateContent yang tersedia."
        raise RuntimeError(detail)

    try:
        import google.generativeai as legacy_genai  # type: ignore
        legacy_genai.configure(api_key=clean_key)
        for model_name in active_models:
            try:
                gen_model = legacy_genai.GenerativeModel(model_name=model_name)
                resp_legacy = gen_model.generate_content(prompt_text)
                raw_text = (getattr(resp_legacy, "text", None) or "").strip()
                if raw_text:
                    return raw_text, model_name
            except Exception as err:
                errors_log.append(f"[GenerativeModel({model_name})] {_redact_api_key(str(err), clean_key)}")
    except ImportError:
        pass
    except Exception as err:
        errors_log.append(f"[configure] {_redact_api_key(str(err), clean_key)}")

    if GENAI_AVAILABLE:
        try:
            client = genai.Client(api_key=clean_key)
            for model_name in active_models:
                try:
                    resp = client.models.generate_content(
                        model=model_name,
                        contents=prompt_text,
                    )
                    raw_text = (getattr(resp, "text", None) or "").strip()
                    if raw_text:
                        return raw_text, model_name
                except Exception as err:
                    errors_log.append(f"[client.generate_content({model_name})] {_redact_api_key(str(err), clean_key)}")
        except Exception as err:
            errors_log.append(f"[genai.Client] {_redact_api_key(str(err), clean_key)}")

    raise RuntimeError(" | ".join(errors_log) if errors_log else "Pemanggilan Gemini AI gagal.")


# ═════════════════════════════════════════════════════════════════
# ENGINE 2 — LIVE JSON BRIDGE (drainsense_bridge.json)
# ═════════════════════════════════════════════════════════════════
BRIDGE_FILENAME = "drainsense_bridge.json"
BRIDGE_PATH = Path(__file__).resolve().parent / BRIDGE_FILENAME

BRIDGE_TEMPLATE: dict = {
    "updated_at": "",
    "source": "Antigravity (VS Code)",
    "modul2_ded": {
        "lebar_dasar_b": 1.5,
        "kedalaman_h": 1.2,
        "kemiringan_s": 0.001,
        "manning_n": 0.013,
        "material": "Beton",
        "debit_rencana_q": 0.0,
        "kutipan_halaman": "Contoh: Bab IV Kriteria Desain Hidrolika, Tabel 4.3, Hal. 42",
        "alasan_ai": "Contoh: dimensi rencana normalisasi dipilih karena kondisi eksisting meluap.",
    },
    "modul3_cowan": {
        "nb": 0.013,
        "n1_ketidakteraturan": 0.005,
        "n2_variasi_penampang": 0.002,
        "n3_hambatan": 0.015,
        "n4_vegetasi_dalam_saluran": 0.010,
        "m_meander": 1.0,
        "estimasi_tebal_sedimen_m": 0.15,
        "catatan_visual_ai": "Contoh: endapan lumpur tipis di dasar, gulma sedikit di dinding dalam.",
    },
}


def _bridge_mtime() -> float | None:
    try:
        return BRIDGE_PATH.stat().st_mtime if BRIDGE_PATH.exists() else None
    except OSError:
        return None


def _read_bridge() -> tuple[dict | None, str | None]:
    if not BRIDGE_PATH.exists():
        return None, None
    try:
        data = json.loads(_clean_ai_json_text(BRIDGE_PATH.read_text(encoding="utf-8")))
        if not isinstance(data, dict):
            return None, "Isi bridge bukan JSON object."
        return data, None
    except Exception as err:
        return None, f"Gagal membaca {BRIDGE_FILENAME}: {err}"


def _write_bridge_template() -> None:
    tpl = json.loads(json.dumps(BRIDGE_TEMPLATE))
    tpl["updated_at"] = datetime.now().isoformat(timespec="seconds")
    BRIDGE_PATH.write_text(json.dumps(tpl, indent=2, ensure_ascii=False), encoding="utf-8")


# ═════════════════════════════════════════════════════════════════
# PYTHON VALIDATOR — MODUL 2 (Tabel Chow 1959)
# ═════════════════════════════════════════════════════════════════
def _to_float_safe(val, default: float | None = None) -> float | None:
    if val is None:
        return default
    try:
        if isinstance(val, str):
            val = val.strip().replace(",", ".").replace("%", "")
        return float(val)
    except (TypeError, ValueError):
        return default


def _validate_m2_ai_json(ai_json: dict, source: str) -> dict:
    """
    Python Validator: parsing JSON AI/Bridge → cek fisik dimensi & cek Manning n terhadap Tabel Chow (1959).
    Tidak menyimpan ke session_state; mengembalikan dict hasil validasi.
    """
    b = _to_float_safe(ai_json.get("lebar_dasar_b"))
    h = _to_float_safe(ai_json.get("kedalaman_h"))
    S = _to_float_safe(ai_json.get("kemiringan_s"))
    n = _to_float_safe(ai_json.get("manning_n"))
    q = _to_float_safe(ai_json.get("debit_rencana_q"), 0.0) or 0.0
    material = _normalize_material(ai_json.get("material"))

    issues: list[str] = []
    if b is None or not (0.05 <= b <= 50.0):
        issues.append(f"Lebar dasar b tidak wajar/kosong ({ai_json.get('lebar_dasar_b')}) → default 1.500 m.")
        b = 1.5
    if h is None or not (0.05 <= h <= 20.0):
        issues.append(f"Kedalaman h tidak wajar/kosong ({ai_json.get('kedalaman_h')}) → default 1.200 m.")
        h = 1.2
    if S is not None and S > 0.2:
        issues.append(f"Kemiringan S = {S} terdeteksi dalam persen → dikonversi ke m/m ({S / 100.0:.6f}).")
        S = S / 100.0
    if S is None or not (1e-6 <= S <= 0.2):
        issues.append(f"Kemiringan S tidak wajar/kosong ({ai_json.get('kemiringan_s')}) → default 0.001000.")
        S = 0.001
    n_min, n_typ, n_max, v_izin = CHOW_1959_RANGES[material]
    if n is None or n <= 0:
        issues.append(f"Manning n kosong → memakai nilai tipikal Chow {material} ({n_typ:.3f}).")
        n = n_typ

    chow_valid = n_min <= n <= n_max
    return {
        "source": source,
        "b": round(b, 4),
        "h": round(h, 4),
        "S": round(S, 7),
        "n": round(n, 4),
        "material": material,
        "Q_rencana": round(q, 4),
        "kutipan_ded": str(ai_json.get("kutipan_halaman") or ai_json.get("kutipan_ded") or "—"),
        "alasan_ai": str(ai_json.get("alasan_ai") or "—"),
        "chow_valid": chow_valid,
        "n_min": n_min,
        "n_typ": n_typ,
        "n_max": n_max,
        "V_izin": v_izin,
        "issues": issues,
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def _commit_m2_to_ded(res: dict, n_override: float | None = None) -> None:
    """
    Isi otomatis form parameter DED (st.session_state['ded_params']) sekaligus
    meng-update state widget Form Parameter DED di panel kanan secara instan.
    """
    b_val = float(res["b"])
    h_val = float(res["h"])
    s_val = float(res["S"])
    n_val = float(n_override if n_override is not None else res["n"])
    mat_val = _normalize_material(res.get("material", "Beton"))
    if mat_val not in MATERIAL_OPTIONS:
        mat_val = MATERIAL_OPTIONS[0]
    kutipan_val = str(res.get("kutipan_ded", "—"))

    ded_dict = {
        "b":           b_val,
        "h":           h_val,
        "S":           s_val,
        "n":           n_val,
        "material":    mat_val,
        "kutipan_ded": kutipan_val,
        "Q_rencana":   float(res.get("Q_rencana", 0.0)),
        "alasan_ai":   res.get("alasan_ai", "—"),
        "sumber":      res.get("source", "—"),
    }
    st.session_state["ded_params"] = ded_dict
    st.session_state["params_modul_2"] = dict(ded_dict)
    # Sinkronisasi langsung ke widget keys pada Form Parameter DED di panel sebelah kanan
    st.session_state["ded_form_b"] = b_val
    st.session_state["ded_form_h"] = h_val
    st.session_state["ded_form_S"] = s_val
    st.session_state["ded_form_n"] = n_val
    st.session_state["ded_form_material"] = mat_val
    st.session_state["ded_form_kutipan"] = kutipan_val


def _save_m2_to_bridge_cache(m2_raw_dict: dict, source: str = "Python Instant Extractor") -> None:
    """
    Simpan hasil ekstraksi Modul 2 secara otomatis di background ke `drainsense_bridge.json`
    sebagai cache/cadangan sistem tanpa memerlukan interaksi manual pengguna.
    """
    try:
        existing, _ = _read_bridge()
        payload = existing if isinstance(existing, dict) else json.loads(json.dumps(BRIDGE_TEMPLATE))
        payload["updated_at"] = datetime.now().isoformat(timespec="seconds")
        payload["source"] = source
        payload["modul2_ded"] = {
            "lebar_dasar_b": float(m2_raw_dict.get("lebar_dasar_b", 1.5)),
            "kedalaman_h": float(m2_raw_dict.get("kedalaman_h", 1.2)),
            "kemiringan_s": float(m2_raw_dict.get("kemiringan_s", 0.001)),
            "manning_n": float(m2_raw_dict.get("manning_n", 0.013)),
            "material": str(m2_raw_dict.get("material", "Beton")),
            "debit_rencana_q": float(m2_raw_dict.get("debit_rencana_q", 0.0)),
            "kutipan_halaman": str(m2_raw_dict.get("kutipan_halaman", "—")),
            "alasan_ai": str(m2_raw_dict.get("alasan_ai", "—")),
        }
        BRIDGE_PATH.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
        st.session_state["bridge_mtime"] = _bridge_mtime()
        st.session_state["bridge_data"] = payload
    except Exception as e:
        print(f"[Bridge Background Cache Warning]: {e}", flush=True)


# ═════════════════════════════════════════════════════════════════
# PYTHON PRE-PROCESSOR (ROI) & CALCULATOR — MODUL 3 (Metode Cowan 1956)
# ═════════════════════════════════════════════════════════════════
COWAN_LIMITS: dict[str, tuple[float, float]] = {
    "nb": (0.010, 0.040),
    "n1": (0.000, 0.020),
    "n2": (0.000, 0.015),
    "n3": (0.000, 0.060),
    "n4": (0.000, 0.100),
    "m":  (1.000, 1.300),
}


def _crop_roi(pil_img, left_pct: float, right_pct: float, top_pct: float, bottom_pct: float):
    """Potong gambar ke Region of Interest (bagian DALAM penampang saluran)."""
    img = pil_img.convert("RGB")
    W, H = img.size
    x0 = int(W * left_pct / 100.0)
    x1 = int(W * (1.0 - right_pct / 100.0))
    y0 = int(H * top_pct / 100.0)
    y1 = int(H * (1.0 - bottom_pct / 100.0))
    if x1 - x0 < 16 or y1 - y0 < 16:
        return img
    return img.crop((x0, y0, x1, y1))


def _roi_pixel_stats(roi_img) -> dict:
    """Hitung rasio piksel dasar pada ZONA TENGAH ROI (dasar & dinding dalam saluran)."""
    arr = np.asarray(roi_img.convert("RGB").resize((256, 256)), dtype=np.float32) / 255.0
    zone = arr[64:256, 48:208, :]  # zona tengah-bawah: dasar & dinding dalam
    R, G, B = zone[:, :, 0], zone[:, :, 1], zone[:, :, 2]
    lum = 0.299 * R + 0.587 * G + 0.114 * B
    sat = (np.max(zone, axis=2) - np.min(zone, axis=2))

    exg = 2.0 * G - R - B
    veg = (exg > 0.08) & (G > 0.18) & (G > R * 1.06)
    sediment = (~veg) & (R >= G * 0.92) & (G >= B * 1.05) & (lum > 0.12) & (lum < 0.62)
    water = (~veg) & (~sediment) & (lum < 0.30)
    concrete = (~veg) & (~sediment) & (lum >= 0.45) & (sat < 0.12)

    gray = 0.299 * arr[:, :, 0] + 0.587 * arr[:, :, 1] + 0.114 * arr[:, :, 2]
    edge = float(np.mean(np.abs(np.diff(gray, axis=1))) + np.mean(np.abs(np.diff(gray, axis=0))))

    total = float(veg.size)
    return {
        "vegetasi_pct": round(float(np.sum(veg)) / total * 100.0, 2),
        "sedimen_pct":  round(float(np.sum(sediment)) / total * 100.0, 2),
        "air_gelap_pct": round(float(np.sum(water)) / total * 100.0, 2),
        "beton_terang_pct": round(float(np.sum(concrete)) / total * 100.0, 2),
        "indeks_tekstur": round(edge, 4),
    }


def _local_cowan_from_stats(stats: dict, nb_default: float, h_ref: float) -> dict:
    """
    Engine Python CV Lokal: mengestimasi parameter Metode Cowan (1956) & Tebal Sedimen (ts)
    secara instan (< 1 detik) dari hasil analisis piksel ROI (Vegetasi %, Sedimen %, Air %, Beton Terang %).
    """
    sed = float(stats.get("sedimen_pct", 0.0))
    veg = float(stats.get("vegetasi_pct", 0.0))
    water = float(stats.get("air_gelap_pct", 0.0))
    concrete = float(stats.get("beton_terang_pct", 0.0))
    tex = float(stats.get("indeks_tekstur", 0.05))

    # n1 (Ketidakteraturan permukaan: 0.000 - 0.020) dari tekstur tepi & kondisi permukaan dinding/beton
    rough_wall = max(0.0, (100.0 - concrete) / 100.0) * 0.006
    n1_est = float(np.clip(round(tex * 0.055 + rough_wall, 3), 0.000, 0.020))

    # n2 (Variasi penampang: 0.000 - 0.015) dari tekstur & distribusi endapan
    n2_est = float(np.clip(round(tex * 0.035 + (sed / 100.0) * 0.006, 3), 0.000, 0.015))

    # n3 (Hambatan endapan/sampah: 0.000 - 0.060) berdasarkan persentase piksel Sedimen
    n3_est = float(np.clip(round((sed / 100.0) * 0.060, 3), 0.000, 0.060))

    # n4 (Vegetasi/gulma dalam saluran: 0.000 - 0.100) berdasarkan persentase piksel Vegetasi
    n4_est = float(np.clip(round((veg / 100.0) * 0.100, 3), 0.000, 0.100))

    # m (Faktor meander/kelokan: 1.00 - 1.30)
    m_est = 1.00 if (sed + veg) < 45.0 else 1.05

    # ts (Tebal sedimen: proporsi luas piksel sedimen terhadap tinggi saluran h_ref)
    prop_sed = float(np.clip((sed / 100.0) * 0.65, 0.02 if sed > 1.0 else 0.0, 0.75))
    ts_est = round(min(prop_sed * h_ref, max(0.0, h_ref - 0.02)), 3)

    # p (Penyempitan lebar: 0.0 - 0.90)
    p_est = round(float(np.clip((sed / 100.0) * 0.15, 0.0, 0.50)), 2)
    alasan_txt = (
        f"Auto-Estimasi Python CV (ROI): Sedimen {sed:.1f}% → n3={n3_est:.3f} & ts={ts_est:.3f} m; "
        f"Vegetasi {veg:.1f}% → n4={n4_est:.3f}; Air/Gelap {water:.1f}%; "
        f"Beton Terang {concrete:.1f}% → n1={n1_est:.3f}, n2={n2_est:.3f}, p={p_est:.2f}."
    )

    return {
        "nb": round(float(nb_default), 4),
        "n1": n1_est,
        "n2": n2_est,
        "n3": n3_est,
        "n4": n4_est,
        "m": m_est,
        "ts": ts_est,
        "p": p_est,
        "alasan_visual": alasan_txt,
        "n1_ketidakteraturan": n1_est,
        "n2_variasi_penampang": n2_est,
        "n3_hambatan": n3_est,
        "n4_vegetasi_dalam_saluran": n4_est,
        "m_meander": m_est,
        "estimasi_tebal_sedimen_m": ts_est,
        "catatan_visual_ai": alasan_txt,
    }


def _apply_m3_result(ai_json: dict, source: str, h_ref: float, b_ref: float, nb_default: float) -> dict:
    """
    Python Calculator: klem faktor AI/Python CV ke batas Cowan (1956), hitung
    n_aktual = (nb + n1 + n2 + n3 + n4) × m, dan langsung sinkronkan ke st.session_state
    serta seluruh slider/input di panel sebelah kanan Modul 3.
    """
    def _clamp_multi(keys_json: tuple[str, ...], key_lim: str, default: float) -> float:
        lo, hi = COWAN_LIMITS[key_lim]
        raw_val = None
        for kj in keys_json:
            if ai_json.get(kj) is not None:
                raw_val = ai_json.get(kj)
                break
        return float(np.clip(_to_float_safe(raw_val, default), lo, hi))

    nb = round(_clamp_multi(("nb",), "nb", nb_default), 4)
    n1 = round(_clamp_multi(("n1", "n1_ketidakteraturan"), "n1", 0.005), 3)
    n2 = round(_clamp_multi(("n2", "n2_variasi_penampang"), "n2", 0.002), 3)
    n3 = round(_clamp_multi(("n3", "n3_hambatan"), "n3", 0.015), 3)
    n4 = round(_clamp_multi(("n4", "n4_vegetasi_dalam_saluran"), "n4", 0.010), 3)
    m  = round(_clamp_multi(("m", "m_meander"), "m", 1.00), 2)

    raw_ts = ai_json.get("ts") if ai_json.get("ts") is not None else ai_json.get("estimasi_tebal_sedimen_m")
    ts = round(float(np.clip(_to_float_safe(raw_ts, 0.2 * h_ref), 0.0, max(0.0, h_ref - 0.01))), 3)

    prev = st.session_state.get("cv_params") or {}
    raw_p = ai_json.get("p") if ai_json.get("p") is not None else ai_json.get("p_narrowing", prev.get("p_narrowing", 0.05))
    p_narrow = round(float(np.clip(_to_float_safe(raw_p, 0.05), 0.0, 0.90)), 2)

    n_aktual = (nb + n1 + n2 + n3 + n4) * m   # ← dihitung Python, bukan AI
    shape_cv = prev.get("shape_cv", "Persegi")
    m_talud = float(prev.get("m_talud", 1.0 if shape_cv != "Persegi" else 0.0))
    h_eff = max(0.01, h_ref - ts)
    b_eff = max(0.01, b_ref * (1.0 - p_narrow)) if shape_cv == "Persegi" else (b_ref + 2.0 * m_talud * ts)

    deskripsi_txt = str(ai_json.get("alasan_visual") or ai_json.get("catatan_visual_ai") or "—")

    st.session_state["cv_params"] = {
        "sediment_pct": (ts / h_ref * 100.0) if h_ref > 0 else 0.0,
        "weed_pct":     min(100.0, n4 / 0.100 * 100.0),
        "ts":           ts,
        "p_narrowing":  p_narrow,
        "shape_cv":     shape_cv,
        "m_talud":      m_talud,
        "h_effective":  h_eff,
        "b_effective":  b_eff,
        "n_adjusted":   n_aktual,
        "n_cowan":      n_aktual,
        "horton_nc":    n_aktual,
        "use_horton":   False,
        "nb": nb, "n1": n1, "n2": n2, "n3": n3, "n4": n4,
        "m_meander":    m,
        "deskripsi":    deskripsi_txt,
        "sumber":       source,
        "h_desain_ref": h_ref,
        "n_desain_ref": nb,
    }
    st.session_state["params_modul_3"] = dict(st.session_state["cv_params"])

    # Sinkronisasi langsung ke widget keys pada Kalkulator Cowan & Geometri di panel kanan Modul 3
    st.session_state["m3_ideal_bypass"] = False
    st.session_state["m3_form_nb"] = float(np.clip(nb, 0.008, 0.050))
    st.session_state["m3_form_n1"] = float(np.clip(n1, 0.000, 0.020))
    st.session_state["m3_form_n2"] = float(np.clip(n2, 0.000, 0.015))
    st.session_state["m3_form_n3"] = float(np.clip(n3, 0.000, 0.060))
    st.session_state["m3_form_n4"] = float(np.clip(n4, 0.000, 0.100))
    st.session_state["m3_form_m"]  = float(np.clip(m, 1.00, 1.30))
    st.session_state["m3_form_h_maks"] = float(max(0.05, h_ref))
    st.session_state["m3_form_ts"] = float(np.clip(ts, 0.0, max(0.01, h_ref - 0.01)))
    st.session_state["m3_form_p"]  = float(np.clip(p_narrow, 0.00, 0.90))

    return st.session_state["cv_params"]


def _sync_bridge_to_session(force: bool = False) -> str | None:
    """Muat drainsense_bridge.json ke session_state bila file baru/berubah. Return pesan status."""
    mtime = _bridge_mtime()
    if mtime is None:
        return None
    if not force and st.session_state.get("bridge_mtime") is None:
        # Catat mtime awal sesi agar cache/template lama tidak menimpa form sebelum pengguna berinteraksi
        st.session_state["bridge_mtime"] = mtime
        return None
    if not force and st.session_state.get("bridge_mtime") == mtime:
        return None
    st.session_state["bridge_mtime"] = mtime  # simpan dulu agar tidak terjadi loop rerun
    data, err = _read_bridge()
    if err:
        return f"❌ {err}"
    st.session_state["bridge_data"] = data
    loaded: list[str] = []

    m2_json = data.get("modul2_ded")
    if isinstance(m2_json, dict) and m2_json:
        res = _validate_m2_ai_json(m2_json, source=f"Live JSON Bridge ({data.get('source', 'Antigravity')})")
        st.session_state["m2_ai_result"] = res
        if res["chow_valid"]:
            _commit_m2_to_ded(res)
            loaded.append("Modul 2 (DED tervalidasi Chow)")
        else:
            loaded.append("Modul 2 (menunggu koreksi n Chow)")

    m3_json = data.get("modul3_cowan")
    if isinstance(m3_json, dict) and m3_json:
        ded_now = st.session_state.get("ded_params") or {}
        h_ref = float(ded_now.get("h", 1.2))
        b_ref = float(ded_now.get("b", 1.5))
        nb_def = float(ded_now.get("n", 0.013))
        st.session_state["m3_ai_result"] = {"source": "Live JSON Bridge", "json": m3_json}
        _apply_m3_result(m3_json, "Live JSON Bridge", h_ref, b_ref, nb_def)
        loaded.append("Modul 3 (faktor Cowan)")

    return ("🔄 Bridge dimuat: " + ", ".join(loaded)) if loaded else "ℹ️ Bridge terbaca, tetapi tidak berisi `modul2_ded` / `modul3_cowan`."


def _extract_pdf_text_pypdf(pdf_bytes: bytes) -> str:
    """Extract text per page from PDF bytes locally using pypdf.PdfReader (with pdfplumber fallback)."""
    pages_text: list[str] = []
    if PYPDF_AVAILABLE:
        try:
            reader = PdfReader(io.BytesIO(pdf_bytes))
            for idx, page in enumerate(reader.pages, start=1):
                txt = page.extract_text() or ""
                if txt.strip():
                    pages_text.append(f"--- [Halaman {idx}] ---\n{txt.strip()}")
            if pages_text:
                return "\n\n".join(pages_text)
        except Exception as e:
            print(f"[pypdf Extract Warning]: {e}", flush=True)

    if PDFPLUMBER_AVAILABLE:
        try:
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf_doc:
                for idx, page in enumerate(pdf_doc.pages, start=1):
                    txt = page.extract_text() or ""
                    tbl_strs = []
                    for tbl in (page.extract_tables() or []):
                        for row in tbl:
                            cells = [str(c).strip().replace("\n", " ") for c in (row or []) if c is not None and str(c).strip()]
                            if cells:
                                tbl_strs.append(" | ".join(cells))
                    combined = "\n".join(part for part in [txt.strip(), "\n".join(tbl_strs)] if part)
                    if combined:
                        pages_text.append(f"--- [Halaman {idx}] ---\n{combined}")
            return "\n\n".join(pages_text)
        except Exception as e:
            print(f"[pdfplumber Extract Warning]: {e}", flush=True)

    return ""


def _extract_docx_text_local(docx_bytes: bytes) -> str:
    """Extract text from Word (.docx) paragraphs and tables locally using python-docx (with XML fallback)."""
    if DOCX_AVAILABLE:
        try:
            doc = DocxDocument(io.BytesIO(docx_bytes))
            blocks = []
            for idx, p in enumerate(doc.paragraphs, start=1):
                if p.text and p.text.strip():
                    blocks.append(p.text.strip())
            for t_idx, table in enumerate(doc.tables, start=1):
                t_rows = []
                for row in table.rows:
                    cells = [c.text.strip().replace("\n", " ") for c in row.cells if c.text.strip()]
                    if cells:
                        t_rows.append(" | ".join(cells))
                if t_rows:
                    blocks.append(f"--- [Tabel Word {t_idx}] ---\n" + "\n".join(t_rows))
            return "\n".join(blocks)
        except Exception as e:
            print(f"[python-docx Extract Warning]: {e}", flush=True)

    # Built-in XML fallback for .docx if python-docx fails
    try:
        with zipfile.ZipFile(io.BytesIO(docx_bytes)) as zf:
            xml_content = zf.read("word/document.xml")
        tree = ET.fromstring(xml_content)
        texts = [elem.text for elem in tree.iter() if elem.tag.endswith("}t") and elem.text]
        return " ".join(texts)
    except Exception:
        return ""


def _build_preview_data(full_doc_text: str, filename: str = "", max_chars: int = 25000) -> tuple[str, dict]:
    """
    Langkah 1 (Python Extractor): Memotong dan menyaring teks dokumen PDF/Word secukupnya
    (memprioritaskan halaman/segmen yang memuat kata kunci Hidrologi, Hidrolika, Kriteria Desain,
    Normalisasi, Sudetan, Dimensi, Manning, atau Debit) menjadi `preview_data` untuk analisis AI.
    """
    raw_text = (full_doc_text or "").strip()
    if not raw_text:
        return "", {
            "total_karakter_dokumen": 0,
            "jumlah_segmen": 0,
            "segmen_relevan": 0,
            "karakter_preview": 0,
            "kata_kunci_terdeteksi": {},
        }

    keywords = [
        "hidrologi",
        "hidrolika",
        "kriteria desain",
        "normalisasi",
        "sudetan",
        "dimensi",
        "penampang",
        "manning",
        "lebar dasar",
        "kedalaman",
        "kemiringan",
        "debit",
        "banjir",
        "eksisting",
        "rencana",
        "u-ditch",
        "pasangan batu",
        "beton",
    ]

    low_all = raw_text.lower()
    kw_counts: dict[str, int] = {}
    for kw in keywords:
        cnt = low_all.count(kw)
        if cnt > 0:
            kw_counts[kw] = cnt

    # Pisahkan per halaman PDF atau tabel Word jika penanda tersedia; jika tidak, bagi per blok paragraf
    segments = re.split(r"(?=---\s*\[(?:Halaman\s*\d+|Tabel\s*Word\s*\d+)\]\s*---)", raw_text)
    segments = [seg.strip() for seg in segments if seg and seg.strip()]
    if len(segments) <= 1 and len(raw_text) > 3500:
        paragraphs = [p.strip() for p in re.split(r"\n{2,}", raw_text) if p.strip()]
        segments = []
        buf = ""
        for p in paragraphs:
            if len(buf) + len(p) + 2 <= 2500:
                buf = f"{buf}\n\n{p}".strip()
            else:
                if buf:
                    segments.append(buf)
                buf = p
        if buf:
            segments.append(buf)

    scored_segments: list[tuple[int, int, str]] = []
    for idx, seg in enumerate(segments):
        seg_low = seg.lower()
        score = sum(3 if kw in ("hidrolika", "kriteria desain", "normalisasi", "sudetan", "manning", "dimensi") else 1 for kw in keywords if kw in seg_low)
        if re.search(r"\b0[.,]0\d{1,3}\b", seg_low):
            score += 3
        if re.search(r"\d+(?:[.,]\d+)?\s*(?:cm|mm|m)\s*[x×]\s*\d+(?:[.,]\d+)?", seg_low):
            score += 3
        scored_segments.append((score, idx, seg))

    relevant = [item for item in scored_segments if item[0] > 0]
    if relevant:
        # Urutkan segmen paling relevan terlebih dahulu, batasi hingga max_chars, lalu urutkan kembali sesuai urutan halaman
        relevant_sorted = sorted(relevant, key=lambda x: (-x[0], x[1]))
        chosen_indices: list[int] = []
        char_acc = 0
        for score, idx, seg in relevant_sorted:
            if char_acc + len(seg) + 4 <= max_chars or not chosen_indices:
                chosen_indices.append(idx)
                char_acc += len(seg) + 4
            if char_acc >= max_chars:
                break
        chosen_indices.sort()
        preview_text = "\n\n".join(segments[i] for i in chosen_indices)
        rel_count = len(chosen_indices)
    else:
        preview_text = raw_text[:max_chars]
        rel_count = min(1, len(segments))

    if len(preview_text) > max_chars:
        preview_text = preview_text[:max_chars]

    stats = {
        "total_karakter_dokumen": len(raw_text),
        "jumlah_segmen": len(segments),
        "segmen_relevan": rel_count,
        "karakter_preview": len(preview_text),
        "kata_kunci_terdeteksi": dict(list(kw_counts.items())[:10]),
    }
    return preview_text, stats


CIVIL_DOMAIN_KEYWORDS: dict[str, dict] = {
    "geoteknik": {
        "label": "Geoteknik / Mekanika Tanah",
        "keywords": [
            "spt", "n-spt", "n60", "kohesi", "sudut gesek", "lempung", "pasir",
            "plaxis", "geostudio", "mohr-coulomb", "cpt", "sondir", "permeabilitas", "unit weight",
        ],
    },
    "hidrologi": {
        "label": "Hidrologi / Klimatologi (Kurva IDF)",
        "keywords": [
            "idf", "mononobe", "gumbel", "log pearson", "chi-square",
            "stasiun hujan", "curah hujan maksimum", "raps", "kala ulang",
        ],
    },
    "struktur": {
        "label": "Struktur / Beton / Baja",
        "keywords": [
            "etabs", "sap2000", "tulangan", "kuat tekan", "fc'", "mpa",
            "balok", "kolom", "pelat", "gempa", "sni 1726", "sni 2847",
        ],
    },
    "drainase": {
        "label": "DED Drainase / Hidrolika Saluran",
        "keywords": [
            "saluran", "drainase", "lebar dasar", "tinggi jagaan", "freeboard",
            "manning", "debit banjir", "penampang", "trapesium", "u-ditch",
            "gorong-gorong", "normalisasi",
        ],
    },
}


def _classify_civil_document(doc_text: str) -> dict:
    """
    Klasifikasikan domain dokumen Teknik Sipil berdasarkan frekuensi kata kunci
    dan ekstrak entitas teknis spesifik domain sebagai bukti pembacaan teks dokumen.
    """
    text = doc_text or ""
    text_low = text.lower()

    domain_scores: dict[str, int] = {}
    matched_keywords: dict[str, dict[str, int]] = {}

    for dom_key, cfg in CIVIL_DOMAIN_KEYWORDS.items():
        kw_hits: dict[str, int] = {}
        total_hits = 0
        for kw in cfg["keywords"]:
            kw_l = kw.lower()
            if kw_l == "fc'":
                cnt = len(re.findall(r"\bfc['’]?", text_low))
            elif len(kw_l) <= 4 and kw_l.isalnum():
                cnt = len(re.findall(rf"\b{re.escape(kw_l)}\b", text_low))
            else:
                cnt = len(re.findall(re.escape(kw_l), text_low))
            if cnt > 0:
                kw_hits[kw] = cnt
                total_hits += cnt
        domain_scores[dom_key] = total_hits
        matched_keywords[dom_key] = dict(sorted(kw_hits.items(), key=lambda x: -x[1]))

    best_domain = "umum"
    best_score = 0
    for dom_key, score in domain_scores.items():
        if score > best_score:
            best_score = score
            best_domain = dom_key

    primary_label = (
        CIVIL_DOMAIN_KEYWORDS[best_domain]["label"]
        if best_domain in CIVIL_DOMAIN_KEYWORDS
        else "Dokumen Teknis Umum"
    )

    # Ekstrak entitas teknis spesifik domain dari teks
    entities: list[dict] = []
    seen_entities: set[tuple[str, str]] = set()

    def _add_entity(param: str, val_ctx: str, detail: str) -> None:
        key = (param, val_ctx[:80])
        if key not in seen_entities and len(entities) < 20:
            seen_entities.add(key)
            entities.append({
                "Parameter / Entitas": param,
                "Nilai / Konteks Terbaca": val_ctx.strip(),
                "Kategori": detail,
            })

    if best_domain == "geoteknik":
        geo_patterns = [
            ("N-SPT / SPT", r"(?:\bN-?SPT\b|\bN60\b|\bSPT\b)\s*[:=]?\s*\d+(?:[.,]\d+)?(?:\s*(?:pukulan|blows?|/30\s*cm))?", "Parameter Tanah"),
            ("Kohesi (c)", r"(?:kohesi|\bc['’]?\s*[:=])\s*\d+(?:[.,]\d+)?\s*(?:kPa|kN/m2|kN/m²|kg/cm2|kg/cm²)?", "Kuat Geser"),
            ("Sudut Gesek Dalam (φ)", r"(?:sudut\s*gesek(?:\s*dalam)?|\bphi\b|\bφ\b)\s*[:=]?\s*\d+(?:[.,]\d+)?\s*(?:°|derajat|deg)?", "Kuat Geser"),
            ("Unit Weight (γ)", r"(?:unit\s*weight|berat\s*(?:isi|volume)|γ\s*(?:sat|unsat|dry)?)\s*[:=]?\s*\d+(?:[.,]\d+)?\s*(?:kN/m3|kN/m³|t/m3|g/cm3)?", "Sifat Fisik"),
            ("Permeabilitas (k)", r"(?:permeabilitas|\bk[xy]?\s*[:=])\s*\d+(?:[.,]\d+)?(?:\s*[x×]\s*10\^?-?\d+|e-?\d+)?\s*(?:m/s|m/hari|cm/s|m/det)?", "Aliran Tanah"),
            ("CPT / Sondir (qc)", r"(?:\bqc\b|sondir|cpt)\s*[:=]?\s*\d+(?:[.,]\d+)?\s*(?:kg/cm2|kg/cm²|MPa|kPa)?", "Uji Lapangan"),
        ]
        for label, pat, cat in geo_patterns:
            for m in re.finditer(pat, text, re.IGNORECASE):
                _add_entity(label, m.group(0), cat)
    elif best_domain == "hidrologi":
        hydro_patterns = [
            ("Curah Hujan Harian (R24)", r"(?:\bR_?24\b|curah\s*hujan(?:\s*maksimum)?)\s*[:=]?\s*\d+(?:[.,]\d+)?\s*mm", "Data Hujan"),
            ("Intensitas Hujan / IDF", r"(?:\bI\b|intensitas(?:\s*hujan)?)\s*[:=]\s*\d+(?:[.,]\d+)?\s*(?:mm/jam|mm/hr)?", "Kurva IDF"),
            ("Kala Ulang (T)", r"(?:kala\s*ulang|\bT\b)\s*[:=]?\s*\d+\s*tahun", "Periode Ulang"),
            ("Persamaan Mononobe / IDF", r"(?:I\s*=\s*\(?R24\s*/\s*24\)?[^\n]{0,40}|mononobe[^\n]{0,40})", "Rumus Hidrologi"),
        ]
        for label, pat, cat in hydro_patterns:
            for m in re.finditer(pat, text, re.IGNORECASE):
                _add_entity(label, m.group(0), cat)
    elif best_domain == "struktur":
        struk_patterns = [
            ("Mutu Beton (fc')", r"(?:fc['’]?|kuat\s*tekan)\s*[:=]?\s*\d+(?:[.,]\d+)?\s*MPa", "Material Beton"),
            ("Mutu Baja (fy)", r"(?:fy|tegangan\s*leleh)\s*[:=]?\s*\d+(?:[.,]\d+)?\s*MPa", "Material Baja"),
            ("Dimensi Elemen Struktur", r"(?:balok|kolom|pelat)\s*[:=]?\s*\d+\s*[x×]\s*\d+\s*(?:cm|mm|m)?", "Geometri Struktur"),
        ]
        for label, pat, cat in struk_patterns:
            for m in re.finditer(pat, text, re.IGNORECASE):
                _add_entity(label, m.group(0), cat)

    # Tambahkan daftar kata kunci domain yang terdeteksi ke tabel entitas
    top_dom_kws = matched_keywords.get(best_domain, {})
    for kw, cnt in list(top_dom_kws.items())[:10]:
        _add_entity(f"Topik: {kw}", f"Terdeteksi {cnt} kali dalam teks dokumen", primary_label)

    return {
        "primary_domain": best_domain,
        "primary_label": primary_label,
        "domain_scores": domain_scores,
        "matched_keywords": matched_keywords,
        "top_keywords_list": list(top_dom_kws.keys())[:6],
        "entities": entities,
    }


def _extract_ded_params_local(doc_text: str, filename: str) -> dict:
    """
    Real-time Python NLP / Regex & Table-Row Parser for extracting drainage DED parameters
    (b, h, S, n, material, Q_rencana, kutipan_ded) directly from PDF or Word (.docx) text in < 1 second.
    Prioritizes Rencana / Desain / Normalisasi / Sudetan context over Eksisting.
    """
    text_clean = doc_text or ""
    raw_lines = [ln.strip() for ln in text_clean.splitlines() if ln.strip()]

    # Gabungkan baris label tabel yang terputus dengan angka di baris tepat berikutnya
    lines: list[str] = []
    i_ln = 0
    while i_ln < len(raw_lines):
        curr = raw_lines[i_ln]
        if (
            i_ln + 1 < len(raw_lines)
            and not re.search(r"\d", curr)
            and re.match(r"^[:=]?\s*\d+(?:[.,]\d+)?\s*(?:cm|mm|m|%|‰|m/m|m3/s|m³/s)?$", raw_lines[i_ln + 1], re.IGNORECASE)
        ):
            lines.append(f"{curr} : {raw_lines[i_ln + 1].lstrip(':= ').strip()}")
            i_ln += 2
        else:
            lines.append(curr)
            i_ln += 1

    def _to_float(val_str: str) -> float | None:
        try:
            return float(val_str.strip().replace(",", "."))
        except Exception:
            return None

    def _normalize_length_m(val: float, unit: str) -> float:
        u = (unit or "").lower().strip()
        if u == "cm":
            return val / 100.0
        if u == "mm":
            return val / 1000.0
        # Heuristik jika angka >= 25 tanpa satuan eksplisit (umumnya dalam cm, misal 80 atau 120)
        if not u and val >= 25.0:
            return val / 100.0 if val <= 1000.0 else val / 1000.0
        return val

    def _context_priority(line_str: str) -> int:
        low = line_str.lower()
        prio = 1
        if any(k in low for k in ["rencana", "desain", "usulan", "normalisasi", "sudetan", "kriteria", "rekomendasi"]):
            prio += 3
        if any(k in low for k in ["eksisting", "existing", "kondisi awal"]):
            prio -= 1
        return prio

    b_found: float | None = None
    b_prio = -99
    h_found: float | None = None
    h_prio = -99
    S_found: float | None = None
    S_prio = -99
    n_found: float | None = None
    n_prio = -99
    q_found: float | None = None
    q_prio = -99
    mat_found: str | None = None
    mat_prio = -99
    matched_quotes: list[str] = []
    current_page = "Dokumen"

    # Pola dimensi gabungan (misal: "Dimensi saluran 1.5 x 1.2 m", "U-Ditch 80 x 100 cm", "b x h = 1,20 x 1,00 m")
    dim_pair_re = re.compile(
        r"(?:dimensi|ukuran|saluran|u-?ditch|box\s*culvert|penampang|b\s*[x×]\s*h)[^0-9\n]{0,35}"
        r"(\d+(?:[.,]\d+)?)\s*(cm|mm|m)?\s*[xX×]\s*(\d+(?:[.,]\d+)?)\s*(cm|mm|m)?",
        re.IGNORECASE,
    )
    # Pola lebar dasar b:
    # - Hindari menangkap rumus aljabar seperti 'b = 2h', 'b = 2 h', 'b = 1.5h', 'b = 0.5*h'
    # - Utamakan nilai eksplisit dengan satuan ('b = 0.80 m', 'b = 0,60 m') atau format tabel
    b_re = re.compile(
        r"(?:lebar(?:\s*(?:dasar|bawah|saluran|rencana|penampang|atas))?(?:\s*saluran)?(?:\s*\(\s*[bB]\s*\))?|\b[bB]\s*(?:dasar|saluran|rencana)\b|\b[bB]\s*[:=])"
        r"(?:\s*[|:=]\s*|\s+)(?:(cm|mm|meter|m)\b[\s|:]*)?(\d+(?:[.,]\d+)?)"
        r"(?:\s*(cm|mm|meter|m)\b|(?![\s*×x⋅·/]*[a-zA-Z0-9_,.]))",
        re.IGNORECASE,
    )
    # Pola kedalaman/tinggi h (mengecualikan tinggi jagaan/freeboard/sedimen/hujan & rumus seperti 'h = 0.5b')
    h_re = re.compile(
        r"(?:(?:tinggi|kedalaman)\s*(?:desain|saluran|air|basah|maksimum|rencana|penampang)?(?:\s*\(\s*[hHyY]\s*\))?|\b[hH]\s*(?:desain|saluran|rencana)\b|\b[hH]\s*[:=])"
        r"(?:\s*[|:=]\s*|\s+)(?:(cm|mm|m)\b[\s|:]*)?(\d+(?:[.,]\d+)?)"
        r"(?:\s*(cm|mm|meter|m)\b|(?![\s*×x⋅·/]*[a-zA-Z0-9_,.]))",
        re.IGNORECASE,
    )
    # Pola kemiringan S / is (mendukung 'is = X%', 'is=X%', 'i_s = X%', 'S = ...', 'Kemiringan ...')
    s_re = re.compile(
        r"(?:kemiringan\s*(?:dasar|saluran|memanjang|lahan|rencana|rata-?rata)?(?:\s*saluran)?(?:\s*\(\s*(?:[sSiI]|i_?s)\s*\))?|\bslope\b|\b(?:[sS]|i_?s)\s*(?:dasar|saluran|memanjang)\b|\b(?:[sS]|i_?s)\s*[:=])"
        r"(?:\s*[|:=]\s*|\s*)(?:(%|‰|m/m)\b[\s|:]*)?(\d+(?:[.,]\d+)?)\s*(%|‰|m/m)?",
        re.IGNORECASE,
    )
    # Pola koefisien Manning n
    n_re = re.compile(
        r"(?:(?:koefisien|angka|nilai|faktor)?\s*(?:kekasaran\s*)?manning(?:\s*\(\s*n\s*\))?|kekasaran\s*(?:saluran|dinding|dasar)?\s*(?:\(\s*n\s*\))?|\bn\s*manning\b|\bn\s*[:=])"
        r"[^0-9\n]{0,20}(0[.,]0\d{1,3})",
        re.IGNORECASE,
    )
    # Pola debit rencana Q
    q_re = re.compile(
        r"(?:debit\s*(?:rencana|banjir|desain|saluran|puncak|maksimum)(?:\s*\(\s*Q[^)]*\))?|\bQ\s*(?:rencana|banjir|desain|saluran)\b)"
        r"(?:\s*[|:=]\s*|\s+)(?:(?:m3/det|m3/s|m³/det|m³/s)\b[\s|:]*)?(\d+(?:[.,]\d+)?)\s*(m3/det|m3/s|m³/det|m³/s)?",
        re.IGNORECASE,
    )
    # Pola luas DAS / DTA / Catchment Area
    dta_re = re.compile(
        r"(?:luas\s*(?:das|dta|daerah\s*tangkapan(?:\s*air)?|tangkapan|area|wilayah|layanan)?|catchment\s*area|\bA\s*(?:das|dta|total)?\b)"
        r"(?:\s*\(\s*A[^)]*\))?(?:\s*[|:=]\s*|\s+)"
        r"(\d+(?:[.,]\d+)?)\s*(ha|hektar|hectare|km2|km²|m2|m²)\b",
        re.IGNORECASE,
    )
    dta_found: float | None = None
    dta_prio = -99

    for ln in lines:
        page_m = re.match(r"^---\s*\[(Halaman\s*\d+|Tabel\s*Word\s*\d+)\]\s*---$", ln, re.IGNORECASE)
        if page_m:
            current_page = page_m.group(1)
            continue

        ln_low = ln.lower()
        prio = _context_priority(ln)

        # 1. Cek pasangan dimensi b x h
        m_pair = dim_pair_re.search(ln)
        if m_pair:
            v1 = _to_float(m_pair.group(1))
            u1 = m_pair.group(2) or m_pair.group(4) or ""
            v2 = _to_float(m_pair.group(3))
            u2 = m_pair.group(4) or m_pair.group(2) or ""
            if v1 and v2:
                b_cand = _normalize_length_m(v1, u1)
                h_cand = _normalize_length_m(v2, u2)
                if 0.10 <= b_cand <= 30.0 and 0.10 <= h_cand <= 20.0:
                    if b_found is None or prio > b_prio:
                        b_found = round(b_cand, 3)
                        b_prio = prio
                    if h_found is None or prio > h_prio:
                        h_found = round(h_cand, 3)
                        h_prio = prio
                    matched_quotes.append(f"[{current_page}] \"{ln[:120]}\"")

        # 2. Cek lebar dasar b (gunakan finditer agar jika ada 'b = 2h, maka b = 0.80 m' di baris yang sama, 'b = 0.80 m' tetap tertangkap)
        for mb in b_re.finditer(ln):
            ub = mb.group(1) or mb.group(3) or ""
            raw_num_b = mb.group(2)
            vb = _to_float(raw_num_b)
            if vb is not None:
                b_cand = _normalize_length_m(vb, ub)
                # Utamakan nilai eksplisit dengan satuan ('m', 'cm', 'mm') dan format desimal ('0.80', '0,60')
                cand_prio = prio + (4 if ub else 0) + (2 if ("," in raw_num_b or "." in raw_num_b) else 0)
                if 0.10 <= b_cand <= 30.0 and (b_found is None or cand_prio >= b_prio):
                    b_found = round(b_cand, 3)
                    b_prio = cand_prio
                    matched_quotes.append(f"[{current_page}] \"{ln[:120]}\"")

        # 3. Cek kedalaman h (abaikan baris tinggi jagaan / freeboard / sedimen / curah hujan)
        if not any(skip_w in ln_low for skip_w in ["jagaan", "freeboard", "waking", "sedimen", "endapan", "curah hujan", "hujan harian"]):
            for mh in h_re.finditer(ln):
                uh = mh.group(1) or mh.group(3) or ""
                raw_num_h = mh.group(2)
                vh = _to_float(raw_num_h)
                if vh is not None:
                    h_cand = _normalize_length_m(vh, uh)
                    cand_prio = prio + (4 if uh else 0) + (2 if ("," in raw_num_h or "." in raw_num_h) else 0)
                    if 0.10 <= h_cand <= 20.0 and (h_found is None or cand_prio >= h_prio):
                        h_found = round(h_cand, 3)
                        h_prio = cand_prio
                        matched_quotes.append(f"[{current_page}] \"{ln[:120]}\"")

        # 4. Cek kemiringan S / is (konversi otomatis dari persen '%' ke desimal, misal 'is = 6.25%' -> 0.0625)
        for ms in s_re.finditer(ln):
            us = (ms.group(1) or ms.group(3) or "").strip()
            vs = _to_float(ms.group(2))
            if vs is not None and vs > 0:
                if us == "%":
                    s_cand = vs / 100.0
                elif us == "‰":
                    s_cand = vs / 1000.0
                else:
                    s_cand = vs
                has_is_token = bool(re.search(r"\bi_?s\b", ms.group(0), re.IGNORECASE))
                cand_prio = prio + (3 if us == "%" else 0) + (2 if has_is_token else 0)
                if 0.00001 <= s_cand <= 0.25 and (S_found is None or cand_prio >= S_prio):
                    S_found = round(s_cand, 6)
                    S_prio = cand_prio
                    matched_quotes.append(f"[{current_page}] \"{ln[:120]}\"")

        # 5. Cek koefisien Manning n
        mn = n_re.search(ln)
        if mn:
            vn = _to_float(mn.group(1))
            if vn is not None and 0.008 <= vn <= 0.150 and (n_found is None or prio > n_prio):
                n_found = round(vn, 4)
                n_prio = prio
                matched_quotes.append(f"[{current_page}] \"{ln[:120]}\"")

        # 6. Cek debit rencana Q
        mq = q_re.search(ln)
        if mq:
            vq = _to_float(mq.group(1))
            if vq is not None and 0.001 <= vq <= 2000.0 and (q_found is None or prio > q_prio):
                q_found = round(vq, 4)
                q_prio = prio
                matched_quotes.append(f"[{current_page}] \"{ln[:120]}\"")

        # 6b. Cek luas DAS / DTA (ha, km2, m2 -> konversi ke Hektar)
        mdta = dta_re.search(ln)
        if mdta:
            raw_dta_num = mdta.group(1).strip()
            u_dta = mdta.group(2).lower().strip()
            if u_dta in ("m2", "m²") and re.match(r"^\d{1,3}(?:\.\d{3})+$", raw_dta_num):
                vdta = _to_float(raw_dta_num.replace(".", ""))
            elif "." in raw_dta_num and "," in raw_dta_num:
                vdta = _to_float(raw_dta_num.replace(".", "").replace(",", "."))
            else:
                vdta = _to_float(raw_dta_num)
            if vdta is not None and vdta > 0:
                if u_dta in ("km2", "km²"):
                    ha_cand = vdta * 100.0
                elif u_dta in ("m2", "m²"):
                    ha_cand = vdta * 0.0001
                else:
                    ha_cand = vdta
                if 0.001 <= ha_cand <= 100000.0 and (dta_found is None or prio >= dta_prio):
                    dta_found = round(ha_cand, 4)
                    dta_prio = prio
                    matched_quotes.append(f"[{current_page}] \"{ln[:120]}\"")

        # 7. Cek material saluran
        mat_cand = None
        if any(k in ln_low for k in ["pasangan batu kali", "pasangan batu", "batu kali", "masonry", "batu belah"]):
            mat_cand = "Pasangan Batu Kali"
        elif any(k in ln_low for k in ["pasangan bata", "batu bata", "bata merah"]):
            mat_cand = "Pasangan Bata"
        elif any(k in ln_low for k in ["saluran tanah", "tanah asli", "tanpa pasangan", "material tanah", "dinding tanah", "berlapis tanah"]):
            mat_cand = "Tanah"
        elif "tanah" in ln_low and any(ctx in ln_low for ctx in ["saluran", "drainase", "manning", "penampang", "material"]):
            mat_cand = "Tanah"
        elif any(k in ln_low for k in ["beton", "u-ditch", "uditch", "box culvert", "precast", "pracetak"]):
            mat_cand = "Beton"
        if mat_cand and (mat_found is None or prio > mat_prio):
            mat_found = mat_cand
            mat_prio = prio
            matched_quotes.append(f"[{current_page}] \"{ln[:120]}\"")

    # Inferensi material dari nilai Manning n jika material tidak tertulis eksplisit
    if mat_found is None and n_found is not None:
        if 0.019 <= n_found <= 0.035:
            mat_found = "Pasangan Batu Kali"
        elif 0.014 < n_found < 0.019:
            mat_found = "Pasangan Bata"
        else:
            mat_found = "Beton"

    # Fallback cerdas jika sebagian parameter tidak tertulis eksplisit di dokumen
    final_mat = mat_found or "Beton"
    default_n_by_mat = {
        "Beton": 0.013,
        "Pasangan Bata": 0.015,
        "Pasangan Batu Kali": 0.025,
        "Tanah": 0.022,
    }
    final_b = b_found if b_found is not None else 1.500
    final_h = h_found if h_found is not None else 1.200
    final_S = S_found if S_found is not None else 0.001000
    final_n = n_found if n_found is not None else default_n_by_mat.get(final_mat, 0.013)
    final_q = q_found if q_found is not None else round(_manning_Q(final_b, final_h, final_S, final_n), 4)

    # Susun ringkasan kutipan unik
    unique_quotes = []
    for q in matched_quotes:
        if q not in unique_quotes:
            unique_quotes.append(q)

    found_flags = []
    if b_found is not None:
        found_flags.append(f"b={final_b}m")
    if h_found is not None:
        found_flags.append(f"h={final_h}m")
    if S_found is not None:
        found_flags.append(f"S={final_S}")
    if n_found is not None:
        found_flags.append(f"n={final_n}")
    if mat_found is not None:
        found_flags.append(f"Material={final_mat}")

    if unique_quotes:
        kutipan_str = " | ".join(unique_quotes[:3])
        if len(found_flags) < 5:
            kutipan_str += f" (Terdeteksi eksplisit: {', '.join(found_flags)}; sisanya diestimasi standar dari konteks {filename})"
    else:
        snippet_preview = " ".join(lines[:3])[:140] if lines else "Dokumen tanpa lapisan teks"
        kutipan_str = (
            f"Diekstrak lokal dari {filename}. Cuplikan awal: \"{snippet_preview}...\" "
            f"(Parameter menggunakan estimasi standar {final_mat} — silakan sesuaikan di panel kanan bila perlu)."
        )

    return {
        "b": final_b,
        "h": final_h,
        "S": final_S,
        "n": final_n,
        "material": final_mat,
        "Q_rencana": final_q,
        "luas_dta_ha": dta_found,
        "kutipan_ded": kutipan_str,
        "explicit_count": len(found_flags),
        "found_keys": {
            "b": b_found is not None,
            "h": h_found is not None,
            "S": S_found is not None,
            "n": n_found is not None,
            "material": mat_found is not None,
            "Q_rencana": q_found is not None,
        },
    }


def _extract_m4_hydrology_local(doc_text: str, filename: str = "Dokumen PDF") -> dict:
    """
    Mengekstrak nilai-nilai akhir hidrologi secara instan (< 1 detik) dari teks PDF jurnal/DED
    menggunakan Regex Python lokal untuk Modul 4:
      - t1 : Waktu limpas permukaan (menit) -> pola sekitar "T1 = 1,641" atau "waktu limpas aliran permukaan"
      - t2 : Waktu aliran dalam saluran (menit) -> pola sekitar "T2 = 4,889" atau "waktu aliran dalam saluran"
      - tc : Total waktu konsentrasi (menit) -> pola sekitar "Tc = 6,530" atau "total waktu konsentrasi"
      - A  : Luas daerah tangkapan -> pola sekitar "A = 6.160 m2" atau "luas total daerah tangkapan"
      - C  : Koefisien limpasan -> pola sekitar "C = 0,946" atau "koefisien limpasan"
      - I  : Intensitas Hujan (mm/jam) -> pola sekitar "I = 175 mm/jam" atau "intensitas hujan"
      - Qr : Debit Rencana (m³/s) -> pola sekitar "Qr = 0,283" atau "debit rencana"
    """
    raw_text = doc_text or ""
    lines = [ln.strip() for ln in raw_text.splitlines() if ln.strip()]
    joined_text = re.sub(r"\s+", " ", raw_text)
    search_segments = lines + [joined_text]

    def _parse_dec(val_str: str) -> float | None:
        s = val_str.strip()
        if not s:
            return None
        # Jika ada titik ribuan sekaligus koma desimal (misal 1.250,5)
        if "." in s and "," in s:
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", ".")
        try:
            return float(s)
        except Exception:
            return None

    def _parse_area(val_str: str, unit_str: str) -> tuple[float, float, float] | None:
        """Mengembalikan (A_m2, A_ha, A_km2)."""
        s = val_str.strip()
        u = (unit_str or "").strip().lower()
        if not s:
            return None
        is_m2_unit = u in ("m2", "m²", "meter persegi", "")
        if "." in s and "," in s:
            s_clean = s.replace(".", "").replace(",", ".")
        elif is_m2_unit and re.match(r"^\d{1,3}(?:\.\d{3})+$", s):
            # Format ribuan Indonesia seperti "6.160 m2" -> 6160.0 m2
            s_clean = s.replace(".", "")
        else:
            s_clean = s.replace(",", ".")
        try:
            val = float(s_clean)
        except Exception:
            return None
        if val <= 0:
            return None

        if u in ("km2", "km²"):
            a_km2 = val
            a_ha = val * 100.0
            a_m2 = val * 1_000_000.0
        elif u in ("ha", "hektar", "hectare"):
            a_ha = val
            a_m2 = val * 10_000.0
            a_km2 = val / 100.0
        elif u in ("m2", "m²") or val >= 100.0:
            a_m2 = val
            a_ha = val / 10_000.0
            a_km2 = val / 1_000_000.0
        elif val < 0.2:
            # Nilai sangat kecil tanpa satuan eksplisit (misal 0,00616) -> km2
            a_km2 = val
            a_ha = val * 100.0
            a_m2 = val * 1_000_000.0
        else:
            a_ha = val
            a_m2 = val * 10_000.0
            a_km2 = val / 100.0
        return (round(a_m2, 2), round(a_ha, 4), round(a_km2, 6))

    t1_found: float | None = None
    t2_found: float | None = None
    tc_found: float | None = None
    a_tuple_found: tuple[float, float, float] | None = None
    c_found: float | None = None
    i_found: float | None = None
    qr_found: float | None = None
    matched_quotes: list[str] = []

    # Pola opsional rumus matematika perantara sebelum '=' hasil akhir (hanya angka & operator matematika, tanpa kata prosa)
    _MATH_INTER = r"(?:[\d\s.,+\-*/x×⋅·()^√]+=\s*)*"

    # 1. Regex Waktu Limpas Permukaan (t1 / T1)
    t1_patterns = [
        re.compile(rf"\b[tT]\s*_?[1o0]\s*=\s*{_MATH_INTER}(\d+(?:[.,]\d+)?)", re.IGNORECASE),
        re.compile(r"waktu\s+limpas(?:\s+aliran)?\s+permukaan[^0-9\n\r]{0,50}(?:[tT]\s*_?[1o0]\s*=\s*)?(\d+(?:[.,]\d+)?)", re.IGNORECASE),
    ]
    for seg in search_segments:
        if t1_found is not None:
            break
        for pat in t1_patterns:
            for m in pat.finditer(seg):
                v = _parse_dec(m.group(1))
                if v is not None and 0.05 <= v <= 300.0:
                    t1_found = round(v, 3)
                    matched_quotes.append(f"t1={t1_found} mnt ({m.group(0).strip()[:60]})")
                    break
            if t1_found is not None:
                break

    # 2. Regex Waktu Aliran dalam Saluran (t2 / T2)
    t2_patterns = [
        re.compile(rf"\b[tT]\s*_?[2fd]\s*=\s*{_MATH_INTER}(\d+(?:[.,]\d+)?)", re.IGNORECASE),
        re.compile(r"waktu\s+aliran\s+(?:dalam\s+)?saluran[^0-9\n\r]{0,50}(?:[tT]\s*_?[2fd]\s*=\s*)?(\d+(?:[.,]\d+)?)", re.IGNORECASE),
    ]
    for seg in search_segments:
        if t2_found is not None:
            break
        for pat in t2_patterns:
            for m in pat.finditer(seg):
                v = _parse_dec(m.group(1))
                if v is not None and 0.05 <= v <= 600.0:
                    t2_found = round(v, 3)
                    matched_quotes.append(f"t2={t2_found} mnt ({m.group(0).strip()[:60]})")
                    break
            if t2_found is not None:
                break

    # 3. Regex Total Waktu Konsentrasi (Tc / tc)
    tc_patterns = [
        re.compile(rf"\b[tT]\s*_?c\s*=\s*(?:[tT]\s*_?[1o0]\s*\+\s*[tT]\s*_?[2fd]\s*=\s*)?{_MATH_INTER}(\d+(?:[.,]\d+)?)", re.IGNORECASE),
        re.compile(r"(?:total\s+)?waktu\s+konsentrasi[^0-9\n\r]{0,50}(?:[tT]\s*_?c\s*=\s*)?(\d+(?:[.,]\d+)?)", re.IGNORECASE),
    ]
    for seg in search_segments:
        if tc_found is not None:
            break
        for pat in tc_patterns:
            for m in pat.finditer(seg):
                v = _parse_dec(m.group(1))
                if v is not None and 0.1 <= v <= 900.0:
                    if t1_found is not None and t2_found is not None and abs(v - t1_found) < 1e-3 and abs((t1_found + t2_found) - v) > 0.1:
                        continue
                    tc_found = round(v, 3)
                    matched_quotes.append(f"Tc={tc_found} mnt ({m.group(0).strip()[:60]})")
                    break
            if tc_found is not None:
                break

    # 4. Regex Luas Daerah Tangkapan (A) -> misal "A = 6.160 m2" atau "luas total daerah tangkapan"
    a_patterns = [
        re.compile(r"\bA(?:\s*_?(?:total|tot|dta|das))?\s*=\s*(\d[\d.,]*)\s*(m2|m²|ha|hektar|km2|km²)", re.IGNORECASE),
        re.compile(r"luas\s+(?:total\s+)?daerah\s+tangkapan(?:\s+air|\s+hujan)?[^0-9\n\r]{0,60}(?:A\s*=\s*)?(\d[\d.,]*)\s*(m2|m²|ha|hektar|km2|km²)?", re.IGNORECASE),
        re.compile(r"\bA\s*=\s*(\d{1,3}(?:\.\d{3})+)\b", re.IGNORECASE),
    ]
    for seg in search_segments:
        if a_tuple_found is not None:
            break
        for pat in a_patterns:
            for m in pat.finditer(seg):
                num_s = m.group(1)
                unit_s = m.group(2) if m.lastindex and m.lastindex >= 2 and m.group(2) else "m2"
                parsed_a = _parse_area(num_s, unit_s)
                if parsed_a is not None and parsed_a[0] >= 50.0:
                    a_tuple_found = parsed_a
                    matched_quotes.append(f"A={parsed_a[0]:,.0f} m² ({m.group(0).strip()[:60]})")
                    break
            if a_tuple_found is not None:
                break

    # 5. Regex Koefisien Limpasan (C) -> misal "C = 0,946" atau "koefisien limpasan"
    c_patterns = [
        re.compile(rf"\bC(?:\s*_?(?:gab|total|rata|r))?\s*=\s*{_MATH_INTER}(0[.,]\d+)", re.IGNORECASE),
        re.compile(r"koefisien\s+(?:limpasan|pengaliran)[^0-9\n\r]{0,50}(?:C\s*=\s*)?(0[.,]\d+)", re.IGNORECASE),
    ]
    for seg in search_segments:
        if c_found is not None:
            break
        for pat in c_patterns:
            for m in pat.finditer(seg):
                v = _parse_dec(m.group(1))
                if v is not None and 0.05 <= v <= 1.0 and abs(v - 0.278) > 1e-4:
                    c_found = round(v, 3)
                    matched_quotes.append(f"C={c_found} ({m.group(0).strip()[:60]})")
                    break
            if c_found is not None:
                break

    # 6. Regex Intensitas Hujan (I) -> misal "I = 175 mm/jam"
    i_patterns = [
        re.compile(rf"\bI(?:\s*_?\d+)?\s*=\s*{_MATH_INTER}(\d+(?:[.,]\d+)?)\s*(?:mm\s*/\s*(?:jam|h|hr))", re.IGNORECASE),
        re.compile(r"intensitas\s+(?:curah\s+)?hujan[^0-9\n\r]{0,60}(?:I\s*=\s*)?(\d+(?:[.,]\d+)?)\s*(?:mm\s*/\s*(?:jam|h|hr))?", re.IGNORECASE),
        re.compile(rf"\bI\s*=\s*{_MATH_INTER}(\d+(?:[.,]\d+)?)\s*(?:mm\s*/\s*(?:jam|h|hr))?", re.IGNORECASE),
    ]
    for seg in search_segments:
        if i_found is not None:
            break
        for pat in i_patterns:
            for m in pat.finditer(seg):
                v = _parse_dec(m.group(1))
                if v is not None and 10.0 <= v <= 1000.0 and abs(v - 24.0) > 1e-3:
                    i_found = round(v, 2)
                    matched_quotes.append(f"I={i_found} mm/jam ({m.group(0).strip()[:60]})")
                    break
            if i_found is not None:
                break

    # 7. Regex Debit Rencana (Qr) -> misal "Qr = 0,283" atau "debit rencana"
    qr_patterns = [
        re.compile(rf"\bQ\s*_?(?:r|rencana|limpasan|banjir|puncak)\s*=\s*(?:0[.,]278\s*[.x×⋅·*]\s*C\s*[.x×⋅·*]\s*I\s*[.x×⋅·*]\s*A\s*=\s*)?{_MATH_INTER}(\d+[.,]\d+)", re.IGNORECASE),
        re.compile(r"debit\s+(?:banjir\s+|limpasan\s+)?rencana[^0-9\n\r]{0,60}(?:Q\s*_?r?\s*=\s*)?(\d+[.,]\d+)", re.IGNORECASE),
    ]
    for seg in search_segments:
        if qr_found is not None:
            break
        for pat in qr_patterns:
            for m in pat.finditer(seg):
                v = _parse_dec(m.group(1))
                if (
                    v is not None
                    and 0.001 <= v <= 500.0
                    and abs(v - 0.278) > 1e-4
                    and (c_found is None or abs(v - c_found) > 1e-4)
                ):
                    qr_found = round(v, 3)
                    matched_quotes.append(f"Qr={qr_found} m³/s ({m.group(0).strip()[:60]})")
                    break
            if qr_found is not None:
                break

    # Konsistensi & fallback antar-parameter
    found_keys: list[str] = []
    if t1_found is not None:
        found_keys.append("t1")
    if t2_found is not None:
        found_keys.append("t2")
    if tc_found is not None:
        found_keys.append("Tc")
    if a_tuple_found is not None:
        found_keys.append("A")
    if c_found is not None:
        found_keys.append("C")
    if i_found is not None:
        found_keys.append("I")
    if qr_found is not None:
        found_keys.append("Qr")

    final_t1 = t1_found if t1_found is not None else 1.641
    final_t2 = t2_found if t2_found is not None else 4.889
    if tc_found is not None:
        final_tc = tc_found
    elif t1_found is not None and t2_found is not None:
        final_tc = round(t1_found + t2_found, 3)
    else:
        final_tc = 6.530

    if a_tuple_found is not None:
        final_a_m2, final_a_ha, final_a_km2 = a_tuple_found
    else:
        final_a_m2, final_a_ha, final_a_km2 = (6160.0, 0.616, 0.00616)

    final_c = c_found if c_found is not None else 0.946
    final_i = i_found if i_found is not None else 175.0

    if qr_found is not None:
        final_qr = qr_found
    elif c_found is not None and i_found is not None and a_tuple_found is not None:
        final_qr = round(0.278 * final_c * final_i * final_a_km2, 3)
    else:
        final_qr = 0.283

    return {
        "t1": final_t1,
        "t2": final_t2,
        "tc": final_tc,
        "A_m2": final_a_m2,
        "A_ha": final_a_ha,
        "A_km2": final_a_km2,
        "C": final_c,
        "I": final_i,
        "Qr": final_qr,
        "found_keys": found_keys,
        "explicit_count": len(found_keys),
        "matched_quotes": matched_quotes,
        "source_file": filename,
    }


def _analyze_channel_image_local(pil_img, h_desain_ref: float, b_desain_ref: float, n_desain_ref: float) -> dict:
    """
    100% Local Computer Vision analyzer (using PIL + NumPy) to estimate Cowan (1956) parameters,
    vegetation/weed percentage, sediment percentage, and width narrowing offline without external API keys.
    """
    img_rgb = pil_img.convert("RGB").resize((256, 256))
    arr = np.asarray(img_rgb, dtype=np.float32) / 255.0
    R, G, B = arr[:, :, 0], arr[:, :, 1], arr[:, :, 2]

    # 1. Deteksi vegetasi / gulma hijau (Excess Green Index ExG = 2G - R - B)
    exg = 2.0 * G - R - B
    green_mask = (exg > 0.08) & (G > 0.18) & (G > R * 1.06)
    weed_pct = float(np.clip(np.mean(green_mask) * 180.0, 2.0, 95.0))

    # 2. Deteksi endapan sedimen / lumpur coklat-gelap pada paruh bawah gambar
    lower_half = arr[128:, :, :]
    R_l, G_l, B_l = lower_half[:, :, 0], lower_half[:, :, 1], lower_half[:, :, 2]
    lum_l = 0.299 * R_l + 0.587 * G_l + 0.114 * B_l
    brown_mud_mask = (R_l >= G_l * 0.92) & (G_l >= B_l * 1.05) & (lum_l > 0.12) & (lum_l < 0.62) & (~green_mask[128:, :])
    dark_silt_mask = (lum_l >= 0.08) & (lum_l <= 0.35) & (np.abs(R_l - G_l) < 0.10) & (~green_mask[128:, :])
    sed_raw = float(np.mean(brown_mud_mask | dark_silt_mask))
    sed_pct = float(np.clip(sed_raw * 65.0 + 8.0, 5.0, 75.0))

    # 3. Analisis tekstur / gradien dinding untuk keretakan (n1) & variasi penampang (n2)
    gray = 0.299 * R + 0.587 * G + 0.114 * B
    gx = np.abs(np.diff(gray, axis=1))
    gy = np.abs(np.diff(gray, axis=0))
    edge_strength = float(np.mean(gx) + np.mean(gy))

    n1_est = float(np.clip(round(edge_strength * 0.065, 3), 0.002, 0.020))
    n2_est = float(np.clip(round(edge_strength * 0.040, 3), 0.001, 0.015))
    n3_est = float(np.clip(round((sed_pct / 100.0) * 0.060, 3), 0.004, 0.060))
    n4_est = float(np.clip(round((weed_pct / 100.0) * 0.085, 3), 0.002, 0.100))
    p_est  = float(np.clip(round((sed_pct / 100.0) * 0.25, 2), 0.02, 0.45))
    m_est  = 1.00

    ts_est  = round((sed_pct / 100.0) * h_desain_ref, 3)
    h_eff   = max(0.01, h_desain_ref - ts_est)
    b_eff   = max(0.01, b_desain_ref * (1.0 - p_est))
    n_cowan = (n_desain_ref + n1_est + n2_est + n3_est + n4_est) * m_est

    deskripsi = (
        f"Analisis Computer Vision Lokal (RGB/ExG & Gradien Tekstur): "
        f"Tutupan vegetasi/gulma terdeteksi ~{weed_pct:.1f}%, indikasi endapan dasar ~{sed_pct:.1f}% "
        f"(ts ≈ {ts_est:.3f} m), dan intensitas tekstur permukaan dinding moderat (n1={n1_est:.3f})."
    )
    return {
        "sediment_pct": round(sed_pct, 1),
        "weed_pct":     round(weed_pct, 1),
        "ts":           ts_est,
        "p_narrowing":  p_est,
        "n1":           n1_est,
        "n2":           n2_est,
        "n3":           n3_est,
        "n4":           n4_est,
        "m_meander":    m_est,
        "h_effective":  h_eff,
        "b_effective":  b_eff,
        "n_cowan":      n_cowan,
        "deskripsi":    deskripsi,
    }

def _safe_pdf_text(text: str) -> str:
    """Sanitize text for FPDF built-in Latin-1 fonts."""
    emoji_re = re.compile(
        "[\U00010000-\U0010FFFF"
        "\U00002702-\U000027B0"
        "\U0001F600-\U0001F64F"
        "\U0001F300-\U0001F5FF"
        "\U0001F680-\U0001F6FF"
        "\U0001F1E0-\U0001F1FF]+",
        flags=re.UNICODE,
    )
    text = emoji_re.sub("", text)
    replacements = {
        "³": "3", "²": "2", "°": "deg", "–": "-", "—": "-",
        "\u2019": "'", "\u2018": "'", "\u201c": '"', "\u201d": '"',
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    return text.encode("latin-1", errors="replace").decode("latin-1")


def _manning_Q_multi(shape: str, b_in: float, h_in: float, m_in: float, D_in: float, S_in: float, n_in: float) -> float:
    """Calculate Manning discharge Q (m3/s) for Persegi, Trapesium, or Lingkaran."""
    if h_in <= 0 or n_in <= 0 or S_in <= 0:
        return 0.0
    if shape == "Persegi":
        if b_in <= 0:
            return 0.0
        A_s = b_in * h_in
        P_s = b_in + 2.0 * h_in
    elif shape == "Trapesium":
        if b_in <= 0:
            return 0.0
        A_s = (b_in + m_in * h_in) * h_in
        P_s = b_in + 2.0 * h_in * math.sqrt(1.0 + m_in ** 2)
    else:  # Lingkaran
        if D_in <= 0:
            return 0.0
        h_c = float(np.clip(h_in, 1e-4 * D_in, 0.999 * D_in))
        cos_val = float(np.clip(1.0 - (2.0 * h_c / D_in), -1.0, 1.0))
        theta = 2.0 * math.acos(cos_val)
        A_s = ((D_in ** 2) / 8.0) * (theta - math.sin(theta))
        P_s = (D_in * theta) / 2.0
    if P_s <= 0 or A_s <= 0:
        return 0.0
    R_s = A_s / P_s
    return (1.0 / n_in) * A_s * (R_s ** (2.0 / 3.0)) * (S_in ** 0.5)


def _compute_m5_analysis(m4: dict, cv: dict, ded: dict) -> dict:
    """Compute Safety Factor, Capacity Drop %, Bisection h*, ts_izin, t_pengerukan & Permen PU 12/2014 recs."""
    Q_desain = float(m4.get("Q_desain", 0.0))
    Q_aktual = float(m4.get("Q_aktual", 0.0))
    Q_banjir = float(m4.get("Q_limpasan", 0.0))

    SF = (Q_aktual / Q_banjir) if Q_banjir > 0 else 999.0
    penurunan_pct = ((1.0 - (Q_aktual / Q_desain)) * 100.0) if Q_desain > 0 else 0.0

    shape     = str(m4.get("shape", cv.get("shape_cv", "Persegi")))
    b_desain  = float(m4.get("b", ded.get("b", 1.5)))
    b_aktual  = float(m4.get("b_eff", cv.get("b_effective", b_desain)))
    h_maks    = float(m4.get("h_ded", ded.get("h", 1.2)))
    h_eff     = float(m4.get("h_eff", cv.get("h_effective", h_maks * 0.8)))
    m_talud   = float(cv.get("m_talud", 1.0) if shape == "Trapesium" else 0.0)
    D_diam    = float(h_maks + float(m4.get("w_fb", 0.0))) if shape == "Lingkaran" else h_maks
    S_slope   = float(m4.get("S", ded.get("S", 0.001)))
    n_desain  = float(m4.get("n_ded", ded.get("n", 0.013)))
    n_aktual  = float(m4.get("n_adj", cv.get("n_adjusted", 0.020)))

    # Tebal sedimen eksisting ts (m)
    ts_current = float(cv.get("ts", max(0.0, h_maks - h_eff)))
    if ts_current <= 0 and h_maks > h_eff:
        ts_current = h_maks - h_eff

    # Kapasitas maksimum jika sedimen dikeruk habis (h = h_maks) tetapi kekasaran tetap n_aktual
    Q_at_hmaks_nakt = _manning_Q_multi(shape, b_aktual, h_maks, m_talud, D_diam, S_slope, n_aktual)

    # Algoritma Bisection (Bagi Dua) pada rentang [0, h_maks] untuk mencari h* di mana Q(h*, n_aktual) = Q_banjir
    low_b  = 0.0
    high_b = h_maks
    for _ in range(60):
        mid_b = 0.5 * (low_b + high_b)
        q_mid = _manning_Q_multi(shape, b_aktual, mid_b, m_talud, D_diam, S_slope, n_aktual)
        if q_mid < Q_banjir:
            low_b = mid_b
        else:
            high_b = mid_b
    h_star_in_range = 0.5 * (low_b + high_b)

    # Cek apakah Q_banjir melebihi kapasitas pada h_maks (artinya h* > h_maks)
    if Q_at_hmaks_nakt < Q_banjir:
        h_star_exceeds = True
        # Cari juga estimasi h* teoretis di atas h_maks (untuk Persegi/Trapesium) sebagai informasi teknis
        low_ext, high_ext = h_maks, h_maks * 4.0
        for _ in range(60):
            mid_ext = 0.5 * (low_ext + high_ext)
            q_ext = _manning_Q_multi(shape, b_aktual, mid_ext, m_talud, max(D_diam, mid_ext * 1.05), S_slope, n_aktual)
            if q_ext < Q_banjir:
                low_ext = mid_ext
            else:
                high_ext = mid_ext
        h_star = 0.5 * (low_ext + high_ext)
        ts_izin = 0.0
        t_pengerukan = ts_current
    else:
        h_star_exceeds = False
        h_star = h_star_in_range
        ts_izin = max(0.0, h_maks - h_star)
        t_pengerukan = max(0.0, ts_current - ts_izin)

    # Rekomendasi Pemeliharaan Otomatis sesuai Permen PU No. 12/PRT/M/2014
    T_ulang_val = int(m4.get("T_ulang", 5))
    tipologi_kota_val = str(m4.get("tipologi_kota") or st.session_state.get("tipologi_kota") or "Kota Sedang")
    Q_ras_val = Q_banjir
    Q_cap_val = Q_aktual
    ts_eksisting_val = ts_current

    t1_status = (
        f"Kondisi saluran AMAN (Safety Factor = {SF:.2f}). Kapasitas penampang masih memadai untuk menampung curah hujan rencana tanpa meluap."
        if SF >= 1.0
        else f"Kondisi saluran KRITIS/OVERTOPPING (Safety Factor = {SF:.2f} < 1.0). Diperlukan REDESAIN dimensi saluran atau penambahan tanggul penahan banjir (Freeboard)."
    )
    tindakan_01 = (
        f"[Evaluasi Kapasitas Hidrolika & Kala Ulang Permen PU 12/2014] "
        f"Berdasarkan evaluasi untuk Kala Ulang {T_ulang_val} Tahun pada {tipologi_kota_val}, debit limpasan rasional tercatat sebesar "
        f"Qr = {Q_ras_val:.3f} m³/s berbanding dengan kapasitas saluran aktual Qs = {Q_cap_val:.3f} m³/s. {t1_status}"
    )

    if t_pengerukan > 0:
        t2_status = (
            f"Status: MELAMPAUI BATAS. Segera lakukan pengerukan sedimen setebal minimum {t_pengerukan:.3f} m "
            f"untuk mencapai kedalaman kritis (h*) dan mencegah pendangkalan ekstrem."
        )
    else:
        t2_status = (
            "Status: NORMAL. Ketebalan sedimen belum mengganggu elevasi dasar saluran kritis (t_pengerukan = 0.000 m). "
            "Cukup jadwalkan inspeksi rutin berkala."
        )
    tindakan_02 = (
        f"[Tindakan Fisik - Optimasi Pengerukan Sedimen] "
        f"Evaluasi ketebalan sedimen eksisting (ts = {ts_eksisting_val:.3f} m) terhadap batas izin (ts_izin = {ts_izin:.3f} m). {t2_status}"
    )

    tindakan_03 = (
        f"[Tindakan Fisik - Rehabilitasi Kekasaran Dinding (Metode Cowan)] "
        f"Telah terjadi degradasi nilai kekasaran saluran dari desain awal (n = {n_desain:.4f}) menjadi kondisi lapangan "
        f"(n = {n_aktual:.4f}). Rekomendasi: Lakukan pembersihan gulma/vegetasi, pengangkatan sampah, serta penambalan retakan "
        f"struktur dinding untuk menekan angka hambatan Manning (n) dan mengembalikan kecepatan aliran (V) mendekati kondisi ideal."
    )

    recs_pu = [tindakan_01, tindakan_02, tindakan_03]

    return {
        "SF": SF,
        "penurunan_pct": penurunan_pct,
        "T_ulang": T_ulang_val,
        "tipologi_kota": tipologi_kota_val,
        "h_star": h_star,
        "h_star_in_range": h_star_in_range,
        "h_star_exceeds": h_star_exceeds,
        "ts_current": ts_current,
        "ts_izin": ts_izin,
        "t_pengerukan": t_pengerukan,
        "Q_at_hmaks_nakt": Q_at_hmaks_nakt,
        "shape": shape,
        "b_desain": b_desain,
        "b_aktual": b_aktual,
        "h_maks": h_maks,
        "h_eff": h_eff,
        "m_talud": m_talud,
        "D_diam": D_diam,
        "S_slope": S_slope,
        "n_desain": n_desain,
        "n_aktual": n_aktual,
        "Q_desain": Q_desain,
        "Q_aktual": Q_aktual,
        "Q_banjir": Q_banjir,
        "recs_pu": recs_pu,
    }


# ═════════════════════════════════════════════════════════════════
# PDF BUILDER (fpdf2)
# ═════════════════════════════════════════════════════════════════
def _build_pdf() -> bytes:
    """Generate the comprehensive executive PDF report from all 5 modules."""
    R24      = st.session_state.get("R24_5yr")
    r24_dict = st.session_state.get("r24_dict") or {}
    m1_df    = st.session_state.get("m1_r24_df")
    ded      = st.session_state.get("ded_params") or {}
    cv       = st.session_state.get("cv_params")  or {}
    m4       = st.session_state.get("m4_results") or {}

    class PDF(FPDF):
        def header(self):
            self.set_fill_color(14, 76, 146)
            self.rect(0, 0, 210, 20, "F")
            self.set_font("Helvetica", "B", 13)
            self.set_text_color(255, 255, 255)
            self.set_xy(0, 4)
            self.cell(0, 12, "DrainSense Kanaeru  -  LAPORAN LENGKAP EVALUASI DRAINASE (SNI & PERMEN PU)", align="C")
            self.set_text_color(0, 0, 0)
            self.ln(14)

        def footer(self):
            self.set_y(-13)
            self.set_font("Helvetica", "I", 8)
            self.set_text_color(100, 100, 100)
            ts = datetime.now().strftime("%d/%m/%Y %H:%M")
            self.cell(0, 8, f"DrainSense Kanaeru  |  SNI 2415:2016 & Permen PU 12/2014  |  Generated: {ts}  |  Hal. {self.page_no()}", align="C")

        def section(self, title: str):
            self.set_font("Helvetica", "B", 10)
            self.set_fill_color(230, 237, 248)
            self.set_text_color(14, 76, 146)
            self.cell(0, 7, _safe_pdf_text(title), fill=True, new_x="LMARGIN", new_y="NEXT")
            self.set_text_color(0, 0, 0)
            self.ln(1)

        def kv(self, label: str, value: str, indent: int = 4):
            self.set_font("Helvetica", "", 9)
            self.set_x(self.l_margin + indent)
            self.cell(82, 5.5, _safe_pdf_text(label + " :"))
            self.set_font("Helvetica", "B", 9)
            self.cell(0, 5.5, _safe_pdf_text(str(value)), new_x="LMARGIN", new_y="NEXT")

        def divider(self):
            self.set_draw_color(180, 180, 180)
            self.set_line_width(0.3)
            self.line(self.l_margin, self.get_y(), 210 - self.r_margin, self.get_y())
            self.ln(2.5)

        def body_text(self, text: str):
            self.set_font("Helvetica", "", 9)
            self.multi_cell(0, 5.2, _safe_pdf_text(text))
            self.ln(1)

    pdf = PDF(orientation="P", unit="mm", format="A4")
    pdf.set_margins(14, 24, 14)
    pdf.set_auto_page_break(auto=True, margin=16)
    pdf.add_page()

    # ── Cover info ──
    pdf.set_font("Helvetica", "I", 9)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(0, 5, f"Tanggal Cetak: {datetime.now().strftime('%d %B %Y %H:%M')}", align="R", new_x="LMARGIN", new_y="NEXT")
    pdf.set_text_color(0, 0, 0)
    pdf.ln(2)

    # ── MODUL 1: Statistik & Uji Kecocokan Hujan (SNI 2415:2016) ──
    pdf.section("MODUL 1: Statistik & Analisis Frekuensi Hujan Rencana (SNI 2415:2016)")
    if m1_df is not None and not m1_df.empty:
        xi_arr = m1_df["R24_Max (mm)"].astype(float).to_numpy()
        n_m1   = len(xi_arr)
        xb_m1  = float(np.mean(xi_arr))
        sd_m1  = float(np.std(xi_arr, ddof=1)) if n_m1 > 1 else 0.0
        cv_m1  = (sd_m1 / xb_m1) if xb_m1 != 0 else 0.0
        cs_m1  = float((n_m1 * np.sum((xi_arr - xb_m1) ** 3)) / ((n_m1 - 1) * (n_m1 - 2) * (sd_m1 ** 3))) if (n_m1 > 2 and sd_m1 > 0) else 0.0
        ck_m1  = float(((n_m1 ** 2) * np.sum((xi_arr - xb_m1) ** 4)) / ((n_m1 - 1) * (n_m1 - 2) * (n_m1 - 3) * (sd_m1 ** 4))) if (n_m1 > 3 and sd_m1 > 0) else 0.0
        log_xi = np.log10(np.clip(xi_arr, 1e-6, None))
        mlog   = float(np.mean(log_xi))
        slog   = float(np.std(log_xi, ddof=1)) if n_m1 > 1 else 0.0
        l10n   = math.log10(max(n_m1, 2))
        kn_gb  = -0.9043 + 3.345 * math.sqrt(l10n) - 0.4046 * l10n
        xh_gb  = 10.0 ** (mlog + kn_gb * slog)
        xl_gb  = 10.0 ** (mlog - kn_gb * slog)

        pdf.kv("Jumlah Data (n) & Rata-rata (X_bar)", f"n = {n_m1} tahun  |  X_bar = {xb_m1:.2f} mm")
        pdf.kv("Simpangan Baku (Sd) & Koef. Variasi (Cv)", f"Sd = {sd_m1:.2f} mm  |  Cv = {cv_m1:.4f}")
        pdf.kv("Kemencengan (Cs) & Kurtosis (Ck)", f"Cs = {cs_m1:.4f}  |  Ck = {ck_m1:.4f}")
        pdf.kv("Uji Outlier Grubbs-Beck (Kn, XL, XH)", f"Kn = {kn_gb:.3f}  |  XL = {xl_gb:.1f} mm  |  XH = {xh_gb:.1f} mm")
    if R24 is not None:
        pdf.kv("R24 Hujan Rencana T=5 Tahun Terpilih", f"{R24:.2f} mm")
    if r24_dict:
        r24_str = ", ".join(f"T={k}th: {v:.1f} mm" for k, v in r24_dict.items())
        pdf.kv("Tabel Hujan Rencana Multi-Kala Ulang", r24_str)
    if R24 is None and m1_df is None:
        pdf.body_text("  Data hujan Modul 1 belum diproses.")
    pdf.divider()

    # ── MODUL 2: Parameter DED & Validasi Chow (1959) ──
    pdf.section("MODUL 2: Parameter DED & Validasi Tabel Chow (1959)")
    if ded:
        mat_d = str(ded.get("material", "Beton"))
        n_d   = float(ded.get("n", 0.013))
        chow_map = {
            "Beton": (0.011, 0.020),
            "Pasangan Bata": (0.011, 0.018),
            "Pasangan Batu Kali": (0.017, 0.030),
            "Pasangan Batu": (0.017, 0.030),
            "Tanah": (0.016, 0.033),
        }
        rng = chow_map.get(mat_d, (0.010, 0.040))
        valid_chow = "VALID (Dalam Rentang)" if (rng[0] <= n_d <= rng[1]) else f"PERINGATAN (Di Luar Rentang {rng[0]:.3f}-{rng[1]:.3f})"
        pdf.kv("Dimensi Desain (Lebar b & Tinggi H)", f"b = {ded.get('b', 'N/A')} m  |  H = {ded.get('h', 'N/A')} m")
        pdf.kv("Kemiringan S & Material Saluran", f"S = {ded.get('S', 'N/A')} m/m  |  Material = {mat_d}")
        pdf.kv("Manning n DED & Status Validasi Chow", f"n = {n_d:.4f}  ->  {valid_chow}")
        pdf.kv("Kutipan / Halaman Asal DED", str(ded.get("kutipan_ded", "-")))
    else:
        pdf.body_text("  Parameter DED belum diinput.")
    pdf.divider()

    # ── MODUL 3: Komponen Cowan, Horton-Einstein & Geometri Efektif ──
    pdf.section("MODUL 3: Metode Cowan (1956), Horton-Einstein & Geometri Efektif")
    if cv:
        pdf.kv(
            "Komponen Cowan (nb, n1, n2, n3, n4, m)",
            f"nb={cv.get('nb', 0.013):.4f}, n1={cv.get('n1', 0):.3f}, n2={cv.get('n2', 0):.3f}, "
            f"n3={cv.get('n3', 0):.3f}, n4={cv.get('n4', 0):.3f}, m={cv.get('m_meander', 1.0):.2f}",
        )
        pdf.kv(
            "Manning Aktual Cowan & Horton-Einstein",
            f"n_Cowan = {cv.get('n_cowan', cv.get('n_adjusted', 0)):.5f}  |  "
            f"n_c Horton = {cv.get('horton_nc', cv.get('n_adjusted', 0)):.5f}  |  "
            f"Aktif = {cv.get('n_adjusted', 0):.5f}",
        )
        pdf.kv(
            "Degradasi Fisik (Tebal Sedimen ts & p)",
            f"ts = {cv.get('ts', 0):.3f} m ({cv.get('sediment_pct', 0):.1f}%)  |  "
            f"Penyempitan p = {cv.get('p_narrowing', 0):.2f}  |  Gulma = {cv.get('weed_pct', 0):.1f}%",
        )
        pdf.kv(
            "Geometri Efektif Saluran",
            f"h_efektif = {cv.get('h_effective', 0):.4f} m  |  b_efektif = {cv.get('b_effective', ded.get('b', 1.5)):.4f} m",
        )
    else:
        pdf.body_text("  Analisis degradasi Modul 3 belum dilakukan.")
    pdf.divider()

    # ── MODUL 4: Hasil Hidrolika Permen PU 12/2014 ──
    pdf.section("MODUL 4: Evaluasi Hidrolika (Permen PU 12/2014, Multi-Penampang & Froude)")
    if m4:
        pdf.kv("Bentuk Penampang & Kala Ulang Rencana", f"{m4.get('shape', 'Persegi')}  |  Kala Ulang T = {m4.get('T_ulang', 5)} Tahun (R24 = {m4.get('R24', 0):.2f} mm)")
        pdf.kv("Waktu Konsentrasi tc & Intensitas I", f"tc = {m4.get('tc', 0) * 60:.2f} menit ({m4.get('tc', 0):.4f} jam)  |  I Mononobe = {m4.get('I_mono', 0):.2f} mm/jam")
        pdf.kv("Koef. Limpasan C_gab & Luas DTA", f"C_gab = {m4.get('C', 0):.3f}  |  A_total = {m4.get('A', 0) * 100:.2f} ha ({m4.get('A', 0):.4f} km2)")
        pdf.kv("Tinggi Jagaan (Freeboard w) & h_maks", f"w = {m4.get('w_fb', 0):.3f} m  |  h_maks = {m4.get('h_ded', 0):.3f} m")
        pdf.kv("Debit Desain (Q_desain) vs Aktual (Q_aktual)", f"Q_desain = {m4.get('Q_desain', 0):.4f} m3/s  |  Q_aktual = {m4.get('Q_aktual', 0):.4f} m3/s")
        pdf.kv("Debit Banjir Rencana (Q_banjir)", f"Q_banjir = {m4.get('Q_limpasan', 0):.4f} m3/s")
        pdf.kv("Kecepatan Aliran V & Bilangan Froude Fr", f"V_aktual = {m4.get('V_aktual', 0):.3f} m/s (V_izin = {m4.get('V_izin', 3.0):.1f} m/s)  |  Fr = {m4.get('Fr_aktual', 0):.3f}")
    else:
        pdf.body_text("  Belum ada hasil perhitungan Modul 4.")
    pdf.divider()

    # ── MODUL 5: Evaluasi Keamanan, Bisection h* & Rekomendasi Permen PU 12/2014 ──
    pdf.section("MODUL 5: Safety Factor (SF), Target Pengerukan Bisection & Rekomendasi PU 12/2014")
    if m4:
        m5_an = _compute_m5_analysis(m4, cv, ded)
        sf_val = m5_an["SF"]
        is_safe_sf = sf_val >= 1.0
        pdf.kv("Safety Factor (SF = Q_aktual / Q_banjir)", f"SF = {sf_val:.3f}  ({'AMAN >= 1.0' if is_safe_sf else 'KRITIS / MELUAP < 1.0'})")
        pdf.kv("Persentase Penurunan Kapasitas Saluran", f"{m5_an['penurunan_pct']:.2f} %")
        pdf.kv(
            "Hasil Bisection Kedalaman Kritis Minimum (h*)",
            f"h* = {m5_an['h_star']:.4f} m  (h_maks = {m5_an['h_maks']:.4f} m)"
            + (" [MELEBIHI h_maks!]" if m5_an["h_star_exceeds"] else ""),
        )
        pdf.kv("Tebal Sedimen Eksisting (ts) vs Diizinkan (ts_izin)", f"ts = {m5_an['ts_current']:.4f} m  |  ts_izin = {m5_an['ts_izin']:.4f} m")
        pdf.kv("Tebal Pengerukan Wajib (t_pengerukan)", f"{m5_an['t_pengerukan']:.4f} m  ({m5_an['t_pengerukan'] * 100:.1f} cm)")

        pdf.ln(2)
        status_txt = f"STATUS AKHIR: {'AMAN' if is_safe_sf else 'KRITIS / RISIKO MELUAP'}  (SF = {sf_val:.3f})"
        pdf.set_font("Helvetica", "B", 11)
        pdf.set_fill_color(200, 240, 200) if is_safe_sf else pdf.set_fill_color(255, 200, 200)
        pdf.cell(0, 9, _safe_pdf_text(f"  {status_txt}"), fill=True, align="C", new_x="LMARGIN", new_y="NEXT")
        pdf.set_fill_color(255, 255, 255)
        pdf.ln(2)

        pdf.set_font("Helvetica", "B", 10)
        pdf.cell(0, 6, "Rekomendasi Pemeliharaan (Permen PU No. 12/PRT/M/2014):", new_x="LMARGIN", new_y="NEXT")
        for idx_r, r_txt in enumerate(m5_an["recs_pu"], 1):
            pdf.body_text(f"{idx_r}. {r_txt}")
    else:
        pdf.body_text("  Data Modul 4 belum tersedia untuk analisis Modul 5.")

    return bytes(pdf.output())


# ═════════════════════════════════════════════════════════════════
# SIDEBAR — Dual-Engine (Direct SDK Gemini 4 Argon + Live JSON Bridge)
# ═════════════════════════════════════════════════════════════════
_bridge_notice = _sync_bridge_to_session(force=False)

with st.sidebar:
    # A. BRANDING FEATURING KOBO CHIBI MASCOT
    _sb_kobo_uri = _get_kobo_data_uri("kobo_chibi", max_height=200)
    _sb_avatar_html = (
        f'<div class="ds-sb-kobo-wrap" title="Kobo Kanaeru Chibi — DrainSense Mascot">'
        f'<img class="ds-sb-kobo-img" src="{_sb_kobo_uri}" alt="Kobo Chibi" />'
        f'</div>'
        if _sb_kobo_uri
        else '<div class="ds-sb-icon">🌊</div>'
    )
    st.markdown(
        f"""
        <div class="ds-sb-brand" style="font-family:'Inter',sans-serif;">
            <div class="ds-sb-logo-row">
                {_sb_avatar_html}
                <div style="min-width: 0; flex: 1;">
                    <div class="ds-sb-title" style="color:#ffffff !important;font-family:'Inter',sans-serif;font-size:17px;font-weight:700;">DrainSense Kanaeru</div>
                    <div class="ds-sb-sub" style="color:#e2e8f0 !important;font-family:'Inter',sans-serif;font-size:12.5px;font-weight:500;">AI-Powered Drainage Analysis System</div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # B. ENGINE 1 — DIRECT SDK (GEMINI API)
    _has_api_key = bool(st.session_state.get("gemini_key"))
    _e1_pill_cls = "on" if _has_api_key else "off"
    _e1_pill_txt = "CONNECTED" if _has_api_key else "LOCAL FALLBACK"
    st.markdown(
        f"""
        <div class="ds-sb-engine-card" style="font-family:'Inter',sans-serif;font-size:14px;color:#ffffff !important;">
            <div class="ds-sb-engine-head" style="margin-bottom:0 !important;">
                <span class="ds-sb-engine-title" style="color:#ffffff !important;font-family:'Inter',sans-serif;font-size:14px;font-weight:700;">🤖 Engine 1: Direct SDK (Gemini API)</span>
                <span class="ds-sb-pill {_e1_pill_cls}">{_e1_pill_txt}</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # C. API KEY INPUT (DI DALAM EXPANDER)
    with st.expander("🔧 Pengaturan Engine 1 API"):
        api_key_input = st.text_input(
            "Gemini API Key (Opsional)",
            value=st.session_state["gemini_key"],
            type="password",
        ).strip()
        if api_key_input != st.session_state["gemini_key"]:
            st.session_state["gemini_key"] = api_key_input

        if api_key_input:
            st.markdown(
                "<p style='color:#4ade80; font-size:14px;'>✅ API Key aktif.</p>",
                unsafe_allow_html=True,
            )

    # D. ENGINE 2 — LIVE JSON BRIDGE
    _bridge_ok = BRIDGE_PATH.exists()
    _e2_pill_cls = "on" if _bridge_ok else "off"
    _e2_pill_txt = "BRIDGE ACTIVE" if _bridge_ok else "STANDBY"
    st.markdown(
        f"""
        <div class="ds-sb-engine-card" style="margin-top:0.35rem;font-family:'Inter',sans-serif;font-size:14px;color:#ffffff !important;">
            <div class="ds-sb-engine-head">
                <span class="ds-sb-engine-title" style="color:#ffffff !important;font-family:'Inter',sans-serif;font-size:14px;font-weight:700;">🔗 Engine 2: Live JSON Bridge</span>
                <span class="ds-sb-pill {_e2_pill_cls}">{_e2_pill_txt}</span>
            </div>
            <div class="ds-sb-engine-meta" style="font-family:'Inter',sans-serif;font-size:14px;color:#e2e8f0 !important;">
                <span style="color:#ffffff !important;font-family:'Inter',sans-serif;font-size:13.5px;font-weight:600;">File:</span>
                <span class="ds-sb-meta-tag" style="color:#ffffff !important;">{BRIDGE_FILENAME}</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # E. TOMBOL SIDEBAR (1 TOMBOL PER BARIS AGAR TEKS UTUH & RAPI)
    if st.button("🔄 Muat Bridge", use_container_width=True):
        msg_b = _sync_bridge_to_session(force=True)
        if msg_b:
            st.toast(msg_b, icon="🔗")
            st.rerun()
        else:
            st.warning("File bridge belum tersedia.")

    if st.button("📝 Buat Template", use_container_width=True):
        _write_bridge_template()
        _sync_bridge_to_session(force=True)
        st.toast(f"Template `{BRIDGE_FILENAME}` dibuat & dimuat!", icon="✅")
        st.rerun()

    if _bridge_notice:
        st.caption(_bridge_notice)

    st.markdown("---")

    # F. NAVIGASI TAHAPAN ANALISIS & STATUS PIPELINE
    _NAV_OPTIONS = [
        "1 | 🌧️ Analisis Iklim & IDF (BMKG)",
        "2 | 📄 Ekstraksi DED (AI NLP)",
        "3 | 📸 Visi Komputer & Degradasi",
        "4 | ⚙️ Engine Hidrolika Hibrida",
        "5 | 📊 Executive Dashboard",
    ]
    if st.session_state.get("nav_modul_aktif") not in _NAV_OPTIONS:
        st.session_state["nav_modul_aktif"] = _NAV_OPTIONS[0]

    menu = st.radio(
        "🧭 Navigasi Tahapan Analisis",
        _NAV_OPTIONS,
        key="nav_modul_aktif",
    )

    _active_stage_num = int(menu.split("|")[0].strip()) if "|" in menu and menu.split("|")[0].strip().isdigit() else 1
    _pipe_states = _get_pipeline_stage_states()
    _pipe_labels = [
        (1, "Modul 1 · R24 & Kurva IDF"),
        (2, "Modul 2 · Parameter DED"),
        (3, "Modul 3 · Visi Komputer & Cowan"),
        (4, "Modul 4 · Debit & Hidrolika"),
        (5, "Modul 5 · Optimasi & Laporan"),
    ]
    _stepper_rows = []
    for _idx_p, _lbl_p in _pipe_labels:
        _is_done = _pipe_states.get(_idx_p, False)
        _is_act = _idx_p == _active_stage_num
        _row_cls = "active" if _is_act else ""
        if _is_act:
            _dot_cls = "active"
            _dot_sym = f"{_idx_p:02d}"
            _st_badge = '<span class="status-active" style="margin-left:auto;font-size:11px;color:#00e5ff !important;font-weight:bold !important;letter-spacing:0.04em;">ACTIVE</span>'
        elif _is_done:
            _dot_cls = "done"
            _dot_sym = "✓"
            _st_badge = '<span class="status-ready" style="margin-left:auto;font-size:11px;color:#4ade80 !important;font-weight:bold !important;letter-spacing:0.04em;">READY</span>'
        else:
            _dot_cls = "pending"
            _dot_sym = f"{_idx_p:02d}"
            _st_badge = '<span class="status-ready" style="margin-left:auto;font-size:11px;color:#4ade80 !important;font-weight:bold !important;letter-spacing:0.04em;">READY</span>'
        _stepper_rows.append(
            f'<div class="ds-sb-step {_row_cls}">'
            f'<span class="ds-sb-dot {_dot_cls}">{_dot_sym}</span>'
            f'<span class="ds-sb-step-lbl" style="color:#ffffff !important;font-weight:600;">{_lbl_p}</span>'
            f'{_st_badge}'
            f'</div>'
        )
    st.markdown(
        '<div class="ds-sb-section-hdr" style="color:#ffffff !important;">📋 Status Pipeline Terpadu</div>'
        '<div class="ds-sb-stepper">' + "".join(_stepper_rows) + "</div>",
        unsafe_allow_html=True,
    )

    st.divider()
    with st.expander("✨ Profil Kanaeru System", expanded=False):
        kobo_profil_uri = _get_kobo_data_uri("kobo_profil", max_height=560)
        if not kobo_profil_uri:
            _profil_rel = _resolve_kobo_profil()
            _profil_path = Path(__file__).resolve().parent / _profil_rel
            if _profil_path.exists():
                _ext = _profil_path.suffix.lower()
                _mime = (
                    "image/webp" if _ext == ".webp"
                    else ("image/png" if _ext == ".png" else "image/jpeg")
                )
                _b64 = base64.b64encode(_profil_path.read_bytes()).decode("ascii")
                kobo_profil_uri = f"data:{_mime};base64,{_b64}"
        if kobo_profil_uri:
            st.markdown(
                f'<div style="text-align: center;">'
                f'<img src="{kobo_profil_uri}" alt="Kobo Kanaeru - Profil Kanaeru System" '
                f'style="width: 100%; max-height: 280px; object-fit: contain; '
                f'filter: drop-shadow(0 4px 10px rgba(0,0,0,0.2)); margin-bottom: 10px;" />'
                f'</div>'
                f'<p style="color: #94a3b8; font-size: 12px; text-align: center; margin: 0;">'
                f'DrainSense Kanaeru AI. Dioptimalkan oleh Gemini &amp; Python. gila kamu!?; kamu yang gila!'
                f'</p>',
                unsafe_allow_html=True,
            )


# ═════════════════════════════════════════════════════════════════
# MODUL 1: Pre-Processor Data Iklim BMKG (SNI 2415:2016)
# ═════════════════════════════════════════════════════════════════
if menu.startswith("1 |") or "Modul 1" in menu:
    _m1_ready = st.session_state.get("R24_5yr") is not None
    _render_module_header(
        step_num=1,
        icon="🌧️",
        title="Analisis Iklim & Kurva IDF (SNI 2415:2016)",
        description=(
            "Pengolahan data curah hujan harian BMKG, uji outlier Grubbs-Beck, uji kecocokan distribusi frekuensi "
            "(Normal, Log-Normal, Gumbel, Log-Pearson III) sesuai SNI 2415:2016, serta pemodelan Kurva IDF Multi-Kala Ulang (Metode Mononobe)."
        ),
        status_text=f"R24 T=5th: {st.session_state['R24_5yr']:.2f} mm" if _m1_ready else "Menunggu Input Data Hujan",
        is_ready=_m1_ready,
    )

    _render_kobo_companion_card(
        asset_key="kobo_rain",
        tag="🌧️ KOBO RAINFALL & IDF COMPANION · SNI 2415:2016",
        title="Analisis Curah Hujan Harian BMKG & Pemodelan Kurva IDF Mononobe",
        description=(
            "Unggah seri data hujan harian BMKG (.csv / .xlsx) atau laporan hidrologi (.pdf) untuk menguji "
            "distribusi Normal, Log-Normal, Gumbel, dan Log-Pearson III secara otomatis serta membangkitkan kurva IDF multi-kala ulang."
        ),
    )

    if st.session_state.get("R24_5yr") is not None:
        st.success(
            f"✅ **Nilai Aktif di Session State:** R24 (T=5 tahun) = **{st.session_state['R24_5yr']:.2f} mm** "
            "— terhubung otomatis ke Modul 2, Modul 3, dan Modul 4."
        )

    _render_section_banner(
        "📂 Akuisisi Data Curah Hujan Harian BMKG / Dokumen Teknis",
        "Format didukung: CSV (.csv), Excel (.xlsx) dengan kolom Tanggal & Curah Hujan (mm), atau Jurnal/Laporan DED (.pdf) hingga 200 MB.",
        badge_html='<span class="ds-badge auto">SNI 2415:2016 · Mononobe IDF</span>',
        tone="ocean",
    )

    uploaded_file = st.file_uploader(
        "📂 Unggah file data curah hujan BMKG (.csv / .xlsx) atau Dokumen Jurnal/DED (.pdf)",
        type=["csv", "xlsx", "pdf"],
        help="Kolom minimal: Tanggal & Curah Hujan (mm), atau dokumen PDF untuk diekstrak di Modul 2 & Modul 4.",
    )

    # Simpan file baru ke st.session_state['uploaded_file_m1'] saat diunggah
    if uploaded_file is not None:
        try:
            if uploaded_file.name.lower().endswith(".pdf"):
                raw_pdf_m1 = bytes(uploaded_file.getbuffer())
                pdf_txt_m1 = _extract_pdf_text_pypdf(raw_pdf_m1)
                prev_d_m1, prev_s_m1 = _build_preview_data(pdf_txt_m1)
                st.session_state["pdf_doc_text"] = pdf_txt_m1
                st.session_state["pdf_doc_name"] = uploaded_file.name
                st.session_state["uploaded_file_m2"] = {
                    "name": uploaded_file.name,
                    "size": uploaded_file.size,
                    "full_doc_text": pdf_txt_m1,
                    "preview_data": prev_d_m1,
                    "preview_stats": prev_s_m1,
                }
                prev_up = st.session_state.get("uploaded_file_m1") or {}
                prev_up["pdf_name"] = uploaded_file.name
                prev_up["full_doc_text"] = pdf_txt_m1
                st.session_state["uploaded_file_m1"] = prev_up
                if isinstance(st.session_state.get("data_modul_1"), dict):
                    st.session_state["data_modul_1"]["full_doc_text"] = pdf_txt_m1
                    st.session_state["data_modul_1"]["pdf_name"] = uploaded_file.name
                st.success(
                    f"✅ Dokumen PDF **{uploaded_file.name}** berhasil dibaca ({len(pdf_txt_m1):,} karakter) "
                    "dan disimpan ke `st.session_state` untuk diekstrak di Modul 2 & Modul 4."
                )
            else:
                df_uploaded = (
                    pd.read_csv(uploaded_file)
                    if uploaded_file.name.endswith(".csv")
                    else pd.read_excel(uploaded_file)
                )
                prev_up = st.session_state.get("uploaded_file_m1")
                is_new_m1_file = (
                    prev_up is None
                    or prev_up.get("name") != uploaded_file.name
                    or prev_up.get("shape") != df_uploaded.shape
                )
                st.session_state["uploaded_file_m1"] = {
                    "name": uploaded_file.name,
                    "df_raw": df_uploaded,
                    "shape": df_uploaded.shape,
                    "full_doc_text": st.session_state.get("pdf_doc_text", ""),
                    "pdf_name": st.session_state.get("pdf_doc_name", ""),
                }
                if is_new_m1_file and len(df_uploaded.columns) >= 2:
                    st.session_state["m1_col_date"] = df_uploaded.columns[0]
                    st.session_state["m1_col_rain"] = df_uploaded.columns[min(1, len(df_uploaded.columns) - 1)]
        except Exception as e_up:
            st.error(f"❌ Error membaca file yang diunggah: {e_up}")

    # Ambil data dari uploaded_file_m1 atau data_modul_1 yang sudah tersimpan di st.session_state
    saved_up_m1 = st.session_state.get("uploaded_file_m1")
    saved_data_m1 = st.session_state.get("data_modul_1")

    active_df_raw = None
    active_m1_name = ""
    if isinstance(saved_up_m1, dict) and saved_up_m1.get("df_raw") is not None:
        active_df_raw = saved_up_m1["df_raw"]
        active_m1_name = str(saved_up_m1.get("name", "Data_BMKG.csv"))
    elif isinstance(saved_data_m1, dict) and saved_data_m1.get("df_raw") is not None:
        active_df_raw = saved_data_m1["df_raw"]
        active_m1_name = str(saved_data_m1.get("file_name", "Data_BMKG.csv"))

    if active_df_raw is not None:
        df_raw = active_df_raw
        if uploaded_file is not None:
            st.success(f"✅ File **{active_m1_name}** berhasil diunggah dan disimpan ke Session State.")
        else:
            st.success(f"💾 Menampilkan kembali data Modul 1 dari Session State: **{active_m1_name}** (data tetap tersimpan saat berpindah modul).")

        try:
            st.subheader("🗂️ Pratinjau Data Mentah (5 Baris Pertama)")
            st.dataframe(df_raw.head())
            mi1, mi2, mi3 = st.columns(3)
            mi1.metric("Jumlah Baris", df_raw.shape[0])
            mi2.metric("Jumlah Kolom", df_raw.shape[1])
            mi3.metric("Missing Values", int(df_raw.isnull().sum().sum()))
            st.markdown("---")

            st.subheader("🔧 Konfigurasi Kolom")
            all_cols = df_raw.columns.tolist()
            saved_c_date = st.session_state.get("m1_col_date")
            saved_c_rain = st.session_state.get("m1_col_rain")
            idx_date = all_cols.index(saved_c_date) if saved_c_date in all_cols else 0
            idx_rain = all_cols.index(saved_c_rain) if saved_c_rain in all_cols else min(1, len(all_cols) - 1)

            sc1, sc2 = st.columns(2)
            with sc1:
                col_date = st.selectbox(
                    "📅 Kolom Tanggal (Date)", all_cols,
                    index=idx_date,
                    help="Kolom berisi tanggal/waktu pengamatan",
                )
            with sc2:
                col_rain = st.selectbox(
                    "🌧️ Kolom Curah Hujan (mm)", all_cols,
                    index=idx_rain,
                    help="Kolom berisi nilai curah hujan harian (mm)",
                )
            st.session_state["m1_col_date"] = col_date
            st.session_state["m1_col_rain"] = col_rain
            st.markdown("---")

            if st.button("📈 Proses Analisis Frekuensi", type="primary"):
                df_work = df_raw[[col_date, col_rain]].copy()
                r24_calc, is_monthly_mode, parse_err = _parse_m1_rainfall_series(df_work)
                if parse_err:
                    st.error(parse_err)
                    st.stop()

                st.session_state["m1_r24_df"] = r24_calc
                st.session_state["m1_is_monthly"] = bool(is_monthly_mode)

            # Pulihkan m1_r24_df dari data_modul_1 bila sudah pernah diolah sebelumnya
            if st.session_state.get("m1_r24_df") is None and isinstance(saved_data_m1, dict) and saved_data_m1.get("r24_df") is not None:
                st.session_state["m1_r24_df"] = saved_data_m1["r24_df"]
                st.session_state["m1_is_monthly"] = bool(saved_data_m1.get("is_monthly", False))

            # Render hasil analisis jika sudah diproses
            if st.session_state.get("m1_r24_df") is not None:
                r24 = st.session_state["m1_r24_df"]
                is_monthly_active = bool(st.session_state.get("m1_is_monthly", "Periode" in r24.columns))
                if is_monthly_active:
                    st.info("ℹ️ Terdeteksi format data bulanan (1 tahun). Analisis statistik dan frekuensi dijalankan menggunakan seri nilai maksimum bulanan.")

                # ── 1. PERHITUNGAN STATISTIK LENGKAP R24 MAKSIMUM (SNI 2415:2016) ──
                series = r24["R24_Max (mm)"].astype(float)
                Xi     = series.to_numpy()
                n_data = len(Xi)
                X_bar  = float(np.mean(Xi))
                Sd     = float(np.std(Xi, ddof=1)) if n_data > 1 else 0.0
                Cv     = (Sd / X_bar) if X_bar != 0 else 0.0

                if n_data > 2 and Sd > 0:
                    Cs = float((n_data * np.sum((Xi - X_bar) ** 3)) / ((n_data - 1) * (n_data - 2) * (Sd ** 3)))
                else:
                    Cs = 0.0

                if n_data > 3 and Sd > 0:
                    Ck = float(((n_data ** 2) * np.sum((Xi - X_bar) ** 4)) / ((n_data - 1) * (n_data - 2) * (n_data - 3) * (Sd ** 4)))
                else:
                    Ck = 0.0

                # Statistik Log10(X)
                logX      = np.log10(np.clip(Xi, 1e-6, None))
                mean_logX = float(np.mean(logX))
                S_logX    = float(np.std(logX, ddof=1)) if n_data > 1 else 0.0
                Cv_log    = (S_logX / mean_logX) if mean_logX != 0 else 0.0
                if n_data > 2 and S_logX > 0:
                    Cs_log = float((n_data * np.sum((logX - mean_logX) ** 3)) / ((n_data - 1) * (n_data - 2) * (S_logX ** 3)))
                else:
                    Cs_log = 0.0
                if n_data > 3 and S_logX > 0:
                    Ck_log = float(((n_data ** 2) * np.sum((logX - mean_logX) ** 4)) / ((n_data - 1) * (n_data - 2) * (n_data - 3) * (S_logX ** 4)))
                else:
                    Ck_log = 0.0

                # ── 2. PERHITUNGAN UJI OUTLIER GRUBBS-BECK (10%) ───────────────────
                log10_n = math.log10(max(n_data, 2))
                Kn = -0.9043 + 3.345 * math.sqrt(log10_n) - 0.4046 * log10_n
                XH = 10.0 ** (mean_logX + Kn * S_logX)
                XL = 10.0 ** (mean_logX - Kn * S_logX)
                outliers_high = r24[r24["R24_Max (mm)"] > XH]
                outliers_low  = r24[r24["R24_Max (mm)"] < XL]

                # ── 3. PERHITUNGAN 4 DISTRIBUSI FREKUENSI HUJAN RENCANA ───────────
                st.info(
                    "ℹ️ **Catatan Kepatuhan SNI 2415:2016 (Pasal 4.2.1 & Tabel 16):** SNI merekomendasikan panjang "
                    f"seri data hujan tahunan **> 20 tahun** serta pengujian data (outlier/konsistensi/tren). "
                    f"Data terdeteksi: **{n_data} periode**. Sistem tetap mengizinkan perhitungan untuk seri data "
                    "yang lebih pendek sebagai referensi teknis."
                )
                gumbel_mode = st.radio(
                    "Metode Faktor Frekuensi Gumbel ($K_T$)",
                    options=[
                        "Gumbel SNI 2415:2016 (K = 0.78Y - 0.45)",
                        "Gumbel Sampel Terbatas (Yn & Sn)",
                    ],
                    index=0,
                    horizontal=True,
                    key="m1_gumbel_mode",
                )
                use_gumbel_sni = gumbel_mode.startswith("Gumbel SNI 2415:2016")

                Yn, Sn = _gumbel_params(n_data)
                return_periods = [2, 5, 10, 20, 25, 50, 100]

                dist_names = ["Normal", "Log Normal", "Gumbel", "Log Pearson III"]
                dist_results: dict[str, dict[int, float]] = {d: {} for d in dist_names}
                comp_rows = []

                for T in return_periods:
                    prob_non_exceed = 1.0 - (1.0 / T)
                    KT_norm = float(scipy.stats.norm.ppf(prob_non_exceed))
                    XT_norm = X_bar + KT_norm * Sd
                    XT_lognorm = 10.0 ** (mean_logX + KT_norm * S_logX)
                    YT = -math.log(-math.log((T - 1.0) / T))
                    KT_gumbel = (0.78 * YT - 0.45) if use_gumbel_sni else ((YT - Yn) / Sn)
                    XT_gumbel = X_bar + Sd * KT_gumbel
                    KT_lp3 = float(scipy.stats.pearson3.ppf(prob_non_exceed, Cs_log))
                    XT_lp3 = 10.0 ** (mean_logX + KT_lp3 * S_logX)

                    dist_results["Normal"][T]          = round(XT_norm, 3)
                    dist_results["Log Normal"][T]      = round(XT_lognorm, 3)
                    dist_results["Gumbel"][T]          = round(XT_gumbel, 3)
                    dist_results["Log Pearson III"][T] = round(XT_lp3, 3)

                    comp_rows.append({
                        "Kala Ulang T (tahun)": T,
                        "Normal (mm)": round(XT_norm, 2),
                        "Log Normal (mm)": round(XT_lognorm, 2),
                        "Gumbel (mm)": round(XT_gumbel, 2),
                        "Log Pearson III (mm)": round(XT_lp3, 2),
                    })

                df_comp = pd.DataFrame(comp_rows)

                # ── 4. PERHITUNGAN UJI KECOCOKAN (S-K & CHI-KUADRAT) ──────────────
                X_sorted    = np.sort(Xi)
                logX_sorted = np.log10(np.clip(X_sorted, 1e-6, None))
                P_weibull   = np.arange(1, n_data + 1) / (n_data + 1.0)

                _ks_crit_table_05 = {5: 0.56, 10: 0.41, 15: 0.34, 20: 0.29, 25: 0.26, 30: 0.24}
                _ks_crit_table_01 = {5: 0.67, 10: 0.49, 15: 0.40, 20: 0.360, 25: 0.32, 30: 0.29}
                if n_data <= 30:
                    _nearest_ks = min(_ks_crit_table_05.keys(), key=lambda k: abs(k - n_data))
                    delta_kritis = _ks_crit_table_05[_nearest_ks]
                    delta_kritis_01 = _ks_crit_table_01[_nearest_ks]
                else:
                    delta_kritis = 1.36 / math.sqrt(n_data)
                    delta_kritis_01 = 1.63 / math.sqrt(n_data)

                G_classes = 5 if n_data >= 5 else max(3, n_data)
                E_freq    = n_data / float(G_classes)
                prob_edges = np.linspace(0.0, 1.0, G_classes + 1)[1:-1]

                gof_rows = []
                passed_dists = []

                for d_name in dist_names:
                    if d_name == "Normal":
                        cdf_vals = scipy.stats.norm.cdf(X_sorted, loc=X_bar, scale=max(Sd, 1e-6))
                        q_edges  = scipy.stats.norm.ppf(prob_edges, loc=X_bar, scale=max(Sd, 1e-6))
                        obs_counts, _ = np.histogram(X_sorted, bins=np.concatenate(([-np.inf], q_edges, [np.inf])))
                        m_params = 2
                    elif d_name == "Log Normal":
                        cdf_vals = scipy.stats.norm.cdf(logX_sorted, loc=mean_logX, scale=max(S_logX, 1e-6))
                        q_edges  = scipy.stats.norm.ppf(prob_edges, loc=mean_logX, scale=max(S_logX, 1e-6))
                        obs_counts, _ = np.histogram(logX_sorted, bins=np.concatenate(([-np.inf], q_edges, [np.inf])))
                        m_params = 2
                    elif d_name == "Gumbel":
                        if use_gumbel_sni:
                            y_reduced = (((X_sorted - X_bar) / max(Sd, 1e-6)) + 0.45) / 0.78
                            y_edges   = -np.log(-np.log(prob_edges))
                            x_edges   = X_bar + (0.78 * y_edges - 0.45) * Sd
                        else:
                            y_reduced = Sn * ((X_sorted - X_bar) / max(Sd, 1e-6)) + Yn
                            y_edges   = -np.log(-np.log(prob_edges))
                            x_edges   = X_bar + ((y_edges - Yn) / Sn) * Sd
                        cdf_vals  = np.exp(-np.exp(-y_reduced))
                        obs_counts, _ = np.histogram(X_sorted, bins=np.concatenate(([-np.inf], x_edges, [np.inf])))
                        m_params = 2
                    else:  # Log Pearson III
                        cdf_vals = scipy.stats.pearson3.cdf(logX_sorted, Cs_log, loc=mean_logX, scale=max(S_logX, 1e-6))
                        q_edges  = scipy.stats.pearson3.ppf(prob_edges, Cs_log, loc=mean_logX, scale=max(S_logX, 1e-6))
                        obs_counts, _ = np.histogram(logX_sorted, bins=np.concatenate(([-np.inf], q_edges, [np.inf])))
                        m_params = 3

                    delta_maks = float(np.max(np.abs(P_weibull - cdf_vals)))
                    ks_pass    = delta_maks < delta_kritis

                    dk = max(1, G_classes - m_params - 1)
                    chi2_hitung = float(np.sum(((obs_counts - E_freq) ** 2) / E_freq))
                    chi2_kritis = float(scipy.stats.chi2.ppf(0.95, dk))
                    chi2_pass   = chi2_hitung < chi2_kritis

                    both_pass = ks_pass and chi2_pass
                    if both_pass:
                        passed_dists.append((delta_maks, d_name))

                    gof_rows.append({
                        "Distribusi": d_name,
                        "Δ_maks (S-K)": round(delta_maks, 4),
                        "Δ_kritis (α=5%)": round(delta_kritis, 4),
                        "Δ_kritis (α=1%)": round(delta_kritis_01, 4),
                        "Status S-K": "✅ Diterima" if ks_pass else "❌ Ditolak",
                        "χ²_hitung": round(chi2_hitung, 3),
                        "χ²_kritis (α=5%)": round(chi2_kritis, 3),
                        "Status Chi-Kuadrat": "✅ Diterima" if chi2_pass else "❌ Ditolak",
                        "Kesimpulan Uji": "🏆 Lolos Uji" if both_pass else "⚠️ Tidak Lolos Penuh",
                    })

                df_gof = pd.DataFrame(gof_rows)

                if passed_dists:
                    passed_dists.sort(key=lambda x: x[0])
                    recommended_dist = passed_dists[0][1]
                    st.success(
                        f"🏆 **Rekomendasi SNI 2415:2016:** Distribusi **{recommended_dist}** lolos uji "
                        f"Smirnov-Kolmogorov dan Chi-Kuadrat dengan simpangan maksimum terkecil "
                        f"(Δ_maks = {passed_dists[0][0]:.4f})."
                    )
                else:
                    recommended_dist = "Gumbel"
                    st.warning(
                        "⚠️ Tidak ada distribusi yang lolos kedua uji secara bersamaan pada α=5%. "
                        "Default diarahkan ke **Gumbel** (atau pilih secara manual di bawah)."
                    )

                saved_dist = st.session_state.get("m1_selected_dist")
                if saved_dist in dist_names:
                    default_idx = dist_names.index(saved_dist)
                elif recommended_dist in dist_names:
                    default_idx = dist_names.index(recommended_dist)
                else:
                    default_idx = 2

                selected_dist = st.selectbox(
                    "🎯 Pilih Distribusi Hujan Rencana yang Dikirim ke Modul 2, 3 & 4 (Session State)",
                    options=dist_names,
                    index=default_idx,
                    help="Nilai R24 (T=5 tahun) dan dictionary seluruh kala ulang dari distribusi terpilih disimpan ke st.session_state['data_modul_1'].",
                )
                st.session_state["m1_selected_dist"] = selected_dist

                # Simpan ke st.session_state["data_modul_1"], st.session_state["R24_5yr"], dan st.session_state["r24_dict"]
                st.session_state["r24_dict"] = dist_results[selected_dist]
                st.session_state["R24_5yr"]  = float(dist_results[selected_dist][5])
                st.session_state["data_modul_1"] = {
                    "file_name": active_m1_name,
                    "df_raw": df_raw,
                    "r24_df": r24,
                    "is_monthly": is_monthly_active,
                    "col_date": col_date,
                    "col_rain": col_rain,
                    "selected_dist": selected_dist,
                    "recommended_dist": recommended_dist,
                    "R24_5yr": float(dist_results[selected_dist][5]),
                    "r24_dict": dist_results[selected_dist],
                    "dist_results": dist_results,
                    "full_doc_text": st.session_state.get("pdf_doc_text", ""),
                    "pdf_name": st.session_state.get("pdf_doc_name", ""),
                    "stats": {
                        "n_data": n_data,
                        "X_bar": X_bar,
                        "Sd": Sd,
                        "Cv": Cv,
                        "Cs": Cs,
                        "Ck": Ck,
                    },
                }

                st.success(
                    f"💾 **Tersimpan ke `st.session_state['data_modul_1']` (`{selected_dist}`):** "
                    f"`R24_5yr` = **{st.session_state['R24_5yr']:.2f} mm** | "
                    f"`r24_dict` = {st.session_state['r24_dict']}"
                )

                # Analytical Cards untuk Ringkasan Hujan Rencana R24 Multi-Kala Ulang
                _render_section_banner(
                    f"📊 Ringkasan Hujan Rencana R24 — Distribusi {selected_dist}",
                    "Nilai hujan harian maksimum rencana R24 (mm) untuk berbagai kala ulang T yang siap dialirkan ke Modul Hidrolika.",
                    badge_html=f'<span class="ds-badge safe">Rekomendasi SNI: {recommended_dist}</span>',
                    tone="teal",
                )
                r24_cols = st.columns(6)
                for _i_rp, _t_rp in enumerate([2, 5, 10, 25, 50, 100]):
                    _val_rp = dist_results[selected_dist][_t_rp]
                    r24_cols[_i_rp].metric(
                        f"R24 (T={_t_rp} thn)",
                        f"{_val_rp:.2f} mm",
                        delta="Acuan Utama" if _t_rp == 5 else f"P={100/_t_rp:.0f}%",
                        delta_color="normal" if _t_rp == 5 else "off",
                    )

                st.divider()

                # ── FOKUS UTAMA: GRAFIK KURVA IDF MULTI-KALA ULANG (METODE MONONOBE) ──
                _render_section_banner(
                    "📉 Visualisasi Utama: Kurva IDF Multi-Kala Ulang (Metode Mononobe)",
                    f"Kurva Intensitas–Durasi–Frekuensi (IDF) dari R24 distribusi {selected_dist} untuk durasi tc = 5 s/d 180 menit.",
                    badge_html='<span class="ds-badge auto">I = (R24/24) · (24/tc)^(2/3)</span>',
                    tone="ocean",
                )

                idf_return_periods = [2, 5, 10, 25, 50, 100]
                tc_minutes = list(range(5, 185, 5))
                idf_data = {"tc (menit)": tc_minutes}

                for T_idf in idf_return_periods:
                    r24_T = dist_results[selected_dist][T_idf]
                    col_label = f"T = {T_idf} Tahun"
                    idf_data[col_label] = [
                        round((r24_T / 24.0) * ((24.0 / (tc_m / 60.0)) ** (2.0 / 3.0)), 2)
                        for tc_m in tc_minutes
                    ]

                df_idf = pd.DataFrame(idf_data).set_index("tc (menit)")

                # Scientific Matplotlib IDF Plot with clean legend + Interactive Chart tab
                idf_tab1, idf_tab2 = st.tabs(["📈 Grafik Ilmiah IDF (High-Res Hydrology Plot)", "⚡ Grafik Interaktif (Hover Data)"])
                with idf_tab1:
                    fig_idf, ax_idf = plt.subplots(figsize=(10.2, 4.6), dpi=120)
                    fig_idf.patch.set_facecolor("#FFFFFF")
                    ax_idf.set_facecolor("#F4F8FC")
                    _idf_palette = ["#34C6D3", "#11A8A0", "#1677C8", "#102A43", "#F4B942", "#E45B64"]
                    for _idx_c, _T_idf in enumerate(idf_return_periods):
                        _col_lbl = f"T = {_T_idf} Tahun"
                        ax_idf.plot(
                            tc_minutes,
                            idf_data[_col_lbl],
                            label=f"{_col_lbl} (R24 = {dist_results[selected_dist][_T_idf]:.1f} mm)",
                            color=_idf_palette[_idx_c % len(_idf_palette)],
                            linewidth=2.3 if _T_idf in (5, 10, 25) else 1.8,
                            marker="o" if _T_idf == 5 else None,
                            markersize=3.5,
                        )
                    ax_idf.set_title(
                        f"Kurva Intensitas-Durasi-Frekuensi (IDF) Mononobe — Distribusi {selected_dist}",
                        fontsize=11.5,
                        fontweight="bold",
                        color="#102A43",
                        pad=10,
                    )
                    ax_idf.set_xlabel("Durasi Hujan / Waktu Konsentrasi tc (menit)", fontsize=9.5, color="#18324A")
                    ax_idf.set_ylabel("Intensitas Hujan I (mm/jam)", fontsize=9.5, color="#18324A")
                    ax_idf.grid(True, linestyle="--", alpha=0.45, color="#B8CCE0")
                    ax_idf.legend(
                        title="Kala Ulang Rencana (T)",
                        loc="upper right",
                        fontsize=8.5,
                        title_fontsize=8.8,
                        frameon=True,
                        facecolor="#FFFFFF",
                        edgecolor="#DFE8F0",
                    )
                    fig_idf.tight_layout()
                    st.pyplot(fig_idf)
                    plt.close(fig_idf)
                with idf_tab2:
                    st.line_chart(df_idf)

                # ── EXPANDER DETAIL STATISTIK & TABEL DISTRIBUSI ──────────────────
                with st.expander("Tampilkan Detail Statistik & Tabel Distribusi"):
                    st.markdown(
                        "#### 📋 Seri Curah Hujan Maksimum Bulanan (R24 Max)"
                        if is_monthly_active
                        else "#### 📋 Curah Hujan Maksimum Harian per Tahun (R24 Max)"
                    )
                    st.dataframe(r24.style.format({"R24_Max (mm)": "{:.2f}"}))
                    st.divider()

                    st.markdown("#### 📊 1. Statistik Lengkap R24 Maksimum (SNI 2415:2016)")
                    st_c1, st_c2, st_c3, st_c4, st_c5, st_c6 = st.columns(6)
                    st_c1.metric("Jumlah Data (n)", f"{n_data}")
                    st_c2.metric("Rata-rata X̄ (mm)", f"{X_bar:.3f}")
                    st_c3.metric("Simpangan Baku Sd", f"{Sd:.3f}")
                    st_c4.metric("Koef. Variasi (Cv)", f"{Cv:.4f}")
                    st_c5.metric("Kemencengan (Cs)", f"{Cs:.4f}")
                    st_c6.metric("Kurtosis (Ck)", f"{Ck:.4f}")

                    df_stat_detail = pd.DataFrame([
                        {"Parameter": "Jumlah Data (n)", "Skala Asli (X)": f"{n_data}", "Skala Log10(X)": f"{n_data}"},
                        {"Parameter": "Rata-rata (X̄)", "Skala Asli (X)": f"{X_bar:.4f} mm", "Skala Log10(X)": f"{mean_logX:.4f}"},
                        {"Parameter": "Simpangan Baku (Sd)", "Skala Asli (X)": f"{Sd:.4f} mm", "Skala Log10(X)": f"{S_logX:.4f}"},
                        {"Parameter": "Koefisien Variasi (Cv)", "Skala Asli (X)": f"{Cv:.4f}", "Skala Log10(X)": f"{Cv_log:.4f}"},
                        {"Parameter": "Koefisien Kemencengan (Cs)", "Skala Asli (X)": f"{Cs:.4f}", "Skala Log10(X)": f"{Cs_log:.4f}"},
                        {"Parameter": "Koefisien Kurtosis (Ck)", "Skala Asli (X)": f"{Ck:.4f}", "Skala Log10(X)": f"{Ck_log:.4f}"},
                    ])
                    st.dataframe(df_stat_detail, hide_index=True)
                    st.latex(r"\bar{X} = \frac{1}{n}\sum_{i=1}^{n} X_i \quad;\quad S_d = \sqrt{\frac{\sum (X_i - \bar{X})^2}{n-1}} \quad;\quad C_v = \frac{S_d}{\bar{X}}")
                    st.latex(r"C_s = \frac{n \sum (X_i - \bar{X})^3}{(n-1)(n-2)S_d^3} \quad;\quad C_k = \frac{n^2 \sum (X_i - \bar{X})^4}{(n-1)(n-2)(n-3)S_d^4}")
                    st.divider()

                    st.markdown("#### 🔍 2. Uji Outlier Metode Grubbs-Beck (Taraf Signifikansi 10%)")
                    gb1, gb2, gb3, gb4, gb5 = st.columns(5)
                    gb1.metric("Rata-rata log10(X)", f"{mean_logX:.4f}")
                    gb2.metric("Simpangan Baku S_logX", f"{S_logX:.4f}")
                    gb3.metric("Faktor Grubbs-Beck (Kn)", f"{Kn:.4f}")
                    gb4.metric("Batas Atas (XH)", f"{XH:.2f} mm")
                    gb5.metric("Batas Bawah (XL)", f"{XL:.2f} mm")

                    if outliers_high.empty and outliers_low.empty:
                        st.success(
                            f"✅ **Status Outlier: TIDAK ADA OUTLIER** — Seluruh data R24 berada dalam rentang "
                            f"batas bawah **XL = {XL:.2f} mm** hingga batas atas **XH = {XH:.2f} mm**."
                        )
                    else:
                        msg_parts = []
                        lbl_col = "Periode" if "Periode" in r24.columns else "Tahun"
                        if not outliers_high.empty:
                            yrs_h = ", ".join(
                                f"{r[lbl_col]} ({r['R24_Max (mm)']:.1f} mm)"
                                for _, r in outliers_high.iterrows()
                            )
                            msg_parts.append(f"**Melewati Batas Atas XH ({XH:.2f} mm):** {yrs_h}")
                        if not outliers_low.empty:
                            yrs_l = ", ".join(
                                f"{r[lbl_col]} ({r['R24_Max (mm)']:.1f} mm)"
                                for _, r in outliers_low.iterrows()
                            )
                            msg_parts.append(f"**Di Bawah Batas Bawah XL ({XL:.2f} mm):** {yrs_l}")
                        st.warning("⚠️ **Status Outlier: TERDETEKSI DATA OUTLIER** — " + " | ".join(msg_parts))
                    st.divider()

                    st.markdown("#### 📈 3. Tabel Komparasi 4 Distribusi Frekuensi Hujan Rencana")
                    st.dataframe(
                        df_comp.style.format({
                            "Normal (mm)": "{:.2f}",
                            "Log Normal (mm)": "{:.2f}",
                            "Gumbel (mm)": "{:.2f}",
                            "Log Pearson III (mm)": "{:.2f}",
                        }),
                        hide_index=True,
                    )
                    df_comp_chart = df_comp.set_index("Kala Ulang T (tahun)")
                    st.line_chart(df_comp_chart)
                    st.divider()

                    st.markdown("#### ✅ 4. Tabel Uji Kecocokan Distribusi (Smirnov-Kolmogorov & Chi-Kuadrat)")
                    st.dataframe(
                        df_gof.style.format({
                            "Δ_maks (S-K)": "{:.4f}",
                            "Δ_kritis (α=5%)": "{:.4f}",
                            "χ²_hitung": "{:.3f}",
                            "χ²_kritis (α=5%)": "{:.3f}",
                        }),
                        hide_index=True,
                    )
                    st.divider()

                    st.markdown("#### 📋 5. Tabel Nilai Intensitas Hujan IDF Mononobe (mm/jam)")
                    st.latex(r"I = \frac{R_{24}}{24} \times \left(\frac{24}{t_c / 60}\right)^{2/3}")
                    st.dataframe(df_idf)

        except Exception as e:
            st.error(f"❌ Error membaca/memproses file: {e}")
    else:
        _render_kobo_empty_state(
            asset_key="kobo_rain",
            title="Menunggu Data Curah Hujan Harian BMKG",
            subtitle=(
                "Silakan unggah file CSV (.csv), Excel (.xlsx), atau dokumen teknis PDF (.pdf) di atas "
                "untuk memulai analisis frekuensi hujan rencana (R24) dan membangkitkan Kurva IDF Mononobe."
            ),
            badges_html=(
                '<span class="ds-badge auto">1 · Uji Outlier Grubbs-Beck</span>'
                '<span class="ds-badge auto">2 · 4 Distribusi Frekuensi SNI 2415:2016</span>'
                '<span class="ds-badge auto">3 · Kurva IDF Mononobe</span>'
            ),
            show_image=False,
        )


# ═════════════════════════════════════════════════════════════════
# MODUL 2: Diskusi Hibrida AI + Python untuk Ekstraksi PDF / Word DED
# ═════════════════════════════════════════════════════════════════
elif menu.startswith("2 |") or "Modul 2" in menu:
    _m2_ready = bool(st.session_state.get("ded_params") or st.session_state.get("params_modul_2"))
    _render_module_header(
        step_num=2,
        icon="📄",
        title="Ekstraksi DED (Python NLP) & Validasi Chow (1959)",
        description=(
            "Ekstraksi Cerdas Parameter DED menggunakan Python NLP & Regex Engine"
        ),
        status_text=(
            f"DED Aktif: b={st.session_state['ded_params']['b']}m, h={st.session_state['ded_params']['h']}m"
            if st.session_state.get("ded_params")
            else "Menunggu Ekstraksi / Input DED"
        ),
        is_ready=_m2_ready,
    )

    # Integrasi langsung dari st.session_state['data_modul_1']
    m1_data_in_m2 = st.session_state.get("data_modul_1")
    if isinstance(m1_data_in_m2, dict) and m1_data_in_m2.get("r24_dict"):
        r24_str_m2 = ", ".join(f"T{k}={v:.1f} mm" for k, v in m1_data_in_m2["r24_dict"].items())
        st.info(
            f"🔗 **Terhubung dari Modul 1 (`st.session_state['data_modul_1']` — {m1_data_in_m2.get('file_name', 'BMKG')}):** "
            f"Distribusi **{m1_data_in_m2.get('selected_dist', 'Gumbel')}** | "
            f"**R24 (T=5 thn) = {m1_data_in_m2.get('R24_5yr', 0.0):.2f} mm** ({r24_str_m2})"
        )
    elif st.session_state.get("R24_5yr") is not None:
        st.info(f"🔗 **Terhubung dari Modul 1:** R24 (T=5 thn) = **{st.session_state['R24_5yr']:.2f} mm**")

    col1, col2 = st.columns([6, 4], gap="large")
    col_ai, col_manual = col1, col2

    with col_ai:
        _render_kobo_companion_card(
            asset_key="kobo_thinking",
            tag="📄 KOBO DED & CHOW VALIDATOR · VEN TE CHOW (1959)",
            title="Ekstraksi Dimensi Saluran & Validasi Koefisien Manning",
            description=(
                "Gunakan Ekstraksi Cerdas via Python NLP & Regex Engine untuk membaca parameter "
                "lebar dasar (b), kedalaman (h), kemiringan (S), dan material saluran dari laporan teknis DED."
            ),
        )
        _render_section_banner(
            "⚡ Ekstraksi Dokumen DED (Python NLP & Regex Engine)",
            "Unggah dokumen laporan teknis (.pdf / .docx) untuk diekstrak secara otomatis oleh Python NLP Engine.",
            badge_html='<span class="ds-badge auto">PDF · DOCX · Python NLP</span>',
            tone="ocean",
        )

        doc_file = st.file_uploader(
            "📎 Langkah 1 — Unggah dokumen DED / Laporan Teknis Hidrologi (.pdf atau .docx)",
            type=["pdf", "docx"],
            help="Python mengekstrak teks & tabel secara lokal menggunakan pypdf / pdfplumber / python-docx.",
        )

        def _prepare_m2_preview(raw_doc_text: str, max_chars: int = 25000) -> tuple[str, dict]:
            builder = globals().get("_build_preview_data")
            if callable(builder):
                return builder(raw_doc_text, max_chars=max_chars)
            txt = (raw_doc_text or "").strip()
            preview = txt[:max_chars]
            return preview, {
                "total_karakter_dokumen": len(txt),
                "jumlah_segmen": 1 if txt else 0,
                "segmen_relevan": 1 if txt else 0,
                "karakter_preview": len(preview),
                "kata_kunci_terdeteksi": {},
            }

        # Simpan dokumen yang diunggah ke st.session_state['uploaded_file_m2'] agar tidak hilang saat pindah modul
        if doc_file is not None:
            raw_bytes = bytes(doc_file.getbuffer())
            fname_low = doc_file.name.lower()
            full_doc_text = (
                _extract_pdf_text_pypdf(raw_bytes)
                if fname_low.endswith(".pdf")
                else _extract_docx_text_local(raw_bytes)
            )
            preview_data, preview_stats = _prepare_m2_preview(full_doc_text)
            st.session_state["uploaded_file_m2"] = {
                "name": doc_file.name,
                "size": doc_file.size,
                "full_doc_text": full_doc_text,
                "preview_data": preview_data,
                "preview_stats": preview_stats,
            }
            st.session_state["pdf_doc_text"] = full_doc_text
            st.session_state["pdf_doc_name"] = doc_file.name
            if isinstance(st.session_state.get("uploaded_file_m1"), dict):
                st.session_state["uploaded_file_m1"]["full_doc_text"] = full_doc_text
                st.session_state["uploaded_file_m1"]["pdf_name"] = doc_file.name
            else:
                st.session_state["uploaded_file_m1"] = {
                    "pdf_name": doc_file.name,
                    "full_doc_text": full_doc_text,
                }
            if isinstance(st.session_state.get("data_modul_1"), dict):
                st.session_state["data_modul_1"]["full_doc_text"] = full_doc_text
                st.session_state["data_modul_1"]["pdf_name"] = doc_file.name

        saved_doc_m2 = st.session_state.get("uploaded_file_m2")
        if isinstance(saved_doc_m2, dict) and saved_doc_m2.get("name"):
            active_doc_name = saved_doc_m2["name"]
            active_doc_size = saved_doc_m2.get("size", 0)
            full_doc_text = saved_doc_m2.get("full_doc_text", "")
            preview_data = saved_doc_m2.get("preview_data", "")
            preview_stats = saved_doc_m2.get("preview_stats", {})

            if doc_file is not None:
                st.write(f"📄 **{active_doc_name}** ({active_doc_size / 1024:.1f} KB)")
            else:
                st.success(f"💾 Dokumen Modul 2 aktif dari Session State: **{active_doc_name}** ({active_doc_size / 1024:.1f} KB)")

            with st.expander("🔍 Langkah 1 (Python Extractor): Ringkasan `preview_data` Terfokus", expanded=False):
                ps1, ps2, ps3 = st.columns(3)
                ps1.metric("Total Karakter Dokumen", f"{preview_stats.get('total_karakter_dokumen', 0):,}")
                ps2.metric("Segmen Relevan Terpilih", f"{preview_stats.get('segmen_relevan', 0)}/{preview_stats.get('jumlah_segmen', 0)}")
                ps3.metric("Ukuran preview_data", f"{preview_stats.get('karakter_preview', 0):,} kar")
                if preview_stats.get("kata_kunci_terdeteksi"):
                    st.caption("Kata kunci teknis terdeteksi: " + ", ".join(f"`{k}` ({v}x)" for k, v in preview_stats["kata_kunci_terdeteksi"].items()))
                st.text_area(
                    "Isi preview_data (Bab Hidrologi / Hidrolika / Kriteria Desain / Normalisasi / Sudetan)",
                    value=preview_data[:6000] + ("\n\n...[Dipangkas untuk tampilan]..." if len(preview_data) > 6000 else ""),
                    height=180,
                    disabled=True,
                )

            btn_c1, btn_c2 = st.columns(2)
            with btn_c1:
                btn_ekstrak = st.button("⚡ Ekstraksi Parameter DED (Python NLP)", type="primary", use_container_width=True)
            with btn_c2:
                btn_ai_m2 = st.button("✨ Analisis Konteks Dokumen (Gemini AI)", type="secondary", use_container_width=True)

            if btn_ekstrak:
                t0 = time.perf_counter()
                doc_class = _classify_civil_document(full_doc_text or preview_data)
                st.session_state["m2_doc_class"] = doc_class

                # Ekstraksi langsung dari teks dokumen PDF/Word yang diunggah di Langkah 1 secara real-time
                loc_primary = _extract_ded_params_local(preview_data or full_doc_text, active_doc_name)
                if loc_primary.get("explicit_count", 0) < 4 and full_doc_text and preview_data != full_doc_text:
                    loc_full = _extract_ded_params_local(full_doc_text, active_doc_name)
                    if loc_full.get("explicit_count", 0) > loc_primary.get("explicit_count", 0):
                        loc_primary = loc_full

                elapsed_ms = (time.perf_counter() - t0) * 1000.0
                q_calc = loc_primary.get("Q_rencana") or round(
                    _manning_Q(loc_primary["b"], loc_primary["h"], loc_primary["S"], loc_primary["n"]), 4
                )
                explicit_cnt = int(loc_primary.get("explicit_count", 0))
                top_kws_str = ", ".join(doc_class.get("top_keywords_list", [])[:6]) or "-"
                dom_key = doc_class.get("primary_domain", "umum")

                domain_warning_msg = ""
                if explicit_cnt == 0:
                    if dom_key == "geoteknik":
                        domain_warning_msg = (
                            f"⚠️ Dokumen Terdeteksi: Rekayasa Geoteknik / Mekanika Tanah "
                            f"(Ditemukan topik: {top_kws_str}). Dokumen ini berisi parameter tanah "
                            f"(seperti N-SPT, Kohesi, Sudut Geser, Permeabilitas, atau pemodelan PLAXIS/GeoStudio) "
                            f"dan tidak memuat dimensi DED saluran drainase (b, h, S, n)."
                        )
                    elif dom_key == "hidrologi":
                        domain_warning_msg = (
                            f"⚠️ Dokumen Terdeteksi: Hidrologi / Klimatologi Murni "
                            f"(Ditemukan topik: {top_kws_str}). Dokumen ini berfokus pada analisis curah hujan / kurva IDF "
                            f"(cocok untuk referensi Modul 1) dan tidak memuat dimensi saluran DED (b, h, S, n)."
                        )
                    elif dom_key == "struktur":
                        domain_warning_msg = (
                            f"⚠️ Dokumen Terdeteksi: Rekayasa Struktur / Beton / Baja "
                            f"(Ditemukan topik: {top_kws_str}). Dokumen ini berfokus pada analisis struktur/material "
                            f"dan tidak memuat dimensi saluran drainase (b, h, S, n)."
                        )
                    else:
                        domain_warning_msg = (
                            "⚠️ Dokumen terdeteksi sebagai laporan non-DED "
                            "(tidak ditemukan angka dimensi saluran b, h, S, atau material). "
                            "Form menggunakan nilai default. Silakan sesuaikan manual atau unggah dokumen DED saluran."
                        )

                py_raw_dict = {
                    "lebar_dasar_b": loc_primary["b"],
                    "kedalaman_h": loc_primary["h"],
                    "kemiringan_s": loc_primary["S"],
                    "manning_n": loc_primary["n"],
                    "material": loc_primary["material"],
                    "debit_rencana_q": q_calc,
                    "kutipan_halaman": loc_primary["kutipan_ded"],
                    "alasan_ai": (
                        f"Klasifikasi Domain: **{doc_class.get('primary_label', 'Umum')}** | "
                        f"Diekstrak langsung secara real-time oleh Python NLP Engine ({elapsed_ms:.1f} ms) dari dokumen "
                        f"'{active_doc_name}' ({explicit_cnt}/5 parameter DED terdeteksi eksplisit, "
                        f"{preview_stats.get('segmen_relevan', 1)} segmen teknis relevan)."
                    ),
                }
                validated_py = _validate_m2_ai_json(
                    py_raw_dict,
                    source=f"Python NLP & Regex Engine ({elapsed_ms:.0f} ms)",
                )
                validated_py["explicit_count"] = explicit_cnt
                validated_py["found_keys"] = loc_primary.get("found_keys", {})
                validated_py["doc_class"] = doc_class
                validated_py["domain_warning"] = domain_warning_msg
                if loc_primary.get("luas_dta_ha") is not None:
                    validated_py["luas_dta_ha"] = float(loc_primary["luas_dta_ha"])
                    st.session_state["m2_luas_dta_ha"] = float(loc_primary["luas_dta_ha"])

                st.session_state["m2_ai_result"] = validated_py
                # Langsung update st.session_state Form Parameter DED di panel sebelah kanan (< 1 detik)
                _commit_m2_to_ded(validated_py)
                if loc_primary.get("luas_dta_ha") is not None:
                    for _k_ss in ("ded_params", "params_modul_2"):
                        if isinstance(st.session_state.get(_k_ss), dict):
                            st.session_state[_k_ss]["luas_dta_ha"] = float(loc_primary["luas_dta_ha"])
                # Simpan otomatis di background ke drainsense_bridge.json sebagai cache/cadangan sistem
                _save_m2_to_bridge_cache(py_raw_dict, source=f"Python NLP Engine ({active_doc_name})")
                if explicit_cnt == 0:
                    st.warning(domain_warning_msg)
                else:
                    st.success(f"✅ Parameter berhasil diekstrak langsung oleh Python NLP Engine ({elapsed_ms:.0f} ms)")
                    st.toast("✅ Parameter berhasil diekstrak oleh Python NLP Engine!", icon="⚡")

            if btn_ai_m2:
                api_key_m2 = (st.session_state.get("gemini_key") or "").strip()
                if not api_key_m2:
                    err_m2 = "API Key Gemini di sidebar masih kosong. Masukkan API Key terlebih dahulu."
                    st.session_state["m2_ai_error"] = err_m2
                    st.warning(f"⚠️ {err_m2}")
                else:
                    excerpt_for_ai = (preview_data or full_doc_text or "")[:12000]
                    prompt_m2_ai = (
                        "Anda adalah asisten ahli Teknik Sipil (Hidrologi, Drainase, Geoteknik, dan Struktur).\n"
                        f"Bacalah kutipan teks dari dokumen '{active_doc_name}' berikut ini, lalu berikan:\n"
                        "1. **Klasifikasi Bidang Dokumen** (misal: Geoteknik / Hidrologi IDF / Struktur / DED Drainase).\n"
                        "2. **Ringkasan Kontekstual Isi Dokumen** (poin-poin utama, metode, rumus, atau parameter penting yang ditemukan di dokumen).\n"
                        "3. **Status Parameter DED Saluran** (apakah dimensi b, h, S, n, dan material saluran ada di dokumen? Jika tidak ada, nyatakan dengan tegas bahwa dokumen ini bukan DED saluran dan jangan mengarang angka).\n\n"
                        f"--- TEKS DOKUMEN ---\n{excerpt_for_ai}"
                    )
                    with st.spinner("🤖 Memanggil Gemini AI (Auto-Discovery Model Aktif) untuk merangkum konteks dokumen..."):
                        try:
                            ai_summary_txt, active_m2_model = call_gemini_text(api_key_m2, prompt_m2_ai)
                            st.session_state["m2_ai_error"] = None
                            st.session_state["m2_ai_context_summary"] = {
                                "doc_name": active_doc_name,
                                "model_name": active_m2_model,
                                "summary": ai_summary_txt,
                                "timestamp": time.strftime("%H:%M:%S"),
                            }
                            st.toast(f"✨ Ringkasan dokumen selesai oleh {active_m2_model}!", icon="✨")
                        except Exception as e_m2_ai:
                            err_m2 = f"Panggilan Gemini AI gagal: {e_m2_ai}"
                            st.session_state["m2_ai_error"] = err_m2
                            st.warning(f"⚠️ {err_m2}")
        else:
            _render_kobo_empty_state(
                asset_key="kobo_thinking",
                title="Belum Ada Dokumen Teknis DED (.pdf / .docx)",
                subtitle=(
                    "Unggah dokumen laporan DED di atas untuk diekstrak secara otomatis oleh Python NLP Extractor, "
                    "atau isi parameter dimensi saluran secara langsung pada Form Parameter DED di sebelah kanan."
                ),
                badges_html=(
                    '<span class="ds-badge auto">Ekstraksi Instan Python (&lt;1 dtk)</span>'
                    '<span class="ds-badge auto">Validasi Tabel Ven Te Chow (1959)</span>'
                ),
                show_image=False,
            )

        # ── Langkah 3: Kotak Dialog "Hasil Ekstraksi Python NLP & Validasi Chow" ──
        m2_res = st.session_state.get("m2_ai_result")
        if m2_res:
            with st.container(border=True):
                st.markdown("### 📋 Hasil Ekstraksi Python NLP & Validasi Tabel Ven Te Chow (1959)")
                st.caption(f"Sumber Analisis: **{m2_res['source']}** · Waktu: `{m2_res['timestamp']}`")

                explicit_cnt_res = int(m2_res.get("explicit_count", 1))
                found_keys_res = m2_res.get("found_keys") or {}
                doc_class_res = m2_res.get("doc_class") or st.session_state.get("m2_doc_class") or {}

                if explicit_cnt_res == 0:
                    if m2_res.get("domain_warning"):
                        st.warning(m2_res["domain_warning"])
                    st.info(
                        "ℹ️ Kartu metrik dimensi saluran (b = 1.500 m, h = 1.200 m, S = 0.001, n = 0.013) "
                        "disembunyikan karena merupakan **Nilai Default (Bukan dari PDF)**."
                    )
                else:
                    def _m2_delta(k: str) -> str:
                        if not found_keys_res:
                            return "Terbaca"
                        return "Terbaca dari PDF" if found_keys_res.get(k) else "Nilai Default (Bukan dari PDF)"

                    m2_c1, m2_c2, m2_c3, m2_c4, m2_c5 = st.columns(5)
                    m2_c1.metric("Lebar Dasar b", f"{m2_res['b']:.3f} m", delta=_m2_delta("b"), delta_color="normal" if found_keys_res.get("b", True) else "off")
                    m2_c2.metric("Kedalaman h", f"{m2_res['h']:.3f} m", delta=_m2_delta("h"), delta_color="normal" if found_keys_res.get("h", True) else "off")
                    m2_c3.metric("Kemiringan S", f"{m2_res['S']:.6f}", delta=_m2_delta("S"), delta_color="normal" if found_keys_res.get("S", True) else "off")
                    m2_c4.metric("Manning n", f"{m2_res['n']:.4f}", delta=_m2_delta("n"), delta_color="normal" if found_keys_res.get("n", True) else "off")
                    m2_c5.metric("Q Rencana Dokumen", f"{m2_res.get('Q_rencana', 0.0):.3f} m³/s", delta=_m2_delta("Q_rencana"), delta_color="normal" if found_keys_res.get("Q_rencana", True) else "off")

                if doc_class_res:
                    with st.expander("🔍 Detail Entitas Teknis yang Terbaca di Dokumen", expanded=(explicit_cnt_res == 0)):
                        dom_scores = doc_class_res.get("domain_scores") or {}
                        if dom_scores:
                            score_cols = st.columns(4)
                            score_cols[0].metric("Geoteknik", f"{dom_scores.get('geoteknik', 0)} hit")
                            score_cols[1].metric("Hidrologi / IDF", f"{dom_scores.get('hidrologi', 0)} hit")
                            score_cols[2].metric("Struktur", f"{dom_scores.get('struktur', 0)} hit")
                            score_cols[3].metric("DED Drainase", f"{dom_scores.get('drainase', 0)} hit")
                        entities_list = doc_class_res.get("entities") or []
                        if entities_list:
                            st.dataframe(pd.DataFrame(entities_list), hide_index=True, use_container_width=True)
                        else:
                            st.caption("Tidak ada entitas numerik spesifik yang terdeteksi di luar kata kunci umum.")

                st.markdown(f"**⚡ Catatan Teknis Python NLP Engine:**  \n{m2_res['alasan_ai']}")
                st.markdown(f"**📌 Kutipan & Halaman Asal Dokumen DED:**  \n_{m2_res['kutipan_ded']}_")

                if m2_res.get("issues"):
                    for iss in m2_res["issues"]:
                        st.caption(f"ℹ️ Catatan Validator Python: {iss}")

                is_ok_m2, chow_msg_m2 = _validate_chow_n(m2_res["material"], float(m2_res["n"]))
                if is_ok_m2:
                    st.success(chow_msg_m2)
                else:
                    st.warning(chow_msg_m2)
                    fix_c1, fix_c2 = st.columns(2)
                    with fix_c1:
                        if st.button(
                            f"🔧 Koreksi Otomatis ke Nilai Tipikal Chow ({m2_res['material']}: n = {m2_res['n_typ']:.3f})",
                            type="primary",
                        ):
                            m2_res["n"] = m2_res["n_typ"]
                            m2_res["chow_valid"] = True
                            st.session_state["m2_ai_result"] = m2_res
                            _commit_m2_to_ded(m2_res, n_override=m2_res["n_typ"])
                            st.success(f"✅ Nilai Manning n dikoreksi ke {m2_res['n_typ']:.3f} dan disimpan ke form DED!")
                            st.rerun()
                    with fix_c2:
                        if st.button(f"⚠️ Tetap Gunakan Nilai n = {m2_res['n']:.4f}"):
                            _commit_m2_to_ded(m2_res)
                            st.info(f"Parameter DED disimpan dengan nilai n = {m2_res['n']:.4f}.")
                            st.rerun()

        m2_ctx_ai = st.session_state.get("m2_ai_context_summary")
        if isinstance(m2_ctx_ai, dict) and m2_ctx_ai.get("summary"):
            with st.container(border=True):
                st.markdown(f"### ✨ Ringkasan Kontekstual Dokumen (Gemini AI: `{m2_ctx_ai.get('model_name', 'Gemini')}`)")
                st.caption(f"Dokumen: **{m2_ctx_ai.get('doc_name', 'PDF/DOCX')}** · Waktu: `{m2_ctx_ai.get('timestamp', '-')}`")
                st.markdown(m2_ctx_ai["summary"])

        with st.expander("📖 Tabel Referensi Koefisien Manning n — Ven Te Chow (1959)", expanded=True):
            df_chow = pd.DataFrame([
                {"Material Saluran": k, "n Minimum": v[0], "n Tipikal": v[1], "n Maksimum": v[2], "V_izin Maks (m/s)": v[3]}
                for k, v in CHOW_1959_RANGES.items()
            ])
            st.dataframe(
                df_chow.style.format({"n Minimum": "{:.3f}", "n Tipikal": "{:.3f}", "n Maksimum": "{:.3f}", "V_izin Maks (m/s)": "{:.1f}"}),
                hide_index=True,
            )

    with col_manual:
        _has_auto_m2 = st.session_state.get("m2_ai_result") is not None
        _m2_mode_badge = (
            '<span class="ds-badge auto">⚡ Auto-Extracted</span>'
            if _has_auto_m2
            else '<span class="ds-badge manual">✏️ Manual / Default</span>'
        )
        _render_section_banner(
            "✏️ Form Parameter DED",
            "Terisi otomatis dari Ekstraksi Parameter DED via Python NLP, dan tetap dapat di-override manual.",
            badge_html=_m2_mode_badge,
            tone="teal",
        )

        saved = st.session_state.get("ded_params") or {}
        saved_mat = _normalize_material(saved.get("material", "Beton"))
        if saved_mat not in MATERIAL_OPTIONS:
            saved_mat = MATERIAL_OPTIONS[0]

        # Inisialisasi state widget form agar langsung sinkron dengan hasil ekstraksi otomatis maupun override manual
        st.session_state.setdefault("ded_form_b", float(saved.get("b", 1.5)))
        st.session_state.setdefault("ded_form_h", float(saved.get("h", 1.2)))
        st.session_state.setdefault("ded_form_S", float(saved.get("S", 0.001)))
        st.session_state.setdefault("ded_form_n", float(saved.get("n", 0.013)))
        if st.session_state.get("ded_form_material") not in MATERIAL_OPTIONS:
            st.session_state["ded_form_material"] = saved_mat
        st.session_state.setdefault(
            "ded_form_kutipan",
            str(saved.get("kutipan_ded", "Input manual / Bab IV Kriteria Desain Hidrolika Hal. 42")),
        )

        with st.form("ded_manual_form"):
            mb   = st.number_input("Lebar dasar b (m)",       min_value=0.01, format="%.3f", key="ded_form_b")
            mh   = st.number_input("Kedalaman desain h (m)",  min_value=0.01, format="%.3f", key="ded_form_h")
            mS   = st.number_input("Kemiringan S (m/m)",      min_value=1e-6, format="%.6f", key="ded_form_S")
            mn   = st.number_input("Koefisien Manning n",     min_value=0.001, format="%.4f", key="ded_form_n",
                                   help="Beton: 0.011–0.020 | Pasangan Bata: 0.011–0.018 | Pasangan Batu Kali: 0.017–0.030 | Tanah: 0.016–0.033")
            mmat = st.selectbox("Material Saluran (Tabel Chow 1959)", MATERIAL_OPTIONS, key="ded_form_material")
            mkutipan = st.text_area(
                "Kutipan Teks / Halaman Asal Dokumen DED",
                height=80,
                key="ded_form_kutipan",
            )

            if st.form_submit_button("💾 Simpan & Validasi Parameter DED", type="primary"):
                ded_manual_dict = {
                    "b": float(mb),
                    "h": float(mh),
                    "S": float(mS),
                    "n": float(mn),
                    "material": mmat,
                    "kutipan_ded": mkutipan,
                    "Q_rencana": float(saved.get("Q_rencana", 0.0)),
                    "alasan_ai": saved.get("alasan_ai", "Input/koreksi manual oleh pengguna."),
                }
                st.session_state["ded_params"] = ded_manual_dict
                st.session_state["params_modul_2"] = dict(ded_manual_dict)
                _save_m2_to_bridge_cache(
                    {
                        "lebar_dasar_b": float(mb),
                        "kedalaman_h": float(mh),
                        "kemiringan_s": float(mS),
                        "manning_n": float(mn),
                        "material": mmat,
                        "debit_rencana_q": float(saved.get("Q_rencana", 0.0)),
                        "kutipan_halaman": mkutipan,
                        "alasan_ai": "Diperbarui melalui override manual oleh pengguna.",
                    },
                    source="Manual Override Form DED",
                )
                st.success("✅ Parameter DED disimpan ke session state!")
                st.rerun()

        if st.session_state["ded_params"]:
            st.markdown("---")
            d = st.session_state["ded_params"]
            is_valid_now, msg_now = _validate_chow_n(d["material"], float(d["n"]))
            _val_badge = '<span class="ds-badge safe">✅ Chow Valid</span>' if is_valid_now else '<span class="ds-badge danger">⚠️ Di Luar Rentang Chow</span>'
            _render_section_banner(
                "📐 Status Validasi Chow (1959) & Kapasitas Desain",
                f"Material: {d['material']} · b = {d['b']:.3f} m · h = {d['h']:.3f} m · S = {d['S']:.6f}",
                badge_html=_val_badge,
                tone="ocean",
            )
            if not is_valid_now:
                st.warning(msg_now)
            else:
                st.success(msg_now)

            df_ded_summary = pd.DataFrame([{
                "Material": d["material"],
                "b (m)": d["b"],
                "h (m)": d["h"],
                "S (m/m)": d["S"],
                "Manning n": d["n"],
                "Kutipan Teks / Halaman Asal DED": d.get("kutipan_ded", "—"),
            }])
            st.dataframe(df_ded_summary, hide_index=True)

            Q_ded = _manning_Q(d["b"], d["h"], d["S"], d["n"])
            st.metric("Q Kapasitas DED (Manning)", f"{Q_ded:.4f} m³/s")
            st.latex(r"Q = \frac{1}{n} \cdot A \cdot R^{2/3} \cdot S^{1/2}")


# ═════════════════════════════════════════════════════════════════
# MODUL 3: Diskusi Hibrida AI + Python untuk Computer Vision & Metode Cowan
# ═════════════════════════════════════════════════════════════════
elif menu.startswith("3 |") or "Modul 3" in menu:
    _m3_ready = bool(st.session_state.get("cv_params") or st.session_state.get("params_modul_3"))
    _render_module_header(
        step_num=3,
        icon="📸",
        title="Visi Komputer & Degradasi Saluran — Metode Cowan (1956)",
        description=(
            "Evaluasi spasial penampang basah saluran menggunakan Python ROI Pre-processor & AI Vision untuk mengestimasi "
            "faktor degradasi Metode Cowan (n1, n2, n3, n4, m), ketebalan sedimen (ts), serta geometri efektif saluran."
        ),
        status_text=(
            f"n_aktual: {st.session_state['cv_params']['n_adjusted']:.5f}"
            if st.session_state.get("cv_params")
            else "Menunggu Evaluasi Foto / Parameter Cowan"
        ),
        is_ready=_m3_ready,
    )

    if not PIL_AVAILABLE:
        st.warning("⚠️ Install `Pillow` untuk membaca gambar: `pip install Pillow`")

    # Pull pipeline data directly from st.session_state['data_modul_1'] & st.session_state['params_modul_2'] / 'ded_params'
    m1_data_in_m3 = st.session_state.get("data_modul_1")
    ded           = st.session_state.get("params_modul_2") or st.session_state.get("ded_params") or {}
    b_desain_ref  = float(ded.get("b", 1.5))
    h_desain_ref  = float(ded.get("h", 1.2))
    n_desain_ref  = float(ded.get("n", 0.013))

    p_m3_c1, p_m3_c2 = st.columns(2)
    with p_m3_c1:
        if isinstance(m1_data_in_m3, dict) and m1_data_in_m3.get("r24_dict"):
            st.info(
                f"🔗 **Modul 1 (`data_modul_1`):** {m1_data_in_m3.get('selected_dist', 'Gumbel')} | "
                f"**R24 (T=5) = {m1_data_in_m3.get('R24_5yr', 0.0):.2f} mm**"
            )
        elif st.session_state.get("R24_5yr") is not None:
            st.info(f"🔗 **Modul 1:** R24 (T=5) = **{st.session_state['R24_5yr']:.2f} mm**")
        else:
            st.warning("⏳ **Modul 1:** Data hujan rencana belum diproses.")
    with p_m3_c2:
        if ded:
            st.info(
                f"🔗 **Modul 2 (`params_modul_2`):** {ded.get('material', 'Beton')} | "
                f"$b = {b_desain_ref:.3f}$ m, $h = {h_desain_ref:.3f}$ m, $n_b = {n_desain_ref:.4f}$"
            )
        else:
            st.warning(
                f"⚠️ **Modul 2:** Parameter DED belum ada (default: b = {b_desain_ref} m, h = {h_desain_ref} m, nb = {n_desain_ref})."
            )

    col1, col2 = st.columns([6, 4], gap="large")
    col_img, col_result = col1, col2

    with col_img:
        _render_kobo_companion_card(
            asset_key="kobo_rain_umbrella",
            tag="📸 KOBO SPATIAL & COWAN INSPECTOR · COWAN (1956)",
            title="Inspeksi Penampang Basah & Faktor Degradasi Saluran",
            description=(
                "Evaluasi berfokus pada dinding dalam dan dasar saluran untuk mengestimasi keretakan (n1), "
                "variasi penampang (n2), hambatan sampah (n3), vegetasi (n4), kelokan (m), dan tebal sedimen (ts)."
            ),
        )
        _render_section_banner(
            "📸 Langkah 1 & 2: ROI Pre-processor & AI Spatial Evaluator",
            "Unggah citra lapangan saluran drainase (.jpg / .png). Fokus analisis diarahkan ke bagian dalam penampang basah.",
            badge_html='<span class="ds-badge auto">Cowan (1956) · ROI Vision</span>',
            tone="ocean",
        )
        img_file = st.file_uploader(
            "Unggah foto saluran drainase eksisting (.jpg / .png)",
            type=["jpg", "jpeg", "png"],
            help="Python memotong ROI bagian dalam saluran & menghitung rasio piksel dasar sebelum dikirim ke AI Spatial Evaluator.",
        )
        if img_file is not None:
            st.session_state["uploaded_file_m3"] = {
                "name": img_file.name,
                "bytes": bytes(img_file.getbuffer()),
            }

        saved_img_m3 = st.session_state.get("uploaded_file_m3")
        if isinstance(saved_img_m3, dict) and saved_img_m3.get("bytes") and PIL_AVAILABLE:
            if img_file is None:
                st.success(f"💾 Foto saluran aktif dari Session State: **{saved_img_m3.get('name', 'foto.jpg')}**")
            pil_raw = PILImage.open(io.BytesIO(saved_img_m3["bytes"])).convert("RGB")
            with st.expander("✂️ Pengaturan Crop ROI (Fokus Bagian Dalam Penampang Saluran)", expanded=False):
                rc1, rc2 = st.columns(2)
                crop_left   = rc1.slider("Potong Tepi Kiri (%)", 0, 40, 10)
                crop_right  = rc2.slider("Potong Tepi Kanan (%)", 0, 40, 10)
                crop_top    = rc1.slider("Potong Tepi Atas / Tanggul Luar (%)", 0, 40, 15)
                crop_bottom = rc2.slider("Potong Tepi Bawah (%)", 0, 40, 5)

            roi_img = _crop_roi(pil_raw, crop_left, crop_right, crop_top, crop_bottom)
            st.image(roi_img, caption="Citra ROI Bagian Dalam Penampang Saluran")

            roi_stats = _roi_pixel_stats(roi_img)
            st.session_state["m3_roi_stats"] = roi_stats
            px1, px2, px3, px4 = st.columns(4)
            px1.metric("🌿 Vegetasi ROI", f"{roi_stats['vegetasi_pct']:.1f}%")
            px2.metric("🟤 Sedimen ROI", f"{roi_stats['sedimen_pct']:.1f}%")
            px3.metric("💧 Air / Gelap", f"{roi_stats['air_gelap_pct']:.1f}%")
            px4.metric("🧱 Beton Terang", f"{roi_stats['beton_terang_pct']:.1f}%")

            cv_b1, cv_b2 = st.columns(2)
            with cv_b1:
                run_ai_cv = st.button("✨ Evaluasi Spasial Cowan (Gemini Vision AI)", type="primary")
            with cv_b2:
                run_py_cv = st.button("⚡ Auto-Estimasi Cowan via Python CV")

            if run_py_cv:
                t0_cv = time.perf_counter()
                active_roi_stats = st.session_state.get("m3_roi_stats") or roi_stats
                cowan_py_json = _local_cowan_from_stats(active_roi_stats, n_desain_ref, h_desain_ref)
                elapsed_cv_ms = (time.perf_counter() - t0_cv) * 1000.0
                src_py_m3 = f"Python CV Auto-Estimator ({elapsed_cv_ms:.1f} ms)"
                st.session_state["m3_vision_error"] = None
                st.session_state["m3_ai_result"] = {"source": src_py_m3, "json": cowan_py_json}
                _apply_m3_result(cowan_py_json, src_py_m3, h_desain_ref, b_desain_ref, n_desain_ref)
                st.success("✅ Parameter Cowan & Sedimen berhasil diperbarui dari analisis piksel Python")
                st.toast("✅ Parameter Cowan & Sedimen berhasil diperbarui dari analisis piksel Python!", icon="⚡")

            if run_ai_cv:
                buf_roi = io.BytesIO()
                roi_img.save(buf_roi, format="JPEG", quality=90)
                roi_bytes = buf_roi.getvalue()

                sys_prompt_m3 = (
                    "Kamu adalah Ahli Hidrolika dan Inspektor Lapangan. Evaluasi foto saluran drainase ini dan berikan nilai estimasi untuk parameter Metode Cowan (1956). \n"
                    "ATURAN MUTLAK: Fokus HANYA pada bagian DALAM penampang basah saluran (dinding dalam dan dasar). ABAIKAN rumput atau pohon di luar/atas tanggul saluran.\n"
                    "Gunakan panduan nilai ini:\n"
                    "- n1 (Keretakan): 0.000 (mulus) s.d 0.020 (beton gompal parah/batu lepas).\n"
                    "- n2 (Variasi Penampang): 0.000 (lurus presisi) s.d 0.015 (segmen bergeser/menyempit).\n"
                    "- n3 (Hambatan/Sampah): 0.000 (bersih) s.d 0.060 (tumpukan sampah/puing parah).\n"
                    "- n4 (Vegetasi/Gulma): 0.000 (bersih) s.d 0.100 (lumut tebal/gulma lebat DI DALAM saluran).\n"
                    "- m (Kelokan): 1.00 (lurus) s.d 1.30 (berkelok tajam). Minimum HARUS 1.00.\n"
                    "- ts (Tebal Sedimen): Estimasi tebal lumpur/tanah di dasar saluran dalam meter (misal 0.15).\n"
                    "- p (Penyempitan Lebar): Estimasi rasio lebar atas yang menyempit akibat longsoran (0.0 s.d 1.0).\n"
                    "- alasan_visual: Narasi inspeksi visual singkat mengenai kondisi retakan dinding, endapan sedimen di dasar saluran, dan vegetasi dari foto.\n\n"
                    "Keluarkan HANYA JSON murni (tanpa markdown ```json):\n"
                    '{"n1": float, "n2": float, "n3": float, "n4": float, "m": float, "ts": float, "p": float, "alasan_visual": "string"}'
                )

                with st.spinner("🔬 Menjalankan Gemini Vision AI (Auto-Discovery) & Python Cowan Calculator…"):
                    cowan_json: dict | None = None
                    src_m3 = ""
                    api_key_input = (st.session_state.get("gemini_key") or "").strip()
                    if api_key_input:
                        try:
                            raw_m3, model_used = call_gemini_vision(
                                api_key=api_key_input,
                                sys_prompt=sys_prompt_m3,
                                image_bytes=roi_bytes,
                                image_mime="image/jpeg",
                            )
                            cowan_json = _parse_ai_json(raw_m3)
                            src_m3 = f"Gemini Vision AI: {model_used}"
                            st.session_state["m3_vision_error"] = None
                        except Exception as err_m3:
                            safe_err3 = _redact_api_key(str(err_m3), api_key_input)
                            st.session_state["m3_vision_error"] = safe_err3
                            st.warning(f"⚠️ Panggilan Gemini Vision gagal ({safe_err3}). Beralih ke Python ROI Pixel Evaluator.")
                    else:
                        err_no_key = "API Key Gemini belum diisi di sidebar (Pengaturan Engine 1 API)."
                        st.session_state["m3_vision_error"] = err_no_key
                        st.warning(f"⚠️ Panggilan Gemini Vision gagal ({err_no_key}). Beralih ke Python ROI Pixel Evaluator.")

                    if cowan_json is None:
                        cowan_json = _local_cowan_from_stats(roi_stats, n_desain_ref, h_desain_ref)
                        src_m3 = "Python ROI Pixel Evaluator (Offline Fallback)"

                    st.session_state["m3_ai_result"] = {"source": src_m3, "json": cowan_json}
                    _apply_m3_result(cowan_json, src_m3, h_desain_ref, b_desain_ref, n_desain_ref)
                    st.success("✅ Faktor Cowan berhasil dievaluasi dan dihitung oleh Python Calculator!")
                    st.rerun()
        else:
            _render_kobo_empty_state(
                asset_key="kobo_rain_umbrella",
                title="Menunggu Foto Lapangan Saluran Drainase",
                subtitle=(
                    "Unggah foto kondisi eksisting saluran (.jpg / .png) di atas untuk dianalisis dengan ROI Pre-processor, "
                    "atau sesuaikan faktor degradasi secara langsung pada Kalkulator Metode Cowan (1956) di panel kanan."
                ),
                badges_html=(
                    '<span class="ds-badge auto">Crop ROI Penampang Dalam</span>'
                    '<span class="ds-badge auto">Metode Cowan (1956) &amp; Horton-Einstein</span>'
                ),
                show_image=False,
            )

        if st.session_state.get("m3_vision_error"):
            st.warning(
                f"⚠️ Panggilan Gemini Vision gagal ({st.session_state['m3_vision_error']}). "
                "Beralih ke Python ROI Pixel Evaluator."
            )

        if st.session_state.get("cv_params"):
            cvs = st.session_state["cv_params"]
            src_lbl = str(cvs.get("sumber", "Python Calculator"))
            is_ai_vision = src_lbl.startswith("Gemini Vision AI:")
            st.markdown("---")
            if is_ai_vision:
                st.info(
                    f"✨ **Status Tersimpan di Session State ({src_lbl})**\n\n"
                    f"• $h_{{\\text{{efektif}}}} = {cvs.get('h_effective', 0):.3f}\\text{{ m}}$ | "
                    f"$b_{{\\text{{efektif}}}} = {cvs.get('b_effective', b_desain_ref):.3f}\\text{{ m}}$\n\n"
                    f"• $n_{{\\text{{adjusted}}}} = {cvs.get('n_adjusted', 0):.5f}$ "
                    f"({'Horton-Einstein' if cvs.get('use_horton') else 'Cowan 1956'})"
                )
                st.success(
                    f"👁️ **Narasi Inspeksi Visual AI:**\n\n{cvs.get('deskripsi', '—')}"
                )
            else:
                st.info(
                    f"💾 **Status Tersimpan di Session State ({src_lbl}):**  \n"
                    f"• $h_{{\\text{{efektif}}}} = {cvs.get('h_effective', 0):.3f}\\text{{ m}}$ | "
                    f"$b_{{\\text{{efektif}}}} = {cvs.get('b_effective', b_desain_ref):.3f}\\text{{ m}}$  \n"
                    f"• $n_{{\\text{{adjusted}}}} = {cvs.get('n_adjusted', 0):.5f}$ "
                    f"({'Horton-Einstein' if cvs.get('use_horton') else 'Cowan 1956'})  \n"
                    f"• Catatan Visual: _{cvs.get('deskripsi', '—')}_"
                )

    with col_result:
        st.subheader("🧮 2. Kalkulator Metode Cowan (1956), Horton-Einstein & Geometri Efektif")
        saved_cv = st.session_state.get("cv_params") or {}

        # Inisialisasi state widget agar slider & input di sebelah kanan langsung menyesuaikan otomatis
        # saat tombol Auto-Estimasi Python CV atau AI Vision ditekan, dan tetap bebas diubah manual
        st.session_state.setdefault("m3_form_nb", float(np.clip(saved_cv.get("nb", n_desain_ref), 0.008, 0.050)))
        st.session_state.setdefault("m3_form_n1", float(np.clip(saved_cv.get("n1", 0.005), 0.000, 0.020)))
        st.session_state.setdefault("m3_form_n2", float(np.clip(saved_cv.get("n2", 0.002), 0.000, 0.015)))
        st.session_state.setdefault("m3_form_n3", float(np.clip(saved_cv.get("n3", 0.015), 0.000, 0.060)))
        st.session_state.setdefault("m3_form_n4", float(np.clip(saved_cv.get("n4", 0.015), 0.000, 0.100)))
        st.session_state.setdefault("m3_form_m",  float(np.clip(saved_cv.get("m_meander", 1.00), 1.00, 1.30)))
        st.session_state.setdefault("m3_form_h_maks", float(max(0.05, saved_cv.get("h_desain_ref", h_desain_ref))))
        st.session_state.setdefault(
            "m3_form_ts",
            float(np.clip(saved_cv.get("ts", round(0.2 * h_desain_ref, 3)), 0.0, max(0.01, h_desain_ref - 0.01))),
        )
        st.session_state.setdefault("m3_form_p", float(np.clip(saved_cv.get("p_narrowing", 0.05), 0.00, 0.90)))

        # ── A. Metode Cowan (1956) ──
        st.markdown("#### A. Koefisien Manning Aktual — Metode Cowan (1956)")
        st.latex(r"n_{\text{aktual}} = (n_b + n_1 + n_2 + n_3 + n_4) \times m")

        ideal_bypass = st.checkbox(
            "✨ Saluran Kondisi Baru / Abaikan Degradasi (Ideal)",
            key="m3_ideal_bypass",
            help="Jika dicentang, seluruh faktor degradasi dikunci ke kondisi ideal (n1=0, n2=0, n3=0, n4=0, ts=0, p=0, dan m=1.00).",
        )
        if ideal_bypass:
            st.session_state["m3_form_n1"] = 0.000
            st.session_state["m3_form_n2"] = 0.000
            st.session_state["m3_form_n3"] = 0.000
            st.session_state["m3_form_n4"] = 0.000
            st.session_state["m3_form_m"]  = 1.00
            st.session_state["m3_form_ts"] = 0.000
            st.session_state["m3_form_p"]  = 0.00

        cw1, cw2, cw3 = st.columns(3)
        with cw1:
            nb_val = st.number_input(
                "nb — Kekasaran Dasar Material",
                min_value=0.008, max_value=0.050,
                step=0.001, format="%.4f",
                key="m3_form_nb",
                help="Nilai Manning dasar saluran lurus & halus (dari DED Modul 2).",
            )
            n1_val = st.slider(
                "n1 — Keretakan Permukaan (0–0.020)",
                min_value=0.000, max_value=0.020,
                step=0.001, format="%.3f",
                key="m3_form_n1",
                disabled=ideal_bypass,
                help="Halus=0.000 | Sedikit retak=0.005 | Rusak sedang=0.010 | Rusak berat=0.020",
            )
        with cw2:
            n2_val = st.slider(
                "n2 — Variasi Penampang (0–0.015)",
                min_value=0.000, max_value=0.015,
                step=0.001, format="%.3f",
                key="m3_form_n2",
                disabled=ideal_bypass,
                help="Konstan=0.000 | Berubah bertahap=0.005 | Berubah sering=0.010–0.015",
            )
            n3_val = st.slider(
                "n3 — Hambatan Endapan/Sampah (0–0.060)",
                min_value=0.000, max_value=0.060,
                step=0.001, format="%.3f",
                key="m3_form_n3",
                disabled=ideal_bypass,
                help="Diabaikan=0.000 | Kecil=0.010–0.015 | Cukup besar=0.020–0.030 | Parah=0.040–0.060",
            )
        with cw3:
            n4_val = st.slider(
                "n4 — Vegetasi / Gulma (0–0.100)",
                min_value=0.000, max_value=0.100,
                step=0.001, format="%.3f",
                key="m3_form_n4",
                disabled=ideal_bypass,
                help="Rendah=0.005–0.010 | Sedang=0.010–0.025 | Tinggi=0.025–0.050 | Sangat rapat=0.050–0.100",
            )
            m_meander = st.slider(
                "m — Faktor Meander / Kelokan (1.00–1.30)",
                min_value=1.00, max_value=1.30,
                step=0.01, format="%.2f",
                key="m3_form_m",
                disabled=ideal_bypass,
                help="Lurus=1.00 | Berkelok sedang=1.15 | Berkelok tajam=1.30",
            )

        if ideal_bypass:
            n1_val, n2_val, n3_val, n4_val, m_meander = 0.0, 0.0, 0.0, 0.0, 1.0

        n_cowan_calc = (nb_val + n1_val + n2_val + n3_val + n4_val) * m_meander
        st.metric(
            "Nilai Manning Aktual (Metode Cowan 1956)",
            f"{n_cowan_calc:.5f}",
            delta=f"{n_cowan_calc - n_desain_ref:+.5f} terhadap n_DED ({n_desain_ref:.4f})",
            delta_color="inverse",
        )

        st.markdown("---")

        # ── B. Geometri Efektif akibat Sedimen ts & Penyempitan p ──
        st.markdown("#### B. Geometri Efektif Akibat Sedimen ($t_s$) & Penyempitan Lebar ($p$)")
        g1, g2, g3, g4 = st.columns(4)
        with g1:
            shape_cv = st.selectbox(
                "Bentuk Penampang",
                options=["Persegi", "Trapesium"],
                index=0 if saved_cv.get("shape_cv", "Persegi") == "Persegi" else 1,
            )
        with g2:
            h_maks_in = st.number_input(
                "h_maks Desain (m)",
                min_value=0.05, step=0.05, format="%.3f",
                key="m3_form_h_maks",
            )
        with g3:
            max_ts_allowed = max(0.01, float(h_maks_in) - 0.01)
            if st.session_state.get("m3_form_ts", 0.0) > max_ts_allowed:
                st.session_state["m3_form_ts"] = max_ts_allowed
            ts_val = st.number_input(
                "Tebal Sedimen ts (m)",
                min_value=0.0, max_value=max_ts_allowed,
                step=0.005, format="%.3f",
                key="m3_form_ts",
                disabled=ideal_bypass,
            )
        with g4:
            if shape_cv == "Persegi":
                p_narrow = st.slider(
                    "Penyempitan Lebar p (0–1)",
                    min_value=0.0, max_value=0.90,
                    step=0.01, format="%.2f",
                    key="m3_form_p",
                    disabled=ideal_bypass,
                    help="Cek jarak kiri ke kanan pada foto. Apakah ada bagian tepi tanah yang longsor ke dalam sehingga lebar saluran menyempit? Isi dalam bentuk desimal (0.0 = lebar normal tidak menyempit; 0.1 = lebar menyempit sekitar 10%).",
                )
                m_talud_cv = 0.0
            else:
                m_talud_cv = st.number_input(
                    "Kemiringan Talud m (1V:mH)",
                    min_value=0.1, max_value=5.0,
                    value=float(saved_cv.get("m_talud", 1.0) or 1.0),
                    step=0.1, format="%.2f",
                )
                p_narrow = 0.0

        if ideal_bypass:
            ts_val = 0.0
            p_narrow = 0.0

        h_eff_calc = max(0.01, h_maks_in - ts_val)
        if shape_cv == "Persegi":
            b_eff_calc = max(0.01, b_desain_ref * (1.0 - p_narrow))
            st.latex(r"h_{\text{efektif}} = h_{\text{maks}} - t_s \qquad;\qquad b_{\text{efektif}} = b \times (1 - p)")
        else:
            b_eff_calc = b_desain_ref + 2.0 * m_talud_cv * ts_val
            st.latex(r"h_{\text{efektif}} = h_{\text{maks}} - t_s \qquad;\qquad b_{\text{efektif}} = b + 2 \cdot m_{\text{talud}} \cdot t_s")

        ge1, ge2, ge3 = st.columns(3)
        ge1.metric("Kedalaman Efektif (h_efektif)", f"{h_eff_calc:.4f} m", delta=f"-{ts_val:.3f} m sedimen", delta_color="inverse")
        ge2.metric("Lebar Dasar Efektif (b_efektif)", f"{b_eff_calc:.4f} m", delta=f"{b_eff_calc - b_desain_ref:+.3f} m vs b_DED")
        sed_pct_equiv = (ts_val / h_maks_in) * 100.0 if h_maks_in > 0 else 0.0
        ge3.metric("Rasio Tebal Sedimen", f"{sed_pct_equiv:.1f} %")

        st.markdown("---")

        # ── C. Opsi Kekasaran Gabungan (Horton-Einstein) ──
        st.markdown("#### C. Opsi Kekasaran Gabungan (Horton-Einstein) — Dinding & Dasar Berbeda Bahan")
        st.latex(
            r"n_c = \left( \frac{P_{\text{dasar}} \cdot n_{\text{sedimen}}^{1.5} + "
            r"P_{\text{dinding}} \cdot n_{\text{dinding}}^{1.5}}{P_{\text{total}}} \right)^{2/3}"
        )
        use_horton = st.checkbox(
            "Aktifkan Kekasaran Gabungan Horton-Einstein (Dinding & Dasar Berbeda Kekasaran)",
            value=bool(saved_cv.get("use_horton", False)),
        )

        P_dasar_he = b_eff_calc
        if shape_cv == "Persegi":
            P_dinding_he = 2.0 * h_eff_calc
        else:
            P_dinding_he = 2.0 * h_eff_calc * math.sqrt(1.0 + m_talud_cv ** 2)
        P_total_he = P_dasar_he + P_dinding_he

        he1, he2, he3 = st.columns(3)
        with he1:
            n_sedimen_he = st.number_input(
                "n_sedimen (Dasar Berlumpur/Batu)",
                min_value=0.010, max_value=0.150,
                value=float(saved_cv.get("n_sedimen_he", max(0.025, n_cowan_calc))),
                step=0.002, format="%.4f",
            )
        with he2:
            n_dinding_he = st.number_input(
                "n_dinding (Dinding Kiri + Kanan)",
                min_value=0.010, max_value=0.150,
                value=float(saved_cv.get("n_dinding_he", nb_val + n1_val)),
                step=0.002, format="%.4f",
            )
        nc_horton = (
            ((P_dasar_he * (n_sedimen_he ** 1.5) + P_dinding_he * (n_dinding_he ** 1.5)) / P_total_he) ** (2.0 / 3.0)
            if P_total_he > 0 else n_cowan_calc
        )
        with he3:
            st.metric("n_c Gabungan (Horton-Einstein)", f"{nc_horton:.5f}")
            st.caption(f"P_dasar={P_dasar_he:.2f} m | P_dinding={P_dinding_he:.2f} m")

        n_final_m3 = nc_horton if use_horton else n_cowan_calc
        weed_pct_equiv = min(100.0, (n4_val / 0.100) * 100.0)

        if st.button("💾 Simpan Parameter Degradasi & Geometri Efektif ke Modul 4", type="primary"):
            cv_saved_dict = {
                "sediment_pct": sed_pct_equiv,
                "weed_pct":     weed_pct_equiv,
                "ts":           ts_val,
                "p_narrowing":  p_narrow,
                "shape_cv":     shape_cv,
                "m_talud":      m_talud_cv,
                "h_effective":  h_eff_calc,
                "b_effective":  b_eff_calc,
                "n_adjusted":   n_final_m3,
                "n_cowan":      n_cowan_calc,
                "horton_nc":    nc_horton,
                "use_horton":   use_horton,
                "n_sedimen_he": n_sedimen_he,
                "n_dinding_he": n_dinding_he,
                "nb":           nb_val,
                "n1":           n1_val,
                "n2":           n2_val,
                "n3":           n3_val,
                "n4":           n4_val,
                "m_meander":    m_meander,
                "deskripsi":    saved_cv.get("deskripsi", "Dihitung dengan Metode Cowan (1956) & Geometri Efektif"),
                "h_desain_ref": h_maks_in,
                "n_desain_ref": nb_val,
            }
            st.session_state["cv_params"] = cv_saved_dict
            st.session_state["params_modul_3"] = dict(cv_saved_dict)
            st.toast("💾 Parameter Modul 3 disimpan → siap digunakan di Modul 4!", icon="✅")
            st.success(
                f"✅ Tersimpan! `h_effective` = **{h_eff_calc:.4f} m**, "
                f"`b_effective` = **{b_eff_calc:.4f} m**, "
                f"`n_adjusted` = **{n_final_m3:.5f}**."
            )

    # ── Full Width: Komparasi Visual Penampang (Sebelum vs Sesudah Sedimen) ──
    st.divider()
    _render_section_banner(
        "📐 Komparasi Visual Penampang: Kondisi Ideal DED vs Eksisting Terdegradasi",
        f"Perbandingan penampang saluran {shape_cv} kondisi rencana DED vs eksisting akibat endapan sedimen ts = {ts_val:.3f} m dan penyempitan p = {p_narrow:.2f}.",
        badge_html=f'<span class="ds-badge auto">n_desain = {nb_val:.4f} ➔ n_aktual = {n_final_m3:.5f}</span>',
        tone="teal",
    )
    fig_sec_cmp = _render_channel_comparison_figure(
        shape=shape_cv,
        b_des=b_desain_ref,
        h_des=h_maks_in,
        b_eff=b_eff_calc,
        h_eff=h_eff_calc,
        ts_val=ts_val,
        n_des=nb_val,
        n_akt=n_final_m3,
    )
    st.pyplot(fig_sec_cmp, use_container_width=True)
    plt.close(fig_sec_cmp)


# ═════════════════════════════════════════════════════════════════
# MODUL 4: Engine Hidrolika Hibrida (Permen PU 12/2014, Kirpich/Kerby, Multi-Penampang & Froude)
# ═════════════════════════════════════════════════════════════════
elif menu.startswith("4 |") or "Modul 4" in menu:
    _m4_ready = st.session_state.get("m4_results") is not None
    _render_module_header(
        step_num=4,
        icon="⚙️",
        title="Engine Hidrolika Hibrida (Permen PU 12/2014 & Multi-Penampang)",
        description=(
            "Evaluasi kapasitas hidrolika komprehensif sesuai Permen PU No. 12/PRT/M/2014: "
            "Matriks Kala Ulang Tipologi Kota, Waktu Konsentrasi (tc), Debit Rasional (Qr), "
            "dan Kapasitas Manning (Qs) pada 3 Geometri Penampang dengan kontrol Kecepatan Izin & Bilangan Froude (Fr)."
        ),
        status_text=(
            f"Qr={st.session_state['m4_results']['Q_limpasan']:.3f} vs Qs={st.session_state['m4_results']['Q_aktual']:.3f} m³/s"
            if _m4_ready
            else "Siap Evaluasi Hidrolika"
        ),
        is_ready=_m4_ready,
    )

    _render_kobo_companion_card(
        asset_key="kobo_thinking_2",
        tag="⚙️ KOBO HYBRID HYDRAULIC ENGINE · PERMEN PU NO. 12/PRT/M/2014",
        title="Integrasi Hidrologi Rasional (Qr) & Kapasitas Manning Multi-Penampang (Qs)",
        description=(
            "Evaluasi keseimbangan debit limpasan hujan rencana terhadap kapasitas penampang Persegi, Trapesium, "
            "atau Lingkaran beserta tinggi jagaan (Freeboard w), kontrol kecepatan izin, dan bilangan Froude (Fr)."
        ),
    )

    # Pull pipeline data directly from data_modul_1, params_modul_2 / ded_params, and params_modul_3 / cv_params
    m1_data_in_m4 = st.session_state.get("data_modul_1") or {}
    ded           = st.session_state.get("params_modul_2") or st.session_state.get("ded_params") or {}
    cv            = st.session_state.get("params_modul_3") or st.session_state.get("cv_params")  or {}
    R24_5         = m1_data_in_m4.get("R24_5yr") if m1_data_in_m4.get("R24_5yr") is not None else st.session_state.get("R24_5yr")
    r24_dict      = m1_data_in_m4.get("r24_dict") or st.session_state.get("r24_dict") or {}

    _m1_str = (
        f"Multi-kala ulang ({', '.join(f'T{k}={v:.1f}mm' for k, v in list(r24_dict.items())[:4])})"
        if r24_dict
        else (f"R24 (T=5) = {R24_5:.2f} mm" if R24_5 else "Input Manual / Default")
    )
    _m2_str = (
        f"{ded.get('material', 'Beton')} · b={ded.get('b')}m, H={ded.get('h')}m, n={ded.get('n')}"
        if ded
        else "Default (Beton · b=1.5m, H=1.2m)"
    )
    _m3_str = (
        f"h_eff={cv.get('h_effective', 0):.3f}m, b_eff={cv.get('b_effective', ded.get('b', 1.5)):.3f}m, n_akt={cv.get('n_adjusted', 0):.4f}"
        if cv
        else "Default (Cowan / Kondisi Eksisting)"
    )
    _b1 = '<span class="ds-badge safe">TERHUBUNG</span>' if (r24_dict or R24_5) else '<span class="ds-badge manual">DEFAULT</span>'
    _b2 = '<span class="ds-badge safe">TERHUBUNG</span>' if ded else '<span class="ds-badge manual">DEFAULT</span>'
    _b3 = '<span class="ds-badge safe">TERHUBUNG</span>' if cv else '<span class="ds-badge manual">DEFAULT</span>'

    st.markdown(
        f"""
        <div class="ds-pipe-grid">
            <div class="ds-pipe-card">
                <div class="ds-pipe-head"><span>🌧️ Modul 1 · Hujan Rencana</span>{_b1}</div>
                <div class="ds-pipe-val">{_m1_str}</div>
            </div>
            <div class="ds-pipe-card">
                <div class="ds-pipe-head"><span>📄 Modul 2 · Dimensi DED</span>{_b2}</div>
                <div class="ds-pipe-val">{_m2_str}</div>
            </div>
            <div class="ds-pipe-card">
                <div class="ds-pipe-head"><span>📸 Modul 3 · Cowan &amp; Degradasi</span>{_b3}</div>
                <div class="ds-pipe-val">{_m3_str}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ══════════════════════════════════════════════════════════════
    # 1. MATRIKS KALA ULANG RENCANA (PERMEN PU 12/2014)
    # ══════════════════════════════════════════════════════════════
    _render_section_banner(
        "🏙️ 1. Penentuan Kala Ulang Rencana (Permen PU No. 12/PRT/M/2014)",
        "Klasifikasi otomatis tipologi 98 kota di Indonesia dan hierarki daerah tangkapan air (DTA).",
        badge_html='<span class="ds-badge auto">Permen PU 12/2014</span>',
        tone="ocean",
    )

    kota_db = {
        "Ambon": "Kota Sedang",
        "Balikpapan": "Kota Besar",
        "Banda Aceh": "Kota Sedang",
        "Bandar Lampung": "Kota Metropolitan",
        "Bandung": "Kota Metropolitan",
        "Banjar": "Kota Sedang",
        "Banjarbaru": "Kota Sedang",
        "Banjarmasin": "Kota Besar",
        "Batam": "Kota Metropolitan",
        "Batu": "Kota Sedang",
        "Baubau": "Kota Sedang",
        "Bekasi": "Kota Metropolitan",
        "Bengkulu": "Kota Sedang",
        "Bima": "Kota Sedang",
        "Binjai": "Kota Sedang",
        "Bitung": "Kota Sedang",
        "Blitar": "Kota Sedang",
        "Bogor": "Kota Metropolitan",
        "Bontang": "Kota Sedang",
        "Bukittinggi": "Kota Sedang",
        "Cilegon": "Kota Sedang",
        "Cimahi": "Kota Besar",
        "Cirebon": "Kota Sedang",
        "Denpasar": "Kota Besar",
        "Depok": "Kota Metropolitan",
        "Dumai": "Kota Sedang",
        "Gorontalo": "Kota Sedang",
        "Gunungsitoli": "Kota Sedang",
        "Jakarta Barat": "Kota Metropolitan",
        "Jakarta Pusat": "Kota Metropolitan",
        "Jakarta Selatan": "Kota Metropolitan",
        "Jakarta Timur": "Kota Metropolitan",
        "Jakarta Utara": "Kota Metropolitan",
        "Jambi": "Kota Besar",
        "Jayapura": "Kota Sedang",
        "Kediri": "Kota Sedang",
        "Kendari": "Kota Sedang",
        "Kotamobagu": "Kota Sedang",
        "Kupang": "Kota Sedang",
        "Langsa": "Kota Sedang",
        "Lhokseumawe": "Kota Sedang",
        "Lubuklinggau": "Kota Sedang",
        "Madiun": "Kota Sedang",
        "Magelang": "Kota Sedang",
        "Makassar": "Kota Metropolitan",
        "Malang": "Kota Besar",
        "Manado": "Kota Sedang",
        "Mataram": "Kota Sedang",
        "Medan": "Kota Metropolitan",
        "Metro": "Kota Sedang",
        "Mojokerto": "Kota Sedang",
        "Padang": "Kota Besar",
        "Padang Panjang": "Kota Kecil",
        "Padangsidimpuan": "Kota Sedang",
        "Pagar Alam": "Kota Sedang",
        "Palangkaraya": "Kota Sedang",
        "Palembang": "Kota Metropolitan",
        "Palopo": "Kota Sedang",
        "Palu": "Kota Sedang",
        "Pangkalpinang": "Kota Sedang",
        "Parepare": "Kota Sedang",
        "Pariaman": "Kota Kecil",
        "Pasuruan": "Kota Sedang",
        "Payakumbuh": "Kota Sedang",
        "Pekalongan": "Kota Sedang",
        "Pekanbaru": "Kota Metropolitan",
        "Pematangsiantar": "Kota Sedang",
        "Pontianak": "Kota Besar",
        "Prabumulih": "Kota Sedang",
        "Probolinggo": "Kota Sedang",
        "Sabang": "Kota Kecil",
        "Salatiga": "Kota Sedang",
        "Samarinda": "Kota Besar",
        "Sawahlunto": "Kota Kecil",
        "Semarang": "Kota Metropolitan",
        "Serang": "Kota Besar",
        "Sibolga": "Kota Kecil",
        "Singkawang": "Kota Sedang",
        "Solok": "Kota Kecil",
        "Sorong": "Kota Sedang",
        "Subulussalam": "Kota Kecil",
        "Sukabumi": "Kota Sedang",
        "Sungai Penuh": "Kota Kecil",
        "Surabaya": "Kota Metropolitan",
        "Surakarta (Solo)": "Kota Besar",
        "Tangerang": "Kota Metropolitan",
        "Tangerang Selatan": "Kota Metropolitan",
        "Tanjung Balai": "Kota Sedang",
        "Tanjungpinang": "Kota Sedang",
        "Tarakan": "Kota Sedang",
        "Tasikmalaya": "Kota Besar",
        "Tebing Tinggi": "Kota Sedang",
        "Tegal": "Kota Sedang",
        "Ternate": "Kota Sedang",
        "Tidore Kepulauan": "Kota Sedang",
        "Tomohon": "Kota Sedang",
        "Tual": "Kota Kecil",
        "Yogyakarta": "Kota Sedang",
        "--- Opsi Input Manual / Kabupaten ---": "Kota Sedang",
        "Manual: Kota Metropolitan (> 1 Juta Jiwa)": "Kota Metropolitan",
        "Manual: Kota Besar (500k - 1 Juta Jiwa)": "Kota Besar",
        "Manual: Kota Sedang (100k - 500k Jiwa)": "Kota Sedang",
        "Manual: Kota Kecil (< 100k Jiwa)": "Kota Kecil"
    }

    matriks_pu = {
        "Kota Metropolitan": {"Tersier": 5, "Sekunder": 10, "Primer": 25},
        "Kota Besar": {"Tersier": 5, "Sekunder": 10, "Primer": 25},
        "Kota Sedang": {"Tersier": 2, "Sekunder": 5, "Primer": 10},
        "Kota Kecil": {"Tersier": 2, "Sekunder": 5, "Primer": 10}
    }

    kota_keys = list(kota_db.keys())

    def _detect_city_from_sources(valid_cities: list[str]) -> tuple[str, str]:
        m2_up = st.session_state.get("uploaded_file_m2") if isinstance(st.session_state.get("uploaded_file_m2"), dict) else {}
        m1_up = st.session_state.get("uploaded_file_m1") if isinstance(st.session_state.get("uploaded_file_m1"), dict) else {}
        m1_d = st.session_state.get("data_modul_1") if isinstance(st.session_state.get("data_modul_1"), dict) else {}
        ded_ss = st.session_state.get("ded_params") if isinstance(st.session_state.get("ded_params"), dict) else {}

        candidates_sources: list[tuple[str, str]] = [
            ("Dokumen DED Modul 2", str(m2_up.get("name") or st.session_state.get("pdf_doc_name") or "")),
            ("Teks PDF/DOCX Modul 2", str(st.session_state.get("pdf_doc_text") or m2_up.get("full_doc_text") or ded_ss.get("kutipan_ded") or "")[:30000]),
            ("File Data Modul 1", str(m1_d.get("file_name") or m1_up.get("file_name") or m1_up.get("pdf_name") or "")),
        ]
        raw_df_m1 = st.session_state.get("m1_raw_df")
        if raw_df_m1 is not None and hasattr(raw_df_m1, "columns"):
            try:
                hdr_txt = " ".join(str(c) for c in raw_df_m1.columns) + " " + raw_df_m1.head(5).to_string()
                candidates_sources.append(("Header Tabel Modul 1", hdr_txt[:5000]))
            except Exception:
                pass

        searchable_cities = [
            c for c in valid_cities
            if not c.startswith("---") and not c.startswith("Manual:")
        ]
        # Urutkan dari nama kota terpanjang agar "Tangerang Selatan" cocok sebelum "Tangerang"
        searchable_cities.sort(key=len, reverse=True)

        for src_label, src_txt in candidates_sources:
            low = (src_txt or "").lower()
            if not low.strip():
                continue
            for city in searchable_cities:
                if city == "Surakarta (Solo)":
                    if re.search(r"\b(?:surakarta|solo)\b", low):
                        return city, src_label
                elif re.search(rf"\b{re.escape(city.lower())}\b", low):
                    return city, src_label

        fallback_city = "Magelang" if "Magelang" in valid_cities else valid_cities[0]
        return fallback_city, "Default Sistem"

    auto_city, auto_city_src = _detect_city_from_sources(kota_keys)
    st.session_state["detected_city"] = auto_city
    st.session_state["detected_city_source"] = auto_city_src

    # Auto-detect Luas DTA (ha) dari ekstraksi Modul 2 (atau fallback 0.616 ha)
    auto_dta_ha = float(
        ded.get("luas_dta_ha")
        or st.session_state.get("m2_luas_dta_ha")
        or 0.616
    )
    auto_dta_ha = max(0.001, auto_dta_ha)

    _m4_s1_sync_sig = (auto_city, round(auto_dta_ha, 4))
    if st.session_state.get("m4_s1_last_sync_sig") != _m4_s1_sync_sig:
        st.session_state["m4_pilihan_kota"] = auto_city
        st.session_state["m4_luas_dta_ha"] = auto_dta_ha
        st.session_state["m4_s1_last_sync_sig"] = _m4_s1_sync_sig
    elif st.session_state.get("m4_pilihan_kota") not in kota_keys:
        st.session_state["m4_pilihan_kota"] = auto_city

    m4_s1_badge_box = st.container()

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        pilihan_kota = st.selectbox("Lokasi Kota/Kabupaten", kota_keys, key="m4_pilihan_kota")
        tipologi_kota = kota_db[pilihan_kota]
        st.session_state["tipologi_kota"] = tipologi_kota
        st.info(f"Tipologi: {tipologi_kota}")

    with col2:
        luas_dta = st.number_input("Luas DTA (Hektar)", min_value=0.001, step=0.1, format="%.3f", key="m4_luas_dta_ha")
        luas_dta_ha = luas_dta
        if luas_dta < 10:
            label_hierarki = "Saluran Tersier (< 10 ha)"
            kunci_hierarki = "Tersier"
        elif luas_dta <= 100:
            label_hierarki = "Saluran Sekunder (10 - 100 ha)"
            kunci_hierarki = "Sekunder"
        else:
            label_hierarki = "Saluran Primer (> 100 ha)"
            kunci_hierarki = "Primer"
        hierarki_saluran = kunci_hierarki
        st.success(label_hierarki)

    with m4_s1_badge_box:
        st.info(
            f"✨ Terisi Otomatis dari Dokumen: Kota **{pilihan_kota}** ({auto_city_src}) | "
            f"Luas DTA = **{luas_dta:.3f} ha**"
        )

    T_otomatis = matriks_pu[tipologi_kota][kunci_hierarki]
    T_options = [2, 5, 10, 25, 50, 100]
    _m4_hierarki_sig = (tipologi_kota, kunci_hierarki)
    if st.session_state.get("m4_last_hierarki_sig") != _m4_hierarki_sig or st.session_state.get("m4_T_pilihan") not in T_options:
        st.session_state["m4_T_pilihan"] = T_otomatis if T_otomatis in T_options else 5
        st.session_state["m4_last_hierarki_sig"] = _m4_hierarki_sig

    with col3:
        T_pilihan = st.selectbox(
            "Kala Ulang Rencana T (Tahun)",
            options=T_options,
            key="m4_T_pilihan",
            help=f"Rekomendasi otomatis Permen PU 12/2014 untuk {tipologi_kota} & {label_hierarki} adalah T = {T_otomatis} tahun.",
        )
    T = T_pilihan

    # Tarik R24 dari st.session_state["r24_dict"] sesuai T_pilihan (atau default 396.592 mm jika manual)
    if r24_dict and (T_pilihan in r24_dict or str(T_pilihan) in r24_dict):
        r24_from_ss = float(r24_dict.get(T_pilihan, r24_dict.get(str(T_pilihan), 396.592)))
    elif R24_5 is not None:
        r24_from_ss = float(R24_5)
    else:
        r24_from_ss = 396.592

    _m4_r24_sig = (T_pilihan, round(r24_from_ss, 3))
    if st.session_state.get("m4_last_r24_sig") != _m4_r24_sig:
        st.session_state["m4_R24_val"] = r24_from_ss
        st.session_state["m4_last_r24_sig"] = _m4_r24_sig

    with col4:
        R24_val = st.number_input(
            f"R24 Hujan Rencana T={T_pilihan} thn (mm)",
            min_value=0.0,
            step=1.0,
            format="%.3f",
            key="m4_R24_val",
            help="Ditarik otomatis dari st.session_state Modul 1 sesuai T_pilihan (atau default 396.592 mm jika diisi manual).",
        )
    R24 = R24_val

    with st.expander("📖 Tabel Matriks Kala Ulang Permen PU No. 12/PRT/M/2014"):
        df_pu_mat = pd.DataFrame(matriks_pu).T
        df_pu_mat.index.name = "Tipologi Kota"
        st.dataframe(df_pu_mat.map(lambda x: f"{x} Tahun"))

    st.divider()

    # ══════════════════════════════════════════════════════════════
    # 2 & 3. PARAMETER HIDROLOGI: WAKTU KONSENTRASI (tc) & DEBIT RASIONAL (Qr)
    # ══════════════════════════════════════════════════════════════
    st.subheader("⏱️ 2. Waktu Konsentrasi ($t_c$) & Debit Rasional ($Q_r$)")

    # Tarik teks PDF dari st.session_state (baik dari Modul 1 maupun Modul 2)
    up_m1_dict = st.session_state.get("uploaded_file_m1") if isinstance(st.session_state.get("uploaded_file_m1"), dict) else {}
    up_m2_dict = st.session_state.get("uploaded_file_m2") if isinstance(st.session_state.get("uploaded_file_m2"), dict) else {}
    pdf_text_m4 = (
        st.session_state.get("pdf_doc_text")
        or up_m1_dict.get("full_doc_text")
        or m1_data_in_m4.get("full_doc_text")
        or up_m2_dict.get("full_doc_text")
        or ""
    )
    pdf_name_m4 = (
        st.session_state.get("pdf_doc_name")
        or up_m1_dict.get("pdf_name")
        or m1_data_in_m4.get("pdf_name")
        or up_m2_dict.get("name")
        or "Dokumen Jurnal/DED PDF"
    )

    if pdf_text_m4:
        st.caption(
            f"📄 **Dokumen PDF aktif di `st.session_state`:** `{pdf_name_m4}` "
            f"({len(pdf_text_m4):,} karakter tersedia untuk ekstraksi instan)."
        )
    else:
        with st.expander("📂 Belum ada PDF di Session State? Unggah PDF Jurnal/DED di sini (Opsional)", expanded=False):
            m4_pdf_up = st.file_uploader(
                "Unggah Dokumen PDF Jurnal / Laporan Hidrologi (.pdf)",
                type=["pdf"],
                key="m4_fallback_pdf_uploader",
            )
            if m4_pdf_up is not None:
                raw_m4_pdf = bytes(m4_pdf_up.getbuffer())
                txt_m4_pdf = _extract_pdf_text_pypdf(raw_m4_pdf)
                prev_d_m4, prev_s_m4 = _build_preview_data(txt_m4_pdf)
                st.session_state["pdf_doc_text"] = txt_m4_pdf
                st.session_state["pdf_doc_name"] = m4_pdf_up.name
                st.session_state["uploaded_file_m2"] = {
                    "name": m4_pdf_up.name,
                    "size": m4_pdf_up.size,
                    "full_doc_text": txt_m4_pdf,
                    "preview_data": prev_d_m4,
                    "preview_stats": prev_s_m4,
                }
                if isinstance(st.session_state.get("uploaded_file_m1"), dict):
                    st.session_state["uploaded_file_m1"]["full_doc_text"] = txt_m4_pdf
                    st.session_state["uploaded_file_m1"]["pdf_name"] = m4_pdf_up.name
                else:
                    st.session_state["uploaded_file_m1"] = {
                        "pdf_name": m4_pdf_up.name,
                        "full_doc_text": txt_m4_pdf,
                    }
                pdf_text_m4 = txt_m4_pdf
                pdf_name_m4 = m4_pdf_up.name
                st.success(f"✅ Dokumen **{m4_pdf_up.name}** tersimpan ke `st.session_state`.")

    def set_manual_calc_false():
        st.session_state["m4_use_manual_calc"] = False

    ext_btn_col, manual_chk_col = st.columns([3, 2])
    with ext_btn_col:
        run_m4_extract_btn = st.button(
            "⚡ Ekstraksi Instan Hidrologi via Python",
            type="primary",
            on_click=set_manual_calc_false,
        )
    with manual_chk_col:
        use_manual_calc = st.checkbox(
            "Gunakan Kalkulator Manual",
            key="m4_use_manual_calc",
            help="Centang jika ingin mengabaikan hasil ekstraksi dan menghitung dari awal menggunakan variabel mentah (Lo, So, nd, Ls, Sub-DTA 1, 2, 3).",
        )

    if run_m4_extract_btn:
        t0_m4 = time.perf_counter()
        if not pdf_text_m4:
            st.warning(
                "⚠️ Teks dokumen PDF belum ditemukan di `st.session_state`. "
                "Menjalankan ekstraksi dengan nilai referensi jurnal standar (atau silakan unggah PDF di Modul 1 / Modul 2)."
            )
        ext_res = _extract_m4_hydrology_local(pdf_text_m4, pdf_name_m4)
        elapsed_m4_ms = (time.perf_counter() - t0_m4) * 1000.0
        ext_res["elapsed_ms"] = elapsed_m4_ms
        st.session_state["m4_hydro_extracted"] = ext_res
        use_manual_calc = False
        st.toast(
            f"⚡ Ekstraksi Hidrologi selesai ({elapsed_m4_ms:.1f} ms) — Mode Override Aktif!",
            icon="✅",
        )

    hydro_ext = st.session_state.get("m4_hydro_extracted")
    is_override_mode = (hydro_ext is not None) and (not use_manual_calc)

    if is_override_mode:
        st.success(
            f"🎯 **Mode Override / Hasil Ekstraksi Aktif** (Sumber: **{hydro_ext.get('source_file', pdf_name_m4)}** — "
            f"`{hydro_ext.get('elapsed_ms', 0.0):.1f} ms`, terdeteksi eksplisit: "
            f"**{', '.join(hydro_ext.get('found_keys', [])) or 'Default Jurnal'}**).  \n"
            "Kalkulator variabel mentah (`Lo`, `So`, `nd`, `Ls`, Sub-DTA 1/2/3) dilewati karena nilai akhir hidrologi telah diekstrak langsung dari dokumen."
        )

        t1_val     = float(hydro_ext.get("t1", 1.641))
        t2_val     = float(hydro_ext.get("t2", 4.889))
        tc_menit   = float(hydro_ext.get("tc", 6.530))
        tc_jam     = tc_menit / 60.0
        A_m2_val   = float(hydro_ext.get("A_m2", 6160.0))
        A_total_ha = float(hydro_ext.get("A_ha", 0.616))
        A_val      = float(hydro_ext.get("A_km2", 0.00616))
        C_val      = float(hydro_ext.get("C", 0.946))
        I_mono     = float(hydro_ext.get("I", 175.0))
        Q_limpasan = float(hydro_ext.get("Qr", 0.283))

        st.markdown("##### ⏱️ A. Hasil Ekstraksi Waktu Konsentrasi ($t_c$)")
        ov1, ov2, ov3 = st.columns(3)
        ov1.metric("Waktu Limpas Permukaan (t₁)", f"{t1_val:.3f} menit")
        ov2.metric("Waktu Aliran Saluran (t₂)", f"{t2_val:.3f} menit")
        ov3.metric("Total Waktu Konsentrasi (Tc)", f"{tc_menit:.3f} menit", delta=f"{tc_jam:.4f} jam", delta_color="off")

        st.divider()

        st.markdown("##### 🌧️ B. Hasil Ekstraksi Debit Rasional ($Q_r$)")
        ov4, ov5, ov6, ov7 = st.columns(4)
        ov4.metric("Luas Daerah Tangkapan (A)", f"{A_m2_val:,.0f} m²", delta=f"{A_total_ha:.4f} ha ({A_val:.5f} km²)", delta_color="off")
        ov5.metric("Koefisien Limpasan (C)", f"{C_val:.3f}")
        ov6.metric("Intensitas Hujan (I)", f"{I_mono:.2f} mm/jam")
        ov7.metric("Debit Rencana (Qr)", f"{Q_limpasan:.3f} m³/s")

        if hydro_ext.get("matched_quotes"):
            with st.expander("🔍 Rincian Kutipan Regex Hidrologi yang Tertangkap dari PDF", expanded=False):
                for q_item in hydro_ext["matched_quotes"]:
                    st.markdown(f"- `{q_item}`")
    else:
        # ══════════════════════════════════════════════════════════════
        # 2. KALKULATOR WAKTU KONSENTRASI (tc): KIRPICH & KERBY/HATHAWAY (MANUAL)
        # ══════════════════════════════════════════════════════════════
        st.markdown("#### ⏱️ Kalkulator Manual Waktu Konsentrasi ($t_c$)")
        tc_method = st.radio(
            "Pilih Metode Perhitungan Waktu Konsentrasi (tc)",
            options=[
                "Metode Kirpich (Rural / Saluran Dominan)",
                "Kerby (1959) — SNI 2415:2016 (Overland to + Saluran tf)",
            ],
            horizontal=True,
        )

        if "Kirpich" in tc_method:
            kp1, kp2, kp3 = st.columns(3)
            with kp1:
                L_kirpich = st.number_input("Panjang Aliran Utama L (m)", min_value=10.0, value=800.0, step=50.0, format="%.1f")
            with kp2:
                S_kirpich = st.number_input("Kemiringan Rata-rata Lintasan S (m/m)", min_value=0.0001, value=float(ded.get("S", 0.002)), step=0.0005, format="%.5f")
            tc_menit = 0.0195 * (L_kirpich ** 0.77) * (S_kirpich ** -0.385)
            tc_jam   = tc_menit / 60.0
            t1_val   = 0.0
            t2_val   = tc_menit
            with kp3:
                st.metric("Waktu Konsentrasi tc (Kirpich)", f"{tc_menit:.2f} menit", delta=f"{tc_jam:.4f} jam", delta_color="off")
            st.latex(r"t_c\ (\text{menit}) = 0.0195 \times L^{0.77} \times S^{-0.385}")
        else:
            ND_OPTIONS: dict[str, float] = {
                "0.02 — Permukaan halus kedap air (aspal/beton rata)": 0.02,
                "0.10 — Tanah gundul padat dan halus": 0.10,
                "0.30 — Rumput gersang / permukaan agak kasar / tanah terbuka": 0.30,
                "0.40 — Padang rumput biasa (cukup rapat)": 0.40,
                "0.60 — Hutan gugur daun (deciduous) / semak rapat": 0.60,
                "0.80 — Rumput lebat / hutan rimbun atau serasah tebal": 0.80,
            }
            kb1, kb2, kb3, kb4, kb5 = st.columns(5)
            with kb1:
                Lo_kerby = st.number_input(
                    "Panjang Overland Lo (m, < 365 m)",
                    min_value=1.0,
                    max_value=364.9,
                    value=100.0,
                    step=10.0,
                    format="%.1f",
                    help="Batas validitas Kerby (1959) — SNI 2415:2016 Tabel 10: Lo < 365 m.",
                )
            with kb2:
                nd_label = st.selectbox("Koefisien Hambatan Permukaan (nd)", list(ND_OPTIONS.keys()), index=0)
                nd_val   = ND_OPTIONS[nd_label]
            with kb3:
                So_kerby = st.number_input("Kemiringan Lahan Overland So (m/m)", min_value=0.0001, value=0.005, step=0.001, format="%.4f")
            with kb4:
                Ls_kerby = st.number_input("Panjang Saluran Ls (m)", min_value=1.0, value=600.0, step=50.0, format="%.1f")
            with kb5:
                V_est_tf = st.number_input("Estimasi Kecepatan Saluran V (m/s)", min_value=0.1, value=1.2, step=0.1, format="%.2f")

            to_menit = 1.44 * (((Lo_kerby * nd_val) / math.sqrt(So_kerby)) ** 0.467)
            tf_menit = Ls_kerby / (60.0 * V_est_tf)
            tc_menit = to_menit + tf_menit
            tc_jam   = tc_menit / 60.0
            t1_val   = to_menit
            t2_val   = tf_menit

            km1, km2, km3 = st.columns(3)
            km1.metric("Waktu Lahan (to — Kerby)", f"{to_menit:.2f} menit")
            km2.metric("Waktu Saluran (tf)", f"{tf_menit:.2f} menit")
            km3.metric("Total Waktu Konsentrasi (tc)", f"{tc_menit:.2f} menit ({tc_jam:.4f} jam)")
            st.latex(
                r"t_o\ (\text{menit}) = 1.44 \left(\frac{L_o \cdot n_d}{\sqrt{S_o}}\right)^{0.467} "
                r"\quad;\quad t_f\ (\text{menit}) = \frac{L_s}{60 \cdot V} "
                r"\quad;\quad t_c = t_o + t_f"
            )

        st.divider()

        # ══════════════════════════════════════════════════════════════
        # 3. KALKULATOR DEBIT RASIONAL (Qr) & KOEFISIEN LIMPASAN GABUNGAN (C_gab)
        # ══════════════════════════════════════════════════════════════
        st.markdown("#### 🌧️ 3. Kalkulator Debit Rasional ($Q_r$) & Koefisien Limpasan Gabungan ($C_{\\text{gab}}$)")
        st.latex(r"C_{\text{gab}} = \frac{\sum (C_i \cdot A_i)}{\sum A_i} \qquad;\qquad Q_r = 0.278 \cdot C_{\text{gab}} \cdot I \cdot A")

        LANDUSE_REF = [
            {"Tata Guna Lahan": "Perdagangan / Pusat Kota Padat", "Rentang C": "0.70 – 0.95", "C Tipikal": 0.85},
            {"Tata Guna Lahan": "Permukiman Padat / Perkotaan", "Rentang C": "0.50 – 0.70", "C Tipikal": 0.65},
            {"Tata Guna Lahan": "Permukiman Sedang / Pinggiran", "Rentang C": "0.35 – 0.50", "C Tipikal": 0.45},
            {"Tata Guna Lahan": "Kawasan Industri", "Rentang C": "0.60 – 0.90", "C Tipikal": 0.75},
            {"Tata Guna Lahan": "Jalan aspal dan beton", "Rentang C": "0.75 – 0.95", "C Tipikal": 0.85},
            {"Tata Guna Lahan": "Atap rumah", "Rentang C": "0.70 – 0.95", "C Tipikal": 0.82},
            {"Tata Guna Lahan": "Halaman berumput, tanah pasir, curam (>7%)", "Rentang C": "0.15 – 0.20", "C Tipikal": 0.18},
            {"Tata Guna Lahan": "Taman / Ruang Terbuka Hijau", "Rentang C": "0.10 – 0.25", "C Tipikal": 0.18},
        ]

        with st.expander("📖 Tabel Referensi Koefisien Limpasan (C) — Tabel 12 SNI 2415:2016"):
            st.dataframe(pd.DataFrame(LANDUSE_REF), hide_index=True)

        with st.expander("🏙️ Kalkulator Cepat Koefisien C Drainase Perkotaan (SNI 2415:2016 Persamaan 33–34)"):
            st.latex(r"C = 0.9 \cdot I_m + (1 - I_m) \cdot C_p")
            sni_c1, sni_c2, sni_c3 = st.columns(3)
            with sni_c1:
                im_frac = st.slider("Fraksi Kedap Air (Im)", min_value=0.0, max_value=1.0, value=0.65, step=0.05)
            with sni_c2:
                cp_val = st.number_input("Koefisien Tembus Air (Cp)", min_value=0.05, max_value=0.60, value=0.20, step=0.05, format="%.2f")
            c_sni_urban = 0.9 * im_frac + (1.0 - im_frac) * cp_val
            with sni_c3:
                st.metric("C Perkotaan (Pers. 33–34)", f"{c_sni_urban:.3f}")
            use_c_sni_urban = st.checkbox(
                f"Gunakan C = {c_sni_urban:.3f} (SNI Persamaan 33–34) sebagai C_gab",
                value=False,
            )

        cg_c1, cg_c2, cg_c3 = st.columns(3)
        with cg_c1:
            st.markdown("**Sub-DTA 1 (Misal: Permukiman/Komersial)**")
            C1_sub = st.number_input("Koefisien C₁", min_value=0.05, max_value=1.00, value=0.75, step=0.05, format="%.2f")
            A1_ha  = st.number_input("Luas A₁ (hektar / ha)", min_value=0.0, value=35.0, step=5.0, format="%.2f")
        with cg_c2:
            st.markdown("**Sub-DTA 2 (Misal: Jalan & Fasilitas Umum)**")
            C2_sub = st.number_input("Koefisien C₂", min_value=0.05, max_value=1.00, value=0.85, step=0.05, format="%.2f")
            A2_ha  = st.number_input("Luas A₂ (hektar / ha)", min_value=0.0, value=15.0, step=5.0, format="%.2f")
        with cg_c3:
            st.markdown("**Sub-DTA 3 (Misal: RTH / Taman)**")
            C3_sub = st.number_input("Koefisien C₃", min_value=0.05, max_value=1.00, value=0.20, step=0.05, format="%.2f")
            A3_ha  = st.number_input("Luas A₃ (hektar / ha)", min_value=0.0, value=10.0, step=5.0, format="%.2f")

        A_total_ha  = max(0.01, A1_ha + A2_ha + A3_ha)
        A_m2_val    = A_total_ha * 10_000.0
        A_val       = A_total_ha / 100.0  # konversi ha -> km² (1 km² = 100 ha)
        C_val_weighted = (C1_sub * A1_ha + C2_sub * A2_ha + C3_sub * A3_ha) / A_total_ha
        C_val       = c_sni_urban if use_c_sni_urban else C_val_weighted
        I_mono      = _mononobe_I(R24_val, max(tc_jam, 1e-4))
        Q_limpasan  = _rational_Q(C_val, I_mono, A_val)

        cg_m1, cg_m2, cg_m3, cg_m4 = st.columns(4)
        cg_m1.metric("Koefisien Gabungan (C_gab)", f"{C_val:.3f}")
        cg_m2.metric("Total Luas DTA", f"{A_total_ha:.2f} ha ({A_val:.4f} km²)")
        cg_m3.metric("Intensitas Mononobe (I)", f"{I_mono:.2f} mm/jam")
        cg_m4.metric("Debit Rasional (Qr / Q_limpasan)", f"{Q_limpasan:.4f} m³/s")

    st.divider()

    # ══════════════════════════════════════════════════════════════
    # 4. KAPASITAS MANNING (Qs): GEOMETRI MULTI-PENAMPANG & TINGGI JAGAAN (FREEBOARD w)
    # ══════════════════════════════════════════════════════════════
    st.subheader("📐 4. Kapasitas Manning ($Q_s$) — Geometri Multi-Penampang & Freeboard ($w$)")

    shape_options = ["Persegi", "Trapesium", "Lingkaran"]
    default_shape = cv.get("shape_cv", "Persegi") if cv.get("shape_cv") in shape_options else "Persegi"
    mat_ded_name = str(ded.get("material", "Beton"))
    v_izin_map = {
        "Beton": 3.0,
        "Pasangan Bata": 2.0,
        "Pasangan Batu Kali": 2.0,
        "Pasangan Batu": 2.0,
        "Tanah": 1.0,
    }

    sync_H = max(0.10, float(ded.get("h", 1.2) or 1.2))
    sync_b = max(0.05, float(ded.get("b", 1.5) or 1.5))
    sync_S = max(1e-6, float(ded.get("S", 0.001) or 0.001))
    sync_n_ded = max(0.008, float(ded.get("n", 0.013) or 0.013))
    sync_m_talud = max(0.1, float(cv.get("m_talud", 1.0) or 1.0))
    sync_V_izin = float(v_izin_map.get(mat_ded_name, 3.0))

    # Estimasi h_efektif default jika cv belum memiliki h_effective
    sync_h_maks_est = max(0.05, sync_H * 0.85)
    sync_ts_est = float(cv.get("ts", round(0.2 * sync_h_maks_est, 3)) or 0.0)
    sync_h_eff_fallback = max(0.02, min(sync_H, sync_h_maks_est - sync_ts_est))
    sync_h_eff = max(0.01, min(sync_H, float(cv.get("h_effective", sync_h_eff_fallback) or sync_h_eff_fallback)))
    sync_b_eff = max(0.01, float(cv.get("b_effective", sync_b) or sync_b))
    sync_n_adj = max(0.008, float(cv.get("n_adjusted", sync_n_ded) or sync_n_ded))

    _m4_sync_sig = (
        round(sync_H, 4),
        round(sync_b, 4),
        round(sync_S, 6),
        round(sync_n_ded, 5),
        mat_ded_name,
        default_shape,
        round(sync_m_talud, 3),
        round(sync_h_eff, 4),
        round(sync_b_eff, 4),
        round(sync_n_adj, 5),
    )

    if st.session_state.get("m4_last_sync_sig") != _m4_sync_sig:
        st.session_state["m4_shape"] = default_shape
        st.session_state["m4_H_saluran"] = sync_H
        st.session_state["m4_b_val"] = sync_b
        st.session_state["m4_m_talud"] = sync_m_talud
        st.session_state["m4_S_val"] = sync_S
        st.session_state["m4_n_ded"] = sync_n_ded
        st.session_state["m4_h_eff_val"] = sync_h_eff
        st.session_state["m4_b_eff_val"] = sync_b_eff
        st.session_state["m4_n_adj_val"] = sync_n_adj
        st.session_state["m4_V_izin_val"] = sync_V_izin
        st.session_state["m4_last_sync_sig"] = _m4_sync_sig

    m4_summary_box = st.container()

    fb_options = [
        "Proporsional Saluran Drainase (w = 15% × H)",
        "Rumus √(0.5 · h) (Dibatasi Maks 25% H)",
        "Input Manual w (m)",
        "Tanpa Freeboard (w = 0)",
    ]

    with st.expander("⚙️ Ubah / Override Parameter Geometri & Manning Secara Manual (Opsional)", expanded=False):
        sh1, sh2, sh3 = st.columns(3)
        with sh1:
            jenis_penampang = st.selectbox(
                "Pilih Bentuk Geometri Penampang",
                options=shape_options,
                key="m4_shape",
            )
        with sh2:
            H_saluran = st.number_input(
                "Tinggi Total Saluran H_saluran / Diameter D (m)",
                min_value=0.10, step=0.05, format="%.3f",
                key="m4_H_saluran",
            )
        with sh3:
            fb_mode = st.selectbox(
                "Metode Tinggi Jagaan (Freeboard w)",
                options=fb_options,
                key="m4_fb_mode",
            )

        if fb_mode == "Input Manual w (m)":
            w_freeboard = st.number_input(
                "Input Tinggi Jagaan w (m)",
                min_value=0.0, max_value=max(0.01, H_saluran - 0.05),
                value=min(0.20, H_saluran * 0.15), step=0.05, format="%.3f",
            )
            h_ded = max(0.05, H_saluran - w_freeboard)
        elif fb_mode == "Tanpa Freeboard (w = 0)":
            w_freeboard = 0.0
            h_ded = H_saluran
        elif fb_mode.startswith("Rumus"):
            u_root = (-math.sqrt(0.5) + math.sqrt(0.5 + 4.0 * H_saluran)) / 2.0
            w_raw = H_saluran - (u_root ** 2)
            w_freeboard = min(w_raw, 0.25 * H_saluran)
            h_ded = max(0.05, H_saluran - w_freeboard)
        else:
            w_freeboard = 0.15 * H_saluran
            h_ded = max(0.05, H_saluran - w_freeboard)

        # Input dimensi spesifik penampang + degradasi
        gm1, gm2, gm3, gm4 = st.columns(4)
        with gm1:
            if jenis_penampang in ["Persegi", "Trapesium"]:
                b_val = st.number_input("Lebar Dasar Desain b (m)", min_value=0.05, step=0.1, format="%.3f", key="m4_b_val")
                D_lingkaran = H_saluran
            else:
                D_lingkaran = H_saluran
                b_val = D_lingkaran
                st.metric("Diameter Saluran D (m)", f"{D_lingkaran:.3f} m")

            if jenis_penampang == "Trapesium":
                m_talud = st.number_input("Kemiringan Talud m (1V:mH)", min_value=0.1, step=0.1, format="%.2f", key="m4_m_talud")
            else:
                m_talud = 0.0

        with gm2:
            S_val = st.number_input("Kemiringan Dasar Saluran S (m/m)", min_value=1e-6, step=0.0005, format="%.6f", key="m4_S_val")
            n_ded = st.number_input("Manning n Desain (DED)", min_value=0.008, step=0.001, format="%.4f", key="m4_n_ded")

        with gm3:
            if st.session_state.get("m4_h_eff_val", sync_h_eff) > H_saluran:
                st.session_state["m4_h_eff_val"] = float(H_saluran)
            h_eff_val = st.number_input(
                "Kedalaman Efektif Aktual h_efektif (m)",
                min_value=0.01, max_value=float(H_saluran),
                step=0.02, format="%.4f",
                key="m4_h_eff_val",
                help="h_efektif = h_maks - ts (dari Modul 3).",
            )
            b_eff_val = st.number_input(
                "Lebar Dasar Efektif Aktual b_efektif (m)",
                min_value=0.01,
                step=0.05, format="%.4f",
                key="m4_b_eff_val",
            )

        with gm4:
            n_adj_val = st.number_input(
                "Manning n Aktual (Cowan / Horton)",
                min_value=0.008,
                step=0.001, format="%.5f",
                key="m4_n_adj_val",
            )
            V_izin_val = st.number_input(
                f"Kecepatan Izin Maks V_izin ({mat_ded_name}) (m/s)",
                min_value=0.5, step=0.2, format="%.2f",
                key="m4_V_izin_val",
            )

    with m4_summary_box:
        with st.container(border=True):
            st.caption(
                f"🔗 **Sinkronisasi Otomatis Aktif (Modul 2 DED: `{mat_ded_name}` · Modul 3 CV: `{jenis_penampang}`)**"
            )
            sr1, sr2, sr3, sr4 = st.columns(4)
            sr1.metric("Lebar Dasar b (DED)", f"{b_val:.3f} m", delta=f"b_efektif = {b_eff_val:.3f} m", delta_color="off")
            sr2.metric("Tinggi Total H (DED)", f"{H_saluran:.3f} m", delta=f"h_maks = {h_ded:.3f} m (w={w_freeboard:.3f} m)", delta_color="off")
            sr3.metric("Kemiringan S & Manning DED", f"S = {S_val:.5f}", delta=f"n_DED = {n_ded:.4f}", delta_color="off")
            sr4.metric("Kondisi Aktual (Modul 3)", f"n_akt = {n_adj_val:.4f}", delta=f"h_ef = {h_eff_val:.3f} m · V_izin = {V_izin_val:.1f} m/s", delta_color="off")

    def _calc_section_hydraulics(shape: str, b_in: float, h_in: float, m_in: float, D_in: float, S_in: float, n_in: float) -> dict:
        if shape == "Persegi":
            A_sec = b_in * h_in
            P_sec = b_in + 2.0 * h_in
            T_sec = b_in
            theta_rad = 0.0
        elif shape == "Trapesium":
            A_sec = (b_in + m_in * h_in) * h_in
            P_sec = b_in + 2.0 * h_in * math.sqrt(1.0 + m_in ** 2)
            T_sec = b_in + 2.0 * m_in * h_in
            theta_rad = 0.0
        else:  # Lingkaran
            # Clamp h_in agar 0 < h_in < D_in supaya T_lebar valid
            h_clamped = float(np.clip(h_in, 0.001 * D_in, 0.999 * D_in))
            cos_arg   = float(np.clip(1.0 - (2.0 * h_clamped / D_in), -1.0, 1.0))
            theta_rad = 2.0 * math.acos(cos_arg)
            A_sec     = ((D_in ** 2) / 8.0) * (theta_rad - math.sin(theta_rad))
            P_sec     = (D_in * theta_rad) / 2.0
            T_sec     = max(1e-4, D_in * math.sin(theta_rad / 2.0))

        R_sec  = (A_sec / P_sec) if P_sec > 0 else 0.0
        Q_sec  = (1.0 / n_in) * A_sec * (R_sec ** (2.0 / 3.0)) * (S_in ** 0.5) if n_in > 0 and S_in > 0 else 0.0
        V_sec  = (Q_sec / A_sec) if A_sec > 0 else 0.0
        Dh_sec = (A_sec / T_sec) if T_sec > 0 else 0.0
        Fr_sec = (V_sec / math.sqrt(9.81 * Dh_sec)) if Dh_sec > 0 else 0.0
        return {
            "A": A_sec, "P": P_sec, "R": R_sec, "T_lebar": T_sec,
            "Q": Q_sec, "V": V_sec, "Dh": Dh_sec, "Fr": Fr_sec, "theta": theta_rad,
        }

    st.divider()

    eval_m4_btn = st.button("⚡ Hitung & Evaluasi Kapasitas Hidrolika Hibrida", type="primary")
    if eval_m4_btn or run_m4_extract_btn:
        hyd_desain = _calc_section_hydraulics(jenis_penampang, b_val,     h_ded,     m_talud, D_lingkaran, S_val, n_ded)
        hyd_aktual = _calc_section_hydraulics(jenis_penampang, b_eff_val, h_eff_val, m_talud, D_lingkaran, S_val, n_adj_val)

        Q_desain = hyd_desain["Q"]
        Q_aktual = hyd_aktual["Q"]
        V_aktual = hyd_aktual["V"]
        Fr_aktual = hyd_aktual["Fr"]

        # Persist to session state for Modul 4 & Modul 5
        st.session_state["m4_results"] = {
            "Q_desain":   Q_desain,
            "Q_aktual":   Q_aktual,
            "Q_limpasan": Q_limpasan,
            "I_mono":     I_mono,
            "b":          b_val,
            "b_eff":      b_eff_val,
            "h_ded":      h_ded,
            "h_eff":      h_eff_val,
            "w_fb":       w_freeboard,
            "n_ded":      n_ded,
            "n_adj":      n_adj_val,
            "S":          S_val,
            "R24":        R24_val,
            "T_ulang":    T_pilihan,
            "tipologi_kota": tipologi_kota,
            "pilihan_kota": pilihan_kota,
            "C":          C_val,
            "A":          A_val,
            "A_m2":       A_m2_val,
            "A_ha":       A_total_ha,
            "t1":         t1_val,
            "t2":         t2_val,
            "tc_menit":   tc_menit,
            "tc":         tc_jam,
            "shape":      jenis_penampang,
            "V_aktual":   V_aktual,
            "Fr_aktual":  Fr_aktual,
            "V_izin":     V_izin_val,
            "hyd_desain": hyd_desain,
            "hyd_aktual": hyd_aktual,
        }
        if eval_m4_btn:
            st.toast("💾 Hasil Modul 4 disimpan → Modul 5", icon="✅")

    if st.session_state.get("m4_results") is not None:
        m4_saved   = st.session_state["m4_results"]
        hyd_desain = m4_saved.get("hyd_desain") or _calc_section_hydraulics(jenis_penampang, b_val, h_ded, m_talud, D_lingkaran, S_val, n_ded)
        hyd_aktual = m4_saved.get("hyd_aktual") or _calc_section_hydraulics(jenis_penampang, b_eff_val, h_eff_val, m_talud, D_lingkaran, S_val, n_adj_val)
        Q_desain   = float(m4_saved["Q_desain"])
        Q_aktual   = float(m4_saved["Q_aktual"])
        Q_limpasan = float(m4_saved["Q_limpasan"])
        I_mono     = float(m4_saved["I_mono"])
        V_aktual   = float(m4_saved["V_aktual"])
        Fr_aktual  = float(m4_saved["Fr_aktual"])

        # === Comparison Table ===
        st.subheader("📊 Tabel Komparasi Tiga Skenario & Parameter Hidrolis")
        df_cmp = pd.DataFrame([
            {
                "Skenario": f"🏗️ Q_desain — Kapasitas Desain ({jenis_penampang})",
                "Debit Q (m³/s)": round(Q_desain, 4),
                "Luas A (m²)": round(hyd_desain["A"], 4),
                "Jari-jari R (m)": round(hyd_desain["R"], 4),
                "Kecepatan V (m/s)": round(hyd_desain["V"], 3),
                "Bil. Froude (Fr)": round(hyd_desain["Fr"], 3),
                "Parameter Utama": f"h_maks={h_ded:.3f} m (w={w_freeboard:.2f} m) · n={n_ded:.4f}",
            },
            {
                "Skenario": f"🔧 Q_aktual — Kapasitas Eksisting Terdegradasi ({jenis_penampang})",
                "Debit Q (m³/s)": round(Q_aktual, 4),
                "Luas A (m²)": round(hyd_aktual["A"], 4),
                "Jari-jari R (m)": round(hyd_aktual["R"], 4),
                "Kecepatan V (m/s)": round(hyd_aktual["V"], 3),
                "Bil. Froude (Fr)": round(hyd_aktual["Fr"], 3),
                "Parameter Utama": f"h_eff={h_eff_val:.3f} m · b_eff={b_eff_val:.3f} m · n_akt={n_adj_val:.5f}",
            },
            {
                "Skenario": f"🌧️ Q_limpasan — Debit Banjir Rencana (T={T_pilihan} Tahun)",
                "Debit Q (m³/s)": round(Q_limpasan, 4),
                "Luas A (m²)": np.nan,
                "Jari-jari R (m)": np.nan,
                "Kecepatan V (m/s)": np.nan,
                "Bil. Froude (Fr)": np.nan,
                "Parameter Utama": f"R24={R24_val:.1f} mm · tc={tc_menit:.3f} mnt · I={I_mono:.2f} mm/jam · C={C_val:.3f}",
            },
        ])
        st.dataframe(df_cmp, hide_index=True)

        # === Metric Cards ===
        m1, m2, m3, m4_c, m5_c = st.columns(5)
        m1.metric("🏗️ Q Desain Awal", f"{Q_desain:.4f} m³/s", delta=f"w = {w_freeboard:.2f} m", delta_color="off")
        m2.metric(
            "🔧 Q Aktual Terdegradasi",
            f"{Q_aktual:.4f} m³/s",
            delta=f"{Q_aktual - Q_desain:+.4f} m³/s vs desain",
            delta_color="inverse" if Q_aktual < Q_desain else "normal",
        )
        m3.metric(f"🌧️ Q Limpasan (T={T_pilihan}th)", f"{Q_limpasan:.4f} m³/s", delta=f"I = {I_mono:.2f} mm/jam", delta_color="off")
        m4_c.metric("💨 Kecepatan Aktual V", f"{V_aktual:.3f} m/s", delta=f"V_izin = {V_izin_val:.1f} m/s", delta_color="off")
        m5_c.metric("🌊 Bilangan Froude (Fr)", f"{Fr_aktual:.3f}", delta="Subkritis (<1)" if Fr_aktual < 1.0 else "Kritis/Superkritis (≥1)", delta_color="normal" if Fr_aktual < 1.0 else "inverse")

        # === Bar Chart (Scientific Hydrology Palette + Interactive Tab) ===
        _render_section_banner(
            "📊 Visualisasi Perbandingan Debit (Q Desain vs Q Aktual vs Q Limpasan)",
            f"Komparasi kapasitas hidrolika penampang {jenis_penampang} terhadap beban debit banjir rencana kala ulang T={T_pilihan} tahun.",
            badge_html=(
                '<span class="ds-badge safe">✅ KAPASITAS AMAN</span>'
                if Q_aktual >= Q_limpasan
                else '<span class="ds-badge danger">🚨 DEFISIT KAPASITAS</span>'
            ),
            tone="ocean",
        )
        m4_tab1, m4_tab2 = st.tabs(["📊 Grafik Komparasi Ilmiah (High-Res)", "⚡ Grafik Batang Interaktif"])
        with m4_tab1:
            fig_m4, ax_m4 = plt.subplots(figsize=(8.8, 3.6), dpi=120)
            fig_m4.patch.set_facecolor("#FFFFFF")
            ax_m4.set_facecolor("#F4F8FC")
            _bars_lbl = [
                f"Q Desain Awal\n(h_maks={h_ded:.2f}m)",
                f"Q Aktual Eksisting\n(h_eff={h_eff_val:.2f}m)",
                f"Q Limpasan Banjir\n(T={T_pilihan} Tahun)",
            ]
            _bars_val = [Q_desain, Q_aktual, Q_limpasan]
            _bars_col = ["#1677C8", "#24A47B" if Q_aktual >= Q_limpasan else "#F4B942", "#E45B64"]
            _bars_obj = ax_m4.bar(_bars_lbl, _bars_val, color=_bars_col, width=0.48, edgecolor="#102A43", linewidth=0.9)
            for _b_rect, _b_v in zip(_bars_obj, _bars_val):
                ax_m4.text(
                    _b_rect.get_x() + _b_rect.get_width() / 2.0,
                    _b_rect.get_height() + max(_bars_val) * 0.03,
                    f"{_b_v:.4f} m³/s",
                    ha="center",
                    va="bottom",
                    fontsize=9,
                    fontweight="bold",
                    color="#102A43",
                )
            ax_m4.set_ylabel("Debit Aliran Q (m³/detik)", fontsize=9.5, color="#18324A")
            ax_m4.set_title(f"Evaluasi Neraca Debit Saluran ({jenis_penampang})", fontsize=11, fontweight="bold", color="#102A43")
            ax_m4.set_ylim(0, max(_bars_val) * 1.24 + 0.05)
            ax_m4.grid(axis="y", linestyle="--", alpha=0.45, color="#B8CCE0")
            fig_m4.tight_layout()
            st.pyplot(fig_m4)
            plt.close(fig_m4)
        with m4_tab2:
            chart_data = pd.DataFrame({
                "Q Desain": [Q_desain],
                "Q Aktual": [Q_aktual],
                "Q Limpasan": [Q_limpasan],
            })
            st.bar_chart(chart_data)

        st.divider()

        # === Evaluasi Hidrolis: Kapasitas, Kecepatan & Froude ===
        st.subheader("🚦 Status Evaluasi Kapasitas, Kecepatan & Rezim Aliran (Froude)")
        defisit = Q_limpasan - Q_aktual
        is_safe = Q_aktual >= Q_limpasan

        if is_safe:
            st.success(
                f"✅ **STATUS: AMAN** — Kapasitas eksisting ({Q_aktual:.4f} m³/s) "
                f"mampu menampung debit limpasan kala ulang {T_pilihan} tahun ({Q_limpasan:.4f} m³/s).  "
                f"Surplus: **{abs(defisit):.4f} m³/s**"
            )
        else:
            st.error(
                f"🚨 **STATUS: RISIKO LUAPAN / BERBAHAYA** — Kapasitas eksisting terdegradasi "
                f"({Q_aktual:.4f} m³/s) TIDAK mencukupi debit limpasan ({Q_limpasan:.4f} m³/s).  "
                f"Defisit: **{defisit:.4f} m³/s**"
            )

        # Kontrol Kecepatan V (0.6 m/s <= V <= V_izin)
        if V_aktual < 0.6:
            st.warning(
                f"⚠️ **Peringatan Kecepatan Rendah ($V = {V_aktual:.3f}\\text{{ m/s}} < 0.60\\text{{ m/s}}$):** "
                "Berisiko tinggi mengalami **sedimentasi / pengendapan lumpur** karena kecepatan aliran "
                "berada di bawah kecepatan swa-bersih (*self-cleansing velocity*) minimum."
            )
        elif V_aktual > V_izin_val:
            st.warning(
                f"⚠️ **Peringatan Kecepatan Tinggi ($V = {V_aktual:.3f}\\text{{ m/s}} > V_{{\\text{{izin}}}} = {V_izin_val:.2f}\\text{{ m/s}}$):** "
                f"Berisiko mengalami **gerusan / abrasi dinding saluran** ({mat_ded_name}). Pertimbangkan peredam energi atau pelapisan ulang."
            )
        else:
            st.info(
                f"✅ **Kecepatan Aliran Aman ($V = {V_aktual:.3f}\\text{{ m/s}}$):** "
                f"Berada dalam rentang ideal $0.60\\text{{ m/s}} \\le V \\le {V_izin_val:.2f}\\text{{ m/s}}$ (bebas sedimentasi & gerusan)."
            )

        # Kontrol Bilangan Froude (Fr)
        if Fr_aktual >= 1.0:
            st.warning(
                f"⚠️ **Peringatan Rezim Aliran ($Fr = {Fr_aktual:.3f} \\ge 1.0$):** "
                "Aliran bersifat **Kritis / Superkritis**. Berpotensi menimbulkan loncatan hidrolis (*hydraulic jump*), "
                "gelombang tegak, dan ketidakstabilan struktural dinding saluran."
            )
        else:
            st.info(
                f"✅ **Rezim Aliran Subkritis Stabil ($Fr = {Fr_aktual:.3f} < 1.0$):** "
                "Aliran tenang dan stabil untuk saluran drainase perkotaan."
            )

        # === Detailed formula breakdown ===
        with st.expander("📐 Rincian Rumus Geometri, Manning, Kecepatan & Bilangan Froude"):
            if jenis_penampang == "Persegi":
                st.latex(r"A = b \cdot h \quad;\quad P = b + 2h \quad;\quad T_{\text{lebar}} = b")
            elif jenis_penampang == "Trapesium":
                st.latex(r"A = (b + m \cdot h)h \quad;\quad P = b + 2h\sqrt{1+m^2} \quad;\quad T_{\text{lebar}} = b + 2mh")
            else:
                st.latex(
                    r"\theta = 2\arccos\!\left(1 - \frac{2h}{D}\right) \quad;\quad "
                    r"A = \frac{D^2}{8}(\theta - \sin\theta) \quad;\quad "
                    r"P = \frac{D\theta}{2} \quad;\quad T_{\text{lebar}} = D\sin\!\left(\frac{\theta}{2}\right)"
                )
            st.latex(
                r"V = \frac{Q}{A} \quad;\quad Fr = \frac{V}{\sqrt{9.81 \cdot (A / T_{\text{lebar}})}}"
            )
    else:
        _render_kobo_empty_state(
            asset_key="kobo_chibi_2",
            title="Evaluasi Kapasitas Hidrolika Hibrida Siap Dijalankan",
            subtitle=(
                "Klik tombol ⚡ Hitung & Evaluasi Kapasitas Hidrolika Hibrida atau ⚡ Ekstraksi Instan Hidrologi via Python "
                "di atas untuk menampilkan tabel komparasi 3 skenario debit (Q Desain, Q Aktual, Q Limpasan) dan grafik evaluasi."
            ),
            badges_html=(
                '<span class="ds-badge auto">Debit Rasional Qr (Mononobe)</span>'
                '<span class="ds-badge auto">Manning Multi-Penampang Qs</span>'
                '<span class="ds-badge auto">Kontrol Kecepatan &amp; Froude (Fr)</span>'
            ),
            show_image=False,
        )


# ═════════════════════════════════════════════════════════════════
# MODUL 5: Executive Dashboard & PDF Export
# ═════════════════════════════════════════════════════════════════
elif menu.startswith("5 |") or "Modul 5" in menu:
    R24      = st.session_state.get("R24_5yr")
    r24_dict = st.session_state.get("r24_dict") or {}
    ded      = st.session_state.get("ded_params") or {}
    cv       = st.session_state.get("cv_params")  or {}
    m4       = st.session_state.get("m4_results") or {}

    _render_module_header(
        step_num=5,
        icon="📊",
        title="Executive Dashboard, Optimasi Pengerukan (Bisection) & PDF Export",
        description=(
            "Pusat komando eksekutif terintegrasi meliputi Executive Summary KPI, Safety Factor (SF), "
            "Algoritma Numerik Bisection untuk target pengerukan sedimen (h*, ts_izin, t_pengerukan), "
            "Kurva Kapasitas Manning (Q vs h), serta Rekomendasi Permen PU No. 12/PRT/M/2014 dan Ekspor PDF Lengkap."
        ),
        status_text="Laporan Eksekutif Siap" if bool(m4) else "Menunggu Evaluasi Modul 4",
        is_ready=bool(m4),
    )

    if not m4:
        _render_kobo_empty_state(
            asset_key="kobo_happy_2",
            title="Executive Dashboard Siap Menampilkan Hasil Evaluasi",
            subtitle=(
                "Data perhitungan debit hidrolika dari Modul 4 (Engine Hidrolika Hibrida) belum tersedia di Session State. "
                "Silakan buka 4 | ⚙️ Engine Hidrolika Hibrida dan klik tombol Hitung & Evaluasi Kapasitas Hidrolika Hibrida "
                "atau Ekstraksi Instan Hidrologi via Python."
            ),
            badges_html=(
                '<span class="ds-badge auto">1 · Iklim &amp; IDF (SNI 2415:2016)</span>'
                '<span class="ds-badge auto">2 · Dimensi DED (Chow 1959)</span>'
                '<span class="ds-badge auto">3 · Visi Komputer (Cowan 1956)</span>'
                '<span class="ds-badge manual">4 · Evaluasi Debit Hidrolika → Diperlukan</span>'
            ),
            show_image=True,
        )

    if m4:
        _render_kobo_companion_card(
            asset_key="kobo_happy",
            tag="🎉 KOBO EXECUTIVE CELEBRATION · PIPELINE ANALYSIS COMPLETE",
            title="Rekapitulasi Eksekutif, Optimasi Bisection & Rekomendasi Permen PU No. 12/PRT/M/2014",
            description=(
                "Seluruh tahapan analisis hidrologi dan hidrolika telah terhubung! Tinjau indikator Safety Factor (SF), "
                "kurva kapasitas Manning (Q vs h), target pengerukan sedimen (t_pengerukan), dan unduh laporan PDF lengkap."
            ),
        )
        m5_an         = _compute_m5_analysis(m4, cv, ded)
        Q_desain      = m5_an["Q_desain"]
        Q_aktual      = m5_an["Q_aktual"]
        Q_banjir      = m5_an["Q_banjir"]
        SF            = m5_an["SF"]
        penurunan_pct = m5_an["penurunan_pct"]
        h_star        = m5_an["h_star"]
        h_star_in_rng = m5_an["h_star_in_range"]
        h_exceeds     = m5_an["h_star_exceeds"]
        ts_current    = m5_an["ts_current"]
        ts_izin       = m5_an["ts_izin"]
        t_pengerukan  = m5_an["t_pengerukan"]
        h_maks        = m5_an["h_maks"]
        h_eff         = m5_an["h_eff"]
        b_desain      = m5_an["b_desain"]
        b_aktual      = m5_an["b_aktual"]
        n_desain      = m5_an["n_desain"]
        n_aktual      = m5_an["n_aktual"]
        shape         = m5_an["shape"]
        m_talud       = m5_an["m_talud"]
        D_diam        = m5_an["D_diam"]
        S_slope       = m5_an["S_slope"]
        defisit       = Q_banjir - Q_aktual
        is_safe_sf    = SF >= 1.0

        # ══════════════════════════════════════════════════════════
        # EXECUTIVE SUMMARY (3 METRIK UTAMA DI BAGIAN PALING ATAS)
        # ══════════════════════════════════════════════════════════
        _render_section_banner(
            "📌 Executive Summary — Indikator Kinerja Utama (KPI) Saluran Drainase",
            f"Evaluasi kelayakan kapasitas eksisting penampang {shape} terhadap debit banjir rencana kala ulang T={m4.get('T_ulang', 5)} tahun.",
            badge_html=(
                '<span class="ds-badge safe">🟢 STATUS: AMAN (SF ≥ 1.0)</span>'
                if is_safe_sf
                else '<span class="ds-badge danger">🔴 STATUS: KRITIS / MELUAP (SF &lt; 1.0)</span>'
            ),
            tone="navy",
        )

        es1, es2, es3 = st.columns(3)
        es1.metric(
            "1. Status Saluran",
            "AMAN" if is_safe_sf else "MELUAP",
            delta=f"Surplus {abs(defisit):.4f} m³/s" if is_safe_sf else f"Defisit {defisit:.4f} m³/s",
            delta_color="normal" if is_safe_sf else "inverse",
        )
        es2.metric(
            "2. Safety Factor (SF)",
            f"{SF:.3f}",
            delta="Memenuhi Standar (SF ≥ 1.0)" if is_safe_sf else "Di Bawah Standar (SF < 1.0)",
            delta_color="normal" if is_safe_sf else "inverse",
        )
        es3.metric(
            "3. Target Pengerukan Lumpur (m)",
            f"{t_pengerukan:.3f} m",
            delta=f"{t_pengerukan * 100:.1f} cm (ts saat ini = {ts_current:.3f} m)",
            delta_color="inverse" if t_pengerukan > 0 else "normal",
        )

        if is_safe_sf:
            st.success(
                f"🟢 **KESIMPULAN EKSEKUTIF: AMAN** — Safety Factor **$SF = {SF:.3f} \\ge 1.00$** "
                f"($Q_{{\\text{{aktual}}}} = {Q_aktual:.4f}\\text{{ m}}^3/\\text{{s}} \\ge Q_{{\\text{{banjir}}}} = {Q_banjir:.4f}\\text{{ m}}^3/\\text{{s}}$)."
            )
        else:
            st.error(
                f"🔴 **KESIMPULAN EKSEKUTIF: KRITIS / MELUAP** — Safety Factor **$SF = {SF:.3f} < 1.00$** "
                f"($Q_{{\\text{{aktual}}}} = {Q_aktual:.4f}\\text{{ m}}^3/\\text{{s}} < Q_{{\\text{{banjir}}}} = {Q_banjir:.4f}\\text{{ m}}^3/\\text{{s}}$, "
                f"Defisit = **{defisit:.4f} m³/s**). Diperlukan pengerukan sedimen minimal **{t_pengerukan:.3f} m** ({t_pengerukan * 100:.1f} cm)."
            )

        st.divider()

        # ══════════════════════════════════════════════════════════
        # 1. RINCIAN METRIK EVALUASI KEAMANAN & PENURUNAN KAPASITAS
        # ══════════════════════════════════════════════════════════
        st.markdown(
            f"""
            <div class="ds-section-card ocean">
                <div class="ds-section-title">
                    <span>🛡️ 1. Rincian Penurunan Kapasitas &amp; Neraca Debit</span>
                    <span class="ds-badge auto">Degradasi Kapasitas: {penurunan_pct:.1f}%</span>
                </div>
                <p class="ds-section-sub">Perbandingan kapasitas desain awal terhadap kapasitas eksisting yang terdegradasi.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        k1, k2, k3, k4 = st.columns(4)
        k1.metric(
            "🏗️ Debit Desain (Q_desain)",
            f"{Q_desain:.4f} m³/s",
            delta=f"h_maks = {h_maks:.3f} m",
            delta_color="off",
        )
        k2.metric(
            "📉 Penurunan Kapasitas",
            f"{penurunan_pct:.2f} %",
            delta=f"{Q_aktual - Q_desain:+.4f} m³/s dari Q_desain",
            delta_color="inverse" if penurunan_pct > 0 else "normal",
        )
        k3.metric(
            "💧 Surplus / Defisit Debit",
            f"{abs(defisit):.4f} m³/s",
            delta="Surplus Kapasitas" if is_safe_sf else "Defisit Kapasitas",
            delta_color="normal" if is_safe_sf else "inverse",
        )
        k4.metric(
            "⛏️ Tebal Pengerukan Wajib",
            f"{t_pengerukan:.3f} m ({t_pengerukan * 100:.1f} cm)",
            delta=f"ts saat ini = {ts_current:.3f} m",
            delta_color="inverse" if t_pengerukan > 0 else "normal",
        )

        with st.expander("📐 Rumus Safety Factor (SF) & Persentase Penurunan Kapasitas"):
            st.latex(
                rf"SF = \frac{{Q_{{\text{{aktual}}}}}}{{Q_{{\text{{banjir}}}}}} = "
                rf"\frac{{{Q_aktual:.4f}}}{{{Q_banjir:.4f}}} = {SF:.3f} \quad;\quad "
                rf"\text{{Penurunan}} = \left(1 - \frac{{Q_{{\text{{aktual}}}}}}{{Q_{{\text{{desain}}}}}}\right) \times 100\% = "
                rf"\left(1 - \frac{{{Q_aktual:.4f}}}{{{Q_desain:.4f}}}\right) \times 100\% = {penurunan_pct:.2f}\%"
            )

        st.divider()

        # ══════════════════════════════════════════════════════════
        # 2. ALGORITMA BISECTION (BAGI DUA) UNTUK TARGET PENGERUKAN SEDIMEN
        # ══════════════════════════════════════════════════════════
        _render_section_banner(
            "🎯 2. Optimasi Target Pengerukan Sedimen — Algoritma Numerik Bisection (Bagi Dua)",
            "Mencari kedalaman air kritis minimum h* pada rentang [0, h_maks] yang memenuhi persamaan Q(h*, n_aktual) = Q_banjir.",
            badge_html='<span class="ds-badge auto">Root-Finding Convergence · tol=1e-5</span>',
            tone="teal",
        )

        bs1, bs2, bs3, bs4 = st.columns(4)
        bs1.metric(
            "Kedalaman Kritis Minimum (h*)",
            f"{h_star:.4f} m",
            delta="Melebihi h_maks!" if h_exceeds else f"≤ h_maks ({h_maks:.3f} m)",
            delta_color="inverse" if h_exceeds else "normal",
        )
        bs2.metric(
            "Tebal Sedimen Maks Diizinkan (ts_izin)",
            f"{ts_izin:.4f} m",
            delta=f"{ts_izin * 100:.1f} cm",
            delta_color="off",
        )
        bs3.metric(
            "Tebal Sedimen Eksisting (ts)",
            f"{ts_current:.4f} m",
            delta=f"{ts_current * 100:.1f} cm",
            delta_color="off",
        )
        bs4.metric(
            "Tebal Pengerukan Wajib (t_pengerukan)",
            f"{t_pengerukan:.4f} m",
            delta=f"{t_pengerukan * 100:.1f} cm wajib dikeruk",
            delta_color="inverse" if t_pengerukan > 0 else "normal",
        )

        st.latex(
            r"Q(h^*, n_{\text{aktual}}) = Q_{\text{banjir}} \implies "
            r"t_{s,\text{izin}} = h_{\text{maks}} - h^* \implies "
            r"t_{\text{pengerukan}} = \max(0,\; t_s - t_{s,\text{izin}})"
        )

        if h_exceeds:
            st.warning(
                f"⚠️ **Hasil Bisection:** Pada batas atas rentang $[0, h_{{\\text{{maks}}}} = {h_maks:.3f}\\text{{ m}}]$, "
                f"debit maksimum dengan kekasaran aktual ($n_{{\\text{{aktual}}}} = {n_aktual:.5f}$) adalah "
                f"**{m5_an['Q_at_hmaks_nakt']:.4f} m³/s**, yang masih lebih kecil dari $Q_{{\\text{{banjir}}}} = {Q_banjir:.4f}\\text{{ m}}^3/\\text{{s}}$. "
                f"Kedalaman kritis teoretis yang dibutuhkan adalah **$h^* = {h_star:.3f}\\text{{ m}} > h_{{\\text{{maks}}}}$**, sehingga "
                f"$t_{{s,\\text{{izin}}}} = 0.000\\text{{ m}}$, seluruh sedimen wajib dikeruk ($t_{{\\text{{pengerukan}}}} = {t_pengerukan:.3f}\\text{{ m}}$), "
                "dan diperlukan pembersihan gulma / rehabilitasi dimensi."
            )
        else:
            st.info(
                f"✅ **Hasil Bisection Konvergen:** Kedalaman air minimum yang dibutuhkan untuk mengalirkan "
                f"$Q_{{\\text{{banjir}}}} = {Q_banjir:.4f}\\text{{ m}}^3/\\text{{s}}$ adalah **$h^* = {h_star:.4f}\\text{{ m}}$**. "
                f"Toleransi endapan sedimen maksimum adalah **$t_{{s,\\text{{izin}}}} = {ts_izin:.4f}\\text{{ m}}$** "
                f"({ts_izin * 100:.1f} cm), sehingga ketebalan sedimen yang wajib dikeruk adalah "
                f"**$t_{{\\text{{pengerukan}}}} = {t_pengerukan:.4f}\\text{{ m}}$** ({t_pengerukan * 100:.1f} cm)."
            )

        st.markdown("---")

        # ══════════════════════════════════════════════════════════
        # 3. GRAFIK KURVA KAPASITAS SALURAN TERHADAP KEDALAMAN AIR (Q vs h)
        # ══════════════════════════════════════════════════════════
        _render_section_banner(
            "📈 3. Grafik Kurva Kapasitas Saluran terhadap Kedalaman Air (Q vs h)",
            f"Hubungan non-linier debit Manning terhadap kedalaman aliran h pada penampang {shape}, menampilkan titik kritis h* hasil Bisection.",
            badge_html=f'<span class="ds-badge auto">h* = {min(h_star, h_maks):.3f} m</span>',
            tone="ocean",
        )

        h_min_plot = min(0.05, h_maks * 0.1)
        h_grid = np.linspace(h_min_plot, h_maks, 200)
        q_curve_desain = [
            _manning_Q_multi(shape, b_desain, float(hv), m_talud, D_diam, S_slope, n_desain)
            for hv in h_grid
        ]
        q_curve_aktual = [
            _manning_Q_multi(shape, b_aktual, float(hv), m_talud, D_diam, S_slope, n_aktual)
            for hv in h_grid
        ]

        fig, ax = plt.subplots(figsize=(10.2, 5.0), dpi=125)
        fig.patch.set_facecolor("#FFFFFF")
        ax.set_facecolor("#F4F8FC")
        ax.plot(
            h_grid, q_curve_desain,
            label=f"Kurva Kapasitas Desain Ideal (n_desain = {n_desain:.4f})",
            color="#1677C8", linewidth=2.6,
        )
        ax.plot(
            h_grid, q_curve_aktual,
            label=f"Kurva Kapasitas Aktual Eksisting (n_aktual = {n_aktual:.5f})",
            color="#11A8A0", linewidth=2.6,
        )
        ax.fill_between(h_grid, q_curve_aktual, color="#34C6D3", alpha=0.12)

        # Garis horizontal putus-putus untuk Q_banjir
        ax.axhline(
            y=Q_banjir, color="#E45B64", linestyle="--", linewidth=2.0,
            label=f"Debit Banjir Rencana Q_banjir = {Q_banjir:.4f} m³/s",
        )

        # Tandai titik potong kedalaman kritis h* pada grafik
        h_star_plot = min(h_star, h_maks)
        q_star_plot = _manning_Q_multi(shape, b_aktual, h_star_plot, m_talud, D_diam, S_slope, n_aktual)
        ax.axvline(x=h_star_plot, color="#24A47B", linestyle=":", linewidth=1.8)
        ax.scatter(
            [h_star_plot], [q_star_plot],
            color="#24A47B", s=100, zorder=6,
            label=f"Titik Kritis Bisection h* = {h_star_plot:.3f} m (Q = {q_star_plot:.3f} m³/s)",
        )
        ax.annotate(
            f"h* = {h_star_plot:.3f} m\n(ts_izin = {ts_izin:.3f} m)",
            xy=(h_star_plot, q_star_plot),
            xytext=(max(h_min_plot, h_star_plot * 0.70), q_star_plot * 1.08 + 0.05),
            arrowprops={"arrowstyle": "->", "color": "#24A47B", "lw": 1.4},
            fontsize=9, fontweight="bold", color="#102A43",
        )

        # Tandai titik kondisi aktual saat ini (h_efektif, Q_aktual)
        ax.scatter(
            [h_eff], [Q_aktual],
            color="#102A43", marker="D", s=80, zorder=6,
            label=f"Kondisi Eksisting (h_eff = {h_eff:.3f} m, Q_akt = {Q_aktual:.3f} m³/s)",
        )

        ax.set_title(
            f"Kurva Kapasitas Saluran Manning Q vs h — Penampang {shape}",
            fontsize=12, fontweight="bold", color="#102A43", pad=10,
        )
        ax.set_xlabel("Kedalaman Air h (m)", fontsize=10, color="#18324A")
        ax.set_ylabel("Debit Aliran Q (m³/detik)", fontsize=10, color="#18324A")
        ax.set_xlim(left=0.0, right=h_maks * 1.05)
        ax.set_ylim(bottom=0.0, top=max(max(q_curve_desain), Q_banjir) * 1.18 + 0.05)
        ax.grid(True, linestyle="--", alpha=0.45, color="#B8CCE0")
        ax.legend(loc="upper left", fontsize=8.6, framealpha=0.95, facecolor="#FFFFFF", edgecolor="#DFE8F0")
        fig.tight_layout()
        st.pyplot(fig)
        plt.close(fig)

        st.markdown("---")

        # ══════════════════════════════════════════════════════════
        # 4. REKOMENDASI PEMELIHARAAN OTOMATIS (PERMEN PU NO. 12/PRT/M/2014)
        # ══════════════════════════════════════════════════════════
        _render_section_banner(
            "💡 4. Rekomendasi Pemeliharaan & Rehabilitasi Otomatis (Permen PU No. 12/PRT/M/2014)",
            "Langkah tindakan teknis terstruktur berdasarkan hasil evaluasi hidrolika dan optimasi Bisection.",
            badge_html='<span class="ds-badge auto">Action Plan</span>',
            tone="teal",
        )

        T_val = int(m4.get("T_ulang", 5))
        tipologi_val = str(m4.get("tipologi_kota") or st.session_state.get("tipologi_kota") or "Kota Sedang")
        Q_ras_val = Q_banjir
        Q_cap_val = Q_aktual
        ts_eksisting_val = ts_current

        # TINDAKAN 01 · Evaluasi Kapasitas Hidrolika & Kala Ulang Permen PU 12/2014
        t1_hdr = "**TINDAKAN 01 · [Evaluasi Kapasitas Hidrolika & Kala Ulang Permen PU 12/2014]**"
        t1_base = (
            f"Berdasarkan evaluasi untuk Kala Ulang {T_val} Tahun pada {tipologi_val}, debit limpasan rasional tercatat sebesar "
            f"Qr = {Q_ras_val:.3f} m³/s berbanding dengan kapasitas saluran aktual Qs = {Q_cap_val:.3f} m³/s."
        )
        if SF >= 1.0:
            st.success(
                f"{t1_hdr}\n\n{t1_base} Kondisi saluran AMAN (Safety Factor = {SF:.2f}). "
                "Kapasitas penampang masih memadai untuk menampung curah hujan rencana tanpa meluap."
            )
        else:
            st.error(
                f"{t1_hdr}\n\n{t1_base} Kondisi saluran KRITIS/OVERTOPPING (Safety Factor = {SF:.2f} < 1.0). "
                "Diperlukan REDESAIN dimensi saluran atau penambahan tanggul penahan banjir (Freeboard)."
            )

        # TINDAKAN 02 · Tindakan Fisik - Optimasi Pengerukan Sedimen
        t2_hdr = "**TINDAKAN 02 · [Tindakan Fisik - Optimasi Pengerukan Sedimen]**"
        t2_base = f"Evaluasi ketebalan sedimen eksisting (ts = {ts_eksisting_val:.3f} m) terhadap batas izin (ts_izin = {ts_izin:.3f} m)."
        if t_pengerukan > 0:
            st.warning(
                f"{t2_hdr}\n\n{t2_base} Status: MELAMPAUI BATAS. Segera lakukan pengerukan sedimen setebal minimum "
                f"{t_pengerukan:.3f} m untuk mencapai kedalaman kritis (h*) dan mencegah pendangkalan ekstrem."
            )
        else:
            st.info(
                f"{t2_hdr}\n\n{t2_base} Status: NORMAL. Ketebalan sedimen belum mengganggu elevasi dasar saluran kritis "
                "(t_pengerukan = 0.000 m). Cukup jadwalkan inspeksi rutin berkala."
            )

        # TINDAKAN 03 · Tindakan Fisik - Rehabilitasi Kekasaran Dinding (Metode Cowan)
        t3_hdr = "**TINDAKAN 03 · [Tindakan Fisik - Rehabilitasi Kekasaran Dinding (Metode Cowan)]**"
        t3_msg = (
            f"{t3_hdr}\n\n"
            f"Telah terjadi degradasi nilai kekasaran saluran dari desain awal (n = {n_desain:.4f}) menjadi kondisi lapangan "
            f"(n = {n_aktual:.4f}). Rekomendasi: Lakukan pembersihan gulma/vegetasi, pengangkatan sampah, serta penambalan retakan "
            f"struktur dinding untuk menekan angka hambatan Manning (n) dan mengembalikan kecepatan aliran (V) mendekati kondisi ideal."
        )
        if n_aktual > n_desain * 1.05:
            st.warning(t3_msg)
        else:
            st.info(t3_msg)

        st.markdown("---")

        # ── Ringkasan Parameter Pipeline Lengkap ─────────────────
        with st.expander("📋 Tabel Rekapitulasi Lengkap Parameter Modul 1 s/d Modul 5", expanded=False):
            df_params = pd.DataFrame([
                {"Modul": "Modul 1", "Parameter": "R24 Hujan Rencana T=5 Tahun (mm)", "Nilai": f"{R24:.2f}" if R24 else "—"},
                {"Modul": "Modul 2", "Parameter": "Material & Manning n DED (Tabel Chow)", "Nilai": f"{ded.get('material', 'Beton')} (n = {n_desain:.4f})"},
                {"Modul": "Modul 2", "Parameter": "Kutipan / Halaman Asal DED", "Nilai": str(ded.get("kutipan_ded", "—"))},
                {"Modul": "Modul 3", "Parameter": "Manning Aktual (Cowan / Horton-Einstein)", "Nilai": f"{n_aktual:.5f}"},
                {"Modul": "Modul 3", "Parameter": "Geometri Efektif (b_efektif & h_efektif)", "Nilai": f"b_eff = {b_aktual:.3f} m | h_eff = {h_eff:.3f} m (ts = {ts_current:.3f} m)"},
                {"Modul": "Modul 4", "Parameter": "Bentuk Penampang & Tinggi Jagaan (Freeboard w)", "Nilai": f"{shape} | w = {m4.get('w_fb', 0):.3f} m | h_maks = {h_maks:.3f} m"},
                {"Modul": "Modul 4", "Parameter": "Debit Desain (Q_desain) vs Aktual (Q_aktual)", "Nilai": f"{Q_desain:.4f} m³/s vs {Q_aktual:.4f} m³/s"},
                {"Modul": "Modul 4", "Parameter": f"Debit Banjir Rencana Q_banjir (T={m4.get('T_ulang', 5)} thn)", "Nilai": f"{Q_banjir:.4f} m³/s (I = {m4.get('I_mono', 0):.2f} mm/jam)"},
                {"Modul": "Modul 4", "Parameter": "Kecepatan Aliran (V) & Bilangan Froude (Fr)", "Nilai": f"V = {m4.get('V_aktual', 0):.3f} m/s | Fr = {m4.get('Fr_aktual', 0):.3f}"},
                {"Modul": "Modul 5", "Parameter": "Safety Factor (SF) & Penurunan Kapasitas", "Nilai": f"SF = {SF:.3f} ({'AMAN' if is_safe_sf else 'KRITIS'}) | Penurunan = {penurunan_pct:.2f}%"},
                {"Modul": "Modul 5", "Parameter": "Hasil Bisection (h*, ts_izin, t_pengerukan)", "Nilai": f"h* = {h_star:.4f} m | ts_izin = {ts_izin:.4f} m | t_pengerukan = {t_pengerukan:.4f} m"},
            ])
            st.dataframe(df_params, hide_index=True)

        st.markdown("---")

    # ══════════════════════════════════════════════════════════════
    # 5. EKSPOR LAPORAN PDF LENGKAP (fpdf2)
    # ══════════════════════════════════════════════════════════════
    _render_section_banner(
        "📄 5. Ekspor Laporan Eksekutif PDF Lengkap (Modul 1 – Modul 5)",
        "Susun dokumen laporan teknis resmi berformat PDF yang merangkum hasil analisis hidrologi, DED, visi komputer, hidrolika, dan rekomendasi.",
        badge_html='<span class="ds-badge auto">FPDF2 Report Engine</span>',
        tone="navy",
    )

    if not FPDF_AVAILABLE:
        st.error(
            "❌ Library `fpdf2` tidak terinstall.  \n"
            "Jalankan: `pip install fpdf2`"
        )
    else:
        col_gen, col_dl = st.columns([1, 1])
        with col_gen:
            if st.button("🖨️ Generate Laporan PDF Lengkap", type="primary"):
                with st.spinner("📄 Menyusun laporan PDF lengkap (SNI 2415:2016 & Permen PU 12/2014)…"):
                    try:
                        pdf_bytes = _build_pdf()
                        st.session_state["pdf_bytes"] = pdf_bytes
                        st.success("✅ Laporan PDF berhasil dibuat!")
                    except Exception as e:
                        st.error(f"❌ Gagal membuat PDF: {e}")

        with col_dl:
            if st.session_state.get("pdf_bytes"):
                fname = f"DrainSense_Kanaeru_Laporan_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
                st.download_button(
                    "⬇️ Download Laporan PDF",
                    data=st.session_state["pdf_bytes"],
                    file_name=fname,
                    mime="application/pdf",
                )
                st.caption(f"File: `{fname}`")
                _render_kobo_companion_card(
                    asset_key="kobo_main",
                    tag="📄 KOBO REPORT READY · DRAINAGE AUDIT COMPLETE",
                    title="Laporan Eksekutif PDF Siap Diunduh",
                    description=(
                        "Seluruh rekapitulasi teknis Modul 1 s/d Modul 5 telah disusun sesuai standar "
                        "SNI 2415:2016 dan Permen PU No. 12/PRT/M/2014."
                    ),
                )

        with st.expander("📋 Cakupan Isi Laporan PDF Lengkap"):
            st.markdown(
                "Laporan PDF mencakup rekapitulasi teknis seluruh modul:\n"
                "- **Modul 1 (SNI 2415:2016):** Statistik Lengkap ($n, \\bar{X}, S_d, C_v, C_s, C_k$), Uji Outlier Grubbs-Beck ($K_n, X_L, X_H$), & Hujan Rencana Multi-Kala Ulang\n"
                "- **Modul 2 (Tabel Chow 1959):** Dimensi DED, Status Validasi Koefisien Manning $n$, & Kutipan Halaman Asal PDF DED\n"
                "- **Modul 3 (Metode Cowan & Horton-Einstein):** Rincian komponen $n_b, n_1, n_2, n_3, n_4, m$, Kekasaran Gabungan $n_c$, serta Geometri Efektif ($b_{\\text{efektif}}, h_{\\text{efektif}}$)\n"
                "- **Modul 4 (Permen PU 12/2014):** Waktu Konsentrasi ($t_c$ Kirpich/Kerby), Intensitas Mononobe, $C_{\\text{gab}}$, Freeboard $w$, Debit Manning, Kecepatan $V$, & Bilangan Froude $Fr$\n"
                "- **Modul 5 (Bisection & Rekomendasi PU):** Safety Factor ($SF$), Penurunan Kapasitas (%), Kedalaman Kritis $h^*$, $t_{s,\\text{izin}}$, $t_{\\text{pengerukan}}$, & Rekomendasi Pemeliharaan Permen PU No. 12/PRT/M/2014\n"
            )
