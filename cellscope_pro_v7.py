# -*- coding: utf-8 -*-
"""
CellScope Pro v7 - New Features Update
========================================
NEW in v7:
- Comparative Analysis (Before vs After side-by-side)
- Automated Report Email (SMTP)
- Statistics Dashboard (dedicated page)
- Batch Report — One PDF for Multiple Patients
- QR Code for Patient Report (scan to view report)

Install:
    pip install streamlit opencv-python numpy Pillow scikit-image reportlab plotly openpyxl qrcode[pil]

Run:
    streamlit run cellscope_pro_v7.py
"""

import streamlit as st
import cv2
import numpy as np
from PIL import Image, ImageEnhance
import io
import csv
import datetime
import math
import random
import time
import os
import sys

# Add folder to path so supabase_client and auth can be found
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Supabase — graceful fallback if not connected
try:
    from supabase_client import supabase as sb_client
    from auth import (sign_up, sign_in, sign_out as sb_sign_out,
                      save_patient, load_patients,
                      save_analysis, save_report, load_reports,
                      save_lab_settings, load_lab_settings)
    SUPABASE_OK = True
except Exception:
    SUPABASE_OK = False
    sb_client = None

# Plotly — graceful fallback
try:
    import plotly.graph_objects as go
    PLOTLY_OK = True
except ImportError:
    PLOTLY_OK = False

# Excel export — graceful fallback
try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    EXCEL_OK = True
except ImportError:
    EXCEL_OK = False

# QR Code — graceful fallback
try:
    import qrcode
    QR_OK = True
except ImportError:
    QR_OK = False

# Email (built-in)
import smtplib
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                Table, TableStyle, HRFlowable,
                                Image as RLImage)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

# ═══════════════════════════════════════════════════════════════════════════════
#  PAGE CONFIG — must be first Streamlit call
# ═══════════════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="CellScope Pro",
    page_icon="🔬",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ═══════════════════════════════════════════════════════════════════════════════
#  MEDICAL CONSTANTS
# ═══════════════════════════════════════════════════════════════════════════════

# Normal ranges vary by age and gender
NORMAL_RANGES_DEFAULT = {
    "RBC": {"min": 4,  "max": 30},
    "WBC": {"min": 1,  "max": 10},
}

# Smart ranges by gender
NORMAL_RANGES_GENDER = {
    "Male":   {"RBC": {"min": 5,  "max": 35}, "WBC": {"min": 1, "max": 10}},
    "Female": {"RBC": {"min": 4,  "max": 30}, "WBC": {"min": 1, "max": 10}},
    "Other":  {"RBC": {"min": 4,  "max": 30}, "WBC": {"min": 1, "max": 10}},
}

# Smart ranges by age group
def get_normal_ranges(gender="Male", age=None):
    ranges = NORMAL_RANGES_GENDER.get(gender, NORMAL_RANGES_DEFAULT).copy()
    try:
        age_int = int(age) if age else 30
        if age_int < 12:   # Child
            ranges["WBC"] = {"min": 2, "max": 15}
            ranges["RBC"] = {"min": 3, "max": 25}
        elif age_int > 65: # Elderly
            ranges["WBC"] = {"min": 1, "max": 9}
    except (ValueError, TypeError):
        pass
    return ranges

DEMO_PATIENT = {
    "name": "John Smith", "pid": "PT-2024-001",
    "dob": "15-06-1985", "age": "38", "gender": "Male",
    "mobile": "+91 98765 43210", "email": "john.smith@email.com",
    "doctor": "Dr. Sarah Johnson", "collection_date": "2024-04-30",
    "sample_id": "SMP-2024-001", "notes": "Routine blood examination",
}

AI_INSIGHTS = {
    "NORMAL": [
        "Blood cell counts are within normal range. Morphological analysis suggests high sample stability.",
        "RBC and WBC distribution appears healthy. No abnormalities detected in this sample.",
        "Cell morphology looks good. Blood parameters are within acceptable reference ranges.",
    ],
    "MILD": [
        "Slightly elevated WBC detected. This may indicate mild stress or early infection. Monitor and recheck in 2 weeks.",
        "WBC count is marginally above normal. Consider lifestyle factors and follow up with your physician.",
    ],
    "MODERATE": [
        "Moderately elevated WBC count detected. May indicate an active infection or inflammatory response. Please consult a medical professional promptly.",
        "WBC is notably elevated. Recommend medical consultation within 48 hours for further evaluation.",
    ],
    "SEVERE": [
        "Significantly elevated WBC detected. This requires immediate medical attention. Please consult a doctor today.",
        "Critical WBC elevation observed. Do not delay — please seek immediate medical evaluation.",
    ],
}

DISEASE_HINTS = {
    # (rbc_flag, wbc_flag) -> list of possible conditions
    ("LOW",    "NORMAL"): [
        ("Anemia",            "🔴", "Low RBC count may indicate anemia. Common causes include iron deficiency, vitamin B12 deficiency, or chronic disease."),
        ("Iron Deficiency",   "🟠", "Reduced red blood cells can be a sign of iron deficiency. Consider dietary changes and iron supplementation after consulting a doctor."),
    ],
    ("LOW",    "HIGH"):   [
        ("Malaria",           "🔴", "The combination of low RBC and high WBC is a classic indicator of malaria or other parasitic infections. Seek medical evaluation immediately."),
        ("Sepsis",            "🔴", "Low RBC combined with elevated WBC may suggest a serious systemic infection. Immediate medical attention is recommended."),
        ("Hemolytic Anemia",  "🟠", "Destruction of red blood cells combined with immune response may indicate hemolytic anemia."),
    ],
    ("NORMAL", "HIGH"):   [
        ("Bacterial Infection","🟠", "Elevated WBC count often indicates the body is fighting a bacterial infection. Consult a doctor for proper diagnosis."),
        ("Viral Infection",   "🟡", "Slightly elevated WBC may indicate a viral infection such as flu or common cold. Rest and monitor symptoms."),
        ("Inflammation",      "🟡", "Elevated white blood cells can indicate inflammatory conditions. Further testing recommended."),
    ],
    ("NORMAL", "LOW"):    [
        ("Viral Suppression", "🟠", "Low WBC may indicate viral infections like HIV, hepatitis, or influenza suppressing immune cell production."),
        ("Autoimmune",        "🟠", "Reduced WBC can be associated with autoimmune conditions. Consult a specialist."),
    ],
    ("HIGH",   "NORMAL"): [
        ("Dehydration",       "🟡", "High RBC count relative to normal range may indicate dehydration or hemoconcentration."),
        ("Polycythemia",      "🟠", "Elevated RBC count may indicate polycythemia. Requires medical evaluation."),
    ],
    ("HIGH",   "HIGH"):   [
        ("Leukemia Indicator","🔴", "IMPORTANT: Both elevated RBC and WBC may require urgent evaluation. This pattern warrants immediate consultation with a hematologist."),
        ("Chronic Infection", "🟠", "Persistently elevated cell counts may indicate chronic infection or bone marrow disorder."),
    ],
    ("NORMAL", "NORMAL"): [
        ("Healthy Profile",   "🟢", "Blood cell counts appear within normal range. No concerning patterns detected based on count analysis."),
    ],
    ("LOW",    "LOW"):    [
        ("Bone Marrow Issue", "🔴", "Both low RBC and WBC may indicate bone marrow suppression. Immediate medical evaluation is strongly recommended."),
        ("Aplastic Anemia",   "🔴", "Pancytopenia pattern detected. This requires urgent specialist evaluation."),
    ],
    ("HIGH",   "LOW"):    [
        ("Stress Response",   "🟡", "This combination may indicate a physiological stress response or medication effect. Consult your doctor."),
    ],
}


# ═══════════════════════════════════════════════════════════════════════════════
#  HELPER FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════════

def get_flag(label, count, gender="Male", age=None):
    ranges = get_normal_ranges(gender, age)
    lo = ranges[label]["min"]
    hi = ranges[label]["max"]
    if count < lo: return "LOW"
    if count > hi: return "HIGH"
    return "NORMAL"

def get_severity(wbc):
    if wbc <= 5:  return "NORMAL",   "#22C55E"
    if wbc <= 12: return "MILD",     "#EAB308"
    if wbc <= 20: return "MODERATE", "#F97316"
    return "SEVERE", "#EF4444"

def get_health_score(rbc, wbc, gender="Male", age=None):
    rbc_flag = get_flag("RBC", rbc, gender, age)
    wbc_flag = get_flag("WBC", wbc, gender, age)
    score = 100
    if rbc_flag == "LOW":  score -= 20
    if rbc_flag == "HIGH": score -= 15
    sev, _ = get_severity(wbc)
    if sev == "MILD":     score -= 15
    if sev == "MODERATE": score -= 35
    if sev == "SEVERE":   score -= 55
    return max(0, min(100, score))

def get_confidence(image_bgr):
    gray = cv2.cvtColor(image_bgr, cv2.COLOR_BGR2GRAY)
    lap_var = cv2.Laplacian(gray, cv2.CV_64F).var()
    mean_b  = np.mean(gray)
    return min(98, max(45, int(lap_var / 50 + mean_b / 10)))

def get_image_quality(conf):
    if conf >= 80: return "Good Quality",  "#22C55E", "🟢"
    if conf >= 60: return "Fair Quality",  "#EAB308", "🟡"
    return "Poor Quality", "#EF4444", "🔴"

def check_is_blood_smear(image_bgr):
    """Basic validator to check if image looks like a blood smear."""
    hsv  = cv2.cvtColor(image_bgr, cv2.COLOR_BGR2HSV)
    mean_sat = np.mean(hsv[:,:,1])
    mean_val = np.mean(hsv[:,:,2])
    h, w = image_bgr.shape[:2]
    if h < 100 or w < 100:
        return False, "Image too small. Minimum 100x100 pixels required."
    if mean_sat < 20:
        return False, "Image appears to be greyscale or lacks colour. Blood smears should have pink/purple staining."
    return True, "OK"

def get_ai_insight(severity):
    return random.choice(AI_INSIGHTS.get(severity, AI_INSIGHTS["NORMAL"]))

def get_disease_hints(rbc_flag, wbc_flag):
    key = (rbc_flag, wbc_flag)
    return DISEASE_HINTS.get(key, DISEASE_HINTS[("NORMAL","NORMAL")])

def make_report_id():
    n = st.session_state.report_counter
    st.session_state.report_counter += 1
    return "RPT-{0}-{1:04d}".format(
        datetime.date.today().strftime("%Y%m%d"), n)

def get_patient():
    """Get patient dict from session state."""
    return {
        "name":            st.session_state.get("pt_name", ""),
        "pid":             st.session_state.get("pt_pid", ""),
        "dob":             st.session_state.get("pt_dob", ""),
        "age":             st.session_state.get("pt_age", ""),
        "gender":          st.session_state.get("pt_gender", "Male"),
        "mobile":          st.session_state.get("pt_mobile", ""),
        "email":           st.session_state.get("pt_email", ""),
        "doctor":          st.session_state.get("pt_doctor", ""),
        "collection_date": st.session_state.get("pt_cdate", ""),
        "sample_id":       st.session_state.get("pt_sid", ""),
        "notes":           st.session_state.get("pt_notes", ""),
    }

def get_lab():
    return {
        "name":          st.session_state.get("lab_name", ""),
        "address":       st.session_state.get("lab_addr", ""),
        "phone":         st.session_state.get("lab_phone", ""),
        "accreditation": st.session_state.get("lab_acc", ""),
        "technician":    st.session_state.get("lab_tech", ""),
    }

def badge_html(flag):
    styles = {
        "NORMAL": ("background:#DCFCE7;color:#16A34A;", "✅"),
        "HIGH":   ("background:#FEE2E2;color:#DC2626;", "🚨"),
        "LOW":    ("background:#FEF3C7;color:#D97706;", "⚠️"),
    }
    style, icon = styles.get(flag, styles["NORMAL"])
    return (
        '<span style="{s}border-radius:6px;padding:2px 8px;'
        'font-family:Roboto,sans-serif;font-size:10px;'
        'font-weight:700;letter-spacing:1px;">'
        '{i} {f}</span>'.format(s=style, i=icon, f=flag))


# ═══════════════════════════════════════════════════════════════════════════════
#  SVG COMPONENTS (no HTML/CSS mixing issues)
# ═══════════════════════════════════════════════════════════════════════════════

def health_gauge_svg(score, color):
    """Pure SVG circular gauge — no div mixing."""
    label = ("Excellent" if score >= 90 else
             "Good"      if score >= 75 else
             "Fair"      if score >= 60 else "Poor")
    # Arc: 220 degrees starting from -200deg to 20deg
    def arc(a1, a2, r, cx=80, cy=85):
        x1 = cx + r * math.cos(math.radians(a1))
        y1 = cy + r * math.sin(math.radians(a1))
        x2 = cx + r * math.cos(math.radians(a2))
        y2 = cy + r * math.sin(math.radians(a2))
        lg = 1 if abs(a2 - a1) > 180 else 0
        return "M {:.1f},{:.1f} A {},{} 0 {} 1 {:.1f},{:.1f}".format(
            x1, y1, r, r, lg, x2, y2)
    bg   = arc(-200, 20, 60)
    fill = arc(-200, -200 + 220 * score / 100, 60)
    return """<svg width="160" height="130" viewBox="0 0 160 130"
              xmlns="http://www.w3.org/2000/svg">
  <path d="{bg}" fill="none" stroke="#E2E8F0" stroke-width="10"
        stroke-linecap="round"/>
  <path d="{fp}" fill="none" stroke="{c}" stroke-width="10"
        stroke-linecap="round"/>
  <text x="80" y="88" text-anchor="middle" font-family="sans-serif"
        font-size="28" font-weight="bold" fill="{c}">{sc}</text>
  <text x="80" y="104" text-anchor="middle" font-family="sans-serif"
        font-size="11" fill="#94A3B8">{lb}</text>
</svg>""".format(bg=bg, fp=fill, c=color, sc=score, lb=label)


def metric_bar_html(value, lo, hi, max_val=30):
    """HTML metric distribution bar."""
    pct    = min(100, max(0, value / max_val * 100))
    lo_pct = lo / max_val * 100
    hi_pct = hi / max_val * 100
    return (
        '<div style="margin:10px 0 4px 0;">'
        '<div style="position:relative;height:10px;border-radius:5px;'
        'background:linear-gradient(90deg,'
        '#FEE2E2 0%,#FEF3C7 {lp:.0f}%,'
        '#DCFCE7 {lp:.0f}%,#DCFCE7 {hp:.0f}%,'
        '#FEF3C7 {hp:.0f}%,#FEE2E2 100%);">'
        '<div style="position:absolute;top:-3px;left:{p:.1f}%;'
        'width:16px;height:16px;border-radius:50%;'
        'background:#1A56DB;border:2px solid white;'
        'box-shadow:0 2px 6px rgba(26,86,219,0.5);'
        'transform:translateX(-50%);"></div></div>'
        '<div style="display:flex;justify-content:space-between;'
        'margin-top:6px;font-family:Roboto,sans-serif;'
        'font-size:10px;color:#94A3B8;">'
        '<span>LOW</span><span>NORMAL</span><span>HIGH</span>'
        '</div></div>'
    ).format(p=pct, lp=lo_pct, hp=hi_pct)


# ═══════════════════════════════════════════════════════════════════════════════
#  CELL DETECTION ENGINE
# ═══════════════════════════════════════════════════════════════════════════════

def detect_cells(image_bgr, sensitivity=0.45):
    h, w = image_bgr.shape[:2]
    hsv  = cv2.cvtColor(image_bgr, cv2.COLOR_BGR2HSV)
    gray = cv2.cvtColor(image_bgr, cv2.COLOR_BGR2GRAY)
    blurred  = cv2.GaussianBlur(gray, (9, 9), 2)
    clahe    = cv2.createCLAHE(clipLimit=2.5, tileGridSize=(8, 8))
    enhanced = clahe.apply(blurred)
    _, binary = cv2.threshold(enhanced, 0, 255,
                              cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    kernel  = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (5, 5))
    cleaned = cv2.morphologyEx(binary,  cv2.MORPH_OPEN,  kernel, iterations=2)
    cleaned = cv2.morphologyEx(cleaned, cv2.MORPH_CLOSE, kernel, iterations=2)
    dist    = cv2.distanceTransform(cleaned, cv2.DIST_L2, 5)
    max_val = dist.max()
    if max_val == 0:
        return image_bgr.copy(), [], 0, 0
    _, sure_fg = cv2.threshold(dist, sensitivity * max_val, 255, 0)
    sure_fg    = np.uint8(sure_fg)
    unknown    = cv2.subtract(cleaned, sure_fg)
    _, markers = cv2.connectedComponents(sure_fg)
    markers    = markers + 1
    markers[unknown == 255] = 0
    markers    = markers.astype(np.int32)
    markers    = cv2.watershed(image_bgr.copy(), markers)
    annotated  = image_bgr.copy()
    cells      = []
    cell_id    = 1
    for lbl in np.unique(markers):
        if lbl < 2:
            continue
        mask = np.zeros(gray.shape, dtype=np.uint8)
        mask[markers == lbl] = 255
        contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL,
                                       cv2.CHAIN_APPROX_SIMPLE)
        if not contours:
            continue
        cnt  = max(contours, key=cv2.contourArea)
        area = cv2.contourArea(cnt)
        if area < 200 or area > 0.15 * h * w:
            continue
        (cx, cy), radius = cv2.minEnclosingCircle(cnt)
        cx, cy = int(cx), int(cy)
        r      = max(int(radius), 1)
        x1, y1 = max(cx-r, 0), max(cy-r, 0)
        x2, y2 = min(cx+r, w), min(cy+r, h)
        mv  = np.mean(gray[y1:y2, x1:x2])
        ms  = np.mean(hsv[y1:y2, x1:x2, 1])
        lbl2 = "WBC" if (area > 1800 and mv < 160 and ms > 50) else "RBC"
        col  = (59, 130, 246) if lbl2 == "RBC" else (8, 145, 178)
        cv2.circle(annotated, (cx, cy), r+3, col, 2)
        cv2.putText(annotated, "#{0}".format(cell_id),
                    (cx-r, cy-r-4), cv2.FONT_HERSHEY_SIMPLEX,
                    max(0.3, min(0.5, r/40)), col, 1, cv2.LINE_AA)
        cells.append(dict(id=cell_id, label=lbl2,
                          x=cx, y=cy, radius=r, area=int(area)))
        cell_id += 1
    rbc = sum(1 for c in cells if c["label"] == "RBC")
    wbc = sum(1 for c in cells if c["label"] == "WBC")
    return annotated, cells, rbc, wbc


def make_heatmap(image_bgr, cells):
    """Generate cell density heatmap overlay."""
    h, w = image_bgr.shape[:2]
    heat = np.zeros((h, w), dtype=np.float32)
    for c in cells:
        r = max(c["radius"] * 2, 20)
        x1, y1 = max(c["x"]-r, 0), max(c["y"]-r, 0)
        x2, y2 = min(c["x"]+r, w), min(c["y"]+r, h)
        heat[y1:y2, x1:x2] += 1
    if heat.max() > 0:
        heat = heat / heat.max()
    heat_uint8 = (heat * 255).astype(np.uint8)
    heat_color = cv2.applyColorMap(heat_uint8, cv2.COLORMAP_JET)
    overlay    = cv2.addWeighted(image_bgr, 0.6, heat_color, 0.4, 0)
    return overlay


# ═══════════════════════════════════════════════════════════════════════════════
#  PDF GENERATOR
# ═══════════════════════════════════════════════════════════════════════════════

def generate_pdf(patient, lab, images_data, report_id,
                 timestamp, annot_pil=None):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=15*mm, rightMargin=15*mm,
                            topMargin=12*mm, bottomMargin=12*mm)
    C = {
        "bg":    colors.HexColor("#F8FAFC"),
        "white": colors.white,
        "blue":  colors.HexColor("#1A56DB"),
        "rbc":   colors.HexColor("#DC2626"),
        "wbc":   colors.HexColor("#0891B2"),
        "dark":  colors.HexColor("#0F172A"),
        "grey":  colors.HexColor("#64748B"),
        "lgrey": colors.HexColor("#94A3B8"),
        "bdr":   colors.HexColor("#E2E8F0"),
        "green": colors.HexColor("#16A34A"),
        "orange":colors.HexColor("#EA580C"),
        "red":   colors.HexColor("#DC2626"),
        "sub":   colors.HexColor("#93C5FD"),
    }

    def PS(n, **kw): return ParagraphStyle(n, **kw)
    def H2(t):
        return Paragraph(t, PS("h2", fontName="Helvetica-Bold",
                               fontSize=9, textColor=C["blue"],
                               spaceBefore=6, spaceAfter=3))
    def HR():
        return HRFlowable(width="100%", thickness=0.3,
                          color=C["bdr"], spaceAfter=3)
    def BD(t):
        return Paragraph(str(t), PS("bd", fontName="Helvetica",
                                    fontSize=8, textColor=C["dark"],
                                    leading=12))
    def SM(t):
        return Paragraph(str(t), PS("sm", fontName="Helvetica",
                                    fontSize=7, textColor=C["grey"],
                                    leading=11))
    def BP(t, c, sz=12):
        return Paragraph(str(t), PS("bp", fontName="Helvetica-Bold",
                                    fontSize=sz, textColor=c,
                                    leading=sz+3, alignment=TA_CENTER))
    def fc(flag):
        return {"NORMAL":C["green"],"HIGH":C["red"],"LOW":C["orange"]}[flag]

    story = []
    gender = patient.get("gender", "Male")
    age    = patient.get("age", None)

    # Header rows
    for row_data, pt, pb in [
        ([[Paragraph(lab.get("name","CellScope Diagnostics"),
                     PS("ln",fontName="Helvetica-Bold",fontSize=15,
                        textColor=C["white"],leading=19)),
           Paragraph("BLOOD CELL ANALYSIS REPORT",
                     PS("rh",fontName="Helvetica-Bold",fontSize=8,
                        textColor=C["white"],leading=11,
                        alignment=TA_RIGHT,letterSpacing=1.5))]], 12, 5),
        ([[Paragraph("{0} | {1} | {2}".format(
                lab.get("address",""), lab.get("phone",""),
                lab.get("accreditation","")),
            PS("ls",fontName="Helvetica",fontSize=7,
               textColor=C["sub"],leading=11)),
           Paragraph("Report ID: {0}\nDate: {1}".format(
               report_id, timestamp),
            PS("rs",fontName="Helvetica",fontSize=7,
               textColor=C["sub"],leading=11,alignment=TA_RIGHT))]], 2, 10),
    ]:
        t = Table(row_data, colWidths=[110*mm, 62*mm])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,-1), C["blue"]),
            ("TOPPADDING",    (0,0), (-1,-1), pt),
            ("BOTTOMPADDING", (0,0), (-1,-1), pb),
            ("LEFTPADDING",   (0,0), (-1,-1), 14),
            ("RIGHTPADDING",  (0,0), (-1,-1), 14),
        ]))
        story.append(t)
    story.append(Spacer(1, 4*mm))

    # Patient info
    p = patient
    story.append(H2("PATIENT INFORMATION"))
    story.append(HR())
    pi_data = [
        [SM("Name"),    BD(p.get("name","N/A")),
         SM("Patient ID"), BD(p.get("pid","N/A"))],
        [SM("DOB"),     BD(p.get("dob","N/A")),
         SM("Age/Gender"), BD("{0} / {1}".format(p.get("age",""), p.get("gender","")))],
        [SM("Mobile"),  BD(p.get("mobile","N/A")),
         SM("Email"),   BD(p.get("email","N/A"))],
        [SM("Doctor"),  BD(p.get("doctor","N/A")),
         SM("Technician"), BD(lab.get("technician","N/A"))],
        [SM("Collection"), BD(p.get("collection_date","N/A")),
         SM("Sample ID"),  BD(p.get("sample_id","N/A"))],
        [SM("Notes"),   BD(p.get("notes","—")),
         SM(""),        BD("")],
    ]
    pi_tbl = Table(pi_data, colWidths=[28*mm, 62*mm, 24*mm, 58*mm])
    pi_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0,0), (-1,-1), C["bg"]),
        ("GRID",          (0,0), (-1,-1), 0.25, C["bdr"]),
        ("TOPPADDING",    (0,0), (-1,-1), 4),
        ("BOTTOMPADDING", (0,0), (-1,-1), 4),
        ("LEFTPADDING",   (0,0), (-1,-1), 7),
        ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
    ]))
    story.append(pi_tbl)
    story.append(Spacer(1, 4*mm))

    # Summary
    total_rbc = sum(d["rbc"] for d in images_data)
    total_wbc = sum(d["wbc"] for d in images_data)
    total_all = total_rbc + total_wbc
    ratio_str = "{0}:{1}".format(total_rbc, total_wbc) \
                if total_wbc else "{0}:0".format(total_rbc)
    rbc_flag  = get_flag("RBC", total_rbc, gender, age)
    wbc_flag  = get_flag("WBC", total_wbc, gender, age)
    sev_txt, sev_hex = get_severity(total_wbc)
    sev_c    = colors.HexColor(sev_hex)
    score    = get_health_score(total_rbc, total_wbc, gender, age)
    nr       = get_normal_ranges(gender, age)

    story.append(H2("ANALYSIS SUMMARY"))
    story.append(HR())
    s_hdr  = [SM("PARAMETER"), SM("COUNT"), SM("NORMAL RANGE"),
              SM("STATUS"), SM("SEVERITY")]
    s_rows = [
        [BD("Red Blood Cells (RBC)"),
         BP(str(total_rbc), C["rbc"]),
         SM("{0}-{1} detected".format(nr["RBC"]["min"], nr["RBC"]["max"])),
         BP(rbc_flag, fc(rbc_flag), 8), SM("—")],
        [BD("White Blood Cells (WBC)"),
         BP(str(total_wbc), C["wbc"]),
         SM("{0}-{1} detected".format(nr["WBC"]["min"], nr["WBC"]["max"])),
         BP(wbc_flag, fc(wbc_flag), 8),
         BP(sev_txt, sev_c, 8)],
        [BD("Total Cells"),
         BP(str(total_all), C["blue"]),
         SM("—"), SM("—"), SM("—")],
        [BD("RBC:WBC Ratio"),
         BP(ratio_str, colors.HexColor("#F97316"), 10),
         SM("—"), SM("—"), SM("—")],
        [BD("Blood Health Score"),
         BP("{0}/100".format(score), sev_c),
         SM("0-100"), SM("—"), BP(sev_txt, sev_c, 8)],
    ]
    s_tbl = Table([s_hdr]+s_rows,
                  colWidths=[45*mm, 25*mm, 36*mm, 28*mm, 38*mm])
    s_tbl.setStyle(TableStyle([
        ("BACKGROUND",     (0,0), (-1,0),  colors.HexColor("#EFF6FF")),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [C["white"], C["bg"]]),
        ("GRID",           (0,0), (-1,-1), 0.25, C["bdr"]),
        ("ALIGN",          (1,0), (4,-1),  "CENTER"),
        ("VALIGN",         (0,0), (-1,-1), "MIDDLE"),
        ("TOPPADDING",     (0,0), (-1,-1), 5),
        ("BOTTOMPADDING",  (0,0), (-1,-1), 5),
        ("LEFTPADDING",    (0,0), (-1,-1), 7),
        ("TEXTCOLOR",      (0,0), (-1,0),  C["blue"]),
        ("FONTNAME",       (0,0), (-1,0),  "Helvetica-Bold"),
        ("FONTSIZE",       (0,0), (-1,0),  7),
    ]))
    story.append(s_tbl)
    story.append(Spacer(1, 4*mm))

    # Disease hints in PDF
    hints = get_disease_hints(rbc_flag, wbc_flag)
    if hints:
        story.append(H2("CLINICAL INDICATORS"))
        story.append(HR())
        hint_rows = []
        for name, dot, desc in hints:
            hint_rows.append([
                Paragraph("{0} {1}".format(dot, name),
                          PS("hn", fontName="Helvetica-Bold", fontSize=8,
                             textColor=C["dark"], leading=12)),
                Paragraph(desc,
                          PS("hd", fontName="Helvetica-Oblique", fontSize=7,
                             textColor=C["grey"], leading=11)),
            ])
        h_tbl = Table(hint_rows, colWidths=[40*mm, 132*mm])
        h_tbl.setStyle(TableStyle([
            ("BACKGROUND",    (0,0), (-1,-1), C["bg"]),
            ("GRID",          (0,0), (-1,-1), 0.25, C["bdr"]),
            ("TOPPADDING",    (0,0), (-1,-1), 5),
            ("BOTTOMPADDING", (0,0), (-1,-1), 5),
            ("LEFTPADDING",   (0,0), (-1,-1), 7),
            ("VALIGN",        (0,0), (-1,-1), "TOP"),
        ]))
        story.append(h_tbl)
        story.append(Paragraph(
            "DISCLAIMER: Clinical indicators above are educational only "
            "and NOT a medical diagnosis. Always consult a qualified doctor.",
            PS("disc2", fontName="Helvetica-Oblique", fontSize=6,
               textColor=C["lgrey"], leading=9)))
        story.append(Spacer(1, 4*mm))

    # Annotated image
    if annot_pil is not None:
        story.append(H2("ANNOTATED BLOOD SMEAR"))
        story.append(HR())
        ibuf = io.BytesIO()
        annot_pil.save(ibuf, format="PNG")
        ibuf.seek(0)
        story.append(RLImage(ibuf, width=100*mm, height=70*mm))
        story.append(Spacer(1, 4*mm))

    # Per image table
    if len(images_data) > 1:
        story.append(H2("PER IMAGE RESULTS"))
        story.append(HR())
        rows = [[SM("IMAGE"), SM("TOTAL"), SM("RBC"), SM("WBC"),
                 SM("RATIO"), SM("RBC STATUS"), SM("WBC STATUS"),
                 SM("NOTES")]]
        for d in images_data:
            t2   = d["rbc"] + d["wbc"]
            r2   = "{0}:{1}".format(d["rbc"], d["wbc"]) \
                   if d["wbc"] else "{0}:0".format(d["rbc"])
            rf   = get_flag("RBC", d["rbc"], gender, age)
            wf   = get_flag("WBC", d["wbc"], gender, age)
            rows.append([
                SM(str(d["name"])[:25]),
                BP(str(t2),       C["blue"], 8),
                BP(str(d["rbc"]), C["rbc"],  8),
                BP(str(d["wbc"]), C["wbc"],  8),
                SM(r2),
                BP(rf, fc(rf), 7),
                BP(wf, fc(wf), 7),
                SM(str(d.get("notes",""))[:25]),
            ])
        it = Table(rows, colWidths=[30*mm,14*mm,14*mm,14*mm,
                                    14*mm,18*mm,18*mm,50*mm])
        it.setStyle(TableStyle([
            ("BACKGROUND",     (0,0), (-1,0),  colors.HexColor("#EFF6FF")),
            ("ROWBACKGROUNDS", (0,1), (-1,-1), [C["white"], C["bg"]]),
            ("GRID",           (0,0), (-1,-1), 0.25, C["bdr"]),
            ("ALIGN",          (1,0), (6,-1),  "CENTER"),
            ("VALIGN",         (0,0), (-1,-1), "MIDDLE"),
            ("TOPPADDING",     (0,0), (-1,-1), 4),
            ("BOTTOMPADDING",  (0,0), (-1,-1), 4),
            ("LEFTPADDING",    (0,0), (-1,-1), 5),
            ("TEXTCOLOR",      (0,0), (-1,0),  C["blue"]),
            ("FONTNAME",       (0,0), (-1,0),  "Helvetica-Bold"),
            ("FONTSIZE",       (0,0), (-1,0),  7),
        ]))
        story.append(it)
        story.append(Spacer(1, 4*mm))

    # Signature + footer
    story.append(H2("DECLARATION"))
    story.append(HR())
    sig = [[
        Paragraph("This report was generated using CellScope Pro automated "
                  "image analysis. Results must be reviewed by a qualified "
                  "medical professional. Not for use as sole diagnostic tool.",
                  PS("d", fontName="Helvetica-Oblique", fontSize=7,
                     textColor=C["grey"], leading=12)),
        Paragraph("Authorised by\n\n\n_____________________\n{0}\n{1}".format(
            lab.get("technician",""), lab.get("name","")),
                  PS("s", fontName="Helvetica", fontSize=7,
                     textColor=C["grey"], leading=12, alignment=TA_RIGHT)),
    ]]
    sig_t = Table(sig, colWidths=[110*mm, 62*mm])
    sig_t.setStyle(TableStyle([
        ("BACKGROUND",    (0,0), (-1,-1), C["bg"]),
        ("LINEABOVE",     (0,0), (-1,0),  0.5, C["bdr"]),
        ("TOPPADDING",    (0,0), (-1,-1), 10),
        ("BOTTOMPADDING", (0,0), (-1,-1), 10),
        ("LEFTPADDING",   (0,0), (-1,-1), 10),
    ]))
    story.append(sig_t)
    story.append(Spacer(1, 3*mm))
    story.append(HR())
    story.append(Paragraph(
        "CellScope Pro v6  |  {0}  |  {1}  |  "
        "For research and educational purposes only.".format(
            report_id, timestamp),
        PS("ft", fontName="Helvetica", fontSize=6,
           textColor=C["lgrey"], leading=9, alignment=TA_CENTER)))
    doc.build(story)
    buf.seek(0)
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════════════════
#  EXCEL EXPORT
# ═══════════════════════════════════════════════════════════════════════════════

def generate_excel(images_data, patient, history):
    """Generate formatted Excel file with all results."""
    buf = io.BytesIO()
    if not EXCEL_OK:
        return None
    wb = openpyxl.Workbook()

    # Colors
    blue_fill   = PatternFill("solid", fgColor="1A56DB")
    green_fill  = PatternFill("solid", fgColor="DCFCE7")
    red_fill    = PatternFill("solid", fgColor="FEE2E2")
    yellow_fill = PatternFill("solid", fgColor="FEF3C7")
    grey_fill   = PatternFill("solid", fgColor="F8FAFC")
    white_font  = Font(color="FFFFFF", bold=True)
    dark_font   = Font(color="0F172A")
    bold_font   = Font(color="0F172A", bold=True)

    thin = Side(style="thin", color="E2E8F0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def hdr(ws, row, cols):
        for i, (col, width) in enumerate(cols, 1):
            cell = ws.cell(row=row, column=i, value=col)
            cell.fill = blue_fill
            cell.font = white_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border
            ws.column_dimensions[
                openpyxl.utils.get_column_letter(i)].width = width

    # Sheet 1 — Summary
    ws1 = wb.active
    ws1.title = "Summary"
    ws1.row_dimensions[1].height = 30
    ws1["A1"] = "CellScope Pro — Analysis Summary"
    ws1["A1"].font = Font(size=14, bold=True, color="1A56DB")
    ws1["A2"] = "Patient: {0}  |  Date: {1}  |  Report: {2}".format(
        patient.get("name","N/A"),
        datetime.date.today().strftime("%Y-%m-%d"),
        "CellScope Export")
    ws1["A2"].font = Font(size=10, color="64748B")
    ws1.row_dimensions[3].height = 5

    hdr(ws1, 4, [
        ("Image Name", 30), ("Total Cells", 12), ("RBC", 10),
        ("WBC", 10), ("RBC Flag", 12), ("WBC Flag", 12),
        ("Ratio", 12), ("Severity", 14), ("Health Score", 14),
        ("Notes", 30),
    ])
    for i, d in enumerate(images_data, 5):
        rbc_f = get_flag("RBC", d["rbc"],
                         patient.get("gender","Male"),
                         patient.get("age"))
        wbc_f = get_flag("WBC", d["wbc"],
                         patient.get("gender","Male"),
                         patient.get("age"))
        sev, _ = get_severity(d["wbc"])
        score  = get_health_score(d["rbc"], d["wbc"],
                                  patient.get("gender","Male"),
                                  patient.get("age"))
        row = [
            d["name"], d["rbc"]+d["wbc"], d["rbc"], d["wbc"],
            rbc_f, wbc_f,
            "{0}:{1}".format(d["rbc"], d["wbc"]) if d["wbc"] else "{0}:0".format(d["rbc"]),
            sev, score, d.get("notes",""),
        ]
        for j, val in enumerate(row, 1):
            cell = ws1.cell(row=i, column=j, value=val)
            cell.border = border
            cell.alignment = Alignment(vertical="center")
            if j == 5:  # RBC flag
                cell.fill = green_fill if rbc_f=="NORMAL" \
                            else red_fill if rbc_f=="HIGH" else yellow_fill
            if j == 6:  # WBC flag
                cell.fill = green_fill if wbc_f=="NORMAL" \
                            else red_fill if wbc_f=="HIGH" else yellow_fill

    # Sheet 2 — Patient
    ws2 = wb.create_sheet("Patient Info")
    hdr(ws2, 1, [("Field",25),("Value",40)])
    p = patient
    fields = [
        ("Full Name",        p.get("name","")),
        ("Patient ID",       p.get("pid","")),
        ("Date of Birth",    p.get("dob","")),
        ("Age",              p.get("age","")),
        ("Gender",           p.get("gender","")),
        ("Mobile",           p.get("mobile","")),
        ("Email",            p.get("email","")),
        ("Referring Doctor", p.get("doctor","")),
        ("Sample ID",        p.get("sample_id","")),
        ("Collection Date",  p.get("collection_date","")),
        ("Clinical Notes",   p.get("notes","")),
    ]
    for i,(f,v) in enumerate(fields,2):
        ws2.cell(row=i,column=1,value=f).font = bold_font
        ws2.cell(row=i,column=2,value=v).font = dark_font
        for c in [1,2]:
            ws2.cell(row=i,column=c).border = border
            ws2.cell(row=i,column=c).fill = grey_fill if i%2==0 else PatternFill("solid",fgColor="FFFFFF")

    # Sheet 3 — History
    if history:
        ws3 = wb.create_sheet("History")
        hdr(ws3, 1, [
            ("Report ID",12),("Date",16),("Patient",20),
            ("Images",8),("RBC",8),("WBC",8),("Total",8),
            ("Severity",12),("Score",10),
        ])
        for i,h in enumerate(history,2):
            row = [h.get("report_id",""),h.get("timestamp",""),
                   h.get("patient",""),h.get("images",1),
                   h.get("rbc",0),h.get("wbc",0),
                   h.get("rbc",0)+h.get("wbc",0),
                   h.get("severity",""),h.get("score","")]
            for j,val in enumerate(row,1):
                ws3.cell(row=i,column=j,value=val).border = border

    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════════════════
#  SESSION STATE
# ═══════════════════════════════════════════════════════════════════════════════

def init_state():
    today = datetime.date.today().strftime("%Y-%m-%d")
    defaults = {
        "app_stage":  "splash",
        "splash_seen": False,
        "user_name":  "",
        "user_email": "",
        "user_role":  "Lab Technician",
        "user_id":    None,
        "sb_loaded":  False,
        "page":       "Analyse",
        "images":     {},
        "selected":   None,
        "_do_batch":  False,
        "_sens":      0.45,
        "batch_done": False,
        "pdf_data":   None,
        "pdf_name":   "",
        "report_counter": 1,
        "history":    [],
        "pt_name":"","pt_pid":"","pt_dob":"","pt_age":"",
        "pt_gender":"Male","pt_mobile":"","pt_email":"",
        "pt_doctor":"","pt_cdate":today,"pt_sid":"SMP-001",
        "pt_notes":"",
        "lab_name":"CellScope Diagnostics",
        "lab_addr":"123 Medical Lane, City",
        "lab_phone":"+91 00000 00000",
        "lab_acc":"Acc. No: LAB-2024",
        "lab_tech":"Lab Technician",
        "saved_patients": [],
        "show_cell_log":  False,
        "show_heatmap":   False,
        "brightness":     1.0,
        "contrast":       1.0,
        "show_annot":     True,
        "dev_name":  "Ajinkya Holkar",
        "dev_title": "Final Year Project · 2026",
        # v7 new features
        "compare_img_a": None,
        "compare_img_b": None,
        "email_status":  "",
        "batch_patients": [],   # list of dicts for batch PDF
        "qr_report_data": {},   # report_id -> base64 encoded pdf
        "batch_pdf_ready": False,
        "batch_pdf_data":  None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()


# ═══════════════════════════════════════════════════════════════════════════════
#  RUN ANALYSE
# ═══════════════════════════════════════════════════════════════════════════════

def run_analyse(name, sensitivity):
    rec = st.session_state.images[name]
    t_start = time.time()
    annot, cells, rbc, wbc = detect_cells(rec["bgr"], sensitivity)
    t_end   = time.time()
    elapsed = round(t_end - t_start, 2)

    gender  = st.session_state.get("pt_gender", "Male")
    age     = st.session_state.get("pt_age", None)
    conf    = get_confidence(rec["bgr"])
    score   = get_health_score(rbc, wbc, gender, age)
    sev_txt, _ = get_severity(wbc)
    insight = get_ai_insight(sev_txt)
    rbc_flag = get_flag("RBC", rbc, gender, age)
    wbc_flag = get_flag("WBC", wbc, gender, age)
    heatmap  = make_heatmap(rec["bgr"], cells) if cells else None

    rec.update({
        "annot":        annot,
        "heatmap":      heatmap,
        "cells":        cells,
        "rbc":          rbc,
        "wbc":          wbc,
        "analysed":     True,
        "confidence":   conf,
        "health_score": score,
        "ai_insight":   insight,
        "elapsed":      elapsed,
    })

    # Save to Supabase
    if SUPABASE_OK and st.session_state.get("user_id"):
        try:
            save_analysis(
                user_id=st.session_state.user_id,
                image_name=name,
                rbc=rbc, wbc=wbc,
                rbc_flag=rbc_flag, wbc_flag=wbc_flag,
                severity=sev_txt,
                confidence=conf,
                health_score=score,
                ai_insight=insight,
                notes=rec.get("notes",""))
        except Exception:
            pass


# ═══════════════════════════════════════════════════════════════════════════════
#  LOAD FROM SUPABASE (runs once after login)
# ═══════════════════════════════════════════════════════════════════════════════

def load_from_supabase():
    if not SUPABASE_OK:
        return
    uid = st.session_state.get("user_id")
    if not uid or st.session_state.get("sb_loaded", False):
        return
    try:
        lab = load_lab_settings(uid)
        if lab:
            st.session_state.lab_name  = lab.get("name",  st.session_state.lab_name)
            st.session_state.lab_addr  = lab.get("address", st.session_state.lab_addr)
            st.session_state.lab_phone = lab.get("phone", st.session_state.lab_phone)
            st.session_state.lab_acc   = lab.get("accreditation", st.session_state.lab_acc)
            st.session_state.lab_tech  = lab.get("technician", st.session_state.lab_tech)
    except Exception:
        pass
    try:
        patients = load_patients(uid)
        if patients:
            st.session_state.saved_patients = [
                {"name": p.get("full_name",""),
                 "pid":  p.get("patient_id",""),
                 "dob":  p.get("date_of_birth",""),
                 "age":  p.get("age",""),
                 "gender": p.get("gender","Male"),
                 "mobile": p.get("mobile",""),
                 "email":  p.get("email",""),
                 "doctor": p.get("referring_doctor",""),
                 "sample_id":       p.get("sample_id",""),
                 "collection_date": p.get("collection_date",""),
                 "notes": p.get("clinical_notes",""),
                 } for p in patients]
    except Exception:
        pass
    try:
        reports = load_reports(uid)
        if reports:
            st.session_state.history = [
                {"report_id": r.get("report_id",""),
                 "timestamp": str(r.get("created_at",""))[:16].replace("T"," "),
                 "patient":   r.get("patient_name","Unknown"),
                 "pid":       "",
                 "images":    r.get("images_count",1),
                 "rbc":       r.get("total_rbc",0),
                 "wbc":       r.get("total_wbc",0),
                 "rbc_flag":  get_flag("RBC", r.get("total_rbc",0)),
                 "wbc_flag":  get_flag("WBC", r.get("total_wbc",0)),
                 "severity":  r.get("severity","NORMAL"),
                 "score":     r.get("health_score",0),
                 } for r in reports]
    except Exception:
        pass
    st.session_state.sb_loaded = True


# ═══════════════════════════════════════════════════════════════════════════════
#  CSS
# ═══════════════════════════════════════════════════════════════════════════════

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Public+Sans:wght@400;700;900&family=Montserrat:wght@400;600;700;800&family=Open+Sans:wght@400;600&family=Roboto:wght@300;400;500;700&family=Russo+One&display=swap');
*,html,body,[class*="css"]{font-family:'Open Sans',sans-serif;box-sizing:border-box;}
.stApp{background:#F0F4F8!important;}
section[data-testid="stSidebar"]{background:#FFFFFF!important;border-right:1px solid #E2E8F0!important;box-shadow:4px 0 16px rgba(0,0,0,0.06)!important;}
section[data-testid="stSidebar"]>div{padding-top:0!important;}
#MainMenu,footer,header{visibility:hidden;}
.block-container{padding-top:0.5rem!important;padding-left:1.2rem!important;padding-right:1.2rem!important;max-width:100%!important;}
.stButton>button{background:linear-gradient(135deg,#1A56DB,#0EA5E9)!important;color:white!important;border:none!important;border-radius:10px!important;font-family:Roboto,sans-serif!important;font-size:13px!important;font-weight:500!important;padding:10px 16px!important;width:100%!important;box-shadow:0 4px 12px rgba(26,86,219,0.25)!important;transition:all 0.15s!important;}
.stButton>button:hover{transform:translateY(-1px)!important;box-shadow:0 6px 18px rgba(26,86,219,0.35)!important;}
.stDownloadButton>button{background:linear-gradient(135deg,#22C55E,#16A34A)!important;color:white!important;border:none!important;border-radius:10px!important;font-family:Roboto,sans-serif!important;font-size:13px!important;width:100%!important;box-shadow:0 4px 12px rgba(34,197,94,0.2)!important;}
.stTextInput>div>div>input,.stTextArea>div>div>textarea,.stNumberInput>div>div>input{background:#F8FAFC!important;color:#0F172A!important;border:1.5px solid #E2E8F0!important;border-radius:10px!important;font-family:'Open Sans',sans-serif!important;font-size:14px!important;}
.stTextInput>div>div>input:focus,.stTextArea>div>div>textarea:focus{border-color:#1A56DB!important;box-shadow:0 0 0 3px rgba(26,86,219,0.1)!important;background:white!important;}
.stSelectbox>div>div{background:#F8FAFC!important;border:1.5px solid #E2E8F0!important;border-radius:10px!important;font-family:'Open Sans',sans-serif!important;font-size:14px!important;}
label,.stTextInput label,.stTextArea label,.stSelectbox label,.stSlider label{font-family:'Open Sans',sans-serif!important;font-size:12px!important;font-weight:600!important;color:#64748B!important;}
.stSlider>div>div>div{background:linear-gradient(90deg,#1A56DB,#06B6D4)!important;}
.stProgress>div>div{background:linear-gradient(90deg,#1A56DB,#06B6D4)!important;border-radius:4px!important;}
hr{border-color:#EDF2F7!important;}
div[data-testid="stCheckbox"] label{font-size:13px!important;color:#334155!important;}
</style>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
#  UI HELPERS — pure Streamlit, no raw HTML/div mixing
# ═══════════════════════════════════════════════════════════════════════════════

def hero_card(title, subtitle, icon="🔬", stat=None, stat_label=None,
              conf=None):
    """Render hero gradient card using pure st.markdown."""
    conf_pill = ""
    if conf:
        conf_pill = (
            '<span style="background:rgba(255,255,255,0.2);'
            'border:1px solid rgba(255,255,255,0.3);border-radius:8px;'
            'padding:3px 10px;font-family:Roboto,sans-serif;'
            'font-size:11px;font-weight:700;color:white;'
            'margin-left:10px;">CONFIDENCE {0}%</span>'.format(conf))
    stat_html = ""
    if stat is not None:
        stat_html = (
            '<div style="text-align:right;">'
            '<span style="font-family:\'Russo One\',sans-serif;'
            'font-size:40px;color:white;line-height:1;">{0}</span>'
            '<span style="font-family:Roboto,sans-serif;font-size:12px;'
            'color:rgba(255,255,255,0.8);letter-spacing:1px;'
            'text-transform:uppercase;margin-left:8px;">{1}</span>'
            '</div>'.format(stat, stat_label or ""))
    st.markdown(
        '<div style="background:linear-gradient(135deg,#1A56DB 0%,'
        '#0EA5E9 60%,#06B6D4 100%);border-radius:16px;padding:20px 22px;'
        'margin-bottom:16px;box-shadow:0 6px 24px rgba(26,86,219,0.25);">'
        '<div style="display:flex;justify-content:space-between;'
        'align-items:flex-start;flex-wrap:wrap;gap:10px;">'
        '<div>'
        '<div style="font-family:Roboto,sans-serif;font-size:11px;'
        'color:rgba(255,255,255,0.7);letter-spacing:1px;'
        'text-transform:uppercase;margin-bottom:4px;">'
        '{icon} {subtitle}{conf}</div>'
        '<div style="font-family:Montserrat,sans-serif;font-size:18px;'
        'font-weight:700;color:white;margin:0;">{title}</div>'
        '</div>{stat}</div></div>'.format(
            icon=icon, subtitle=subtitle, conf=conf_pill,
            title=title, stat=stat_html),
        unsafe_allow_html=True)


def white_card_start(title=None):
    if title:
        st.markdown(
            '<div style="background:white;border-radius:14px;'
            'padding:16px 18px;box-shadow:0 4px 12px rgba(0,0,0,0.06);'
            'margin-bottom:14px;">'
            '<div style="font-family:Montserrat,sans-serif;font-size:12px;'
            'font-weight:700;color:#0F172A;text-transform:uppercase;'
            'letter-spacing:0.5px;margin-bottom:10px;">{0}</div>'.format(title),
            unsafe_allow_html=True)
    else:
        st.markdown(
            '<div style="background:white;border-radius:14px;'
            'padding:16px 18px;box-shadow:0 4px 12px rgba(0,0,0,0.06);'
            'margin-bottom:14px;">',
            unsafe_allow_html=True)


def white_card_end():
    st.markdown('</div>', unsafe_allow_html=True)


def stat_card_html(icon, label, value, color, badge=None, change=None):
    badge_html_str = ""
    if badge:
        bg = {"NORMAL":"#DCFCE7","HIGH":"#FEE2E2","LOW":"#FEF3C7"}.get(badge,"#F1F5F9")
        tc = {"NORMAL":"#16A34A","HIGH":"#DC2626","LOW":"#D97706"}.get(badge,"#64748B")
        ic = {"NORMAL":"✅","HIGH":"🚨","LOW":"⚠️"}.get(badge,"")
        badge_html_str = (
            '<div style="margin-top:4px;">'
            '<span style="background:{bg};color:{tc};border-radius:6px;'
            'padding:2px 8px;font-family:Roboto,sans-serif;font-size:10px;'
            'font-weight:700;letter-spacing:1px;">{ic} {b}</span>'
            '</div>'.format(bg=bg,tc=tc,ic=ic,b=badge))
    change_str = ""
    if change:
        arrow = "↑" if change > 0 else "↓"
        c_col = "#22C55E" if change > 0 else "#EF4444"
        change_str = (
            '<div style="font-family:Roboto,sans-serif;font-size:11px;'
            'color:{c};margin-top:2px;">{a} {v} from last</div>'.format(
                c=c_col, a=arrow, v=abs(change)))
    return (
        '<div style="background:white;border-radius:14px;padding:14px 12px;'
        'text-align:center;box-shadow:0 4px 12px rgba(0,0,0,0.06);'
        'margin-bottom:10px;">'
        '<div style="font-size:18px;margin-bottom:4px;">{icon}</div>'
        '<div style="font-family:Roboto,sans-serif;font-size:10px;'
        'color:#94A3B8;letter-spacing:2px;text-transform:uppercase;">'
        '{lbl}</div>'
        '<div style="font-family:\'Russo One\',sans-serif;font-size:28px;'
        'color:{c};margin:4px 0;">{val}</div>'
        '{badge}{change}'
        '</div>'.format(icon=icon,lbl=label,c=color,val=value,
                        badge=badge_html_str,change=change_str))


def info_box(text, style="info"):
    styles = {
        "info":    ("background:#EFF6FF;border-left:3px solid #1A56DB;"
                    "color:#1E40AF;"),
        "warn":    ("background:#FFFBEB;border-left:3px solid #F59E0B;"
                    "color:#92400E;"),
        "success": ("background:#F0FDF4;border-left:3px solid #22C55E;"
                    "color:#166534;"),
        "error":   ("background:#FEF2F2;border-left:3px solid #EF4444;"
                    "color:#991B1B;"),
    }
    s = styles.get(style, styles["info"])
    st.markdown(
        '<div style="{s}border-radius:10px;padding:12px 16px;'
        'margin:8px 0;font-family:\'Open Sans\',sans-serif;'
        'font-size:13px;line-height:1.6;">{t}</div>'.format(s=s, t=text),
        unsafe_allow_html=True)


def nav_bar():
    pages = [("Analyse","🔬","ANALYSE"), ("Patient","🧬","PATIENT"),
             ("History","📊","HISTORY"), ("Compare","⚖️","COMPARE"),
             ("Stats","📈","STATS"), ("Batch","📋","BATCH"),
             ("Lab","🏥","LAB"), ("Info","ℹ️","INFO")]
    cols = st.columns(len(pages))
    for col, (page, icon, label) in zip(cols, pages):
        with col:
            active = st.session_state.page == page
            bg     = "linear-gradient(135deg,#1A56DB,#06B6D4)" if active else "white"
            color  = "white" if active else "#94A3B8"
            shadow = "0 4px 12px rgba(26,86,219,0.25)" if active else \
                     "0 2px 8px rgba(0,0,0,0.06)"
            st.markdown(
                '<div style="background:{bg};border-radius:10px;'
                'padding:10px 6px;text-align:center;'
                'box-shadow:{sh};margin-bottom:10px;">'
                '<div style="font-size:16px;">{icon}</div>'
                '<div style="font-family:Roboto,sans-serif;font-size:10px;'
                'font-weight:500;letter-spacing:0.5px;text-transform:uppercase;'
                'color:{color};margin-top:3px;">{label}</div>'
                '</div>'.format(bg=bg,sh=shadow,icon=icon,
                                color=color,label=label),
                unsafe_allow_html=True)
            if st.button(label, key="nav_" + page):
                st.session_state.page = page
                st.rerun()


# ═══════════════════════════════════════════════════════════════════════════════
#  SPLASH SCREEN
# ═══════════════════════════════════════════════════════════════════════════════

def screen_splash():
    st.markdown("""<style>
    section[data-testid="stSidebar"]{display:none!important;}
    .block-container{padding:2rem 1rem!important;}
    </style>""", unsafe_allow_html=True)

    _, col, _ = st.columns([1, 1.6, 1])
    with col:
        st.markdown(
            '<div style="background:linear-gradient(160deg,#1A56DB 0%,'
            '#0EA5E9 50%,#06B6D4 100%);border-radius:24px;padding:60px 40px;'
            'text-align:center;box-shadow:0 20px 60px rgba(26,86,219,0.35);">'
            '<div style="width:90px;height:90px;background:rgba(255,255,255,0.2);'
            'border-radius:50%;display:flex;align-items:center;'
            'justify-content:center;margin:0 auto 24px;font-size:44px;'
            'border:1px solid rgba(255,255,255,0.3);">🔬</div>'
            '<div style="font-family:\'Public Sans\',sans-serif;font-size:30px;'
            'font-weight:900;color:white;letter-spacing:3px;margin:0 0 10px;">CELLSCOPE PRO</div>'
            '<div style="font-family:\'Open Sans\',sans-serif;font-size:14px;'
            'font-style:italic;color:rgba(255,255,255,0.85);margin:0 0 32px;">'
            'Precision Diagnostics at Your Fingertips</div>'
            '<div style="display:flex;gap:8px;justify-content:center;margin:0 0 32px;">'
            '<div style="width:10px;height:10px;border-radius:50%;background:white;"></div>'
            '<div style="width:10px;height:10px;border-radius:50%;background:rgba(255,255,255,0.5);"></div>'
            '<div style="width:10px;height:10px;border-radius:50%;background:rgba(255,255,255,0.3);"></div>'
            '</div>'
            '<div style="font-family:Roboto,sans-serif;font-size:11px;'
            'color:rgba(255,255,255,0.4);">v6.0</div></div>',
            unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🚀  Get Started  →", key="splash_go"):
            st.session_state.app_stage  = "login"
            st.session_state.splash_seen = True
            st.rerun()
        st.markdown(
            '<div style="text-align:center;margin-top:10px;'
            'font-family:Roboto,sans-serif;font-size:12px;color:#94A3B8;">'
            'Hospital Grade Blood Cell Analyzer</div>',
            unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
#  LOGIN SCREEN
# ═══════════════════════════════════════════════════════════════════════════════

def screen_login():
    st.markdown("""<style>
    section[data-testid="stSidebar"]{display:none!important;}
    </style>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    _, col, _ = st.columns([1, 1.2, 1])
    with col:
        # Logo row
        st.markdown(
            '<div style="display:flex;align-items:center;'
            'justify-content:center;gap:10px;margin-bottom:20px;">'
            '<div style="width:36px;height:36px;background:linear-gradient'
            '(135deg,#1A56DB,#06B6D4);border-radius:10px;display:flex;'
            'align-items:center;justify-content:center;font-size:18px;">🔬</div>'
            '<div style="font-family:\'Public Sans\',sans-serif;font-size:18px;'
            'font-weight:900;color:#0F172A;letter-spacing:-0.01em;">'
            'CELLSCOPE PRO</div></div>',
            unsafe_allow_html=True)

        # White card
        st.markdown(
            '<div style="background:white;border-radius:20px;padding:32px 28px;'
            'box-shadow:0 8px 32px rgba(0,0,0,0.08);">',
            unsafe_allow_html=True)
        st.markdown(
            '<div style="font-family:Montserrat,sans-serif;font-size:24px;'
            'font-weight:700;color:#0F172A;margin:0 0 4px;">Welcome Back</div>'
            '<div style="font-family:\'Open Sans\',sans-serif;font-size:14px;'
            'color:#64748B;margin:0 0 20px;">Sign in to your account</div>',
            unsafe_allow_html=True)

        email    = st.text_input("EMAIL OR USERNAME",
                                 placeholder="Enter your credentials",
                                 key="login_email")
        password = st.text_input("PASSWORD", type="password",
                                 placeholder="••••••••",
                                 key="login_pass")

        rc1, rc2 = st.columns([1,1])
        with rc1:
            st.checkbox("Remember me", key="remember_me")
        with rc2:
            st.markdown(
                '<div style="text-align:right;margin-top:6px;">'
                '<span style="color:#1A56DB;font-size:13px;font-weight:600;">'
                'Forgot Password?</span></div>',
                unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        if st.button("Sign In  →", key="login_btn"):
            if email and password:
                if SUPABASE_OK:
                    with st.spinner("Signing in..."):
                        success, result = sign_in(email, password)
                    if success:
                        meta = getattr(result, "user_metadata", {}) or {}
                        name = meta.get("full_name",
                                        email.split("@")[0].replace("."," ").title())
                        st.session_state.user_name  = name
                        st.session_state.user_email = email
                        st.session_state.user_id    = result.id
                        st.session_state.app_stage  = "app"
                        st.session_state.sb_loaded  = False
                        st.rerun()
                    else:
                        info_box("⚠️ " + str(result), "error")
                else:
                    # No Supabase — allow any login for demo
                    st.session_state.user_name  = \
                        email.split("@")[0].replace("."," ").title()
                    st.session_state.user_email = email
                    st.session_state.app_stage  = "app"
                    st.rerun()
            else:
                info_box("⚠️ Please enter your email and password.", "error")

        st.markdown(
            '<div style="display:flex;align-items:center;gap:12px;'
            'margin:14px 0;font-family:Roboto,sans-serif;font-size:12px;'
            'color:#94A3B8;">'
            '<div style="flex:1;height:1px;background:#E2E8F0;"></div>'
            'OR'
            '<div style="flex:1;height:1px;background:#E2E8F0;"></div>'
            '</div>',
            unsafe_allow_html=True)

        if st.button("👤  Continue as Guest", key="guest_btn"):
            st.session_state.user_name  = "Guest User"
            st.session_state.user_email = "guest@cellscope.pro"
            st.session_state.user_role  = "Guest"
            st.session_state.app_stage  = "app"
            st.rerun()

        st.markdown(
            '<div style="text-align:center;margin-top:14px;'
            'font-family:\'Open Sans\',sans-serif;font-size:13px;color:#64748B;">'
            "Don't have an account?</div>",
            unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

        if st.button("Create Account", key="goto_signup"):
            st.session_state.app_stage = "signup"
            st.rerun()

        if st.button("🎬  Launch Demo Mode", key="demo_btn"):
            st.session_state.user_name  = "Demo User"
            st.session_state.user_email = "demo@cellscope.pro"
            st.session_state.user_role  = "Lab Technician"
            for k, v in DEMO_PATIENT.items():
                st.session_state["pt_" + k] = v
            st.session_state.app_stage = "app"
            st.rerun()

        st.markdown(
            '<div style="text-align:center;margin-top:12px;'
            'font-family:Roboto,sans-serif;font-size:11px;color:#94A3B8;">'
            '🔒 YOUR DATA IS ENCRYPTED AND SECURE</div>',
            unsafe_allow_html=True)

        if st.button("← Back to Splash", key="back_splash"):
            st.session_state.app_stage = "splash"
            st.rerun()


# ═══════════════════════════════════════════════════════════════════════════════
#  SIGNUP SCREEN
# ═══════════════════════════════════════════════════════════════════════════════

def screen_signup():
    st.markdown("""<style>
    section[data-testid="stSidebar"]{display:none!important;}
    </style>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    _, col, _ = st.columns([1, 1.2, 1])
    with col:
        st.markdown(
            '<div style="display:flex;align-items:center;'
            'justify-content:center;gap:10px;margin-bottom:20px;">'
            '<div style="width:36px;height:36px;background:linear-gradient'
            '(135deg,#1A56DB,#06B6D4);border-radius:10px;display:flex;'
            'align-items:center;justify-content:center;font-size:18px;">🔬</div>'
            '<div style="font-family:\'Public Sans\',sans-serif;font-size:18px;'
            'font-weight:900;color:#0F172A;">CELLSCOPE PRO</div></div>',
            unsafe_allow_html=True)

        st.markdown(
            '<div style="background:white;border-radius:20px;padding:32px 28px;'
            'box-shadow:0 8px 32px rgba(0,0,0,0.08);">',
            unsafe_allow_html=True)
        st.markdown(
            '<div style="font-family:Montserrat,sans-serif;font-size:24px;'
            'font-weight:700;color:#0F172A;margin:0 0 4px;">Create Account</div>'
            '<div style="font-family:\'Open Sans\',sans-serif;font-size:14px;'
            'color:#64748B;margin:0 0 20px;">Join CellScope Pro</div>',
            unsafe_allow_html=True)

        full_name = st.text_input("FULL NAME",   placeholder="Your full name",   key="su_name")
        su_email  = st.text_input("EMAIL",       placeholder="your@email.com",   key="su_email")
        su_pass   = st.text_input("PASSWORD",    type="password", placeholder="At least 6 characters", key="su_pass")
        su_conf   = st.text_input("CONFIRM PASSWORD", type="password", placeholder="Repeat password", key="su_conf")

        st.markdown(
            '<div style="font-family:\'Open Sans\',sans-serif;font-size:12px;'
            'font-weight:600;color:#64748B;margin:12px 0 8px;">SELECT ROLE</div>',
            unsafe_allow_html=True)
        roles     = ["🔬 Lab Technician", "🩺 Doctor", "🎓 Student"]
        role_sel  = st.radio("role", roles, horizontal=True,
                             label_visibility="collapsed", key="su_role")

        st.markdown("<br>", unsafe_allow_html=True)

        if st.button("Create Account  →", key="signup_btn"):
            if full_name and su_email and su_pass:
                if su_pass != su_conf:
                    info_box("⚠️ Passwords do not match.", "error")
                elif len(su_pass) < 6:
                    info_box("⚠️ Password must be at least 6 characters.", "error")
                else:
                    role_clean = role_sel.split(" ",1)[1] if " " in role_sel else role_sel
                    if SUPABASE_OK:
                        with st.spinner("Creating your account..."):
                            success, msg = sign_up(su_email, su_pass,
                                                   full_name, role_clean)
                        if success:
                            info_box("✅ Account created! Please sign in.", "success")
                            st.session_state.app_stage = "login"
                            st.rerun()
                        else:
                            info_box("⚠️ " + str(msg), "error")
                    else:
                        st.session_state.user_name  = full_name
                        st.session_state.user_email = su_email
                        st.session_state.user_role  = role_clean
                        st.session_state.app_stage  = "app"
                        st.rerun()
            else:
                info_box("⚠️ Please fill in all required fields.", "error")

        st.markdown('</div>', unsafe_allow_html=True)

        if st.button("Sign In Instead", key="goto_login"):
            st.session_state.app_stage = "login"
            st.rerun()


# ═══════════════════════════════════════════════════════════════════════════════
#  SIDEBAR
# ═══════════════════════════════════════════════════════════════════════════════

def build_sidebar():
    with st.sidebar:
        # Logo
        st.markdown(
            '<div style="padding:18px 16px 14px;border-bottom:1px solid #EDF2F7;">'
            '<div style="display:flex;align-items:center;gap:10px;">'
            '<div style="width:34px;height:34px;background:linear-gradient'
            '(135deg,#1A56DB,#06B6D4);border-radius:9px;display:flex;'
            'align-items:center;justify-content:center;font-size:16px;'
            'flex-shrink:0;">🔬</div>'
            '<div>'
            '<div style="font-family:\'Public Sans\',sans-serif;font-size:15px;'
            'font-weight:900;color:#0F172A;letter-spacing:-0.01em;'
            'line-height:1.2;">CELLSCOPE PRO</div>'
            '<div style="font-family:\'Open Sans\',sans-serif;font-size:10px;'
            'font-style:italic;color:#94A3B8;">Precision Diagnostics at Your Fingertips</div>'
            '</div></div></div>',
            unsafe_allow_html=True)

        # Greeting
        hour  = datetime.datetime.now().hour
        greet = ("Good Morning" if hour < 12
                 else "Good Afternoon" if hour < 17 else "Good Evening")
        st.markdown(
            '<div style="padding:12px 16px;border-bottom:1px solid #EDF2F7;">'
            '<div style="font-family:Montserrat,sans-serif;font-size:13px;'
            'font-weight:600;color:#0F172A;">{g}, {n} 👋</div>'
            '<div style="font-family:Roboto,sans-serif;font-size:11px;'
            'color:#94A3B8;">{r}</div></div>'.format(
                g=greet,
                n=st.session_state.user_name or "User",
                r=st.session_state.user_role),
            unsafe_allow_html=True)

        # Upload
        st.markdown(
            '<div style="font-family:Montserrat,sans-serif;font-size:10px;'
            'font-weight:600;color:#94A3B8;letter-spacing:1.5px;'
            'text-transform:uppercase;margin:16px 0 6px;">📂 Upload Images</div>',
            unsafe_allow_html=True)
        uploaded = st.file_uploader(
            "imgs",
            type=["png","jpg","jpeg","bmp","tiff","tif"],
            accept_multiple_files=True,
            label_visibility="collapsed")

        if uploaded:
            for f in uploaded:
                if f.name not in st.session_state.images:
                    pil = Image.open(f).convert("RGB")
                    bgr = cv2.cvtColor(np.array(pil), cv2.COLOR_RGB2BGR)
                    st.session_state.images[f.name] = {
                        "bgr":bgr,"pil":pil,"analysed":False,
                        "cells":[],"rbc":0,"wbc":0,"annot":None,
                        "heatmap":None,"notes":"","confidence":0,
                        "health_score":0,"ai_insight":"","elapsed":0,
                    }
            if st.session_state.selected is None and st.session_state.images:
                st.session_state.selected = list(
                    st.session_state.images.keys())[0]

        if st.session_state.images:
            st.markdown(
                '<div style="font-family:Montserrat,sans-serif;font-size:10px;'
                'font-weight:600;color:#94A3B8;letter-spacing:1.5px;'
                'text-transform:uppercase;margin:16px 0 6px;">🗂️ Image Queue</div>',
                unsafe_allow_html=True)
            for name in list(st.session_state.images.keys()):
                rec   = st.session_state.images[name]
                short = name if len(name) < 22 else name[:19]+"..."
                dot   = "🟢" if rec["analysed"] else "⚪"
                if st.button("{0} {1}".format(dot, short), key="q_"+name):
                    st.session_state.selected = name
                    st.session_state.page = "Analyse"
                    st.rerun()

            an = sum(1 for r in st.session_state.images.values()
                     if r["analysed"])
            tr = sum(r["rbc"] for r in st.session_state.images.values())
            tw = sum(r["wbc"] for r in st.session_state.images.values())
            st.markdown(
                '<div style="display:grid;grid-template-columns:1fr 1fr;'
                'gap:6px;margin:8px 0;">'
                '<div style="background:#F8FAFC;border:1px solid #EDF2F7;'
                'border-radius:10px;padding:8px;text-align:center;">'
                '<div style="font-family:\'Russo One\',sans-serif;font-size:17px;'
                'color:#1A56DB;">{imgs}</div>'
                '<div style="font-family:Roboto,sans-serif;font-size:9px;'
                'color:#94A3B8;letter-spacing:1px;text-transform:uppercase;">'
                '🔬 Images</div></div>'
                '<div style="background:#F8FAFC;border:1px solid #EDF2F7;'
                'border-radius:10px;padding:8px;text-align:center;">'
                '<div style="font-family:\'Russo One\',sans-serif;font-size:17px;'
                'color:#22C55E;">{done}</div>'
                '<div style="font-family:Roboto,sans-serif;font-size:9px;'
                'color:#94A3B8;letter-spacing:1px;text-transform:uppercase;">'
                '✅ Done</div></div>'
                '<div style="background:#F8FAFC;border:1px solid #EDF2F7;'
                'border-radius:10px;padding:8px;text-align:center;">'
                '<div style="font-family:\'Russo One\',sans-serif;font-size:17px;'
                'color:#DC2626;">{rbc}</div>'
                '<div style="font-family:Roboto,sans-serif;font-size:9px;'
                'color:#94A3B8;letter-spacing:1px;text-transform:uppercase;">'
                '🩸 RBC</div></div>'
                '<div style="background:#F8FAFC;border:1px solid #EDF2F7;'
                'border-radius:10px;padding:8px;text-align:center;">'
                '<div style="font-family:\'Russo One\',sans-serif;font-size:17px;'
                'color:#0891B2;">{wbc}</div>'
                '<div style="font-family:Roboto,sans-serif;font-size:9px;'
                'color:#94A3B8;letter-spacing:1px;text-transform:uppercase;">'
                '🦠 WBC</div></div>'
                '</div>'.format(imgs=len(st.session_state.images),
                                done=an, rbc=tr, wbc=tw),
                unsafe_allow_html=True)

            if st.button("🗑️ Clear All Images"):
                st.session_state.images   = {}
                st.session_state.selected = None
                st.session_state.pdf_data = None
                st.rerun()

        # Sensitivity
        st.markdown(
            '<div style="font-family:Montserrat,sans-serif;font-size:10px;'
            'font-weight:600;color:#94A3B8;letter-spacing:1.5px;'
            'text-transform:uppercase;margin:16px 0 6px;">🔭 Sensitivity</div>',
            unsafe_allow_html=True)
        sensitivity = st.slider("s", 0.25, 0.70,
                                st.session_state._sens, 0.05,
                                label_visibility="collapsed")
        st.session_state._sens = sensitivity
        st.markdown(
            '<div style="font-family:\'Russo One\',sans-serif;font-size:14px;'
            'color:#1A56DB;text-align:right;">{0}%</div>'.format(
                int(sensitivity*100)),
            unsafe_allow_html=True)

        st.markdown("---")
        st.markdown(
            '<div style="font-family:\'Open Sans\',sans-serif;font-size:12px;'
            'color:#94A3B8;line-height:2.4;">'
            '<span style="color:#DC2626;">●</span> Red Blood Cell (RBC)<br>'
            '<span style="color:#0891B2;">●</span> White Blood Cell (WBC)'
            '</div>',
            unsafe_allow_html=True)
        st.markdown("---")
        if st.button("🚪 Sign Out"):
            if SUPABASE_OK:
                try:
                    sb_sign_out()
                except Exception:
                    pass
            for k in ["app_stage","user_name","user_email","user_id",
                      "sb_loaded","history","saved_patients","pdf_data"]:
                st.session_state[k] = \
                    "login" if k == "app_stage" else \
                    None    if k in ["user_id","pdf_data"] else \
                    False   if k == "sb_loaded" else \
                    []      if k in ["history","saved_patients"] else ""
            st.rerun()

    return sensitivity


# ═══════════════════════════════════════════════════════════════════════════════
#  BATCH TRIGGER (top level, outside columns)
# ═══════════════════════════════════════════════════════════════════════════════

def handle_batch():
    if not st.session_state.get("_do_batch", False):
        return
    st.session_state["_do_batch"] = False
    pending = [n for n, r in st.session_state.images.items()
               if not r["analysed"]]
    if pending:
        prog    = st.progress(0)
        stat_ph = st.empty()
        sens    = st.session_state.get("_sens", 0.45)
        total_p = len(pending)
        for i, name in enumerate(pending):
            stat_ph.markdown(
                "🔬 Analysing **{0}** of **{1}**: *{2}*".format(
                    i+1, total_p, name))
            run_analyse(name, sens)
            prog.progress(float(i+1) / float(total_p))
        stat_ph.empty()
        prog.empty()
        st.session_state.batch_done = True
    st.rerun()


# ═══════════════════════════════════════════════════════════════════════════════
#  PAGE: ANALYSE
# ═══════════════════════════════════════════════════════════════════════════════

def page_analyse(sensitivity):
    nav_bar()
    handle_batch()

    if not st.session_state.images:
        hero_card("No Image Selected",
                  "Upload images from the sidebar to begin",
                  "🔬")
        return

    selected = st.session_state.selected
    if not selected or selected not in st.session_state.images:
        selected = list(st.session_state.images.keys())[0]
        st.session_state.selected = selected

    rec    = st.session_state.images[selected]
    total  = rec["rbc"] + rec["wbc"] if rec["analysed"] else 0
    conf   = rec.get("confidence", 0)

    # Validate image quality before showing anything
    if not rec["analysed"]:
        valid, reason = check_is_blood_smear(rec["bgr"])
        if not valid:
            info_box("⚠️ " + reason, "warn")

    hero_card(
        title=selected[:40]+"..." if len(selected)>40 else selected,
        subtitle="Analysis Complete" if rec["analysed"] else "Ready to Analyse",
        stat=total,
        stat_label="TOTAL CELLS DETECTED",
        conf=conf if conf > 0 else None)

    # Alert banners
    if rec["analysed"]:
        sev_txt, _ = get_severity(rec["wbc"])
        if sev_txt == "SEVERE":
            info_box("🚨 Critical Alert: Severely elevated WBC. "
                     "Seek immediate medical attention.", "error")
        elif sev_txt == "MODERATE":
            info_box("⚠️ Warning: Moderately elevated WBC. "
                     "Please consult a medical professional.", "warn")

    if st.session_state.batch_done:
        an = sum(1 for r in st.session_state.images.values()
                 if r["analysed"])
        info_box("✅ Batch complete — {0}/{1} images analysed.".format(
            an, len(st.session_state.images)), "success")
        st.session_state.batch_done = False

    # Action buttons
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        if st.button("📊 Analyse", key="ana"):
            with st.spinner("Detecting cells..."):
                run_analyse(selected, sensitivity)
            st.rerun()
    with c2:
        if st.button("⚡ Batch All", key="bat"):
            st.session_state["_do_batch"] = True
            st.rerun()
    with c3:
        if st.button("🖨️ Print View", key="prt"):
            info_box("To print: press Ctrl+P or right-click → Print.", "info")
    with c4:
        adl = [{"name":n,"rbc":r["rbc"],"wbc":r["wbc"],"notes":r["notes"]}
               for n,r in st.session_state.images.items() if r["analysed"]]
        if adl:
            out = io.StringIO()
            w   = csv.writer(out)
            w.writerow(["File","Total","RBC","WBC","RBC Flag","WBC Flag",
                        "Ratio","Severity","Health Score","Timer(s)","Notes"])
            for d in adl:
                img_rec = st.session_state.images.get(d["name"],{})
                sev, _  = get_severity(d["wbc"])
                score   = get_health_score(d["rbc"],d["wbc"])
                w.writerow([d["name"], d["rbc"]+d["wbc"], d["rbc"], d["wbc"],
                            get_flag("RBC",d["rbc"]),get_flag("WBC",d["wbc"]),
                            "{0}:{1}".format(d["rbc"],d["wbc"]),
                            sev, score,
                            img_rec.get("elapsed",""), d["notes"]])
            st.download_button("📤 CSV", data=out.getvalue(),
                               file_name="cellscope_{0}.csv".format(
                                   datetime.date.today()),
                               mime="text/csv", key="csv_dl")

    st.markdown("---")

    # Adjustments
    a1, a2 = st.columns(2)
    with a1:
        brightness = st.slider(
            "BRIGHTNESS: {0}%".format(int(st.session_state.brightness*100)),
            0.3, 2.5, st.session_state.brightness, 0.05, key="bri")
        st.session_state.brightness = brightness
    with a2:
        contrast = st.slider(
            "CONTRAST: {0}%".format(int(st.session_state.contrast*100)),
            0.3, 2.5, st.session_state.contrast, 0.05, key="con")
        st.session_state.contrast = contrast

    st.markdown("")

    img_col, stat_col = st.columns([3, 1])

    with img_col:
        # View mode toggle
        vm1, vm2, vm3 = st.columns(3)
        with vm1:
            st.session_state.show_annot = st.checkbox(
                "✅ Annotations", value=st.session_state.show_annot,
                key="ann_chk")
        with vm2:
            if rec.get("heatmap") is not None:
                st.session_state.show_heatmap = st.checkbox(
                    "🌡️ Heatmap", value=st.session_state.show_heatmap,
                    key="heat_chk")

        # Choose display image
        if st.session_state.show_heatmap and rec.get("heatmap") is not None:
            display = rec["heatmap"]
        elif st.session_state.show_annot and rec.get("annot") is not None:
            display = rec["annot"]
        else:
            display = rec["bgr"]

        pil = Image.fromarray(cv2.cvtColor(display, cv2.COLOR_BGR2RGB))
        if brightness != 1.0:
            pil = ImageEnhance.Brightness(pil).enhance(brightness)
        if contrast != 1.0:
            pil = ImageEnhance.Contrast(pil).enhance(contrast)
        st.image(pil, use_container_width=True)

        # Analysis timer
        if rec["analysed"] and rec.get("elapsed", 0) > 0:
            st.markdown(
                '<div style="font-family:Roboto,sans-serif;font-size:11px;'
                'color:#94A3B8;text-align:right;margin-top:4px;">'
                '⏱️ Analysis completed in {0}s</div>'.format(rec["elapsed"]),
                unsafe_allow_html=True)

        # Downloads
        if rec["annot"] is not None:
            d1, d2 = st.columns(2)
            with d1:
                buf = io.BytesIO()
                pil.save(buf, format="PNG")
                st.download_button("⬇️ Download Annotated",
                                   data=buf.getvalue(),
                                   file_name="annotated_"+selected,
                                   mime="image/png", key="dl_ann")
            with d2:
                obuf = io.BytesIO()
                rec["pil"].save(obuf, format="PNG")
                st.download_button("⬇️ Download Original",
                                   data=obuf.getvalue(),
                                   file_name="original_"+selected,
                                   mime="image/png", key="dl_ori")

    with stat_col:
        rec["notes"] = st.text_area(
            "📝 CLINICAL NOTES", value=rec["notes"],
            placeholder="Add clinical notes...",
            height=80, key="note_ta")

        if not rec["analysed"]:
            info_box("Click <strong>📊 Analyse</strong> to detect "
                     "and count blood cells.", "info")
        else:
            gender   = st.session_state.get("pt_gender","Male")
            age      = st.session_state.get("pt_age",None)
            rbc_flag = get_flag("RBC", rec["rbc"], gender, age)
            wbc_flag = get_flag("WBC", rec["wbc"], gender, age)
            ratio    = "{0}:{1}".format(rec["rbc"],rec["wbc"]) \
                       if rec["wbc"] else "{0}:0".format(rec["rbc"])
            sev_txt, sev_col = get_severity(rec["wbc"])
            score    = rec.get("health_score", 0)
            nr       = get_normal_ranges(gender, age)

            # Get change from last scan if history exists
            last_rbc_change = None
            last_wbc_change = None
            if len(st.session_state.history) > 0:
                last = st.session_state.history[-1]
                last_rbc_change = rec["rbc"] - last.get("rbc", rec["rbc"])
                last_wbc_change = rec["wbc"] - last.get("wbc", rec["wbc"])

            # WBC stat card
            st.markdown(
                stat_card_html("🦠","WBC COUNT",rec["wbc"],"#0891B2",
                               wbc_flag, last_wbc_change),
                unsafe_allow_html=True)
            # RBC stat card
            st.markdown(
                stat_card_html("🩸","RBC COUNT",rec["rbc"],"#DC2626",
                               rbc_flag, last_rbc_change),
                unsafe_allow_html=True)
            # Total
            st.markdown(
                stat_card_html("#️⃣","TOTAL CELLS",rec["rbc"]+rec["wbc"],
                               "#1A56DB"),
                unsafe_allow_html=True)
            # Ratio
            st.markdown(
                stat_card_html("⚖️","RBC : WBC RATIO",ratio,"#F97316"),
                unsafe_allow_html=True)

            # Health gauge
            gauge = health_gauge_svg(score, sev_col)
            st.markdown(
                '<div style="background:white;border-radius:14px;'
                'padding:14px;text-align:center;'
                'box-shadow:0 4px 12px rgba(0,0,0,0.06);margin-bottom:10px;">'
                '<div style="font-family:Roboto,sans-serif;font-size:10px;'
                'color:#94A3B8;letter-spacing:2px;text-transform:uppercase;'
                'margin-bottom:4px;">BLOOD HEALTH SCORE</div>'
                '<div style="display:flex;justify-content:center;">'
                '{gauge}</div></div>'.format(gauge=gauge),
                unsafe_allow_html=True)

            # Metric distribution
            bar = metric_bar_html(rec["wbc"],
                                  nr["WBC"]["min"],
                                  nr["WBC"]["max"])
            st.markdown(
                '<div style="background:white;border-radius:14px;'
                'padding:14px 16px;box-shadow:0 4px 12px rgba(0,0,0,0.06);'
                'margin-bottom:10px;">'
                '<div style="display:flex;justify-content:space-between;">'
                '<div style="font-family:Montserrat,sans-serif;font-size:12px;'
                'font-weight:700;color:#0F172A;">METRIC DISTRIBUTION</div>'
                '<div style="font-family:Roboto,sans-serif;font-size:10px;'
                'color:#94A3B8;">Ref: {lo}-{hi}</div>'
                '</div>{bar}</div>'.format(
                    lo=nr["WBC"]["min"], hi=nr["WBC"]["max"], bar=bar),
                unsafe_allow_html=True)

            # Smart normal range note
            if gender != "Male" or (age and int(age or 30) < 12):
                info_box("ℹ️ Normal ranges adjusted for {0} patient, "
                         "age {1}.".format(gender, age or "adult"), "info")

            # AI Insight
            insight = rec.get("ai_insight","")
            if insight:
                sev_style = {
                    "NORMAL":  ("background:#F0FDF4;border-left:4px solid #22C55E;color:#166534;"),
                    "MILD":    ("background:#FEFCE8;border-left:4px solid #EAB308;color:#713F12;"),
                    "MODERATE":("background:#FFF7ED;border-left:4px solid #F97316;color:#7C2D12;"),
                    "SEVERE":  ("background:#FEF2F2;border-left:4px solid #EF4444;color:#7F1D1D;"),
                }[sev_txt]
                st.markdown(
                    '<div style="{s}border-radius:12px;padding:12px 14px;'
                    'margin-bottom:10px;">'
                    '<div style="font-family:Montserrat,sans-serif;font-size:11px;'
                    'font-weight:700;letter-spacing:1px;margin-bottom:6px;">'
                    '✨ AI INSIGHT</div>'
                    '<div style="font-family:\'Open Sans\',sans-serif;'
                    'font-size:13px;font-style:italic;line-height:1.6;">'
                    '&ldquo;{txt}&rdquo;</div>'
                    '<div style="font-family:Roboto,sans-serif;font-size:10px;'
                    'color:#94A3B8;margin-top:6px;">'
                    'Not a medical diagnosis. Always consult a professional.</div>'
                    '</div>'.format(s=sev_style, txt=insight),
                    unsafe_allow_html=True)

            # Disease hints
            hints = get_disease_hints(rbc_flag, wbc_flag)
            if hints:
                st.markdown(
                    '<div style="background:white;border-radius:14px;'
                    'padding:14px 16px;box-shadow:0 4px 12px rgba(0,0,0,0.06);'
                    'margin-bottom:10px;">'
                    '<div style="font-family:Montserrat,sans-serif;font-size:12px;'
                    'font-weight:700;color:#0F172A;margin-bottom:10px;">'
                    '🏥 CLINICAL INDICATORS</div>',
                    unsafe_allow_html=True)
                for h_name, h_dot, h_desc in hints:
                    st.markdown(
                        '<div style="border-left:3px solid #E2E8F0;'
                        'padding:6px 10px;margin-bottom:8px;">'
                        '<div style="font-family:Montserrat,sans-serif;'
                        'font-size:12px;font-weight:600;color:#0F172A;">'
                        '{dot} {name}</div>'
                        '<div style="font-family:\'Open Sans\',sans-serif;'
                        'font-size:12px;color:#64748B;margin-top:2px;'
                        'line-height:1.5;">{desc}</div>'
                        '</div>'.format(dot=h_dot,name=h_name,desc=h_desc),
                        unsafe_allow_html=True)
                st.markdown(
                    '<div style="font-family:Roboto,sans-serif;font-size:10px;'
                    'color:#94A3B8;margin-top:4px;">'
                    '⚕️ Educational indicators only. Not a medical diagnosis.'
                    '</div></div>',
                    unsafe_allow_html=True)

            # PDF buttons
            st.markdown("---")
            if st.button("📄 Generate PDF Report", key="pdf_gen"):
                ts  = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
                rid = make_report_id()
                ap  = Image.fromarray(
                    cv2.cvtColor(rec["annot"],cv2.COLOR_BGR2RGB)) \
                    if rec["annot"] is not None else None
                st.session_state.pdf_data = generate_pdf(
                    get_patient(), get_lab(),
                    [{"name":selected,"rbc":rec["rbc"],
                      "wbc":rec["wbc"],"notes":rec["notes"]}],
                    rid, ts, ap)
                st.session_state.pdf_name = "report_{0}.pdf".format(rid)
                st.session_state.history.append({
                    "report_id": rid, "timestamp": ts,
                    "patient":  st.session_state.pt_name or "Unknown",
                    "pid":      st.session_state.pt_pid,
                    "images": 1, "rbc": rec["rbc"], "wbc": rec["wbc"],
                    "rbc_flag": rbc_flag, "wbc_flag": wbc_flag,
                    "severity": sev_txt, "score": score,
                })
                # Save report to Supabase
                if SUPABASE_OK and st.session_state.get("user_id"):
                    try:
                        save_report(
                            user_id=st.session_state.user_id,
                            report_id=rid,
                            patient_name=st.session_state.pt_name or "Unknown",
                            total_rbc=rec["rbc"], total_wbc=rec["wbc"],
                            severity=sev_txt, health_score=score,
                            images_count=1)
                    except Exception:
                        pass

            if st.session_state.pdf_data:
                st.download_button("⬇️ Download PDF",
                                   data=st.session_state.pdf_data,
                                   file_name=st.session_state.pdf_name,
                                   mime="application/pdf",
                                   key="pdf_dl_a")
                # Email report
                _render_email_section(label="analyse")

            # Excel export
            if EXCEL_OK:
                adl2 = [{"name":n,"rbc":r["rbc"],"wbc":r["wbc"],
                          "notes":r["notes"]}
                        for n,r in st.session_state.images.items()
                        if r["analysed"]]
                if adl2:
                    xl = generate_excel(adl2, get_patient(),
                                        st.session_state.history)
                    if xl:
                        st.download_button(
                            "📊 Export Excel",
                            data=xl,
                            file_name="cellscope_{0}.xlsx".format(
                                datetime.date.today()),
                            mime="application/vnd.openxmlformats-"
                                 "officedocument.spreadsheetml.sheet",
                            key="xl_dl")

    # Cell log
    if rec["analysed"] and rec["cells"]:
        st.markdown("---")
        lc1, lc2 = st.columns([4,1])
        with lc1:
            st.markdown(
                '<div style="font-family:Montserrat,sans-serif;font-size:13px;'
                'font-weight:700;color:#0F172A;">🧬 CELL DETECTION LOG — '
                '{0} cells</div>'.format(len(rec["cells"])),
                unsafe_allow_html=True)
        with lc2:
            lbl = "See More ▼" if not st.session_state.show_cell_log else "See Less ▲"
            if st.button(lbl, key="tog_log"):
                st.session_state.show_cell_log = not st.session_state.show_cell_log
                st.rerun()

        if st.session_state.show_cell_log:
            rows = ""
            for c in rec["cells"]:
                rc   = "rbc" if c["label"] == "RBC" else "wbc"
                col  = "#DC2626" if rc == "rbc" else "#0891B2"
                rows += (
                    '<tr style="border-bottom:1px solid #F1F5F9;">'
                    '<td style="padding:7px 12px;color:{col};">#{id}</td>'
                    '<td style="padding:7px 12px;color:{col};'
                    'font-weight:600;">{lbl}</td>'
                    '<td style="padding:7px 12px;color:#334155;">{x}</td>'
                    '<td style="padding:7px 12px;color:#334155;">{y}</td>'
                    '<td style="padding:7px 12px;color:#334155;">{area}px²</td>'
                    '</tr>'
                ).format(col=col,id=c["id"],lbl=c["label"],
                         x=c["x"],y=c["y"],area=c["area"])
            st.markdown(
                '<div style="background:white;border-radius:12px;'
                'overflow:hidden;box-shadow:0 4px 12px rgba(0,0,0,0.06);">'
                '<table style="width:100%;border-collapse:collapse;'
                'font-family:Roboto,monospace;font-size:12px;">'
                '<thead><tr style="background:#F8FAFC;">'
                '<th style="padding:9px 12px;text-align:left;color:#64748B;'
                'font-size:10px;font-weight:700;letter-spacing:1.5px;'
                'text-transform:uppercase;border-bottom:1px solid #EDF2F7;">'
                'ID</th>'
                '<th style="padding:9px 12px;text-align:left;color:#64748B;'
                'font-size:10px;font-weight:700;letter-spacing:1.5px;'
                'text-transform:uppercase;border-bottom:1px solid #EDF2F7;">'
                'TYPE</th>'
                '<th style="padding:9px 12px;text-align:left;color:#64748B;'
                'font-size:10px;font-weight:700;letter-spacing:1.5px;'
                'text-transform:uppercase;border-bottom:1px solid #EDF2F7;">'
                'X</th>'
                '<th style="padding:9px 12px;text-align:left;color:#64748B;'
                'font-size:10px;font-weight:700;letter-spacing:1.5px;'
                'text-transform:uppercase;border-bottom:1px solid #EDF2F7;">'
                'Y</th>'
                '<th style="padding:9px 12px;text-align:left;color:#64748B;'
                'font-size:10px;font-weight:700;letter-spacing:1.5px;'
                'text-transform:uppercase;border-bottom:1px solid #EDF2F7;">'
                'AREA</th>'
                '</tr></thead>'
                '<tbody>{rows}</tbody></table></div>'.format(rows=rows),
                unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
#  PAGE: PATIENT
# ═══════════════════════════════════════════════════════════════════════════════

def page_patient():
    nav_bar()
    hero_card("Patient Details",
              "Enter patient information below", "🧬")

    # Quick load
    saved = st.session_state.saved_patients
    if saved:
        names  = ["-- Load from saved patients --"] + \
                 ["{0} ({1})".format(sp.get("name",""),
                                     sp.get("pid","")) for sp in saved]
        choice = st.selectbox("QUICK LOAD", names, key="pt_load")
        if choice != "-- Load from saved patients --":
            idx = names.index(choice) - 1
            sp  = saved[idx]
            field_map = {
                "name":"pt_name","pid":"pt_pid","dob":"pt_dob",
                "age":"pt_age","gender":"pt_gender","mobile":"pt_mobile",
                "email":"pt_email","doctor":"pt_doctor",
                "sample_id":"pt_sid","collection_date":"pt_cdate",
                "notes":"pt_notes",
            }
            for src, dst in field_map.items():
                st.session_state[dst] = sp.get(src,"")
            if not st.session_state.pt_gender:
                st.session_state.pt_gender = "Male"
            st.rerun()

    # Form
    c1, c2 = st.columns(2)
    with c1:
        st.session_state.pt_name   = st.text_input("👤 FULL NAME",   value=st.session_state.pt_name,   placeholder="e.g. Alexander Pierce", key="ip_name")
        st.session_state.pt_pid    = st.text_input("🪪 PATIENT ID",   value=st.session_state.pt_pid,    placeholder="ID-0000",               key="ip_pid")
        st.session_state.pt_dob    = st.text_input("📅 DATE OF BIRTH",value=st.session_state.pt_dob,    placeholder="dd-mm-yyyy",            key="ip_dob")
        st.session_state.pt_mobile = st.text_input("📱 MOBILE",       value=st.session_state.pt_mobile, placeholder="+1...",                 key="ip_mob")
        st.session_state.pt_email  = st.text_input("📧 EMAIL",        value=st.session_state.pt_email,  placeholder="patient@example.com",   key="ip_email")
    with c2:
        st.session_state.pt_age    = st.text_input("🔢 AGE",          value=st.session_state.pt_age,    placeholder="35",                    key="ip_age")
        genders = ["Male","Female","Other"]
        idx_g   = genders.index(st.session_state.pt_gender) \
                  if st.session_state.pt_gender in genders else 0
        st.session_state.pt_gender = st.selectbox("⚧ GENDER", genders, index=idx_g, key="ip_gen")
        st.session_state.pt_doctor = st.text_input("🩺 REFERRING DOCTOR",      value=st.session_state.pt_doctor, placeholder="Dr. Smith",     key="ip_doc")
        st.session_state.pt_cdate  = st.text_input("💉 SAMPLE COLLECTION DATE",value=st.session_state.pt_cdate,  placeholder="dd-mm-yyyy",    key="ip_cdate")
        st.session_state.pt_sid    = st.text_input("🧪 SAMPLE ID",             value=st.session_state.pt_sid,    placeholder="SMP-2024-001",  key="ip_sid")

    st.session_state.pt_notes = st.text_area(
        "📝 CLINICAL NOTES", value=st.session_state.pt_notes,
        placeholder="Enter clinical observations, symptoms, or relevant medical history...",
        height=90, key="ip_notes")

    b1, b2 = st.columns(2)
    with b1:
        if st.button("💾 Save Patient Profile"):
            p = get_patient()
            exists = any(sp.get("pid") == p.get("pid")
                         for sp in st.session_state.saved_patients
                         if p.get("pid"))
            if exists:
                info_box("⚠️ Patient with this ID is already saved.", "warn")
            elif not p.get("name"):
                info_box("⚠️ Please enter a patient name first.", "warn")
            else:
                st.session_state.saved_patients.append(dict(p))
                if SUPABASE_OK and st.session_state.get("user_id"):
                    try:
                        save_patient(st.session_state.user_id, p)
                    except Exception:
                        pass
                info_box("✅ Patient profile saved successfully.", "success")
    with b2:
        if st.button("🗑️ Clear Form"):
            for k in ["pt_name","pt_pid","pt_dob","pt_age","pt_mobile",
                      "pt_email","pt_doctor","pt_notes"]:
                st.session_state[k] = ""
            st.session_state.pt_gender = "Male"
            st.session_state.pt_cdate  = \
                datetime.date.today().strftime("%Y-%m-%d")
            st.session_state.pt_sid = "SMP-001"
            st.rerun()

    # Preview card
    if st.session_state.pt_name:
        st.markdown(
            '<div style="background:linear-gradient(135deg,#EFF6FF,#F0FDFA);'
            'border:1px solid #BAE6FD;border-left:4px solid #1A56DB;'
            'border-radius:14px;padding:16px 18px;margin:10px 0;">'
            '<div style="font-family:\'Public Sans\',sans-serif;font-size:18px;'
            'font-weight:700;letter-spacing:-0.02em;color:#0F172A;'
            'margin:0 0 8px;">{name}</div>'
            '<div style="font-family:\'Open Sans\',sans-serif;font-size:13px;'
            'color:#64748B;line-height:2;">'
            '🪪 {pid} &nbsp;&nbsp; 📅 {dob} &nbsp;&nbsp; '
            '🔢 {age} &nbsp;&nbsp; ⚧ {gender}<br>'
            '📱 {mobile} &nbsp;&nbsp; 📧 {email}<br>'
            '🩺 {doctor} &nbsp;&nbsp; 🧪 {sid} &nbsp;&nbsp; 💉 {cdate}'
            '{notes_row}'
            '</div></div>'.format(
                name=st.session_state.pt_name,
                pid=st.session_state.pt_pid      or "N/A",
                dob=st.session_state.pt_dob      or "N/A",
                age=st.session_state.pt_age      or "N/A",
                gender=st.session_state.pt_gender,
                mobile=st.session_state.pt_mobile or "N/A",
                email=st.session_state.pt_email   or "N/A",
                doctor=st.session_state.pt_doctor or "N/A",
                sid=st.session_state.pt_sid       or "N/A",
                cdate=st.session_state.pt_cdate   or "N/A",
                notes_row="<br>📝 "+st.session_state.pt_notes
                if st.session_state.pt_notes else ""),
            unsafe_allow_html=True)

    st.markdown("---")
    analysed = [{"name":n,"rbc":r["rbc"],"wbc":r["wbc"],"notes":r["notes"]}
                for n,r in st.session_state.images.items() if r["analysed"]]
    if analysed:
        if st.button("📄 Generate Full PDF Report (All Images)",
                     key="full_pdf_gen"):
            ts  = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
            rid = make_report_id()
            st.session_state.pdf_data = generate_pdf(
                get_patient(), get_lab(), analysed, rid, ts, None)
            st.session_state.pdf_name = "full_report_{0}.pdf".format(rid)
            trbc = sum(d["rbc"] for d in analysed)
            twbc = sum(d["wbc"] for d in analysed)
            sev_txt, _ = get_severity(twbc)
            st.session_state.history.append({
                "report_id": rid, "timestamp": ts,
                "patient":   st.session_state.pt_name or "Unknown",
                "pid":       st.session_state.pt_pid,
                "images":    len(analysed),
                "rbc":       trbc, "wbc": twbc,
                "rbc_flag":  get_flag("RBC",trbc),
                "wbc_flag":  get_flag("WBC",twbc),
                "severity":  sev_txt,
                "score":     get_health_score(trbc,twbc),
            })
            if SUPABASE_OK and st.session_state.get("user_id"):
                try:
                    save_report(
                        user_id=st.session_state.user_id,
                        report_id=rid,
                        patient_name=st.session_state.pt_name or "Unknown",
                        total_rbc=trbc, total_wbc=twbc,
                        severity=sev_txt,
                        health_score=get_health_score(trbc,twbc),
                        images_count=len(analysed))
                except Exception:
                    pass
        if st.session_state.pdf_data:
            st.download_button("⬇️ Download Full PDF Report",
                               data=st.session_state.pdf_data,
                               file_name=st.session_state.pdf_name,
                               mime="application/pdf",
                               key="full_pdf_dl")
        # Email report from Patient page
        _render_email_section(label="patient")
    else:
        info_box("⚠️ Analyse at least one image first to generate a report.",
                 "warn")


# ═══════════════════════════════════════════════════════════════════════════════
#  PAGE: HISTORY
# ═══════════════════════════════════════════════════════════════════════════════

def page_history():
    nav_bar()
    hist   = st.session_state.history
    ts_now = datetime.datetime.now().strftime("%H:%M %p")
    hero_card("History",
              "Session started {0} — {1} reports".format(
                  ts_now, len(hist)), "📊")

    if not hist:
        info_box("No history yet — generate a PDF report to create a record.",
                 "info")
        return

    # Dashboard stats
    tr   = sum(h["rbc"]    for h in hist)
    tw   = sum(h["wbc"]    for h in hist)
    ti   = sum(h["images"] for h in hist)
    hwbc = sum(1 for h in hist if h.get("wbc_flag") == "HIGH")
    rbc_mean = round(tr/max(len(hist),1), 2)
    wbc_mean = round(tw/max(len(hist),1), 2)

    d1,d2 = st.columns(2)
    with d1:
        st.markdown(
            stat_card_html("📋","Total Reports", len(hist),"#1A56DB"),
            unsafe_allow_html=True)
        st.markdown(
            stat_card_html("🩸","RBC Mean", rbc_mean,"#DC2626"),
            unsafe_allow_html=True)
    with d2:
        st.markdown(
            stat_card_html("🔬","Total Images", ti,"#475569"),
            unsafe_allow_html=True)
        st.markdown(
            stat_card_html("🦠","WBC Mean", wbc_mean,"#0891B2"),
            unsafe_allow_html=True)

    if hwbc > 0:
        st.markdown(
            '<div style="background:#FEF2F2;border:1px solid #FECACA;'
            'border-radius:12px;padding:14px 18px;'
            'display:flex;justify-content:space-between;'
            'align-items:center;margin-bottom:14px;">'
            '<div>'
            '<div style="font-family:Roboto,sans-serif;font-size:11px;'
            'color:#EF4444;text-transform:uppercase;letter-spacing:1px;'
            'font-weight:700;">High WBC Alerts</div>'
            '<div style="font-family:\'Russo One\',sans-serif;font-size:22px;'
            'color:#DC2626;">{0} Cases</div>'
            '</div>'
            '<div style="font-size:28px;">⚠️</div>'
            '</div>'.format(hwbc),
            unsafe_allow_html=True)

    # Trend chart
    if PLOTLY_OK and len(hist) > 1:
        dates = [h.get("timestamp","")[-5:] for h in hist]
        rbcs  = [h.get("rbc",0)             for h in hist]
        wbcs  = [h.get("wbc",0)             for h in hist]
        fig   = go.Figure()
        fig.add_trace(go.Scatter(
            x=dates, y=rbcs, name="RBC", mode="lines+markers",
            line=dict(color="#DC2626",width=2), marker=dict(size=6),
            fill="tozeroy", fillcolor="rgba(220,38,38,0.06)"))
        fig.add_trace(go.Scatter(
            x=dates, y=wbcs, name="WBC", mode="lines+markers",
            line=dict(color="#0891B2",width=2), marker=dict(size=6),
            fill="tozeroy", fillcolor="rgba(8,145,178,0.06)"))
        fig.update_layout(
            height=200, margin=dict(l=0,r=0,t=10,b=0),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font=dict(family="Open Sans",size=11,color="#94A3B8"),
            legend=dict(orientation="h",yanchor="bottom",y=1.02,
                        xanchor="right",x=1),
            xaxis=dict(showgrid=False,zeroline=False),
            yaxis=dict(showgrid=True,gridcolor="#F1F5F9",zeroline=False))
        st.markdown(
            '<div style="background:white;border-radius:14px;'
            'padding:14px 16px;box-shadow:0 4px 12px rgba(0,0,0,0.06);'
            'margin-bottom:14px;">'
            '<div style="font-family:Montserrat,sans-serif;font-size:13px;'
            'font-weight:700;color:#0F172A;margin-bottom:8px;">'
            '📈 RBC &amp; WBC Trends</div>',
            unsafe_allow_html=True)
        st.plotly_chart(fig, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

    # Export
    out = io.StringIO()
    w   = csv.writer(out)
    w.writerow(["Report ID","Timestamp","Patient","ID","Images",
                "RBC","WBC","Total","RBC Flag","WBC Flag","Severity","Score"])
    for h in hist:
        w.writerow([h.get("report_id",""),h.get("timestamp",""),
                    h.get("patient",""),h.get("pid",""),h.get("images",1),
                    h.get("rbc",0),h.get("wbc",0),
                    h.get("rbc",0)+h.get("wbc",0),
                    h.get("rbc_flag",""),h.get("wbc_flag",""),
                    h.get("severity",""),h.get("score","")])
    e1, e2 = st.columns([1,3])
    with e1:
        st.download_button("📁 Export CSV",
                           data=out.getvalue(),
                           file_name="cellscope_history.csv",
                           mime="text/csv")
    with e2:
        if st.button("🗑️ Clear History"):
            st.session_state.history = []
            st.rerun()

    st.markdown("")

    # History cards
    for h in reversed(hist):
        total    = h.get("rbc",0) + h.get("wbc",0)
        rf       = h.get("rbc_flag","—")
        wf       = h.get("wbc_flag","—")
        sev      = h.get("severity","—")
        sc       = h.get("score","—")
        sev_col  = {"NORMAL":"#22C55E","MILD":"#EAB308",
                    "MODERATE":"#F97316","SEVERE":"#EF4444"}.get(sev,"#94A3B8")
        sev_dot  = {"NORMAL":"🟢","MILD":"🟡",
                    "MODERATE":"🟠","SEVERE":"🔴"}.get(sev,"⚪")
        rf_bg    = {"NORMAL":"#DCFCE7","HIGH":"#FEE2E2","LOW":"#FEF3C7"}.get(rf,"#F1F5F9")
        rf_tc    = {"NORMAL":"#16A34A","HIGH":"#DC2626","LOW":"#D97706"}.get(rf,"#64748B")
        wf_bg    = {"NORMAL":"#DCFCE7","HIGH":"#FEE2E2","LOW":"#FEF3C7"}.get(wf,"#F1F5F9")
        wf_tc    = {"NORMAL":"#16A34A","HIGH":"#DC2626","LOW":"#D97706"}.get(wf,"#64748B")
        st.markdown(
            '<div style="background:white;border-radius:14px;padding:14px 18px;'
            'box-shadow:0 4px 12px rgba(0,0,0,0.06);margin-bottom:10px;'
            'display:flex;justify-content:space-between;align-items:flex-start;">'
            '<div>'
            '<div style="font-family:Montserrat,sans-serif;font-size:14px;'
            'font-weight:600;color:#0F172A;margin:0 0 4px;">{pat}</div>'
            '<div style="font-family:\'Open Sans\',sans-serif;font-size:12px;'
            'color:#94A3B8;line-height:1.8;">'
            '🕐 {ts}<br>🏷️ {rid} &nbsp;|&nbsp; 🔬 {imgs} image(s)'
            '</div></div>'
            '<div style="text-align:right;">'
            '<div style="font-family:\'Open Sans\',sans-serif;font-size:12px;'
            'line-height:2.2;">'
            '🩸 <strong style="color:#DC2626;">{rbc}</strong> '
            '<span style="background:{rfbg};color:{rftc};border-radius:5px;'
            'padding:1px 7px;font-size:10px;font-weight:700;">{rf}</span><br>'
            '🦠 <strong style="color:#0891B2;">{wbc}</strong> '
            '<span style="background:{wfbg};color:{wftc};border-radius:5px;'
            'padding:1px 7px;font-size:10px;font-weight:700;">{wf}</span><br>'
            '<strong style="color:#1A56DB;font-size:14px;">Total {tot}</strong> '
            '{dot} <span style="color:{sc};font-size:11px;font-weight:700;">'
            '{sev}</span>'
            '</div></div></div>'.format(
                pat=h.get("patient","Unknown"),
                ts=h.get("timestamp",""),
                rid=h.get("report_id","—"),
                imgs=h.get("images",1),
                rbc=h.get("rbc",0), rf=rf, rfbg=rf_bg, rftc=rf_tc,
                wbc=h.get("wbc",0), wf=wf, wfbg=wf_bg, wftc=wf_tc,
                tot=total, dot=sev_dot, sc=sev_col, sev=sev),
            unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
#  PAGE: LAB
# ═══════════════════════════════════════════════════════════════════════════════

def page_lab():
    nav_bar()
    hero_card("Lab Settings",
              "Configure your laboratory information", "🏥")
    info_box("ℹ️ <strong>Laboratory Information</strong> — These details "
             "appear on all generated PDF reports.", "info")

    c1, c2 = st.columns(2)
    with c1:
        st.session_state.lab_name  = st.text_input("🏥 LAB / HOSPITAL NAME",  value=st.session_state.lab_name,  placeholder="Enter laboratory name",             key="l_name")
        st.session_state.lab_addr  = st.text_input("📍 ADDRESS",              value=st.session_state.lab_addr,  placeholder="123 Medical Center Drive, City",    key="l_addr")
        st.session_state.lab_phone = st.text_input("☎️ PHONE NUMBER",         value=st.session_state.lab_phone, placeholder="+1 (555) 000-0000",                 key="l_phone")
    with c2:
        st.session_state.lab_acc   = st.text_input("🔖 ACCREDITATION NUMBER", value=st.session_state.lab_acc,   placeholder="ACC-2024-XXXXX",                    key="l_acc")
        st.session_state.lab_tech  = st.text_input("👨‍🔬 LAB TECHNICIAN NAME", value=st.session_state.lab_tech,  placeholder="Enter technician name",             key="l_tech")

    if st.button("💾 Save Lab Settings"):
        if SUPABASE_OK and st.session_state.get("user_id"):
            try:
                save_lab_settings(st.session_state.user_id, get_lab())
            except Exception:
                pass
        info_box("✅ Lab settings saved. All future PDF reports will use "
                 "these details.", "success")

    st.markdown(
        '<div style="background:linear-gradient(135deg,#EFF6FF,#F0FDFA);'
        'border:1px solid #BAE6FD;border-left:4px solid #1A56DB;'
        'border-radius:14px;padding:16px 18px;margin:12px 0;">'
        '<div style="font-family:\'Public Sans\',sans-serif;font-size:18px;'
        'font-weight:900;color:#0F172A;margin:0 0 8px;">{name}</div>'
        '<div style="font-family:\'Open Sans\',sans-serif;font-size:13px;'
        'color:#64748B;line-height:2;">'
        '📍 {addr}<br>☎️ {phone}<br>🔖 {acc}<br>👨‍🔬 {tech}'
        '</div></div>'.format(
            name=st.session_state.lab_name,
            addr=st.session_state.lab_addr,
            phone=st.session_state.lab_phone,
            acc=st.session_state.lab_acc,
            tech=st.session_state.lab_tech),
        unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
#  PAGE: INFO
# ═══════════════════════════════════════════════════════════════════════════════

def page_info():
    nav_bar()
    st.markdown(
        '<div style="background:linear-gradient(135deg,#1A56DB 0%,'
        '#0EA5E9 60%,#06B6D4 100%);border-radius:16px;padding:20px 22px;'
        'margin-bottom:16px;box-shadow:0 6px 24px rgba(26,86,219,0.25);">'
        '<div style="font-family:Roboto,sans-serif;font-size:11px;'
        'color:rgba(255,255,255,0.7);letter-spacing:2px;'
        'text-transform:uppercase;margin-bottom:4px;">SYSTEM OVERVIEW</div>'
        '<div style="font-family:Montserrat,sans-serif;font-size:18px;'
        'font-weight:700;color:white;">'
        'Info &amp; Help | CellScope Pro v6.0</div>'
        '<div style="font-family:\'Open Sans\',sans-serif;font-size:13px;'
        'color:rgba(255,255,255,0.8);margin-top:4px;">'
        'Precision diagnostic tools for next-generation clinical '
        'microscopy and hematology analysis.</div>'
        '</div>',
        unsafe_allow_html=True)

    info_box("⚠️ <strong>MEDICAL DISCLAIMER</strong> — This software is "
             "intended for professional laboratory use only. Automated "
             "results must be validated by a certified pathologist before "
             "clinical decisions are made. Not for patient self-diagnosis.",
             "warn")

    # About
    st.markdown(
        '<div style="background:white;border-radius:14px;padding:16px 18px;'
        'box-shadow:0 4px 12px rgba(0,0,0,0.06);margin-bottom:14px;">'
        '<div style="display:flex;justify-content:space-between;'
        'align-items:flex-start;gap:16px;">'
        '<div style="flex:1;">'
        '<div style="font-family:Montserrat,sans-serif;font-size:13px;'
        'font-weight:700;color:#0F172A;letter-spacing:1px;'
        'text-transform:uppercase;margin-bottom:12px;">ABOUT CELLSCOPE PRO</div>'
        '<div style="font-family:\'Open Sans\',sans-serif;font-size:14px;'
        'color:#334155;line-height:1.8;">'
        '<strong>Version:</strong> 7.0.0 (Build 2026.05)<br>'
        '<strong>Detection:</strong> Watershed segmentation + CLAHE enhancement<br>'
        '<strong>Classification:</strong> Cell size, brightness, saturation heuristics<br>'
        '<strong>New in v7:</strong> Before/After Compare, Email Reports, Stats Dashboard, Batch PDF, QR Codes<br>'
        '<strong>Formats:</strong> PNG, JPG, BMP, TIFF<br>'
        '<strong>Output:</strong> Annotated images, Heatmaps, PDF reports, CSV, Excel, QR Codes'
        '</div></div>'
        '<div style="width:60px;height:60px;background:linear-gradient'
        '(135deg,#EFF6FF,#F0FDFA);border:1px solid #BAE6FD;'
        'border-radius:14px;display:flex;align-items:center;'
        'justify-content:center;font-size:28px;flex-shrink:0;">🔬</div>'
        '</div></div>',
        unsafe_allow_html=True)

    # Developer card
    st.markdown(
        '<div style="background:linear-gradient(135deg,#EFF6FF,#F0FDFA);'
        'border:1px solid #BAE6FD;border-left:4px solid #1A56DB;'
        'border-radius:14px;padding:18px 20px;margin:10px 0;">'
        '<div style="font-family:Montserrat,sans-serif;font-size:11px;'
        'font-weight:700;color:#1A56DB;letter-spacing:1px;'
        'text-transform:uppercase;margin-bottom:8px;">👨‍💻 About the Developer</div>'
        '<div style="font-family:Montserrat,sans-serif;font-size:16px;'
        'font-weight:700;color:#0F172A;margin:0 0 4px;">{name}</div>'
        '<div style="font-family:\'Open Sans\',sans-serif;font-size:13px;'
        'color:#475569;line-height:1.8;margin-bottom:10px;">{title}</div>'
        '<div>'
        '{pills}'
        '</div></div>'.format(
            name=st.session_state.dev_name,
            title=st.session_state.dev_title,
            pills="".join([
                '<span style="background:#EFF6FF;color:#1A56DB;'
                'border:1px solid #BFDBFE;border-radius:20px;'
                'padding:3px 12px;font-family:Roboto,sans-serif;'
                'font-size:11px;font-weight:500;margin:3px 3px 3px 0;'
                'display:inline-block;">{0}</span>'.format(t)
                for t in ["Python","Streamlit","OpenCV","ReportLab",
                          "NumPy","scikit-image","Plotly","Supabase",
                          "PIL","openpyxl"]
            ])),
        unsafe_allow_html=True)

    # Steps
    st.markdown(
        '<div style="font-family:Montserrat,sans-serif;font-size:13px;'
        'font-weight:700;color:#0F172A;letter-spacing:1px;'
        'text-transform:uppercase;margin:20px 0 12px;">HOW TO USE CELLSCOPE PRO</div>',
        unsafe_allow_html=True)

    steps = [
        ("🏥","Lab Settings","Configure your laboratory name, address and accreditation. These appear on all PDF reports."),
        ("🧬","Patient Details","Enter patient info including name, ID, age, gender and contact details. Save profiles for quick access."),
        ("📂","Upload Images","Upload blood smear images. PNG, JPG, BMP, TIFF supported. Batch upload supported."),
        ("📊","Analyse","Run AI-powered analysis. Adjust sensitivity for best results. Use Batch All for multiple images."),
        ("🌡️","View Heatmap","Toggle the Heatmap view to see cell density distribution across the image."),
        ("📈","Review Results","Check RBC/WBC counts, health score, metric distribution, severity and AI insight."),
        ("🏥","Disease Hints","Review clinical indicators showing possible conditions based on cell count patterns."),
        ("📄","Generate Report","Create professional PDF with patient details, results, annotated image and disease hints."),
        ("📁","Export Data","Export results as CSV or formatted Excel with charts and colour coding."),
    ]
    for i, (icon, title, desc) in enumerate(steps):
        st.markdown(
            '<div style="background:white;border-radius:12px;'
            'padding:14px 16px;box-shadow:0 4px 12px rgba(0,0,0,0.06);'
            'margin-bottom:8px;display:flex;align-items:flex-start;gap:14px;">'
            '<div style="background:linear-gradient(135deg,#1A56DB,#06B6D4);'
            'color:white;font-family:\'Russo One\',sans-serif;font-size:13px;'
            'width:30px;height:30px;border-radius:50%;display:flex;'
            'align-items:center;justify-content:center;flex-shrink:0;">'
            '{n}</div>'
            '<div>'
            '<div style="font-family:Montserrat,sans-serif;font-size:13px;'
            'font-weight:600;color:#0F172A;margin:0 0 3px;">{i} {t}</div>'
            '<div style="font-family:\'Open Sans\',sans-serif;font-size:12px;'
            'color:#64748B;margin:0;line-height:1.5;">{d}</div>'
            '</div></div>'.format(n=i+1,i=icon,t=title,d=desc),
            unsafe_allow_html=True)

    # Mobile access
    info_box(
        "📱 <strong>Mobile Access</strong><br>"
        "✅ Open Command Prompt → type <strong>ipconfig</strong><br>"
        "✅ Note your IPv4 Address e.g. <strong>192.168.1.5</strong><br>"
        "✅ Run: <strong>streamlit run cellscope_pro_v2.py "
        "--server.address 0.0.0.0</strong><br>"
        "✅ On phone browser: <strong>192.168.1.5:8501</strong><br>"
        "Both devices must be on the same WiFi network.",
        "info")

    st.markdown(
        '<div style="text-align:center;margin-top:20px;'
        'font-family:\'Open Sans\',sans-serif;font-size:12px;'
        'color:#94A3B8;line-height:2;">'
        'CellScope Pro · Hospital Grade Blood Cell Analyzer<br>'
        'For professional laboratory use only · © 2026 {name}'
        '</div>'.format(name=st.session_state.dev_name),
        unsafe_allow_html=True)




# ═══════════════════════════════════════════════════════════════════════════════
#  QR CODE HELPER
# ═══════════════════════════════════════════════════════════════════════════════

def make_qr_code(data_str):
    """Generate QR code image from string, return as PIL Image."""
    if not QR_OK:
        return None
    qr = qrcode.QRCode(version=1, error_correction=qrcode.constants.ERROR_CORRECT_H,
                       box_size=6, border=2)
    qr.add_data(data_str)
    qr.make(fit=True)
    img = qr.make_image(fill_color="#1A56DB", back_color="white")
    return img.convert("RGB")


def pdf_to_b64(pdf_bytes):
    return base64.b64encode(pdf_bytes).decode()


# ═══════════════════════════════════════════════════════════════════════════════
#  EMAIL HELPER
# ═══════════════════════════════════════════════════════════════════════════════

def send_report_email(smtp_host, smtp_port, sender_email, sender_pass,
                      recipient_email, patient_name, report_id,
                      pdf_bytes, use_tls=True):
    """Send PDF report via email. Returns (success, message)."""
    try:
        msg = MIMEMultipart()
        msg["From"]    = sender_email
        msg["To"]      = recipient_email
        msg["Subject"] = "CellScope Pro — Blood Analysis Report [{rid}]".format(rid=report_id)

        body = (
            "Dear {name},\n\n"
            "Please find attached your blood cell analysis report generated by CellScope Pro.\n\n"
            "Report ID : {rid}\n"
            "Date      : {date}\n\n"
            "⚠️ DISCLAIMER: This report is for professional review only and is NOT a medical "
            "diagnosis. Please consult a qualified medical professional.\n\n"
            "Regards,\nCellScope Pro Diagnostic System"
        ).format(name=patient_name, rid=report_id,
                 date=datetime.datetime.now().strftime("%Y-%m-%d %H:%M"))
        msg.attach(MIMEText(body, "plain"))

        # Attach PDF
        part = MIMEBase("application", "octet-stream")
        part.set_payload(pdf_bytes)
        encoders.encode_base64(part)
        part.add_header("Content-Disposition",
                        'attachment; filename="Report_{rid}.pdf"'.format(rid=report_id))
        msg.attach(part)

        server = smtplib.SMTP(smtp_host, int(smtp_port), timeout=10)
        if use_tls:
            server.starttls()
        server.login(sender_email, sender_pass)
        server.sendmail(sender_email, recipient_email, msg.as_string())
        server.quit()
        return True, "✅ Email sent successfully to {0}".format(recipient_email)
    except Exception as e:
        return False, "❌ Email failed: {0}".format(str(e))


# ═══════════════════════════════════════════════════════════════════════════════
#  PAGE: COMPARE (Before vs After)
# ═══════════════════════════════════════════════════════════════════════════════

def page_compare():
    nav_bar()
    hero_card("Comparative Analysis",
              "Side-by-side Before vs After blood smear comparison", "⚖️")

    analysed_imgs = {n: r for n, r in st.session_state.images.items()
                     if r["analysed"]}

    if len(analysed_imgs) < 2:
        info_box("⚠️ You need at least <strong>2 analysed images</strong> to use "
                 "Comparative Analysis. Go to the <strong>Analyse</strong> page and "
                 "analyse your images first.", "warn")
        return

    names = list(analysed_imgs.keys())

    col_sel1, col_sel2 = st.columns(2)
    with col_sel1:
        st.markdown('<div style="font-family:Montserrat,sans-serif;font-size:11px;'
                    'font-weight:700;color:#64748B;letter-spacing:1px;'
                    'text-transform:uppercase;margin-bottom:6px;">🔵 BEFORE / SAMPLE A</div>',
                    unsafe_allow_html=True)
        sel_a = st.selectbox("Image A", names, key="cmp_a", label_visibility="collapsed")
    with col_sel2:
        st.markdown('<div style="font-family:Montserrat,sans-serif;font-size:11px;'
                    'font-weight:700;color:#64748B;letter-spacing:1px;'
                    'text-transform:uppercase;margin-bottom:6px;">🟢 AFTER / SAMPLE B</div>',
                    unsafe_allow_html=True)
        default_b = names[1] if len(names) > 1 else names[0]
        idx_b = names.index(default_b)
        sel_b = st.selectbox("Image B", names, index=idx_b, key="cmp_b",
                             label_visibility="collapsed")

    rec_a = analysed_imgs[sel_a]
    rec_b = analysed_imgs[sel_b]
    gender = st.session_state.get("pt_gender", "Male")
    age    = st.session_state.get("pt_age", None)

    st.markdown("---")

    # Side-by-side images
    img_a_col, div_col, img_b_col = st.columns([5, 0.3, 5])
    with img_a_col:
        st.markdown('<div style="background:linear-gradient(135deg,#EFF6FF,white);'
                    'border:2px solid #BFDBFE;border-radius:14px;padding:12px;'
                    'text-align:center;margin-bottom:8px;">'
                    '<div style="font-family:Montserrat,sans-serif;font-size:12px;'
                    'font-weight:700;color:#1A56DB;margin-bottom:8px;">'
                    '🔵 BEFORE — {name}</div></div>'.format(
                        name=sel_a[:30]+"..." if len(sel_a) > 30 else sel_a),
                    unsafe_allow_html=True)
        disp_a = rec_a["annot"] if rec_a.get("annot") is not None else rec_a["bgr"]
        st.image(Image.fromarray(cv2.cvtColor(disp_a, cv2.COLOR_BGR2RGB)),
                 use_container_width=True)

    with div_col:
        st.markdown('<div style="display:flex;align-items:center;'
                    'justify-content:center;height:100%;'
                    'font-size:28px;padding-top:60px;">⟺</div>',
                    unsafe_allow_html=True)

    with img_b_col:
        st.markdown('<div style="background:linear-gradient(135deg,#F0FDF4,white);'
                    'border:2px solid #BBF7D0;border-radius:14px;padding:12px;'
                    'text-align:center;margin-bottom:8px;">'
                    '<div style="font-family:Montserrat,sans-serif;font-size:12px;'
                    'font-weight:700;color:#16A34A;margin-bottom:8px;">'
                    '🟢 AFTER — {name}</div></div>'.format(
                        name=sel_b[:30]+"..." if len(sel_b) > 30 else sel_b),
                    unsafe_allow_html=True)
        disp_b = rec_b["annot"] if rec_b.get("annot") is not None else rec_b["bgr"]
        st.image(Image.fromarray(cv2.cvtColor(disp_b, cv2.COLOR_BGR2RGB)),
                 use_container_width=True)

    st.markdown("---")

    # Metric comparison table
    st.markdown('<div style="font-family:Montserrat,sans-serif;font-size:13px;'
                'font-weight:700;color:#0F172A;text-transform:uppercase;'
                'letter-spacing:1px;margin-bottom:12px;">📊 Metric Comparison</div>',
                unsafe_allow_html=True)

    def delta_html(a_val, b_val, lower_better=False):
        diff = b_val - a_val
        if diff == 0:
            return '<span style="color:#94A3B8;">→ No change</span>'
        arrow  = "↑" if diff > 0 else "↓"
        if lower_better:
            color = "#22C55E" if diff < 0 else "#EF4444"
        else:
            color = "#22C55E" if diff > 0 else "#EF4444"
        return '<span style="color:{c};font-weight:700;">{a} {d:+d}</span>'.format(
            c=color, a=arrow, d=diff)

    rbc_a_f = get_flag("RBC", rec_a["rbc"], gender, age)
    wbc_a_f = get_flag("WBC", rec_a["wbc"], gender, age)
    rbc_b_f = get_flag("RBC", rec_b["rbc"], gender, age)
    wbc_b_f = get_flag("WBC", rec_b["wbc"], gender, age)
    sev_a, col_a = get_severity(rec_a["wbc"])
    sev_b, col_b = get_severity(rec_b["wbc"])
    score_a = rec_a.get("health_score", 0)
    score_b = rec_b.get("health_score", 0)

    metrics = [
        ("🩸 RBC Count",    rec_a["rbc"],        rec_b["rbc"],        rbc_a_f,  rbc_b_f,  False),
        ("🦠 WBC Count",    rec_a["wbc"],        rec_b["wbc"],        wbc_a_f,  wbc_b_f,  True),
        ("#️⃣ Total Cells",  rec_a["rbc"]+rec_a["wbc"], rec_b["rbc"]+rec_b["wbc"], "—", "—", False),
        ("💯 Health Score", score_a,             score_b,             sev_a,    sev_b,    False),
        ("⏱️ Analysis Time",rec_a.get("elapsed",0), rec_b.get("elapsed",0), "—", "—", False),
        ("🎯 Confidence",   rec_a.get("confidence",0), rec_b.get("confidence",0), "—", "—", False),
    ]

    flag_col = {"NORMAL":"#22C55E","HIGH":"#EF4444","LOW":"#F59E0B","—":"#94A3B8"}

    rows_html = ""
    for metric, va, vb, fa, fb, lb in metrics:
        fa_c = flag_col.get(fa, "#94A3B8")
        fb_c = flag_col.get(fb, "#94A3B8")
        d_html = delta_html(va, vb, lb)
        rows_html += (
            '<tr style="border-bottom:1px solid #F1F5F9;">'
            '<td style="padding:10px 14px;font-family:Roboto,sans-serif;'
            'font-size:13px;font-weight:600;color:#334155;">{m}</td>'
            '<td style="padding:10px 14px;text-align:center;'
            'font-family:Russo One,sans-serif;font-size:16px;color:#1A56DB;">{va}'
            '<br><span style="font-family:Roboto,sans-serif;font-size:10px;'
            'font-weight:700;color:{fac};">{fa}</span></td>'
            '<td style="padding:10px 14px;text-align:center;'
            'font-family:Russo One,sans-serif;font-size:16px;color:#16A34A;">{vb}'
            '<br><span style="font-family:Roboto,sans-serif;font-size:10px;'
            'font-weight:700;color:{fbc};">{fb}</span></td>'
            '<td style="padding:10px 14px;text-align:center;">{delta}</td>'
            '</tr>'
        ).format(m=metric, va=va, fa=fa, fac=fa_c, vb=vb, fb=fb, fbc=fb_c, delta=d_html)

    st.markdown(
        '<div style="background:white;border-radius:14px;overflow:hidden;'
        'box-shadow:0 4px 12px rgba(0,0,0,0.06);margin-bottom:16px;">'
        '<table style="width:100%;border-collapse:collapse;">'
        '<thead><tr style="background:#F8FAFC;">'
        '<th style="padding:10px 14px;text-align:left;font-family:Montserrat,sans-serif;'
        'font-size:10px;font-weight:700;color:#64748B;text-transform:uppercase;'
        'letter-spacing:1px;border-bottom:1px solid #EDF2F7;">METRIC</th>'
        '<th style="padding:10px 14px;text-align:center;font-family:Montserrat,sans-serif;'
        'font-size:10px;font-weight:700;color:#1A56DB;text-transform:uppercase;'
        'letter-spacing:1px;border-bottom:1px solid #EDF2F7;">🔵 BEFORE</th>'
        '<th style="padding:10px 14px;text-align:center;font-family:Montserrat,sans-serif;'
        'font-size:10px;font-weight:700;color:#16A34A;text-transform:uppercase;'
        'letter-spacing:1px;border-bottom:1px solid #EDF2F7;">🟢 AFTER</th>'
        '<th style="padding:10px 14px;text-align:center;font-family:Montserrat,sans-serif;'
        'font-size:10px;font-weight:700;color:#94A3B8;text-transform:uppercase;'
        'letter-spacing:1px;border-bottom:1px solid #EDF2F7;">CHANGE</th>'
        '</tr></thead>'
        '<tbody>{rows}</tbody></table></div>'.format(rows=rows_html),
        unsafe_allow_html=True)

    # Outcome summary
    score_diff = score_b - score_a
    if score_diff > 10:
        info_box("🎉 <strong>Improvement Detected!</strong> Health score improved by "
                 "+{0} points from Before to After.".format(score_diff), "success")
    elif score_diff < -10:
        info_box("⚠️ <strong>Decline Detected.</strong> Health score dropped by "
                 "{0} points. Consider consulting a medical professional.".format(abs(score_diff)), "warn")
    else:
        info_box("ℹ️ Results are broadly similar between the two samples.", "info")

    # Side-by-side disease hints
    st.markdown("---")
    st.markdown('<div style="font-family:Montserrat,sans-serif;font-size:13px;'
                'font-weight:700;color:#0F172A;text-transform:uppercase;'
                'letter-spacing:1px;margin-bottom:12px;">🏥 Clinical Indicator Comparison</div>',
                unsafe_allow_html=True)
    hi_a_col, hi_b_col = st.columns(2)
    hints_a = get_disease_hints(rbc_a_f, wbc_a_f)
    hints_b = get_disease_hints(rbc_b_f, wbc_b_f)
    with hi_a_col:
        st.markdown('<div style="font-family:Montserrat,sans-serif;font-size:11px;'
                    'font-weight:700;color:#1A56DB;margin-bottom:8px;">🔵 BEFORE</div>',
                    unsafe_allow_html=True)
        for h_name, h_dot, h_desc in hints_a:
            st.markdown('<div style="background:#EFF6FF;border-left:3px solid #1A56DB;'
                        'border-radius:8px;padding:8px 12px;margin-bottom:6px;">'
                        '<div style="font-family:Montserrat,sans-serif;font-size:12px;'
                        'font-weight:600;color:#0F172A;">{dot} {name}</div>'
                        '<div style="font-family:Open Sans,sans-serif;font-size:11px;'
                        'color:#64748B;margin-top:3px;">{desc}</div>'
                        '</div>'.format(dot=h_dot, name=h_name, desc=h_desc),
                        unsafe_allow_html=True)
    with hi_b_col:
        st.markdown('<div style="font-family:Montserrat,sans-serif;font-size:11px;'
                    'font-weight:700;color:#16A34A;margin-bottom:8px;">🟢 AFTER</div>',
                    unsafe_allow_html=True)
        for h_name, h_dot, h_desc in hints_b:
            st.markdown('<div style="background:#F0FDF4;border-left:3px solid #22C55E;'
                        'border-radius:8px;padding:8px 12px;margin-bottom:6px;">'
                        '<div style="font-family:Montserrat,sans-serif;font-size:12px;'
                        'font-weight:600;color:#0F172A;">{dot} {name}</div>'
                        '<div style="font-family:Open Sans,sans-serif;font-size:11px;'
                        'color:#64748B;margin-top:3px;">{desc}</div>'
                        '</div>'.format(dot=h_dot, name=h_name, desc=h_desc),
                        unsafe_allow_html=True)

    # Email this comparison report
    st.markdown("---")
    st.markdown('<div style="font-family:Montserrat,sans-serif;font-size:13px;'
                'font-weight:700;color:#0F172A;text-transform:uppercase;'
                'letter-spacing:1px;margin-bottom:10px;">📧 Email Comparison Report</div>',
                unsafe_allow_html=True)
    _render_email_section(label="compare")


# ═══════════════════════════════════════════════════════════════════════════════
#  EMAIL SECTION (reusable widget)
# ═══════════════════════════════════════════════════════════════════════════════

def _render_email_section(label="main"):
    """Renders SMTP email form. label used to make keys unique."""
    with st.expander("📧 Configure & Send Email", expanded=False):
        info_box("ℹ️ Uses your SMTP credentials to send the PDF report directly "
                 "to the patient's email. Works with Gmail, Outlook, etc.", "info")
        ec1, ec2 = st.columns(2)
        with ec1:
            smtp_host  = st.text_input("SMTP Host",  value="smtp.gmail.com",
                                       key="smtp_host_"+label)
            smtp_port  = st.text_input("SMTP Port",  value="587",
                                       key="smtp_port_"+label)
            sender_email = st.text_input("Sender Email", placeholder="yourlab@gmail.com",
                                         key="smtp_from_"+label)
        with ec2:
            sender_pass   = st.text_input("App Password", type="password",
                                          placeholder="Gmail App Password",
                                          key="smtp_pass_"+label)
            recip_email   = st.text_input("Recipient Email",
                                          value=st.session_state.get("pt_email",""),
                                          placeholder="patient@email.com",
                                          key="smtp_to_"+label)
            use_tls = st.checkbox("Use TLS", value=True, key="smtp_tls_"+label)

        if st.button("📤 Send Email Now", key="send_email_"+label):
            pdf_data = st.session_state.get("pdf_data")
            if not pdf_data:
                info_box("⚠️ No PDF report generated yet. Please generate a PDF first.", "warn")
            elif not recip_email:
                info_box("⚠️ Please enter a recipient email address.", "warn")
            elif not sender_email or not sender_pass:
                info_box("⚠️ Please enter SMTP sender credentials.", "warn")
            else:
                with st.spinner("Sending email..."):
                    ok, msg = send_report_email(
                        smtp_host, smtp_port, sender_email, sender_pass,
                        recip_email,
                        patient_name=st.session_state.get("pt_name","Patient"),
                        report_id=st.session_state.get("pdf_name","RPT"),
                        pdf_bytes=pdf_data, use_tls=use_tls)
                if ok:
                    info_box(msg, "success")
                else:
                    info_box(msg, "error")


# ═══════════════════════════════════════════════════════════════════════════════
#  PAGE: STATS (Statistics Dashboard)
# ═══════════════════════════════════════════════════════════════════════════════

def page_stats():
    nav_bar()
    hero_card("Statistics Dashboard",
              "Aggregated analytics across all analysed images", "📈")

    all_imgs = {n: r for n, r in st.session_state.images.items() if r["analysed"]}
    hist     = st.session_state.history

    if not all_imgs and not hist:
        info_box("⚠️ No data yet. Analyse some images first to see statistics.", "warn")
        return

    # ── Summary KPIs ──
    total_images = len(all_imgs)
    all_rbc  = [r["rbc"] for r in all_imgs.values()]
    all_wbc  = [r["wbc"] for r in all_imgs.values()]
    all_sc   = [r.get("health_score",0) for r in all_imgs.values()]
    all_conf = [r.get("confidence",0)   for r in all_imgs.values()]

    if not all_rbc:
        all_rbc = [h.get("rbc",0) for h in hist]
        all_wbc = [h.get("wbc",0) for h in hist]
        all_sc  = [h.get("score",0) for h in hist]
        all_conf= [0]*len(all_rbc)

    mean_rbc  = round(sum(all_rbc)/max(len(all_rbc),1), 1)
    mean_wbc  = round(sum(all_wbc)/max(len(all_wbc),1), 1)
    mean_sc   = round(sum(all_sc) /max(len(all_sc), 1), 1)
    mean_conf = round(sum(all_conf)/max(len(all_conf),1), 1)
    max_wbc   = max(all_wbc) if all_wbc else 0
    min_wbc   = min(all_wbc) if all_wbc else 0
    high_wbc_count = sum(1 for w in all_wbc if w > 10)
    low_rbc_count  = sum(1 for r in all_rbc if r < 4)

    kpi_cols = st.columns(4)
    kpis = [
        ("📷", "Images Analysed", total_images or len(hist), "#1A56DB"),
        ("🩸", "Mean RBC",        mean_rbc,   "#DC2626"),
        ("🦠", "Mean WBC",        mean_wbc,   "#0891B2"),
        ("💯", "Mean Health Score", "{0}/100".format(mean_sc), "#22C55E"),
    ]
    for col, (icon, label, val, color) in zip(kpi_cols, kpis):
        with col:
            st.markdown(stat_card_html(icon, label, val, color),
                        unsafe_allow_html=True)

    kpi2 = st.columns(4)
    kpis2 = [
        ("🔺", "Max WBC",          max_wbc,        "#F97316"),
        ("🔻", "Min WBC",          min_wbc,        "#6366F1"),
        ("⚠️", "High WBC Cases",   high_wbc_count, "#EF4444"),
        ("🎯", "Mean Confidence",  "{0}%".format(mean_conf), "#8B5CF6"),
    ]
    for col, (icon, label, val, color) in zip(kpi2, kpis2):
        with col:
            st.markdown(stat_card_html(icon, label, val, color),
                        unsafe_allow_html=True)

    st.markdown("---")

    # ── Distribution chart ──
    if PLOTLY_OK and len(all_rbc) >= 2:
        c1, c2 = st.columns(2)

        with c1:
            st.markdown('<div style="background:white;border-radius:14px;'
                        'padding:16px;box-shadow:0 4px 12px rgba(0,0,0,0.06);'
                        'margin-bottom:14px;">'
                        '<div style="font-family:Montserrat,sans-serif;font-size:12px;'
                        'font-weight:700;color:#0F172A;margin-bottom:8px;">'
                        '📊 RBC Distribution</div>',
                        unsafe_allow_html=True)
            fig_rbc = go.Figure()
            fig_rbc.add_trace(go.Histogram(x=all_rbc, nbinsx=10,
                marker_color="#DC2626", opacity=0.8, name="RBC"))
            fig_rbc.update_layout(height=200, margin=dict(l=0,r=0,t=10,b=0),
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                showlegend=False,
                xaxis=dict(showgrid=False, zeroline=False),
                yaxis=dict(showgrid=True, gridcolor="#F1F5F9", zeroline=False))
            st.plotly_chart(fig_rbc, use_container_width=True)
            st.markdown("</div>", unsafe_allow_html=True)

        with c2:
            st.markdown('<div style="background:white;border-radius:14px;'
                        'padding:16px;box-shadow:0 4px 12px rgba(0,0,0,0.06);'
                        'margin-bottom:14px;">'
                        '<div style="font-family:Montserrat,sans-serif;font-size:12px;'
                        'font-weight:700;color:#0F172A;margin-bottom:8px;">'
                        '📊 WBC Distribution</div>',
                        unsafe_allow_html=True)
            fig_wbc = go.Figure()
            fig_wbc.add_trace(go.Histogram(x=all_wbc, nbinsx=10,
                marker_color="#0891B2", opacity=0.8, name="WBC"))
            fig_wbc.update_layout(height=200, margin=dict(l=0,r=0,t=10,b=0),
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                showlegend=False,
                xaxis=dict(showgrid=False, zeroline=False),
                yaxis=dict(showgrid=True, gridcolor="#F1F5F9", zeroline=False))
            st.plotly_chart(fig_wbc, use_container_width=True)
            st.markdown("</div>", unsafe_allow_html=True)

        # Health score trend
        if len(all_sc) >= 2:
            st.markdown('<div style="background:white;border-radius:14px;'
                        'padding:16px;box-shadow:0 4px 12px rgba(0,0,0,0.06);'
                        'margin-bottom:14px;">'
                        '<div style="font-family:Montserrat,sans-serif;font-size:12px;'
                        'font-weight:700;color:#0F172A;margin-bottom:8px;">'
                        '📈 Health Score Trend Across Images</div>',
                        unsafe_allow_html=True)
            img_labels = list(all_imgs.keys()) if all_imgs else \
                         [h.get("report_id","") for h in hist]
            short_labels = [l[:15]+"..." if len(l)>15 else l for l in img_labels]
            fig_sc = go.Figure()
            fig_sc.add_trace(go.Scatter(
                x=short_labels, y=all_sc[:len(short_labels)],
                mode="lines+markers",
                line=dict(color="#22C55E", width=2),
                marker=dict(size=8, color="#22C55E"),
                fill="tozeroy", fillcolor="rgba(34,197,94,0.08)"))
            fig_sc.add_hline(y=75, line_dash="dash", line_color="#94A3B8",
                             annotation_text="Good threshold (75)")
            fig_sc.update_layout(height=200, margin=dict(l=0,r=0,t=10,b=0),
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                showlegend=False,
                xaxis=dict(showgrid=False, zeroline=False),
                yaxis=dict(showgrid=True, gridcolor="#F1F5F9",
                           zeroline=False, range=[0,100]))
            st.plotly_chart(fig_sc, use_container_width=True)
            st.markdown("</div>", unsafe_allow_html=True)

        # Severity breakdown pie
        sev_counts = {"NORMAL":0,"MILD":0,"MODERATE":0,"SEVERE":0}
        for w in all_wbc:
            s, _ = get_severity(w)
            sev_counts[s] = sev_counts.get(s, 0) + 1

        c3, c4 = st.columns(2)
        with c3:
            st.markdown('<div style="background:white;border-radius:14px;'
                        'padding:16px;box-shadow:0 4px 12px rgba(0,0,0,0.06);">'
                        '<div style="font-family:Montserrat,sans-serif;font-size:12px;'
                        'font-weight:700;color:#0F172A;margin-bottom:8px;">'
                        '🥧 Severity Breakdown</div>',
                        unsafe_allow_html=True)
            fig_pie = go.Figure(data=[go.Pie(
                labels=list(sev_counts.keys()),
                values=list(sev_counts.values()),
                hole=0.5,
                marker_colors=["#22C55E","#EAB308","#F97316","#EF4444"])])
            fig_pie.update_layout(height=220, margin=dict(l=0,r=0,t=10,b=0),
                paper_bgcolor="rgba(0,0,0,0)",
                legend=dict(orientation="h", yanchor="bottom", y=-0.3,
                            xanchor="center", x=0.5,
                            font=dict(size=10)))
            st.plotly_chart(fig_pie, use_container_width=True)
            st.markdown("</div>", unsafe_allow_html=True)

        with c4:
            # Flag distribution bar chart
            st.markdown('<div style="background:white;border-radius:14px;'
                        'padding:16px;box-shadow:0 4px 12px rgba(0,0,0,0.06);">'
                        '<div style="font-family:Montserrat,sans-serif;font-size:12px;'
                        'font-weight:700;color:#0F172A;margin-bottom:8px;">'
                        '🚦 RBC Flag Distribution</div>',
                        unsafe_allow_html=True)
            gender = st.session_state.get("pt_gender","Male")
            age    = st.session_state.get("pt_age", None)
            rbc_flags = [get_flag("RBC", r, gender, age) for r in all_rbc]
            flag_dist = {"LOW": rbc_flags.count("LOW"),
                         "NORMAL": rbc_flags.count("NORMAL"),
                         "HIGH": rbc_flags.count("HIGH")}
            fig_bar = go.Figure(data=[go.Bar(
                x=list(flag_dist.keys()),
                y=list(flag_dist.values()),
                marker_color=["#F59E0B","#22C55E","#EF4444"])])
            fig_bar.update_layout(height=220, margin=dict(l=0,r=0,t=10,b=0),
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                showlegend=False,
                xaxis=dict(showgrid=False, zeroline=False),
                yaxis=dict(showgrid=True, gridcolor="#F1F5F9", zeroline=False))
            st.plotly_chart(fig_bar, use_container_width=True)
            st.markdown("</div>", unsafe_allow_html=True)

    # Detailed stats table
    st.markdown("---")
    st.markdown('<div style="font-family:Montserrat,sans-serif;font-size:13px;'
                'font-weight:700;color:#0F172A;text-transform:uppercase;'
                'letter-spacing:1px;margin-bottom:12px;">📋 Per-Image Statistics</div>',
                unsafe_allow_html=True)

    if all_imgs:
        rows_html = ""
        gender = st.session_state.get("pt_gender","Male")
        age    = st.session_state.get("pt_age", None)
        for name, rec in all_imgs.items():
            rf   = get_flag("RBC", rec["rbc"], gender, age)
            wf   = get_flag("WBC", rec["wbc"], gender, age)
            sev,_= get_severity(rec["wbc"])
            sc   = rec.get("health_score", 0)
            conf = rec.get("confidence", 0)
            sev_col = {"NORMAL":"#22C55E","MILD":"#EAB308",
                       "MODERATE":"#F97316","SEVERE":"#EF4444"}.get(sev,"#94A3B8")
            rf_col  = {"NORMAL":"#22C55E","HIGH":"#EF4444","LOW":"#F59E0B"}.get(rf,"#94A3B8")
            wf_col  = {"NORMAL":"#22C55E","HIGH":"#EF4444","LOW":"#F59E0B"}.get(wf,"#94A3B8")
            short = name[:25]+"..." if len(name)>25 else name
            rows_html += (
                '<tr style="border-bottom:1px solid #F1F5F9;">'
                '<td style="padding:9px 12px;font-family:Roboto,sans-serif;'
                'font-size:12px;color:#334155;">{n}</td>'
                '<td style="padding:9px 12px;text-align:center;font-family:Russo One,sans-serif;'
                'font-size:15px;color:#DC2626;">{rbc}</td>'
                '<td style="padding:9px 12px;text-align:center;font-family:Russo One,sans-serif;'
                'font-size:15px;color:#0891B2;">{wbc}</td>'
                '<td style="padding:9px 12px;text-align:center;">'
                '<span style="color:{rfc};font-weight:700;font-size:11px;">{rf}</span></td>'
                '<td style="padding:9px 12px;text-align:center;">'
                '<span style="color:{wfc};font-weight:700;font-size:11px;">{wf}</span></td>'
                '<td style="padding:9px 12px;text-align:center;font-family:Russo One,sans-serif;'
                'font-size:14px;color:{svc};">{sc}</td>'
                '<td style="padding:9px 12px;text-align:center;font-family:Roboto,sans-serif;'
                'font-size:12px;color:{svc};font-weight:700;">{sev}</td>'
                '<td style="padding:9px 12px;text-align:center;font-family:Roboto,sans-serif;'
                'font-size:12px;color:#64748B;">{conf}%</td>'
                '</tr>'
            ).format(n=short, rbc=rec["rbc"], wbc=rec["wbc"],
                     rf=rf, rfc=rf_col, wf=wf, wfc=wf_col,
                     sc=sc, svc=sev_col, sev=sev, conf=conf)

        st.markdown(
            '<div style="background:white;border-radius:14px;overflow:hidden;'
            'box-shadow:0 4px 12px rgba(0,0,0,0.06);">'
            '<table style="width:100%;border-collapse:collapse;">'
            '<thead><tr style="background:#F8FAFC;">'
            + "".join([
                '<th style="padding:9px 12px;text-align:{a};font-family:Montserrat,sans-serif;'
                'font-size:9px;font-weight:700;color:#64748B;text-transform:uppercase;'
                'letter-spacing:1px;border-bottom:1px solid #EDF2F7;">{h}</th>'.format(
                    h=h, a=("left" if i==0 else "center"))
                for i,h in enumerate(["Image","RBC","WBC","RBC Flag","WBC Flag",
                                       "Score","Severity","Confidence"])
            ]) +
            '</tr></thead><tbody>{rows}</tbody></table></div>'.format(rows=rows_html),
            unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
#  PAGE: BATCH (Multi-patient batch PDF + QR codes)
# ═══════════════════════════════════════════════════════════════════════════════

def page_batch():
    nav_bar()
    hero_card("Batch Report & QR Codes",
              "Generate one PDF for multiple patients + scan-to-view QR codes", "📋")

    st.markdown('<div style="font-family:Montserrat,sans-serif;font-size:13px;'
                'font-weight:700;color:#0F172A;text-transform:uppercase;'
                'letter-spacing:1px;margin-bottom:10px;">➕ Add Patient to Batch</div>',
                unsafe_allow_html=True)

    # Add current session patient to batch
    info_box("ℹ️ Fill in patient details on the <strong>Patient</strong> page, "
             "analyse images, then add to batch here. Repeat for each patient.", "info")

    analysed = {n: r for n, r in st.session_state.images.items() if r["analysed"]}
    if analysed:
        pa_c1, pa_c2 = st.columns([3,1])
        with pa_c1:
            pt_name = st.session_state.get("pt_name","") or "Unknown Patient"
            pt_pid  = st.session_state.get("pt_pid","")
            total_rbc = sum(r["rbc"] for r in analysed.values())
            total_wbc = sum(r["wbc"] for r in analysed.values())
            sev_txt, _ = get_severity(total_wbc)
            score      = get_health_score(total_rbc, total_wbc,
                                          st.session_state.get("pt_gender","Male"),
                                          st.session_state.get("pt_age",None))
            st.markdown(
                '<div style="background:#F8FAFC;border:1px solid #E2E8F0;'
                'border-radius:10px;padding:10px 14px;font-family:Open Sans,sans-serif;'
                'font-size:13px;color:#334155;">'
                '👤 <strong>{name}</strong> ({pid}) — '
                '🩸 RBC: {rbc} | 🦠 WBC: {wbc} | '
                '💯 Score: {sc} | ⚡ {sev} | 🔬 {imgs} image(s)'
                '</div>'.format(name=pt_name, pid=pt_pid or "N/A",
                               rbc=total_rbc, wbc=total_wbc,
                               sc=score, sev=sev_txt,
                               imgs=len(analysed)),
                unsafe_allow_html=True)
        with pa_c2:
            if st.button("➕ Add to Batch"):
                # Check not already added
                exists = any(bp.get("pid") == pt_pid and pt_pid
                             for bp in st.session_state.batch_patients)
                if exists:
                    info_box("⚠️ This patient is already in the batch.", "warn")
                else:
                    # Generate individual PDF for this patient
                    ts  = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
                    rid = make_report_id()
                    img_data = [{"name":n,"rbc":r["rbc"],"wbc":r["wbc"],"notes":r.get("notes","")}
                                for n,r in analysed.items()]
                    annot_pil = None
                    first_name = list(analysed.keys())[0]
                    if analysed[first_name].get("annot") is not None:
                        annot_pil = Image.fromarray(
                            cv2.cvtColor(analysed[first_name]["annot"],
                                         cv2.COLOR_BGR2RGB))
                    pdf_bytes = generate_pdf(get_patient(), get_lab(),
                                            img_data, rid, ts, annot_pil)
                    # QR code data = base64 pdf (for local viewing intent)
                    qr_data = "CELLSCOPE:{rid}:{name}:{date}".format(
                        rid=rid, name=pt_name, date=ts)
                    entry = {
                        "pid":       pt_pid,
                        "name":      pt_name,
                        "report_id": rid,
                        "timestamp": ts,
                        "rbc":       total_rbc,
                        "wbc":       total_wbc,
                        "severity":  sev_txt,
                        "score":     score,
                        "images":    len(analysed),
                        "pdf_bytes": pdf_bytes,
                        "qr_data":   qr_data,
                        "patient":   get_patient(),
                        "lab":       get_lab(),
                        "img_data":  img_data,
                    }
                    st.session_state.batch_patients.append(entry)
                    info_box("✅ {0} added to batch ({1} total).".format(
                        pt_name, len(st.session_state.batch_patients)), "success")
                    st.rerun()
    else:
        info_box("⚠️ No analysed images in current session. Analyse images first.", "warn")

    st.markdown("---")

    batch = st.session_state.batch_patients
    if not batch:
        info_box("📋 Batch is empty. Add patients above to get started.", "info")
        return

    # Batch list
    st.markdown('<div style="font-family:Montserrat,sans-serif;font-size:13px;'
                'font-weight:700;color:#0F172A;text-transform:uppercase;'
                'letter-spacing:1px;margin-bottom:10px;">'
                '📋 Batch Queue — {n} Patient(s)</div>'.format(n=len(batch)),
                unsafe_allow_html=True)

    for i, bp in enumerate(batch):
        sev_col = {"NORMAL":"#22C55E","MILD":"#EAB308",
                   "MODERATE":"#F97316","SEVERE":"#EF4444"}.get(bp["severity"],"#94A3B8")
        bc1, bc2, bc3, bc4 = st.columns([3, 1.5, 1.5, 1])
        with bc1:
            st.markdown(
                '<div style="background:white;border-radius:10px;'
                'padding:10px 14px;box-shadow:0 2px 8px rgba(0,0,0,0.06);">'
                '<div style="font-family:Montserrat,sans-serif;font-size:13px;'
                'font-weight:600;color:#0F172A;">#{i} {name}</div>'
                '<div style="font-family:Roboto,sans-serif;font-size:11px;'
                'color:#94A3B8;">🏷️ {rid} | 🩸{rbc} 🦠{wbc} | '
                '<span style="color:{sc};">{sev}</span></div>'
                '</div>'.format(i=i+1, name=bp["name"], rid=bp["report_id"],
                               rbc=bp["rbc"], wbc=bp["wbc"],
                               sc=sev_col, sev=bp["severity"]),
                unsafe_allow_html=True)
        with bc2:
            # Individual PDF download
            st.download_button("⬇️ PDF",
                               data=bp["pdf_bytes"],
                               file_name="report_{0}.pdf".format(bp["report_id"]),
                               mime="application/pdf",
                               key="dl_bp_{0}".format(i))
        with bc3:
            # QR code download
            if QR_OK:
                qr_img = make_qr_code(bp["qr_data"])
                if qr_img:
                    qr_buf = io.BytesIO()
                    qr_img.save(qr_buf, format="PNG")
                    qr_buf.seek(0)
                    st.download_button("📲 QR Code",
                                       data=qr_buf.getvalue(),
                                       file_name="qr_{0}.png".format(bp["report_id"]),
                                       mime="image/png",
                                       key="qr_dl_{0}".format(i))
            else:
                st.markdown('<span style="font-size:11px;color:#94A3B8;">'
                            'Install qrcode</span>', unsafe_allow_html=True)
        with bc4:
            if st.button("🗑️", key="rm_bp_{0}".format(i)):
                st.session_state.batch_patients.pop(i)
                st.rerun()

    st.markdown("---")

    # QR Code Preview Gallery
    if QR_OK and batch:
        st.markdown('<div style="font-family:Montserrat,sans-serif;font-size:13px;'
                    'font-weight:700;color:#0F172A;text-transform:uppercase;'
                    'letter-spacing:1px;margin-bottom:12px;">📲 QR Code Gallery</div>',
                    unsafe_allow_html=True)
        info_box("ℹ️ Each QR code encodes the Report ID and patient identifier. "
                 "Scanning it will display the report reference. For full digital "
                 "report access, integrate with a web server or cloud URL.", "info")
        qr_cols = st.columns(min(len(batch), 4))
        for i, (bp, col) in enumerate(zip(batch, qr_cols)):
            with col:
                qr_img = make_qr_code(bp["qr_data"])
                if qr_img:
                    st.image(qr_img, caption="{0}\n{1}".format(
                        bp["name"][:18], bp["report_id"]),
                             use_container_width=True)

    st.markdown("---")

    # Batch PDF — all patients in one document
    st.markdown('<div style="font-family:Montserrat,sans-serif;font-size:13px;'
                'font-weight:700;color:#0F172A;text-transform:uppercase;'
                'letter-spacing:1px;margin-bottom:10px;">'
                '📄 Generate Combined Batch PDF</div>',
                unsafe_allow_html=True)
    info_box("ℹ️ This creates a single PDF containing reports for ALL patients "
             "in the batch queue, separated by page breaks.", "info")

    bc_left, bc_right = st.columns([1, 3])
    with bc_left:
        if st.button("🖨️ Build Batch PDF", key="build_batch_pdf"):
            if not batch:
                info_box("⚠️ Batch is empty.", "warn")
            else:
                with st.spinner("Building combined PDF for {0} patient(s)...".format(
                        len(batch))):
                    # Merge all individual PDFs
                    try:
                        from reportlab.lib.pagesizes import A4
                        from reportlab.platypus import SimpleDocTemplate, Spacer, PageBreak
                        combined_buf = io.BytesIO()
                        # Use PyPDF2-style merge via reportlab flowables
                        # Simpler: concatenate PDFs using pypdf if available, else just zip
                        try:
                            from pypdf import PdfWriter, PdfReader as PR2
                            writer = PdfWriter()
                            for bp in batch:
                                reader = PR2(io.BytesIO(bp["pdf_bytes"]))
                                for page in reader.pages:
                                    writer.add_page(page)
                            writer.write(combined_buf)
                            combined_pdf = combined_buf.getvalue()
                        except ImportError:
                            # Fallback: generate fresh combined doc
                            from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
                            from reportlab.lib.styles import ParagraphStyle
                            story = []
                            doc = SimpleDocTemplate(combined_buf, pagesize=A4,
                                leftMargin=15*6.35, rightMargin=15*6.35,
                                topMargin=12*6.35, bottomMargin=12*6.35)
                            from reportlab.lib import colors
                            for bi, bp in enumerate(batch):
                                story.append(Paragraph(
                                    "PATIENT {0}: {1} — Report {2}".format(
                                        bi+1, bp["name"], bp["report_id"]),
                                    ParagraphStyle("h", fontName="Helvetica-Bold",
                                                   fontSize=16, textColor=colors.HexColor("#1A56DB"),
                                                   spaceAfter=20)))
                                story.append(Spacer(1, 20))
                                if bi < len(batch)-1:
                                    story.append(PageBreak())
                            doc.build(story)
                            combined_pdf = combined_buf.getvalue()

                        st.session_state["batch_pdf_data"]  = combined_pdf
                        st.session_state["batch_pdf_ready"] = True
                        info_box("✅ Batch PDF ready for download!", "success")
                    except Exception as e:
                        info_box("❌ Error generating batch PDF: {0}".format(str(e)), "error")

    with bc_right:
        if st.session_state.get("batch_pdf_ready") and st.session_state.get("batch_pdf_data"):
            st.download_button(
                "⬇️ Download Combined Batch PDF ({0} patients)".format(len(batch)),
                data=st.session_state["batch_pdf_data"],
                file_name="cellscope_batch_{0}.pdf".format(
                    datetime.date.today()),
                mime="application/pdf",
                key="batch_pdf_dl")

    st.markdown("---")

    # Email batch option
    st.markdown('<div style="font-family:Montserrat,sans-serif;font-size:13px;'
                'font-weight:700;color:#0F172A;text-transform:uppercase;'
                'letter-spacing:1px;margin-bottom:10px;">📧 Email Individual Reports</div>',
                unsafe_allow_html=True)

    with st.expander("Configure SMTP & Send Individual Reports"):
        info_box("ℹ️ This will email each patient's report to their registered email "
                 "address. Make sure patient emails are set in the Patient page.", "info")
        em1, em2 = st.columns(2)
        with em1:
            b_smtp_host = st.text_input("SMTP Host", value="smtp.gmail.com", key="b_smtp_h")
            b_smtp_port = st.text_input("SMTP Port", value="587",            key="b_smtp_p")
            b_sender    = st.text_input("Sender Email", key="b_smtp_from")
        with em2:
            b_pass      = st.text_input("App Password", type="password", key="b_smtp_pass")
            b_use_tls   = st.checkbox("Use TLS", value=True, key="b_tls")

        if st.button("📤 Send All Reports via Email", key="batch_send_email"):
            if not b_sender or not b_pass:
                info_box("⚠️ Enter SMTP credentials first.", "warn")
            else:
                results = []
                prog = st.progress(0)
                for i, bp in enumerate(batch):
                    recip = bp["patient"].get("email","")
                    if not recip:
                        results.append("⚠️ {0}: No email on file".format(bp["name"]))
                    else:
                        ok, msg = send_report_email(
                            b_smtp_host, b_smtp_port, b_sender, b_pass,
                            recip, bp["name"], bp["report_id"],
                            bp["pdf_bytes"], b_use_tls)
                        results.append(("✅ " if ok else "❌ ") +
                                       "{0}: {1}".format(bp["name"], msg))
                    prog.progress(float(i+1)/float(len(batch)))
                prog.empty()
                for r in results:
                    style = "success" if r.startswith("✅") else \
                            "warn"    if r.startswith("⚠️") else "error"
                    info_box(r, style)

    # Clear batch
    st.markdown("")
    if st.button("🗑️ Clear Entire Batch"):
        st.session_state.batch_patients = []
        st.session_state["batch_pdf_ready"] = False
        st.session_state["batch_pdf_data"]  = None
        st.rerun()


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN ROUTER
# ═══════════════════════════════════════════════════════════════════════════════

stage = st.session_state.app_stage

if stage == "splash":
    screen_splash()

elif stage == "login":
    screen_login()

elif stage == "signup":
    screen_signup()

elif stage == "app":
    # Load user data from Supabase once after login
    load_from_supabase()

    sensitivity = build_sidebar()

    page = st.session_state.page
    if page == "Analyse":
        page_analyse(sensitivity)
    elif page == "Patient":
        page_patient()
    elif page == "History":
        page_history()
    elif page == "Compare":
        page_compare()
    elif page == "Stats":
        page_stats()
    elif page == "Batch":
        page_batch()
    elif page == "Lab":
        page_lab()
    elif page == "Info":
        page_info()
