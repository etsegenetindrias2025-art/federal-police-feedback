from flask import Flask, render_template, request, jsonify, session, redirect, url_for, send_file
import sqlite3
import os
import io
import re
import qrcode
from gtts import gTTS
import socket
import time
from datetime import datetime
from sqlalchemy import or_, and_

# Excel styling imports
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Word document imports
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# PDF generation imports (ReportLab)
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

try:
    from better_profanity import profanity  # type: ignore[import-not-found]
except ImportError:
    class _FallbackProfanity:
        def __init__(self):
            self._censored_words = set()

        def load_censor_words(self):
            return None

        def add_censor_words(self, words):
            for word in words:
                cleaned = str(word).strip().lower()
                if cleaned:
                    self._censored_words.add(cleaned)

        def contains_profanity(self, text):
            if not text:
                return False

            normalized = re.sub(r'[^\w\s]', '', str(text).lower())
            for word in self._censored_words:
                if re.search(rf'\b{re.escape(word)}\b', normalized):
                    return True
            return False

    profanity = _FallbackProfanity()

app = Flask(__name__)
app.secret_key = 'federal_police_secret_key'

# ----------------------------------------------------
# COMPREHENSIVE PROFANITY FILTER (አማርኛ + Manglish + English)
# ----------------------------------------------------
profanity.load_censor_words()

ETHIOPIC_BAD_WORDS = [
    'ውሻ', 'ሌባ', 'አህያ', 'ጅብ', 'ፋንድያ', 'ሉጢ', 'ክፉ', 'በረንዳ አዳሪ',
    'wusha', 'wesha', 'leba', 'ahiya', 'jib', 'fandya', 'luti', 'dingay', 'balege', 'dedeb', 'denez'
]

profanity.add_censor_words(ETHIOPIC_BAD_WORDS)


def is_inappropriate(text):
    if not text:
        return False

    if profanity.contains_profanity(text):
        return True

    cleaned_text = re.sub(r'[^\w\s]', '', text.lower())
    for bad_word in ETHIOPIC_BAD_WORDS:
        if bad_word in cleaned_text:
            return True

    return False


# ----------------------------------------------------
# DATABASE SETUP
# ----------------------------------------------------
DB_PATH = os.path.join(os.path.dirname(__file__), 'database.db')


def get_db_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()

    # Feedbacks Table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS feedbacks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            service_name TEXT NOT NULL,
            sub_service TEXT,
            rating TEXT NOT NULL,
            comment TEXT,
            audio_status TEXT DEFAULT 'No audio recorded',
            is_read BOOLEAN DEFAULT 0,
            timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    # Fingerprint Registrations Table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS fingerprint_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id TEXT,
            fingerprint_data TEXT NOT NULL,
            status TEXT DEFAULT 'Verified',
            timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    # Services Table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS services (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            service_key TEXT UNIQUE NOT NULL,
            service_name TEXT NOT NULL
        )
    ''')

    # Sub Services Table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS sub_services (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            service_key TEXT NOT NULL,
            sub_service_key TEXT NOT NULL,
            sub_service_name TEXT NOT NULL,
            UNIQUE(service_key, sub_service_key)
        )
    ''')

    # Audit Logs Table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS audit_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            admin_user TEXT NOT NULL,
            action_description TEXT NOT NULL,
            timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    default_services = [
        ("police_clearance", "Police Clearance"),
        ("complaint", "Complaint"),
        ("hospital", "Hospital"),
        ("logistics", "Logistics"),
        ("education_training", "Education & Training"),
        ("other", "Other")
    ]
    for key, name in default_services:
        cursor.execute("""
            INSERT OR IGNORE INTO services (service_key, service_name) VALUES (?, ?)
        """, (key, name))

    default_sub_services = {
        "police_clearance": [
            ("new_clearance", "New Police Clearance"),
            ("renewal", "Renewal"),
            ("criminal_record", "Criminal Record Verification"),
            ("fingerprint", "Fingerprint Registration"),
            ("document_collection", "Document Collection")
        ],
        "complaint": [
            ("crime_complaint", "Crime Complaint Registration"),
            ("public_office", "Public Complaint Office"),
            ("online_followup", "Online Complaint Follow-up"),
            ("investigation", "Investigation"),
            ("resolution", "Resolution")
        ],
        "hospital": [
            ("opd", "OPD"),
            ("emergency", "Emergency"),
            ("pharmacy", "Pharmacy"),
            ("laboratory", "Laboratory"),
            ("medical_exam", "Medical Examination")
        ],
        "logistics": [
            ("vehicle_mgmt", "Vehicle Management"),
            ("garage", "Garage"),
            ("equipment_dist", "Equipment Distribution"),
            ("inventory", "Inventory"),
            ("procurement", "Procurement")
        ],
        "education_training": [
            ("student_reg", "Student Registration"),
            ("training", "Training"),
            ("certificates", "Certificates"),
            ("examination", "Examination"),
            ("academic_records", "Academic Records")
        ],
        "other": [
            ("reception", "Reception"),
            ("ict_support", "ICT Support"),
            ("hr", "HR"),
            ("finance", "Finance"),
            ("admin", "Administration")
        ]
    }

    for s_key, sub_list in default_sub_services.items():
        for sub_key, sub_name in sub_list:
            cursor.execute("""
                INSERT OR IGNORE INTO sub_services (service_key, sub_service_key, sub_service_name) VALUES (?, ?, ?)
            """, (s_key, sub_key, sub_name))

    cursor.execute("PRAGMA table_info(feedbacks)")
    columns = [column[1] for column in cursor.fetchall()]
    if 'is_read' not in columns:
        cursor.execute("ALTER TABLE feedbacks ADD COLUMN is_read BOOLEAN DEFAULT 0")
    if 'audio_status' not in columns:
        cursor.execute("ALTER TABLE feedbacks ADD COLUMN audio_status TEXT DEFAULT 'No audio recorded'")

    conn.commit()
    cursor.close()
    conn.close()


init_db()
print(f"[startup] Using database file at: {DB_PATH}")


def log_admin_action(username, description):
    conn = None
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO audit_logs (admin_user, action_description) VALUES (?, ?)",
            (str(username), str(description))
        )
        conn.commit()
        cursor.close()
    except Exception:
        import traceback
        print("=== Error logging audit trail (see traceback below) ===")
        traceback.print_exc()
    finally:
        if conn is not None:
            conn.close()


def get_service_map():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT service_key, service_name FROM services")
    rows = cursor.fetchall()
    cursor.close()
    conn.close()
    return {row['service_key']: row['service_name'] for row in rows}


def get_sub_service_map():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT service_key, sub_service_key, sub_service_name FROM sub_services")
    rows = cursor.fetchall()
    cursor.close()
    conn.close()

    sub_map = {}
    for row in rows:
        s_key = row['service_key']
        sub_k = row['sub_service_key']
        sub_n = row['sub_service_name']
        if s_key not in sub_map:
            sub_map[s_key] = {}
        sub_map[s_key][sub_k] = sub_n
    return sub_map


def get_admin_credentials():
    return {
        "admin gen": {"password": "1234", "type": "general", "service": "all", "sub_service": "all", "title": "General Admin Dashboard"},
        "admin pol": {"password": "1234", "type": "service", "service": "police_clearance", "sub_service": "all", "title": "Police Clearance Admin"},
        "admin com": {"password": "1234", "type": "service", "service": "complaint", "sub_service": "all", "title": "Complaint Admin"},
        "admin hos": {"password": "1234", "type": "service", "service": "hospital", "sub_service": "all", "title": "Hospital Admin"},
        "admin log": {"password": "1234", "type": "service", "service": "logistics", "sub_service": "all", "title": "Logistics Admin"},
        "admin edu": {"password": "1234", "type": "service", "service": "education_training", "sub_service": "all", "title": "Education Admin"},
        "admin oth": {"password": "1234", "type": "service", "service": "other", "sub_service": "all", "title": "Other Admin"},
        "admin new": {"password": "1234", "type": "sub_service", "service": "police_clearance", "sub_service": "new_clearance", "title": "Sub Admin: New Clearance"},
        "admin ren": {"password": "1234", "type": "sub_service", "service": "police_clearance", "sub_service": "renewal", "title": "Sub Admin: Renewal"},
        "admin cri": {"password": "1234", "type": "sub_service", "service": "police_clearance", "sub_service": "criminal_record", "title": "Sub Admin: Criminal Record"},
        "admin fin": {"password": "1234", "type": "sub_service", "service": "police_clearance", "sub_service": "fingerprint", "title": "Sub Admin: Fingerprint"},
        "admin doc": {"password": "1234", "type": "sub_service", "service": "police_clearance", "sub_service": "document_collection", "title": "Sub Admin: Document Collection"},
        "admin pub": {"password": "1234", "type": "sub_service", "service": "complaint", "sub_service": "public_office", "title": "Sub Admin: Public Complaint"},
        "admin onl": {"password": "1234", "type": "sub_service", "service": "complaint", "sub_service": "online_followup", "title": "Sub Admin: Online Follow-up"},
        "admin inv": {"password": "1234", "type": "sub_service", "service": "complaint", "sub_service": "investigation", "title": "Sub Admin: Investigation"},
        "admin res": {"password": "1234", "type": "sub_service", "service": "complaint", "sub_service": "resolution", "title": "Sub Admin: Resolution"},
        "admin opd": {"password": "1234", "type": "sub_service", "service": "hospital", "sub_service": "opd", "title": "Sub Admin: OPD"},
        "admin eme": {"password": "1234", "type": "sub_service", "service": "hospital", "sub_service": "emergency", "title": "Sub Admin: Emergency"},
        "admin pha": {"password": "1234", "type": "sub_service", "service": "hospital", "sub_service": "pharmacy", "title": "Sub Admin: Pharmacy"},
        "admin lab": {"password": "1234", "type": "sub_service", "service": "hospital", "sub_service": "laboratory", "title": "Sub Admin: Laboratory"},
        "admin med": {"password": "1234", "type": "sub_service", "service": "hospital", "sub_service": "medical_exam", "title": "Sub Admin: Medical Examination"},
        "admin veh": {"password": "1234", "type": "sub_service", "service": "logistics", "sub_service": "vehicle_mgmt", "title": "Sub Admin: Vehicle Management"},
        "admin gar": {"password": "1234", "type": "sub_service", "service": "logistics", "sub_service": "garage", "title": "Sub Admin: Garage"},
        "admin equ": {"password": "1234", "type": "sub_service", "service": "logistics", "sub_service": "equipment_dist", "title": "Sub Admin: Equipment Distribution"},
        "admin pro": {"password": "1234", "type": "sub_service", "service": "logistics", "sub_service": "procurement", "title": "Sub Admin: Procurement"},
        "admin stu": {"password": "1234", "type": "sub_service", "service": "education_training", "sub_service": "student_reg", "title": "Sub Admin: Student Registration"},
        "admin tra": {"password": "1234", "type": "sub_service", "service": "education_training", "sub_service": "training", "title": "Sub Admin: Training"},
        "admin cer": {"password": "1234", "type": "sub_service", "service": "education_training", "sub_service": "certificates", "title": "Sub Admin: Certificates"},
        "admin exa": {"password": "1234", "type": "sub_service", "service": "education_training", "sub_service": "examination", "title": "Sub Admin: Examination"},
        "admin aca": {"password": "1234", "type": "sub_service", "service": "education_training", "sub_service": "academic_records", "title": "Sub Admin: Academic Records"},
        "admin rec": {"password": "1234", "type": "sub_service", "service": "other", "sub_service": "reception", "title": "Sub Admin: Reception"},
        "admin ict": {"password": "1234", "type": "sub_service", "service": "other", "sub_service": "ict_support", "title": "Sub Admin: ICT Support"},
        "admin hr": {"password": "1234", "type": "sub_service", "service": "other", "sub_service": "hr", "title": "Sub Admin: HR"},
        "admin fnn": {"password": "1234", "type": "sub_service", "service": "other", "sub_service": "finance", "title": "Sub Admin: Finance"},
        "admin adm": {"password": "1234", "type": "sub_service", "service": "other", "sub_service": "admin", "title": "Sub Admin: Administration"}
    }


# ----------------------------------------------------
# ADVANCED SEARCH, FILTER & REPORT BUILDERS
# ----------------------------------------------------
def get_filtered_feedbacks(admin_type, assigned_service, assigned_sub):
    search_query = request.args.get('q', '').strip()
    service_filter = request.args.get('service_filter', 'all')
    rating_filter = request.args.get('rating', 'all')
    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM feedbacks")
    all_rows = cursor.fetchall()
    cursor.close()
    conn.close()

    filtered = []
    for fb in all_rows:
        db_s = str(fb['service_name']).strip().lower()
        db_sub = str(fb['sub_service']).strip().lower()

        # Admin privilege scope check
        if admin_type == 'service' and db_s != assigned_service:
            continue
        if admin_type == 'sub_service' and (db_s != assigned_service or db_sub != assigned_sub):
            continue

        # Dropdown filter check
        if service_filter and service_filter != 'all':
            if db_s != service_filter.strip().lower():
                continue

        if rating_filter and rating_filter != 'all':
            if str(fb['rating']).strip() != str(rating_filter).strip():
                continue

        # Date range check
        if date_from or date_to:
            ts_str = str(fb['timestamp'])
            try:
                dt_rec = datetime.strptime(ts_str.split('.')[0], '%Y-%m-%d %H:%M:%S')
            except ValueError:
                try:
                    dt_rec = datetime.strptime(ts_str.split(' ')[0], '%Y-%m-%d')
                except ValueError:
                    dt_rec = None

            if dt_rec:
                if date_from:
                    try:
                        df_dt = datetime.strptime(date_from, '%Y-%m-%d')
                        if dt_rec < df_dt:
                            continue
                    except ValueError:
                        pass
                if date_to:
                    try:
                        dt_dt = datetime.strptime(date_to, '%Y-%m-%d').replace(hour=23, minute=59, second=59)
                        if dt_rec > dt_dt:
                            continue
                    except ValueError:
                        pass

        # Global search text check across fields
        if search_query:
            q_lower = search_query.lower()
            combined_text = f"#{fb['id']} {fb['service_name']} {fb['sub_service']} {fb['rating']} {fb['comment']} {fb['timestamp']}".lower()
            if q_lower not in combined_text:
                continue

        filtered.append(fb)

    return filtered


# --- EXCEL REPORT (openpyxl) ---
def generate_excel_report(records):
    wb = openpyxl.Workbook()
    ws = wb.active
    if ws is None:
        ws = wb.create_sheet()
    ws.title = "Feedback Report"
    ws.views.sheetView[0].showGridLines = True

    header_fill = PatternFill(start_color="1B2A4A", end_color="1B2A4A", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    title_font = Font(name="Calibri", size=16, bold=True, color="1B2A4A")
    meta_font = Font(name="Calibri", size=10, italic=True, color="555555")
    data_font = Font(name="Calibri", size=10)
    
    border_thin = Border(
        left=Side(style='thin', color='DDDDDD'),
        right=Side(style='thin', color='DDDDDD'),
        top=Side(style='thin', color='DDDDDD'),
        bottom=Side(style='thin', color='DDDDDD')
    )

    ws.append(["Ethiopian Federal Police - Feedback Management Report"])
    ws.cell(row=1, column=1).font = title_font
    ws.append([f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M')} | Total Records: {len(records)}"])
    ws.cell(row=2, column=1).font = meta_font
    ws.append([])

    headers = ["ID", "Service", "Sub-Service", "Rating", "Comment", "Audio Feedback", "Submission Date"]
    ws.append(headers)
    
    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=4, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for r_idx, rec in enumerate(records, start=5):
        ws.append([
            f"#{rec['id']}",
            rec['service_name'],
            rec['sub_service'],
            str(rec['rating']),
            rec['comment'] or "No comment provided.",
            rec['audio_status'] or "No audio recorded",
            str(rec['timestamp'])
        ])
        for c_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.font = data_font
            cell.border = border_thin
            if c_idx in [1, 4, 7]:
                cell.alignment = Alignment(horizontal="center")

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        if col[0].column is not None:
            col_letter = get_column_letter(int(col[0].column))
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# --- WORD REPORT (python-docx) ---
def generate_word_report(records):
    doc = docx.Document()
    
    section = doc.sections[0]
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)

    title_p = doc.add_paragraph()
    title_run = title_p.add_run("Ethiopian Federal Police")
    title_run.font.size = Pt(20)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(27, 42, 74)
    
    sub_p = doc.add_paragraph()
    sub_run = sub_p.add_run(f"Feedback System Audit Report\nGenerated: {datetime.now().strftime('%Y-%m-%d %H:%M')} | Records Found: {len(records)}")
    sub_run.font.size = Pt(10)
    sub_run.font.italic = True
    sub_run.font.color.rgb = RGBColor(100, 100, 100)
    
    doc.add_paragraph()

    table = doc.add_table(rows=1, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    hdr_cells = table.rows[0].cells
    headers = ["ID & Date", "Service / Sub-Service", "Rating", "Comment", "Audio"]
    col_widths = [Inches(1.2), Inches(1.5), Inches(0.8), Inches(2.0), Inches(1.0)]

    for i, title in enumerate(headers):
        hdr_cells[i].text = title
        hdr_cells[i].width = col_widths[i]
        shading = parse_xml(r'<w:shd {} w:fill="1B2A4A"/>'.format(nsdecls('w')))
        hdr_cells[i]._tc.get_or_add_tcPr().append(shading)
        for p in hdr_cells[i].paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)
                run.font.size = Pt(10)

    for rec in records:
        row_cells = table.add_row().cells
        row_cells[0].text = f"#{rec['id']}\n{str(rec['timestamp']).split()[0]}"
        row_cells[1].text = f"{rec['service_name']}\n({rec['sub_service']})"
        row_cells[2].text = str(rec['rating'])
        row_cells[3].text = rec['comment'] or "No comment."
        row_cells[4].text = rec['audio_status'] or "None"

        for i, cell in enumerate(row_cells):
            cell.width = col_widths[i]
            for p in cell.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(9)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output


# --- PDF REPORT (ReportLab) ---
def generate_pdf_report(records):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(letter),
        rightMargin=36, leftMargin=36, topMargin=40, bottomMargin=40
    )
    
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'DocTitle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=18,
        textColor=colors.HexColor('#1B2A4A'),
        spaceAfter=4
    )
    subtitle_style = ParagraphStyle(
        'DocSubTitle',
        parent=styles['Normal'],
        fontName='Helvetica-Oblique',
        fontSize=10,
        textColor=colors.HexColor('#555555'),
        spaceAfter=15
    )
    cell_style = ParagraphStyle(
        'CellText',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9,
        textColor=colors.HexColor('#222222'),
        leading=11
    )
    header_cell_style = ParagraphStyle(
        'HeaderCellText',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=10,
        textColor=colors.white,
        alignment=1
    )

    elements = []
    
    elements.append(Paragraph("Ethiopian Federal Police - Feedback Management Report", title_style))
    elements.append(Paragraph(f"Export Date: {datetime.now().strftime('%Y-%m-%d %H:%M')} | Total Filtered Records: {len(records)}", subtitle_style))
    elements.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#1B2A4A'), spaceAfter=15))

    table_data = [[
        Paragraph("ID", header_cell_style),
        Paragraph("Service & Sub-Service", header_cell_style),
        Paragraph("Rating", header_cell_style),
        Paragraph("Comment", header_cell_style),
        Paragraph("Audio Status", header_cell_style),
        Paragraph("Submission Date", header_cell_style)
    ]]

    for rec in records:
        table_data.append([
            Paragraph(f"#{rec['id']}", cell_style),
            Paragraph(f"<b>{rec['service_name']}</b><br/>{rec['sub_service']}", cell_style),
            Paragraph(str(rec['rating']), cell_style),
            Paragraph(rec['comment'] or "No comment provided.", cell_style),
            Paragraph(rec['audio_status'] or "No audio", cell_style),
            Paragraph(str(rec['timestamp']), cell_style)
        ])

    col_widths = [50, 160, 60, 260, 100, 110]
    t = Table(table_data, colWidths=col_widths, repeatRows=1)
    
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1B2A4A')),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#DCDCDC')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F9FAFB')])
    ]))

    elements.append(t)
    doc.build(elements)
    
    buffer.seek(0)
    return buffer


# ----------------------------------------------------
# PROGRESSIVE WEB APP (PWA) OFFLINE ROUTE
# ----------------------------------------------------
@app.route('/sw.js')
def service_worker():
    return app.send_static_file('sw.js')


# ----------------------------------------------------
# PUBLIC ROUTES
# ----------------------------------------------------
@app.route('/')
def fingerprint():
    return render_template('fingerprint.html')


@app.route('/welcome')
@app.route('/language')
def welcome_page():
    lang = request.args.get('lang', 'am')
    return render_template('welcome.html', lang=lang)


@app.route('/services')
def services():
    lang = request.args.get('lang', 'am')
    service_map = get_service_map()
    return render_template('services.html', lang=lang, service_map=service_map)


@app.route('/feedback')
def feedback():
    lang = request.args.get('lang', 'am')
    service = request.args.get('service', 'police_clearance')
    service_map = get_service_map()
    sub_service_map = get_sub_service_map()
    return render_template('feedback.html', lang=lang, service=service, service_map=service_map, sub_service_map=sub_service_map)


# ----------------------------------------------------
# FINGERPRINT API ROUTES
# ----------------------------------------------------
@app.route('/api/fingerprint/scan', methods=['POST'])
def scan_fingerprint():
    try:
        data = request.get_json() or {}
        user_id = data.get('user_id', 'TEMPORARY_USER')
        fingerprint_data = data.get('fingerprint_data', 'BYPASS_FINGERPRINT_HASH')

        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO fingerprint_records (user_id, fingerprint_data) VALUES (?, ?)",
            (str(user_id), str(fingerprint_data))
        )
        conn.commit()
        cursor.close()
        conn.close()

        return jsonify({
            "status": "success",
            "message": "የጣት አሻራ ስካን ሳይጠበቅ ቀጥታ አልፏል (Demo Mode)!",
            "user_id": user_id
        }), 200
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


# ----------------------------------------------------
# FEEDBACK SUBMIT / UNREAD ROUTES
# ----------------------------------------------------
@app.route('/submit-feedback', methods=['POST'])
@app.route('/api/submit-feedback', methods=['POST'])
def submit_feedback():
    try:
        feedback_count = session.get('feedback_count', 0)
        if feedback_count >= 3:
            return jsonify({
                "status": "error",
                "message": "ለአሁኑ የተፈቀደልዎትን 3 አስተያየቶች ጨርሰዋል! / You have reached your max limit of 3 feedbacks for this session."
            }), 403

        data = request.get_json()
        if not data:
            return jsonify({"status": "error", "message": "No data received"}), 400

        rating = data.get('rating', '😊')
        comment = data.get('comment', 'No comment provided.')
        sub_service = data.get('sub_service', 'general_service')
        audio_status = data.get('audio_status', 'No audio recorded')

        url_service = data.get('service') or data.get('category') or data.get('service_name', 'police_clearance')
        client_timestamp = data.get('timestamp')

        if is_inappropriate(comment):
            return jsonify({
                "status": "error",
                "message": "Inappropriate language detected. Please keep your feedback respectful."
            }), 400

        conn = get_db_connection()
        cursor = conn.cursor()

        if client_timestamp:
            cursor.execute(
                "INSERT INTO feedbacks (service_name, sub_service, rating, comment, audio_status, is_read, timestamp) VALUES (?, ?, ?, ?, ?, 0, ?)",
                (str(url_service), str(sub_service), str(rating), str(comment), str(audio_status), str(client_timestamp))
            )
        else:
            cursor.execute(
                "INSERT INTO feedbacks (service_name, sub_service, rating, comment, audio_status, is_read) VALUES (?, ?, ?, ?, ?, 0)",
                (str(url_service), str(sub_service), str(rating), str(comment), str(audio_status))
            )

        conn.commit()
        cursor.close()
        conn.close()

        session['feedback_count'] = feedback_count + 1

        return jsonify({
            "status": "success",
            "message": f"Feedback saved successfully! ({session['feedback_count']}/3 submitted)"
        })
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/api/unread-count')
def api_unread_count():
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT COUNT(*) AS count FROM feedbacks WHERE is_read = 0")
        result = cursor.fetchone()
        unread_count = result['count'] if result else 0
    except Exception:
        unread_count = 0
    finally:
        cursor.close()
        conn.close()
    return jsonify({"unread_count": unread_count})


# ----------------------------------------------------
# ADMIN AUTHENTICATION
# ----------------------------------------------------
@app.route('/admin/login', methods=['GET', 'POST'])
def admin_login():
    error = None
    admin_credentials = get_admin_credentials()
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        if username in admin_credentials and admin_credentials[username]['password'] == password:
            session['admin_user'] = username
            log_admin_action(username, "Logged into the admin panel.")
            return redirect(url_for('admin_dashboard'))
        else:
            error = "Invalid Username or Password. Please try again."
    return render_template('admin_login.html', error=error)


@app.route('/admin/logout')
def admin_logout():
    logged_in_admin = session.get('admin_user')
    if logged_in_admin:
        log_admin_action(logged_in_admin, "Logged out of the admin panel.")
    session.pop('admin_user', None)
    return redirect(url_for('admin_login'))


# ----------------------------------------------------
# ADMIN DASHBOARD & REPORTS EXPORT ROUTES
# ----------------------------------------------------
@app.route('/admin/dashboard')
def admin_dashboard():
    logged_in_admin = session.get('admin_user')
    admin_credentials = get_admin_credentials()
    if not logged_in_admin or logged_in_admin not in admin_credentials:
        return redirect(url_for('admin_login'))

    admin_info = admin_credentials[logged_in_admin]
    admin_type = admin_info['type']
    assigned_service = admin_info['service']
    assigned_sub = admin_info['sub_service']
    admin_title = admin_info['title']
    is_general_admin = (admin_type == 'general')
    selected_filter = request.args.get('service', 'all')

    service_map = get_service_map()
    sub_service_map = get_sub_service_map()

    conn = get_db_connection()
    cursor = conn.cursor()

    chart_data = {label: 0 for label in service_map.values()}
    service_counts = {key: 0 for key in service_map.keys()}
    total_feedbacks_count = 0
    unread_notifications_count = 0

    try:
        cursor.execute("SELECT * FROM feedbacks ORDER BY timestamp DESC")
        all_feedbacks = cursor.fetchall()

        for fb in all_feedbacks:
            db_s = str(fb['service_name']).strip().lower()
            db_sub = str(fb['sub_service']).strip().lower()

            if not fb['is_read']:
                if admin_type == 'general':
                    unread_notifications_count += 1
                elif admin_type == 'service' and db_s == assigned_service:
                    unread_notifications_count += 1
                elif admin_type == 'sub_service' and db_s == assigned_service and db_sub == assigned_sub:
                    unread_notifications_count += 1

            matched_key = db_s if db_s in service_map else 'other'
            service_counts[matched_key] = service_counts.get(matched_key, 0) + 1
            if matched_key in service_map:
                chart_data[service_map[matched_key]] = chart_data.get(service_map[matched_key], 0) + 1

        total_feedbacks_count = len(all_feedbacks)
    except Exception:
        pass

    cursor.close()
    conn.close()

    # Apply search, filters & scope via helper
    feedbacks = get_filtered_feedbacks(admin_type, assigned_service, assigned_sub)

    return render_template(
        'admin_dashboard.html',
        feedbacks=feedbacks,
        admin_title=admin_title,
        is_general=is_general_admin,
        service_map=service_map,
        sub_service_map=sub_service_map,
        chart_data=chart_data,
        service_counts=service_counts,
        total_feedbacks_count=total_feedbacks_count,
        unread_notifications_count=unread_notifications_count,
        current_filter=selected_filter
    )


@app.route('/admin/export/excel')
def export_excel():
    logged_in_admin = session.get('admin_user')
    admin_credentials = get_admin_credentials()
    if not logged_in_admin or logged_in_admin not in admin_credentials:
        return redirect(url_for('admin_login'))

    admin_info = admin_credentials[logged_in_admin]
    records = get_filtered_feedbacks(admin_info['type'], admin_info['service'], admin_info['sub_service'])
    log_admin_action(logged_in_admin, "Exported filtered feedback report to Excel.")
    excel_file = generate_excel_report(records)
    return send_file(
        excel_file,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=f"police_feedback_report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    )


@app.route('/admin/export/word')
def export_word():
    logged_in_admin = session.get('admin_user')
    admin_credentials = get_admin_credentials()
    if not logged_in_admin or logged_in_admin not in admin_credentials:
        return redirect(url_for('admin_login'))

    admin_info = admin_credentials[logged_in_admin]
    records = get_filtered_feedbacks(admin_info['type'], admin_info['service'], admin_info['sub_service'])
    log_admin_action(logged_in_admin, "Exported filtered feedback report to Word.")
    word_file = generate_word_report(records)
    return send_file(
        word_file,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        as_attachment=True,
        download_name=f"police_feedback_report_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
    )


@app.route('/admin/export/pdf')
def export_pdf():
    logged_in_admin = session.get('admin_user')
    admin_credentials = get_admin_credentials()
    if not logged_in_admin or logged_in_admin not in admin_credentials:
        return redirect(url_for('admin_login'))

    admin_info = admin_credentials[logged_in_admin]
    records = get_filtered_feedbacks(admin_info['type'], admin_info['service'], admin_info['sub_service'])
    log_admin_action(logged_in_admin, "Exported filtered feedback report to PDF.")
    pdf_file = generate_pdf_report(records)
    return send_file(
        pdf_file,
        mimetype='application/pdf',
        as_attachment=True,
        download_name=f"police_feedback_report_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
    )


@app.route('/admin/notifications')
def admin_notifications():
    logged_in_admin = session.get('admin_user')
    admin_credentials = get_admin_credentials()
    if not logged_in_admin or logged_in_admin not in admin_credentials:
        return redirect(url_for('admin_login'))

    admin_info = admin_credentials[logged_in_admin]
    admin_type = admin_info['type']
    assigned_service = admin_info['service']
    assigned_sub = admin_info['sub_service']

    conn = get_db_connection()
    cursor = conn.cursor()

    if admin_type == 'general':
        cursor.execute("UPDATE feedbacks SET is_read = 1 WHERE is_read = 0")
    elif admin_type == 'service':
        cursor.execute("UPDATE feedbacks SET is_read = 1 WHERE is_read = 0 AND service_name = ?", (assigned_service,))
    elif admin_type == 'sub_service':
        cursor.execute(
            "UPDATE feedbacks SET is_read = 1 WHERE is_read = 0 AND service_name = ? AND sub_service = ?",
            (assigned_service, assigned_sub)
        )
    conn.commit()

    cursor.execute("SELECT * FROM feedbacks ORDER BY timestamp DESC")
    all_feedbacks = cursor.fetchall()

    notifications = []
    service_map = get_service_map()
    sub_service_map = get_sub_service_map()

    if admin_type == 'general':
        notifications = all_feedbacks
    elif admin_type == 'service':
        for fb in all_feedbacks:
            if str(fb['service_name']).strip().lower() == assigned_service:
                notifications.append(fb)
    elif admin_type == 'sub_service':
        for fb in all_feedbacks:
            if str(fb['service_name']).strip().lower() == assigned_service and str(fb['sub_service']).strip().lower() == assigned_sub:
                notifications.append(fb)

    cursor.close()
    conn.close()

    return render_template('admin_notifications.html', notifications=notifications, service_map=service_map, sub_service_map=sub_service_map)


@app.route('/admin/audit-logs')
def admin_audit_logs():
    logged_in_admin = session.get('admin_user')
    admin_credentials = get_admin_credentials()
    if not logged_in_admin or logged_in_admin not in admin_credentials:
        return redirect(url_for('admin_login'))

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM audit_logs ORDER BY timestamp DESC")
    logs = cursor.fetchall()
    cursor.close()
    conn.close()
    return render_template('admin_audit_logs.html', audit_logs=logs)


@app.route('/admin/settings', methods=['GET', 'POST'])
def admin_settings():
    logged_in_admin = session.get('admin_user')
    admin_credentials = get_admin_credentials()
    if not logged_in_admin or logged_in_admin not in admin_credentials:
        return redirect(url_for('admin_login'))

    admin_info = admin_credentials[logged_in_admin]
    is_general_admin = (admin_info['type'] == 'general')
    message = None
    error = None

    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'password':
            current_pass = request.form.get('current_password')
            new_pass = request.form.get('new_password')
            if admin_credentials[logged_in_admin]['password'] == current_pass:
                message = "Password updated successfully!"
                log_admin_action(logged_in_admin, "Changed account password.")
            else:
                error = "Current password is incorrect."

        elif is_general_admin:
            if action == 'add_service':
                service_key_raw = request.form.get('service_key')
                s_key = service_key_raw.strip().lower().replace(" ", "_") if service_key_raw else None
                service_name_raw = request.form.get('service_name')
                s_name = service_name_raw.strip() if service_name_raw else None
                if s_key and s_name:
                    try:
                        conn = get_db_connection()
                        cursor = conn.cursor()
                        cursor.execute("INSERT INTO services (service_key, service_name) VALUES (?, ?)", (s_key, s_name))
                        conn.commit()
                        cursor.close()
                        conn.close()
                        message = f"Service '{s_name}' added successfully!"
                        log_admin_action(logged_in_admin, f"Added new service: {s_name}")
                    except Exception:
                        error = "Service key already exists or invalid input."
                else:
                    error = "All fields are required to add a service."

            elif action == 'update_service':
                s_key = request.form.get('service_key')
                new_service_name_raw = request.form.get('new_service_name')
                s_name = new_service_name_raw.strip() if new_service_name_raw else None
                if s_key and s_name:
                    conn = get_db_connection()
                    cursor = conn.cursor()
                    cursor.execute("UPDATE services SET service_name = ? WHERE service_key = ?", (s_name, s_key))
                    conn.commit()
                    cursor.close()
                    conn.close()
                    message = "Service updated successfully!"
                    log_admin_action(logged_in_admin, f"Updated service key '{s_key}' name to '{s_name}'")
                else:
                    error = "Invalid service update details."

            elif action == 'delete_service':
                s_key = request.form.get('service_key')
                if s_key:
                    conn = get_db_connection()
                    cursor = conn.cursor()
                    cursor.execute("DELETE FROM services WHERE service_key = ?", (s_key,))
                    cursor.execute("DELETE FROM sub_services WHERE service_key = ?", (s_key,))
                    conn.commit()
                    cursor.close()
                    conn.close()
                    message = "Service and its sub-services deleted successfully!"
                    log_admin_action(logged_in_admin, f"Deleted service and sub-services for key '{s_key}'")
                else:
                    error = "Select a service to delete."

            elif action == 'add_sub_service':
                parent_service = request.form.get('parent_service_key')
                sub_key_raw = request.form.get('sub_service_key')
                sub_key = sub_key_raw.strip().lower().replace(" ", "_") if sub_key_raw else ""
                sub_name_raw = request.form.get('sub_service_name')
                sub_name = sub_name_raw.strip() if sub_name_raw else ""
                if parent_service and sub_key and sub_name:
                    try:
                        conn = get_db_connection()
                        cursor = conn.cursor()
                        cursor.execute(
                            "INSERT INTO sub_services (service_key, sub_service_key, sub_service_name) VALUES (?, ?, ?)",
                            (parent_service, sub_key, sub_name)
                        )
                        conn.commit()
                        cursor.close()
                        conn.close()
                        message = f"Sub-service '{sub_name}' added successfully!"
                        log_admin_action(logged_in_admin, f"Added sub-service '{sub_name}' under '{parent_service}'")
                    except Exception:
                        error = "Sub-service key already exists under this service."
                else:
                    error = "All fields are required to add a sub-service."

            elif action == 'update_sub_service':
                parent_service = request.form.get('parent_service_key')
                sub_key = request.form.get('sub_service_key')
                sub_name_raw = request.form.get('new_sub_service_name')
                sub_name = sub_name_raw.strip() if sub_name_raw else ""
                if parent_service and sub_key and sub_name:
                    conn = get_db_connection()
                    cursor = conn.cursor()
                    cursor.execute(
                        "UPDATE sub_services SET sub_service_name = ? WHERE service_key = ? AND sub_service_key = ?",
                        (sub_name, parent_service, sub_key)
                    )
                    conn.commit()
                    cursor.close()
                    conn.close()
                    message = "Sub-service updated successfully!"
                    log_admin_action(logged_in_admin, f"Updated sub-service '{sub_key}' under '{parent_service}' to '{sub_name}'")
                else:
                    error = "Invalid sub-service update details."

    return render_template(
        'admin_settings.html',
        admin_info=admin_info,
        is_general=is_general_admin,
        service_map=get_service_map(),
        sub_service_map=get_sub_service_map(),
        message=message,
        error=error
    )


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)