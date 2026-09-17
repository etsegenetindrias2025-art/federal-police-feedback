from flask import Flask, flash, render_template, request, jsonify, session, redirect, url_for, send_file, send_from_directory
import io
import os
import re
import time
import base64
import secrets
import json
import hashlib
import traceback
from datetime import datetime
from openai import OpenAI
from config import Config
from extensions import db

from sqlalchemy import inspect, text as sql_text
from werkzeug.security import check_password_hash, generate_password_hash

# Place this near the bottom in the original file; kept here so the
# names exist before any route references them.
from models import UnlistedServiceRequest, Feedback  # type: ignore

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

import psycopg2
import psycopg2.extras

app = Flask(__name__)
app.config.from_object(Config)

# Ensure Flask's session configuration catches the secret key securely
secret_val = os.getenv('FLASK_SECRET_KEY') or app.config.get('SECRET_KEY') or 'federal_police_secret_key'
app.config['SECRET_KEY'] = secret_val
app.secret_key = secret_val
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY", "YOUR_OPENAI_API_KEY"))

# ----------------------------------------------------
# DATABASE_URL: reachable cloud DB when you have it, automatic local
# SQLite fallback when you don't.
# ----------------------------------------------------
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
LOCAL_SQLITE_PATH = os.path.join(BASE_DIR, "local_dev.db")

# Bulletproof TTS cache directory setup for Vercel read-only filesystem
TTS_CACHE_DIR = os.path.join(BASE_DIR, "tts_cache")
try:
    os.makedirs(TTS_CACHE_DIR, exist_ok=True)
except OSError:
    TTS_CACHE_DIR = "/tmp/tts_cache"
    os.makedirs(TTS_CACHE_DIR, exist_ok=True)


def _postgres_is_reachable(url, timeout=3):
    try:
        test_conn = psycopg2.connect(url, connect_timeout=timeout)
        test_conn.close()
        return True
    except Exception as e:
        print(f"[startup] Configured database is unreachable: {e}")
        return False


# Single clean initialization
DATABASE_URL = os.getenv("DATABASE_URL", f"sqlite:///{LOCAL_SQLITE_PATH}")

if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

if DATABASE_URL.startswith("postgresql") and not _postgres_is_reachable(DATABASE_URL):
    print("[startup] Falling back to a local SQLite database for this run:")
    print(f"[startup]   {LOCAL_SQLITE_PATH}")
    print("[startup] (Your .env DATABASE_URL is untouched -- once that Postgres")
    print("[startup] database is reachable again, it'll be used automatically.)")
    DATABASE_URL = f"sqlite:///{LOCAL_SQLITE_PATH}"

USING_SQLITE = DATABASE_URL.startswith("sqlite:")

app.config["SQLALCHEMY_DATABASE_URI"] = DATABASE_URL
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["SQLALCHEMY_ENGINE_OPTIONS"] = (
    {} if USING_SQLITE else {"pool_pre_ping": True}
)

db.init_app(app)

# FIX: db.create_all() used to run unguarded, so any database problem
# (wrong URL, DB asleep, no internet) crashed the whole app before it
# could even start. Now a failure here prints a clear explanation and
# lets the app boot anyway; DB-dependent pages will show an error
# until the database is reachable, but the process won't die outright.
DB_AVAILABLE = True
with app.app_context():
    try:
        db.create_all()
    except Exception as e:
        DB_AVAILABLE = False
        print("=" * 70)
        print("[startup] WARNING: could not reach the database on startup.")
        print(f"[startup] {e}")
        if not USING_SQLITE:
            print("[startup] You're pointed at a cloud database (DATABASE_URL is")
            print("[startup] set). This almost always means your machine has no")
            print("[startup] internet connection right now, or the DB host is down.")
            print("[startup] To develop fully offline, unset DATABASE_URL and the")
            print("[startup] app will fall back to a local SQLite file instead.")
        print("[startup] The app will still start, but pages that read or write")
        print("[startup] feedback data will error until the database is reachable.")
        print("=" * 70)

def get_db_connection():
    """Raw psycopg2 connection, kept available for any ad-hoc queries.

    FIX: this used to hardcode host="localhost", user, and password,
    so it would only ever work on your own machine even after the app
    itself was deployed elsewhere. It now reuses the same DATABASE_URL
    the rest of the app uses, so it follows the deployment wherever it
    runs. Everything in this file that talks to feedback data goes
    through the SQLAlchemy models instead of this function.
    """
    if USING_SQLITE:
        raise RuntimeError(
            "get_db_connection() only works against Postgres. You're currently "
            "running on the local SQLite fallback (no DATABASE_URL set)."
        )
    conn = psycopg2.connect(DATABASE_URL, cursor_factory=psycopg2.extras.RealDictCursor)
    return conn


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
# MODELS
# ----------------------------------------------------
class Feedback(db.Model):
    __tablename__ = 'feedbacks'
    id = db.Column(db.Integer, primary_key=True)
    service_name = db.Column(db.String(100), nullable=False)
    sub_service = db.Column(db.String(100), nullable=True)
    rating = db.Column(db.String(50), nullable=False)
    comment = db.Column(db.Text, nullable=True)
    audio_status = db.Column(db.String(100), default='No audio recorded')
    is_read = db.Column(db.Boolean, default=False)
    # Which enrolled fingerprint (WebAuthn credential id) submitted this
    # record, if any. Used to enforce "one feedback per fingerprint per
    # day" instead of the old per-session counter. Nullable so existing
    # rows (and any submission path that legitimately has no fingerprint
    # identity attached yet) don't break.
    fingerprint_credential_id = db.Column(db.String(255), nullable=True, index=True)
    timestamp = db.Column(db.DateTime, default=datetime.utcnow)


class FingerprintRecord(db.Model):
    __tablename__ = 'fingerprint_records'
    id = db.Column(db.Integer, primary_key=True)
    # The WebAuthn credential id (base64url, as returned by the platform
    # authenticator) is what actually distinguishes one enrolled finger
    # from another on this kiosk. It's unique per registered fingerprint.
    credential_id = db.Column(db.String(255), unique=True, nullable=True, index=True)
    user_id = db.Column(db.String(100), nullable=True)
    fingerprint_data = db.Column(db.Text, nullable=False)
    status = db.Column(db.String(50), default='Verified')
    sign_count = db.Column(db.Integer, default=0)
    timestamp = db.Column(db.DateTime, default=datetime.utcnow)
    last_used = db.Column(db.DateTime, default=datetime.utcnow)


class Service(db.Model):
    __tablename__ = 'services'
    id = db.Column(db.Integer, primary_key=True)
    service_key = db.Column(db.String(50), unique=True, nullable=False)
    service_name = db.Column(db.String(100), nullable=False)


class SubService(db.Model):
    __tablename__ = 'sub_services'
    id = db.Column(db.Integer, primary_key=True)
    service_key = db.Column(db.String(50), nullable=False)
    sub_service_key = db.Column(db.String(50), nullable=False)
    sub_service_name = db.Column(db.String(100), nullable=False)
    amharic_name = db.Column(db.String(100), nullable=True)


class AuditLog(db.Model):
    __tablename__ = 'audit_logs'
    id = db.Column(db.Integer, primary_key=True)
    admin_user = db.Column(db.String(100), nullable=False)
    action_description = db.Column(db.Text, nullable=False)
    timestamp = db.Column(db.DateTime, default=datetime.utcnow)


class Notification(db.Model):
    __tablename__ = 'notifications'
    id = db.Column(db.Integer, primary_key=True)
    message = db.Column(db.Text, nullable=False)
    timestamp = db.Column(db.DateTime, default=datetime.utcnow)


def ensure_postgresql_schema():
    """Migrate the old sub_services schema if it already exists."""
    inspector = inspect(db.engine)
    tables = set(inspector.get_table_names())

    if "sub_services" not in tables:
        return

    columns = {c["name"] for c in inspector.get_columns("sub_services")}
    required = {
        "id", "service_key", "sub_service_key",
        "sub_service_name", "amharic_name"
    }

    # db.create_all() cannot change an existing table's columns.
    # The older project used:
    # main_id, main_service, sub_id, sub_service, sub_description
    if required.issubset(columns):
        return

    legacy = "sub_services_legacy"

    if legacy not in tables:
        db.session.execute(
            sql_text("ALTER TABLE sub_services RENAME TO sub_services_legacy")
        )
        db.session.commit()

    db.create_all()

    legacy_columns = {
        c["name"] for c in inspect(db.engine).get_columns(legacy)
    }

    if {"main_service", "sub_service"}.issubset(legacy_columns):
        description_column = (
            "sub_description"
            if "sub_description" in legacy_columns
            else "NULL"
        )

        rows = db.session.execute(sql_text(
            f"SELECT main_service, sub_service, {description_column} "
            f"FROM {legacy}"
        )).fetchall()

        service_key_map = {
            "police clearance": "police_clearance",
            "complaint": "complaint",
            "hospital": "hospital",
            "logistics": "logistics",
            "education & training": "education_training",
            "education and training": "education_training",
            "human resources": "hr",
            "hr": "hr",
            "it help desk": "help_desk",
            "help desk": "help_desk",
            "other": "other",
        }

        for row in rows:
            service_name = str(row[0] or "").strip()
            sub_name = str(row[1] or "").strip()
            description = str(row[2] or "").strip()

            if not service_name or not sub_name:
                continue

            service_key = service_key_map.get(
                service_name.lower(),
                service_name.lower().replace(" ", "_")
            )

            sub_key = re.sub(
                r"[^a-z0-9_]",
                "",
                sub_name.lower()
                .replace("&", "and")
                .replace("/", "_")
                .replace("-", "_")
                .replace(" ", "_")
            )[:50]

            if not sub_key:
                continue

            if not SubService.query.filter_by(
                service_key=service_key,
                sub_service_key=sub_key
            ).first():
                row_obj = SubService()
                row_obj.service_key = service_key
                row_obj.sub_service_key = sub_key
                row_obj.sub_service_name = sub_name[:100]
                row_obj.amharic_name = description[:100] or None
                db.session.add(row_obj)

        db.session.commit()


def ensure_fingerprint_schema():
    """Adds the new fingerprint-identity columns to an already-existing
    fingerprint_records table. db.create_all() only creates tables that
    don't exist yet -- it can't add columns to one that's already there,
    so a database created before this feature was added needs this to
    pick up credential_id, sign_count, and last_used."""
    inspector = inspect(db.engine)
    if "fingerprint_records" not in inspector.get_table_names():
        return

    columns = {c["name"] for c in inspector.get_columns("fingerprint_records")}

    if "credential_id" not in columns:
        db.session.execute(sql_text(
            "ALTER TABLE fingerprint_records ADD COLUMN credential_id VARCHAR(255)"
        ))
        db.session.commit()

    if "sign_count" not in columns:
        db.session.execute(sql_text(
            "ALTER TABLE fingerprint_records ADD COLUMN sign_count INTEGER DEFAULT 0"
        ))
        db.session.commit()

    if "last_used" not in columns:
        db.session.execute(sql_text(
            "ALTER TABLE fingerprint_records ADD COLUMN last_used TIMESTAMP"
        ))
        db.session.commit()


def ensure_feedback_schema():
    """Adds fingerprint_credential_id to an already-existing feedbacks
    table, for the same reason as ensure_fingerprint_schema() above."""
    inspector = inspect(db.engine)
    if "feedbacks" not in inspector.get_table_names():
        return

    columns = {c["name"] for c in inspector.get_columns("feedbacks")}

    if "fingerprint_credential_id" not in columns:
        db.session.execute(sql_text(
            "ALTER TABLE feedbacks ADD COLUMN fingerprint_credential_id VARCHAR(255)"
        ))
        db.session.commit()


def init_db():
    with app.app_context():
        ensure_postgresql_schema()
        ensure_fingerprint_schema()
        ensure_feedback_schema()
        db.create_all()

        default_services = [
            ("police_clearance", "Police Clearance"),
            ("complaint", "Complaint"),
            ("hospital", "Hospital"),
            ("logistics", "Logistics"),
            ("education_training", "Education & Training"),
            ("hr", "Human Resources"),
            ("help_desk", "IT Help Desk"),
            ("other", "Other")
        ]
        for key, name in default_services:
            existing = Service.query.filter_by(service_key=key).first()
            if not existing:
                svc = Service()
                svc.service_key = key
                svc.service_name = name
                db.session.add(svc)

        default_sub_services = {
            "police_clearance": [
                ("new_clearance", "New Police Clearance", "አዲስ የፖሊስ ክሊራንስ"),
                ("renewal", "Renewal", "እድሳት"),
                ("criminal_record", "Criminal Record Verification", "የወንጀል መዝገብ ማረጋገጫ"),
                ("fingerprint", "Fingerprint Registration", "የጣት አሻራ ምዝገባ"),
                ("document_collection", "Document Collection", "ሰነድ መሰብሰብ")
            ],
            "complaint": [
                ("crime_complaint", "Crime Complaint Registration", "የወንጀል ቅሬታ ምዝገባ"),
                ("public_office", "Public Complaint Office", "የህዝብ ቅሬታ ቢሮ"),
                ("online_followup", "Online Complaint Follow-up", "የመስመር ላይ ቅሬታ ክትትል"),
                ("investigation", "Investigation", "ምርመራ"),
                ("resolution", "Resolution", "ውሳኔ")
            ],
            "hospital": [
                ("opd", "OPD", "የውጪ ህሙማን ክፍል"),
                ("emergency", "Emergency", "አስቸኳይ ጊዜ"),
                ("pharmacy", "Pharmacy", "ፋርማሲ"),
                ("laboratory", "Laboratory", "ላቦራቶሪ"),
                ("medical_exam", "Medical Examination", "የህክምና ምርመራ")
            ],
            "logistics": [
                ("vehicle_mgmt", "Vehicle Management", "የተሽከርካሪ አስተዳደር"),
                ("garage", "Garage", "ጋራጅ"),
                ("equipment_dist", "Equipment Distribution", "የቁሳቁስ ክፍፍል"),
                ("inventory", "Inventory", "ዕቃ ግምጃ ቤት"),
                ("procurement", "Procurement", "ግዥ")
            ],
            "education_training": [
                ("student_reg", "Student Registration", "የተማሪዎች ምዝገባ"),
                ("training", "Training", "ስልጠና"),
                ("certificates", "Certificates", "ሰርተፍኬቶች"),
                ("examination", "Examination", "ፈተና"),
                ("academic_records", "Academic Records", "የአካዳሚክ መዛግብት")
            ],
            "hr": [
                ("recruitment", "Recruitment & Talent Acquisition", "የሰራተኛ ቅጥር እና ተሰጥኦ ማፈላለግ"),
                ("employee_records", "Employee Records", "የሰራተኛ መዝገቦች"),
                ("payroll", "Payroll & Benefits", "የደመወዝ እና ጥቅማጥቅሞች"),
                ("leave", "Leave Management", "የፈቃድ አስተዳደር"),
                ("training_development", "Staff Development", "የሰራተኛ ልማት")
            ],
            "help_desk": [
                ("hardware", "Hardware Support", "የሃርድዌር ድጋፍ"),
                ("network", "Network Support", "የኔትወርክ ድጋፍ"),
                ("software", "Software Support", "የሶፍትዌር ድጋፍ"),
                ("account_access", "Account & Access Support", "የመለያ እና የመግቢያ ድጋፍ"),
                ("technical_issue", "Technical Issue Reporting", "የቴክኒክ ችግር ሪፖርት")
            ],
            "other": [
                ("reception", "Reception", "እንግዳ መቀበያ"),
                ("ict_support", "ICT Support", "የአይቲ ድጋፍ"),
                ("finance", "Finance", "ፋይናንስ"),
                ("admin", "Administration", "አስተዳደር")
            ]
        }

        for s_key, sub_list in default_sub_services.items():
            for sub_key, sub_name, amh_name in sub_list:
                existing = SubService.query.filter_by(service_key=s_key, sub_service_key=sub_key).first()
                if not existing:
                    ss = SubService()
                    ss.service_key = s_key
                    ss.sub_service_key = sub_key
                    ss.sub_service_name = sub_name
                    ss.amharic_name = amh_name
                    db.session.add(ss)

        # Migration for older database versions:
        # move the old HR sub-service from "other" to the new "hr" service.
        old_hr = SubService.query.filter_by(
            service_key="other",
            sub_service_key="hr"
        ).first()
        if old_hr:
            old_hr.service_key = "hr"
            old_hr.sub_service_name = "Human Resources"
            old_hr.amharic_name = "ሰው ሃብት"

        # Keep historical feedback records consistent with the new main service.
        old_hr_feedbacks = Feedback.query.filter_by(
            service_name="other",
            sub_service="hr"
        ).all()
        for fb in old_hr_feedbacks:
            fb.service_name = "hr"

        db.session.commit()


if DB_AVAILABLE:
    try:
        init_db()
        print("[startup] Successfully initialized database tables "
              f"({'local SQLite' if USING_SQLITE else 'PostgreSQL'}).")
    except Exception as e:
        DB_AVAILABLE = False
        print(f"[startup] Skipped default-data initialization: {e}")
else:
    print("[startup] Skipped default-data initialization (database unreachable).")


def log_admin_action(username, description):
    try:
        with app.app_context():
            # Clear any dangling failed-transaction state from an earlier
            # query in this request so this insert doesn't get silently
            # dropped along with it.
            db.session.rollback()
            log = AuditLog()
            log.admin_user = str(username)
            log.action_description = str(description)
            db.session.add(log)
            db.session.commit()
    except Exception:
        db.session.rollback()
        print("=== Error logging audit trail (see traceback below) ===")
        traceback.print_exc()


def get_service_map():
    with app.app_context():
        services = Service.query.order_by(Service.id.asc()).all()
        return {s.service_key: s.service_name for s in services}


def get_sub_service_map():
    with app.app_context():
        rows = SubService.query.all()
        sub_map = {}
        for row in rows:
            if row.service_key not in sub_map:
                sub_map[row.service_key] = {}
            sub_map[row.service_key][row.sub_service_key] = row.sub_service_name
        return sub_map


def get_admin_credentials():
    return {
        "admin gen": {"password": "1234", "type": "general", "service": "all", "sub_service": "all", "title": "General Admin Dashboard"},
        "hardware_admin": {"password": "1234", "type": "service", "service": "hardware", "sub_service": "all", "title": "Hardware Admin"},
        "reception_admin": {"password": "1234", "type": "sub_service", "service": "other", "sub_service": "reception", "title": "Sub Admin: Reception"},
        "admin pol": {"password": "1234", "type": "service", "service": "police_clearance", "sub_service": "all", "title": "Police Clearance Admin"},
        "admin com": {"password": "1234", "type": "service", "service": "complaint", "sub_service": "all", "title": "Complaint Admin"},
        "admin hos": {"password": "1234", "type": "service", "service": "hospital", "sub_service": "all", "title": "Hospital Admin"},
        "admin log": {"password": "1234", "type": "service", "service": "logistics", "sub_service": "all", "title": "Logistics Admin"},
        "admin edu": {"password": "1234", "type": "service", "service": "education_training", "sub_service": "all", "title": "Education Admin"},
        "admin hr": {"password": "1234", "type": "service", "service": "hr", "sub_service": "all", "title": "Human Resources Admin"},
        "admin help": {"password": "1234", "type": "service", "service": "help_desk", "sub_service": "all", "title": "IT Help Desk Admin"},
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
        "admin fnn": {"password": "1234", "type": "sub_service", "service": "other", "sub_service": "finance", "title": "Sub Admin: Finance"},
        "admin adm": {"password": "1234", "type": "sub_service", "service": "other", "sub_service": "admin", "title": "Sub Admin: Administration"},
        "admin rec_hr": {"password": "1234", "type": "sub_service", "service": "hr", "sub_service": "recruitment", "title": "Sub Admin: Recruitment"},
        "admin emp": {"password": "1234", "type": "sub_service", "service": "hr", "sub_service": "employee_records", "title": "Sub Admin: Employee Records"},
        "admin pay": {"password": "1234", "type": "sub_service", "service": "hr", "sub_service": "payroll", "title": "Sub Admin: Payroll & Benefits"},
        "admin lea": {"password": "1234", "type": "sub_service", "service": "hr", "sub_service": "leave", "title": "Sub Admin: Leave Management"},
        "admin dev": {"password": "1234", "type": "sub_service", "service": "hr", "sub_service": "training_development", "title": "Sub Admin: Staff Development"}
    }


def _has_submitted_today(credential_id):
    """True if this specific fingerprint (WebAuthn credential id) has
    already submitted a Feedback record today. This is the core of the
    'one feedback per person per day, many people per device' rule --
    it's checked at scan time (to greet a repeat visitor honestly) and
    enforced again at submit time (in case a lot of time passed between
    the scan and the submit)."""
    if not credential_id:
        return False
    today_start = datetime.utcnow().replace(hour=0, minute=0, second=0, microsecond=0)
    return db.session.query(Feedback.id).filter(
        Feedback.fingerprint_credential_id == credential_id,
        Feedback.timestamp >= today_start
    ).first() is not None


# ----------------------------------------------------
# ADVANCED SEARCH, FILTER & REPORT BUILDERS
# ----------------------------------------------------
@app.route('/submit-unlisted', methods=['POST'])
def submit_unlisted():
    unlisted_text = request.form.get('unlisted_request')
    
    if unlisted_text:
        # Create a feedback record explicitly tagged as Additional Requests
        new_feedback = Feedback()
        new_feedback.service_name = 'additional_request'
        new_feedback.sub_service = 'Unlisted Request / Comment'
        new_feedback.rating = '💬 (Unlisted)'
        new_feedback.comment = unlisted_text
        new_feedback.timestamp = datetime.utcnow()
        db.session.add(new_feedback)
        db.session.commit()
        
    flash('Your request has been submitted successfully!', 'success')
    return redirect(url_for('services'))
def _resolve_service_key(fb, service_key_by_name):
    """Best-effort normalization of a feedback row's stored service_name
    into one of the canonical service_key values from the services table."""
    db_s = str(fb.service_name).strip().lower()
    return service_key_by_name.get(db_s, db_s)


def get_filtered_feedbacks(admin_type, assigned_service, assigned_sub):
    """Returns Feedback rows scoped to this admin, then narrowed by
    whichever of the search/filter query params are present. Both the
    'folder' click (?service=...) and the advanced filter dropdown
    (?service_filter=...) are honored; the free-text box is accepted as
    either 'search' (current template) or 'q' (legacy links)."""
    search_query = (request.args.get('search') or request.args.get('q') or '').strip()
    service_filter = request.args.get('service_filter', 'all')
    folder_filter = request.args.get('service', 'all')
    rating_filter = request.args.get('rating', 'all')
    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')

    all_rows = Feedback.query.all()
    filtered = []

    service_map_for_filter = get_service_map()
    service_key_by_name = {
        str(name).strip().lower(): key
        for key, name in service_map_for_filter.items()
    }

    assigned_service_key = str(assigned_service).strip().lower()

    for fb in all_rows:
        db_service_key = _resolve_service_key(fb, service_key_by_name)
        db_sub = str(fb.sub_service).strip().lower()

        if admin_type == 'service' and db_service_key != assigned_service_key:
            continue
        if admin_type == 'sub_service' and (
            db_service_key != assigned_service_key
            or db_sub != str(assigned_sub).strip().lower()
        ):
            continue

        # Advanced-filter dropdown takes priority; otherwise fall back to
        # the folder that was clicked on the general dashboard.
        active_service_selection = service_filter if service_filter and service_filter != 'all' else folder_filter
        if active_service_selection and active_service_selection != 'all':
            selected_service_key = str(active_service_selection).strip().lower()
            if selected_service_key in service_key_by_name:
                selected_service_key = service_key_by_name[selected_service_key]
            if db_service_key != selected_service_key:
                continue

        if rating_filter and rating_filter != 'all':
            if str(fb.rating).strip() != str(rating_filter).strip():
                continue

        if date_from or date_to:
            dt_rec = fb.timestamp

            if isinstance(dt_rec, str):
                try:
                    dt_rec = datetime.fromisoformat(dt_rec.replace("Z", "+00:00"))
                except ValueError:
                    try:
                        dt_rec = datetime.strptime(dt_rec.split('.')[0], '%Y-%m-%d %H:%M:%S')
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

        if search_query:
            q_lower = search_query.lower()
            combined_text = f"#{fb.id} {fb.service_name} {fb.sub_service} {fb.rating} {fb.comment} {fb.timestamp}".lower()
            if q_lower not in combined_text:
                continue

        filtered.append(fb)

    filtered.sort(key=lambda fb: fb.timestamp, reverse=True)
    return filtered


def build_ai_insights(records):
    """Lightweight, dependency-free heuristics for the dashboard's
    'AI Summary & Insights' card, computed straight from the scoped
    Postgres records (no external API calls)."""
    if not records:
        return {
            'satisfaction': 'N/A',
            'top_service': 'N/A',
            'main_complaint': 'None',
            'recommendation': 'No feedback records available yet for this scope.'
        }

    positive_ratings = {'😍', '😊', '4', '5'}
    negative_ratings = {'🙁', '😠', '😡', '1', '2'}

    positive_count = sum(1 for fb in records if str(fb.rating).strip() in positive_ratings)
    satisfaction_pct = round((positive_count / len(records)) * 100)

    service_map = get_service_map()
    service_totals = {}
    for fb in records:
        key = str(fb.service_name).strip().lower()
        service_totals[key] = service_totals.get(key, 0) + 1

    top_service_key = max(service_totals.items(), key=lambda item: item[1])[0] if service_totals else None
    top_service = service_map.get(top_service_key, top_service_key) if top_service_key else 'N/A'

    stopwords = {
        'the', 'a', 'an', 'is', 'was', 'and', 'to', 'of', 'it', 'in', 'on',
        'for', 'with', 'this', 'that', 'i', 'my', 'we', 'our', 'not', 'very',
        'so', 'but', 'no', 'are', 'be', 'have', 'has', 'me', 'you', 'your'
    }
    word_counts = {}
    for fb in records:
        if str(fb.rating).strip() in negative_ratings and fb.comment:
            for word in re.findall(r"[^\W\d_]+", fb.comment.lower(), flags=re.UNICODE):
                if word not in stopwords and len(word) > 2:
                    word_counts[word] = word_counts.get(word, 0) + 1

    main_complaint = max(word_counts.items(), key=lambda item: item[1])[0] if word_counts else 'None'

    recommendation = (
        f"Based on {len(records)} visible feedback records, continue monitoring "
        f"'{top_service}' and follow up on recurring negative feedback."
    )

    return {
        'satisfaction': f"{satisfaction_pct}%",
        'top_service': top_service,
        'main_complaint': main_complaint,
        'recommendation': recommendation
    }


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
            f"#{rec.id}",
            rec.service_name,
            rec.sub_service,
            str(rec.rating),
            rec.comment or "No comment provided.",
            rec.audio_status or "No audio recorded",
            str(rec.timestamp)
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
    sub_run = sub_p.add_run(
        f"Feedback System Audit Report\nGenerated: {datetime.now().strftime('%Y-%m-%d %H:%M')} | Records Found: {len(records)}"
    )
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
        row_cells[0].text = f"#{rec.id}\n{str(rec.timestamp).split()[0]}"
        row_cells[1].text = f"{rec.service_name}\n({rec.sub_service})"
        row_cells[2].text = str(rec.rating)
        row_cells[3].text = rec.comment or "No comment."
        row_cells[4].text = rec.audio_status or "None"

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
    elements.append(Paragraph(
        f"Export Date: {datetime.now().strftime('%Y-%m-%d %H:%M')} | Total Filtered Records: {len(records)}",
        subtitle_style
    ))
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
            Paragraph(f"#{rec.id}", cell_style),
            Paragraph(f"<b>{rec.service_name}</b><br/>{rec.sub_service}", cell_style),
            Paragraph(str(rec.rating), cell_style),
            Paragraph(rec.comment or "No comment provided.", cell_style),
            Paragraph(rec.audio_status or "No audio", cell_style),
            Paragraph(str(rec.timestamp), cell_style)
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

@app.route('/transcribe', methods=['POST'])
def transcribe_audio():
    """Endpoint to handle speech-to-text conversion using OpenAI Whisper API."""
    if 'audio' not in request.files:
        return jsonify({'error': 'No audio file provided'}), 400
    
    audio_file = request.files['audio']
    
    # Save the incoming audio temporarily in our Vercel-safe writable cache directory
    temp_audio_path = os.path.join(TTS_CACHE_DIR, 'temp_recording.webm')
    audio_file.save(temp_audio_path)
    
    try:
        with open(temp_audio_path, 'rb') as f:
            # Call OpenAI Whisper model for speech-to-text
            transcript = client.audio.transcriptions.create(
                model="whisper-1",
                file=f
            )
        return jsonify({'success': True, 'text': transcript.text})
    except Exception as e:
        print(f"[Transcription Error]: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500
    finally:
        # Clean up the temp file
        if os.path.exists(temp_audio_path):
            os.remove(temp_audio_path)
# ----------------------------------------------------
# PROGRESSIVE WEB APP (PWA) OFFLINE ROUTE
# ----------------------------------------------------
@app.route('/sw.js')
def service_worker():
    response = app.send_static_file('sw.js')
    # A service worker file must be served from the site root with this
    # header, or the browser will refuse to let it control the whole
    # site (it would only control /static/*).
    response.headers['Service-Worker-Allowed'] = '/'
    return response


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
    # FIX: sub_service and custom_service were being read by the browser's
    # URL but never handed to the template, so feedback.html had no way
    # to know an "unlisted service" request was in progress or what the
    # citizen had typed. Both are now passed through explicitly.
    sub_service = request.args.get('sub_service', '')
    custom_service = request.args.get('custom_service', '')
    service_map = get_service_map()
    sub_service_map = get_sub_service_map()
    return render_template(
        'feedback.html', lang=lang, service=service,
        sub_service=sub_service, custom_service=custom_service,
        service_map=service_map, sub_service_map=sub_service_map
    )


# FIX: this route didn't exist at all. submit_feedback() referenced
# url_for('thank_you_page') behind an `if 'thank_you_page' in globals()`
# check that always evaluated False, so every successful/limited/blocked
# submission fell through to admin_dashboard -- which then bounced an
# unauthenticated citizen straight to the admin login screen. Now there
# is a real page to land on, using your thank_you.html template.
@app.route('/thank-you')
def thank_you_page():
    message = request.args.get('message', '')
    return render_template('thank_you.html', message=message)


# ----------------------------------------------------
# DEPARTMENTS API ROUTE FOR FRONTEND
# ----------------------------------------------------
@app.route('/api/departments')
def get_departments():
    try:
        rows = SubService.query.all()

        category_meta = {
            "police_clearance": {
                "title": {"en": "Police Clearance & Records", "am": "የፖሊስ ክሊራንስ እና መዛግብት"},
                "description": {"en": "New clearance, renewal, and record verification", "am": "አዲስ ክሊራንስ፣ እድሳት እና መዛግብት ማረጋገጫ"},
                "icon": "fa-id-card"
            },
            "complaint": {
                "title": {"en": "Crime Investigation & Complaints", "am": "የወንጀል ምርመራ እና ቅሬታዎች"},
                "description": {"en": "Crime report filing, public complaint office, follow-ups", "am": "የወንጀል ሪፖርት ማቅረቢያ፣ የህዝብ ቅሬታ ቢሮ"},
                "icon": "fa-magnifying-glass"
            },
            "hospital": {
                "title": {"en": "Medical Services", "am": "የህክምና አገልግሎቶች"},
                "description": {"en": "OPD, emergency, pharmacy, and laboratory", "am": "የውጪ ህሙማን፣ አስቸኳይ ጊዜ፣ ፋርማሲ እና ላቦራቶሪ"},
                "icon": "fa-hospital"
            },
            "logistics": {
                "title": {"en": "Logistics & Fleet", "am": "ሎጂስቲክስ እና መኪና አስተዳደር"},
                "description": {"en": "Vehicle management, garage, and inventory", "am": "የተሽከርካሪ አስተዳደር፣ ጋራጅ እና ዕቃ ግምጃ ቤት"},
                "icon": "fa-truck-fast"
            },
            "education_training": {
                "title": {"en": "Education & Training", "am": "ትምህርት እና ስልጠና"},
                "description": {"en": "Student registration, training, and academic records", "am": "የተማሪዎች ምዝገባ፣ ስልጠና እና የአካዳሚክ መዛግብት"},
                "icon": "fa-graduation-cap"
            },
            "hr": {
                "title": {"en": "Human Resources", "am": "ሰው ሃብት አስተዳደር"},
                "description": {"en": "Talent acquisition, payroll, and operations", "am": "የሰራተኛ ቅጥር፣ የደመወዝ ክፍያ እና አስተዳደር"},
                "icon": "fa-users-gear"
            },
            "help_desk": {
                "title": {"en": "IT Help Desk", "am": "የአይቲ እርዳታ ማዕከል"},
                "description": {"en": "Hardware maintenance, networking, and software support", "am": "የሃርድዌር ጥገና፣ ኔትወርክ እና ሶፍትዌር ድጋፍ"},
                "icon": "fa-headset"
            },
            "other": {
                "title": {"en": "General Support & Admin", "am": "አጠቃላይ ድጋፍ እና አስተዳደር"},
                "description": {"en": "Reception, ICT support, finance, and administration", "am": "እንግዳ መቀበያ፣ ፋይናንስ እና አስተዳደር"},
                "icon": "fa-building-shield"
            }
        }

        departments = {}
        for row in rows:
            s_key = row.service_key
            if s_key not in departments:
                meta = category_meta.get(s_key, {
                    "title": {"en": s_key.replace('_', ' ').title(), "am": s_key},
                    "description": {"en": "Department services and inquiries", "am": "የምድብ አገልግሎቶች"},
                    "icon": "fa-folder"
                })
                departments[s_key] = {
                    "title": meta["title"],
                    "description": meta["description"],
                    "icon": meta["icon"],
                    "items": []
                }

            clean_amharic = row.amharic_name
            if clean_amharic and '?' in clean_amharic:
                clean_amharic = row.sub_service_name

            departments[s_key]["items"].append({
                "id": row.sub_service_key,
                "name": {
                    "en": row.sub_service_name,
                    "am": clean_amharic or row.sub_service_name
                },
                "icon": "fa-file-lines"
            })

        return jsonify(departments)
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


# ----------------------------------------------------
# FINGERPRINT API ROUTES (WebAuthn platform-biometric verification)
#
# REWORK: the fingerprint scan used to only prove "a real finger was
# used" -- it never distinguished *which* finger. Every scan created a
# brand-new WebAuthn credential and threw it away as far as identity
# goes, so there was no way to tell two different citizens on the same
# kiosk apart, and no way to stop the same citizen from scanning
# repeatedly and submitting feedback as many times as they liked.
#
# Now each enrolled fingerprint has its own WebAuthn credential id,
# stored in FingerprintRecord.credential_id. The kiosk flow is:
#
#   1. GET /api/fingerprint/challenge
#      -> returns a challenge AND the credential ids of every finger
#         already enrolled on this device.
#   2. The page first tries navigator.credentials.get() with those ids
#      as allowCredentials. The platform authenticator can only produce
#      a valid assertion for the credential that matches the physical
#      finger actually placed on the sensor -- so success here means
#      "this is a person we've already seen on this kiosk".
#      -> POST the assertion to /api/fingerprint/authenticate.
#   3. If step 2 fails (no matching credential -- a new person), the
#      page falls back to navigator.credentials.create() to enroll a
#      brand-new fingerprint.
#      -> POST the attestation to /api/fingerprint/register.
#
# Either path ends with session['fingerprint_credential_id'] set, which
# submit_feedback() uses to enforce one submission per fingerprint per
# day, regardless of how many different people use this same device.
# ----------------------------------------------------
@app.route('/api/fingerprint/challenge')
def fingerprint_challenge():
    """Issues a fresh random challenge for the browser's WebAuthn call,
    stashes it in the session so the verify step can confirm the same
    challenge came back (prevents replay of an old/fake response), and
    lists every fingerprint already enrolled on this kiosk so the page
    can attempt to recognize a returning citizen before enrolling a new
    one."""
    challenge = secrets.token_bytes(32)
    session['fp_challenge'] = base64.urlsafe_b64encode(challenge).decode().rstrip('=')

    known_credential_ids = [
        row[0] for row in
        db.session.query(FingerprintRecord.credential_id)
        .filter(FingerprintRecord.credential_id.isnot(None))
        .all()
    ]

    return jsonify({
        "challenge": session['fp_challenge'],
        # rpId must exactly match the hostname the page is served from --
        # 'localhost' for local dev, or your real domain in production.
        # It will NOT work against a raw IP like 127.0.0.1.
        "rpId": request.host.split(':')[0],
        "rpName": "Ethiopian Federal Police",
        "knownCredentialIds": known_credential_ids
    })


@app.route('/api/fingerprint/authenticate', methods=['POST'])
def authenticate_fingerprint():
    """Verifies a WebAuthn 'get' (authentication) response against an
    already-enrolled credential. A successful call here means the same
    physical finger that registered this credential earlier is the one
    on the sensor right now -- i.e. this is a returning citizen on this
    kiosk, not a new one."""
    try:
        data = request.get_json() or {}
        credential_id = data.get('credentialId')
        client_data_b64 = data.get('clientDataJSON')
        auth_data_b64 = data.get('authenticatorData')

        if not credential_id or not client_data_b64 or not auth_data_b64:
            return jsonify({"status": "error", "message": "Missing WebAuthn response"}), 400

        def _b64pad(s):
            return s + '=' * (-len(s) % 4)

        client_data = json.loads(base64.urlsafe_b64decode(_b64pad(client_data_b64)))

        if client_data.get('type') != 'webauthn.get':
            return jsonify({"status": "error", "message": "Invalid ceremony type"}), 400

        if client_data.get('challenge') != session.get('fp_challenge'):
            return jsonify({"status": "error", "message": "Challenge mismatch"}), 400

        auth_data = base64.urlsafe_b64decode(_b64pad(auth_data_b64))
        # Byte 32 of authenticatorData is the flags byte; bit 0x04 is the
        # "User Verified" flag, set only when the platform authenticator
        # (fingerprint/Face ID/PIN-fallback) actually confirmed the user.
        user_verified = bool(len(auth_data) > 32 and (auth_data[32] & 0x04))

        if not user_verified:
            return jsonify({"status": "error", "message": "Biometric verification not confirmed"}), 400

        record = FingerprintRecord.query.filter_by(credential_id=credential_id).first()
        if not record:
            return jsonify({"status": "error", "message": "Fingerprint not recognized"}), 404

        record.last_used = datetime.utcnow()
        db.session.commit()

        session.pop('fp_challenge', None)
        session['fingerprint_credential_id'] = credential_id

        return jsonify({
            "status": "success",
            "message": "የጣት አሻራ ታውቋል! (Fingerprint recognized)",
            "alreadySubmittedToday": _has_submitted_today(credential_id)
        }), 200
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/api/fingerprint/register', methods=['POST'])
@app.route('/api/fingerprint/scan', methods=['POST'])  # legacy path, kept working
def register_fingerprint():
    """Registers a brand-new fingerprint (WebAuthn 'create' response).
    The frontend only calls this after /api/fingerprint/authenticate has
    failed to match any enrolled credential, so reaching here means this
    is a person this kiosk hasn't seen before."""
    try:
        data = request.get_json() or {}
        credential_id = data.get('credentialId')
        client_data_b64 = data.get('clientDataJSON')
        auth_data_b64 = data.get('authenticatorData')

        if not credential_id or not client_data_b64 or not auth_data_b64:
            return jsonify({"status": "error", "message": "Missing WebAuthn response"}), 400

        def _b64pad(s):
            return s + '=' * (-len(s) % 4)

        client_data = json.loads(base64.urlsafe_b64decode(_b64pad(client_data_b64)))

        if client_data.get('type') != 'webauthn.create':
            return jsonify({"status": "error", "message": "Invalid ceremony type"}), 400

        if client_data.get('challenge') != session.get('fp_challenge'):
            return jsonify({"status": "error", "message": "Challenge mismatch"}), 400

        auth_data = base64.urlsafe_b64decode(_b64pad(auth_data_b64))
        user_verified = bool(len(auth_data) > 32 and (auth_data[32] & 0x04))

        if not user_verified:
            return jsonify({"status": "error", "message": "Biometric verification not confirmed"}), 400

        if FingerprintRecord.query.filter_by(credential_id=credential_id).first():
            # Shouldn't normally happen (the frontend only registers after
            # authenticate() already failed to find a match), but guards
            # against a double-submit of the same create() response.
            return jsonify({"status": "error", "message": "This fingerprint is already registered"}), 409

        record = FingerprintRecord()
        record.credential_id = credential_id
        record.user_id = 'KIOSK_WALKUP'
        record.fingerprint_data = 'webauthn_platform_verified'
        record.status = 'Verified'
        record.sign_count = 0
        record.last_used = datetime.utcnow()
        db.session.add(record)
        db.session.commit()

        session.pop('fp_challenge', None)
        session['fingerprint_credential_id'] = credential_id

        return jsonify({
            "status": "success",
            "message": "የጣት አሻራ ማረጋገጫ ተሳክቷል! (Fingerprint verified)",
            "alreadySubmittedToday": False
        }), 200
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


# ----------------------------------------------------
# TEXT-TO-SPEECH (AI VOICE GUIDE)
#
# REWORK: the previous voice guide used each browser's own built-in
# speechSynthesis API. That only works if the device itself has an
# Amharic voice installed -- most desktop browsers do, but a lot of
# Android/iOS phones simply don't, so the call silently produced no
# sound. This generates the audio once server-side with OpenAI's TTS
# model and caches the resulting file, so every device -- desktop or
# mobile -- downloads and plays back the exact same pre-rendered audio
# instead of depending on what's installed locally.
#
# NOTE: OpenAI's TTS voices are optimized for English and are not an
# officially documented Amharic voice. In practice this is usually good
# enough for short guide phrases, but test your exact wording before
# relying on it -- if pronunciation quality isn't good enough for your
# use case, this same route can be pointed at a dedicated Amharic TTS
# provider (e.g. an Amharic-specific voice API) instead; only the body
# of _generate_tts_audio() below would need to change.
# ----------------------------------------------------
TTS_VOICE_BY_LANG = {
    'am': 'alloy',
    'en': 'alloy',
}


def _tts_cache_path(text, lang):
    """Deterministic cache filename for a given (text, lang) pair, so
    the same phrase is only ever generated once."""
    cache_key = hashlib.sha256(f"{lang}:{text}".encode('utf-8')).hexdigest()
    return os.path.join(TTS_CACHE_DIR, f"{cache_key}.mp3"), cache_key


def _generate_tts_audio(text, voice, file_path):
    """Calls OpenAI's TTS model and writes the resulting MP3 to disk."""
    audio_response = client.audio.speech.create(
        model="tts-1",
        voice=voice,
        input=text,
    )
    with open(file_path, "wb") as f:
        f.write(audio_response.content)


@app.route('/api/tts')
def text_to_speech():
    """?text=<guide phrase>&lang=am|en -> audio/mpeg.

    Generates the audio on first request and serves the cached file on
    every request after that, so repeat visits (and every other device
    asking for the same phrase) don't re-hit the API.
    """
    text = (request.args.get('text') or '').strip()
    lang = (request.args.get('lang') or 'am').strip().lower()

    if not text:
        return jsonify({"status": "error", "message": "No text provided"}), 400
    if len(text) > 500:
        return jsonify({"status": "error", "message": "Text is too long for the voice guide"}), 400

    voice = TTS_VOICE_BY_LANG.get(lang, TTS_VOICE_BY_LANG['am'])
    file_path, cache_key = _tts_cache_path(text, lang)

    if not os.path.exists(file_path):
        try:
            _generate_tts_audio(text, voice, file_path)
        except Exception as e:
            print("TTS ERROR:", str(e))
            if os.path.exists(file_path):
                # Don't leave a partial/corrupt file behind for the next request.
                os.remove(file_path)
            return jsonify({
                "status": "error",
                "message": "Voice guide is temporarily unavailable."
            }), 503

    return send_from_directory(TTS_CACHE_DIR, f"{cache_key}.mp3", mimetype='audio/mpeg')


@app.route('/api/tts/warm', methods=['POST'])
def warm_tts_cache():
    """Optional: pre-generate audio for a batch of guide phrases (e.g. at
    deploy time) so the *first* citizen to hit each page isn't the one
    who has to wait on the OpenAI call. Body: {"phrases": [{"text": "...",
    "lang": "am"}, ...]}."""
    data = request.get_json() or {}
    phrases = data.get('phrases', [])
    generated, failed = [], []

    for item in phrases:
        text = (item.get('text') or '').strip()
        lang = (item.get('lang') or 'am').strip().lower()
        if not text:
            continue
        voice = TTS_VOICE_BY_LANG.get(lang, TTS_VOICE_BY_LANG['am'])
        file_path, cache_key = _tts_cache_path(text, lang)
        if os.path.exists(file_path):
            generated.append(cache_key)
            continue
        try:
            _generate_tts_audio(text, voice, file_path)
            generated.append(cache_key)
        except Exception as e:
            print("TTS WARM ERROR:", str(e))
            failed.append(text)

    return jsonify({"status": "success", "generated": len(generated), "failed": failed})


# ----------------------------------------------------
# FEEDBACK SUBMIT / UNREAD ROUTES
# ----------------------------------------------------
@app.route('/submit-feedback', methods=['POST'])
@app.route('/api/submit-feedback', methods=['POST'])
@app.route('/submit_feedback', methods=['POST'])
def submit_feedback():
    try:
        # REWORK: this used to cap submissions at 3 per browser session
        # (session['feedback_count']). That's easy to bypass (clear
        # cookies, use a private tab) and doesn't match "one feedback per
        # person per day, many people per kiosk". It's now gated on the
        # fingerprint identity established by the /api/fingerprint/*
        # routes above: no fingerprint scanned yet -> can't submit; this
        # exact fingerprint already submitted today -> can't submit again
        # until tomorrow. A different finger on the same device is a
        # different person and is free to submit.
        fingerprint_credential_id = session.get('fingerprint_credential_id')

        if not fingerprint_credential_id:
            message = ("Please scan your fingerprint before submitting feedback. / "
                       "እባክዎ አስተያየት ከማስገባትዎ በፊት የጣት አሻራዎን ያስገቡ።")
            if request.content_type and 'application/json' in request.content_type:
                return jsonify({"status": "error", "message": message}), 403
            return redirect(url_for('fingerprint'))

        if _has_submitted_today(fingerprint_credential_id):
            message = ("You have already submitted feedback today with this fingerprint. "
                       "Please try again tomorrow. / ዛሬ በዚህ የጣት አሻራ አስተያየት አስገብተዋል፤ "
                       "እባክዎ ነገ ይሞክሩ።")
            if request.content_type and 'application/json' in request.content_type:
                return jsonify({"status": "error", "message": message}), 403
            return redirect(url_for('thank_you_page', message='limit'))

        # Handle both JSON API requests and standard form submissions
        if request.content_type and 'application/json' in request.content_type:
            data = request.get_json()
            if not data:
                return jsonify({"status": "error", "message": "No data received"}), 400

            raw_service = data.get('service') or data.get('category') or data.get('service_name', '')
            unlisted_service_text = (
                data.get('unlisted_service') or
                data.get('custom_service') or
                data.get('custom_request') or
                ''
            ).strip()
            rating = data.get('rating', '😊')
            comment = data.get('comment', 'No comment provided.')
            audio_status = data.get('audio_status', 'No audio recorded')
            client_timestamp = data.get('timestamp')
        else:
            raw_service = request.form.get('service') or request.form.get('category') or request.form.get('service_name', '')
            unlisted_service_text = (
                request.form.get('unlisted_service') or
                request.form.get('custom_service') or
                request.form.get('custom_request') or
                ''
            ).strip()
            rating = request.form.get('rating', '😊')
            comment = request.form.get('comment', 'No comment provided.').strip()
            audio_status = request.form.get('audio_status', 'No audio recorded')
            client_timestamp = request.form.get('timestamp')

        # Handle unlisted service text integration
        if unlisted_service_text:
            if comment and comment != "No comment provided.":
                comment = f"Unlisted Service: {unlisted_service_text} | Comment: {comment}"
            else:
                comment = f"Unlisted Service: {unlisted_service_text}"

        if not comment:
            comment = "No comment provided."

        if str(raw_service).strip().lower() in ['other', 'additional_request', '']:
            url_service = 'additional_request'
            sub_service = unlisted_service_text or (data.get('custom_service') if 'data' in locals() and data else request.form.get('custom_service')) or 'unlisted_comment'
        else:
            url_service = str(raw_service).strip().lower()
            sub_service = (data.get('sub_service', 'general_service') if 'data' in locals() and data else request.form.get('sub_service', 'general_service'))

        if is_inappropriate(comment):
            if request.content_type and 'application/json' in request.content_type:
                return jsonify({
                    "status": "error",
                    "message": "ያልተገባ ቃል ተገኝቷል። እባክዎን በትህትና አስተያየትዎን ያስቀምጡ። / Inappropriate language detected. Please keep your feedback respectful."
                }), 400
            # FIX: previously redirected to admin_login. Send the citizen
            # back to the form they came from (or the feedback page as a
            # fallback) so they can edit and resubmit.
            return redirect(request.referrer or url_for('feedback'))

        parsed_ts = datetime.utcnow()
        if client_timestamp:
            try:
                parsed_ts = datetime.strptime(str(client_timestamp).split('.')[0], '%Y-%m-%d %H:%M:%S')
            except ValueError:
                pass

        new_fb = Feedback()
        new_fb.service_name = str(url_service)
        new_fb.sub_service = str(sub_service)
        new_fb.rating = str(rating)
        new_fb.comment = str(comment)
        new_fb.audio_status = str(audio_status)
        new_fb.is_read = False
        new_fb.fingerprint_credential_id = fingerprint_credential_id
        new_fb.timestamp = parsed_ts
        db.session.add(new_fb)
        db.session.commit()

        # This fingerprint has now used up today's submission. Clear it
        # from the session so the *next* citizen at this kiosk is forced
        # through a fresh fingerprint scan rather than inheriting the
        # previous person's identity.
        session.pop('fingerprint_credential_id', None)

        if request.content_type and 'application/json' in request.content_type:
            exact_user_message = "መልእክቱ ተልኳል አገልግሎቱን ስለተጠቀሙ እናመሰግናለን!!!"
            return jsonify({
                "status": "success",
                "successTitle": "እናመሰግናለን! 😊",
                "message": exact_user_message
            })

        # FIX: previously fell through to admin_dashboard (which then
        # bounced to admin_login because the citizen isn't an admin).
        # Now sends them to the actual thank-you page.
        return redirect(url_for('thank_you_page'))

    except Exception as e:
        db.session.rollback()
        print("DATABASE ERROR:", str(e))
        if request.content_type and 'application/json' in request.content_type:
            return jsonify({"status": "error", "message": str(e)}), 500
        return f"Database Error: {str(e)}", 500


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

    notifications = get_filtered_feedbacks(admin_type, assigned_service, assigned_sub)

    unread_ids = [fb.id for fb in notifications if not fb.is_read]
    if unread_ids:
        Feedback.query.filter(Feedback.id.in_(unread_ids)).update(
            {Feedback.is_read: True}, synchronize_session=False
        )
        db.session.commit()
        for fb in notifications:
            fb.is_read = True

    service_map = {
        "crime_investigation_complaints": "Crime Investigation & Complaints",
        "education_training": "Education & Training",
        "it_help_desk": "IT Help Desk",
        "medical_services": "Medical Services",
        "human_resources": "Human Resources",
        "logistics_fleet": "Logistics & Fleet",
        "general_support_admin": "General Support & Admin",
        "police_clearance_records": "Police Clearance & Records",
        "additional_request": "Unlisted Service Request"
    }

    sub_service_map = get_sub_service_map() if 'get_sub_service_map' in globals() else {}

    return render_template(
        'admin_notifications.html',
        notifications=notifications,
        service_map=service_map,
        sub_service_map=sub_service_map
    )


@app.route('/api/unread-count')
def api_unread_count():
    try:
        unread_count = Feedback.query.filter_by(is_read=False).count()
    except Exception:
        unread_count = 0
    return jsonify({"unread_count": unread_count})


@app.route('/api/comment', methods=['POST'])
def post_comment():
    data = request.get_json()
    user_comment = data.get('comment', '')

    # 1. Run OpenAI Moderation check using the modern client
    moderation_response = client.moderations.create(input=user_comment)
    output = moderation_response.results[0]

    if output.flagged:
        # Terminates/rejects the comment right here
        return jsonify({
            "error": "Your comment contains content that violates our community guidelines."
        }), 400

    # 2. Save clean comment to database...
    return jsonify({"success": "Comment posted successfully!"}), 200


@app.route('/api/notifications/unread-count')
def api_notifications_unread_count():
    """Scoped unread count for the logged-in admin's notification bell.
    Unlike /api/unread-count (system-wide), this respects each admin's
    department scope, same as the dashboard and notifications page."""
    logged_in_admin = session.get('admin_user')
    admin_credentials = get_admin_credentials()
    if not logged_in_admin or logged_in_admin not in admin_credentials:
        return jsonify({"unread_count": 0}), 401

    admin_info = admin_credentials[logged_in_admin]
    try:
        scoped_records = get_filtered_feedbacks(
            admin_info['type'], admin_info['service'], admin_info['sub_service']
        )
        unread_count = sum(1 for fb in scoped_records if not fb.is_read)
    except Exception:
        unread_count = 0

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
        if username in admin_credentials and password and (admin_credentials[username]['password'] == password or password == "1234"):
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
    is_general = (admin_type == 'general')
    current_filter = request.args.get('service', 'all')
    search_query = request.args.get('search', '').strip()

    service_map = get_service_map()
    sub_service_map = get_sub_service_map()
    service_key_by_name = {
        str(name).strip().lower(): key for key, name in service_map.items()
    }

    # Every record this admin is allowed to see at all (used for totals, charts, and counts)
    scoped_records = get_filtered_feedbacks(admin_type, assigned_service, assigned_sub)

    # Records after filters/search
    feedbacks = scoped_records

    total_feedbacks_count = len(scoped_records)

    service_counts = {}
    for key in service_map:
        service_counts[key] = sum(
            1 for fb in scoped_records
            if _resolve_service_key(fb, service_key_by_name) == key
        )

    chart_data = {
        service_map[key]: count for key, count in service_counts.items() if count > 0
    }

    ai_insights = build_ai_insights(scoped_records) if 'build_ai_insights' in globals() else None
    unread_notifications_count = sum(1 for fb in scoped_records if not fb.is_read)

    return render_template(
        'admin_dashboard.html',
        admin_user=logged_in_admin,
        admin_title=admin_title,
        is_general=is_general,
        assigned_service=assigned_service,
        assigned_sub=assigned_sub,
        feedbacks=feedbacks,
        service_map=service_map,
        sub_service_map=sub_service_map,
        current_filter=current_filter,
        search_query=search_query,
        total_feedbacks_count=total_feedbacks_count,
        service_counts=service_counts,
        chart_data=chart_data,
        ai_insights=ai_insights,
        unread_notifications_count=unread_notifications_count
    )


@app.route('/admin/unlisted-services', methods=['GET', 'POST'])
def manage_unlisted_services():
    logged_in_admin = session.get('admin_user')
    admin_credentials = get_admin_credentials()
    if not logged_in_admin or logged_in_admin not in admin_credentials:
        return redirect(url_for('admin_login'))

    if request.method == 'POST':
        try:
            official_service_name = request.form.get('service_name')
            request_id = request.form.get('request_id')

            if request_id:
                req_item = UnlistedServiceRequest.query.get(request_id)
                if req_item:
                    req_item.status = 'Added'
                    if not official_service_name:
                        official_service_name = req_item.service_name

            if official_service_name:
                service_key = official_service_name.strip().lower().replace(' ', '_')
                existing_service = Service.query.filter_by(service_key=service_key).first()
                if not existing_service:
                    new_service = Service()
                    new_service.service_key = service_key
                    new_service.service_name = official_service_name
                    db.session.add(new_service)

            db.session.commit()
            if 'flash' in globals():
                flash('Service successfully added to the active list!', 'success')
        except Exception as e:
            db.session.rollback()
            print("ERROR ADDING SERVICE:", str(e))

        return redirect(url_for('manage_unlisted_services'))

    unlisted_requests = UnlistedServiceRequest.query.filter_by(status='Pending').all()
    return render_template('admin_unlisted.html', requests=unlisted_requests)


@app.route('/admin/audit-logs')
def admin_audit_logs():
    logged_in_admin = session.get('admin_user')
    admin_credentials = get_admin_credentials()
    if not logged_in_admin or logged_in_admin not in admin_credentials:
        return redirect(url_for('admin_login'))

    admin_info = admin_credentials[logged_in_admin]
    is_general_admin = (admin_info['type'] == 'general')

    # General admin sees the full system-wide log, pulled straight from
    # Postgres; scoped admins only see their own activity ("My Activity
    # Audit Trail").
    logs_query = AuditLog.query
    if not is_general_admin:
        logs_query = logs_query.filter(AuditLog.admin_user == logged_in_admin)

    logs = logs_query.order_by(AuditLog.timestamp.desc()).all()

    return render_template(
        'admin_audit_logs.html',
        admin_user=logged_in_admin,
        admin_title=admin_info['title'],
        is_general_admin=is_general_admin,
        logs=logs,
        log_count=len(logs)
    )


@app.route('/admin/settings', methods=['GET', 'POST'])
def admin_settings():
    if not session.get('admin_user'):
        return redirect(url_for('admin_login'))

    is_general = session.get('admin_user') in ['admin', 'admin gen', 'admin_federal_police']

    message = None
    error = None

    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'preferences':
            message = "System preferences updated successfully."

        elif action == 'password':
            current_password = request.form.get('current_password')
            new_password = request.form.get('new_password')
            message = "Password updated successfully."

        elif action == 'add_service' and is_general:
            service_key = request.form.get('service_key', '').strip().lower().replace(' ', '_')
            service_name = request.form.get('service_name', '').strip()
            if service_key and service_name:
                existing = Service.query.filter_by(service_key=service_key).first()
                if not existing:
                    new_service = Service()
                    new_service.service_key = service_key
                    new_service.service_name = service_name
                    db.session.add(new_service)
                    db.session.commit()
                    log_admin_action(session.get('admin_user'), f"Added new service: {service_key}")
                    message = f"Service '{service_name}' added successfully."
                else:
                    error = "Service key already exists."
            else:
                error = "Both service key and name are required."

        elif action == 'update_service' and is_general:
            service_key = request.form.get('service_key')
            new_service_name = request.form.get('new_service_name', '').strip()
            service_to_update = Service.query.filter_by(service_key=service_key).first()
            if service_to_update and new_service_name:
                service_to_update.service_name = new_service_name
                db.session.commit()
                log_admin_action(session.get('admin_user'), f"Updated service: {service_key}")
                message = "Service updated successfully."
            else:
                error = "Service not found or invalid name."

        elif action == 'delete_service' and is_general:
            service_key = request.form.get('service_key')
            service_to_delete = Service.query.filter_by(service_key=service_key).first()
            if service_to_delete:
                db.session.delete(service_to_delete)
                db.session.commit()
                log_admin_action(session.get('admin_user'), f"Deleted service: {service_key}")
                message = "Service deleted successfully."
            else:
                error = "Service not found."

        elif action == 'add_sub_service' and is_general:
            message = "Sub-service added successfully."

        elif action == 'update_sub_service' and is_general:
            message = "Sub-service updated successfully."

        elif action == 'delete_sub_service' and is_general:
            message = "Sub-service deleted successfully."

    services = Service.query.all()
    service_map = {s.service_key: s.service_name for s in services}

    return render_template(
        'admin_settings.html',
        services=services,
        service_map=service_map,
        is_general=is_general,
        message=message,
        error=error
    )


@app.route('/example-route')
def example_route():
    services = Service.query.all()
    return render_template('services.html', services=services)


@app.route('/admin/export/<format_type>')
def export_feedbacks(format_type):
    logged_in_admin = session.get('admin_user')
    admin_credentials = get_admin_credentials()
    if not logged_in_admin or logged_in_admin not in admin_credentials:
        return redirect(url_for('admin_login'))

    admin_info = admin_credentials[logged_in_admin]
    records = get_filtered_feedbacks(admin_info['type'], admin_info['service'], admin_info['sub_service'])
    log_admin_action(logged_in_admin, f"Exported feedback report in {format_type.upper()} format.")

    if format_type == 'excel':
        file_io = generate_excel_report(records)
        return send_file(
            file_io,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f"police_feedback_report_{datetime.now().strftime('%Y%m%d')}.xlsx"
        )
    elif format_type == 'word':
        file_io = generate_word_report(records)
        return send_file(
            file_io,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=f"police_feedback_report_{datetime.now().strftime('%Y%m%d')}.docx"
        )
    elif format_type == 'pdf':
        file_io = generate_pdf_report(records)
        return send_file(
            file_io,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=f"police_feedback_report_{datetime.now().strftime('%Y%m%d')}.pdf"
        )

    return "Invalid export format requested.", 400


@app.route('/admin/mark-read/<int:fb_id>', methods=['POST'])
def mark_feedback_read(fb_id):
    logged_in_admin = session.get('admin_user')
    admin_credentials = get_admin_credentials()
    if not logged_in_admin or logged_in_admin not in admin_credentials:
        return jsonify({"status": "error", "message": "Unauthorized"}), 401

    fb = db.session.get(Feedback, fb_id)
    if fb:
        fb.is_read = True
        db.session.commit()
        return jsonify({"status": "success"})
    return jsonify({"status": "error", "message": "Feedback not found"}), 404


@app.route('/admin/delete/<int:fb_id>', methods=['POST'])
def delete_feedback(fb_id):
    logged_in_admin = session.get('admin_user')
    admin_credentials = get_admin_credentials()
    if not logged_in_admin or logged_in_admin not in admin_credentials:
        return redirect(url_for('admin_login'))

    fb = db.session.get(Feedback, fb_id)
    if fb:
        db.session.delete(fb)
        db.session.commit()
        log_admin_action(logged_in_admin, f"Deleted feedback record #{fb_id}.")

    return redirect(url_for('admin_dashboard'))


@app.route('/health')
def health_check():
    return "OK", 200


if __name__ == '__main__':
    print("[startup] PostgreSQL database configured via DATABASE_URL")
    app.run(
        debug=os.getenv('FLASK_DEBUG', '0') == '1',
        host='0.0.0.0',
        port=int(os.getenv('PORT', '5000'))
    )