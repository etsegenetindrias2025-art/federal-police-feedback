"""
Ethiopian Federal Police - Citizen Feedback System

Render-ready Flask application.

Local development:
    python app.py

Render production start command:
    gunicorn --bind 0.0.0.0:$PORT app:app

Public deployment URL:
    https://ethiopian-federal-police-feedback.onrender.com
"""

from flask import Flask, render_template, request, jsonify, session, redirect, url_for, send_file
import sqlite3
import os
import io
import re
import qrcode
from gtts import gTTS
import socket
import time
from datetime import datetime, timedelta

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

# ---------------------------------------------------------------------------
# DEPLOYMENT / SECURITY CONFIGURATION
# ---------------------------------------------------------------------------
# On Render, set FLASK_SECRET_KEY to a long random value in the Environment
# settings. The fallback below is only for local development.
app.secret_key = os.environ.get('FLASK_SECRET_KEY', 'dev-only-federal-police-secret-key-change-me')

# Public URL used by QR codes/links when the deployed address is needed.
# Render provides the real host automatically, but this value can be overridden
# with PUBLIC_BASE_URL in the Render Environment settings.
PUBLIC_BASE_URL = os.environ.get(
    'PUBLIC_BASE_URL',
    'https://ethiopian-federal-police-feedback.onrender.com'
).rstrip('/')

# Secure cookies are enabled for the HTTPS deployment and disabled locally.
app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'
app.config['SESSION_COOKIE_SECURE'] = PUBLIC_BASE_URL.startswith('https://')
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(minutes=30)

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
# SQLite database location.
# Local development: ./database.db
# Render with a persistent disk mounted at /var/data: /var/data/database.db
# If /var/data is not available, the app safely falls back to the project folder.
render_data_dir = '/var/data'
if os.path.isdir(render_data_dir) and os.access(render_data_dir, os.W_OK):
    DB_PATH = os.path.join(render_data_dir, 'database.db')
else:
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

    conn.commit()
    cursor.close()
    conn.close()


init_db()
print(f"[startup] Using database file at: {DB_PATH}")  # confirm which database.db this run reads/writes


def log_admin_action(username, description):
    """Insert one row into audit_logs. Errors are printed WITH a full
    traceback (not just str(e)) so a failed insert is never silent —
    check your terminal if entries stop appearing."""
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
    credentials = {
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
    return credentials


ADMIN_PASSWORD_OVERRIDES = {}


def get_current_admin_credentials():
    """Return admin credentials including runtime password changes."""
    credentials = get_admin_credentials()
    for username, password in ADMIN_PASSWORD_OVERRIDES.items():
        if username in credentials:
            credentials[username]['password'] = password
    return credentials


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
                "INSERT INTO feedbacks (service_name, sub_service, rating, comment, is_read, timestamp) VALUES (?, ?, ?, ?, 0, ?)",
                (str(url_service), str(sub_service), str(rating), str(comment), str(client_timestamp))
            )
        else:
            cursor.execute(
                "INSERT INTO feedbacks (service_name, sub_service, rating, comment, is_read) VALUES (?, ?, ?, ?, 0)",
                (str(url_service), str(sub_service), str(rating), str(comment))
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
    admin_credentials = get_current_admin_credentials()
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
# ADMIN DASHBOARD
# ----------------------------------------------------
@app.route('/admin/dashboard')
def admin_dashboard():
    logged_in_admin = session.get('admin_user')
    admin_credentials = get_current_admin_credentials()
    if not logged_in_admin or logged_in_admin not in admin_credentials:
        return redirect(url_for('admin_login'))

    admin_info = admin_credentials[logged_in_admin]
    admin_type = admin_info['type']
    assigned_service = admin_info['service']
    assigned_sub = admin_info['sub_service']
    admin_title = admin_info['title']
    is_general_admin = (admin_type == 'general')

    # Filters (service/rating/date range) applied on top of the
    # role-based visibility rules below.
    selected_filter = request.args.get('service', 'all')
    rating_filter = request.args.get('rating_filter', '')
    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')

    service_map = get_service_map()
    sub_service_map = get_sub_service_map()

    conn = get_db_connection()
    cursor = conn.cursor()

    chart_data = {label: 0 for label in service_map.values()}
    service_counts = {key: 0 for key in service_map.keys()}
    total_feedbacks_count = 0
    unread_notifications_count = 0
    feedbacks = []

    try:
        cursor.execute("SELECT * FROM feedbacks ORDER BY timestamp DESC")
        all_feedbacks = cursor.fetchall()

        # Parse optional date-range bounds once, up front.
        dt_from = None
        dt_to_end = None
        if date_from:
            try:
                dt_from = datetime.strptime(date_from, '%Y-%m-%d')
            except ValueError:
                dt_from = None
        if date_to:
            try:
                dt_to_end = datetime.strptime(date_to, '%Y-%m-%d') + timedelta(days=1) - timedelta(seconds=1)
            except ValueError:
                dt_to_end = None

        def passes_extra_filters(fb):
            if rating_filter and str(fb['rating']) != rating_filter:
                return False
            if dt_from or dt_to_end:
                try:
                    fb_dt = datetime.fromisoformat(str(fb['timestamp']))
                except (TypeError, ValueError):
                    return False
                if dt_from and fb_dt < dt_from:
                    return False
                if dt_to_end and fb_dt > dt_to_end:
                    return False
            return True

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

        if admin_type == 'general':
            for fb in all_feedbacks:
                if selected_filter != 'all' and str(fb['service_name']).strip().lower() != selected_filter:
                    continue
                if not passes_extra_filters(fb):
                    continue
                feedbacks.append(fb)
            total_feedbacks_count = len(all_feedbacks)
        elif admin_type == 'service':
            for fb in all_feedbacks:
                if str(fb['service_name']).strip().lower() != assigned_service:
                    continue
                if not passes_extra_filters(fb):
                    continue
                feedbacks.append(fb)
            total_feedbacks_count = len(feedbacks)
        elif admin_type == 'sub_service':
            for fb in all_feedbacks:
                if str(fb['service_name']).strip().lower() != assigned_service or str(fb['sub_service']).strip().lower() != assigned_sub:
                    continue
                if not passes_extra_filters(fb):
                    continue
                feedbacks.append(fb)
            total_feedbacks_count = len(feedbacks)
    except Exception:
        feedbacks = []

    cursor.close()
    conn.close()

    # --------------------------------------------------------
    # AI INSIGHTS
    # Calculated from the feedback records visible to this admin.
    # --------------------------------------------------------
    total_count = len(feedbacks)
    if total_count > 0:
        positive_ratings = sum(
            1 for fb in feedbacks
            if str(fb['rating']).strip() in {'4', '5', '😊', '😍'}
        )
        satisfaction_pct = f"{int((positive_ratings / total_count) * 100)}%"

        visible_service_counts = {}
        for fb in feedbacks:
            service_name = str(fb['service_name'] or 'General')
            visible_service_counts[service_name] = visible_service_counts.get(service_name, 0) + 1
        # get the service with highest visible count in a type-safe way
        top_service = max(visible_service_counts.items(), key=lambda kv: kv[1])[0]

        # Use the lowest-rated visible feedback as a simple complaint indicator.
        negative_feedback = [
            fb for fb in feedbacks
            if str(fb['rating']).strip() in {'1', '2', '😡', '😞', '😐'}
        ]
        if negative_feedback:
            main_complaint = str(negative_feedback[0]['comment'] or 'Negative feedback received.')
        else:
            main_complaint = 'None reported'

        ai_insights = {
            'satisfaction': satisfaction_pct,
            'top_service': top_service,
            'main_complaint': main_complaint,
            'recommendation': (
                f"Based on {total_count} visible feedback records, continue monitoring "
                f"service quality and response times for {top_service}."
            ),
        }
    else:
        ai_insights = {
            'satisfaction': 'N/A',
            'top_service': 'N/A',
            'main_complaint': 'None',
            'recommendation': 'No feedback records available to generate insights.',
        }

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
        ai_insights=ai_insights,
        unread_notifications_count=unread_notifications_count,
        current_filter=selected_filter,
        rating_filter=rating_filter,
        date_from=date_from,
        date_to=date_to
    )


@app.route('/admin/notifications')
def admin_notifications():
    logged_in_admin = session.get('admin_user')
    admin_credentials = get_current_admin_credentials()
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


@app.route('/admin/settings', methods=['GET', 'POST'])
def admin_settings():
    logged_in_admin = session.get('admin_user')
    admin_credentials = get_current_admin_credentials()
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
                if not new_pass or len(new_pass) < 4:
                    error = "New password must contain at least 4 characters."
                else:
                    ADMIN_PASSWORD_OVERRIDES[logged_in_admin] = new_pass
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

            elif action == 'delete_sub_service':
                parent_service = request.form.get('parent_service_key')
                sub_key = request.form.get('sub_service_key')
                if parent_service and sub_key:
                    conn = get_db_connection()
                    cursor = conn.cursor()
                    cursor.execute("DELETE FROM sub_services WHERE service_key = ? AND sub_service_key = ?", (parent_service, sub_key))
                    conn.commit()
                    cursor.close()
                    conn.close()
                    message = "Sub-service deleted successfully!"
                    log_admin_action(logged_in_admin, f"Deleted sub-service '{sub_key}' under '{parent_service}'")
                else:
                    error = "Select a sub-service to delete."
        else:
            error = "Unauthorized action: Only General Admin can perform CRUD operations on services."

    service_map = get_service_map()
    sub_service_map = get_sub_service_map()
    return render_template(
        'admin_settings.html',
        message=message,
        error=error,
        service_map=service_map,
        sub_service_map=sub_service_map,
        is_general=is_general_admin
    )


@app.route('/admin/export/<format_type>')
def export_report(format_type):
    logged_in_admin = session.get('admin_user')
    admin_credentials = get_current_admin_credentials()
    if not logged_in_admin or logged_in_admin not in admin_credentials:
        return redirect(url_for('admin_login'))

    admin_info = admin_credentials[logged_in_admin]
    admin_type = admin_info['type']
    assigned_service = admin_info['service']
    assigned_sub = admin_info['sub_service']

    # Optional filters from the query string — combined with the
    # role-based visibility rules below, not a replacement for them.
    service_filter = request.args.get('service_filter', '')
    rating_filter = request.args.get('rating_filter', '')
    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')

    conn = get_db_connection()
    cursor = conn.cursor()

    query = "SELECT * FROM feedbacks WHERE 1=1"
    params = []

    # Role-based visibility: non-general admins are hard-scoped to their
    # own service/sub-service regardless of what's in the query string.
    if admin_type == 'general':
        if service_filter:
            query += " AND service_name = ?"
            params.append(service_filter)
    elif admin_type == 'service':
        query += " AND service_name = ?"
        params.append(assigned_service)
    elif admin_type == 'sub_service':
        query += " AND service_name = ? AND sub_service = ?"
        params.extend([assigned_service, assigned_sub])
    else:
        cursor.close()
        conn.close()
        return "Unauthorized.", 403

    if rating_filter:
        query += " AND rating = ?"
        params.append(rating_filter)
    if date_from:
        query += " AND timestamp >= ?"
        params.append(date_from)
    if date_to:
        query += " AND timestamp <= ?"
        params.append(date_to + " 23:59:59")

    query += " ORDER BY timestamp ASC"

    cursor.execute(query, params)
    records = cursor.fetchall()
    cursor.close()
    conn.close()

    log_admin_action(logged_in_admin, f"Exported feedback report in format: {format_type}")

    file_stream = io.BytesIO()

    # --- WORD EXPORT ---
    if format_type == 'word':
        from docx import Document
        doc = Document()
        doc.add_heading('Ethiopian Federal Police - Feedback Report', 0)

        for r in records:
            doc.add_paragraph(f"Service: {r['service_name']} | Rating: {r['rating']} | Date: {r['timestamp']}")
            if r['comment']:
                doc.add_paragraph(f"Comment: {r['comment']}")

        doc.save(file_stream)
        file_stream.seek(0)
        return send_file(
            file_stream,
            as_attachment=True,
            download_name='police_feedback_report.docx',
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    # --- EXCEL EXPORT ---
    elif format_type == 'excel':
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        if ws is None:
            ws = wb.create_sheet()
        ws.title = "Feedback Reports"

        ws.append(["ID", "Service", "Sub-Service", "Rating", "Comment", "Date"])
        for r in records:
            ws.append([r['id'], r['service_name'], r['sub_service'], r['rating'], r['comment'], str(r['timestamp'])])

        wb.save(file_stream)
        file_stream.seek(0)
        return send_file(
            file_stream,
            as_attachment=True,
            download_name='police_feedback_report.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    # --- PDF EXPORT ---
    elif format_type == 'pdf':
        from reportlab.lib.pagesizes import letter
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

        doc = SimpleDocTemplate(file_stream, pagesize=letter)
        styles = getSampleStyleSheet()
        story = []

        title_style = ParagraphStyle(
            'TitleStyle',
            parent=styles['Heading1'],
            fontSize=16,
            spaceAfter=15
        )
        story.append(Paragraph("Ethiopian Federal Police - Feedback Report", title_style))
        story.append(Spacer(1, 10))

        for r in records:
            text = f"<b>Service:</b> {r['service_name']} | <b>Rating:</b> {r['rating']} | <b>Date:</b> {r['timestamp']}"
            story.append(Paragraph(text, styles['Normal']))
            if r['comment']:
                story.append(Paragraph(f"<b>Comment:</b> {r['comment']}", styles['Normal']))
            story.append(Spacer(1, 6))

        doc.build(story)
        file_stream.seek(0)
        return send_file(
            file_stream,
            as_attachment=True,
            download_name='police_feedback_report.pdf',
            mimetype='application/pdf'
        )

    return "Invalid format type specified. Use 'word', 'excel', or 'pdf'.", 400


def gregorian_to_ethiopian(dt):
    """Return a simple Ethiopian-calendar display string.

    This is intended for display only. It does not perform an exact
    Gregorian-to-Ethiopian calendar conversion for every date.
    """
    g_year = dt.year
    g_month = dt.month
    g_day = dt.day

    if g_month < 9 or (g_month == 9 and g_day < 11):
        et_year = g_year - 8
    else:
        et_year = g_year - 7

    et_months = [
        "መስከረም (Meskerem)",
        "ጥቅምት (Tikimt)",
        "ኅዳር (Hidar)",
        "ታኅሣሥ (Tahsas)",
        "ጥር (Tir)",
        "የካቲት (Yekatit)",
        "መጋቢት (Megabit)",
        "ሚያዝያ (Miyazia)",
        "ግንቦት (Ginbot)",
        "ሰኔ (Sene)",
        "ሐምሌ (Hamle)",
        "ነሐሴ (Nehase)",
        "ጳጉሜ (Pagume)"
    ]

    # This preserves the original project's display-oriented mapping.
    month_index = (g_month - 9) % 13
    return (
        f"{dt.day} {et_months[month_index]} {et_year} — "
        f"{dt.strftime('%H:%M:%S')}"
    )


@app.route('/admin/audit-logs')
def admin_audit_logs():
    """Display audit logs with role-based visibility."""
    logged_in_admin = session.get('admin_user')
    admin_credentials = get_current_admin_credentials()

    if not logged_in_admin or logged_in_admin not in admin_credentials:
        return redirect(url_for('admin_login'))

    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        if logged_in_admin == 'admin gen':
            cursor.execute(
                "SELECT id, admin_user, action_description, timestamp "
                "FROM audit_logs ORDER BY timestamp DESC"
            )
        else:
            cursor.execute(
                "SELECT id, admin_user, action_description, timestamp "
                "FROM audit_logs WHERE admin_user = ? ORDER BY timestamp DESC",
                (logged_in_admin,)
            )

        rows = cursor.fetchall()
        formatted_logs = []

        for row in rows:
            timestamp = row['timestamp']

            try:
                ethiopian_timestamp = (
                    gregorian_to_ethiopian(datetime.fromisoformat(timestamp))
                    if timestamp
                    else "N/A"
                )
            except (TypeError, ValueError):
                ethiopian_timestamp = timestamp or "N/A"

            formatted_logs.append({
                'id': row['id'],
                'admin_user': row['admin_user'],
                'action': row['action_description'],
                'action_description': row['action_description'],
                'timestamp': timestamp,
                'ethiopian_timestamp': ethiopian_timestamp
            })
    finally:
        cursor.close()
        conn.close()

    return render_template(
        'admin_audit_logs.html',
        audit_logs=formatted_logs,
        logs=formatted_logs,
        current_admin=logged_in_admin
    )


# ----------------------------------------------------
# DEPLOYMENT HEALTH CHECK
# ----------------------------------------------------
@app.route('/health')
def health_check():
    return jsonify({
        'status': 'ok',
        'service': 'Ethiopian Federal Police Citizen Feedback System'
    }), 200


# ----------------------------------------------------
# APPLICATION ENTRY POINT
# ----------------------------------------------------
if __name__ == '__main__':
    # Local: PORT defaults to 5000.
    # Render: PORT is supplied by the platform.
    port = int(os.environ.get('PORT', '5000'))
    debug_mode = os.environ.get('FLASK_DEBUG', 'false').lower() == 'true'
    app.run(host='0.0.0.0', port=port, debug=debug_mode)