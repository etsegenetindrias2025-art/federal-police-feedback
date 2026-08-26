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

from flask import (
    Flask,
    render_template,
    request,
    jsonify,
    session,
    redirect,
    url_for,
    send_file
)

import sqlite3
import os
import io
import re
import html
from datetime import datetime, timedelta

import qrcode
from gtts import gTTS


# ============================================================
# OPTIONAL PROFANITY LIBRARY
# ============================================================

try:
    from better_profanity import profanity  # type: ignore

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

            normalized = re.sub(
                r"[^\w\s]",
                "",
                str(text).lower()
            )

            for word in self._censored_words:

                if re.search(
                    rf"\b{re.escape(word)}\b",
                    normalized
                ):
                    return True

            return False

    profanity = _FallbackProfanity()


# ============================================================
# FLASK APPLICATION
# ============================================================

app = Flask(__name__)


# ============================================================
# SECURITY / DEPLOYMENT CONFIGURATION
# ============================================================

app.secret_key = os.environ.get(
    "FLASK_SECRET_KEY",
    "dev-only-federal-police-secret-key-change-me"
)

PUBLIC_BASE_URL = os.environ.get(
    "PUBLIC_BASE_URL",
    "https://ethiopian-federal-police-feedback.onrender.com"
).rstrip("/")

app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_SECURE"] = PUBLIC_BASE_URL.startswith(
    "https://"
)
app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(
    minutes=30
)


# ============================================================
# PROFANITY FILTER
# ============================================================

profanity.load_censor_words()

ETHIOPIC_BAD_WORDS = [
    "ውሻ",
    "ሌባ",
    "አህያ",
    "ጅብ",
    "ፋንድያ",
    "ሉጢ",
    "ክፉ",
    "በረንዳ አዳሪ",

    "wusha",
    "wesha",
    "leba",
    "ahiya",
    "jib",
    "fandya",
    "luti",
    "dingay",
    "balege",
    "dedeb",
    "denez"
]

profanity.add_censor_words(
    ETHIOPIC_BAD_WORDS
)


def is_inappropriate(text):

    if not text:
        return False

    if profanity.contains_profanity(text):
        return True

    cleaned_text = re.sub(
        r"[^\w\s]",
        "",
        str(text).lower()
    )

    for bad_word in ETHIOPIC_BAD_WORDS:

        if bad_word in cleaned_text:
            return True

    return False


# ============================================================
# DATABASE CONFIGURATION
# ============================================================

render_data_dir = "/var/data"

if (
    os.path.isdir(render_data_dir)
    and os.access(render_data_dir, os.W_OK)
):
    DB_PATH = os.path.join(
        render_data_dir,
        "database.db"
    )
else:
    DB_PATH = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "database.db"
    )


def get_db_connection():

    conn = sqlite3.connect(
        DB_PATH,
        timeout=30
    )

    conn.row_factory = sqlite3.Row

    return conn


# ============================================================
# DATABASE INITIALIZATION
# ============================================================

def init_db():

    conn = get_db_connection()
    cursor = conn.cursor()

    # --------------------------------------------------------
    # FEEDBACKS
    # --------------------------------------------------------

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS feedbacks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            service_name TEXT NOT NULL,
            sub_service TEXT,
            rating TEXT NOT NULL,
            comment TEXT,
            is_read BOOLEAN DEFAULT 0,
            timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
        )
        """
    )

    # --------------------------------------------------------
    # FINGERPRINT RECORDS
    # --------------------------------------------------------

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS fingerprint_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id TEXT,
            fingerprint_data TEXT NOT NULL,
            status TEXT DEFAULT 'Verified',
            timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
        )
        """
    )

    # --------------------------------------------------------
    # SERVICES
    # --------------------------------------------------------

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS services (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            service_key TEXT UNIQUE NOT NULL,
            service_name TEXT NOT NULL
        )
        """
    )

    # --------------------------------------------------------
    # SUB SERVICES
    # --------------------------------------------------------

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS sub_services (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            service_key TEXT NOT NULL,
            sub_service_key TEXT NOT NULL,
            sub_service_name TEXT NOT NULL,
            UNIQUE(service_key, sub_service_key)
        )
        """
    )

    # --------------------------------------------------------
    # AUDIT LOGS
    # --------------------------------------------------------

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS audit_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            admin_user TEXT NOT NULL,
            action_description TEXT NOT NULL,
            timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
        )
        """
    )

    # --------------------------------------------------------
    # DEFAULT SERVICES
    # --------------------------------------------------------

    default_services = [
        (
            "police_clearance",
            "Police Clearance"
        ),
        (
            "complaint",
            "Complaint"
        ),
        (
            "hospital",
            "Hospital"
        ),
        (
            "logistics",
            "Logistics"
        ),
        (
            "education_training",
            "Education & Training"
        ),
        (
            "other",
            "Other"
        )
    ]

    for key, name in default_services:

        cursor.execute(
            """
            INSERT OR IGNORE INTO services
            (service_key, service_name)
            VALUES (?, ?)
            """,
            (key, name)
        )

    # --------------------------------------------------------
    # DEFAULT SUB SERVICES
    # --------------------------------------------------------

    default_sub_services = {

        "police_clearance": [
            (
                "new_clearance",
                "New Police Clearance"
            ),
            (
                "renewal",
                "Renewal"
            ),
            (
                "criminal_record",
                "Criminal Record Verification"
            ),
            (
                "fingerprint",
                "Fingerprint Registration"
            ),
            (
                "document_collection",
                "Document Collection"
            )
        ],

        "complaint": [
            (
                "crime_complaint",
                "Crime Complaint Registration"
            ),
            (
                "public_office",
                "Public Complaint Office"
            ),
            (
                "online_followup",
                "Online Complaint Follow-up"
            ),
            (
                "investigation",
                "Investigation"
            ),
            (
                "resolution",
                "Resolution"
            )
        ],

        "hospital": [
            (
                "opd",
                "OPD"
            ),
            (
                "emergency",
                "Emergency"
            ),
            (
                "pharmacy",
                "Pharmacy"
            ),
            (
                "laboratory",
                "Laboratory"
            ),
            (
                "medical_exam",
                "Medical Examination"
            )
        ],

        "logistics": [
            (
                "vehicle_mgmt",
                "Vehicle Management"
            ),
            (
                "garage",
                "Garage"
            ),
            (
                "equipment_dist",
                "Equipment Distribution"
            ),
            (
                "inventory",
                "Inventory"
            ),
            (
                "procurement",
                "Procurement"
            )
        ],

        "education_training": [
            (
                "student_reg",
                "Student Registration"
            ),
            (
                "training",
                "Training"
            ),
            (
                "certificates",
                "Certificates"
            ),
            (
                "examination",
                "Examination"
            ),
            (
                "academic_records",
                "Academic Records"
            )
        ],

        "other": [
            (
                "reception",
                "Reception"
            ),
            (
                "ict_support",
                "ICT Support"
            ),
            (
                "hr",
                "HR"
            ),
            (
                "finance",
                "Finance"
            ),
            (
                "admin",
                "Administration"
            )
        ]
    }

    for service_key, sub_list in default_sub_services.items():

        for sub_key, sub_name in sub_list:

            cursor.execute(
                """
                INSERT OR IGNORE INTO sub_services
                (
                    service_key,
                    sub_service_key,
                    sub_service_name
                )
                VALUES (?, ?, ?)
                """,
                (
                    service_key,
                    sub_key,
                    sub_name
                )
            )

    # --------------------------------------------------------
    # DATABASE MIGRATION
    # --------------------------------------------------------

    cursor.execute(
        "PRAGMA table_info(feedbacks)"
    )

    columns = [
        column[1]
        for column in cursor.fetchall()
    ]

    if "is_read" not in columns:

        cursor.execute(
            """
            ALTER TABLE feedbacks
            ADD COLUMN is_read BOOLEAN DEFAULT 0
            """
        )

    conn.commit()

    cursor.close()
    conn.close()


init_db()

print(
    f"[startup] Using database file at: {DB_PATH}"
)


# ============================================================
# ADMIN AUDIT LOG
# ============================================================

def log_admin_action(
    username,
    description
):

    conn = None

    try:

        conn = get_db_connection()
        cursor = conn.cursor()

        cursor.execute(
            """
            INSERT INTO audit_logs
            (
                admin_user,
                action_description
            )
            VALUES (?, ?)
            """,
            (
                str(username),
                str(description)
            )
        )

        conn.commit()

        cursor.close()

    except Exception:

        import traceback

        print(
            "=== Error logging audit trail ==="
        )

        traceback.print_exc()

    finally:

        if conn is not None:
            conn.close()


# ============================================================
# SERVICE HELPERS
# ============================================================

def get_service_map():

    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute(
        """
        SELECT service_key, service_name
        FROM services
        ORDER BY id
        """
    )

    rows = cursor.fetchall()

    cursor.close()
    conn.close()

    return {
        row["service_key"]:
        row["service_name"]
        for row in rows
    }


def get_sub_service_map():

    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute(
        """
        SELECT
            service_key,
            sub_service_key,
            sub_service_name
        FROM sub_services
        ORDER BY id
        """
    )

    rows = cursor.fetchall()

    cursor.close()
    conn.close()

    sub_map = {}

    for row in rows:

        service_key = row["service_key"]

        if service_key not in sub_map:
            sub_map[service_key] = {}

        sub_map[service_key][
            row["sub_service_key"]
        ] = row["sub_service_name"]

    return sub_map


# ============================================================
# ADMIN CREDENTIALS
# ============================================================

def get_admin_credentials():

    return {

        "admin gen": {
            "password": "1234",
            "type": "general",
            "service": "all",
            "sub_service": "all",
            "title": "General Admin Dashboard"
        },

        "admin pol": {
            "password": "1234",
            "type": "service",
            "service": "police_clearance",
            "sub_service": "all",
            "title": "Police Clearance Admin"
        },

        "admin com": {
            "password": "1234",
            "type": "service",
            "service": "complaint",
            "sub_service": "all",
            "title": "Complaint Admin"
        },

        "admin hos": {
            "password": "1234",
            "type": "service",
            "service": "hospital",
            "sub_service": "all",
            "title": "Hospital Admin"
        },

        "admin log": {
            "password": "1234",
            "type": "service",
            "service": "logistics",
            "sub_service": "all",
            "title": "Logistics Admin"
        },

        "admin edu": {
            "password": "1234",
            "type": "service",
            "service": "education_training",
            "sub_service": "all",
            "title": "Education Admin"
        },

        "admin oth": {
            "password": "1234",
            "type": "service",
            "service": "other",
            "sub_service": "all",
            "title": "Other Admin"
        },

        "admin new": {
            "password": "1234",
            "type": "sub_service",
            "service": "police_clearance",
            "sub_service": "new_clearance",
            "title": "Sub Admin: New Clearance"
        },

        "admin ren": {
            "password": "1234",
            "type": "sub_service",
            "service": "police_clearance",
            "sub_service": "renewal",
            "title": "Sub Admin: Renewal"
        },

        "admin cri": {
            "password": "1234",
            "type": "sub_service",
            "service": "police_clearance",
            "sub_service": "criminal_record",
            "title": "Sub Admin: Criminal Record"
        },

        "admin fin": {
            "password": "1234",
            "type": "sub_service",
            "service": "police_clearance",
            "sub_service": "fingerprint",
            "title": "Sub Admin: Fingerprint"
        },

        "admin doc": {
            "password": "1234",
            "type": "sub_service",
            "service": "police_clearance",
            "sub_service": "document_collection",
            "title": "Sub Admin: Document Collection"
        },

        "admin pub": {
            "password": "1234",
            "type": "sub_service",
            "service": "complaint",
            "sub_service": "public_office",
            "title": "Sub Admin: Public Complaint"
        },

        "admin onl": {
            "password": "1234",
            "type": "sub_service",
            "service": "complaint",
            "sub_service": "online_followup",
            "title": "Sub Admin: Online Follow-up"
        },

        "admin inv": {
            "password": "1234",
            "type": "sub_service",
            "service": "complaint",
            "sub_service": "investigation",
            "title": "Sub Admin: Investigation"
        },

        "admin res": {
            "password": "1234",
            "type": "sub_service",
            "service": "complaint",
            "sub_service": "resolution",
            "title": "Sub Admin: Resolution"
        },

        "admin opd": {
            "password": "1234",
            "type": "sub_service",
            "service": "hospital",
            "sub_service": "opd",
            "title": "Sub Admin: OPD"
        },

        "admin eme": {
            "password": "1234",
            "type": "sub_service",
            "service": "hospital",
            "sub_service": "emergency",
            "title": "Sub Admin: Emergency"
        },

        "admin pha": {
            "password": "1234",
            "type": "sub_service",
            "service": "hospital",
            "sub_service": "pharmacy",
            "title": "Sub Admin: Pharmacy"
        },

        "admin lab": {
            "password": "1234",
            "type": "sub_service",
            "service": "hospital",
            "sub_service": "laboratory",
            "title": "Sub Admin: Laboratory"
        },

        "admin med": {
            "password": "1234",
            "type": "sub_service",
            "service": "hospital",
            "sub_service": "medical_exam",
            "title": "Sub Admin: Medical Examination"
        },

        "admin veh": {
            "password": "1234",
            "type": "sub_service",
            "service": "logistics",
            "sub_service": "vehicle_mgmt",
            "title": "Sub Admin: Vehicle Management"
        },

        "admin gar": {
            "password": "1234",
            "type": "sub_service",
            "service": "logistics",
            "sub_service": "garage",
            "title": "Sub Admin: Garage"
        },

        "admin equ": {
            "password": "1234",
            "type": "sub_service",
            "service": "logistics",
            "sub_service": "equipment_dist",
            "title": "Sub Admin: Equipment Distribution"
        },

        "admin pro": {
            "password": "1234",
            "type": "sub_service",
            "service": "logistics",
            "sub_service": "procurement",
            "title": "Sub Admin: Procurement"
        },

        "admin stu": {
            "password": "1234",
            "type": "sub_service",
            "service": "education_training",
            "sub_service": "student_reg",
            "title": "Sub Admin: Student Registration"
        },

        "admin tra": {
            "password": "1234",
            "type": "sub_service",
            "service": "education_training",
            "sub_service": "training",
            "title": "Sub Admin: Training"
        },

        "admin cer": {
            "password": "1234",
            "type": "sub_service",
            "service": "education_training",
            "sub_service": "certificates",
            "title": "Sub Admin: Certificates"
        },

        "admin exa": {
            "password": "1234",
            "type": "sub_service",
            "service": "education_training",
            "sub_service": "examination",
            "title": "Sub Admin: Examination"
        },

        "admin aca": {
            "password": "1234",
            "type": "sub_service",
            "service": "education_training",
            "sub_service": "academic_records",
            "title": "Sub Admin: Academic Records"
        },

        "admin rec": {
            "password": "1234",
            "type": "sub_service",
            "service": "other",
            "sub_service": "reception",
            "title": "Sub Admin: Reception"
        },

        "admin ict": {
            "password": "1234",
            "type": "sub_service",
            "service": "other",
            "sub_service": "ict_support",
            "title": "Sub Admin: ICT Support"
        },

        "admin hr": {
            "password": "1234",
            "type": "sub_service",
            "service": "other",
            "sub_service": "hr",
            "title": "Sub Admin: HR"
        },

        "admin fnn": {
            "password": "1234",
            "type": "sub_service",
            "service": "other",
            "sub_service": "finance",
            "title": "Sub Admin: Finance"
        },

        "admin adm": {
            "password": "1234",
            "type": "sub_service",
            "service": "other",
            "sub_service": "admin",
            "title": "Sub Admin: Administration"
        }
    }


ADMIN_PASSWORD_OVERRIDES = {}


def get_current_admin_credentials():

    credentials = get_admin_credentials()

    for username, password in ADMIN_PASSWORD_OVERRIDES.items():

        if username in credentials:
            credentials[username]["password"] = password

    return credentials


# ============================================================
# DATE HELPERS
# ============================================================

def parse_date_start(value):

    if not value:
        return None

    try:

        return datetime.strptime(
            value,
            "%Y-%m-%d"
        )

    except ValueError:

        return None


def parse_date_end(value):

    if not value:
        return None

    try:

        return (
            datetime.strptime(
                value,
                "%Y-%m-%d"
            )
            + timedelta(days=1)
        )

    except ValueError:

        return None


def parse_database_timestamp(value):

    if value is None:
        return None

    value = str(value).strip()

    if not value:
        return None

    # SQLite commonly stores:
    # 2026-08-26 10:30:00
    # or
    # 2026-08-26T10:30:00

    value = value.replace(
        "Z",
        ""
    )

    try:

        return datetime.fromisoformat(
            value
        )

    except ValueError:

        pass

    formats = [
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d"
    ]

    for fmt in formats:

        try:

            return datetime.strptime(
                value,
                fmt
            )

        except ValueError:
            continue

    return None


def normalize_value(value):

    if value is None:
        return ""

    return str(value).strip().lower()


# ============================================================
# GET FILTER VALUES
# ============================================================

def get_dashboard_filters():

    """
    Supports both the new parameter names and the older
    parameter names so existing admin_dashboard.html files
    do not immediately break.
    """

    service_filter = (
        request.args.get("service_filter")
        or request.args.get("service")
        or "all"
    )

    rating_filter = (
        request.args.get("rating_filter")
        or request.args.get("rating")
        or ""
    )

    date_from = (
        request.args.get("date_from")
        or ""
    )

    date_to = (
        request.args.get("date_to")
        or ""
    )

    return {
        "service_filter": normalize_value(
            service_filter
        ) if service_filter != "all" else "all",

        "rating_filter": str(
            rating_filter
        ).strip(),

        "date_from": str(
            date_from
        ).strip(),

        "date_to": str(
            date_to
        ).strip()
    }


# ============================================================
# FILTER A FEEDBACK RECORD
# ============================================================

def feedback_passes_filters(
    feedback,
    service_filter="all",
    rating_filter="",
    date_from="",
    date_to=""
):

    # --------------------------------------------------------
    # SERVICE FILTER
    # --------------------------------------------------------

    db_service = normalize_value(
        feedback["service_name"]
    )

    if (
        service_filter
        and service_filter != "all"
        and db_service != normalize_value(service_filter)
    ):
        return False

    # --------------------------------------------------------
    # RATING FILTER
    # --------------------------------------------------------

    if rating_filter:

        db_rating = str(
            feedback["rating"]
        ).strip()

        if db_rating != str(
            rating_filter
        ).strip():

            return False

    # --------------------------------------------------------
    # DATE FROM
    # --------------------------------------------------------

    dt_from = parse_date_start(
        date_from
    )

    if dt_from:

        feedback_dt = parse_database_timestamp(
            feedback["timestamp"]
        )

        if feedback_dt is None:
            return False

        if feedback_dt < dt_from:
            return False

    # --------------------------------------------------------
    # DATE TO
    #
    # Use exclusive next-day boundary.
    #
    # Example:
    # date_to = 2026-08-26
    #
    # Includes:
    # 2026-08-26 00:00:00
    # 2026-08-26 10:30:00
    # 2026-08-26 23:59:59
    # --------------------------------------------------------

    dt_to_exclusive = parse_date_end(
        date_to
    )

    if dt_to_exclusive:

        feedback_dt = parse_database_timestamp(
            feedback["timestamp"]
        )

        if feedback_dt is None:
            return False

        if feedback_dt >= dt_to_exclusive:
            return False

    return True


# ============================================================
# ROLE VISIBILITY
# ============================================================

def feedback_visible_to_admin(
    feedback,
    admin_type,
    assigned_service,
    assigned_sub
):

    db_service = normalize_value(
        feedback["service_name"]
    )

    db_sub = normalize_value(
        feedback["sub_service"]
    )

    if admin_type == "general":
        return True

    if admin_type == "service":

        return (
            db_service
            == normalize_value(assigned_service)
        )

    if admin_type == "sub_service":

        return (
            db_service
            == normalize_value(assigned_service)
            and
            db_sub
            == normalize_value(assigned_sub)
        )

    return False


# ============================================================
# FILTER FEEDBACK LIST
# ============================================================

def get_filtered_feedbacks(
    admin_type,
    assigned_service,
    assigned_sub,
    service_filter="all",
    rating_filter="",
    date_from="",
    date_to=""
):

    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute(
        """
        SELECT *
        FROM feedbacks
        ORDER BY timestamp DESC, id DESC
        """
    )

    all_feedbacks = cursor.fetchall()

    cursor.close()
    conn.close()

    filtered = []

    for feedback in all_feedbacks:

        # First enforce admin role permissions.
        if not feedback_visible_to_admin(
            feedback,
            admin_type,
            assigned_service,
            assigned_sub
        ):
            continue

        # Then enforce user-selected filters.
        if not feedback_passes_filters(
            feedback,
            service_filter,
            rating_filter,
            date_from,
            date_to
        ):
            continue

        filtered.append(
            feedback
        )

    return filtered


# ============================================================
# PROGRESSIVE WEB APP
# ============================================================

@app.route("/sw.js")
def service_worker():

    return app.send_static_file(
        "sw.js"
    )


# ============================================================
# PUBLIC ROUTES
# ============================================================

@app.route("/")
def fingerprint():

    return render_template(
        "fingerprint.html"
    )


@app.route("/welcome")
@app.route("/language")
def welcome_page():

    lang = request.args.get(
        "lang",
        "am"
    )

    return render_template(
        "welcome.html",
        lang=lang
    )


@app.route("/services")
def services():

    lang = request.args.get(
        "lang",
        "am"
    )

    service_map = get_service_map()

    return render_template(
        "services.html",
        lang=lang,
        service_map=service_map
    )


@app.route("/feedback")
def feedback():

    lang = request.args.get(
        "lang",
        "am"
    )

    service = request.args.get(
        "service",
        "police_clearance"
    )

    service_map = get_service_map()
    sub_service_map = get_sub_service_map()

    return render_template(
        "feedback.html",
        lang=lang,
        service=service,
        service_map=service_map,
        sub_service_map=sub_service_map
    )


# ============================================================
# FINGERPRINT API
# ============================================================

@app.route(
    "/api/fingerprint/scan",
    methods=["POST"]
)
def scan_fingerprint():

    try:

        data = request.get_json() or {}

        user_id = data.get(
            "user_id",
            "TEMPORARY_USER"
        )

        fingerprint_data = data.get(
            "fingerprint_data",
            "BYPASS_FINGERPRINT_HASH"
        )

        conn = get_db_connection()
        cursor = conn.cursor()

        cursor.execute(
            """
            INSERT INTO fingerprint_records
            (
                user_id,
                fingerprint_data
            )
            VALUES (?, ?)
            """,
            (
                str(user_id),
                str(fingerprint_data)
            )
        )

        conn.commit()

        cursor.close()
        conn.close()

        return jsonify(
            {
                "status": "success",
                "message": (
                    "የጣት አሻራ ስካን ሳይጠበቅ "
                    "ቀጥታ አልፏል (Demo Mode)!"
                ),
                "user_id": user_id
            }
        ), 200

    except Exception as e:

        return jsonify(
            {
                "status": "error",
                "message": str(e)
            }
        ), 500


# ============================================================
# FEEDBACK SUBMISSION
# ============================================================

@app.route(
    "/submit-feedback",
    methods=["POST"]
)
@app.route(
    "/api/submit-feedback",
    methods=["POST"]
)
def submit_feedback():

    try:

        feedback_count = session.get(
            "feedback_count",
            0
        )

        if feedback_count >= 3:

            return jsonify(
                {
                    "status": "error",
                    "message": (
                        "ለአሁኑ የተፈቀደልዎትን "
                        "3 አስተያየቶች ጨርሰዋል! / "
                        "You have reached your max "
                        "limit of 3 feedbacks for this session."
                    )
                }
            ), 403

        data = request.get_json(
            silent=True
        )

        if not data:

            return jsonify(
                {
                    "status": "error",
                    "message": "No data received"
                }
            ), 400

        rating = data.get(
            "rating",
            "😊"
        )

        comment = data.get(
            "comment",
            "No comment provided."
        )

        sub_service = data.get(
            "sub_service",
            "general_service"
        )

        service = (
            data.get("service")
            or data.get("category")
            or data.get("service_name")
            or "police_clearance"
        )

        client_timestamp = data.get(
            "timestamp"
        )

        if is_inappropriate(
            comment
        ):

            return jsonify(
                {
                    "status": "error",
                    "message": (
                        "Inappropriate language detected. "
                        "Please keep your feedback respectful."
                    )
                }
            ), 400

        conn = get_db_connection()
        cursor = conn.cursor()

        # ----------------------------------------------------
        # IMPORTANT:
        # Do NOT blindly accept arbitrary client timestamps.
        # If the frontend sends a valid ISO timestamp, use it.
        # Otherwise SQLite generates the timestamp.
        # ----------------------------------------------------

        timestamp_value = None

        if client_timestamp:

            parsed_client_timestamp = parse_database_timestamp(
                client_timestamp
            )

            if parsed_client_timestamp:

                timestamp_value = (
                    parsed_client_timestamp.strftime(
                        "%Y-%m-%d %H:%M:%S"
                    )
                )

        if timestamp_value:

            cursor.execute(
                """
                INSERT INTO feedbacks
                (
                    service_name,
                    sub_service,
                    rating,
                    comment,
                    is_read,
                    timestamp
                )
                VALUES (?, ?, ?, ?, 0, ?)
                """,
                (
                    str(service),
                    str(sub_service),
                    str(rating),
                    str(comment),
                    timestamp_value
                )
            )

        else:

            cursor.execute(
                """
                INSERT INTO feedbacks
                (
                    service_name,
                    sub_service,
                    rating,
                    comment,
                    is_read
                )
                VALUES (?, ?, ?, ?, 0)
                """,
                (
                    str(service),
                    str(sub_service),
                    str(rating),
                    str(comment)
                )
            )

        conn.commit()

        cursor.close()
        conn.close()

        session["feedback_count"] = (
            feedback_count + 1
        )

        return jsonify(
            {
                "status": "success",
                "message": (
                    "Feedback saved successfully! "
                    f"({session['feedback_count']}/3 submitted)"
                )
            }
        )

    except Exception as e:

        import traceback

        traceback.print_exc()

        return jsonify(
            {
                "status": "error",
                "message": str(e)
            }
        ), 500


# ============================================================
# UNREAD COUNT
# ============================================================

@app.route("/api/unread-count")
def api_unread_count():

    conn = get_db_connection()
    cursor = conn.cursor()

    try:

        cursor.execute(
            """
            SELECT COUNT(*) AS count
            FROM feedbacks
            WHERE is_read = 0
            """
        )

        result = cursor.fetchone()

        unread_count = (
            result["count"]
            if result
            else 0
        )

    except Exception:

        unread_count = 0

    finally:

        cursor.close()
        conn.close()

    return jsonify(
        {
            "unread_count": unread_count
        }
    )


# ============================================================
# ADMIN LOGIN
# ============================================================

@app.route(
    "/admin/login",
    methods=["GET", "POST"]
)
def admin_login():

    error = None

    admin_credentials = (
        get_current_admin_credentials()
    )

    if request.method == "POST":

        username = request.form.get(
            "username",
            ""
        ).strip()

        password = request.form.get(
            "password",
            ""
        )

        if (
            username in admin_credentials
            and
            admin_credentials[username]["password"]
            == password
        ):

            session["admin_user"] = username
            session.permanent = True

            log_admin_action(
                username,
                "Logged into the admin panel."
            )

            return redirect(
                url_for("admin_dashboard")
            )

        error = (
            "Invalid Username or Password. "
            "Please try again."
        )

    return render_template(
        "admin_login.html",
        error=error
    )


# ============================================================
# ADMIN LOGOUT
# ============================================================

@app.route("/admin/logout")
def admin_logout():

    logged_in_admin = session.get(
        "admin_user"
    )

    if logged_in_admin:

        log_admin_action(
            logged_in_admin,
            "Logged out of the admin panel."
        )

    session.pop(
        "admin_user",
        None
    )

    return redirect(
        url_for("admin_login")
    )


# ============================================================
# ADMIN DASHBOARD
# ============================================================

@app.route("/admin/dashboard")
def admin_dashboard():

    # --------------------------------------------------------
    # AUTHENTICATION
    # --------------------------------------------------------

    logged_in_admin = session.get(
        "admin_user"
    )

    admin_credentials = (
        get_current_admin_credentials()
    )

    if (
        not logged_in_admin
        or
        logged_in_admin not in admin_credentials
    ):

        return redirect(
            url_for("admin_login")
        )

    # --------------------------------------------------------
    # ADMIN INFORMATION
    # --------------------------------------------------------

    admin_info = admin_credentials[
        logged_in_admin
    ]

    admin_type = admin_info["type"]

    assigned_service = admin_info[
        "service"
    ]

    assigned_sub = admin_info[
        "sub_service"
    ]

    admin_title = admin_info[
        "title"
    ]

    is_general_admin = (
        admin_type == "general"
    )

    # --------------------------------------------------------
    # FILTERS
    # --------------------------------------------------------

    filters = get_dashboard_filters()

    service_filter = filters[
        "service_filter"
    ]

    rating_filter = filters[
        "rating_filter"
    ]

    date_from = filters[
        "date_from"
    ]

    date_to = filters[
        "date_to"
    ]

    # --------------------------------------------------------
    # SERVICE DATA
    # --------------------------------------------------------

    service_map = get_service_map()
    sub_service_map = get_sub_service_map()

    # --------------------------------------------------------
    # GET FILTERED RECORDS
    # --------------------------------------------------------

    feedbacks = get_filtered_feedbacks(
        admin_type=admin_type,
        assigned_service=assigned_service,
        assigned_sub=assigned_sub,
        service_filter=service_filter,
        rating_filter=rating_filter,
        date_from=date_from,
        date_to=date_to
    )

    # --------------------------------------------------------
    # TOTAL FILTERED FEEDBACKS
    # --------------------------------------------------------

    total_feedbacks_count = len(
        feedbacks
    )

    # --------------------------------------------------------
    # SERVICE COUNTS
    #
    # IMPORTANT:
    # Counts now use FILTERED records.
    # --------------------------------------------------------

    service_counts = {
        key: 0
        for key in service_map.keys()
    }

    chart_data = {
        label: 0
        for label in service_map.values()
    }

    for feedback in feedbacks:

        service_key = normalize_value(
            feedback["service_name"]
        )

        if service_key in service_map:

            service_counts[
                service_key
            ] += 1

            service_label = service_map[
                service_key
            ]

            chart_data[
                service_label
            ] += 1

        else:

            if "other" in service_counts:

                service_counts[
                    "other"
                ] += 1

                chart_data[
                    service_map["other"]
                ] += 1

    # --------------------------------------------------------
    # UNREAD NOTIFICATIONS
    #
    # Notification count respects admin permissions.
    # It does not depend on dashboard filters.
    # --------------------------------------------------------

    conn = get_db_connection()
    cursor = conn.cursor()

    try:

        cursor.execute(
            """
            SELECT *
            FROM feedbacks
            WHERE is_read = 0
            ORDER BY timestamp DESC, id DESC
            """
        )

        unread_records = cursor.fetchall()

    finally:

        cursor.close()
        conn.close()

    unread_notifications_count = 0

    for feedback in unread_records:

        if feedback_visible_to_admin(
            feedback,
            admin_type,
            assigned_service,
            assigned_sub
        ):

            unread_notifications_count += 1

    # --------------------------------------------------------
    # AI INSIGHTS
    # --------------------------------------------------------

    total_count = len(
        feedbacks
    )

    if total_count > 0:

        positive_ratings = sum(
            1
            for feedback in feedbacks
            if str(
                feedback["rating"]
            ).strip()
            in {
                "4",
                "5",
                "😊",
                "😍"
            }
        )

        satisfaction_pct = (
            f"{int((positive_ratings / total_count) * 100)}%"
        )

        visible_service_counts = {}

        for feedback in feedbacks:

            service_name = str(
                feedback["service_name"]
                or "General"
            )

            visible_service_counts[
                service_name
            ] = (
                visible_service_counts.get(
                    service_name,
                    0
                ) + 1
            )

        top_service = max(
            visible_service_counts.items(),
            key=lambda item: item[1]
        )[0]

        negative_feedback = [
            feedback
            for feedback in feedbacks
            if str(
                feedback["rating"]
            ).strip()
            in {
                "1",
                "2",
                "😡",
                "😞",
                "😐"
            }
        ]

        if negative_feedback:

            main_complaint = str(
                negative_feedback[0]["comment"]
                or "Negative feedback received."
            )

        else:

            main_complaint = (
                "None reported"
            )

        ai_insights = {
            "satisfaction": satisfaction_pct,

            "top_service": top_service,

            "main_complaint": main_complaint,

            "recommendation": (
                f"Based on {total_count} visible "
                f"filtered feedback records, continue "
                f"monitoring service quality and "
                f"response times for {top_service}."
            )
        }

    else:

        ai_insights = {

            "satisfaction": "N/A",

            "top_service": "N/A",

            "main_complaint": "None",

            "recommendation": (
                "No feedback records match "
                "the selected filters."
            )
        }

    # --------------------------------------------------------
    # RENDER DASHBOARD
    # --------------------------------------------------------

    return render_template(
        "admin_dashboard.html",

        feedbacks=feedbacks,

        admin_title=admin_title,

        is_general=is_general_admin,

        service_map=service_map,

        sub_service_map=sub_service_map,

        chart_data=chart_data,

        service_counts=service_counts,

        total_feedbacks_count=total_feedbacks_count,

        ai_insights=ai_insights,

        unread_notifications_count=(
            unread_notifications_count
        ),

        # New names
        service_filter=service_filter,

        rating_filter=rating_filter,

        date_from=date_from,

        date_to=date_to,

        # Backward-compatible names
        current_filter=service_filter
    )


# ============================================================
# ADMIN NOTIFICATIONS
# ============================================================

@app.route("/admin/notifications")
def admin_notifications():

    logged_in_admin = session.get(
        "admin_user"
    )

    admin_credentials = (
        get_current_admin_credentials()
    )

    if (
        not logged_in_admin
        or
        logged_in_admin not in admin_credentials
    ):

        return redirect(
            url_for("admin_login")
        )

    admin_info = admin_credentials[
        logged_in_admin
    ]

    admin_type = admin_info["type"]

    assigned_service = admin_info[
        "service"
    ]

    assigned_sub = admin_info[
        "sub_service"
    ]

    conn = get_db_connection()
    cursor = conn.cursor()

    # --------------------------------------------------------
    # MARK VISIBLE NOTIFICATIONS AS READ
    # --------------------------------------------------------

    if admin_type == "general":

        cursor.execute(
            """
            UPDATE feedbacks
            SET is_read = 1
            WHERE is_read = 0
            """
        )

    elif admin_type == "service":

        cursor.execute(
            """
            UPDATE feedbacks
            SET is_read = 1
            WHERE is_read = 0
            AND service_name = ?
            """,
            (
                assigned_service,
            )
        )

    elif admin_type == "sub_service":

        cursor.execute(
            """
            UPDATE feedbacks
            SET is_read = 1
            WHERE is_read = 0
            AND service_name = ?
            AND sub_service = ?
            """,
            (
                assigned_service,
                assigned_sub
            )
        )

    conn.commit()

    # --------------------------------------------------------
    # LOAD NOTIFICATIONS
    # --------------------------------------------------------

    cursor.execute(
        """
        SELECT *
        FROM feedbacks
        ORDER BY timestamp DESC, id DESC
        """
    )

    all_feedbacks = cursor.fetchall()

    notifications = []

    for feedback in all_feedbacks:

        if feedback_visible_to_admin(
            feedback,
            admin_type,
            assigned_service,
            assigned_sub
        ):

            notifications.append(
                feedback
            )

    cursor.close()
    conn.close()

    service_map = get_service_map()
    sub_service_map = get_sub_service_map()

    return render_template(
        "admin_notifications.html",
        notifications=notifications,
        service_map=service_map,
        sub_service_map=sub_service_map
    )


# ============================================================
# ADMIN SETTINGS
# ============================================================

@app.route(
    "/admin/settings",
    methods=["GET", "POST"]
)
def admin_settings():

    logged_in_admin = session.get(
        "admin_user"
    )

    admin_credentials = (
        get_current_admin_credentials()
    )

    if (
        not logged_in_admin
        or
        logged_in_admin not in admin_credentials
    ):

        return redirect(
            url_for("admin_login")
        )

    admin_info = admin_credentials[
        logged_in_admin
    ]

    is_general_admin = (
        admin_info["type"] == "general"
    )

    message = None
    error = None

    if request.method == "POST":

        action = request.form.get(
            "action"
        )

        # ----------------------------------------------------
        # CHANGE PASSWORD
        # ----------------------------------------------------

        if action == "password":

            current_pass = request.form.get(
                "current_password",
                ""
            )

            new_pass = request.form.get(
                "new_password",
                ""
            )

            if (
                admin_credentials[
                    logged_in_admin
                ]["password"]
                == current_pass
            ):

                if (
                    not new_pass
                    or
                    len(new_pass) < 4
                ):

                    error = (
                        "New password must contain "
                        "at least 4 characters."
                    )

                else:

                    ADMIN_PASSWORD_OVERRIDES[
                        logged_in_admin
                    ] = new_pass

                    message = (
                        "Password updated successfully!"
                    )

                    log_admin_action(
                        logged_in_admin,
                        "Changed account password."
                    )

            else:

                error = (
                    "Current password is incorrect."
                )

        # ----------------------------------------------------
        # GENERAL ADMIN CRUD
        # ----------------------------------------------------

        elif is_general_admin:

            # -----------------------------------------------
            # ADD SERVICE
            # -----------------------------------------------

            if action == "add_service":

                service_key_raw = (
                    request.form.get(
                        "service_key"
                    )
                )

                service_key = (
                    service_key_raw.strip()
                    .lower()
                    .replace(" ", "_")
                    if service_key_raw
                    else None
                )

                service_name_raw = (
                    request.form.get(
                        "service_name"
                    )
                )

                service_name = (
                    service_name_raw.strip()
                    if service_name_raw
                    else None
                )

                if service_key and service_name:

                    try:

                        conn = get_db_connection()
                        cursor = conn.cursor()

                        cursor.execute(
                            """
                            INSERT INTO services
                            (
                                service_key,
                                service_name
                            )
                            VALUES (?, ?)
                            """,
                            (
                                service_key,
                                service_name
                            )
                        )

                        conn.commit()

                        cursor.close()
                        conn.close()

                        message = (
                            f"Service '{service_name}' "
                            "added successfully!"
                        )

                        log_admin_action(
                            logged_in_admin,
                            (
                                f"Added new service: "
                                f"{service_name}"
                            )
                        )

                    except Exception:

                        error = (
                            "Service key already exists "
                            "or invalid input."
                        )

                else:

                    error = (
                        "All fields are required "
                        "to add a service."
                    )

            # -----------------------------------------------
            # UPDATE SERVICE
            # -----------------------------------------------

            elif action == "update_service":

                service_key = request.form.get(
                    "service_key"
                )

                new_service_name_raw = (
                    request.form.get(
                        "new_service_name"
                    )
                )

                service_name = (
                    new_service_name_raw.strip()
                    if new_service_name_raw
                    else None
                )

                if service_key and service_name:

                    conn = get_db_connection()
                    cursor = conn.cursor()

                    cursor.execute(
                        """
                        UPDATE services
                        SET service_name = ?
                        WHERE service_key = ?
                        """,
                        (
                            service_name,
                            service_key
                        )
                    )

                    conn.commit()

                    cursor.close()
                    conn.close()

                    message = (
                        "Service updated successfully!"
                    )

                    log_admin_action(
                        logged_in_admin,
                        (
                            f"Updated service key "
                            f"'{service_key}' name "
                            f"to '{service_name}'"
                        )
                    )

                else:

                    error = (
                        "Invalid service update details."
                    )

            # -----------------------------------------------
            # DELETE SERVICE
            # -----------------------------------------------

            elif action == "delete_service":

                service_key = request.form.get(
                    "service_key"
                )

                if service_key:

                    conn = get_db_connection()
                    cursor = conn.cursor()

                    cursor.execute(
                        """
                        DELETE FROM services
                        WHERE service_key = ?
                        """,
                        (
                            service_key,
                        )
                    )

                    cursor.execute(
                        """
                        DELETE FROM sub_services
                        WHERE service_key = ?
                        """,
                        (
                            service_key,
                        )
                    )

                    conn.commit()

                    cursor.close()
                    conn.close()

                    message = (
                        "Service and its sub-services "
                        "deleted successfully!"
                    )

                    log_admin_action(
                        logged_in_admin,
                        (
                            "Deleted service and "
                            f"sub-services for key "
                            f"'{service_key}'"
                        )
                    )

                else:

                    error = (
                        "Select a service to delete."
                    )

            # -----------------------------------------------
            # ADD SUB-SERVICE
            # -----------------------------------------------

            elif action == "add_sub_service":

                parent_service = request.form.get(
                    "parent_service_key"
                )

                sub_key_raw = request.form.get(
                    "sub_service_key"
                )

                sub_key = (
                    sub_key_raw.strip()
                    .lower()
                    .replace(" ", "_")
                    if sub_key_raw
                    else ""
                )

                sub_name_raw = request.form.get(
                    "sub_service_name"
                )

                sub_name = (
                    sub_name_raw.strip()
                    if sub_name_raw
                    else ""
                )

                if (
                    parent_service
                    and sub_key
                    and sub_name
                ):

                    try:

                        conn = get_db_connection()
                        cursor = conn.cursor()

                        cursor.execute(
                            """
                            INSERT INTO sub_services
                            (
                                service_key,
                                sub_service_key,
                                sub_service_name
                            )
                            VALUES (?, ?, ?)
                            """,
                            (
                                parent_service,
                                sub_key,
                                sub_name
                            )
                        )

                        conn.commit()

                        cursor.close()
                        conn.close()

                        message = (
                            f"Sub-service '{sub_name}' "
                            "added successfully!"
                        )

                        log_admin_action(
                            logged_in_admin,
                            (
                                f"Added sub-service "
                                f"'{sub_name}' under "
                                f"'{parent_service}'"
                            )
                        )

                    except Exception:

                        error = (
                            "Sub-service key already "
                            "exists under this service."
                        )

                else:

                    error = (
                        "All fields are required "
                        "to add a sub-service."
                    )

            # -----------------------------------------------
            # UPDATE SUB-SERVICE
            # -----------------------------------------------

            elif action == "update_sub_service":

                parent_service = request.form.get(
                    "parent_service_key"
                )

                sub_key = request.form.get(
                    "sub_service_key"
                )

                sub_name_raw = request.form.get(
                    "new_sub_service_name"
                )

                sub_name = (
                    sub_name_raw.strip()
                    if sub_name_raw
                    else ""
                )

                if (
                    parent_service
                    and sub_key
                    and sub_name
                ):

                    conn = get_db_connection()
                    cursor = conn.cursor()

                    cursor.execute(
                        """
                        UPDATE sub_services
                        SET sub_service_name = ?
                        WHERE service_key = ?
                        AND sub_service_key = ?
                        """,
                        (
                            sub_name,
                            parent_service,
                            sub_key
                        )
                    )

                    conn.commit()

                    cursor.close()
                    conn.close()

                    message = (
                        "Sub-service updated successfully!"
                    )

                    log_admin_action(
                        logged_in_admin,
                        (
                            f"Updated sub-service "
                            f"'{sub_key}' under "
                            f"'{parent_service}' "
                            f"to '{sub_name}'"
                        )
                    )

                else:

                    error = (
                        "Invalid sub-service "
                        "update details."
                    )

            # -----------------------------------------------
            # DELETE SUB-SERVICE
            # -----------------------------------------------

            elif action == "delete_sub_service":

                parent_service = request.form.get(
                    "parent_service_key"
                )

                sub_key = request.form.get(
                    "sub_service_key"
                )

                if parent_service and sub_key:

                    conn = get_db_connection()
                    cursor = conn.cursor()

                    cursor.execute(
                        """
                        DELETE FROM sub_services
                        WHERE service_key = ?
                        AND sub_service_key = ?
                        """,
                        (
                            parent_service,
                            sub_key
                        )
                    )

                    conn.commit()

                    cursor.close()
                    conn.close()

                    message = (
                        "Sub-service deleted successfully!"
                    )

                    log_admin_action(
                        logged_in_admin,
                        (
                            f"Deleted sub-service "
                            f"'{sub_key}' under "
                            f"'{parent_service}'"
                        )
                    )

                else:

                    error = (
                        "Select a sub-service to delete."
                    )

        else:

            error = (
                "Unauthorized action: Only General "
                "Admin can perform CRUD operations "
                "on services."
            )

    service_map = get_service_map()
    sub_service_map = get_sub_service_map()

    return render_template(
        "admin_settings.html",
        message=message,
        error=error,
        service_map=service_map,
        sub_service_map=sub_service_map,
        is_general=is_general_admin
    )


# ============================================================
# EXPORT REPORT
# ============================================================

@app.route(
    "/admin/export/<format_type>"
)
def export_report(format_type):

    # --------------------------------------------------------
    # AUTHENTICATION
    # --------------------------------------------------------

    logged_in_admin = session.get(
        "admin_user"
    )

    admin_credentials = (
        get_current_admin_credentials()
    )

    if (
        not logged_in_admin
        or
        logged_in_admin not in admin_credentials
    ):

        return redirect(
            url_for("admin_login")
        )

    # --------------------------------------------------------
    # ADMIN ROLE
    # --------------------------------------------------------

    admin_info = admin_credentials[
        logged_in_admin
    ]

    admin_type = admin_info[
        "type"
    ]

    assigned_service = admin_info[
        "service"
    ]

    assigned_sub = admin_info[
        "sub_service"
    ]

    # --------------------------------------------------------
    # FILTERS
    #
    # Accept both:
    # service_filter
    # service
    #
    # rating_filter
    # rating
    # --------------------------------------------------------

    filters = get_dashboard_filters()

    service_filter = filters[
        "service_filter"
    ]

    rating_filter = filters[
        "rating_filter"
    ]

    date_from = filters[
        "date_from"
    ]

    date_to = filters[
        "date_to"
    ]

    # --------------------------------------------------------
    # GET EXACT SAME FILTERED DATA AS DASHBOARD
    # --------------------------------------------------------

    records = get_filtered_feedbacks(
        admin_type=admin_type,
        assigned_service=assigned_service,
        assigned_sub=assigned_sub,
        service_filter=service_filter,
        rating_filter=rating_filter,
        date_from=date_from,
        date_to=date_to
    )

    # --------------------------------------------------------
    # LOG EXPORT
    # --------------------------------------------------------

    log_admin_action(
        logged_in_admin,
        (
            f"Exported filtered feedback report "
            f"in format: {format_type}; "
            f"service={service_filter}; "
            f"rating={rating_filter}; "
            f"date_from={date_from}; "
            f"date_to={date_to}; "
            f"records={len(records)}"
        )
    )

    # --------------------------------------------------------
    # FILE STREAM
    # --------------------------------------------------------

    file_stream = io.BytesIO()

    # ========================================================
    # WORD
    # ========================================================

    if format_type == "word":

        from docx import Document

        doc = Document()

        doc.add_heading(
            "Ethiopian Federal Police - Feedback Report",
            0
        )

        # ----------------------------------------------------
        # REPORT FILTER INFORMATION
        # ----------------------------------------------------

        doc.add_paragraph(
            f"Service Filter: {service_filter}"
        )

        doc.add_paragraph(
            f"Rating Filter: "
            f"{rating_filter or 'All'}"
        )

        doc.add_paragraph(
            f"Date From: "
            f"{date_from or 'All'}"
        )

        doc.add_paragraph(
            f"Date To: "
            f"{date_to or 'All'}"
        )

        doc.add_paragraph(
            f"Total Records: {len(records)}"
        )

        doc.add_paragraph("")

        # ----------------------------------------------------
        # RECORDS
        # ----------------------------------------------------

        if not records:

            doc.add_paragraph(
                "No feedback records match "
                "the selected filters."
            )

        else:

            for record in records:

                doc.add_paragraph(
                    (
                        f"Service: "
                        f"{record['service_name']}"
                    )
                )

                doc.add_paragraph(
                    (
                        f"Sub-Service: "
                        f"{record['sub_service']}"
                    )
                )

                doc.add_paragraph(
                    (
                        f"Rating: "
                        f"{record['rating']}"
                    )
                )

                doc.add_paragraph(
                    (
                        f"Date: "
                        f"{record['timestamp']}"
                    )
                )

                if record["comment"]:

                    doc.add_paragraph(
                        (
                            f"Comment: "
                            f"{record['comment']}"
                        )
                    )

                doc.add_paragraph(
                    "-" * 60
                )

        doc.save(
            file_stream
        )

        file_stream.seek(0)

        return send_file(
            file_stream,
            as_attachment=True,
            download_name=(
                "police_feedback_report.docx"
            ),
            mimetype=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            )
        )

    # ========================================================
    # EXCEL
    # ========================================================

    elif format_type == "excel":

        import openpyxl

        from openpyxl.styles import Font

        wb = openpyxl.Workbook()

        ws = wb.active

        if ws is None:
            ws = wb.create_sheet()

        ws.title = "Feedback Reports"

        # ----------------------------------------------------
        # REPORT INFORMATION
        # ----------------------------------------------------

        ws.append(
            [
                "Ethiopian Federal Police "
                "Feedback Report"
            ]
        )

        ws["A1"].font = Font(
            bold=True
        )

        ws.append(
            [
                "Service Filter",
                service_filter
            ]
        )

        ws.append(
            [
                "Rating Filter",
                rating_filter or "All"
            ]
        )

        ws.append(
            [
                "Date From",
                date_from or "All"
            ]
        )

        ws.append(
            [
                "Date To",
                date_to or "All"
            ]
        )

        ws.append(
            [
                "Total Records",
                len(records)
            ]
        )

        ws.append([])

        # ----------------------------------------------------
        # TABLE HEADER
        # ----------------------------------------------------

        headers = [
            "ID",
            "Service",
            "Sub-Service",
            "Rating",
            "Comment",
            "Date"
        ]

        ws.append(
            headers
        )

        for cell in ws[8]:

            cell.font = Font(
                bold=True
            )

        # ----------------------------------------------------
        # RECORDS
        # ----------------------------------------------------

        for record in records:

            ws.append(
                [
                    record["id"],
                    record["service_name"],
                    record["sub_service"],
                    record["rating"],
                    record["comment"],
                    str(record["timestamp"])
                ]
            )

        # ----------------------------------------------------
        # COLUMN WIDTHS
        # ----------------------------------------------------

        widths = {
            "A": 10,
            "B": 25,
            "C": 35,
            "D": 15,
            "E": 60,
            "F": 25
        }

        for column, width in widths.items():

            ws.column_dimensions[
                column
            ].width = width

        # ----------------------------------------------------
        # FREEZE HEADER
        # ----------------------------------------------------

        ws.freeze_panes = "A9"

        wb.save(
            file_stream
        )

        file_stream.seek(0)

        return send_file(
            file_stream,
            as_attachment=True,
            download_name=(
                "police_feedback_report.xlsx"
            ),
            mimetype=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            )
        )

    # ========================================================
    # PDF
    # ========================================================

    elif format_type == "pdf":

        from reportlab.lib.pagesizes import letter
        from reportlab.platypus import (
            SimpleDocTemplate,
            Paragraph,
            Spacer
        )
        from reportlab.lib.styles import (
            getSampleStyleSheet,
            ParagraphStyle
        )

        doc = SimpleDocTemplate(
            file_stream,
            pagesize=letter,
            rightMargin=36,
            leftMargin=36,
            topMargin=36,
            bottomMargin=36
        )

        styles = getSampleStyleSheet()

        title_style = ParagraphStyle(
            "TitleStyle",
            parent=styles["Heading1"],
            fontSize=16,
            leading=20,
            spaceAfter=15
        )

        normal_style = ParagraphStyle(
            "NormalStyle",
            parent=styles["Normal"],
            fontSize=9,
            leading=12,
            spaceAfter=6
        )

        story = []

        story.append(
            Paragraph(
                "Ethiopian Federal Police - "
                "Feedback Report",
                title_style
            )
        )

        story.append(
            Paragraph(
                (
                    f"<b>Service Filter:</b> "
                    f"{html.escape(service_filter)}"
                ),
                normal_style
            )
        )

        story.append(
            Paragraph(
                (
                    f"<b>Rating Filter:</b> "
                    f"{html.escape(rating_filter or 'All')}"
                ),
                normal_style
            )
        )

        story.append(
            Paragraph(
                (
                    f"<b>Date From:</b> "
                    f"{html.escape(date_from or 'All')}"
                ),
                normal_style
            )
        )

        story.append(
            Paragraph(
                (
                    f"<b>Date To:</b> "
                    f"{html.escape(date_to or 'All')}"
                ),
                normal_style
            )
        )

        story.append(
            Paragraph(
                (
                    f"<b>Total Records:</b> "
                    f"{len(records)}"
                ),
                normal_style
            )
        )

        story.append(
            Spacer(1, 10)
        )

        if not records:

            story.append(
                Paragraph(
                    "No feedback records match "
                    "the selected filters.",
                    normal_style
                )
            )

        else:

            for record in records:

                service_text = html.escape(
                    str(
                        record["service_name"]
                        or ""
                    )
                )

                sub_service_text = html.escape(
                    str(
                        record["sub_service"]
                        or ""
                    )
                )

                rating_text = html.escape(
                    str(
                        record["rating"]
                        or ""
                    )
                )

                timestamp_text = html.escape(
                    str(
                        record["timestamp"]
                        or ""
                    )
                )

                story.append(
                    Paragraph(
                        (
                            f"<b>Service:</b> "
                            f"{service_text}"
                        ),
                        normal_style
                    )
                )

                story.append(
                    Paragraph(
                        (
                            f"<b>Sub-Service:</b> "
                            f"{sub_service_text}"
                        ),
                        normal_style
                    )
                )

                story.append(
                    Paragraph(
                        (
                            f"<b>Rating:</b> "
                            f"{rating_text}"
                        ),
                        normal_style
                    )
                )

                story.append(
                    Paragraph(
                        (
                            f"<b>Date:</b> "
                            f"{timestamp_text}"
                        ),
                        normal_style
                    )
                )

                if record["comment"]:

                    comment_text = html.escape(
                        str(
                            record["comment"]
                        )
                    )

                    story.append(
                        Paragraph(
                            (
                                f"<b>Comment:</b> "
                                f"{comment_text}"
                            ),
                            normal_style
                        )
                    )

                story.append(
                    Spacer(1, 8)
                )

        doc.build(
            story
        )

        file_stream.seek(0)

        return send_file(
            file_stream,
            as_attachment=True,
            download_name=(
                "police_feedback_report.pdf"
            ),
            mimetype="application/pdf"
        )

    # ========================================================
    # INVALID FORMAT
    # ========================================================

    return (
        "Invalid format type specified. "
        "Use 'word', 'excel', or 'pdf'.",
        400
    )


# ============================================================
# ETHIOPIAN CALENDAR DISPLAY
# ============================================================

def gregorian_to_ethiopian(dt):

    """
    Display-oriented Ethiopian calendar conversion.

    Note:
    This preserves the original project's simple display
    approach and is not intended as a full astronomical
    Ethiopian calendar conversion.
    """

    g_year = dt.year
    g_month = dt.month
    g_day = dt.day

    if (
        g_month < 9
        or (
            g_month == 9
            and g_day < 11
        )
    ):

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

    month_index = (
        g_month - 9
    ) % 13

    return (
        f"{dt.day} "
        f"{et_months[month_index]} "
        f"{et_year} — "
        f"{dt.strftime('%H:%M:%S')}"
    )


# ============================================================
# ADMIN AUDIT LOGS
# ============================================================

@app.route("/admin/audit-logs")
def admin_audit_logs():

    logged_in_admin = session.get(
        "admin_user"
    )

    admin_credentials = (
        get_current_admin_credentials()
    )

    if (
        not logged_in_admin
        or
        logged_in_admin not in admin_credentials
    ):

        return redirect(
            url_for("admin_login")
        )

    conn = get_db_connection()
    cursor = conn.cursor()

    try:

        if logged_in_admin == "admin gen":

            cursor.execute(
                """
                SELECT
                    id,
                    admin_user,
                    action_description,
                    timestamp
                FROM audit_logs
                ORDER BY timestamp DESC, id DESC
                """
            )

        else:

            cursor.execute(
                """
                SELECT
                    id,
                    admin_user,
                    action_description,
                    timestamp
                FROM audit_logs
                WHERE admin_user = ?
                ORDER BY timestamp DESC, id DESC
                """,
                (
                    logged_in_admin,
                )
            )

        rows = cursor.fetchall()

        formatted_logs = []

        for row in rows:

            timestamp = row[
                "timestamp"
            ]

            try:

                parsed_timestamp = (
                    parse_database_timestamp(
                        timestamp
                    )
                )

                if parsed_timestamp:

                    ethiopian_timestamp = (
                        gregorian_to_ethiopian(
                            parsed_timestamp
                        )
                    )

                else:

                    ethiopian_timestamp = (
                        timestamp
                        or
                        "N/A"
                    )

            except Exception:

                ethiopian_timestamp = (
                    timestamp
                    or
                    "N/A"
                )

            formatted_logs.append(
                {
                    "id": row["id"],

                    "admin_user":
                        row["admin_user"],

                    "action":
                        row["action_description"],

                    "action_description":
                        row["action_description"],

                    "timestamp":
                        timestamp,

                    "ethiopian_timestamp":
                        ethiopian_timestamp
                }
            )

    finally:

        cursor.close()
        conn.close()

    return render_template(
        "admin_audit_logs.html",
        audit_logs=formatted_logs,
        logs=formatted_logs,
        current_admin=logged_in_admin
    )


# ============================================================
# HEALTH CHECK
# ============================================================

@app.route("/health")
def health_check():

    return jsonify(
        {
            "status": "ok",
            "service": (
                "Ethiopian Federal Police "
                "Citizen Feedback System"
            )
        }
    ), 200


# ============================================================
# OPTIONAL TEXT-TO-SPEECH API
#
# This keeps compatibility with a frontend that calls:
# POST /text-to-speech
# {
#     "text": "...",
#     "lang": "en"
# }
# ============================================================

@app.route(
    "/text-to-speech",
    methods=["POST"]
)
def text_to_speech():

    try:

        data = request.get_json(
            silent=True
        ) or {}

        text = str(
            data.get(
                "text",
                ""
            )
        ).strip()

        lang = str(
            data.get(
                "lang",
                "en"
            )
        ).strip()

        if not text:

            return jsonify(
                {
                    "status": "error",
                    "message": "Text is required."
                }
            ), 400

        # gTTS commonly uses:
        # am = Amharic
        # en = English
        # om = Oromo may depend on gTTS support.
        #
        # For unsupported languages, fall back to English.

        supported_languages = {
            "en",
            "am",
            "fr",
            "de",
            "es",
            "it",
            "pt",
            "ar",
            "ru",
            "zh-CN",
            "hi",
            "ja",
            "ko"
        }

        if lang not in supported_languages:

            lang = "en"

        audio_buffer = io.BytesIO()

        tts = gTTS(
            text=text,
            lang=lang
        )

        tts.write_to_fp(
            audio_buffer
        )

        audio_buffer.seek(0)

        return send_file(
            audio_buffer,
            mimetype="audio/mpeg"
        )

    except Exception as e:

        return jsonify(
            {
                "status": "error",
                "message": str(e)
            }
        ), 500


# ============================================================
# APPLICATION ENTRY POINT
# ============================================================

if __name__ == "__main__":

    port = int(
        os.environ.get(
            "PORT",
            "5000"
        )
    )

    debug_mode = (
        os.environ.get(
            "FLASK_DEBUG",
            "false"
        ).lower()
        == "true"
    )

    app.run(
        host="0.0.0.0",
        port=port,
        debug=debug_mode
    )