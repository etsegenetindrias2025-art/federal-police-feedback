from flask import Flask, render_template, request, jsonify, session, redirect, url_for, send_file
import psycopg2
import psycopg2.extras
import os
import io
import qrcode
from gtts import gTTS
import socket
from datetime import datetime

app = Flask(__name__)
app.secret_key = 'federal_police_secret_key'

# PostgreSQL connection configurations
DB_HOST = os.environ.get('DB_HOST', 'localhost')
DB_NAME = os.environ.get('DB_NAME', 'federal_police_db')
DB_USER = os.environ.get('DB_USER', 'postgres')
DB_PASSWORD = os.environ.get('DB_PASSWORD', '2323')  # Update with your exact password
DB_PORT = os.environ.get('DB_PORT', '5432')

def get_db_connection():
    conn = psycopg2.connect(
        host=DB_HOST,
        database=DB_NAME,
        user=DB_USER,
        password=DB_PASSWORD,
        port=DB_PORT,
        cursor_factory=psycopg2.extras.RealDictCursor
    )
    return conn

def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS feedbacks (
            id SERIAL PRIMARY KEY,
            service_name TEXT NOT NULL,
            sub_service TEXT,
            rating TEXT NOT NULL,
            comment TEXT,
            is_read BOOLEAN DEFAULT FALSE,
            feedback_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS services (
            id SERIAL PRIMARY KEY,
            service_key TEXT UNIQUE NOT NULL,
            service_name TEXT NOT NULL
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS sub_services (
            id SERIAL PRIMARY KEY,
            service_key TEXT NOT NULL,
            sub_service_key TEXT NOT NULL,
            sub_service_name TEXT NOT NULL,
            UNIQUE(service_key, sub_service_key)
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
            INSERT INTO services (service_key, service_name) VALUES (%s, %s)
            ON CONFLICT (service_key) DO NOTHING
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
                INSERT INTO sub_services (service_key, sub_service_key, sub_service_name) VALUES (%s, %s, %s)
                ON CONFLICT (service_key, sub_service_key) DO NOTHING
            """, (s_key, sub_key, sub_name))

    cursor.execute('''
        DO $$ 
        BEGIN 
            IF NOT EXISTS (SELECT 1 FROM information_schema.columns WHERE table_name='feedbacks' and column_name='is_read') THEN 
                ALTER TABLE feedbacks ADD COLUMN is_read BOOLEAN DEFAULT FALSE; 
            END IF; 
            IF NOT EXISTS (SELECT 1 FROM information_schema.columns WHERE table_name='feedbacks' and column_name='feedback_date') THEN 
                ALTER TABLE feedbacks ADD COLUMN feedback_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP; 
            END IF;
        END $$;
    ''')
        
    conn.commit()
    cursor.close()
    conn.close()

init_db()

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
        "admin gen": {
            "password": "1234", 
            "type": "general",
            "service": "all",
            "sub_service": "all",
            "title": "General Admin Dashboard (All Services & Sub-services)"
        },
        "admin pol": {
            "password": "1234",
            "type": "service",
            "service": "police_clearance",
            "sub_service": "all",
            "title": "Police Clearance Department Admin"
        },
        "admin com": {
            "password": "1234",
            "type": "service",
            "service": "complaint",
            "sub_service": "all",
            "title": "Complaint Department Admin"
        },
        "admin hos": {
            "password": "1234",
            "type": "service",
            "service": "hospital",
            "sub_service": "all",
            "title": "Hospital Department Admin"
        },
        "admin log": {
            "password": "1234",
            "type": "service",
            "service": "logistics",
            "sub_service": "all",
            "title": "Logistics Department Admin"
        },
        "admin edu": {
            "password": "1234",
            "type": "service",
            "service": "education_training",
            "sub_service": "all",
            "title": "Education & Training Department Admin"
        },
        "admin oth": {
            "password": "1234",
            "type": "service",
            "service": "other",
            "sub_service": "all",
            "title": "Other Department Admin"
        },
        "admin new": {
            "password": "1234", "type": "sub_service", "service": "police_clearance", "sub_service": "new_clearance", "title": "Sub-Service Admin: New Police Clearance"
        },
        "admin ren": {
            "password": "1234", "type": "sub_service", "service": "police_clearance", "sub_service": "renewal", "title": "Sub-Service Admin: Renewal"
        },
        "admin cri": {
            "password": "1234", "type": "sub_service", "service": "police_clearance", "sub_service": "criminal_record", "title": "Sub-Service Admin: Criminal Record Verification"
        },
        "admin fin": {
            "password": "1234", "type": "sub_service", "service": "police_clearance", "sub_service": "fingerprint", "title": "Sub-Service Admin: Fingerprint Registration"
        },
        "admin doc": {
            "password": "1234", "type": "sub_service", "service": "police_clearance", "sub_service": "document_collection", "title": "Sub-Service Admin: Document Collection"
        },
        "admin pub": {
            "password": "1234", "type": "sub_service", "service": "complaint", "sub_service": "public_office", "title": "Sub-Service Admin: Public Complaint Office"
        },
        "admin onl": {
            "password": "1234", "type": "sub_service", "service": "complaint", "sub_service": "online_followup", "title": "Sub-Service Admin: Online Complaint Follow-up"
        },
        "admin inv": {
            "password": "1234", "type": "sub_service", "service": "complaint", "sub_service": "investigation", "title": "Sub-Service Admin: Investigation"
        },
        "admin res": {
            "password": "1234", "type": "sub_service", "service": "complaint", "sub_service": "resolution", "title": "Sub-Service Admin: Resolution"
        },
        "admin opd": {
            "password": "1234", "type": "sub_service", "service": "hospital", "sub_service": "opd", "title": "Sub-Service Admin: OPD"
        },
        "admin eme": {
            "password": "1234", "type": "sub_service", "service": "hospital", "sub_service": "emergency", "title": "Sub-Service Admin: Emergency"
        },
        "admin pha": {
            "password": "1234", "type": "sub_service", "service": "hospital", "sub_service": "pharmacy", "title": "Sub-Service Admin: Pharmacy"
        },
        "admin lab": {
            "password": "1234", "type": "sub_service", "service": "hospital", "sub_service": "laboratory", "title": "Sub-Service Admin: Laboratory"
        },
        "admin med": {
            "password": "1234", "type": "sub_service", "service": "hospital", "sub_service": "medical_exam", "title": "Sub-Service Admin: Medical Examination"
        },
        "admin veh": {
            "password": "1234", "type": "sub_service", "service": "logistics", "sub_service": "vehicle_mgmt", "title": "Sub-Service Admin: Vehicle Management"
        },
        "admin gar": {
            "password": "1234", "type": "sub_service", "service": "logistics", "sub_service": "garage", "title": "Sub-Service Admin: Garage"
        },
        "admin equ": {
            "password": "1234", "type": "sub_service", "service": "logistics", "sub_service": "equipment_dist", "title": "Sub-Service Admin: Equipment Distribution"
        },
        "admin pro": {
            "password": "1234", "type": "sub_service", "service": "logistics", "sub_service": "procurement", "title": "Sub-Service Admin: Procurement"
        },
        "admin stu": {
            "password": "1234", "type": "sub_service", "service": "education_training", "sub_service": "student_reg", "title": "Sub-Service Admin: Student Registration"
        },
        "admin tra": {
            "password": "1234", "type": "sub_service", "service": "education_training", "sub_service": "training", "title": "Sub-Service Admin: Training"
        },
        "admin cer": {
            "password": "1234", "type": "sub_service", "service": "education_training", "sub_service": "certificates", "title": "Sub-Service Admin: Certificates"
        },
        "admin exa": {
            "password": "1234", "type": "sub_service", "service": "education_training", "sub_service": "examination", "title": "Sub-Service Admin: Examination"
        },
        "admin aca": {
            "password": "1234", "type": "sub_service", "service": "education_training", "sub_service": "academic_records", "title": "Sub-Service Admin: Academic Records"
        },
        "admin rec": {
            "password": "1234", "type": "sub_service", "service": "other", "sub_service": "reception", "title": "Sub-Service Admin: Reception"
        },
        "admin ict": {
            "password": "1234", "type": "sub_service", "service": "other", "sub_service": "ict_support", "title": "Sub-Service Admin: ICT Support"
        },
        "admin hr": {
            "password": "1234", "type": "sub_service", "service": "other", "sub_service": "hr", "title": "Sub-Service Admin: HR"
        },
        "admin fnn": {
            "password": "1234", "type": "sub_service", "service": "other", "sub_service": "finance", "title": "Sub-Service Admin: Finance"
        },
        "admin adm": {
            "password": "1234", "type": "sub_service", "service": "other", "sub_service": "admin", "title": "Sub-Service Admin: Administration"
        }
    }
    return credentials

# Public Routes
@app.route('/')
def welcome():
    return render_template('welcome.html')

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

@app.route('/submit-feedback', methods=['POST'])
def submit_feedback():
    try:
        data = request.get_json()
        if not data:
            return jsonify({"status": "error", "message": "No data received"}), 400
            
        rating = data.get('rating', '😊')
        comment = data.get('comment', 'No comment provided.')
        sub_service = data.get('sub_service', 'general_service')
        url_service = data.get('service') or data.get('service_name', 'police_clearance')

        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO feedbacks (service_name, sub_service, rating, comment, is_read, feedback_date) VALUES (%s, %s, %s, %s, FALSE, CURRENT_TIMESTAMP)",
            (str(url_service), str(sub_service), str(rating), str(comment))
        )
        conn.commit()
        cursor.close()
        conn.close()
        
        return jsonify({"status": "success", "message": "Feedback saved successfully!"})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/api/unread-count')
def api_unread_count():
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT COUNT(*) AS count FROM feedbacks WHERE is_read = FALSE")
        result = cursor.fetchone()
        unread_count = result['count'] if result else 0
    except Exception:
        unread_count = 0
    finally:
        cursor.close()
        conn.close()
        
    return jsonify({"unread_count": unread_count})

# Integrated Reports Filter Route
@app.route('/api/reports/filter', methods=['GET'])
def filter_reports():
    # Get parameters from request (e.g., ?year=2026&month=08&day=05)
    day = request.args.get('day')
    month = request.args.get('month')
    year = request.args.get('year')

    conn = get_db_connection()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)

    # Base query mapped to table schema
    query = "SELECT id, feedback_date, service_name, rating, is_read FROM feedbacks WHERE 1=1"
    params = []

    if year:
        query += " AND EXTRACT(YEAR FROM feedback_date) = %s"
        params.append(year)
    if month:
        query += " AND EXTRACT(MONTH FROM feedback_date) = %s"
        params.append(month)
    if day:
        query += " AND EXTRACT(DAY FROM feedback_date) = %s"
        params.append(day)

    cur.execute(query, params)
    rows = cur.fetchall()
    
    cur.close()
    conn.close()
    
    # Format results as JSON
    results = []
    for row in rows:
        results.append({
            "id": row['id'],
            "date": str(row['feedback_date']),
            "service_type": row['service_name'],
            "rating": row['rating'],
            "status": "Read" if row['is_read'] else "Unread"
        })

    return jsonify(results)

# Admin Authentication & Dashboard
@app.route('/admin/login', methods=['GET', 'POST'])
def admin_login():
    error = None
    admin_credentials = get_admin_credentials()
    
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        
        if username in admin_credentials and admin_credentials[username]['password'] == password:
            session['admin_user'] = username
            return redirect(url_for('admin_dashboard'))
        else:
            error = "Invalid Username or Password. Please try again."
            
    return render_template('admin_login.html', error=error)

@app.route('/admin/logout')
def admin_logout():
    session.pop('admin_user', None)
    return redirect(url_for('admin_login'))

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
    feedbacks = [] 
    
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

        if admin_type == 'general':
            if selected_filter != 'all':
                for fb in all_feedbacks:
                    if str(fb['service_name']).strip().lower() == selected_filter:
                        feedbacks.append(fb)
            else:
                feedbacks = all_feedbacks
            total_feedbacks_count = len(all_feedbacks)
            
        elif admin_type == 'service':
            for fb in all_feedbacks:
                if str(fb['service_name']).strip().lower() == assigned_service:
                    feedbacks.append(fb)
            total_feedbacks_count = len(feedbacks)
            
        elif admin_type == 'sub_service':
            for fb in all_feedbacks:
                if str(fb['service_name']).strip().lower() == assigned_service and str(fb['sub_service']).strip().lower() == assigned_sub:
                    feedbacks.append(fb)
            total_feedbacks_count = len(feedbacks)
                
    except Exception as e:
        feedbacks = []
        
    cursor.close()
    conn.close()
    
    return render_template('admin_dashboard.html', 
                           feedbacks=feedbacks, 
                           admin_title=admin_title, 
                           is_general=is_general_admin,
                           service_map=service_map,
                           sub_service_map=sub_service_map,
                           chart_data=chart_data,
                           service_counts=service_counts,
                           total_feedbacks_count=total_feedbacks_count,
                           unread_notifications_count=unread_notifications_count,
                           current_filter=selected_filter)

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
        cursor.execute("UPDATE feedbacks SET is_read = TRUE WHERE is_read = FALSE")
    elif admin_type == 'service':
        cursor.execute("UPDATE feedbacks SET is_read = TRUE WHERE is_read = FALSE AND service_name = %s", (assigned_service,))
    elif admin_type == 'sub_service':
        cursor.execute("UPDATE feedbacks SET is_read = TRUE WHERE is_read = FALSE AND service_name = %s AND sub_service = %s", (assigned_service, assigned_sub))
        
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
    
    return render_template('admin_notifications.html', 
                           notifications=notifications, 
                           service_map=service_map, 
                           sub_service_map=sub_service_map)

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
            else:
                error = "Current password is incorrect."
                
        elif is_general_admin:
            if action == 'add_service':
                s_key = request.form.get('service_key').strip().lower().replace(" ", "_")
                s_name = request.form.get('service_name').strip()
                if s_key and s_name:
                    try:
                        conn = get_db_connection()
                        cursor = conn.cursor()
                        cursor.execute("INSERT INTO services (service_key, service_name) VALUES (%s, %s)", (s_key, s_name))
                        conn.commit()
                        cursor.close()
                        conn.close()
                        message = f"Service '{s_name}' added successfully!"
                    except Exception as e:
                        conn.rollback()
                        error = "Service key already exists or invalid input."
                else:
                    error = "All fields are required to add a service."
                    
            elif action == 'update_service':
                s_key = request.form.get('service_key')
                s_name = request.form.get('new_service_name').strip()
                if s_key and s_name:
                    conn = get_db_connection()
                    cursor = conn.cursor()
                    cursor.execute("UPDATE services SET service_name = %s WHERE service_key = %s", (s_name, s_key))
                    conn.commit()
                    cursor.close()
                    conn.close()
                    message = "Service updated successfully!"
                else:
                    error = "Invalid service update details."

            elif action == 'delete_service':
                s_key = request.form.get('service_key')
                if s_key:
                    conn = get_db_connection()
                    cursor = conn.cursor()
                    cursor.execute("DELETE FROM services WHERE service_key = %s", (s_key,))
                    cursor.execute("DELETE FROM sub_services WHERE service_key = %s", (s_key,))
                    conn.commit()
                    cursor.close()
                    conn.close()
                    message = "Service and its sub-services deleted successfully!"
                else:
                    error = "Select a service to delete."

            elif action == 'add_sub_service':
                parent_service = request.form.get('parent_service_key')
                sub_key = request.form.get('sub_service_key').strip().lower().replace(" ", "_")
                sub_name = request.form.get('sub_service_name').strip()
                if parent_service and sub_key and sub_name:
                    try:
                        conn = get_db_connection()
                        cursor = conn.cursor()
                        cursor.execute("INSERT INTO sub_services (service_key, sub_service_key, sub_service_name) VALUES (%s, %s, %s)", 
                                       (parent_service, sub_key, sub_name))
                        conn.commit()
                        cursor.close()
                        conn.close()
                        message = f"Sub-service '{sub_name}' added successfully!"
                    except Exception as e:
                        conn.rollback()
                        error = "Sub-service key already exists under this service."
                else:
                    error = "All fields are required to add a sub-service."

            elif action == 'update_sub_service':
                parent_service = request.form.get('parent_service_key')
                sub_key = request.form.get('sub_service_key')
                sub_name = request.form.get('new_sub_service_name').strip()
                if parent_service and sub_key and sub_name:
                    conn = get_db_connection()
                    cursor = conn.cursor()
                    cursor.execute("UPDATE sub_services SET sub_service_name = %s WHERE service_key = %s AND sub_service_key = %s", 
                                   (sub_name, parent_service, sub_key))
                    conn.commit()
                    cursor.close()
                    conn.close()
                    message = "Sub-service updated successfully!"
                else:
                    error = "Invalid sub-service update details."

            elif action == 'delete_sub_service':
                parent_service = request.form.get('parent_service_key')
                sub_key = request.form.get('sub_service_key')
                if parent_service and sub_key:
                    conn = get_db_connection()
                    cursor = conn.cursor()
                    cursor.execute("DELETE FROM sub_services WHERE service_key = %s AND sub_service_key = %s", (parent_service, sub_key))
                    conn.commit()
                    cursor.close()
                    conn.close()
                    message = "Sub-service deleted successfully!"
                else:
                    error = "Select a sub-service to delete."
        else:
            error = "Unauthorized action: Only General Admin can perform CRUD operations on services."
                
    service_map = get_service_map()
    sub_service_map = get_sub_service_map()
    
    return render_template('admin_settings.html', 
                           message=message, 
                           error=error, 
                           service_map=service_map, 
                           sub_service_map=sub_service_map,
                           is_general=is_general_admin)

@app.route('/admin/export/<format>')
def export_report(format):
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
        service_filter = request.args.get('service', 'all')
        if service_filter == 'all':
            cursor.execute("SELECT * FROM feedbacks ORDER BY timestamp ASC")
        else:
            cursor.execute("SELECT * FROM feedbacks WHERE service_name = %s ORDER BY timestamp ASC", (service_filter,))
    elif admin_type == 'service':
        cursor.execute("SELECT * FROM feedbacks WHERE service_name = %s ORDER BY timestamp ASC", (assigned_service,))
    elif admin_type == 'sub_service':
        cursor.execute("SELECT * FROM feedbacks WHERE service_name = %s AND sub_service = %s ORDER BY timestamp ASC", (assigned_service, assigned_sub))
        
    feedbacks = cursor.fetchall()
    cursor.close()
    conn.close()

    service_map = get_service_map()
    sub_service_map = get_sub_service_map()

    rating_text_map = {
        "😍": "Very Satisfied",
        "😊": "Satisfied",
        "😐": "Neutral",
        "🙁": "Not Satisfied",
        "😠": "Very Dissatisfied"
    }

    if format == 'excel':
        import pandas as pd
        data = []
        for index, fb in enumerate(feedbacks, start=1):
            raw_rating = str(fb['rating']).strip()
            mapped_rating = rating_text_map.get(raw_rating, raw_rating)
            rating_display = f"{raw_rating} ({mapped_rating})" if mapped_rating else raw_rating

            data.append({
                'ID': index,
                'Service': service_map.get(fb['service_name'], fb['service_name']),
                'Sub-Service': sub_service_map.get(fb['service_name'], {}).get(fb['sub_service'], fb['sub_service']),
                'Rating': rating_display,
                'Comment': fb['comment'],
                'Timestamp': str(fb['timestamp'])
            })
        
        df = pd.DataFrame(data)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Feedback Report')
            
            worksheet = writer.sheets['Feedback Report']
            for col in worksheet.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = col[0].column_letter
                worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)
                
        output.seek(0)
        
        return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                         as_attachment=True, download_name='feedback_report.xlsx')

    elif format == 'pdf':
        from reportlab.lib.pagesizes import letter
        from reportlab.pdfgen import canvas
        
        output = io.BytesIO()
        p = canvas.Canvas(output, pagesize=letter)
        p.drawString(50, 750, "Ethiopian Federal Police - Report")
        
        y = 710
        for index, fb in enumerate(feedbacks, start=1):
            s_name = service_map.get(fb['service_name'], fb['service_name'])
            sub_name = sub_service_map.get(fb['service_name'], {}).get(fb['sub_service'], fb['sub_service'])
            raw_rating = str(fb['rating']).strip()
            mapped_rating = rating_text_map.get(raw_rating, raw_rating)
            p.drawString(50, y, f"ID: {index} | Service: {s_name} ({sub_name}) | Rating: {raw_rating} ({mapped_rating})")
            y -= 20
            if y < 50:
                p.showPage()
                y = 750
        p.save()
        output.seek(0)
        
        return send_file(output, mimetype='application/pdf', as_attachment=True, download_name='feedback_report.pdf')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)