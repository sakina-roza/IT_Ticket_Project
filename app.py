from flask import Flask, render_template, jsonify, request, redirect, url_for, session, Response
import pandas as pd
import os, io, json, random, re, secrets, uuid
from datetime import datetime, timedelta
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = "it_project_2026_secure"

EXCEL_FILE   = 'IT_Ticket_Performance_Data.xlsx'
USERS_FILE   = 'users.json'
COMMENTS_FILE= 'comments.json'
HISTORY_FILE = 'ticket_history.json'
NOTIF_FILE   = 'notifications.json'
NOTES_FILE   = 'internal_notes.json'
CANNED_FILE  = 'canned_responses.json'
CHAT_FILE    = 'chat_messages.json'

DEFAULT_CANNED = [
    {"id": "cr1", "label": "Acknowledge",          "body": "Thank you for reaching out. I've picked up your ticket and will begin investigating shortly."},
    {"id": "cr2", "label": "Need More Info",        "body": "Could you please provide more details? Specifically: when did it start, what steps you've already tried, and any error messages you're seeing."},
    {"id": "cr3", "label": "Resolved",              "body": "I'm happy to let you know your issue has been resolved. Please reopen the ticket if you experience any further problems."},
    {"id": "cr4", "label": "Escalating",            "body": "Your ticket is being escalated to a senior technician for further investigation. We'll keep you updated on progress."},
    {"id": "cr5", "label": "Scheduled Maintenance", "body": "This issue is related to scheduled maintenance. Normal service will be restored shortly. Apologies for the inconvenience."},
]
UPLOAD_FOLDER= os.path.join('static', 'uploads')
ALLOWED_EXTENSIONS = {'png','jpg','jpeg','gif','pdf','txt','docx','xlsx','zip','log'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

SLA_HOURS = {'Critical': 4, 'High': 8, 'Medium': 24, 'Low': 72}

# ── Flask-Mail (optional) ──────────────────────────────────────────────────────
MAIL_ENABLED = False
mail = None
try:
    from flask_mail import Mail, Message as MailMessage
    app.config['MAIL_SERVER']         = os.environ.get('MAIL_SERVER',   'smtp.gmail.com')
    app.config['MAIL_PORT']           = int(os.environ.get('MAIL_PORT', 587))
    app.config['MAIL_USE_TLS']        = True
    app.config['MAIL_USERNAME']       = os.environ.get('MAIL_USERNAME', '')
    app.config['MAIL_PASSWORD']       = os.environ.get('MAIL_PASSWORD', '')
    app.config['MAIL_DEFAULT_SENDER'] = os.environ.get('MAIL_USERNAME', 'noreply@it-tickets.com')
    if app.config['MAIL_USERNAME']:
        mail = Mail(app)
        MAIL_ENABLED = True
except ImportError:
    pass

# ── Keyword classifier rules ───────────────────────────────────────────────────
CLASSIFY_RULES = {
    'Network':  ['network','wifi','internet','vpn','firewall','ping','dns','bandwidth','router','switch','lan','wan'],
    'Hardware': ['hardware','laptop','computer','monitor','keyboard','mouse','screen','device','cable','battery','blue screen','bsod','overheating'],
    'Software': ['software','application','app','install','update','error','bug','program','windows','office','browser','license'],
    'Security': ['security','virus','malware','phishing','hack','breach','unauthorized','suspicious','ransomware','spam'],
    'Access':   ['access','permission','login','account','locked','credential','admin rights','2fa','mfa'],
    'Email':    ['email','outlook','mail','inbox','calendar','teams','exchange'],
    'Printer':  ['printer','print','scan','copier','toner','paper jam','fax'],
}

# ── JSON helpers ───────────────────────────────────────────────────────────────

def _read_json(path, default=None):
    if default is None:
        default = {}
    if not os.path.exists(path):
        return default
    try:
        with open(path, encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return default

def _write_json(path, data):
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

# ── User helpers ───────────────────────────────────────────────────────────────

def get_users():
    if not os.path.exists(USERS_FILE):
        default = {
            "admin@it.com": {
                "password": generate_password_hash("password123"),
                "role": "admin", "name": "Admin",
                "theme": "dark", "email_notifications": True,
                "reset_token": None, "reset_token_expiry": None
            }
        }
        _write_json(USERS_FILE, default)
        return default
    return _read_json(USERS_FILE, {})

def save_users(users):
    _write_json(USERS_FILE, users)

def current_user_info():
    email = session.get('user')
    if not email:
        return None
    u = get_users().get(email, {})
    return {
        "email": email,
        "name":  u.get("name", email),
        "role":  u.get("role", "user"),
        "theme": u.get("theme", "dark"),
        "email_notifications": u.get("email_notifications", True)
    }

def is_admin():
    u = current_user_info()
    return u and u["role"] == "admin"

def is_agent():
    u = current_user_info()
    return u and u["role"] in ("agent", "admin")

def is_admin_or_agent():
    u = current_user_info()
    return u and u["role"] in ("admin", "agent")

# ── Agent workload + skill-based routing ───────────────────────────────────────

def _get_agent_workload(email):
    """Count open/in-progress tickets currently assigned to an agent."""
    df = get_safe_data()
    if df.empty or 'Assigned_To' not in df.columns:
        return 0
    users = get_users()
    name  = users.get(email, {}).get('name', '')
    mask  = (
        (df['Assigned_To'].astype(str).str.strip() == email) |
        (df['Assigned_To'].astype(str).str.strip() == name)
    ) & (~df['Status'].isin(['Resolved', 'Closed']))
    return int(mask.sum())

def _find_best_agent(category, priority):
    """
    Strict role-based, availability-aware, load-balanced agent selection.
    Only assigns to an agent whose assigned roles include `category`.
    Returns the matching agent email with the lowest workload, or None.
    If no agent has a matching role the ticket goes to the manual queue.
    """
    users      = get_users()
    candidates = []
    for email, u in users.items():
        if u.get('role') != 'agent':
            continue
        if u.get('availability_status', 'online') != 'online':
            continue
        # Strict role check — must have the category in their assigned roles
        if category not in u.get('skills', []):
            continue
        workload = _get_agent_workload(email)
        if workload >= u.get('max_workload', 10):
            continue
        candidates.append((email, workload))

    if not candidates:
        return None

    # Lowest workload first
    candidates.sort(key=lambda x: x[1])
    return candidates[0][0]

# ── Context processor ──────────────────────────────────────────────────────────
@app.context_processor
def inject_user():
    u = current_user_info()
    return dict(
        current_user=u,
        is_admin_user=is_admin(),
        is_agent_user=(u and u["role"] == "agent")
    )

# ── Data helpers ───────────────────────────────────────────────────────────────

def get_safe_data():
    if not os.path.exists(EXCEL_FILE):
        df_empty = pd.DataFrame(columns=[
            'Ticket_ID','Status','Priority','Category',
            'Assigned_To','Created_Date','Resolution_Time_Hours',
            'Created_By','Description','Last_Updated','Attachments'
        ])
        df_empty.to_excel(EXCEL_FILE, index=False)
        return df_empty
    try:
        df = pd.read_excel(EXCEL_FILE, engine='openpyxl')
        for col in ('Created_By','Description','Last_Updated','Attachments'):
            if col not in df.columns:
                df[col] = ''
            else:
                df[col] = df[col].astype(object).fillna('')
        return df
    except Exception as e:
        print(f"Excel Read Error: {e}")
        return pd.DataFrame()

def generate_ticket_id():
    df = get_safe_data()
    existing = set(df['Ticket_ID'].astype(str).str.strip().values)
    for _ in range(100):
        tid = f"TKT-{datetime.now().strftime('%Y%m%d')}-{random.randint(1000,9999)}"
        if tid not in existing:
            return tid
    return f"TKT-{int(datetime.now().timestamp())}"

def get_sla_status(priority, res_hours):
    limit = SLA_HOURS.get(str(priority).strip(), 24)
    try:
        h = float(res_hours)
    except (TypeError, ValueError):
        return 'unknown'
    if h > limit:           return 'breached'
    elif h > limit * 0.75:  return 'near_breach'
    else:                   return 'on_track'

def _build_trend(df, period='weekly'):
    if 'Created_Date' not in df.columns:
        return {"labels": [], "values": []}
    try:
        df = df.copy()
        df['Created_Date'] = pd.to_datetime(df['Created_Date'], errors='coerce')
        df = df.dropna(subset=['Created_Date'])
        if period == 'monthly':
            df['period'] = df['Created_Date'].dt.to_period('M').astype(str)
            counts = df.groupby('period').size().sort_index().tail(12)
        else:
            counts = df.groupby(df['Created_Date'].dt.date).size().sort_index().tail(7)
        return {"labels": [str(d) for d in counts.index], "values": counts.tolist()}
    except Exception:
        return {"labels": [], "values": []}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# ── Comment helpers ────────────────────────────────────────────────────────────

def get_comments(ticket_id):
    data = _read_json(COMMENTS_FILE, {})
    return data.get(str(ticket_id), [])

def save_comment(ticket_id, author_email, author_name, body):
    data = _read_json(COMMENTS_FILE, {})
    tid  = str(ticket_id)
    if tid not in data:
        data[tid] = []
    data[tid].append({
        "id":          str(uuid.uuid4()),
        "author":      author_name,
        "author_email":author_email,
        "body":        body,
        "time":        datetime.now().strftime('%Y-%m-%d %H:%M')
    })
    _write_json(COMMENTS_FILE, data)

# ── History / Timeline helpers ─────────────────────────────────────────────────

def log_ticket_event(ticket_id, actor_email, actor_name, event_type, detail):
    data = _read_json(HISTORY_FILE, {})
    tid  = str(ticket_id)
    if tid not in data:
        data[tid] = []
    data[tid].append({
        "timestamp":  datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        "actor":      actor_name,
        "actor_email":actor_email,
        "event":      event_type,
        "detail":     detail
    })
    _write_json(HISTORY_FILE, data)

# ── Notification helpers ───────────────────────────────────────────────────────

def add_notification(user_email, message, ticket_id, notif_type='info'):
    data = _read_json(NOTIF_FILE, {})
    if user_email not in data:
        data[user_email] = []
    data[user_email].append({
        "id":        str(uuid.uuid4()),
        "message":   message,
        "ticket_id": str(ticket_id),
        "type":      notif_type,
        "read":      False,
        "time":      datetime.now().strftime('%Y-%m-%d %H:%M')
    })
    data[user_email] = data[user_email][-50:]
    _write_json(NOTIF_FILE, data)

# ── Email helper ───────────────────────────────────────────────────────────────

def send_email_notification(to_email, subject, body_html):
    users = get_users()
    u     = users.get(to_email, {})
    if not u.get('email_notifications', True):
        return
    if not MAIL_ENABLED:
        print(f"[Email skipped - mail not configured]: {subject} -> {to_email}")
        return
    try:
        msg = MailMessage(subject, recipients=[to_email], html=body_html)
        mail.send(msg)
    except Exception as e:
        print(f"[Email error]: {e}")

# ── Routes ─────────────────────────────────────────────────────────────────────

@app.route('/')
def root():
    return render_template('landing.html')

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        email = request.form.get('email', '').strip()
        pw    = request.form.get('password', '')
        users = get_users()
        u = users.get(email)
        if u and check_password_hash(u['password'], pw):
            session['user'] = email
            role = u.get('role', 'user')
            if role == 'admin':
                return redirect(url_for('dashboard'))
            elif role == 'agent':
                return redirect(url_for('agent_queue'))
            else:
                return redirect(url_for('my_tickets'))
        return render_template('login.html', error="Invalid email or password.")
    return render_template('login.html')

@app.route('/signup', methods=['GET', 'POST'])
def signup():
    if request.method == 'POST':
        name  = request.form.get('name', '').strip()
        email = request.form.get('email', '').strip()
        pw    = request.form.get('password', '')
        pw2   = request.form.get('confirm_password', '')
        if not name or not email or not pw:
            return render_template('signup.html', error="All fields are required.")
        if pw != pw2:
            return render_template('signup.html', error="Passwords do not match.")
        users = get_users()
        if email in users:
            return render_template('signup.html', error="Email already registered.")
        users[email] = {
            "password": generate_password_hash(pw),
            "role": "user", "name": name,
            "theme": "dark", "email_notifications": True,
            "reset_token": None, "reset_token_expiry": None
        }
        save_users(users)
        session['user'] = email
        return redirect(url_for('my_tickets'))
    return render_template('signup.html')

@app.route('/logout')
def logout():
    session.pop('user', None)
    return redirect(url_for('login'))

@app.route('/forgot_password', methods=['GET', 'POST'])
def forgot_password():
    if request.method == 'POST':
        email = request.form.get('email', '').strip()
        users = get_users()
        if email in users:
            token  = secrets.token_urlsafe(32)
            expiry = (datetime.now() + timedelta(hours=1)).isoformat()
            users[email]['reset_token']        = token
            users[email]['reset_token_expiry'] = expiry
            save_users(users)
            reset_url = url_for('reset_password', token=token, _external=True)
            send_email_notification(
                email,
                'Password Reset - IT Ticket System',
                f'<p>Click to reset your password: <a href="{reset_url}">{reset_url}</a></p>'
                f'<p>This link expires in 1 hour.</p>'
            )
            print(f"[Password Reset] Token for {email}: {reset_url}")
        return render_template('forgot_password.html',
                               message="If that email exists, a reset link has been sent.")
    return render_template('forgot_password.html')

@app.route('/reset_password/<token>', methods=['GET', 'POST'])
def reset_password(token):
    users        = get_users()
    target_email = None
    for email, u in users.items():
        if u.get('reset_token') == token:
            expiry = u.get('reset_token_expiry')
            if expiry:
                try:
                    if datetime.fromisoformat(expiry) > datetime.now():
                        target_email = email
                except Exception:
                    pass
            break
    if not target_email:
        return render_template('reset_password.html',
                               error="Invalid or expired reset link.", token=token)
    if request.method == 'POST':
        new_pw = request.form.get('password', '')
        if len(new_pw) < 6:
            return render_template('reset_password.html',
                                   error="Password must be at least 6 characters.", token=token)
        users[target_email]['password']           = generate_password_hash(new_pw)
        users[target_email]['reset_token']        = None
        users[target_email]['reset_token_expiry'] = None
        save_users(users)
        return redirect(url_for('login'))
    return render_template('reset_password.html', token=token)

# ── Admin-only routes ──────────────────────────────────────────────────────────

@app.route('/dashboard')
def dashboard():
    if 'user' not in session: return redirect(url_for('login'))
    if not is_admin():        return redirect(url_for('my_tickets'))
    return render_template('index.html')

@app.route('/tickets')
def tickets():
    if 'user' not in session:    return redirect(url_for('login'))
    if not is_admin_or_agent():  return redirect(url_for('my_tickets'))
    return render_template('tickets.html')

@app.route('/manage')
def manage():
    if 'user' not in session: return redirect(url_for('login'))
    if not is_admin():        return redirect(url_for('my_tickets'))
    return render_template('manage.html')

@app.route('/manage_users')
def manage_users():
    if 'user' not in session: return redirect(url_for('login'))
    if not is_admin():        return redirect(url_for('my_tickets'))
    return render_template('manage_users.html')


# ── Ticket detail (accessible by admin, agents, and ticket owner) ──────────────

@app.route('/tickets/<ticket_id>')
def ticket_detail(ticket_id):
    if 'user' not in session: return redirect(url_for('login'))
    return render_template('ticket_detail.html', tid=ticket_id)

# ── Profile ────────────────────────────────────────────────────────────────────

@app.route('/profile', methods=['GET', 'POST'])
def profile():
    if 'user' not in session: return redirect(url_for('login'))
    users = get_users()
    email = session['user']
    u     = users.get(email, {})

    if request.method == 'POST':
        action = request.form.get('action', '')

        if action == 'update_name':
            name = request.form.get('name', '').strip()
            if name:
                users[email]['name'] = name
                save_users(users)
                return jsonify({"success": True, "name": name})
            return jsonify({"success": False, "error": "Name cannot be empty"})

        elif action == 'change_password':
            current_pw = request.form.get('current_password', '')
            new_pw     = request.form.get('new_password', '')
            if not check_password_hash(u.get('password',''), current_pw):
                return jsonify({"success": False, "error": "Current password is incorrect."})
            if len(new_pw) < 6:
                return jsonify({"success": False, "error": "Password must be at least 6 characters."})
            users[email]['password'] = generate_password_hash(new_pw)
            save_users(users)
            return jsonify({"success": True})

        elif action == 'update_prefs':
            users[email]['theme']               = request.form.get('theme', 'dark')
            users[email]['email_notifications'] = request.form.get('email_notifications') == 'true'
            save_users(users)
            return jsonify({"success": True})

        return jsonify({"success": False, "error": "Unknown action"})

    return render_template('profile.html')

# ── Agent dashboard ────────────────────────────────────────────────────────────

@app.route('/agent/dashboard')
def agent_dashboard():
    if 'user' not in session:   return redirect(url_for('login'))
    if not is_admin_or_agent(): return redirect(url_for('my_tickets'))
    return render_template('agent_dashboard.html')

# ── Agent queue ────────────────────────────────────────────────────────────────

@app.route('/agent/queue')
def agent_queue():
    if 'user' not in session:    return redirect(url_for('login'))
    if not is_admin_or_agent():  return redirect(url_for('my_tickets'))
    return render_template('agent_queue.html')

# ── User routes ────────────────────────────────────────────────────────────────

@app.route('/create_ticket')
def create_ticket():
    if 'user' not in session: return redirect(url_for('login'))
    return render_template('create_ticket.html')

@app.route('/my_tickets')
def my_tickets():
    if 'user' not in session: return redirect(url_for('login'))
    return render_template('my_tickets.html')

# ── APIs ───────────────────────────────────────────────────────────────────────

@app.route('/api/tickets/count')
def ticket_count():
    if 'user' not in session: return jsonify({"count": 0})
    df = get_safe_data()
    return jsonify({"count": len(df)})

@app.route('/api/stats')
def stats():
    df       = get_safe_data()
    period   = request.args.get('period', 'weekly')
    date_from = request.args.get('date_from')
    date_to   = request.args.get('date_to')

    empty = {
        "stats": {"total":0,"open":0,"critical":0,"avg_res":0},
        "trend": {"labels":[],"values":[]},
        "recent": [], "sync_time": "NO DATA",
        "categories": [],
        "sla": {"on_track":0,"near_breach":0,"breached":0},
        "insights": {"tickets_today":0,"sla_breaches":0,"avg_res":0},
        "priority_counts": {}, "status_counts": {}, "category_counts": {}
    }
    if df.empty: return jsonify(empty)

    try:
        df['Status']   = df['Status'].fillna('Open')
        df['Priority'] = df['Priority'].fillna('Low')

        if date_from or date_to:
            df['Created_Date'] = pd.to_datetime(df['Created_Date'], errors='coerce')
            if date_from: df = df[df['Created_Date'] >= pd.to_datetime(date_from)]
            if date_to:   df = df[df['Created_Date'] <= pd.to_datetime(date_to)]

        sla_counts = {"on_track":0,"near_breach":0,"breached":0}
        if 'Resolution_Time_Hours' in df.columns:
            for _, row in df.iterrows():
                s = get_sla_status(row.get('Priority','Low'), row.get('Resolution_Time_Hours'))
                if s in sla_counts: sla_counts[s] += 1

        tickets_today = 0
        if 'Created_Date' in df.columns:
            df['Created_Date'] = pd.to_datetime(df['Created_Date'], errors='coerce')
            tickets_today = int((df['Created_Date'].dt.date == datetime.now().date()).sum())

        avg_res = round(float(df['Resolution_Time_Hours'].mean()), 1) \
                  if 'Resolution_Time_Hours' in df.columns else 0

        recent_df = df.tail(8).copy().fillna('')
        if 'Resolution_Time_Hours' in recent_df.columns:
            recent_df['SLA_Status'] = recent_df.apply(
                lambda r: get_sla_status(r['Priority'], r['Resolution_Time_Hours']), axis=1)

        return jsonify({
            "stats": {
                "total":    int(len(df)),
                "open":     int(len(df[df['Status'].str.contains('Open', case=False, na=False)])),
                "critical": int(len(df[df['Priority'].str.contains('Critical', case=False, na=False)])),
                "avg_res":  avg_res
            },
            "trend":    _build_trend(df, period),
            "recent":   recent_df.to_dict(orient='records'),
            "sync_time": datetime.now().strftime("%H:%M:%S"),
            "categories": df['Category'].dropna().unique().tolist() if 'Category' in df.columns else [],
            "sla":      sla_counts,
            "insights": {"tickets_today": tickets_today, "sla_breaches": sla_counts['breached'], "avg_res": avg_res},
            "priority_counts":  df['Priority'].value_counts().to_dict(),
            "status_counts":    df['Status'].value_counts().to_dict(),
            "category_counts":  df['Category'].value_counts().to_dict() if 'Category' in df.columns else {},
            "agent_perf":       _build_agent_perf(df)
        })
    except Exception as e:
        print(f"API Error: {e}")
        return jsonify({"error": str(e)})

@app.route('/api/all_tickets')
def all_tickets():
    if not is_admin_or_agent(): return jsonify({"data":[],"total":0,"page":1,"page_size":10,"total_pages":0})
    df = get_safe_data()
    if df.empty: return jsonify({"data":[],"total":0,"page":1,"page_size":10,"total_pages":0})

    status    = request.args.get('status')
    priority  = request.args.get('priority')
    assignee  = request.args.get('assignee')
    search    = request.args.get('search')
    date_from = request.args.get('date_from')
    date_to   = request.args.get('date_to')
    page      = int(request.args.get('page', 1))
    page_size = int(request.args.get('page_size', 10))
    no_page   = request.args.get('no_page', 'false').lower() == 'true'

    u = current_user_info()
    # Agents see all tickets (for queue management)
    if status   and status   != 'All': df = df[df['Status'].astype(str).str.contains(status,   case=False, na=False)]
    if priority and priority != 'All': df = df[df['Priority'].astype(str).str.contains(priority, case=False, na=False)]
    if assignee and assignee != 'All': df = df[df['Assigned_To'].astype(str).str.contains(assignee, case=False, na=False)]
    if search:
        mask = (df['Ticket_ID'].astype(str).str.contains(search, case=False, na=False) |
                df['Category'].astype(str).str.contains(search, case=False, na=False) |
                df.get('Description', pd.Series([''] * len(df), index=df.index)).astype(str).str.contains(search, case=False, na=False))
        df = df[mask]
    if date_from or date_to:
        df['Created_Date'] = pd.to_datetime(df['Created_Date'], errors='coerce')
        if date_from: df = df[df['Created_Date'] >= pd.to_datetime(date_from)]
        if date_to:   df = df[df['Created_Date'] <= pd.to_datetime(date_to)]

    if 'Resolution_Time_Hours' in df.columns and 'Priority' in df.columns:
        df = df.copy()
        df['SLA_Status'] = df.apply(lambda r: get_sla_status(r['Priority'], r['Resolution_Time_Hours']), axis=1)

    if no_page:
        return jsonify(df.fillna('').to_dict(orient='records'))

    total       = len(df)
    total_pages = max(1, (total + page_size - 1) // page_size)
    start       = (page - 1) * page_size
    end         = start + page_size
    page_df     = df.iloc[start:end]

    return jsonify({
        "data":        page_df.fillna('').to_dict(orient='records'),
        "total":       total,
        "page":        page,
        "page_size":   page_size,
        "total_pages": total_pages
    })

@app.route('/api/my_tickets')
def api_my_tickets():
    if 'user' not in session: return jsonify([])
    df = get_safe_data()
    if df.empty: return jsonify([])
    email = session['user']
    df    = df[df['Created_By'].astype(str).str.strip() == email]
    df    = df.fillna('')
    PROGRESS = {'Open': 10, 'In Progress': 50, 'Resolved': 80, 'Closed': 100}
    df = df.copy()
    df['Progress'] = df['Status'].apply(lambda s: PROGRESS.get(s, 10))
    return jsonify(df.to_dict(orient='records'))

@app.route('/api/submit_ticket', methods=['POST'])
def submit_ticket():
    if 'user' not in session: return jsonify({"success": False, "error": "Not logged in"})
    data     = request.get_json()
    category = data.get('Category', '').strip()
    priority = data.get('Priority', 'Low')
    desc     = data.get('Description', '').strip()
    if not category:
        return jsonify({"success": False, "error": "Category is required"})

    df  = get_safe_data()
    tid = generate_ticket_id()
    u   = current_user_info()

    # Try auto-assignment by role (agent skills matching the ticket category)
    auto_agent = _find_best_agent(category, priority)

    new_row = {
        'Ticket_ID':             tid,
        'Status':                'Open',        # stays Open until agent accepts
        'Priority':              priority,
        'Category':              category,
        'Assigned_To':           auto_agent if auto_agent else 'Unassigned',
        'Created_Date':          datetime.now().strftime('%Y-%m-%d %H:%M'),
        'Resolution_Time_Hours': '',
        'Created_By':            session['user'],
        'Description':           desc,
        'Last_Updated':          datetime.now().strftime('%Y-%m-%d %H:%M'),
        'Attachments':           ''
    }
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    df.to_excel(EXCEL_FILE, index=False)

    if auto_agent:
        log_ticket_event(tid, u['email'], u['name'], 'created',
                         f"Ticket created — Category: {category}, Priority: {priority}. "
                         f"Auto-assigned to {auto_agent} by role match.")
        # Notify assigned agent
        add_notification(auto_agent,
            f"Ticket {tid} auto-assigned to you by role: {category} ({priority}). Accept or Decline.",
            tid, 'new_ticket')
        send_email_notification(auto_agent,
            f"[Action Required] Ticket {tid} Assigned to You",
            f"<p>Ticket <b>{tid}</b> ({category}, {priority}) was auto-assigned to you based on your role.</p>"
            f"<p>Description: {desc[:200]}</p>"
            f"<p><b>Please log in and Accept or Decline from your queue.</b></p>")
    else:
        log_ticket_event(tid, u['email'], u['name'], 'created',
                         f"Ticket created — Category: {category}, Priority: {priority}. "
                         f"No matching agent role — queued for admin assignment.")
        # Notify admins only when no agent could be auto-assigned
        for email, user_data in get_users().items():
            if user_data.get('role') == 'admin':
                add_notification(email,
                    f"New ticket {tid} needs manual assignment (no matching role): {category} ({priority})",
                    tid, 'new_ticket')

    send_email_notification(
        session['user'],
        f"Ticket {tid} Submitted",
        f"<p>Your ticket <b>{tid}</b> ({category}) has been submitted successfully.</p>"
        f"<p>Priority: {priority}. An agent will be assigned shortly.</p>"
    )

    # Auto-merge if this ticket is a duplicate of an existing one
    _auto_merge_check(tid)

    return jsonify({"success": True, "ticket_id": tid})

@app.route('/api/update_ticket', methods=['POST'])
def update_ticket():
    if 'user' not in session: return jsonify({"success": False, "error": "Unauthorized"})
    u = current_user_info()
    if u['role'] not in ('admin', 'agent'):
        return jsonify({"success": False, "error": "Unauthorized"})

    data      = request.get_json()
    ticket_id = data.get('Ticket_ID', '').strip()
    if not ticket_id: return jsonify({"success": False, "error": "Ticket_ID required"})

    df   = get_safe_data()
    mask = df['Ticket_ID'].astype(str).str.strip() == ticket_id
    if not mask.any(): return jsonify({"success": False, "error": "Ticket not found"})

    # Agents can only update their assigned tickets
    if u['role'] == 'agent':
        row = df[mask].iloc[0]
        if str(row.get('Assigned_To', '')) not in (u['email'], u['name']):
            return jsonify({"success": False, "error": "Not your ticket"})

    old_row = df[mask].iloc[0].to_dict()
    changes = []
    for field in ('Status', 'Priority', 'Assigned_To'):
        if field in data:
            old_val = str(old_row.get(field, ''))
            new_val = data[field]
            if old_val != str(new_val):
                changes.append(f"{field}: {old_val} → {new_val}")
            df.loc[mask, field] = new_val

    df.loc[mask, 'Last_Updated'] = datetime.now().strftime('%Y-%m-%d %H:%M')
    if data.get('Status') in ('Resolved', 'Closed'):
        if 'Resolved_Date' not in df.columns:
            df['Resolved_Date'] = ''
        df.loc[mask, 'Resolved_Date'] = datetime.now().strftime('%Y-%m-%d %H:%M')

    df.to_excel(EXCEL_FILE, index=False)

    if changes:
        log_ticket_event(ticket_id, u['email'], u['name'], 'updated', '; '.join(changes))
        owner = str(old_row.get('Created_By', ''))
        if owner and owner != u['email']:
            add_notification(owner,
                f"Ticket {ticket_id} updated: {'; '.join(changes)}", ticket_id, 'update')
            send_email_notification(owner, f"Ticket {ticket_id} Updated",
                f"<p>Your ticket <b>{ticket_id}</b> was updated:</p><p>{'; '.join(changes)}</p>")

    return jsonify({"success": True})

@app.route('/api/bulk_update', methods=['POST'])
def bulk_update():
    if not is_admin(): return jsonify({"success": False, "error": "Unauthorized"})
    data       = request.get_json()
    ticket_ids = data.get('ticket_ids', [])
    updates    = data.get('updates', {})
    if not ticket_ids: return jsonify({"success": False, "error": "No tickets selected"})

    df   = get_safe_data()
    mask = df['Ticket_ID'].astype(str).str.strip().isin([str(t).strip() for t in ticket_ids])
    for field, value in updates.items():
        if field in df.columns:
            df.loc[mask, field] = value
    df.loc[mask, 'Last_Updated'] = datetime.now().strftime('%Y-%m-%d %H:%M')
    df.to_excel(EXCEL_FILE, index=False)
    return jsonify({"success": True, "updated": int(mask.sum())})

@app.route('/api/export')
def export():
    if not is_admin_or_agent(): return jsonify({"error": "Unauthorized"}), 403
    df = get_safe_data()
    if df.empty: return jsonify({"error": "No data"}), 400

    status   = request.args.get('status')
    priority = request.args.get('priority')
    search   = request.args.get('search')
    if status   and status   != 'All': df = df[df['Status'].astype(str).str.contains(status,   case=False, na=False)]
    if priority and priority != 'All': df = df[df['Priority'].astype(str).str.contains(priority, case=False, na=False)]
    if search:
        mask = (df['Ticket_ID'].astype(str).str.contains(search, case=False, na=False) |
                df['Category'].astype(str).str.contains(search, case=False, na=False))
        df = df[mask]

    output = io.StringIO()
    df.to_csv(output, index=False)
    output.seek(0)
    return Response(output.getvalue(), mimetype='text/csv',
                    headers={"Content-Disposition": "attachment;filename=tickets_export.csv"})

@app.route('/api/get_ticket/<ticket_id>')
def get_ticket(ticket_id):
    if 'user' not in session: return jsonify({"found": False})
    df   = get_safe_data()
    mask = df['Ticket_ID'].astype(str).str.strip() == ticket_id.strip()
    if not mask.any(): return jsonify({"found": False})
    row  = df[mask].iloc[0].fillna('').to_dict()
    u    = current_user_info()
    if u['role'] not in ('admin', 'agent') and str(row.get('Created_By', '')) != u['email']:
        return jsonify({"found": False, "error": "Unauthorized"})
    row['found'] = True
    if 'Priority' in row and 'Resolution_Time_Hours' in row:
        row['SLA_Status'] = get_sla_status(row['Priority'], row['Resolution_Time_Hours'])
    return jsonify(row)

@app.route('/api/ticket_ids')
def ticket_ids():
    if not is_admin_or_agent(): return jsonify([])
    df = get_safe_data()
    if df.empty: return jsonify([])
    return jsonify(df['Ticket_ID'].dropna().astype(str).tolist())

@app.route('/api/add_ticket', methods=['POST'])
def add_ticket():
    if not is_admin(): return jsonify({"success": False, "error": "Unauthorized"})
    data = request.get_json()
    for field in ['Ticket_ID', 'Status', 'Priority', 'Category', 'Assigned_To']:
        if not str(data.get(field, '')).strip():
            return jsonify({"success": False, "error": f"{field} is required"})

    df = get_safe_data()
    if data['Ticket_ID'].strip() in df['Ticket_ID'].astype(str).str.strip().values:
        return jsonify({"success": False, "error": "Ticket ID already exists"})

    new_row = {
        'Ticket_ID':   data['Ticket_ID'].strip(), 'Status':      data['Status'],
        'Priority':    data['Priority'],           'Category':    data['Category'].strip(),
        'Assigned_To': data['Assigned_To'].strip(),
        'Created_Date': datetime.now().strftime('%Y-%m-%d %H:%M'),
        'Resolution_Time_Hours': data.get('Resolution_Time_Hours', ''),
        'Created_By':  session.get('user', ''),
        'Description': data.get('Description', ''),
        'Last_Updated': datetime.now().strftime('%Y-%m-%d %H:%M'),
        'Attachments': ''
    }
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    df.to_excel(EXCEL_FILE, index=False)
    u = current_user_info()
    log_ticket_event(new_row['Ticket_ID'], u['email'], u['name'], 'created',
                     f"Manually added by admin. Priority: {new_row['Priority']}")
    return jsonify({"success": True})

@app.route('/api/delete_ticket', methods=['POST'])
def delete_ticket():
    if not is_admin(): return jsonify({"success": False, "error": "Unauthorized"})
    data      = request.get_json()
    ticket_id = data.get('Ticket_ID', '').strip()
    if not ticket_id: return jsonify({"success": False, "error": "Ticket_ID required"})

    df   = get_safe_data()
    mask = df['Ticket_ID'].astype(str).str.strip() == ticket_id
    if not mask.any(): return jsonify({"success": False, "error": "Ticket not found"})
    df = df[~mask]
    df.to_excel(EXCEL_FILE, index=False)
    return jsonify({"success": True})

# ── Comments API ───────────────────────────────────────────────────────────────

@app.route('/api/tickets/<ticket_id>/comments', methods=['GET'])
def get_ticket_comments(ticket_id):
    if 'user' not in session: return jsonify([])
    df   = get_safe_data()
    mask = df['Ticket_ID'].astype(str).str.strip() == ticket_id.strip()
    if not mask.any(): return jsonify([])
    u = current_user_info()
    if u['role'] not in ('admin', 'agent'):
        row = df[mask].iloc[0]
        if str(row.get('Created_By', '')) != u['email']:
            return jsonify([])
    return jsonify(get_comments(ticket_id))

@app.route('/api/tickets/<ticket_id>/comments', methods=['POST'])
def post_ticket_comment(ticket_id):
    if 'user' not in session: return jsonify({"success": False}), 401
    data = request.get_json()
    body = data.get('body', '').strip()
    if not body: return jsonify({"success": False, "error": "Comment cannot be empty"})

    u = current_user_info()
    save_comment(ticket_id, u['email'], u['name'], body)
    log_ticket_event(ticket_id, u['email'], u['name'], 'comment', f"Comment posted: {body[:80]}")

    # Notify ticket owner if commenter is admin/agent
    df   = get_safe_data()
    mask = df['Ticket_ID'].astype(str).str.strip() == ticket_id.strip()
    if mask.any():
        owner = str(df[mask].iloc[0].get('Created_By', ''))
        if owner and owner != u['email']:
            add_notification(owner,
                f"{u['name']} commented on ticket {ticket_id}", ticket_id, 'comment')
            send_email_notification(owner, f"New Comment on {ticket_id}",
                f"<p><b>{u['name']}</b> commented on your ticket <b>{ticket_id}</b>:</p><p>{body}</p>")
    return jsonify({"success": True})

# ── Timeline / History API ─────────────────────────────────────────────────────

@app.route('/api/tickets/<ticket_id>/history')
def get_ticket_history(ticket_id):
    if 'user' not in session: return jsonify([])
    data = _read_json(HISTORY_FILE, {})
    return jsonify(data.get(str(ticket_id), []))

# ── Attachments API ────────────────────────────────────────────────────────────

@app.route('/api/tickets/<ticket_id>/attachments', methods=['POST'])
def upload_attachment(ticket_id):
    if 'user' not in session: return jsonify({"success": False}), 401
    if 'file' not in request.files:
        return jsonify({"success": False, "error": "No file provided"})
    file = request.files['file']
    if file.filename == '' or not allowed_file(file.filename):
        return jsonify({"success": False, "error": "Invalid file type"})

    ext      = file.filename.rsplit('.', 1)[1].lower()
    filename = secure_filename(
        f"{ticket_id}_{datetime.now().strftime('%Y%m%d%H%M%S')}_{file.filename}"
    )
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    file.save(os.path.join(UPLOAD_FOLDER, filename))

    df   = get_safe_data()
    mask = df['Ticket_ID'].astype(str).str.strip() == ticket_id.strip()
    if mask.any():
        if 'Attachments' not in df.columns:
            df['Attachments'] = ''
        raw_val = df.loc[mask, 'Attachments'].iloc[0]
        existing = '' if pd.isna(raw_val) else str(raw_val)
        df.loc[mask, 'Attachments'] = (existing.strip(',') + ',' + filename).strip(',')
        df.to_excel(EXCEL_FILE, index=False)

    u = current_user_info()
    log_ticket_event(ticket_id, u['email'], u['name'], 'attachment', f"File uploaded: {file.filename}")
    return jsonify({"success": True, "filename": filename, "url": f"/static/uploads/{filename}"})

@app.route('/api/tickets/<ticket_id>/attachments', methods=['GET'])
def get_attachments(ticket_id):
    if 'user' not in session: return jsonify([])
    df   = get_safe_data()
    mask = df['Ticket_ID'].astype(str).str.strip() == ticket_id.strip()
    if not mask.any() or 'Attachments' not in df.columns:
        return jsonify([])
    raw_val = df[mask].iloc[0].get('Attachments', '')
    raw     = '' if pd.isna(raw_val) else str(raw_val)
    files   = [f.strip() for f in raw.split(',') if f.strip()]
    return jsonify([{"filename": f, "url": f"/static/uploads/{f}"} for f in files])

# ── Internal Notes API (agent/admin only) ─────────────────────────────────────

@app.route('/api/tickets/<ticket_id>/notes', methods=['GET'])
def get_internal_notes(ticket_id):
    if 'user' not in session:   return jsonify([])
    if not is_admin_or_agent(): return jsonify([])
    data = _read_json(NOTES_FILE, {})
    return jsonify(data.get(ticket_id, []))

@app.route('/api/tickets/<ticket_id>/notes', methods=['POST'])
def post_internal_note(ticket_id):
    if 'user' not in session:   return jsonify({"success": False}), 401
    if not is_admin_or_agent(): return jsonify({"success": False, "error": "Agents only"}), 403
    body = (request.get_json() or {}).get('body', '').strip()
    if not body:
        return jsonify({"success": False, "error": "Empty note"})
    u    = current_user_info()
    note = {
        "id":     str(uuid.uuid4())[:8],
        "author": u['name'],
        "email":  u['email'],
        "body":   body,
        "time":   datetime.now().strftime('%Y-%m-%d %H:%M')
    }
    data = _read_json(NOTES_FILE, {})
    data.setdefault(ticket_id, []).append(note)
    _write_json(NOTES_FILE, data)
    return jsonify({"success": True, "note": note})

# ── Notifications API ──────────────────────────────────────────────────────────

@app.route('/api/notifications')
def notifications():
    if 'user' not in session: return jsonify([])
    data       = _read_json(NOTIF_FILE, {})
    user_notifs = data.get(session['user'], [])
    unread     = [n for n in user_notifs if not n.get('read')]
    if unread:
        return jsonify(unread[-10:])
    # Fallback: return recent tickets as generic notifications
    df = get_safe_data()
    if df.empty: return jsonify([])
    notifs = []
    if 'Created_Date' in df.columns:
        df['Created_Date'] = pd.to_datetime(df['Created_Date'], errors='coerce')
        cutoff = datetime.now() - timedelta(days=1)
        recent = df[df['Created_Date'] >= cutoff].tail(10)
        for _, row in recent.iterrows():
            notifs.append({
                "id":       str(row.get('Ticket_ID','')),
                "message":  f"{row.get('Category','General')} — {row.get('Status','')}",
                "priority": str(row.get('Priority','')),
                "time":     str(row.get('Created_Date',''))
            })
    return jsonify(notifs[:10])

@app.route('/api/notifications/mark_read', methods=['POST'])
def mark_notifications_read():
    if 'user' not in session: return jsonify({"success": False})
    req_data  = request.get_json() or {}
    notif_id  = req_data.get('id')
    data      = _read_json(NOTIF_FILE, {})
    email     = session['user']
    if email in data:
        for n in data[email]:
            if notif_id is None or n.get('id') == notif_id:
                n['read'] = True
    _write_json(NOTIF_FILE, data)
    return jsonify({"success": True})

# ── Agent queue API ────────────────────────────────────────────────────────────

@app.route('/api/agent_queue')
def api_agent_queue():
    if 'user' not in session:    return jsonify([])
    if not is_admin_or_agent():  return jsonify([])
    df = get_safe_data()
    if df.empty: return jsonify([])
    u    = current_user_info()
    email = u['email']
    name  = u['name']
    mask  = (df['Assigned_To'].astype(str).str.strip() == email) | \
            (df['Assigned_To'].astype(str).str.strip() == name)
    df    = df[mask].fillna('')
    now   = datetime.now()
    records = []
    for _, row in df.iterrows():
        r = row.to_dict()
        priority = str(r.get('Priority', 'Low'))
        sla_hrs  = SLA_HOURS.get(priority, 24)
        status   = str(r.get('Status', ''))
        created  = row.get('Created_Date')
        if status not in ('Resolved', 'Closed') and pd.notna(created):
            try:
                created_dt = pd.to_datetime(created)
                deadline   = created_dt + timedelta(hours=sla_hrs)
                remaining_s = (deadline - now).total_seconds()
                elapsed_s   = (now - created_dt).total_seconds()
                elapsed_h   = elapsed_s / 3600
                pct         = (elapsed_h / sla_hrs) * 100
                r['SLA_Remaining_Seconds'] = int(remaining_s)
                r['SLA_Total_Seconds']     = int(sla_hrs * 3600)
                if remaining_s < 0:
                    r['SLA_Status'] = 'breached'
                elif pct >= 75:
                    r['SLA_Status'] = 'near_breach'
                else:
                    r['SLA_Status'] = 'on_track'
            except Exception:
                r['SLA_Status'] = 'unknown'
                r['SLA_Remaining_Seconds'] = None
                r['SLA_Total_Seconds']     = None
        else:
            r['SLA_Status'] = 'unknown'
            r['SLA_Remaining_Seconds'] = None
            r['SLA_Total_Seconds']     = None
        records.append(r)
    return jsonify(records)

# ── Admin user management API ──────────────────────────────────────────────────

@app.route('/api/admin/users', methods=['GET'])
def admin_get_users():
    if not is_admin(): return jsonify([])
    users = get_users()
    result = []
    for e, u in users.items():
        workload = _get_agent_workload(e) if u.get('role') == 'agent' else 0
        result.append({
            "email":               e,
            "name":                u.get("name", ""),
            "role":                u.get("role", "user"),
            "availability_status": u.get("availability_status", "online"),
            "skills":              u.get("skills", []),
            "max_workload":        u.get("max_workload", 10),
            "current_workload":    workload
        })
    return jsonify(result)

@app.route('/api/admin/users', methods=['POST'])
def admin_update_user():
    if not is_admin(): return jsonify({"success": False, "error": "Unauthorized"})
    data  = request.get_json()
    email = data.get('email', '').strip()
    users = get_users()
    if email not in users:
        return jsonify({"success": False, "error": "User not found"})
    u = users[email]
    if 'role' in data and data['role'] in ('admin', 'agent', 'user'):
        u['role'] = data['role']
    if 'name' in data and str(data['name']).strip():
        u['name'] = str(data['name']).strip()
    if 'skills' in data and isinstance(data['skills'], list):
        u['skills'] = [s for s in data['skills'] if s in list(CLASSIFY_RULES.keys()) + ['Other']]
    if 'max_workload' in data:
        try:
            u['max_workload'] = max(1, int(data['max_workload']))
        except (ValueError, TypeError):
            pass
    if 'availability_status' in data and data['availability_status'] in ('online', 'away', 'busy'):
        u['availability_status'] = data['availability_status']
    save_users(users)
    return jsonify({"success": True})

@app.route('/api/admin/users/create', methods=['POST'])
def admin_create_user():
    if not is_admin(): return jsonify({"success": False, "error": "Unauthorized"})
    data  = request.get_json()
    email = data.get('email', '').strip()
    name  = data.get('name', '').strip()
    pw    = data.get('password', '').strip()
    role  = data.get('role', 'user')
    if not name or not pw:
        return jsonify({"success": False, "error": "Name and password are required"})
    if len(pw) < 6:
        return jsonify({"success": False, "error": "Password must be at least 6 characters"})
    if role not in ('admin', 'agent', 'user'):
        role = 'user'
    # Auto-generate a unique key if no email provided
    if not email:
        email = f"user_{uuid.uuid4().hex[:8]}"
    users = get_users()
    if email in users:
        return jsonify({"success": False, "error": "Email already exists"})
    users[email] = {
        "password":            generate_password_hash(pw),
        "role":                role,
        "name":                name,
        "theme":               "dark",
        "email_notifications": True,
        "reset_token":         None,
        "reset_token_expiry":  None,
        "availability_status": "online",
        "skills":              data.get('skills', []),
        "max_workload":        int(data.get('max_workload', 10))
    }
    save_users(users)
    return jsonify({"success": True})

@app.route('/api/admin/users/delete', methods=['POST'])
def admin_delete_user():
    if not is_admin(): return jsonify({"success": False, "error": "Unauthorized"})
    data  = request.get_json()
    email = data.get('email', '').strip()
    if not email:
        return jsonify({"success": False, "error": "Email required"})
    if email == session.get('user'):
        return jsonify({"success": False, "error": "Cannot delete your own account"})
    users = get_users()
    if email not in users:
        return jsonify({"success": False, "error": "User not found"})
    del users[email]
    save_users(users)
    return jsonify({"success": True})

# ── Agent availability API ─────────────────────────────────────────────────────

@app.route('/api/agent/availability', methods=['POST'])
def set_agent_availability():
    if 'user' not in session:   return jsonify({"success": False, "error": "Not logged in"})
    if not is_admin_or_agent(): return jsonify({"success": False, "error": "Unauthorized"})
    data   = request.get_json()
    status = data.get('status', 'online')
    if status not in ('online', 'away', 'busy'):
        return jsonify({"success": False, "error": "Invalid status"})
    users = get_users()
    email = session['user']
    users[email]['availability_status'] = status
    save_users(users)
    return jsonify({"success": True, "status": status})


@app.route('/api/auto_assign/<ticket_id>', methods=['POST'])
def auto_assign_ticket(ticket_id):
    if not is_admin(): return jsonify({"success": False, "error": "Unauthorized"})
    df   = get_safe_data()
    mask = df['Ticket_ID'].astype(str).str.strip() == ticket_id.strip()
    if not mask.any():
        return jsonify({"success": False, "error": "Ticket not found"})
    row      = df[mask].iloc[0]
    category = str(row.get('Category', ''))
    priority = str(row.get('Priority', 'Low'))
    agent    = _find_best_agent(category, priority)
    if not agent:
        return jsonify({"success": False, "error": "No available agent found"})
    u = current_user_info()
    df.loc[mask, 'Assigned_To']  = agent
    df.loc[mask, 'Status']       = 'In Progress'
    df.loc[mask, 'Last_Updated'] = datetime.now().strftime('%Y-%m-%d %H:%M')
    df.to_excel(EXCEL_FILE, index=False)
    log_ticket_event(ticket_id, u['email'], u['name'], 'assigned',
                     f"Auto-assigned to {agent} (skill: {category})")
    add_notification(agent,
        f"Ticket {ticket_id} assigned to you: {category} ({priority})", ticket_id, 'new_ticket')
    return jsonify({"success": True, "assigned_to": agent})

# ── Classify API ───────────────────────────────────────────────────────────────

def _classify_keywords(text):
    tl = text.lower()
    scores = {}
    for category, keywords in CLASSIFY_RULES.items():
        score = sum(1 for kw in keywords if kw in tl)
        if score > 0:
            scores[category] = score
    best_category = max(scores, key=scores.get) if scores else 'Other'

    priority = 'Low'
    if any(w in tl for w in ['urgent','critical','asap','emergency','down','outage','breach','virus','ransomware']):
        priority = 'Critical'
    elif any(w in tl for w in ['important','slow','error','cannot','unable','broken','not working']):
        priority = 'High'
    elif any(w in tl for w in ['issue','problem','help','please','trouble','weird']):
        priority = 'Medium'
    return jsonify({"category": best_category, "priority": priority, "confidence": 0.7})

@app.route('/api/classify', methods=['POST'])
def classify():
    if 'user' not in session: return jsonify({"error": "Not logged in"}), 401
    data = request.get_json()
    text = data.get('text', '').strip()
    if not text: return jsonify({"category": None, "priority": None})
    return _classify_keywords(text)


# ── SLA Alerts API ─────────────────────────────────────────────────────────────

@app.route('/api/sla_alerts')
def sla_alerts():
    if not is_admin_or_agent(): return jsonify([])
    df  = get_safe_data()
    if df.empty: return jsonify([])
    alerts = []
    now    = datetime.now()
    if 'Created_Date' in df.columns and 'Priority' in df.columns:
        df['Created_Date'] = pd.to_datetime(df['Created_Date'], errors='coerce')
        active = df[~df['Status'].isin(['Resolved', 'Closed'])].copy()
        for _, row in active.iterrows():
            priority   = str(row.get('Priority', 'Low'))
            sla_hours  = SLA_HOURS.get(priority, 24)
            created    = row.get('Created_Date')
            if pd.isna(created): continue
            elapsed    = (now - created).total_seconds() / 3600
            remaining  = sla_hours - elapsed
            pct        = (elapsed / sla_hours) * 100
            if pct >= 75:
                alerts.append({
                    "ticket_id":     str(row.get('Ticket_ID','')),
                    "priority":      priority,
                    "category":      str(row.get('Category','')),
                    "assigned_to":   str(row.get('Assigned_To','Unassigned')),
                    "elapsed_hours": round(elapsed, 1),
                    "sla_hours":     sla_hours,
                    "remaining_hours": round(remaining, 1),
                    "sla_pct":       round(pct, 1),
                    "status":        'breached' if remaining < 0 else 'near_breach'
                })
    alerts.sort(key=lambda x: x['sla_pct'], reverse=True)
    return jsonify(alerts[:20])

def _build_agent_perf(df):
    """Build agent performance summary — used by /api/stats."""
    if df.empty or 'Assigned_To' not in df.columns:
        return []
    results = []
    for agent, group in df.groupby('Assigned_To'):
        total     = len(group)
        resolved  = len(group[group['Status'].isin(['Resolved','Closed'])])
        res_times = group['Resolution_Time_Hours'].dropna() if 'Resolution_Time_Hours' in group else pd.Series([], dtype=float)
        avg_res   = round(float(res_times.mean()), 1) if len(res_times) > 0 else 0
        breaches  = sum(1 for _, r in group.iterrows()
                        if get_sla_status(r.get('Priority'), r.get('Resolution_Time_Hours')) == 'breached')
        results.append({
            "agent":           str(agent),
            "total":           total,
            "resolved":        resolved,
            "avg_res":         avg_res if not pd.isna(avg_res) else 0,
            "resolution_rate": round((resolved/total)*100, 1) if total > 0 else 0,
            "sla_breaches":    breaches
        })
    return sorted(results, key=lambda x: x['total'], reverse=True)

# ── Search autocomplete API ────────────────────────────────────────────────────

@app.route('/api/search')
def search_autocomplete():
    if 'user' not in session: return jsonify([])
    query = request.args.get('q', '').strip()
    if len(query) < 2: return jsonify([])

    df = get_safe_data()
    u  = current_user_info()
    if u['role'] == 'user':
        df = df[df['Created_By'].astype(str) == u['email']]

    desc_col = df['Description'].astype(str) if 'Description' in df.columns else pd.Series([''] * len(df), index=df.index)
    mask = (
        df['Ticket_ID'].astype(str).str.contains(query, case=False, na=False) |
        df['Category'].astype(str).str.contains(query, case=False, na=False) |
        desc_col.str.contains(query, case=False, na=False)
    )
    results = df[mask].head(8).fillna('')
    return jsonify([{
        "id":    str(r['Ticket_ID']),
        "label": f"{r['Ticket_ID']} — {r['Category']} ({r['Status']})",
        "url":   f"/tickets/{r['Ticket_ID']}"
    } for _, r in results.iterrows()])

# ── Accept / Decline with re-routing ──────────────────────────────────────────

@app.route('/api/tickets/<ticket_id>/accept', methods=['POST'])
def accept_ticket(ticket_id):
    if 'user' not in session:   return jsonify({"success": False, "error": "Unauthorized"}), 401
    if not is_admin_or_agent(): return jsonify({"success": False, "error": "Unauthorized"}), 403
    u    = current_user_info()
    df   = get_safe_data()
    mask = df['Ticket_ID'].astype(str).str.strip() == ticket_id.strip()
    if not mask.any():
        return jsonify({"success": False, "error": "Ticket not found"})
    row = df[mask].iloc[0]
    if str(row.get('Assigned_To', '')) not in (u['email'], u['name']) and u['role'] != 'admin':
        return jsonify({"success": False, "error": "Not assigned to you"})
    df.loc[mask, 'Status']       = 'In Progress'
    df.loc[mask, 'Last_Updated'] = datetime.now().strftime('%Y-%m-%d %H:%M')
    df.to_excel(EXCEL_FILE, index=False)
    log_ticket_event(ticket_id, u['email'], u['name'], 'updated', f"Ticket accepted by {u['name']}")
    owner = str(row.get('Created_By', ''))
    if owner and owner != u['email']:
        add_notification(owner, f"{u['name']} accepted ticket {ticket_id}", ticket_id, 'update')
    return jsonify({"success": True})


@app.route('/api/tickets/<ticket_id>/decline', methods=['POST'])
def decline_ticket(ticket_id):
    if 'user' not in session:   return jsonify({"success": False, "error": "Unauthorized"}), 401
    if not is_admin_or_agent(): return jsonify({"success": False, "error": "Unauthorized"}), 403
    u      = current_user_info()
    reason = (request.get_json() or {}).get('reason', '').strip() or 'No reason given'
    df     = get_safe_data()
    mask   = df['Ticket_ID'].astype(str).str.strip() == ticket_id.strip()
    if not mask.any():
        return jsonify({"success": False, "error": "Ticket not found"})
    row      = df[mask].iloc[0]
    if str(row.get('Assigned_To', '')) not in (u['email'], u['name']) and u['role'] != 'admin':
        return jsonify({"success": False, "error": "Not assigned to you"})
    category = str(row.get('Category', ''))
    priority = str(row.get('Priority', 'Low'))

    # Re-route: find next best agent excluding the declining agent
    users = get_users()
    available = [
        (email, usr) for email, usr in users.items()
        if usr.get('role') == 'agent'
        and usr.get('availability_status', 'online') == 'online'
        and email != u['email']
    ]
    candidates = []
    for email, usr in available:
        workload  = _get_agent_workload(email)
        max_wl    = usr.get('max_workload', 10)
        if workload >= max_wl: continue
        has_skill = category in usr.get('skills', [])
        candidates.append((email, workload, has_skill))
    candidates.sort(key=lambda x: (not x[2], x[1]))
    new_agent = candidates[0][0] if candidates else 'Unassigned'

    df.loc[mask, 'Assigned_To']  = new_agent
    df.loc[mask, 'Status']       = 'Open' if new_agent == 'Unassigned' else 'In Progress'
    df.loc[mask, 'Last_Updated'] = datetime.now().strftime('%Y-%m-%d %H:%M')
    df.to_excel(EXCEL_FILE, index=False)

    log_ticket_event(ticket_id, u['email'], u['name'], 'assigned',
                     f"Declined by {u['name']} ({reason}). Re-routed to {new_agent}")
    if new_agent != 'Unassigned':
        add_notification(new_agent,
            f"Ticket {ticket_id} re-assigned to you: {category} ({priority})", ticket_id, 'new_ticket')
    return jsonify({"success": True, "re_assigned_to": new_agent})


# ── Canned responses (quick reply templates) ───────────────────────────────────

def _get_canned():
    data = _read_json(CANNED_FILE, [])
    if not data:
        _write_json(CANNED_FILE, DEFAULT_CANNED)
        return DEFAULT_CANNED
    return data

@app.route('/api/canned_responses', methods=['GET'])
def get_canned_responses():
    if not is_admin_or_agent(): return jsonify([])
    return jsonify(_get_canned())

@app.route('/api/canned_responses', methods=['POST'])
def add_canned_response():
    if not is_admin(): return jsonify({"success": False, "error": "Admins only"}), 403
    body  = request.get_json() or {}
    label = body.get('label', '').strip()
    text  = body.get('body', '').strip()
    if not label or not text:
        return jsonify({"success": False, "error": "label and body required"})
    items = _get_canned()
    new   = {"id": str(uuid.uuid4())[:8], "label": label, "body": text}
    items.append(new)
    _write_json(CANNED_FILE, items)
    return jsonify({"success": True, "item": new})

@app.route('/api/canned_responses/<cid>', methods=['DELETE'])
def delete_canned_response(cid):
    if not is_admin(): return jsonify({"success": False, "error": "Admins only"}), 403
    items = [r for r in _get_canned() if r.get('id') != cid]
    _write_json(CANNED_FILE, items)
    return jsonify({"success": True})


# ── Ticket transfer ────────────────────────────────────────────────────────────

@app.route('/api/tickets/<ticket_id>/transfer', methods=['POST'])
def transfer_ticket(ticket_id):
    if 'user' not in session:   return jsonify({"success": False, "error": "Unauthorized"}), 401
    if not is_admin_or_agent(): return jsonify({"success": False, "error": "Unauthorized"}), 403
    u        = current_user_info()
    body     = request.get_json() or {}
    to_agent = body.get('to_agent', '').strip()
    reason   = body.get('reason', '').strip() or 'No reason given'
    if not to_agent:
        return jsonify({"success": False, "error": "to_agent required"})
    users = get_users()
    if to_agent not in users:
        return jsonify({"success": False, "error": "Agent not found"})
    df   = get_safe_data()
    mask = df['Ticket_ID'].astype(str).str.strip() == ticket_id.strip()
    if not mask.any():
        return jsonify({"success": False, "error": "Ticket not found"})
    row = df[mask].iloc[0]
    if u['role'] == 'agent' and str(row.get('Assigned_To', '')) not in (u['email'], u['name']):
        return jsonify({"success": False, "error": "Not your ticket"})
    old_agent = str(row.get('Assigned_To', 'Unassigned'))
    df.loc[mask, 'Assigned_To']  = to_agent
    df.loc[mask, 'Status']       = 'In Progress'
    df.loc[mask, 'Last_Updated'] = datetime.now().strftime('%Y-%m-%d %H:%M')
    df.to_excel(EXCEL_FILE, index=False)
    log_ticket_event(ticket_id, u['email'], u['name'], 'assigned',
                     f"Transferred from {old_agent} to {to_agent}. Reason: {reason}")
    add_notification(to_agent,
        f"Ticket {ticket_id} transferred to you by {u['name']}", ticket_id, 'new_ticket')
    return jsonify({"success": True, "assigned_to": to_agent})


# ── Agent personal stats ──────────────────────────────────────────────────────

@app.route('/api/agent/stats')
def agent_stats():
    if not is_admin_or_agent(): return jsonify({}), 403
    u   = current_user_info()
    df  = get_safe_data()
    now = datetime.now()

    if df.empty:
        return jsonify({
            "resolved_today": 0, "resolved_week": 0, "avg_resolution": 0,
            "sla_compliance": 100, "queue_open": 0, "queue_inprogress": 0,
            "queue_near_breach": 0, "queue_breached": 0,
            "total_assigned": 0, "total_resolved": 0,
            "trend": {"labels": [], "values": []}
        })

    mask  = (df['Assigned_To'].astype(str).str.strip() == u['email']) | \
            (df['Assigned_To'].astype(str).str.strip() == u['name'])
    my_df = df[mask].copy()

    today      = now.date()
    week_start = today - timedelta(days=today.weekday())

    resolved = my_df[my_df['Status'].isin(['Resolved', 'Closed'])].copy()
    resolved_today = 0
    resolved_week  = 0
    if 'Last_Updated' in resolved.columns and len(resolved):
        resolved['_upd'] = pd.to_datetime(resolved['Last_Updated'], errors='coerce')
        resolved_today   = int((resolved['_upd'].dt.date == today).sum())
        resolved_week    = int((resolved['_upd'].dt.date >= week_start).sum())

    avg_res = 0
    if 'Resolution_Time_Hours' in resolved.columns and len(resolved):
        vals    = pd.to_numeric(resolved['Resolution_Time_Hours'], errors='coerce').dropna()
        avg_res = round(float(vals.mean()), 1) if len(vals) else 0

    total_resolved = len(resolved)
    sla_ok = sum(
        1 for _, row in resolved.iterrows()
        if get_sla_status(row.get('Priority', 'Low'), row.get('Resolution_Time_Hours', 0)) != 'breached'
    )
    sla_pct = round(sla_ok / total_resolved * 100, 1) if total_resolved else 100.0

    active       = my_df[~my_df['Status'].isin(['Resolved', 'Closed'])]
    queue_open   = int((active['Status'] == 'Open').sum())
    queue_inprog = int((active['Status'] == 'In Progress').sum())
    near_breach  = 0
    sla_breached = 0
    for _, row in active.iterrows():
        created  = row.get('Created_Date')
        sla_hrs  = SLA_HOURS.get(str(row.get('Priority', 'Low')), 24)
        if pd.notna(created):
            try:
                elapsed = (now - pd.to_datetime(created)).total_seconds() / 3600
                pct     = (elapsed / sla_hrs) * 100
                if pct >= 100: sla_breached += 1
                elif pct >= 75: near_breach  += 1
            except Exception:
                pass

    # 7-day ticket trend
    trend_labels, trend_values = [], []
    if 'Created_Date' in my_df.columns:
        my_df['_date'] = pd.to_datetime(my_df['Created_Date'], errors='coerce').dt.date
        for i in range(6, -1, -1):
            d = (now - timedelta(days=i)).date()
            trend_labels.append(d.strftime('%a'))
            trend_values.append(int((my_df['_date'] == d).sum()))

    return jsonify({
        "resolved_today":    resolved_today,
        "resolved_week":     resolved_week,
        "avg_resolution":    avg_res,
        "sla_compliance":    sla_pct,
        "queue_open":        queue_open,
        "queue_inprogress":  queue_inprog,
        "queue_near_breach": near_breach,
        "queue_breached":    sla_breached,
        "total_assigned":    len(my_df),
        "total_resolved":    total_resolved,
        "trend":             {"labels": trend_labels, "values": trend_values}
    })


# ── Agent-to-agent chat ────────────────────────────────────────────────────────

@app.route('/api/chat/messages', methods=['GET'])
def get_chat_messages():
    if not is_admin_or_agent(): return jsonify([])
    u     = current_user_info()
    room  = request.args.get('room', 'general')
    since = request.args.get('since', '')
    msgs  = _read_json(CHAT_FILE, [])

    if room == 'general':
        visible = [m for m in msgs if m.get('to') == 'all']
    else:
        visible = [m for m in msgs if
                   (m.get('from_email') == u['email'] and m.get('to') == room) or
                   (m.get('from_email') == room       and m.get('to') == u['email'])]

    if since:
        try:
            since_dt = datetime.fromisoformat(since)
            visible  = [m for m in visible if datetime.fromisoformat(m['timestamp']) > since_dt]
        except Exception:
            pass

    # Mark visible messages as read
    changed = False
    ids_visible = {m['id'] for m in visible}
    for m in msgs:
        if m['id'] in ids_visible and u['email'] not in m.get('read_by', []):
            m.setdefault('read_by', []).append(u['email'])
            changed = True
    if changed:
        _write_json(CHAT_FILE, msgs)

    return jsonify(visible[-100:])


@app.route('/api/chat/messages', methods=['POST'])
def post_chat_message():
    if not is_admin_or_agent(): return jsonify({"success": False}), 403
    u    = current_user_info()
    data = request.get_json() or {}
    to   = data.get('to', 'all').strip()
    body = data.get('body', '').strip()
    if not body:
        return jsonify({"success": False, "error": "Empty message"})
    msg  = {
        "id":         str(uuid.uuid4())[:8],
        "from_email": u['email'],
        "from_name":  u['name'],
        "to":         to,
        "body":       body,
        "timestamp":  datetime.now().isoformat(),
        "read_by":    [u['email']]
    }
    msgs = _read_json(CHAT_FILE, [])
    msgs.append(msg)
    if len(msgs) > 1000:
        msgs = msgs[-1000:]
    _write_json(CHAT_FILE, msgs)
    return jsonify({"success": True, "message": msg})


@app.route('/api/chat/unread')
def chat_unread():
    if not is_admin_or_agent(): return jsonify({"count": 0})
    u    = current_user_info()
    msgs = _read_json(CHAT_FILE, [])
    count = sum(
        1 for m in msgs
        if u['email'] not in m.get('read_by', [])
        and m.get('from_email') != u['email']
        and (m.get('to') == 'all' or m.get('to') == u['email'])
    )
    return jsonify({"count": count})


# ── Unassigned ticket count (for nav badge) ───────────────────────────────────

@app.route('/api/unassigned_count')
def unassigned_count():
    if not is_admin(): return jsonify({"count": 0})
    df = get_safe_data()
    if df.empty: return jsonify({"count": 0})
    unassigned_mask   = df['Assigned_To'].astype(str).str.strip().isin(['Unassigned', '', 'nan'])
    pending_mask      = (df['Status'].astype(str).str.strip() == 'Open') & ~unassigned_mask
    return jsonify({"count": int(unassigned_mask.sum() + pending_mask.sum())})


# ── Assignment workflow ────────────────────────────────────────────────────────

@app.route('/assign')
def assign_tickets():
    if 'user' not in session: return redirect(url_for('login'))
    if not is_admin():        return redirect(url_for('my_tickets'))
    return render_template('assign_tickets.html')


@app.route('/api/assignment_queue')
def assignment_queue():
    if not is_admin(): return jsonify({"error": "Unauthorized"}), 403
    df    = get_safe_data()
    users = get_users()
    now   = datetime.now()

    # Only truly unassigned tickets need admin assignment.
    # Once Assigned_To is set (even if still Open), the ticket is in the agent's
    # pending-acceptance queue — remove it from this panel immediately.
    unassigned = []
    if not df.empty:
        mask = (
            df['Assigned_To'].astype(str).str.strip().isin(['Unassigned', '', 'nan'])
        )
        q_df = df[mask].fillna('').copy()
        for _, row in q_df.iterrows():
            r        = row.to_dict()
            priority = str(r.get('Priority', 'Low'))
            sla_hrs  = SLA_HOURS.get(priority, 24)
            created  = row.get('Created_Date')
            if pd.notna(created) and str(created).strip():
                try:
                    created_dt   = pd.to_datetime(created)
                    elapsed_s    = (now - created_dt).total_seconds()
                    remaining_s  = (sla_hrs * 3600) - elapsed_s
                    pct          = (elapsed_s / (sla_hrs * 3600)) * 100
                    r['SLA_Remaining_Seconds'] = int(remaining_s)
                    r['SLA_Pct']               = round(pct, 1)
                    if remaining_s < 0:
                        r['SLA_Status'] = 'breached'
                    elif pct >= 75:
                        r['SLA_Status'] = 'near_breach'
                    else:
                        r['SLA_Status'] = 'on_track'
                except Exception:
                    r['SLA_Remaining_Seconds'] = None
                    r['SLA_Pct']               = 0
                    r['SLA_Status']            = 'unknown'
            else:
                r['SLA_Remaining_Seconds'] = None
                r['SLA_Pct']               = 0
                r['SLA_Status']            = 'unknown'
            unassigned.append(r)

    # Sort: Critical first, then by SLA_Pct descending
    priority_order = {'Critical': 0, 'High': 1, 'Medium': 2, 'Low': 3}
    unassigned.sort(key=lambda x: (priority_order.get(x.get('Priority', 'Low'), 3), -(x.get('SLA_Pct') or 0)))

    # All agents with live workload
    agents = []
    for email, u in users.items():
        if u.get('role') != 'agent':
            continue
        workload = _get_agent_workload(email)
        max_wl   = u.get('max_workload', 10)
        agents.append({
            "email":            email,
            "name":             u.get('name', email),
            "availability":     u.get('availability_status', 'online'),
            "skills":           u.get('skills', []),
            "current_workload": workload,
            "max_workload":     max_wl,
            "capacity_pct":     round((workload / max_wl) * 100, 1) if max_wl > 0 else 100
        })
    agents.sort(key=lambda x: (x['availability'] != 'online', x['capacity_pct']))

    # Pending acceptance: Open tickets that already have an agent assigned
    pending_acceptance = 0
    if not df.empty:
        assigned_mask = ~df['Assigned_To'].astype(str).str.strip().isin(['Unassigned', '', 'nan'])
        open_mask     = df['Status'].astype(str).str.strip() == 'Open'
        pending_acceptance = int((assigned_mask & open_mask).sum())

    return jsonify({
        "unassigned_tickets": unassigned,
        "agents":             agents,
        "summary": {
            "total_unassigned":   len(unassigned),
            "pending_acceptance": pending_acceptance,
            "agents_online":      sum(1 for a in agents if a['availability'] == 'online'),
            "agents_busy":        sum(1 for a in agents if a['availability'] == 'busy'),
            "agents_away":        sum(1 for a in agents if a['availability'] == 'away'),
        }
    })


@app.route('/api/tickets/<ticket_id>/assign', methods=['POST'])
def assign_ticket(ticket_id):
    if not is_admin(): return jsonify({"success": False, "error": "Unauthorized"}), 403
    body        = request.get_json() or {}
    agent_email = body.get('agent_email', '').strip()
    if not agent_email:
        return jsonify({"success": False, "error": "agent_email required"})
    users = get_users()
    if agent_email not in users or users[agent_email].get('role') not in ('agent', 'admin'):
        return jsonify({"success": False, "error": "Agent not found"})

    df   = get_safe_data()
    mask = df['Ticket_ID'].astype(str).str.strip() == ticket_id.strip()
    if not mask.any():
        return jsonify({"success": False, "error": "Ticket not found"})

    old_row   = df[mask].iloc[0].to_dict()
    old_agent = str(old_row.get('Assigned_To', 'Unassigned'))

    # Keep status as 'Open' so agent must explicitly Accept or Decline
    current_status = str(old_row.get('Status', 'Open'))
    new_status = current_status if current_status not in ('Resolved', 'Closed') else 'Open'
    if new_status not in ('Open', 'In Progress'):
        new_status = 'Open'

    df.loc[mask, 'Assigned_To']  = agent_email
    df.loc[mask, 'Status']       = new_status
    df.loc[mask, 'Last_Updated'] = datetime.now().strftime('%Y-%m-%d %H:%M')
    df.to_excel(EXCEL_FILE, index=False)

    u = current_user_info()
    log_ticket_event(ticket_id, u['email'], u['name'], 'assigned',
                     f"Manually assigned to {agent_email} by admin (was: {old_agent}). Awaiting agent acceptance.")

    category   = str(old_row.get('Category', ''))
    priority   = str(old_row.get('Priority', 'Low'))
    agent_name = users[agent_email].get('name', agent_email)

    add_notification(agent_email,
        f"Ticket {ticket_id} assigned to you by admin — please Accept or Decline: {category} ({priority})",
        ticket_id, 'new_ticket')
    send_email_notification(agent_email,
        f"[Action Required] Ticket {ticket_id} Assigned to You",
        f"<p>Admin <b>{u['name']}</b> assigned ticket <b>{ticket_id}</b> to you.</p>"
        f"<p>Category: {category} | Priority: {priority}</p>"
        f"<p>Description: {str(old_row.get('Description',''))[:200]}</p>"
        f"<p><b>Please log in and Accept or Decline this ticket from your queue.</b></p>")

    # Notify ticket creator
    creator = str(old_row.get('Created_By', ''))
    if creator and creator != agent_email:
        add_notification(creator,
            f"Your ticket {ticket_id} has been assigned to {agent_name} (pending acceptance)",
            ticket_id, 'update')

    return jsonify({"success": True, "assigned_to": agent_email, "agent_name": agent_name})


# ── Agent list for transfer modal ──────────────────────────────────────────────

@app.route('/api/agents/available')
def list_agents_available():
    """All agents with workload + role info — used by the agent overview tab."""
    if not is_admin(): return jsonify([])
    users = get_users()
    result = []
    for email, usr in users.items():
        if usr.get('role') != 'agent':
            continue
        workload = _get_agent_workload(email)
        result.append({
            "email":               email,
            "name":                usr.get('name', email),
            "availability_status": usr.get('availability_status', 'online'),
            "skills":              usr.get('skills', []),
            "max_workload":        usr.get('max_workload', 10),
            "current_workload":    workload,
        })
    result.sort(key=lambda x: (x['availability_status'] != 'online', x['current_workload']))
    return jsonify(result)


@app.route('/api/agents')
def list_agents():
    if not is_admin_or_agent(): return jsonify([])
    users = get_users()
    u     = current_user_info()
    return jsonify([
        {"email": email, "name": usr.get("name", email),
         "availability": usr.get("availability_status", "online"),
         "workload": _get_agent_workload(email)}
        for email, usr in users.items()
        if usr.get('role') in ('agent', 'admin') and email != u['email']
    ])


# ── Duplicate detection & auto-merge (backend only) ───────────────────────────

def _jaccard_similarity(text_a, text_b):
    """Word-level Jaccard similarity between two strings (0.0 – 1.0)."""
    stop = {'the','a','an','is','it','in','on','at','to','of','and','or','for',
            'my','i','me','we','our','your','can','not','with','this','that',
            'are','was','have','has','be','been','do','did','please','help'}
    def tokens(t):
        return {w for w in re.sub(r'[^a-z0-9 ]', ' ', t.lower()).split() if w and w not in stop}
    a, b = tokens(text_a), tokens(text_b)
    if not a or not b:
        return 0.0
    return len(a & b) / len(a | b)


def _perform_merge(df, master_id, duplicate_ids, actor_name, actor_email):
    """Core merge logic (no HTTP context required).

    Closes each duplicate, copies its comments/history into the master,
    appends its description, and notifies all affected users.
    Returns the updated DataFrame (caller must save to Excel).
    """
    now_str       = datetime.now().strftime('%Y-%m-%d %H:%M')
    comments_data = _read_json(COMMENTS_FILE, {})
    history_data  = _read_json(HISTORY_FILE,  {})
    extra_descs   = []
    notified      = set()

    for dup_id in duplicate_ids:
        dup_mask = df['Ticket_ID'].astype(str).str.strip() == dup_id
        if not dup_mask.any():
            continue
        dup_row  = df[dup_mask].iloc[0].to_dict()

        # Collect description snippet
        dup_desc = str(dup_row.get('Description', '')).strip()
        if dup_desc:
            extra_descs.append(f"[From {dup_id}] {dup_desc}")

        # Copy comments → master
        for c in comments_data.get(dup_id, []):
            comments_data.setdefault(master_id, []).append({
                **c,
                "id":   str(uuid.uuid4()),
                "body": f"[Merged from {dup_id}] {c['body']}",
                "time": c.get('time', now_str)
            })

        # Copy history → master
        for h in history_data.get(dup_id, []):
            history_data.setdefault(master_id, []).append({
                **h,
                "detail":    f"[Merged from {dup_id}] {h.get('detail','')}",
                "timestamp": h.get('timestamp', now_str)
            })

        # Close the duplicate
        df.loc[dup_mask, 'Status']       = 'Closed'
        df.loc[dup_mask, 'Last_Updated'] = now_str
        if 'Resolved_Date' not in df.columns:
            df['Resolved_Date'] = ''
        df.loc[dup_mask, 'Resolved_Date'] = now_str

        # Leave a note on the duplicate
        comments_data.setdefault(dup_id, []).append({
            "id":     str(uuid.uuid4())[:8],
            "author": actor_name,
            "email":  actor_email,
            "body":   f"This ticket was automatically merged into {master_id} as a duplicate.",
            "time":   now_str
        })
        log_ticket_event(dup_id, actor_email, actor_name, 'merged',
                         f"Auto-merged into {master_id} (duplicate detected)")

        # Notify duplicate creator
        creator = str(dup_row.get('Created_By', ''))
        if creator and creator not in notified:
            notified.add(creator)
            add_notification(creator,
                f"Your ticket {dup_id} was automatically merged into {master_id} as a duplicate.",
                master_id, 'update')

    # Update master description
    master_mask = df['Ticket_ID'].astype(str).str.strip() == master_id
    master_row  = df[master_mask].iloc[0].to_dict()
    if extra_descs:
        base = str(master_row.get('Description', '')).strip()
        df.loc[master_mask, 'Description'] = (
            base + '\n\n--- Auto-merged duplicates ---\n' + '\n'.join(extra_descs)
        )
    df.loc[master_mask, 'Last_Updated'] = now_str

    _write_json(COMMENTS_FILE, comments_data)
    _write_json(HISTORY_FILE,  history_data)

    log_ticket_event(master_id, actor_email, actor_name, 'merged',
                     f"Auto-merged duplicates: {', '.join(duplicate_ids)}")

    # Notify master creator
    master_creator = str(master_row.get('Created_By', ''))
    if master_creator:
        add_notification(master_creator,
            f"Ticket {master_id} absorbed {len(duplicate_ids)} duplicate ticket(s) automatically.",
            master_id, 'update')

    # Notify all admins
    for email, u in get_users().items():
        if u.get('role') == 'admin':
            add_notification(email,
                f"Auto-merged {len(duplicate_ids)} duplicate(s) into {master_id}",
                master_id, 'info')

    return df


def _auto_merge_check(new_ticket_id):
    """Called after a new ticket is saved.

    Compares the new ticket against all existing open tickets.
    If a duplicate is found (same category + >= 50 % word overlap),
    the new ticket is merged into the older master automatically.
    """
    AUTO_MERGE_THRESHOLD = 0.50   # 50 % Jaccard similarity required

    try:
        df = get_safe_data()
        if df.empty:
            return

        new_mask = df['Ticket_ID'].astype(str).str.strip() == new_ticket_id
        if not new_mask.any():
            return
        new_row  = df[new_mask].iloc[0].to_dict()
        new_desc = str(new_row.get('Description', '')).strip()
        new_cat  = str(new_row.get('Category', '')).strip()

        # Compare against all other active (non-closed/resolved) tickets
        active = df[
            ~df['Ticket_ID'].astype(str).str.strip().eq(new_ticket_id) &
            ~df['Status'].astype(str).str.strip().isin(['Closed', 'Resolved'])
        ].fillna('')

        best_match_id  = None
        best_sim       = 0.0

        for _, row in active.iterrows():
            if str(row.get('Category', '')).strip() != new_cat:
                continue   # must share same category
            sim = _jaccard_similarity(new_desc, str(row.get('Description', '')))
            if sim >= AUTO_MERGE_THRESHOLD and sim > best_sim:
                best_sim       = sim
                best_match_id  = str(row['Ticket_ID'])

        if best_match_id:
            # The older ticket is the master; the new one is the duplicate
            df = _perform_merge(df, best_match_id, [new_ticket_id],
                                actor_name='System', actor_email='system@auto-merge')
            df.to_excel(EXCEL_FILE, index=False)
            print(f"[Auto-Merge] {new_ticket_id} merged into {best_match_id} "
                  f"(similarity {best_sim:.0%})")
    except Exception as e:
        print(f"[Auto-Merge Error] {e}")


if __name__ == '__main__':
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    app.run(debug=True, port=5000)
