from flask import Flask, render_template, jsonify, request, redirect, url_for, session, Response
import pandas as pd
import os, io, json, random, re, sqlite3, secrets, uuid
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
        with open(path) as f:
            return json.load(f)
    except Exception:
        return default

def _write_json(path, data):
    with open(path, 'w') as f:
        json.dump(data, f, indent=2)

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
        print(f"[Email skipped - mail not configured]: {subject} → {to_email}")
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

@app.route('/analytics')
def analytics():
    if 'user' not in session: return redirect(url_for('login'))
    if not is_admin():        return redirect(url_for('my_tickets'))
    return render_template('analytics.html')

@app.route('/sql')
def sql_console():
    if 'user' not in session: return redirect(url_for('login'))
    if not is_admin():        return redirect(url_for('my_tickets'))
    return render_template('sql_query.html')

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
            "category_counts":  df['Category'].value_counts().to_dict() if 'Category' in df.columns else {}
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

    new_row = {
        'Ticket_ID':             tid,
        'Status':                'Open',
        'Priority':              priority,
        'Category':              category,
        'Assigned_To':           'Unassigned',
        'Created_Date':          datetime.now().strftime('%Y-%m-%d %H:%M'),
        'Resolution_Time_Hours': '',
        'Created_By':            session['user'],
        'Description':           desc,
        'Last_Updated':          datetime.now().strftime('%Y-%m-%d %H:%M'),
        'Attachments':           ''
    }
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    df.to_excel(EXCEL_FILE, index=False)

    log_ticket_event(tid, u['email'], u['name'], 'created',
                     f"Ticket created — Category: {category}, Priority: {priority}")

    # Notify all admins
    for email, user_data in get_users().items():
        if user_data.get('role') == 'admin':
            add_notification(email,
                f"New ticket {tid}: {category} ({priority})", tid, 'new_ticket')

    send_email_notification(
        session['user'],
        f"Ticket {tid} Submitted",
        f"<p>Your ticket <b>{tid}</b> ({category}) has been submitted successfully.</p>"
        f"<p>Priority: {priority}</p>"
    )
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
        existing = str(df.loc[mask, 'Attachments'].iloc[0])
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
    raw   = str(df[mask].iloc[0].get('Attachments', ''))
    files = [f.strip() for f in raw.split(',') if f.strip()]
    return jsonify([{"filename": f, "url": f"/static/uploads/{f}"} for f in files])

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
    if 'Resolution_Time_Hours' in df.columns and 'Priority' in df.columns:
        df = df.copy()
        df['SLA_Status'] = df.apply(
            lambda r: get_sla_status(r['Priority'], r['Resolution_Time_Hours']), axis=1)
    return jsonify(df.to_dict(orient='records'))

# ── Admin user management API ──────────────────────────────────────────────────

@app.route('/api/admin/users', methods=['GET'])
def admin_get_users():
    if not is_admin(): return jsonify([])
    users = get_users()
    return jsonify([
        {"email": e, "name": u.get("name",""), "role": u.get("role","user")}
        for e, u in users.items()
    ])

@app.route('/api/admin/users', methods=['POST'])
def admin_update_user():
    if not is_admin(): return jsonify({"success": False})
    data  = request.get_json()
    email = data.get('email', '')
    users = get_users()
    if email not in users:
        return jsonify({"success": False, "error": "User not found"})
    if 'role' in data and data['role'] in ('admin', 'agent', 'user'):
        users[email]['role'] = data['role']
        save_users(users)
    return jsonify({"success": True})

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

def _classify_ai(text):
    try:
        import anthropic as anthropic_sdk
        client     = anthropic_sdk.Anthropic(api_key=os.environ['ANTHROPIC_API_KEY'])
        categories = list(CLASSIFY_RULES.keys()) + ['Other']
        message    = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=256,
            messages=[{
                "role": "user",
                "content": (
                    f"You are an IT helpdesk classifier. Given this ticket, return JSON with:\n"
                    f"- \"category\": one of {categories}\n"
                    f"- \"priority\": one of Critical, High, Medium, Low\n"
                    f"- \"confidence\": 0.0-1.0\n\n"
                    f"Ticket: \"{text}\"\n\nReturn only valid JSON."
                )
            }]
        )
        result = json.loads(message.content[0].text)
        return jsonify(result)
    except Exception as e:
        print(f"[AI classify error]: {e}")
        return _classify_keywords(text)

@app.route('/api/classify', methods=['POST'])
def classify():
    if 'user' not in session: return jsonify({"error": "Not logged in"}), 401
    data   = request.get_json()
    text   = data.get('text', '').strip()
    use_ai = data.get('use_ai', False)
    if not text: return jsonify({"category": None, "priority": None})
    if use_ai and os.environ.get('ANTHROPIC_API_KEY'):
        return _classify_ai(text)
    return _classify_keywords(text)

# ── SQL Console API ────────────────────────────────────────────────────────────

@app.route('/api/sql', methods=['POST'])
def run_sql():
    if not is_admin(): return jsonify({"error": "Unauthorized"}), 403
    data  = request.get_json()
    query = data.get('query', '').strip()
    if not re.match(r'^\s*SELECT\b', query, re.IGNORECASE):
        return jsonify({"error": "Only SELECT queries are allowed."})

    conn = sqlite3.connect(':memory:')
    try:
        df = get_safe_data()
        df.to_sql('tickets', conn, if_exists='replace', index=False)
        users      = get_users()
        users_rows = [{"email": e, "name": u.get("name",""), "role": u.get("role","user")}
                      for e, u in users.items()]
        pd.DataFrame(users_rows).to_sql('users', conn, if_exists='replace', index=False)

        cursor = conn.cursor()
        cursor.execute(query)
        columns = [d[0] for d in cursor.description] if cursor.description else []
        rows    = [[str(v) if v is not None else '' for v in row] for row in cursor.fetchall()]
        return jsonify({"columns": columns, "rows": rows, "count": len(rows)})
    except Exception as e:
        return jsonify({"error": str(e)})
    finally:
        conn.close()

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

# ── Agent Performance API ──────────────────────────────────────────────────────

@app.route('/api/agent_perf')
def agent_perf():
    if not is_admin(): return jsonify([])
    df = get_safe_data()
    if df.empty or 'Assigned_To' not in df.columns: return jsonify([])

    date_from = request.args.get('date_from')
    date_to   = request.args.get('date_to')
    if date_from or date_to:
        df['Created_Date'] = pd.to_datetime(df['Created_Date'], errors='coerce')
        if date_from: df = df[df['Created_Date'] >= pd.to_datetime(date_from)]
        if date_to:   df = df[df['Created_Date'] <= pd.to_datetime(date_to)]

    results = []
    for agent, group in df.groupby('Assigned_To'):
        total      = len(group)
        resolved   = len(group[group['Status'].isin(['Resolved','Closed'])])
        open_count = len(group[group['Status'] == 'Open'])
        critical   = len(group[group['Priority'] == 'Critical'])
        res_times  = group['Resolution_Time_Hours'].dropna() if 'Resolution_Time_Hours' in group else pd.Series([], dtype=float)
        avg_res    = round(float(res_times.mean()), 1) if len(res_times) > 0 else 0
        breaches   = sum(1 for _, r in group.iterrows()
                         if get_sla_status(r.get('Priority'), r.get('Resolution_Time_Hours')) == 'breached')
        results.append({
            "agent":            str(agent),
            "total":            total,
            "resolved":         resolved,
            "open":             open_count,
            "critical":         critical,
            "avg_res":          avg_res if not pd.isna(avg_res) else 0,
            "resolution_rate":  round((resolved/total)*100, 1) if total > 0 else 0,
            "sla_breaches":     breaches
        })
    return jsonify(sorted(results, key=lambda x: x['total'], reverse=True))

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

if __name__ == '__main__':
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    app.run(debug=True, port=5000)
