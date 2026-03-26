from flask import Flask, render_template, jsonify, request, redirect, url_for, session, Response
import pandas as pd
import os, io, json, random, re, threading, smtplib, sqlite3
from datetime import datetime, timedelta
from werkzeug.security import generate_password_hash, check_password_hash

app = Flask(__name__)
app.secret_key = "it_project_2026_secure"
EXCEL_FILE = 'IT_Ticket_Performance_Data.xlsx'
USERS_FILE = 'users.json'

SLA_HOURS = {'Critical': 4, 'High': 8, 'Medium': 24, 'Low': 72}
COMMENTS_FILE = 'comments.json'

# Email settings (can be moved to .env)
SMTP_HOST = os.environ.get('SMTP_HOST', '')
SMTP_PORT = int(os.environ.get('SMTP_PORT', 587))
SMTP_USER = os.environ.get('SMTP_USER', '')
SMTP_PASS = os.environ.get('SMTP_PASS', '')
SMTP_FROM = os.environ.get('SMTP_FROM', '')

CATEGORY_KEYWORDS = {
    'Network':            ['network', 'vpn', 'wifi', 'internet', 'dns', 'lan', 'firewall', 'bandwidth'],
    'Access Management':  ['password', 'login', 'access', 'account', 'permission', 'authentication', 'auth', 'locked'],
    'Hardware':           ['printer', 'scanner', 'hardware', 'laptop', 'computer', 'monitor', 'keyboard', 'mouse', 'device'],
    'Software':           ['software', 'install', 'update', 'crash', 'error', 'bug', 'application', 'app', 'patch'],
    'Email':              ['email', 'outlook', 'office', 'mailbox', 'teams', 'calendar'],
    'Infrastructure':     ['server', 'database', 'storage', 'backup', 'cloud', 'vm', 'virtual'],
}

# ── User helpers ──────────────────────────────────────────────────────────────

def get_users():
    if not os.path.exists(USERS_FILE):
        default = {
            "admin@it.com": {
                "password": generate_password_hash("password123"),
                "role": "admin",
                "name": "Admin"
            }
        }
        with open(USERS_FILE, 'w') as f:
            json.dump(default, f, indent=2)
        return default
    with open(USERS_FILE) as f:
        return json.load(f)

def save_users(users):
    with open(USERS_FILE, 'w') as f:
        json.dump(users, f, indent=2)

def current_user_info():
    email = session.get('user')
    if not email:
        return None
    users = get_users()
    u = users.get(email, {})
    return {"email": email, "name": u.get("name", email), "role": u.get("role", "user")}

def is_admin():
    u = current_user_info()
    return u and u["role"] == "admin"

def classify_ticket(text_input):
    lower = text_input.lower()
    for cat, kws in CATEGORY_KEYWORDS.items():
        for kw in kws:
            if kw in lower: return cat
    return 'General'

def get_comments():
    if not os.path.exists(COMMENTS_FILE): return {}
    with open(COMMENTS_FILE) as f: return json.load(f)

def save_comments(comments):
    with open(COMMENTS_FILE, 'w') as f: json.dump(comments, f, indent=2)

def send_email_async(to_addr, subject, body):
    if not SMTP_HOST or not SMTP_USER: return
    def _send():
        try:
            msg = f"Subject: {subject}\n\n{body}"
            with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
                server.starttls()
                server.login(SMTP_USER, SMTP_PASS)
                server.sendmail(SMTP_FROM or SMTP_USER, [to_addr], msg)
        except Exception as e: print(f"Email error: {e}")
    threading.Thread(target=_send, daemon=True).start()

# ── Context processor: injects user info into every template ──────────────────
@app.context_processor
def inject_user():
    u = current_user_info()
    return dict(current_user=u, is_admin_user=is_admin())

# ── Data helpers ──────────────────────────────────────────────────────────────

def get_safe_data():
    if not os.path.exists(EXCEL_FILE):
        df_empty = pd.DataFrame(columns=[
            'Ticket_ID','Status','Priority','Category',
            'Assigned_To','Created_Date','Resolution_Time_Hours','Created_By'
        ])
        df_empty.to_excel(EXCEL_FILE, index=False)
        return df_empty
    try:
        df = pd.read_excel(EXCEL_FILE, engine='openpyxl')
        if 'Created_By' not in df.columns:
            df['Created_By'] = ''
        return df
    except Exception as e:
        print(f"❌ Excel Read Error: {e}")
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
    if h > limit:        return 'breached'
    elif h > limit * 0.75: return 'near_breach'
    else:                  return 'on_track'

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

# ── Routes ────────────────────────────────────────────────────────────────────

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
            return redirect(url_for('dashboard') if u['role'] == 'admin' else url_for('my_tickets'))
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
        users[email] = {"password": generate_password_hash(pw), "role": "user", "name": name}
        save_users(users)
        session['user'] = email
        return redirect(url_for('my_tickets'))
    return render_template('signup.html')

@app.route('/logout')
def logout():
    session.pop('user', None)
    return redirect(url_for('login'))

# Admin-only routes
@app.route('/dashboard')
def dashboard():
    if 'user' not in session: return redirect(url_for('login'))
    if not is_admin():        return redirect(url_for('my_tickets'))
    return render_template('index.html')

@app.route('/tickets')
def tickets():
    if 'user' not in session: return redirect(url_for('login'))
    if not is_admin():        return redirect(url_for('my_tickets'))
    return render_template('tickets.html')

@app.route('/manage')
def manage():
    if 'user' not in session: return redirect(url_for('login'))
    if not is_admin():        return redirect(url_for('my_tickets'))
    return render_template('manage.html')

@app.route('/sql')
def sql_page():
    if 'user' not in session or not is_admin(): return redirect(url_for('login'))
    return render_template('sql_query.html')

@app.route('/analytics')
def analytics_page():
    if 'user' not in session or not is_admin(): return redirect(url_for('login'))
    return render_template('analytics.html')

@app.route('/profile')
def profile_page():
    if 'user' not in session: return redirect(url_for('login'))
    return render_template('profile.html')

@app.route('/tickets/<ticket_id>')
def ticket_detail(ticket_id):
    if 'user' not in session: return redirect(url_for('login'))
    # Check if ticket exists
    df = get_safe_data()
    mask = df['Ticket_ID'].astype(str).str.strip() == ticket_id.strip()
    if not mask.any(): return redirect(url_for('tickets'))
    return render_template('ticket_detail.html', tid=ticket_id)

# User routes
@app.route('/create_ticket')
def create_ticket():
    if 'user' not in session: return redirect(url_for('login'))
    return render_template('create_ticket.html')

@app.route('/my_tickets')
def my_tickets():
    if 'user' not in session: return redirect(url_for('login'))
    return render_template('my_tickets.html')

# ── APIs ──────────────────────────────────────────────────────────────────────

@app.route('/api/stats')
def stats():
    df = get_safe_data()
    period = request.args.get('period', 'weekly')

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

        sla_counts = {"on_track":0,"near_breach":0,"breached":0}
        if 'Resolution_Time_Hours' in df.columns:
            for _, row in df.iterrows():
                s = get_sla_status(row.get('Priority','Low'), row.get('Resolution_Time_Hours'))
                if s in sla_counts: sla_counts[s] += 1

        tickets_today = 0
        if 'Created_Date' in df.columns:
            df['Created_Date'] = pd.to_datetime(df['Created_Date'], errors='coerce')
            tickets_today = int((df['Created_Date'].dt.date == datetime.now().date()).sum())

        avg_res = round(float(df['Resolution_Time_Hours'].mean()), 1) if 'Resolution_Time_Hours' in df.columns else 0

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
    if not is_admin(): return jsonify([])
    df = get_safe_data()
    if df.empty: return jsonify([])

    status   = request.args.get('status')
    priority = request.args.get('priority')
    assignee = request.args.get('assignee')
    search   = request.args.get('search')
    date_from = request.args.get('date_from')
    date_to   = request.args.get('date_to')

    if status   and status   != 'All': df = df[df['Status'].astype(str).str.contains(status,   case=False, na=False)]
    if priority and priority != 'All': df = df[df['Priority'].astype(str).str.contains(priority, case=False, na=False)]
    if assignee and assignee != 'All': df = df[df['Assigned_To'].astype(str).str.contains(assignee, case=False, na=False)]
    if search:
        mask = (df['Ticket_ID'].astype(str).str.contains(search, case=False, na=False) |
                df['Category'].astype(str).str.contains(search, case=False, na=False))
        df = df[mask]
    if date_from or date_to:
        df['Created_Date'] = pd.to_datetime(df['Created_Date'], errors='coerce')
        if date_from: df = df[df['Created_Date'] >= pd.to_datetime(date_from)]
        if date_to:   df = df[df['Created_Date'] <= pd.to_datetime(date_to)]

    if 'Resolution_Time_Hours' in df.columns and 'Priority' in df.columns:
        df = df.copy()
        df['SLA_Status'] = df.apply(lambda r: get_sla_status(r['Priority'], r['Resolution_Time_Hours']), axis=1)

    return jsonify(df.fillna('').to_dict(orient='records'))

@app.route('/api/my_tickets')
def api_my_tickets():
    if 'user' not in session: return jsonify([])
    df = get_safe_data()
    if df.empty: return jsonify([])
    email = session['user']
    df = df[df['Created_By'].astype(str).str.strip() == email]
    df = df.fillna('')
    # Add progress %
    PROGRESS = {'Open': 10, 'In Progress': 50, 'Resolved': 80, 'Closed': 100}
    df = df.copy()
    df['Progress'] = df['Status'].apply(lambda s: PROGRESS.get(s, 10))
    return jsonify(df.to_dict(orient='records'))

@app.route('/api/submit_ticket', methods=['POST'])
def submit_ticket():
    if 'user' not in session: return jsonify({"success": False, "error": "Not logged in"})
    data = request.get_json()
    category = data.get('Category', '').strip()
    priority  = data.get('Priority', 'Low')
    desc      = data.get('Description', '').strip()
    if not category:
        return jsonify({"success": False, "error": "Category is required"})

    df = get_safe_data()
    tid = generate_ticket_id()
    new_row = {
        'Ticket_ID':             tid,
        'Status':                'Open',
        'Priority':              priority,
        'Category':              category,
        'Assigned_To':           'Unassigned',
        'Created_Date':          datetime.now().strftime('%Y-%m-%d'),
        'Resolution_Time_Hours': '',
        'Created_By':            session['user']
    }
    if 'Description' not in df.columns:
        df['Description'] = ''
    new_row['Description'] = desc
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    df.to_excel(EXCEL_FILE, index=False)
    return jsonify({"success": True, "ticket_id": tid})

@app.route('/api/update_ticket', methods=['POST'])
def update_ticket():
    if not is_admin(): return jsonify({"success": False, "error": "Unauthorized"})
    data = request.get_json()
    ticket_id = data.get('Ticket_ID', '').strip()
    if not ticket_id: return jsonify({"success": False, "error": "Ticket_ID is required"})

    df = get_safe_data()
    mask = df['Ticket_ID'].astype(str).str.strip() == ticket_id
    if not mask.any(): return jsonify({"success": False})

    for field in ('Status', 'Priority', 'Assigned_To'):
        if field in data:
            df.loc[mask, field] = data[field]

    df.to_excel(EXCEL_FILE, index=False)
    return jsonify({"success": True})

@app.route('/api/bulk_update', methods=['POST'])
def bulk_update():
    if not is_admin(): return jsonify({"success": False, "error": "Unauthorized"})
    data = request.get_json()
    ticket_ids = data.get('ticket_ids', [])
    updates    = data.get('updates', {})
    if not ticket_ids: return jsonify({"success": False, "error": "No tickets selected"})

    df = get_safe_data()
    mask = df['Ticket_ID'].astype(str).str.strip().isin([str(t).strip() for t in ticket_ids])
    for field, value in updates.items():
        if field in df.columns:
            df.loc[mask, field] = value
    df.to_excel(EXCEL_FILE, index=False)
    return jsonify({"success": True, "updated": int(mask.sum())})

@app.route('/api/export')
def export():
    if not is_admin(): return jsonify({"error": "Unauthorized"}), 403
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
    if not is_admin(): return jsonify({"found": False})
    df = get_safe_data()
    mask = df['Ticket_ID'].astype(str).str.strip() == ticket_id.strip()
    if not mask.any(): return jsonify({"found": False})
    row = df[mask].iloc[0].fillna('').to_dict()
    row['found'] = True
    return jsonify(row)

@app.route('/api/ticket_ids')
def ticket_ids():
    if not is_admin(): return jsonify([])
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
        'Ticket_ID': data['Ticket_ID'].strip(), 'Status': data['Status'],
        'Priority': data['Priority'], 'Category': data['Category'].strip(),
        'Assigned_To': data['Assigned_To'].strip(),
        'Created_Date': datetime.now().strftime('%Y-%m-%d'),
        'Resolution_Time_Hours': data.get('Resolution_Time_Hours', ''),
        'Created_By': session.get('user', ''),
        'Description': data.get('Description', '').strip()
    }
    if 'Description' not in df.columns:
        df['Description'] = ''
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    df.to_excel(EXCEL_FILE, index=False)
    return jsonify({"success": True})

@app.route('/api/delete_ticket', methods=['POST'])
def delete_ticket():
    if not is_admin(): return jsonify({"success": False, "error": "Unauthorized"})
    data = request.get_json()
    ticket_id = data.get('Ticket_ID', '').strip()
    if not ticket_id: return jsonify({"success": False, "error": "Ticket_ID is required"})

    df = get_safe_data()
    mask = df['Ticket_ID'].astype(str).str.strip() == ticket_id
    if not mask.any(): return jsonify({"success": False, "error": "Ticket not found"})

    df = df[~mask]
    df.to_excel(EXCEL_FILE, index=False)
    return jsonify({"success": True})

@app.route('/api/notifications')
def notifications():
    df = get_safe_data()
    if df.empty: return jsonify([])
    notifs = []
    if 'Created_Date' in df.columns:
        df['Created_Date'] = pd.to_datetime(df['Created_Date'], errors='coerce')
        cutoff = datetime.now() - timedelta(days=1)
        recent = df[df['Created_Date'] >= cutoff].tail(10)
        for _, row in recent.iterrows():
            notifs.append({
                "id": str(row.get('Ticket_ID','')),
                "message": f"{row.get('Category','General')} — {row.get('Status','')}",
                "priority": str(row.get('Priority','')),
                "time": str(row.get('Created_Date',''))
            })
    if not notifs:
        for _, row in df.tail(5).fillna('').iterrows():
            notifs.append({"id": str(row.get('Ticket_ID','')),
                           "message": f"{row.get('Category','General')} — {row.get('Status','')}",
                           "priority": str(row.get('Priority','')),
                           "time": str(row.get('Created_Date',''))})
    return jsonify(notifs[:10])

@app.route('/api/classify', methods=['POST'])
def classify_api():
    data = request.json or {}
    text_input = data.get('text', '')
    return jsonify({'category': classify_ticket(text_input)})

@app.route('/api/tickets/<ticket_id>/comments', methods=['GET', 'POST'])
def ticket_comments(ticket_id):
    comments = get_comments()
    if request.method == 'POST':
        data = request.json or {}
        body = data.get('body', '').strip()
        if not body: return jsonify({"error": "Empty comment"}), 400
        
        c_list = comments.get(ticket_id, [])
        c_list.append({
            "author": session.get('user', 'Unknown'),
            "body": body,
            "time": datetime.now().strftime('%Y-%m-%d %H:%M')
        })
        comments[ticket_id] = c_list
        save_comments(comments)
        return jsonify({"success": True})
    
    return jsonify(comments.get(ticket_id, []))

@app.route('/api/sql', methods=['POST'])
def run_sql():
    if not is_admin(): return jsonify({"error": "Unauthorized"}), 403
    data = request.json or {}
    query = data.get('query', '').strip()
    if not query: return jsonify({"error": "Empty query"}), 400
    
    # Safety check: Only SELECT allowed
    if not re.search(r'^\s*SELECT', query, re.IGNORECASE):
        return jsonify({"error": "Only SELECT queries are allowed for safety."}), 403
    
    try:
        df = get_safe_data()
        # Create in-memory SQLite from Excel data
        conn = sqlite3.connect(':memory:')
        df.to_sql('tickets', conn, index=False)
        
        # Also include users for more interesting queries
        users_df = pd.DataFrame([{"email": k, "name": v["name"], "role": v["role"]} 
                               for k, v in get_users().items()])
        users_df.to_sql('users', conn, index=False)

        # Run query
        result_df = pd.read_sql_query(query, conn)
        return jsonify({
            "columns": result_df.columns.tolist(),
            "rows": result_df.fillna('').values.tolist(),
            "count": len(result_df)
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 400

if __name__ == '__main__':
    app.run(debug=True, port=5000)
