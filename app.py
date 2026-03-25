from flask import Flask, render_template, jsonify, request, redirect, url_for, session
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)
app.secret_key = "it_project_2026_secure"
EXCEL_FILE = 'IT_Ticket_Performance_Data.xlsx'

# Mock User
USERS = {"admin@it.com": "password123"}

def get_safe_data():
    if not os.path.exists(EXCEL_FILE):
        print("⚠️ Excel file not found. Creating a blank one...")
        df_empty = pd.DataFrame(columns=['Ticket_ID', 'Status', 'Priority', 'Category', 'Assigned_To', 'Created_Date', 'Resolution_Time_Hours'])
        df_empty.to_excel(EXCEL_FILE, index=False)
        return df_empty
    try:
        return pd.read_excel(EXCEL_FILE, engine='openpyxl')
    except Exception as e:
        print(f"❌ Excel Read Error: {e}")
        return pd.DataFrame()

@app.route('/')
def root():
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        email = request.form.get('email')
        pw = request.form.get('password')
        if email in USERS and USERS[email] == pw:
            session['user'] = email
            return redirect(url_for('dashboard'))
        return render_template('login.html', error="Invalid Credentials")
    return render_template('login.html')

@app.route('/dashboard')
def dashboard():
    if 'user' not in session: return redirect(url_for('login'))
    return render_template('index.html')

@app.route('/manage')
def manage():
    if 'user' not in session: return redirect(url_for('login'))
    return render_template('manage.html')

@app.route('/logout')
def logout():
    session.pop('user', None)
    return redirect(url_for('login'))

# --- API SECTION ---
@app.route('/api/stats')
def stats():
    df = get_safe_data()
    
    # Fallback if Excel is totally empty
    if df.empty or len(df) == 0:
        return jsonify({
            "stats": {"total": 0, "open": 0, "critical": 0, "avg_res": 0},
            "trend": {"labels": ["No Data"], "values": [0]},
            "recent": [],
            "sync_time": "NO DATA FOUND",
            "categories": ["General"]
        })

    try:
        # Data Cleanup
        df['Status'] = df['Status'].fillna('Open')
        df['Priority'] = df['Priority'].fillna('Low')
        
        # Calculations
        res_data = {
            "stats": {
                "total": int(len(df)),
                "open": int(len(df[df['Status'].str.contains('Open', case=False, na=False)])),
                "critical": int(len(df[df['Priority'].str.contains('Critical', case=False, na=False)])),
                "avg_res": round(float(df['Resolution_Time_Hours'].mean()), 1) if 'Resolution_Time_Hours' in df else 0
            },
            "trend": {"labels": ["Day 1", "Day 2", "Day 3"], "values": [len(df)//2, len(df), len(df)+2]},
            "recent": df.tail(8).to_dict(orient='records'),
            "sync_time": datetime.now().strftime("%H:%M:%S"),
            "categories": df['Category'].unique().tolist() if 'Category' in df else []
        }
        return jsonify(res_data)
    except Exception as e:
        print(f"API Error: {e}")
        return jsonify({"error": str(e)})

if __name__ == '__main__':
    app.run(debug=True, port=5000)