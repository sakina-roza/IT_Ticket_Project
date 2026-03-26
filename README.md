# IT Ticket Management System

A comprehensive, Flask-based IT ticketing and performance tracking application designed for small to medium-sized IT teams. This system features a robust Admin Dashboard, real-time analytics, and automated SLA monitoring.

## 🚀 Features

### For Users
- **Secure Authentication**: Signup and Login with role-based access control.
- **Ticket Lifecycle**: Create, view, and track personal support tickets.
- **Interactive Comments**: Communicate directly with IT staff on specific tickets.
- **Real-time Progress**: Visual tracking of ticket status (Open, In Progress, Resolved, Closed).

### For Admins
- **Interactive Dashboard**: Real-time statistics including total tickets, critical issues, and average resolution time.
- **Performance Trends**: Weekly and monthly ticket volume visualization.
- **Advanced Management**: Bulk update ticket statuses, priorities, and assignees.
- **SLA Monitoring**: Automated tracking of Service Level Agreements (Critical, High, Medium, Low) with breach alerts.
- **Analytics & SQL**: In-memory SQL query interface for custom data exploration and advanced reporting.
- **Data Export**: Export ticket data to CSV for external processing.

## 🛠 Tech Stack
- **Backend**: Python (Flask)
- **Data Storage**: Microsoft Excel (`.xlsx`) via Pandas & SQLite (for analytics)
- **Frontend**: HTML5, Vanilla CSS, JS
- **Security**: Werkzeug password hashing
- **Automation**: Threaded email notifications

## 📥 Installation

1. **Clone the repository**:
   ```bash
   git clone <repository-url>
   cd IT_Ticket_Project
   ```

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Set up Environment Variables** (Optional for Email):
   - `SMTP_HOST`, `SMTP_PORT`, `SMTP_USER`, `SMTP_PASS`, `SMTP_FROM`

4. **Run the application**:
   ```bash
   python app.py
   ```
   The app will be available at `http://localhost:5000`.

## 📂 Project Structure
- `app.py`: Main Flask application and API logic.
- `templates/`: HTML templates for the dashboard, ticket management, etc.
- `IT_Ticket_Performance_Data.xlsx`: Central data store for tickets.
- `users.json`: User registry and authentication data.
- `comments.json`: Ticket comment history.

## ⚖ SLA Definitions
- **Critical**: 4 Hours
- **High**: 8 Hours
- **Medium**: 24 Hours
- **Low**: 72 Hours
