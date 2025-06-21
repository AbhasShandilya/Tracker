from flask import Flask, render_template, request, send_file, redirect, url_for
from flask_dance.contrib.google import make_google_blueprint, google
from flask_login import LoginManager, login_required, UserMixin, login_user, logout_user, current_user
import pandas as pd
import datetime
import os
from openpyxl.styles import Font, PatternFill
from openpyxl import Workbook

app = Flask(__name__)
app.secret_key = "supersecretkey"  # Replace with a strong key or use environment variable

# === Google OAuth Setup ===
google_bp = make_google_blueprint(
    client_id="294586716366-843n82lvud6f66f8ihul2jbh0i7lptgi.apps.googleusercontent.com",
    client_secret="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",
    redirect_url="https://tracker-77o2.onrender.com/login/callback",
    scope=["profile", "email"]
)
app.register_blueprint(google_bp, url_prefix="/login")

# === User Login Setup ===
login_manager = LoginManager()
login_manager.login_view = "google.login"
login_manager.init_app(app)

# Temporary user model
class User(UserMixin):
    def __init__(self, id, email):
        self.id = id
        self.email = email

# Store users in memory for demo (replace with DB in real app)
users = {}

@login_manager.user_loader
def load_user(user_id):
    return users.get(user_id)

@app.route("/login/callback")
def google_login():
    if not google.authorized:
        return redirect(url_for("google.login"))
    resp = google.get("/oauth2/v2/userinfo")
    if not resp.ok:
        return f"Failed to fetch user info: {resp.text}", 400
    user_info = resp.json()
    email = user_info["email"]
    user = User(id=email, email=email)
    users[email] = user
    login_user(user)
    return redirect(url_for("index"))

@app.route("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for("index"))

@app.route('/')
@login_required
def index():
    return render_template('index.html', user=current_user.email)

@app.route('/create_tracker', methods=['POST'])
@login_required
def create_tracker():
    project_name = request.form['project_name']
    tasks = request.form.getlist('task[]')
    statuses = request.form.getlist('status[]')
    owners = request.form.getlist('owner[]')
    start_dates = request.form.getlist('start_date[]')
    end_dates = request.form.getlist('end_date[]')
    priorities = request.form.getlist('priority[]')
    progress = request.form.getlist('progress[]')
    remarks = request.form.getlist('remark[]')

    df = pd.DataFrame({
        "Task": tasks,
        "Status": statuses,
        "Owner": owners,
        "Start Date": start_dates,
        "End Date": end_dates,
        "Priority": priorities,
        "Progress (%)": progress,
        "Remarks": remarks
    })

    today = datetime.date.today().strftime("%Y-%m-%d")
    filename = f"{project_name.replace(' ', '_')}_Tracker_{today}.xlsx"
    filepath = os.path.join("static", filename)

    with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Tracker')
        workbook = writer.book
        worksheet = writer.sheets['Tracker']

        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="4A90E2", end_color="4A90E2", fill_type="solid")

        for cell in worksheet[1]:
            cell.font = header_font
            cell.fill = header_fill

    return send_file(filepath, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
