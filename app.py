from flask import Flask, render_template, request, redirect, url_for, send_file
from flask_login import LoginManager, UserMixin, login_user, logout_user, login_required, current_user
from flask_dance.contrib.google import make_google_blueprint, google
import pandas as pd
import os
import datetime

app = Flask(__name__)
app.secret_key = "your_very_secret_key"

# Allow OAuth over HTTP (for local development only)
os.environ["OAUTHLIB_INSECURE_TRANSPORT"] = "1"

# Flask-Login Setup
login_manager = LoginManager()
login_manager.login_view = "google.login"
login_manager.init_app(app)

# User class and in-memory storage
class User(UserMixin):
    def __init__(self, id_):
        self.id = id_

users = {}

@login_manager.user_loader
def load_user(user_id):
    return users.get(user_id)

# Google OAuth Blueprint
google_bp = make_google_blueprint(
    client_id="294586716366-843n82lvud6f66f8ihul2jbh0i7lptgi.apps.googleusercontent.com",
    client_secret="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",
    redirect_url="/login/callback",
    scope=["profile", "email"]
)
app.register_blueprint(google_bp, url_prefix="/login")

# Homepage — protected
@app.route('/')
@login_required
def index():
    return render_template('index.html')

# Google OAuth callback
@app.route('/login/callback')
def login_callback():
    if not google.authorized:
        return redirect(url_for("google.login"))

    resp = google.get("/oauth2/v2/userinfo")
    if not resp.ok:
        return "❌ Failed to fetch user info from Google", 400

    user_info = resp.json()
    user_id = user_info["id"]
    user = User(user_id)
    users[user_id] = user
    login_user(user)

    return redirect(url_for('index'))

# Logout
@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('index'))

# Excel Tracker Generator
@app.route('/create_tracker', methods=['POST'])
@login_required
def create_tracker():
    project_name = request.form['project_name']
    today = datetime.date.today().strftime("%Y-%m-%d")
    file_name = f"{project_name.replace(' ', '_')}_Tracker_{today}.xlsx"
    file_path = os.path.join("trackers", file_name)

    # Extract form data
    tasks = request.form.getlist('task')
    statuses = request.form.getlist('status')
    owners = request.form.getlist('owner')
    start_dates = request.form.getlist('start_date')
    end_dates = request.form.getlist('end_date')
    priorities = request.form.getlist('priority')
    progresses = request.form.getlist('progress')
    remarks = request.form.getlist('remarks')

    # Create DataFrame
    data = list(zip(tasks, statuses, owners, start_dates, end_dates, priorities, progresses, remarks))
    df = pd.DataFrame(data, columns=["Task", "Status", "Owner", "Start Date", "End Date", "Priority", "Progress (%)", "Remarks"])

    # Save Excel file
    os.makedirs("trackers", exist_ok=True)
    df.to_excel(file_path, index=False)

    return send_file(file_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
