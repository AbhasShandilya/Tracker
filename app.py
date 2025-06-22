from flask import Flask, render_template, request, send_file, redirect, url_for
from flask_dance.contrib.google import make_google_blueprint, google
import pandas as pd
import datetime
import os
from openpyxl.styles import Font, PatternFill
from openpyxl import Workbook

app = Flask(__name__)
app.secret_key = "supersecretkey"  # Replace with a secure secret key in production

# ✅ Google OAuth Setup using DEFAULT REDIRECT URI
google_bp = make_google_blueprint(
    client_id="294586716366-843n82lvud6f66f8ihul2jbh0i7lptgi.apps.googleusercontent.com",
    client_secret="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",
    redirect_url="https://tracker-77o2.onrender.com/login/google/authorized",
    scope=["profile", "email"]
)
app.register_blueprint(google_bp, url_prefix="/login")


@app.route('/')
def index():
    if not google.authorized:
        return redirect(url_for("google.login"))

    # Get user info
    resp = google.get("/oauth2/v2/userinfo")
    if not resp.ok:
        return f"Failed to fetch user info: {resp.text}", 400
    user_info = resp.json()
    email = user_info.get("email")
    return render_template("index.html", user=email)


@app.route('/create_tracker', methods=['POST'])
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

        # Header styling
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="4A90E2", end_color="4A90E2", fill_type="solid")

        for cell in worksheet[1]:
            cell.font = header_font
            cell.fill = header_fill

    return send_file(filepath, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True)
