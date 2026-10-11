
from flask import Flask, request, redirect, render_template, send_file, flash
from flask_login import (
    LoginManager, UserMixin, login_user,
    login_required, logout_user, current_user
)
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import Workbook
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from io import BytesIO
import sqlite3
import os

website = Flask(__name__)
website.secret_key = os.environ.get(
    "FLASK_SECRET_KEY", "local-practice-secret-change-later"
)

login_manager = LoginManager()
login_manager.init_app(website)
login_manager.login_view = "login"

ADMIN_USERNAME = "admin"
ADMIN_PASSWORD_HASH = generate_password_hash("Admin@123")


class Admin(UserMixin):
    def __init__(self, username):
        self.id = username


@login_manager.user_loader
def load_user(user_id):
    if user_id == ADMIN_USERNAME:
        return Admin(user_id)
    return None


def connect_db():
    return sqlite3.connect("employees.db")


def create_table():
    db = connect_db()
    db.execute("""
        CREATE TABLE IF NOT EXISTS employees (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            role TEXT NOT NULL,
            salary REAL NOT NULL
        )
    """)
    db.commit()
    db.close()


def get_all_employees():
    db = connect_db()
    employees = db.execute("""
        SELECT id, name, role, salary
        FROM employees
        ORDER BY id DESC
    """).fetchall()
    db.close()
    return employees


@website.route("/login", methods=["GET", "POST"])
def login():
    if current_user.is_authenticated:
        return redirect("/")

    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")

        if (
            username == ADMIN_USERNAME
            and check_password_hash(ADMIN_PASSWORD_HASH, password)
        ):
            login_user(Admin(username))
            return redirect("/")

        flash("Incorrect username or password.")

    return render_template("login.html")


@website.route("/logout")
@login_required
def logout():
    logout_user()
    return redirect("/login")


@website.route("/", methods=["GET", "POST"])
@login_required
def home():
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        role = request.form.get("role", "").strip()

        try:
            salary = float(request.form.get("salary", ""))
            if not name or not role or salary < 0:
                raise ValueError
        except ValueError:
            return "Please enter valid employee details.", 400

        db = connect_db()
        db.execute("""
            INSERT INTO employees (name, role, salary)
            VALUES (?, ?, ?)
        """, (name, role, salary))
        db.commit()
        db.close()
        return redirect("/")

    search = request.args.get("search", "").strip()
    db = connect_db()

    if search:
        employees = db.execute("""
            SELECT id, name, role, salary
            FROM employees
            WHERE name LIKE ? OR role LIKE ?
            ORDER BY id DESC
        """, (f"%{search}%", f"%{search}%")).fetchall()
    else:
        employees = db.execute("""
            SELECT id, name, role, salary
            FROM employees
            ORDER BY id DESC
        """).fetchall()

    total_employees = db.execute(
        "SELECT COUNT(*) FROM employees"
    ).fetchone()[0]

    average_salary = db.execute(
        "SELECT COALESCE(AVG(salary), 0) FROM employees"
    ).fetchone()[0]

    db.close()

    return render_template(
        "index.html",
        employees=employees,
        search=search,
        total_employees=total_employees,
        average_salary=average_salary
    )


@website.route("/edit/<int:employee_id>", methods=["GET", "POST"])
@login_required
def edit_employee(employee_id):
    db = connect_db()

    if request.method == "POST":
        name = request.form.get("name", "").strip()
        role = request.form.get("role", "").strip()

        try:
            salary = float(request.form.get("salary", ""))
            if not name or not role or salary < 0:
                raise ValueError
        except ValueError:
            db.close()
            return "Please enter valid employee details.", 400

        db.execute("""
            UPDATE employees
            SET name = ?, role = ?, salary = ?
            WHERE id = ?
        """, (name, role, salary, employee_id))

        db.commit()
        db.close()
        return redirect("/")

    employee = db.execute("""
        SELECT name, role, salary
        FROM employees
        WHERE id = ?
    """, (employee_id,)).fetchone()

    db.close()

    if employee is None:
        return "Employee not found", 404

    return render_template(
        "edit.html",
        employee=employee,
        employee_id=employee_id
    )


@website.route("/delete/<int:employee_id>", methods=["POST"])
@login_required
def delete_employee(employee_id):
    db = connect_db()
    db.execute(
        "DELETE FROM employees WHERE id = ?",
        (employee_id,)
    )
    db.commit()
    db.close()
    return redirect("/")


@website.route("/download-excel")
@login_required
def download_excel():
    employees = get_all_employees()

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Employees"
    sheet.append(["ID", "Name", "Role", "Salary"])

    for employee in employees:
        sheet.append(list(employee))

    for cell in sheet[1]:
        cell.font = cell.font.copy(bold=True)

    for column in sheet.columns:
        max_length = max(
            len(str(cell.value or "")) for cell in column
        )
        sheet.column_dimensions[
            column[0].column_letter
        ].width = max_length + 3

    output = BytesIO()
    workbook.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="employee_report.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@website.route("/download-pdf")
@login_required
def download_pdf():
    employees = get_all_employees()

    output = BytesIO()
    document = SimpleDocTemplate(
        output,
        pagesize=landscape(A4)
    )

    styles = getSampleStyleSheet()
    elements = [
        Paragraph("Employee Management Report", styles["Title"]),
        Spacer(1, 18)
    ]

    data = [["ID", "Name", "Role", "Salary"]]

    for employee in employees:
        data.append([
            str(employee[0]),
            str(employee[1]),
            str(employee[2]),
            f"Rs. {float(employee[3]):,.2f}"
        ])

    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#172554")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1),
         [colors.white, colors.HexColor("#EFF6FF")]),
        ("PADDING", (0, 0), (-1, -1), 8),
    ]))

    elements.append(table)
    document.build(elements)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="employee_report.pdf",
        mimetype="application/pdf"
    )


if __name__ == "__main__":
    create_table()
    website.run(host="0.0.0.0", port=5000, debug=False)
