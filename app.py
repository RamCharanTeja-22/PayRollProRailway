import os
import pandas as pd
import smtplib
import logging
from datetime import datetime, date
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from io import BytesIO
from flask import Flask, render_template_string, request, redirect, url_for, flash, send_file, jsonify, session
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, login_user, logout_user, login_required, current_user
from sqlalchemy.orm import DeclarativeBase
from werkzeug.utils import secure_filename
from werkzeug.middleware.proxy_fix import ProxyFix
from reportlab.lib.pagesizes import letter, A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

# Configure logging
logging.basicConfig(level=logging.DEBUG)


class Base(DeclarativeBase):
    pass


db = SQLAlchemy(model_class=Base)

# create the app
app = Flask(__name__)
app.secret_key = os.environ.get("SESSION_SECRET", "dev-secret-key-change-in-production")
app.wsgi_app = ProxyFix(app.wsgi_app, x_proto=1, x_host=1)  # needed for url_for to generate with https

# configure the database, relative to the app instance folder
app.config["SQLALCHEMY_DATABASE_URI"] = os.environ.get("DATABASE_URL", "sqlite:///payroll.db")
app.config["SQLALCHEMY_ENGINE_OPTIONS"] = {
    "pool_recycle": 300,
    "pool_pre_ping": True,
}
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# initialize the app with the extension, flask-sqlalchemy >= 3.0.x
db.init_app(app)

# Setup Flask-Login
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'
login_manager.login_message = 'Please log in to access this page.'

# Ensure upload directory exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

with app.app_context():
    # Make sure to import the models here or their tables won't be created
    import models  # noqa: F401
    db.create_all()


@login_manager.user_loader
def load_user(user_id):
    from models import User
    return db.session.get(User, int(user_id))


# Initialize default users if they don't exist
def init_default_users():
    from models import User
    
    try:
        # Check if users already exist
        user_count = db.session.query(User).count()
        app.logger.info(f"Current user count: {user_count}")
        
        if user_count == 0:
            # Create default users
            admin_user = User(username='admin', email='admin@payrollpro.com', role='admin')
            admin_user.set_password('admin123')
            
            hr_user = User(username='hr', email='hr@payrollpro.com', role='hr')
            hr_user.set_password('hr123')
            
            accounts_user = User(username='accounts', email='accounts@payrollpro.com', role='accounts')
            accounts_user.set_password('accounts123')
            
            db.session.add_all([admin_user, hr_user, accounts_user])
            db.session.commit()
            app.logger.info("Default users created: admin/admin, hr/hr, accounts/accounts")
            print("Default users created: admin/admin123, hr/hr123, accounts/accounts123")
        else:
            # Check if existing users can authenticate properly
            admin_user = db.session.query(User).filter_by(username='admin').first()
            if admin_user and not admin_user.check_password('admin'):
                app.logger.info("Password verification failed, updating user passwords")
                
                # Update passwords for existing users
                for username, password in [('admin', 'admin'), ('hr', 'hr'), ('accounts', 'accounts')]:
                    user = db.session.query(User).filter_by(username=username).first()
                    if user:
                        user.set_password(password)
                        app.logger.info(f"Updated password for user: {username}")
                
                db.session.commit()
                app.logger.info("User passwords updated successfully")
                print("User passwords updated successfully")
            else:
                app.logger.info("Users exist and passwords are valid")
            
    except Exception as e:
        db.session.rollback()
        app.logger.error(f"Error in init_default_users: {str(e)}")
        print(f"Error initializing default users: {str(e)}")


# Role-based access decorator
def role_required(*roles):
    def decorator(f):
        def decorated_function(*args, **kwargs):
            if not current_user.is_authenticated:
                return redirect(url_for('login'))
            if current_user.role not in roles:
                flash('You do not have permission to access this page.', 'error')
                return redirect(url_for('dashboard'))
            return f(*args, **kwargs)
        decorated_function.__name__ = f.__name__
        return decorated_function
    return decorator


# Call init_default_users after models are imported
def initialize_database():
    try:
        with app.app_context():
            # Ensure all tables are created
            db.create_all()
            # Initialize default users
            init_default_users()
    except Exception as e:
        print(f"Error initializing database: {e}")
        app.logger.error(f"Database initialization error: {e}")

# Initialize the database
initialize_database()


# HTML Template with embedded CSS and JS
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PayrollPro - Dashboard</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        :root {
            --primary-color: #2c3e50;
            --secondary-color: #3498db;
            --success-color: #27ae60;
            --warning-color: #f39c12;
            --danger-color: #e74c3c;
            --light-bg: #f8f9fa;
            --card-shadow: 0 0.125rem 0.25rem rgba(0,0,0,0.075);
        }

        body {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }

        .dashboard-container {
            background: white;
            margin: 20px;
            border-radius: 15px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.1);
            overflow: hidden;
        }

        .header {
            background: linear-gradient(135deg, var(--primary-color) 0%, var(--secondary-color) 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }

        .header h1 {
            margin: 0;
            font-size: 2.5rem;
            font-weight: 300;
        }

        .header p {
            margin: 10px 0 0 0;
            opacity: 0.9;
        }

        .stats-row {
            padding: 30px;
            background: var(--light-bg);
        }

        .stat-card {
            background: white;
            border-radius: 12px;
            padding: 25px;
            text-align: center;
            box-shadow: var(--card-shadow);
            transition: transform 0.3s ease, box-shadow 0.3s ease;
            border-left: 4px solid;
            margin-bottom: 20px;
        }

        .stat-card:hover {
            transform: translateY(-5px);
            box-shadow: 0 8px 25px rgba(0,0,0,0.15);
        }

        .stat-card.employees { border-left-color: var(--secondary-color); }
        .stat-card.payroll { border-left-color: var(--success-color); }
        .stat-card.pending { border-left-color: var(--warning-color); }
        .stat-card.total { border-left-color: var(--danger-color); }

        .stat-card i {
            font-size: 2.5rem;
            margin-bottom: 15px;
            opacity: 0.8;
        }

        .stat-card h3 {
            font-size: 2rem;
            font-weight: bold;
            margin: 0;
        }

        .stat-card p {
            margin: 0;
            color: #666;
            font-size: 0.9rem;
        }

        .main-content {
            padding: 30px;
        }

        .section-card {
            background: white;
            border-radius: 12px;
            padding: 25px;
            margin-bottom: 30px;
            box-shadow: var(--card-shadow);
            border: 1px solid #e9ecef;
        }

        .section-title {
            color: var(--primary-color);
            font-size: 1.3rem;
            font-weight: 600;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 2px solid #e9ecef;
        }

        .btn-custom {
            border-radius: 8px;
            padding: 10px 20px;
            font-weight: 500;
            transition: all 0.3s ease;
            border: none;
            margin: 5px;
        }

        .btn-primary-custom {
            background: linear-gradient(135deg, var(--secondary-color) 0%, #2980b9 100%);
            color: white;
        }

        .btn-primary-custom:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(52, 152, 219, 0.4);
            color: white;
        }

        .btn-success-custom {
            background: linear-gradient(135deg, var(--success-color) 0%, #219a52 100%);
            color: white;
        }

        .btn-success-custom:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(39, 174, 96, 0.4);
            color: white;
        }

        .btn-warning-custom {
            background: linear-gradient(135deg, var(--warning-color) 0%, #d68910 100%);
            color: white;
        }

        .btn-warning-custom:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(243, 156, 18, 0.4);
            color: white;
        }

        .btn-info-custom {
            background: linear-gradient(135deg, #17a2b8 0%, #138496 100%);
            color: white;
        }

        .btn-info-custom:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(23, 162, 184, 0.4);
            color: white;
        }

        .form-control, .form-select {
            border-radius: 8px;
            border: 1px solid #ddd;
            padding: 12px;
            transition: all 0.3s ease;
        }

        .form-control:focus, .form-select:focus {
            border-color: var(--secondary-color);
            box-shadow: 0 0 0 0.2rem rgba(52, 152, 219, 0.25);
        }

        .table {
            border-radius: 8px;
            overflow: hidden;
            box-shadow: var(--card-shadow);
        }

        .table thead {
            background: var(--primary-color);
            color: white;
        }

        .table tbody tr:hover {
            background-color: rgba(52, 152, 219, 0.1);
        }

        .alert {
            border-radius: 8px;
            border: none;
            box-shadow: var(--card-shadow);
        }

        .modal-content {
            border-radius: 12px;
            border: none;
            box-shadow: 0 10px 30px rgba(0,0,0,0.2);
        }

        .modal-header {
            background: var(--primary-color);
            color: white;
            border-radius: 12px 12px 0 0;
        }

        .file-upload-area {
            border: 2px dashed #ddd;
            border-radius: 8px;
            padding: 30px;
            text-align: center;
            background: #f8f9fa;
            transition: all 0.3s ease;
        }

        .file-upload-area:hover {
            border-color: var(--secondary-color);
            background: rgba(52, 152, 219, 0.1);
        }

        .chart-container {
            position: relative;
            height: 300px;
            margin: 20px 0;
        }

        @media (max-width: 768px) {
            .dashboard-container {
                margin: 10px;
            }
            
            .header {
                padding: 20px;
            }
            
            .header h1 {
                font-size: 2rem;
            }
            
            .stats-row, .main-content {
                padding: 20px;
            }
        }
    </style>
</head>
<body>
    <div class="dashboard-container">
        <!-- Header -->
        <div class="header d-flex justify-content-between align-items-center">
            <div>
                <h1><i class="fas fa-calculator"></i> PayrollPro</h1>
                <p>Comprehensive Payroll Management System</p>
            </div>
            <div class="user-info">
                <span class="me-3 text-white">
                    <i class="fas fa-user"></i> {{ current_user.username }} ({{ current_user.role.title() }})
                </span>
                <a href="{{ url_for('logout') }}" class="btn btn-outline-light btn-sm">
                    <i class="fas fa-sign-out-alt"></i> Logout
                </a>
            </div>
        </div>

        <!-- Flash Messages -->
        {% with messages = get_flashed_messages(with_categories=true) %}
            {% if messages %}
                <div class="container mt-3">
                    {% for category, message in messages %}
                        <div class="alert alert-{{ 'danger' if category == 'error' else 'success' }} alert-dismissible fade show" role="alert">
                            {{ message }}
                            <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
                        </div>
                    {% endfor %}
                </div>
            {% endif %}
        {% endwith %}

        <!-- Statistics Row -->
        <div class="stats-row">
            <div class="row">
                <div class="col-md-3">
                    <div class="stat-card employees">
                        <i class="fas fa-users text-info"></i>
                        <h3>{{ total_employees }}</h3>
                        <p>Total Employees</p>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="stat-card payroll">
                        <i class="fas fa-money-bill-wave text-success"></i>
                        <h3>{{ recent_payroll|length }}</h3>
                        <p>Recent Payrolls</p>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="stat-card pending">
                        <i class="fas fa-clock text-warning"></i>
                        <h3>{{ monthly_stats|length }}</h3>
                        <p>Monthly Records</p>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="stat-card total">
                        <i class="fas fa-chart-line text-danger"></i>
                        <h3>₹{{ "%.0f"|format(monthly_stats[0].total_payout if monthly_stats else 0) }}</h3>
                        <p>Latest Month Payout</p>
                    </div>
                </div>
            </div>
        </div>

        <!-- Main Content -->
        <div class="main-content">
            <div class="row">
                <!-- Employee Management -->
                <div class="col-lg-6">
                    <div class="section-card">
                        <h3 class="section-title"><i class="fas fa-user-plus"></i> Employee Management</h3>
                        
                        <!-- Add Single Employee -->
                        <div class="mb-4">
                            <button class="btn btn-primary-custom btn-custom" data-bs-toggle="modal" data-bs-target="#addEmployeeModal">
                                <i class="fas fa-user-plus"></i> Add Employee
                            </button>
                            <button class="btn btn-success-custom btn-custom" data-bs-toggle="modal" data-bs-target="#bulkEmployeeModal">
                                <i class="fas fa-upload"></i> Bulk Add
                            </button>
                            <a href="{{ url_for('download_employee_template') }}" class="btn btn-warning-custom btn-custom">
                                <i class="fas fa-download"></i> Template
                            </a>
                        </div>

                        <!-- Recent Employees -->
                        <div class="table-responsive">
                            <table class="table table-hover">
                                <thead>
                                    <tr>
                                        <th>ID</th>
                                        <th>Name</th>
                                        <th>CTC</th>
                                        <th>Department</th>
                                        <th>Actions</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {% for emp in employees[:5] %}
                                    <tr>
                                        <td>{{ emp.emp_id }}</td>
                                        <td>{{ emp.name }}</td>
                                        <td>₹{{ "%.0f"|format(emp.ctc_monthly) }}</td>
                                        <td>{{ emp.department or 'N/A' }}</td>
                                        <td>
                                            <button class="btn btn-sm btn-info-custom" onclick="viewEmployeeCostBreakdown('{{ emp.emp_id }}', '{{ emp.name }}', {{ emp.ctc_monthly }})">
                                                <i class="fas fa-calculator"></i> View
                                            </button>
                                            <a href="/employee_dashboard/{{ emp.emp_id }}" class="btn btn-sm btn-success-custom ms-1">
                                                <i class="fas fa-calendar-check"></i> Leave
                                            </a>
                                        </td>
                                    </tr>
                                    {% endfor %}
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>

                <!-- Payroll Management -->
                <div class="col-lg-6">
                    <div class="section-card">
                        <h3 class="section-title"><i class="fas fa-calculator"></i> Payroll Management</h3>
                        
                        <!-- Process Payroll -->
                        <div class="mb-4">
                            <button class="btn btn-primary-custom btn-custom" data-bs-toggle="modal" data-bs-target="#processPayrollModal">
                                <i class="fas fa-calculator"></i> Process Payroll
                            </button>
                            <button class="btn btn-success-custom btn-custom" data-bs-toggle="modal" data-bs-target="#bulkPayrollModal">
                                <i class="fas fa-upload"></i> Bulk Process
                            </button>
                            <a href="{{ url_for('download_payroll_template') }}" class="btn btn-warning-custom btn-custom">
                                <i class="fas fa-download"></i> Template
                            </a>
                        </div>

                        <!-- Recent Payroll -->
                        <div class="table-responsive">
                            <table class="table table-hover">
                                <thead>
                                    <tr>
                                        <th>Employee</th>
                                        <th>Month</th>
                                        <th>Net Salary</th>
                                        <th>Actions</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {% for payroll in recent_payroll %}
                                    <tr>
                                        <td>{{ payroll.name }}</td>
                                        <td>{{ payroll.month }} {{ payroll.year }}</td>
                                        <td>₹{{ "%.0f"|format(payroll.net_salary) }}</td>
                                        <td>
                                            <button class="btn btn-sm btn-primary-custom" onclick="viewPayslip('{{ payroll.emp_id }}', '{{ payroll.month }}', {{ payroll.year }})">
                                                <i class="fas fa-eye"></i> View
                                            </button>
                                            <a href="/download_payslip/{{ payroll.emp_id }}/{{ payroll.month }}/{{ payroll.year }}" class="btn btn-sm btn-success-custom">
                                                <i class="fas fa-download"></i> PDF
                                            </a>
                                        </td>
                                    </tr>
                                    {% endfor %}
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>

            <!-- Enhanced Payroll Management -->
            <div class="row">
                <div class="col-lg-12">
                    <div class="section-card">
                        <h3 class="section-title"><i class="fas fa-file-invoice-dollar"></i> All Payroll Records</h3>
                        
                        <!-- Filter Options -->
                        <div class="row mb-4">
                            <div class="col-md-3">
                                <select class="form-select" id="monthFilter">
                                    <option value="">All Months</option>
                                    <option value="January">January</option>
                                    <option value="February">February</option>
                                    <option value="March">March</option>
                                    <option value="April">April</option>
                                    <option value="May">May</option>
                                    <option value="June">June</option>
                                    <option value="July">July</option>
                                    <option value="August">August</option>
                                    <option value="September">September</option>
                                    <option value="October">October</option>
                                    <option value="November">November</option>
                                    <option value="December">December</option>
                                </select>
                            </div>
                            <div class="col-md-3">
                                <select class="form-select" id="yearFilter">
                                    <option value="">All Years</option>
                                    <option value="2024">2024</option>
                                    <option value="2023">2023</option>
                                    <option value="2025">2025</option>
                                </select>
                            </div>
                            <div class="col-md-6 text-end">
                                <button class="btn btn-primary-custom btn-custom" onclick="loadAllPayrolls()">
                                    <i class="fas fa-sync-alt"></i> Refresh
                                </button>
                                <button class="btn btn-success-custom btn-custom" data-bs-toggle="modal" data-bs-target="#emailModal">
                                    <i class="fas fa-paper-plane"></i> Send Payslips
                                </button>
                                <button class="btn btn-warning-custom btn-custom" onclick="downloadReport()">
                                    <i class="fas fa-file-excel"></i> Download Report
                                </button>
                            </div>
                        </div>

                        <!-- All Payroll Records Table -->
                        <div class="table-responsive" id="allPayrollTable">
                            <table class="table table-hover">
                                <thead>
                                    <tr>
                                        <th>Employee ID</th>
                                        <th>Name</th>
                                        <th>Month/Year</th>
                                        <th>Leaves Taken</th>
                                        <th>Gross Salary</th>
                                        <th>Net Salary</th>
                                        <th>Actions</th>
                                    </tr>
                                </thead>
                                <tbody id="payrollTableBody">
                                    {% for payroll in all_payroll %}
                                    <tr>
                                        <td>{{ payroll.emp_id }}</td>
                                        <td>{{ payroll.name }}</td>
                                        <td>{{ payroll.month }} {{ payroll.year }}</td>
                                        <td>{{ payroll.leaves_taken }} days</td>
                                        <td>₹{{ "%.2f"|format(payroll.gross_salary) }}</td>
                                        <td>₹{{ "%.2f"|format(payroll.net_salary) }}</td>
                                        <td>
                                            <button class="btn btn-sm btn-primary-custom" onclick="viewPayslip('{{ payroll.emp_id }}', '{{ payroll.month }}', {{ payroll.year }})">
                                                <i class="fas fa-eye"></i> View
                                            </button>
                                            <a href="/download_payslip/{{ payroll.emp_id }}/{{ payroll.month }}/{{ payroll.year }}" class="btn btn-sm btn-success-custom">
                                                <i class="fas fa-download"></i> PDF
                                            </a>
                                        </td>
                                    </tr>
                                    {% endfor %}
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>

            <!-- Additional Actions -->
            <div class="row">
                <div class="col-lg-6">
                    <div class="section-card">
                        <h3 class="section-title"><i class="fas fa-arrow-up"></i> Salary Hike Management</h3>
                        <button class="btn btn-primary-custom btn-custom" data-bs-toggle="modal" data-bs-target="#hikeModal">
                            <i class="fas fa-arrow-up"></i> Apply Hike
                        </button>
                        <p class="mt-2 text-muted">Select employee and apply salary hike during payroll processing</p>
                    </div>
                </div>
                <div class="col-lg-6">
                    <div class="section-card">
                        <h3 class="section-title"><i class="fas fa-users"></i> Employee Overview</h3>
                        <button class="btn btn-warning-custom btn-custom" data-bs-toggle="modal" data-bs-target="#allEmployeesModal">
                            <i class="fas fa-list"></i> View All Employees
                        </button>
                        <p class="mt-2 text-muted">View complete employee list with details and salary information</p>
                    </div>
                </div>
            </div>

            <!-- Charts -->
            <div class="row">
                <div class="col-12">
                    <div class="section-card">
                        <h3 class="section-title"><i class="fas fa-chart-bar"></i> Monthly Payroll Analytics</h3>
                        <div class="chart-container">
                            <canvas id="payrollChart"></canvas>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <!-- Modals -->
    
    <!-- Add Employee Modal -->
    <div class="modal fade" id="addEmployeeModal" tabindex="-1">
        <div class="modal-dialog modal-lg">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title"><i class="fas fa-user-plus"></i> Add Employee</h5>
                    <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
                </div>
                <form method="POST" action="{{ url_for('add_employee') }}">
                    <div class="modal-body">
                        <div class="row">
                            <div class="col-md-6">
                                <div class="mb-3">
                                    <label for="emp_id" class="form-label">Employee ID</label>
                                    <input type="text" class="form-control" name="emp_id" required>
                                </div>
                            </div>
                            <div class="col-md-6">
                                <div class="mb-3">
                                    <label for="name" class="form-label">Full Name</label>
                                    <input type="text" class="form-control" name="name" required>
                                </div>
                            </div>
                        </div>
                        <div class="row">
                            <div class="col-md-6">
                                <div class="mb-3">
                                    <label for="email" class="form-label">Email Address</label>
                                    <input type="email" class="form-control" name="email" required>
                                </div>
                            </div>
                            <div class="col-md-6">
                                <div class="mb-3">
                                    <label for="designation" class="form-label">Designation</label>
                                    <input type="text" class="form-control" name="designation">
                                </div>
                            </div>
                        </div>
                        <div class="row">
                            <div class="col-md-6">
                                <div class="mb-3">
                                    <label for="department" class="form-label">Department</label>
                                    <input type="text" class="form-control" name="department">
                                </div>
                            </div>
                            <div class="col-md-6">
                                <div class="mb-3">
                                    <label for="joining_date" class="form-label">Joining Date</label>
                                    <input type="date" class="form-control" name="joining_date">
                                </div>
                            </div>
                        </div>
                        <div class="row">
                            <div class="col-md-6">
                                <div class="mb-3">
                                    <label for="ctc_monthly" class="form-label">Monthly CTC (₹)</label>
                                    <input type="number" class="form-control" name="ctc_monthly" step="0.01" required>
                                </div>
                            </div>
                            <div class="col-md-6">
                                <div class="mb-3">
                                    <label class="form-label">PF Opted</label>
                                    <div class="form-check mt-2">
                                        <input class="form-check-input" type="checkbox" name="pf_opted" checked>
                                        <label class="form-check-label">Employee opts for Provident Fund</label>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
                        <button type="submit" class="btn btn-primary-custom">Add Employee</button>
                    </div>
                </form>
            </div>
        </div>
    </div>

    <!-- Bulk Add Employee Modal -->
    <div class="modal fade" id="bulkEmployeeModal" tabindex="-1">
        <div class="modal-dialog">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title"><i class="fas fa-upload"></i> Bulk Add Employees</h5>
                    <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
                </div>
                <form method="POST" action="{{ url_for('bulk_add_employees') }}" enctype="multipart/form-data">
                    <div class="modal-body">
                        <div class="file-upload-area">
                            <i class="fas fa-cloud-upload-alt fa-3x mb-3 text-muted"></i>
                            <h5>Upload Employee Excel File</h5>
                            <p class="text-muted">Select your filled employee template file</p>
                            <input type="file" class="form-control" name="file" accept=".xlsx,.xls" required>
                        </div>
                        <div class="mt-3">
                            <small class="text-muted">
                                <strong>Template Format:</strong> emp_id, name, email, designation, department, joining_date, ctc_monthly, pf_opted
                            </small>
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
                        <button type="submit" class="btn btn-success-custom">Upload & Process</button>
                    </div>
                </form>
            </div>
        </div>
    </div>

    <!-- Process Payroll Modal -->
    <div class="modal fade" id="processPayrollModal" tabindex="-1">
        <div class="modal-dialog modal-lg">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title"><i class="fas fa-calculator"></i> Process Individual Payroll</h5>
                    <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
                </div>
                <form method="POST" action="{{ url_for('process_individual_payroll') }}">
                    <div class="modal-body">
                        <div class="row">
                            <div class="col-md-6">
                                <div class="mb-3">
                                    <label for="emp_id_select" class="form-label">Select Employee</label>
                                    <select class="form-select" name="emp_id" required>
                                        <option value="">Choose Employee...</option>
                                        {% for emp in employees %}
                                        <option value="{{ emp.emp_id }}">{{ emp.emp_id }} - {{ emp.name }}</option>
                                        {% endfor %}
                                    </select>
                                </div>
                            </div>
                            <div class="col-md-6">
                                <div class="mb-3">
                                    <label for="leaves_taken" class="form-label">Leaves Taken</label>
                                    <input type="number" class="form-control" name="leaves_taken" value="0" min="0" max="31" required>
                                    <small class="text-muted">Monthly allocation: 1.5 days. Balance carries forward.</small>
                                </div>
                            </div>
                        </div>
                        <div class="row">
                            <div class="col-md-6">
                                <div class="mb-3">
                                    <label for="month" class="form-label">Month</label>
                                    <select class="form-select" name="month" required>
                                        <option value="January">January</option>
                                        <option value="February">February</option>
                                        <option value="March">March</option>
                                        <option value="April">April</option>
                                        <option value="May">May</option>
                                        <option value="June">June</option>
                                        <option value="July">July</option>
                                        <option value="August">August</option>
                                        <option value="September">September</option>
                                        <option value="October">October</option>
                                        <option value="November">November</option>
                                        <option value="December">December</option>
                                    </select>
                                </div>
                            </div>
                            <div class="col-md-6">
                                <div class="mb-3">
                                    <label for="year" class="form-label">Year</label>
                                    <input type="number" class="form-control" name="year" value="2024" min="2020" max="2030" required>
                                </div>
                            </div>
                        </div>
                        <div class="row">
                            <div class="col-md-6">
                                <div class="mb-3">
                                    <label class="form-label">PF Opted</label>
                                    <div class="form-check mt-2">
                                        <input class="form-check-input" type="checkbox" name="pf_opted" checked>
                                        <label class="form-check-label">Employee opts for PF deduction</label>
                                    </div>
                                </div>
                            </div>
                            <div class="col-md-6">
                                <div class="mb-3">
                                    <label for="hike_amount" class="form-label">Hike Amount (₹)</label>
                                    <input type="number" class="form-control" name="hike_amount" value="0" step="0.01">
                                    <small class="text-muted">Optional: Add hike to monthly CTC (this month only)</small>
                                </div>
                            </div>
                        </div>
                        <div class="row">
                            <div class="col-md-6">
                                <div class="mb-3">
                                    <label for="deduction_amount" class="form-label">Additional Deduction (₹)</label>
                                    <input type="number" class="form-control" name="deduction_amount" value="0" step="0.01">
                                    <small class="text-muted">Optional: Add deduction for this month only</small>
                                </div>
                            </div>
                            <div class="col-md-6">
                                <div class="mb-3">
                                    <label for="deduction_reason" class="form-label">Deduction Reason</label>
                                    <input type="text" class="form-control" name="deduction_reason" placeholder="e.g., Advance, Fine, etc.">
                                    <small class="text-muted">Optional: Reason for deduction</small>
                                </div>
                            </div>
                        </div>
                        <hr>
                        <div class="row">
                            <div class="col-md-4">
                                <div class="mb-3">
                                    <label for="overtime_amount" class="form-label">Overtime (₹)</label>
                                    <input type="number" class="form-control" name="overtime_amount" value="0" step="0.01">
                                </div>
                            </div>
                            <div class="col-md-4">
                                <div class="mb-3">
                                    <label for="expenses" class="form-label">Expenses (₹)</label>
                                    <input type="number" class="form-control" name="expenses" value="0" step="0.01">
                                </div>
                            </div>
                            <div class="col-md-4">
                                <div class="mb-3">
                                    <label for="bonus" class="form-label">Bonus (₹)</label>
                                    <input type="number" class="form-control" name="bonus" value="0" step="0.01">
                                </div>
                            </div>
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
                        <button type="submit" class="btn btn-primary-custom">Process Payroll</button>
                    </div>
                </form>
            </div>
        </div>
    </div>

    <!-- Bulk Process Payroll Modal -->
    <div class="modal fade" id="bulkPayrollModal" tabindex="-1">
        <div class="modal-dialog">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title"><i class="fas fa-upload"></i> Bulk Process Payroll</h5>
                    <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
                </div>
                <form method="POST" action="{{ url_for('bulk_process_payroll') }}" enctype="multipart/form-data">
                    <div class="modal-body">
                        <div class="row">
                            <div class="col-md-6">
                                <div class="mb-3">
                                    <label for="month" class="form-label">Month</label>
                                    <select class="form-select" name="month" required>
                                        <option value="January">January</option>
                                        <option value="February">February</option>
                                        <option value="March">March</option>
                                        <option value="April">April</option>
                                        <option value="May">May</option>
                                        <option value="June">June</option>
                                        <option value="July">July</option>
                                        <option value="August">August</option>
                                        <option value="September">September</option>
                                        <option value="October">October</option>
                                        <option value="November">November</option>
                                        <option value="December">December</option>
                                    </select>
                                </div>
                            </div>
                            <div class="col-md-6">
                                <div class="mb-3">
                                    <label for="year" class="form-label">Year</label>
                                    <input type="number" class="form-control" name="year" value="2024" min="2020" max="2030" required>
                                </div>
                            </div>
                        </div>
                        <div class="file-upload-area">
                            <i class="fas fa-cloud-upload-alt fa-3x mb-3 text-muted"></i>
                            <h5>Upload Payroll Excel File</h5>
                            <p class="text-muted">Select your filled payroll template file</p>
                            <input type="file" class="form-control" name="file" accept=".xlsx,.xls" required>
                        </div>
                        <div class="mt-3">
                            <small class="text-muted">
                                <strong>Template Format:</strong> emp_id, name, leaves_taken, pf_opted
                            </small>
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
                        <button type="submit" class="btn btn-success-custom">Upload & Process</button>
                    </div>
                </form>
            </div>
        </div>
    </div>

    <!-- Salary Hike Modal -->
    <div class="modal fade" id="hikeModal" tabindex="-1">
        <div class="modal-dialog">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title"><i class="fas fa-arrow-up"></i> Apply Salary Hike</h5>
                    <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
                </div>
                <form method="POST" action="{{ url_for('apply_hike') }}">
                    <div class="modal-body">
                        <div class="mb-3">
                            <label for="hike_emp_id" class="form-label">Select Employee</label>
                            <select class="form-select" name="emp_id" required>
                                <option value="">Choose Employee...</option>
                                {% for emp in employees %}
                                <option value="{{ emp.emp_id }}">{{ emp.emp_id }} - {{ emp.name }} (Current: ₹{{ "%.0f"|format(emp.ctc_monthly) }})</option>
                                {% endfor %}
                            </select>
                        </div>
                        <div class="mb-3">
                            <label for="hike_amount" class="form-label">Hike Amount (₹)</label>
                            <input type="number" class="form-control" name="hike_amount" step="0.01" required>
                            <small class="text-muted">This amount will be added to the current monthly CTC</small>
                        </div>
                        <div class="mb-3">
                            <label for="hike_reason" class="form-label">Reason for Hike</label>
                            <textarea class="form-control" name="hike_reason" rows="3" placeholder="Performance increment, promotion, etc."></textarea>
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
                        <button type="submit" class="btn btn-primary-custom">Apply Hike</button>
                    </div>
                </form>
            </div>
        </div>
    </div>

    <!-- Email Modal -->
    <div class="modal fade" id="emailModal" tabindex="-1">
        <div class="modal-dialog">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title"><i class="fas fa-envelope"></i> Send Payslips</h5>
                    <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
                </div>
                <form method="POST" action="{{ url_for('send_payslips') }}">
                    <div class="modal-body">
                        <div class="row">
                            <div class="col-md-6">
                                <div class="mb-3">
                                    <label for="email_month" class="form-label">Month</label>
                                    <select class="form-select" name="month" required>
                                        <option value="January">January</option>
                                        <option value="February">February</option>
                                        <option value="March">March</option>
                                        <option value="April">April</option>
                                        <option value="May">May</option>
                                        <option value="June">June</option>
                                        <option value="July">July</option>
                                        <option value="August">August</option>
                                        <option value="September">September</option>
                                        <option value="October">October</option>
                                        <option value="November">November</option>
                                        <option value="December">December</option>
                                    </select>
                                </div>
                            </div>
                            <div class="col-md-6">
                                <div class="mb-3">
                                    <label for="email_year" class="form-label">Year</label>
                                    <input type="number" class="form-control" name="year" value="2024" min="2020" max="2030" required>
                                </div>
                            </div>
                        </div>
                        <div class="alert alert-info">
                            <i class="fas fa-info-circle"></i> 
                            This will send payslips to all employees who have processed payroll for the selected month/year.
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
                        <button type="submit" class="btn btn-success-custom">Send Payslips</button>
                    </div>
                </form>
            </div>
        </div>
    </div>

    <!-- Payslip View Modal -->
    <div class="modal fade" id="payslipModal" tabindex="-1">
        <div class="modal-dialog modal-xl">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title"><i class="fas fa-file-invoice-dollar"></i> Payslip Details</h5>
                    <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
                </div>
                <div class="modal-body" id="payslipContent">
                    <!-- Payslip content will be loaded here -->
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Close</button>
                    <button type="button" class="btn btn-success-custom" id="downloadPayslipBtn">
                        <i class="fas fa-download"></i> Download PDF
                    </button>
                </div>
            </div>
        </div>
    </div>

    <!-- All Employees Modal -->
    <div class="modal fade" id="allEmployeesModal" tabindex="-1">
        <div class="modal-dialog modal-xl">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title"><i class="fas fa-users"></i> All Employees</h5>
                    <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
                </div>
                <div class="modal-body">
                    <div class="table-responsive">
                        <table class="table table-hover">
                            <thead>
                                <tr>
                                    <th>Employee ID</th>
                                    <th>Name</th>
                                    <th>Email</th>
                                    <th>Designation</th>
                                    <th>Department</th>
                                    <th>Monthly CTC</th>
                                    <th>PF Opted</th>
                                    <th>Joining Date</th>
                                </tr>
                            </thead>
                            <tbody>
                                {% for emp in employees %}
                                <tr>
                                    <td>{{ emp.emp_id }}</td>
                                    <td>{{ emp.name }}</td>
                                    <td>{{ emp.email }}</td>
                                    <td>{{ emp.designation or 'N/A' }}</td>
                                    <td>{{ emp.department or 'N/A' }}</td>
                                    <td>₹{{ "%.2f"|format(emp.ctc_monthly) }}</td>
                                    <td>
                                        {% if emp.pf_opted %}
                                            <span class="badge bg-success">Yes</span>
                                        {% else %}
                                            <span class="badge bg-danger">No</span>
                                        {% endif %}
                                    </td>
                                    <td>{{ emp.joining_date or 'N/A' }}</td>
                                </tr>
                                {% endfor %}
                            </tbody>
                        </table>
                    </div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Close</button>
                </div>
            </div>
        </div>
    </div>

    <!-- Employee Cost Breakdown Modal -->
    <div class="modal fade" id="employeeCostModal" tabindex="-1">
        <div class="modal-dialog modal-lg">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title"><i class="fas fa-calculator"></i> Employee Cost Breakdown</h5>
                    <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
                </div>
                <div class="modal-body" id="employeeCostContent">
                    <!-- Cost breakdown content will be loaded here -->
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Close</button>
                </div>
            </div>
        </div>
    </div>

    <!-- Bootstrap JS -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    
    <script>
        // Chart.js for payroll analytics
        document.addEventListener('DOMContentLoaded', function() {
            const ctx = document.getElementById('payrollChart').getContext('2d');
            
            // Sample data - this would come from the backend in a real application
            const monthlyData = {{ monthly_stats|tojson }};
            const labels = [];
            const data = [];
            
            monthlyData.forEach(function(item) {
                labels.push(item.month + ' ' + item.year);
                data.push(item.total_payout || 0);
            });
            
            new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: labels.reverse(),
                    datasets: [{
                        label: 'Monthly Payout (₹)',
                        data: data.reverse(),
                        backgroundColor: 'rgba(52, 152, 219, 0.8)',
                        borderColor: 'rgba(52, 152, 219, 1)',
                        borderWidth: 1
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {
                        y: {
                            beginAtZero: true,
                            ticks: {
                                callback: function(value) {
                                    return '₹' + value.toLocaleString();
                                }
                            }
                        }
                    },
                    plugins: {
                        legend: {
                            display: true,
                            position: 'top'
                        },
                        tooltip: {
                            callbacks: {
                                label: function(context) {
                                    return 'Total Payout: ₹' + context.parsed.y.toLocaleString();
                                }
                            }
                        }
                    }
                }
            });
        });

        // Auto-dismiss alerts after 5 seconds
        setTimeout(function() {
            const alerts = document.querySelectorAll('.alert');
            alerts.forEach(function(alert) {
                const bsAlert = new bootstrap.Alert(alert);
                bsAlert.close();
            });
        }, 5000);

        // View payslip function
        function viewPayslip(empId, month, year) {
            fetch(`/api/payslip/${empId}/${month}/${year}`)
                .then(response => response.json())
                .then(data => {
                    if (data.success) {
                        document.getElementById('payslipContent').innerHTML = data.html;
                        document.getElementById('downloadPayslipBtn').onclick = function() {
                            window.open(`/download_payslip/${empId}/${month}/${year}`, '_blank');
                        };
                        new bootstrap.Modal(document.getElementById('payslipModal')).show();
                    } else {
                        alert('Error loading payslip: ' + data.message);
                    }
                })
                .catch(error => {
                    console.error('Error:', error);
                    alert('Error loading payslip');
                });
        }

        // Load all payrolls with filtering
        function loadAllPayrolls() {
            const month = document.getElementById('monthFilter').value;
            const year = document.getElementById('yearFilter').value;
            
            let url = '/api/payrolls';
            const params = new URLSearchParams();
            if (month) params.append('month', month);
            if (year) params.append('year', year);
            if (params.toString()) url += '?' + params.toString();
            
            fetch(url)
                .then(response => response.json())
                .then(data => {
                    if (data.success) {
                        updatePayrollTable(data.payrolls);
                    } else {
                        alert('Error loading payrolls: ' + data.message);
                    }
                })
                .catch(error => {
                    console.error('Error:', error);
                    alert('Error loading payrolls');
                });
        }

        function updatePayrollTable(payrolls) {
            const tbody = document.getElementById('payrollTableBody');
            tbody.innerHTML = '';
            
            payrolls.forEach(payroll => {
                const row = `
                    <tr>
                        <td>${payroll.emp_id}</td>
                        <td>${payroll.name}</td>
                        <td>${payroll.month} ${payroll.year}</td>
                        <td>${payroll.days_worked} days</td>
                        <td>₹${parseFloat(payroll.gross_salary).toFixed(2)}</td>
                        <td>₹${parseFloat(payroll.net_salary).toFixed(2)}</td>
                        <td>
                            <button class="btn btn-sm btn-primary-custom" onclick="viewPayslip('${payroll.emp_id}', '${payroll.month}', ${payroll.year})">
                                <i class="fas fa-eye"></i> View
                            </button>
                            <a href="/download_payslip/${payroll.emp_id}/${payroll.month}/${payroll.year}" class="btn btn-sm btn-success-custom">
                                <i class="fas fa-download"></i> PDF
                            </a>
                        </td>
                    </tr>
                `;
                tbody.innerHTML += row;
            });
        }

        // Add event listeners for filters
        document.getElementById('monthFilter').addEventListener('change', loadAllPayrolls);
        document.getElementById('yearFilter').addEventListener('change', loadAllPayrolls);

        // View employee cost breakdown
        function viewEmployeeCostBreakdown(empId, name, ctcMonthly) {
            const breakdown = calculateSalaryBreakdown(ctcMonthly);
            
            const costHtml = `
                <div class="cost-breakdown-container">
                    <div class="employee-header" style="text-align: center; margin-bottom: 30px; padding: 20px; background: linear-gradient(135deg, #2c3e50 0%, #3498db 100%); color: white; border-radius: 8px;">
                        <h3 style="margin: 0;">${name} (${empId})</h3>
                        <p style="margin: 5px 0 0 0; opacity: 0.9;">Monthly CTC: ₹${ctcMonthly.toFixed(2)}</p>
                    </div>
                    
                    <div class="cost-table" style="background: white; border: 1px solid #ddd; border-radius: 8px; overflow: hidden;">
                        <div class="section-header" style="background: #2c3e50; color: white; padding: 15px; text-align: center;">
                            <h4 style="margin: 0;">SALARY STRUCTURE</h4>
                        </div>
                        
                        <table style="width: 100%; border-collapse: collapse;">
                            <thead>
                                <tr style="background: #3498db; color: white;">
                                    <th style="padding: 12px; text-align: left; border: 1px solid #ddd;">Component</th>
                                    <th style="padding: 12px; text-align: right; border: 1px solid #ddd;">Amount (₹)</th>
                                    <th style="padding: 12px; text-align: right; border: 1px solid #ddd;">Percentage</th>
                                </tr>
                            </thead>
                            <tbody>
                                <tr><td style="padding: 10px; border: 1px solid #ddd;">Basic Salary</td><td style="padding: 10px; text-align: right; border: 1px solid #ddd;">₹${breakdown.basic_salary.toFixed(2)}</td><td style="padding: 10px; text-align: right; border: 1px solid #ddd;">40%</td></tr>
                                <tr><td style="padding: 10px; border: 1px solid #ddd;">HRA</td><td style="padding: 10px; text-align: right; border: 1px solid #ddd;">₹${breakdown.hra.toFixed(2)}</td><td style="padding: 10px; text-align: right; border: 1px solid #ddd;">20%</td></tr>
                                <tr><td style="padding: 10px; border: 1px solid #ddd;">Travel Allowance</td><td style="padding: 10px; text-align: right; border: 1px solid #ddd;">₹${breakdown.travel_allowance.toFixed(2)}</td><td style="padding: 10px; text-align: right; border: 1px solid #ddd;">10%</td></tr>
                                <tr><td style="padding: 10px; border: 1px solid #ddd;">Medical Allowance</td><td style="padding: 10px; text-align: right; border: 1px solid #ddd;">₹${breakdown.medical_allowance.toFixed(2)}</td><td style="padding: 10px; text-align: right; border: 1px solid #ddd;">5%</td></tr>
                                <tr><td style="padding: 10px; border: 1px solid #ddd;">LTA</td><td style="padding: 10px; text-align: right; border: 1px solid #ddd;">₹${breakdown.lta.toFixed(2)}</td><td style="padding: 10px; text-align: right; border: 1px solid #ddd;">8%</td></tr>
                                <tr><td style="padding: 10px; border: 1px solid #ddd;">Special Allowance</td><td style="padding: 10px; text-align: right; border: 1px solid #ddd;">₹${breakdown.special_allowance.toFixed(2)}</td><td style="padding: 10px; text-align: right; border: 1px solid #ddd;">17%</td></tr>
                                <tr style="background: #e8f5e8; font-weight: bold;"><td style="padding: 12px; border: 1px solid #ddd;">GROSS SALARY</td><td style="padding: 12px; text-align: right; border: 1px solid #ddd;">₹${breakdown.gross_salary.toFixed(2)}</td><td style="padding: 12px; text-align: right; border: 1px solid #ddd;">100%</td></tr>
                                <tr><td style="padding: 10px; border: 1px solid #ddd; color: #e74c3c; font-weight: bold;">Potential Deductions:</td><td style="padding: 10px; text-align: right; border: 1px solid #ddd;"></td><td style="padding: 10px; text-align: right; border: 1px solid #ddd;"></td></tr>
                                <tr><td style="padding: 10px; border: 1px solid #ddd; padding-left: 30px;">PF Contribution (12%)</td><td style="padding: 10px; text-align: right; border: 1px solid #ddd; color: #e74c3c;">₹${breakdown.pf_deduction.toFixed(2)}</td><td style="padding: 10px; text-align: right; border: 1px solid #ddd;">-12%</td></tr>
                                <tr style="background: #2c3e50; color: white; font-weight: bold; font-size: 1.1em;"><td style="padding: 15px; border: 1px solid #ddd;">POTENTIAL NET SALARY</td><td style="padding: 15px; text-align: right; border: 1px solid #ddd;">₹${breakdown.net_salary.toFixed(2)}</td><td style="padding: 15px; text-align: right; border: 1px solid #ddd;">88%</td></tr>
                            </tbody>
                        </table>
                    </div>
                    
                    <div style="margin-top: 20px; padding: 15px; background: #fff3cd; border: 1px solid #ffeaa7; border-radius: 8px;">
                        <i class="fas fa-info-circle"></i> <strong>Note:</strong> This is the salary structure breakdown. Actual deductions may vary based on PF opt-in, attendance, and other factors.
                    </div>
                </div>
            `;
            
            document.getElementById('employeeCostContent').innerHTML = costHtml;
            new bootstrap.Modal(document.getElementById('employeeCostModal')).show();
        }

        // Download report function
        function downloadReport() {
            const month = document.getElementById('monthFilter').value;
            const year = document.getElementById('yearFilter').value;
            
            if (!month || !year) {
                alert('Please select both month and year to download report');
                return;
            }
            
            window.open(`/download_report/${month}/${year}`, '_blank');
        }

        // Helper function to calculate salary breakdown
        function calculateSalaryBreakdown(ctcMonthly) {
            const basic_salary = ctcMonthly * 0.40;
            const hra = ctcMonthly * 0.20;
            const travel_allowance = ctcMonthly * 0.10;
            const medical_allowance = ctcMonthly * 0.05;
            const lta = ctcMonthly * 0.08;
            const special_allowance = ctcMonthly * 0.17;
            const gross_salary = basic_salary + hra + travel_allowance + medical_allowance + lta + special_allowance;
            const pf_deduction = basic_salary * 0.12;
            const net_salary = gross_salary - pf_deduction;
            
            return {
                basic_salary,
                hra,
                travel_allowance,
                medical_allowance,
                lta,
                special_allowance,
                gross_salary,
                pf_deduction,
                net_salary
            };
        }
    </script>
</body>
</html>
"""


# SQLAlchemy models are imported and tables created via models.py


def calculate_salary_components(ctc_monthly, pf_opted=True, additional_deduction=0.0):
    """Calculate salary components based on new payslip structure"""
    components = {}
    
    # Basic salary (50% of CTC)
    components['basic'] = round(ctc_monthly * 0.50, 2)
    
    # HRA (20% of basic)
    components['hra'] = round(components['basic'] * 0.20, 2)
    
    # New salary components for updated payslip
    components['conveyance_allowance'] = round(ctc_monthly * 0.05, 2)  # 5% of CTC
    components['medical_allowance'] = round(ctc_monthly * 0.05, 2)  # 5% of CTC
    
    # Variable components (can be 0 by default)
    components['overtime_amount'] = 0.0
    components['expenses'] = 0.0
    components['bonus'] = 0.0
    components['leave_balance_amount'] = 0.0
    
    # Calculate special allowance to balance CTC
    fixed_components = (components['basic'] + components['hra'] + 
                       components['conveyance_allowance'] + components['medical_allowance'])
    remaining = ctc_monthly - fixed_components
    components['special_allowance'] = round(remaining, 2)
    
    # Deductions - PF calculation
    if pf_opted:
        # Employee PF (12% of basic, capped at 1800)
        pf_amount = min(components['basic'] * 0.12, 1800)
        components['pf_employee'] = round(pf_amount, 2)
        components['pf_employer'] = round(pf_amount, 2)  # Employer contribution
    else:
        components['pf_employee'] = 0
        components['pf_employer'] = 0
    
    # Other deductions (can be 0 by default)
    components['vpf'] = 0.0
    components['pt'] = 0.0
    components['charity'] = 0.0
    components['misc_deduction'] = 0.0
    components['additional_deduction'] = additional_deduction
    
    # Calculate totals
    components['gross_salary'] = round(
        components['basic'] + components['hra'] + components['special_allowance'] +
        components['conveyance_allowance'] + components['medical_allowance'] + 
        components['overtime_amount'] + components['expenses'] + components['bonus'] +
        components['leave_balance_amount'], 2)
    
    components['total_deductions'] = round(
        components['pf_employee'] + components['vpf'] + components['pt'] + 
        components['charity'] + components['misc_deduction'] + components['additional_deduction'], 2)
    
    components['net_salary'] = round(
        components['gross_salary'] - components['total_deductions'], 2)
    
    return components


def calculate_leave_balance(employee, leaves_taken, month, year):
    """Calculate leave balance and salary adjustments for leave management"""
    from models import LeaveTransaction
    
    # Monthly leave allocation (1.5 days)
    monthly_allocation = 1.5
    
    # Get previous leave transactions for this employee
    prev_transaction = db.session.query(LeaveTransaction).filter_by(
        emp_id=employee.emp_id
    ).order_by(LeaveTransaction.year.desc(), LeaveTransaction.month.desc()).first()
    
    # Calculate current balance
    if prev_transaction:
        current_balance = prev_transaction.balance_after + monthly_allocation
    else:
        current_balance = employee.leave_balance + monthly_allocation
    
    # Determine how much leave can be covered by balance
    leave_balance_used = min(leaves_taken, current_balance)
    loss_of_pay_days = max(0, leaves_taken - current_balance)
    
    # Calculate remaining balance
    remaining_balance = current_balance - leave_balance_used
    
    # Create leave transaction record
    leave_transaction = LeaveTransaction()
    leave_transaction.emp_id = employee.emp_id
    leave_transaction.month = month
    leave_transaction.year = year
    leave_transaction.leaves_allocated = monthly_allocation
    leave_transaction.leaves_used = leave_balance_used
    leave_transaction.balance_before = current_balance - monthly_allocation
    leave_transaction.balance_after = remaining_balance
    
    db.session.add(leave_transaction)
    
    # Update employee's leave balance
    employee.leave_balance = remaining_balance
    
    return {
        'leave_balance_used': leave_balance_used,
        'loss_of_pay_days': loss_of_pay_days,
        'remaining_balance': remaining_balance,
        'paid_days': 30 - loss_of_pay_days,
        'present_days': 30 - leaves_taken
    }


def generate_payslip_pdf(employee_data, payroll_data):
    """Generate PDF payslip for an employee"""
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()

    # Custom styles
    title_style = ParagraphStyle('CustomTitle',
                                 parent=styles['Heading1'],
                                 fontSize=18,
                                 spaceAfter=30,
                                 alignment=TA_CENTER,
                                 textColor=colors.darkblue)

    heading_style = ParagraphStyle('CustomHeading',
                                   parent=styles['Heading2'],
                                   fontSize=14,
                                   spaceAfter=12,
                                   textColor=colors.darkblue)

    content = []

    # Title
    content.append(Paragraph("PAYSLIP", title_style))
    content.append(Spacer(1, 20))

    # Employee Information
    emp_info = [['Employee ID:', employee_data.get('emp_id', getattr(employee_data, 'emp_id', 'N/A'))],
                ['Name:', employee_data.get('name', getattr(employee_data, 'name', 'N/A'))],
                ['Designation:', employee_data.get('designation', getattr(employee_data, 'designation', 'N/A')) or 'N/A'],
                ['Department:', employee_data.get('department', getattr(employee_data, 'department', 'N/A')) or 'N/A'],
                [
                    'Pay Period:',
                    f"{payroll_data.get('month', getattr(payroll_data, 'month', 'N/A'))} {payroll_data.get('year', getattr(payroll_data, 'year', 'N/A'))}"
                ], 
                ['Total Days:', str(payroll_data.get('total_days', getattr(payroll_data, 'total_days', 30)))],
                ['Present Days:', str(payroll_data.get('present_days', getattr(payroll_data, 'present_days', 30)))],
                ['Leaves Taken:', str(payroll_data.get('leaves_taken', getattr(payroll_data, 'leaves_taken', 0)))],
                ['Paid Days:', str(payroll_data.get('paid_days', getattr(payroll_data, 'paid_days', 30)))],
                ['Loss of Pay Days:', str(payroll_data.get('loss_of_pay_days', getattr(payroll_data, 'loss_of_pay_days', 0)))]]

    emp_table = Table(emp_info, colWidths=[2 * inch, 3 * inch])
    emp_table.setStyle(
        TableStyle([('BACKGROUND', (0, 0), (0, -1), colors.lightgrey),
                    ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                    ('FONTSIZE', (0, 0), (-1, -1), 10),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black)]))

    content.append(emp_table)
    content.append(Spacer(1, 20))

    # Salary Breakdown
    content.append(Paragraph("SALARY BREAKDOWN", heading_style))

    # Handle both dict and object access for payroll data
    def get_payroll_value(data, key, default=0):
        if isinstance(data, dict):
            return data.get(key, default)
        else:
            return getattr(data, key, default)

    # Create the exact 5-column table structure matching the email PDF format
    header_data = [['Rate of Salary/ wages', 'Earning', 'Arrear', 'Deductions', 'Attendance/Leave']]
    
    # Salary components rows
    salary_rows = [
        ['BASIC', f"{get_payroll_value(payroll_data, 'basic_salary', 0):.2f}", '0', f"PF: {get_payroll_value(payroll_data, 'pf_employee', 0):.2f}", f"Total Days: {get_payroll_value(payroll_data, 'total_days', 30)}"],
        ['HRA', f"{get_payroll_value(payroll_data, 'hra', 0):.2f}", '0', f"PF(Employer): {get_payroll_value(payroll_data, 'pf_employer', 0):.2f}", f"Present Days: {get_payroll_value(payroll_data, 'present_days', 30)}"],
        ['Special Allowance', f"{get_payroll_value(payroll_data, 'special_allowance', 0):.2f}", '0', 'VPF', f"Leave: {get_payroll_value(payroll_data, 'leaves_taken', 0):.1f}"],
        ['Conveyance Allowance', f"{get_payroll_value(payroll_data, 'conveyance_allowance', 0):.2f}", '0', 'PT', ''],
        ['Medical Allowance', f"{get_payroll_value(payroll_data, 'medical_allowance', 0):.2f}", '0', 'Charity', f"Paid Days: {get_payroll_value(payroll_data, 'paid_days', 30)}"],
        ['Over time', f"{get_payroll_value(payroll_data, 'overtime_amount', 0):.2f}", '0', f"Misc. Deduction: {get_payroll_value(payroll_data, 'pf_employee', 0):.2f}", f"Loss Of Pay: {get_payroll_value(payroll_data, 'loss_of_pay_days', 0):.1f}"],
        ['Expenses', f"{get_payroll_value(payroll_data, 'expenses', 0):.2f}", '0', '', ''],
        ['Bonus', f"{get_payroll_value(payroll_data, 'bonus', 0):.2f}", '0', '', ''],
        ['Leave Balance amount', f"{get_payroll_value(payroll_data, 'leave_balance_amount', 0):.2f}", '', '', '']
    ]
    
    # Totals row
    totals_row = [
        ['Earning', f"{get_payroll_value(payroll_data, 'gross_salary', 0):.2f}", '0', 
         'Deduction', f"{get_payroll_value(payroll_data, 'total_deductions', 0):.2f}"]
    ]
    
    # Combine all data
    all_salary_data = header_data + salary_rows + totals_row
    
    salary_table = Table(all_salary_data, colWidths=[2.2 * inch, 1.1 * inch, 0.7 * inch, 1.3 * inch, 1.2 * inch])
    salary_table.setStyle(
        TableStyle([
            # Header styling
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            # Bold headers
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            # Bold totals row
            ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
            ('BACKGROUND', (0, -1), (-1, -1), colors.lightblue),
            # Align amounts to right
            ('ALIGN', (1, 1), (1, -2), 'RIGHT'),
            ('ALIGN', (2, 1), (2, -2), 'RIGHT'),
            ('ALIGN', (3, 1), (3, -2), 'LEFT'),
            ('ALIGN', (4, 1), (4, -2), 'RIGHT'),
        ]))

    content.append(salary_table)
    content.append(Spacer(1, 20))
    
    # Net Salary section
    def convert_number_to_words(amount):
        # Simple number to words conversion for the payslip
        # This is a basic implementation
        return f"Rupees {amount:.2f}"
    
    net_salary_data = [
        ['Net Salary/Wages(In Words)', ''],
        [f'{convert_number_to_words(get_payroll_value(payroll_data, "net_salary", 0))} Only', f'₹ {get_payroll_value(payroll_data, "net_salary", 0):.2f}']
    ]
    
    net_table = Table(net_salary_data, colWidths=[5 * inch, 2 * inch])
    net_table.setStyle(
        TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('ALIGN', (1, 1), (1, 1), 'RIGHT'),
        ]))

    content.append(net_table)
    content.append(Spacer(1, 20))
    
    # Footer note
    content.append(Paragraph("***This is computer generated salary slip No signature required***", 
                           ParagraphStyle('Footer', parent=styles['Normal'], alignment=TA_CENTER, fontSize=8)))
    
    content.append(Spacer(1, 10))

    # Footer
    footer_text = f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    content.append(Paragraph(footer_text, styles['Normal']))

    doc.build(content)
    buffer.seek(0)
    return buffer


def send_payslip_email(employee_email, employee_name, payslip_pdf, month,
                       year):
    """Send payslip via email"""
    try:
        msg = MIMEMultipart()
        msg['From'] = "stud.studentsmart@gmail.com"
        msg['To'] = employee_email
        msg['Subject'] = f"Payslip for {month} {year}"

        body = f"""
        Dear {employee_name},
        
        Please find attached your payslip for {month} {year}.
        
        If you have any questions regarding your payslip, please contact HR.
        
        Best regards,
        HR Team
        """

        msg.attach(MIMEText(body, 'plain'))

        # Attach PDF
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(payslip_pdf.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition',
                        f'attachment; filename="payslip_{month}_{year}.pdf"')
        msg.attach(part)

        # SMTP configuration
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login("stud.studentsmart@gmail.com", "jygr uhcl odmk flve")
        server.send_message(msg)
        server.quit()

        return True
    except Exception as e:
        logging.error(f"Email sending failed: {str(e)}")
        return False


@app.route('/login', methods=['GET', 'POST'])
def login():
    """Login page"""
    if request.method == 'POST':
        from models import User
        
        username = request.form['username'].strip()
        password = request.form['password'].strip()
        
        # Debug: Check if user exists
        user = db.session.query(User).filter_by(username=username).first()
        app.logger.info(f"Login attempt for username: {username}")
        app.logger.info(f"User found: {user is not None}")
        
        if user:
            app.logger.info(f"User role: {user.role}")
            password_valid = user.check_password(password)
            app.logger.info(f"Password valid: {password_valid}")
            
            if password_valid:
                login_user(user)
                flash(f'Welcome {user.username}! You are logged in as {user.role}.', 'success')
                return redirect(url_for('dashboard'))
        
        # Debug: List all users if login fails
        all_users = db.session.query(User).all()
        app.logger.info(f"All users in database: {[(u.username, u.role) for u in all_users]}")
        flash('Invalid username or password', 'error')
    
    # Login form HTML
    login_html = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PayrollPro - Login</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" rel="stylesheet">
    <style>
        body {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
        }
        .login-card {
            background: white;
            border-radius: 15px;
            box-shadow: 0 15px 35px rgba(0,0,0,0.1);
            padding: 2rem;
            width: 100%;
            max-width: 400px;
        }
        .login-header {
            text-align: center;
            margin-bottom: 2rem;
        }
        .login-header h2 {
            color: #2c3e50;
            margin-bottom: 0.5rem;
        }
        .form-control {
            border-radius: 10px;
            padding: 12px;
            border: 1px solid #ddd;
            margin-bottom: 1rem;
        }
        .btn-login {
            width: 100%;
            padding: 12px;
            border-radius: 10px;
            background: linear-gradient(135deg, #3498db 0%, #2980b9 100%);
            border: none;
            color: white;
            font-weight: 500;
            margin-top: 1rem;
        }
        .btn-login:hover {
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(52, 152, 219, 0.4);
        }
        .user-info {
            background: #f8f9fa;
            border-radius: 10px;
            padding: 1rem;
            margin-top: 1.5rem;
            font-size: 0.9rem;
        }
        .alert {
            border-radius: 10px;
        }
    </style>
</head>
<body>
    <div class="login-card">
        <div class="login-header">
            <i class="fas fa-calculator fa-3x text-primary mb-3"></i>
            <h2>PayrollPro</h2>
            <p class="text-muted">Employee Payroll Management System</p>
        </div>
        
        {% with messages = get_flashed_messages(with_categories=true) %}
            {% if messages %}
                {% for category, message in messages %}
                    <div class="alert alert-{{ 'danger' if category == 'error' else 'success' }} alert-dismissible fade show">
                        {{ message }}
                        <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
                    </div>
                {% endfor %}
            {% endif %}
        {% endwith %}
        
        <form method="POST">
            <div class="mb-3">
                <label for="username" class="form-label">
                    <i class="fas fa-user"></i> Username
                </label>
                <input type="text" class="form-control" id="username" name="username" required>
            </div>
            
            <div class="mb-3">
                <label for="password" class="form-label">
                    <i class="fas fa-lock"></i> Password
                </label>
                <input type="password" class="form-control" id="password" name="password" required>
            </div>
            
            <button type="submit" class="btn btn-login">
                <i class="fas fa-sign-in-alt"></i> Login
            </button>
        </form>
        
        
    </div>
    
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>
    """
    return render_template_string(login_html)


@app.route('/logout')
@login_required
def logout():
    """Logout user"""
    logout_user()
    flash('You have been logged out successfully.', 'success')
    return redirect(url_for('login'))


@app.route('/debug_users')
def debug_users():
    """Debug route to check and recreate users - REMOVE IN PRODUCTION"""
    from models import User
    
    try:
        # List all users
        users = db.session.query(User).all()
        user_info = [(u.username, u.email, u.role) for u in users]
        
        # If no users exist, create them
        if len(users) == 0:
            init_default_users()
            users = db.session.query(User).all()
            user_info = [(u.username, u.email, u.role) for u in users]
        
        return f"""
        <h2>User Debug Information</h2>
        <p>Total users: {len(users)}</p>
        <ul>
        {''.join([f'<li>{username} ({email}) - {role}</li>' for username, email, role in user_info])}
        </ul>
        <p><a href="/login">Go to Login</a></p>
        <p><strong>Note:</strong> Remove this debug route in production!</p>
        """
        
    except Exception as e:
        return f"Error: {str(e)}"


@app.route('/')
@login_required
def dashboard():
    from models import Employee, Payroll
    from sqlalchemy import func
    
    # Data isolation: filter by user unless admin
    if current_user.role == 'admin':
        # Admin sees everything
        employee_filter = True  # No filter for employees
        payroll_filter = True   # No filter for payroll
    else:
        # HR and Accounts only see their own data
        employee_filter = Employee.created_by == current_user.username
        payroll_filter = Payroll.processed_by == current_user.username
    
    # Get dashboard statistics with user isolation
    if current_user.role == 'admin':
        total_employees = db.session.query(Employee).count()
    else:
        total_employees = db.session.query(Employee).filter(employee_filter).count()

    # Get recent payroll entries with user isolation
    if current_user.role == 'admin':
        recent_payroll_query = db.session.query(Payroll, Employee.name).join(
            Employee, Payroll.emp_id == Employee.emp_id
        ).order_by(Payroll.processed_at.desc()).limit(5).all()
    else:
        recent_payroll_query = db.session.query(Payroll, Employee.name).join(
            Employee, Payroll.emp_id == Employee.emp_id
        ).filter(payroll_filter).order_by(Payroll.processed_at.desc()).limit(5).all()

    recent_payroll = []
    for payroll, name in recent_payroll_query:
        recent_payroll.append({
            'emp_id': payroll.emp_id,
            'name': name,
            'month': payroll.month,
            'year': payroll.year,
            'net_salary': payroll.net_salary
        })

    # Get employees with user isolation
    if current_user.role == 'admin':
        employees_query = db.session.query(Employee).order_by(Employee.name).all()
    else:
        employees_query = db.session.query(Employee).filter(employee_filter).order_by(Employee.name).all()
        
    employees = []
    for emp in employees_query:
        employees.append({
            'emp_id': emp.emp_id,
            'name': emp.name,
            'email': emp.email,
            'designation': emp.designation,
            'department': emp.department,
            'ctc_monthly': emp.ctc_monthly,
            'pf_opted': emp.pf_opted,
            'joining_date': emp.joining_date
        })

    # Get all payroll records with user isolation
    if current_user.role == 'admin':
        all_payroll_query = db.session.query(Payroll, Employee.name).join(
            Employee, Payroll.emp_id == Employee.emp_id
        ).order_by(Payroll.processed_at.desc()).limit(20).all()
    else:
        all_payroll_query = db.session.query(Payroll, Employee.name).join(
            Employee, Payroll.emp_id == Employee.emp_id
        ).filter(payroll_filter).order_by(Payroll.processed_at.desc()).limit(20).all()

    # Convert to dictionaries for JSON serialization
    all_payroll = []
    for payroll, name in all_payroll_query:
        all_payroll.append({
            'emp_id': payroll.emp_id,
            'name': name,
            'month': payroll.month,
            'year': payroll.year,
            'leaves_taken': payroll.leaves_taken,
            'total_days': payroll.total_days,
            'present_days': payroll.present_days,
            'paid_days': payroll.paid_days,
            'loss_of_pay_days': payroll.loss_of_pay_days,
            'gross_salary': payroll.gross_salary,
            'net_salary': payroll.net_salary,
            'basic_salary': payroll.basic_salary,
            'hra': payroll.hra,
            'conveyance_allowance': payroll.conveyance_allowance,
            'medical_allowance': payroll.medical_allowance,
            'special_allowance': payroll.special_allowance,
            'overtime_amount': payroll.overtime_amount,
            'expenses': payroll.expenses,
            'bonus': payroll.bonus,
            'pf_employee': payroll.pf_employee,
            'hike_amount': payroll.hike_amount
        })

    # Monthly payroll stats with user isolation
    if current_user.role == 'admin':
        monthly_stats_query = db.session.query(
            Payroll.month,
            Payroll.year,
            func.count().label('count'),
            func.sum(Payroll.net_salary).label('total_payout')
        ).group_by(Payroll.month, Payroll.year).order_by(
            Payroll.year.desc(), Payroll.month.desc()
        ).limit(6).all()
    else:
        monthly_stats_query = db.session.query(
            Payroll.month,
            Payroll.year,
            func.count().label('count'),
            func.sum(Payroll.net_salary).label('total_payout')
        ).filter(payroll_filter).group_by(Payroll.month, Payroll.year).order_by(
            Payroll.year.desc(), Payroll.month.desc()
        ).limit(6).all()

    monthly_stats = []
    for stat in monthly_stats_query:
        monthly_stats.append({
            'month': stat.month,
            'year': stat.year,
            'count': stat.count,
            'total_payout': float(stat.total_payout) if stat.total_payout else 0
        })

    return render_template_string(HTML_TEMPLATE,
                                  total_employees=total_employees,
                                  recent_payroll=recent_payroll,
                                  employees=employees,
                                  all_payroll=all_payroll,
                                  monthly_stats=monthly_stats)


@app.route('/employee_dashboard/<emp_id>')
def employee_dashboard(emp_id):
    """Employee-specific dashboard showing leave balance and usage"""
    from datetime import datetime
    from dateutil import relativedelta
    from models import Employee, Payroll
    
    try:
        employee = db.session.query(Employee).filter_by(emp_id=emp_id).first()
        if not employee:
            return f"<div class='alert alert-danger'>Employee not found</div>"
        
        # Get current month and year
        current_date = datetime.now()
        current_month = current_date.strftime('%B')
        current_year = current_date.year
        
        # Calculate leave balance for current month
        leave_info = calculate_leave_balance(employee, 0, current_month, current_year)
        
        # Get leave usage history for last 6 months
        leave_history = []
        for i in range(6):
            month_date = current_date - relativedelta.relativedelta(months=i)
            month_name = month_date.strftime('%B')
            year = month_date.year
            
            # Get payroll record for this month if exists
            payroll = db.session.query(Payroll).filter_by(
                emp_id=emp_id, 
                month=month_name, 
                year=year
            ).first()
            
            if payroll:
                leave_history.append({
                    'month': month_name,
                    'year': year,
                    'leaves_taken': payroll.leaves_taken,
                    'leave_balance_used': payroll.leave_balance_used,
                    'loss_of_pay_days': payroll.loss_of_pay_days
                })
            else:
                leave_history.append({
                    'month': month_name,
                    'year': year,
                    'leaves_taken': 0,
                    'leave_balance_used': 0,
                    'loss_of_pay_days': 0
                })
        
        # Create employee dashboard HTML
        dashboard_content = f'''
        <div class="container-fluid">
            <div class="row mb-4">
                <div class="col">
                    <h2 class="page-title">
                        <i class="fas fa-user-circle"></i> {employee.name} - Leave Dashboard
                    </h2>
                    <nav aria-label="breadcrumb">
                        <ol class="breadcrumb">
                            <li class="breadcrumb-item"><a href="/">Dashboard</a></li>
                            <li class="breadcrumb-item active">Employee Leave Dashboard</li>
                        </ol>
                    </nav>
                </div>
            </div>

            <!-- Current Leave Status -->
            <div class="row mb-4">
                <div class="col-md-3">
                    <div class="card">
                        <div class="card-body text-center">
                            <div class="stat-icon text-primary mb-2">
                                <i class="fas fa-calendar-check fa-2x"></i>
                            </div>
                            <h5 class="card-title">Available Leave Balance</h5>
                            <h3 class="text-primary">{leave_info['remaining_balance']:.1f} days</h3>
                        </div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="card">
                        <div class="card-body text-center">
                            <div class="stat-icon text-success mb-2">
                                <i class="fas fa-plus-circle fa-2x"></i>
                            </div>
                            <h5 class="card-title">Monthly Allocation</h5>
                            <h3 class="text-success">1.5 days</h3>
                        </div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="card">
                        <div class="card-body text-center">
                            <div class="stat-icon text-info mb-2">
                                <i class="fas fa-calendar-alt fa-2x"></i>
                            </div>
                            <h5 class="card-title">Current Month</h5>
                            <h3 class="text-info">{current_month} {current_year}</h3>
                        </div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="card">
                        <div class="card-body text-center">
                            <div class="stat-icon text-warning mb-2">
                                <i class="fas fa-info-circle fa-2x"></i>
                            </div>
                            <h5 class="card-title">Employee ID</h5>
                            <h3 class="text-warning">{employee.emp_id}</h3>
                        </div>
                    </div>
                </div>
            </div>

            <!-- Leave Usage History -->
            <div class="row">
                <div class="col-12">
                    <div class="card">
                        <div class="card-header">
                            <h5 class="mb-0">
                                <i class="fas fa-history"></i> Leave Usage History (Last 6 Months)
                            </h5>
                        </div>
                        <div class="card-body">
                            <div class="table-responsive">
                                <table class="table table-striped">
                                    <thead>
                                        <tr>
                                            <th>Month</th>
                                            <th>Year</th>
                                            <th>Total Leave Days Taken</th>
                                            <th>Paid Leave Used</th>
                                            <th>Loss of Pay Days</th>
                                            <th>Status</th>
                                        </tr>
                                    </thead>
                                    <tbody>
        '''
        
        for history in leave_history:
            status_badge = "success" if history['loss_of_pay_days'] == 0 else "warning"
            status_text = "Paid Leave" if history['loss_of_pay_days'] == 0 else "Loss of Pay Applied"
            
            dashboard_content += f'''
                                        <tr>
                                            <td>{history['month']}</td>
                                            <td>{history['year']}</td>
                                            <td>{history['leaves_taken']:.1f}</td>
                                            <td>{history['leave_balance_used']:.1f}</td>
                                            <td>{history['loss_of_pay_days']:.1f}</td>
                                            <td><span class="badge bg-{status_badge}">{status_text}</span></td>
                                        </tr>
            '''
        
        dashboard_content += '''
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

            <!-- Leave Policy Information -->
            <div class="row mt-4">
                <div class="col-12">
                    <div class="card">
                        <div class="card-header bg-info text-white">
                            <h5 class="mb-0">
                                <i class="fas fa-info-circle"></i> Leave Policy Information
                            </h5>
                        </div>
                        <div class="card-body">
                            <ul class="list-group list-group-flush">
                                <li class="list-group-item">
                                    <strong>Monthly Allocation:</strong> You get 1.5 leave days every month
                                </li>
                                <li class="list-group-item">
                                    <strong>Carry Forward:</strong> Unused leave days are automatically carried forward to the next month
                                </li>
                                <li class="list-group-item">
                                    <strong>Loss of Pay:</strong> If you take more leave than your available balance, it will be deducted as Loss of Pay
                                </li>
                                <li class="list-group-item">
                                    <strong>Current Balance:</strong> Your current available balance includes carried forward days from previous months
                                </li>
                            </ul>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        '''
        
        return render_template_string(HTML_TEMPLATE,
                                    total_employees=0,
                                    recent_payroll=[],
                                    employees=[],
                                    all_payroll=[],
                                    monthly_stats=[],
                                    title=f"{employee.name} - Leave Dashboard",
                                    content=dashboard_content)
                             
    except Exception as e:
        app.logger.error(f"Employee dashboard error: {str(e)}")
        return f"<div class='alert alert-danger'>Error loading employee dashboard: {str(e)}</div>"


@app.route('/add_employee', methods=['POST'])
@role_required('admin', 'hr')
def add_employee():
    from models import Employee
    from datetime import datetime
    
    try:
        # Validate required fields
        emp_id = request.form.get('emp_id', '').strip()
        name = request.form.get('name', '').strip()
        email = request.form.get('email', '').strip()
        ctc_monthly_str = request.form.get('ctc_monthly', '').strip()
        
        if not emp_id:
            flash('Employee ID is required!', 'error')
            return redirect(url_for('dashboard'))
        
        if not name:
            flash('Employee name is required!', 'error')
            return redirect(url_for('dashboard'))
            
        if not email:
            flash('Email address is required!', 'error')
            return redirect(url_for('dashboard'))
            
        if not ctc_monthly_str:
            flash('Monthly CTC is required!', 'error')
            return redirect(url_for('dashboard'))
        
        # Check if employee ID already exists
        existing_employee = db.session.query(Employee).filter_by(emp_id=emp_id).first()
        if existing_employee:
            flash(f'Employee with ID "{emp_id}" already exists! Please use a different Employee ID.', 'error')
            return redirect(url_for('dashboard'))
        
        # Check if email already exists
        existing_email = db.session.query(Employee).filter_by(email=email).first()
        if existing_email:
            flash(f'Employee with email "{email}" already exists! Please use a different email address.', 'error')
            return redirect(url_for('dashboard'))
        
        # Validate CTC amount
        try:
            ctc_monthly = float(ctc_monthly_str)
            if ctc_monthly <= 0:
                flash('Monthly CTC must be greater than 0!', 'error')
                return redirect(url_for('dashboard'))
        except ValueError:
            flash('Monthly CTC must be a valid number!', 'error')
            return redirect(url_for('dashboard'))
        
        # Validate email format
        import re
        email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        if not re.match(email_pattern, email):
            flash('Please enter a valid email address!', 'error')
            return redirect(url_for('dashboard'))
        
        designation = request.form.get('designation', '').strip()
        department = request.form.get('department', '').strip()
        joining_date_str = request.form.get('joining_date', date.today().isoformat())
        
        # Parse joining date
        try:
            joining_date = datetime.strptime(joining_date_str, '%Y-%m-%d').date()
        except ValueError:
            flash('Invalid joining date format! Please use YYYY-MM-DD format.', 'error')
            return redirect(url_for('dashboard'))
        
        ctc_annual = ctc_monthly * 12
        pf_opted = True if request.form.get('pf_opted') == 'on' else False

        new_employee = Employee()
        new_employee.emp_id = emp_id
        new_employee.name = name
        new_employee.email = email
        new_employee.designation = designation
        new_employee.department = department
        new_employee.joining_date = joining_date
        new_employee.ctc_monthly = ctc_monthly
        new_employee.ctc_annual = ctc_annual
        new_employee.pf_opted = pf_opted
        new_employee.leave_balance = 1.5  # Default monthly leave allocation
        new_employee.created_by = current_user.username  # Track who created this employee
        
        db.session.add(new_employee)
        db.session.commit()

        flash(f'Employee "{name}" (ID: {emp_id}) has been added successfully!', 'success')
    except Exception as e:
        db.session.rollback()
        logging.error(f"Error adding employee: {str(e)}")
        flash('Failed to add employee due to a database error. Please try again or contact administrator.', 'error')

    return redirect(url_for('dashboard'))


@app.route('/download_employee_template')
def download_employee_template():
    """Generate and download employee template Excel file"""
    data = {
        'emp_id': ['EMP001'],
        'name': ['John Doe'],
        'email': ['john@example.com'],
        'designation': ['Developer'],
        'department': ['IT'],
        'joining_date': ['2024-01-01'],
        'ctc_monthly': [50000],
        'pf_opted': ['Yes']
    }

    df = pd.DataFrame(data)

    output = BytesIO()
    df.to_excel(output,
                sheet_name='Employee_Template',
                index=False,
                engine='openpyxl')

    output.seek(0)
    return send_file(
        output,
        mimetype=
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='employee_template.xlsx')


@app.route('/bulk_add_employees', methods=['POST'])
@role_required('admin', 'hr')
def bulk_add_employees():
    from models import Employee
    from datetime import datetime
    
    if 'file' not in request.files:
        flash('No file selected!', 'error')
        return redirect(url_for('dashboard'))

    file = request.files['file']
    if file.filename == '':
        flash('No file selected!', 'error')
        return redirect(url_for('dashboard'))

    try:
        # Check file extension
        filename = file.filename.lower()
        if not (filename.endswith('.xlsx') or filename.endswith('.xls')):
            flash('Please upload an Excel file (.xlsx or .xls format only)!', 'error')
            return redirect(url_for('dashboard'))
        
        df = pd.read_excel(file)
        success_count = 0
        error_count = 0
        error_details = []

        # Validate required columns
        required_columns = ['emp_id', 'name', 'email', 'ctc_monthly']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            flash(f'Missing required columns in Excel file: {", ".join(missing_columns)}', 'error')
            return redirect(url_for('dashboard'))

        for row_num, row in df.iterrows():
            try:
                # Validate required fields
                emp_id = str(row['emp_id']).strip() if pd.notna(row['emp_id']) else ''
                name = str(row['name']).strip() if pd.notna(row['name']) else ''
                email = str(row['email']).strip() if pd.notna(row['email']) else ''
                
                if not emp_id:
                    error_details.append(f"Row {row_num + 2}: Employee ID is missing")
                    error_count += 1
                    continue
                    
                if not name:
                    error_details.append(f"Row {row_num + 2}: Employee name is missing")
                    error_count += 1
                    continue
                    
                if not email:
                    error_details.append(f"Row {row_num + 2}: Email address is missing")
                    error_count += 1
                    continue
                
                # Check for duplicates
                existing_employee = db.session.query(Employee).filter_by(emp_id=emp_id).first()
                if existing_employee:
                    error_details.append(f"Row {row_num + 2}: Employee ID '{emp_id}' already exists")
                    error_count += 1
                    continue
                
                existing_email = db.session.query(Employee).filter_by(email=email).first()
                if existing_email:
                    error_details.append(f"Row {row_num + 2}: Email '{email}' already exists")
                    error_count += 1
                    continue
                
                # Validate CTC
                try:
                    ctc_monthly = float(row['ctc_monthly'])
                    if ctc_monthly <= 0:
                        error_details.append(f"Row {row_num + 2}: Monthly CTC must be greater than 0")
                        error_count += 1
                        continue
                except (ValueError, TypeError):
                    error_details.append(f"Row {row_num + 2}: Invalid CTC amount")
                    error_count += 1
                    continue
                
                # Validate email format
                import re
                email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
                if not re.match(email_pattern, email):
                    error_details.append(f"Row {row_num + 2}: Invalid email format")
                    error_count += 1
                    continue
                
                pf_opted = True if str(row.get('pf_opted', '')).lower() in ['yes', '1', 'true'] else False
                ctc_annual = ctc_monthly * 12
                
                # Parse joining date
                joining_date_str = str(row.get('joining_date', date.today().isoformat()))
                try:
                    if joining_date_str and joining_date_str != 'nan':
                        joining_date = datetime.strptime(joining_date_str, '%Y-%m-%d').date()
                    else:
                        joining_date = date.today()
                except ValueError:
                    joining_date = date.today()

                new_employee = Employee()
                new_employee.emp_id = emp_id
                new_employee.name = name
                new_employee.email = email
                new_employee.designation = str(row.get('designation', ''))
                new_employee.department = str(row.get('department', ''))
                new_employee.joining_date = joining_date
                new_employee.ctc_monthly = ctc_monthly
                new_employee.ctc_annual = ctc_annual
                new_employee.pf_opted = pf_opted
                new_employee.leave_balance = 1.5  # Default monthly leave allocation
                new_employee.created_by = current_user.username  # Track who created this employee
                
                db.session.add(new_employee)
                success_count += 1
            except Exception as e:
                error_details.append(f"Row {row_num + 2}: Unexpected error - {str(e)}")
                error_count += 1
                continue

        if success_count > 0:
            db.session.commit()

        # Provide detailed feedback
        if success_count > 0 and error_count == 0:
            flash(f'Successfully added all {success_count} employees from the Excel file!', 'success')
        elif success_count > 0 and error_count > 0:
            flash(f'Successfully added {success_count} employees. {error_count} rows had errors.', 'warning')
            if error_details:
                flash(f'Error details: {"; ".join(error_details[:5])}{"..." if len(error_details) > 5 else ""}', 'info')
        else:
            flash(f'No employees were added. All {error_count} rows had errors.', 'error')
            if error_details:
                flash(f'Error details: {"; ".join(error_details[:5])}{"..." if len(error_details) > 5 else ""}', 'error')
                
    except Exception as e:
        db.session.rollback()
        logging.error(f"Error processing Excel file: {str(e)}")
        if "No sheet named" in str(e):
            flash('Excel file format error. Please ensure the file has the correct structure.', 'error')
        elif "Missing column" in str(e):
            flash('Excel file is missing required columns. Please download the template and use the correct format.', 'error')
        else:
            flash('Failed to process Excel file. Please check the file format and try again.', 'error')

    return redirect(url_for('dashboard'))


@app.route('/download_payroll_template')
def download_payroll_template():
    """Generate and download payroll template Excel file"""
    data = {
        'emp_id': ['EMP001'],
        'name': ['John Doe'],
        'leaves_taken': [2.0],
        'pf_opted': ['Yes']
    }

    df = pd.DataFrame(data)

    output = BytesIO()
    df.to_excel(output,
                sheet_name='Payroll_Template',
                index=False,
                engine='openpyxl')

    output.seek(0)
    return send_file(
        output,
        mimetype=
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='payroll_template.xlsx')


@app.route('/process_individual_payroll', methods=['POST'])
@role_required('admin', 'hr', 'accounts')
def process_individual_payroll():
    from models import Employee, Payroll
    
    try:
        emp_id = request.form['emp_id']
        leaves_taken = float(request.form.get('leaves_taken', 0))
        month = request.form['month']
        year = int(request.form['year'])
        pf_opted = True if request.form.get('pf_opted') == 'on' else False
        hike_amount = float(request.form.get('hike_amount', 0))
        deduction_amount = float(request.form.get('deduction_amount', 0))
        deduction_reason = request.form.get('deduction_reason', '')
        
        # Optional additional components
        overtime_amount = float(request.form.get('overtime_amount', 0))
        expenses = float(request.form.get('expenses', 0))
        bonus = float(request.form.get('bonus', 0))

        # Get employee data
        employee = db.session.query(Employee).filter_by(emp_id=emp_id).first()
        if not employee:
            flash('Employee not found!', 'error')
            return redirect(url_for('dashboard'))

        # Apply hike if specified
        ctc_monthly = employee.ctc_monthly + hike_amount

        # Calculate salary components
        salary_components = calculate_salary_components(ctc_monthly, pf_opted, deduction_amount)
        
        # Calculate leave balance and attendance
        leave_info = calculate_leave_balance(employee, leaves_taken, month, year)
        
        # Adjust salary for loss of pay days
        if leave_info['loss_of_pay_days'] > 0:
            loss_ratio = leave_info['loss_of_pay_days'] / 30.0
            for key in ['basic', 'hra', 'special_allowance', 'conveyance_allowance', 'medical_allowance']:
                salary_components[key] = round(salary_components[key] * (1 - loss_ratio), 2)
        
        # Add variable components
        salary_components['overtime_amount'] = overtime_amount
        salary_components['expenses'] = expenses
        salary_components['bonus'] = bonus
        salary_components['leave_balance_amount'] = leave_info['remaining_balance'] * 100  # Monetize leave balance
        
        # Recalculate totals with adjustments
        salary_components['gross_salary'] = round(
            salary_components['basic'] + salary_components['hra'] + salary_components['special_allowance'] +
            salary_components['conveyance_allowance'] + salary_components['medical_allowance'] + 
            salary_components['overtime_amount'] + salary_components['expenses'] + salary_components['bonus'] +
            salary_components['leave_balance_amount'], 2)
        
        salary_components['net_salary'] = round(
            salary_components['gross_salary'] - salary_components['total_deductions'], 2)

        # Check if payroll already exists for this employee/month/year
        existing = db.session.query(Payroll).filter_by(
            emp_id=emp_id, month=month, year=year
        ).first()

        if existing:
            # Update existing payroll
            existing.total_days = 30
            existing.present_days = int(leave_info['present_days'])
            existing.leaves_taken = leaves_taken
            existing.paid_days = int(leave_info['paid_days'])
            existing.loss_of_pay_days = leave_info['loss_of_pay_days']
            existing.leave_balance_used = leave_info['leave_balance_used']
            
            existing.basic_salary = salary_components['basic']
            existing.hra = salary_components['hra']
            existing.special_allowance = salary_components['special_allowance']
            existing.conveyance_allowance = salary_components['conveyance_allowance']
            existing.medical_allowance = salary_components['medical_allowance']
            existing.overtime_amount = salary_components['overtime_amount']
            existing.expenses = salary_components['expenses']
            existing.bonus = salary_components['bonus']
            existing.leave_balance_amount = salary_components['leave_balance_amount']
            
            existing.pf_employee = salary_components['pf_employee']
            existing.pf_employer = salary_components['pf_employer']
            existing.vpf = salary_components['vpf']
            existing.pt = salary_components['pt']
            existing.charity = salary_components['charity'] 
            existing.misc_deduction = salary_components['misc_deduction']
            existing.additional_deduction = salary_components['additional_deduction']
            existing.deduction_reason = deduction_reason
            existing.total_deductions = salary_components['total_deductions']
            existing.gross_salary = salary_components['gross_salary']
            existing.net_salary = salary_components['net_salary']
            existing.hike_amount = hike_amount
            existing.processed_by = current_user.username  # Track who processed this payroll
            existing.processed_at = datetime.utcnow()
        else:
            # Insert new payroll
            new_payroll = Payroll()
            new_payroll.emp_id = emp_id
            new_payroll.month = month
            new_payroll.year = year
            new_payroll.total_days = 30
            new_payroll.present_days = int(leave_info['present_days'])
            new_payroll.leaves_taken = leaves_taken
            new_payroll.paid_days = int(leave_info['paid_days'])
            new_payroll.loss_of_pay_days = leave_info['loss_of_pay_days']
            new_payroll.leave_balance_used = leave_info['leave_balance_used']
            
            new_payroll.basic_salary = salary_components['basic']
            new_payroll.hra = salary_components['hra']
            new_payroll.special_allowance = salary_components['special_allowance']
            new_payroll.conveyance_allowance = salary_components['conveyance_allowance']
            new_payroll.medical_allowance = salary_components['medical_allowance']
            new_payroll.overtime_amount = salary_components['overtime_amount']
            new_payroll.expenses = salary_components['expenses']
            new_payroll.bonus = salary_components['bonus']
            new_payroll.leave_balance_amount = salary_components['leave_balance_amount']
            
            new_payroll.pf_employee = salary_components['pf_employee']
            new_payroll.pf_employer = salary_components['pf_employer']
            new_payroll.vpf = salary_components['vpf']
            new_payroll.pt = salary_components['pt']
            new_payroll.charity = salary_components['charity']
            new_payroll.misc_deduction = salary_components['misc_deduction']
            new_payroll.additional_deduction = salary_components['additional_deduction']
            new_payroll.deduction_reason = deduction_reason
            new_payroll.total_deductions = salary_components['total_deductions']
            new_payroll.gross_salary = salary_components['gross_salary']
            new_payroll.net_salary = salary_components['net_salary']
            new_payroll.hike_amount = hike_amount
            new_payroll.processed_by = current_user.username  # Track who processed this payroll
            new_payroll.processed_at = datetime.utcnow()
            db.session.add(new_payroll)

        # Note: Hike amount is only applied for this month's calculation
        # Employee's base CTC remains unchanged for future months

        db.session.commit()

        flash(f'Payroll processed successfully for {employee.name}! Leave balance: {leave_info["remaining_balance"]:.1f} days', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Error processing payroll: {str(e)}', 'error')

    return redirect(url_for('dashboard'))


@app.route('/bulk_process_payroll', methods=['POST'])
@role_required('admin', 'hr', 'accounts')
def bulk_process_payroll():
    from models import Employee, Payroll
    
    if 'file' not in request.files:
        flash('No file selected!', 'error')
        return redirect(url_for('dashboard'))

    file = request.files['file']
    month = request.form['month']
    year = int(request.form['year'])

    if file.filename == '':
        flash('No file selected!', 'error')
        return redirect(url_for('dashboard'))

    try:
        df = pd.read_excel(file)
        success_count = 0
        error_count = 0

        for _, row in df.iterrows():
            try:
                emp_id = str(row['emp_id'])
                days_worked = int(float(row.get('days_worked', 30)))
                pf_opted = str(row.get('pf_opted', 'Yes')).lower() in ['yes', '1', 'true']

                # Get employee data
                employee = db.session.query(Employee).filter_by(emp_id=emp_id).first()
                if not employee:
                    error_count += 1
                    continue

                # Calculate salary components
                salary_components = calculate_salary_components(employee.ctc_monthly, pf_opted)

                # Adjust for days worked
                if days_worked != 30:
                    ratio = days_worked / 30.0
                    for key in salary_components:
                        if key != 'employer_pf':  # Employer PF remains constant
                            salary_components[key] = round(
                                salary_components[key] * ratio, 2)

                # Check if payroll already exists
                existing = db.session.query(Payroll).filter_by(
                    emp_id=emp_id, month=month, year=year
                ).first()

                if existing:
                    # Update existing payroll with new schema
                    existing.total_days = days_worked
                    existing.present_days = days_worked
                    existing.leaves_taken = 0
                    existing.paid_days = days_worked
                    existing.loss_of_pay_days = 30 - days_worked if days_worked < 30 else 0
                    existing.leave_balance_used = 0
                    existing.basic_salary = salary_components['basic']
                    existing.hra = salary_components['hra']
                    existing.conveyance_allowance = salary_components['conveyance_allowance']
                    existing.medical_allowance = salary_components['medical_allowance']
                    existing.special_allowance = salary_components['special_allowance']
                    existing.overtime_amount = salary_components['overtime_amount']
                    existing.expenses = salary_components['expenses']
                    existing.bonus = salary_components['bonus']
                    existing.leave_balance_amount = salary_components['leave_balance_amount']
                    existing.pf_employer = salary_components['employer_pf']
                    existing.pf_employee = salary_components['employee_pf']
                    existing.gross_salary = salary_components['gross_salary']
                    existing.total_deductions = salary_components['total_deductions']
                    existing.net_salary = salary_components['net_salary']
                    existing.processed_by = current_user.username  # Track who processed this payroll
                    existing.processed_at = datetime.utcnow()
                else:
                    # Insert new payroll with new schema
                    new_payroll = Payroll()
                    new_payroll.emp_id = emp_id
                    new_payroll.month = month
                    new_payroll.year = year
                    new_payroll.total_days = days_worked
                    new_payroll.present_days = days_worked
                    new_payroll.leaves_taken = 0
                    new_payroll.paid_days = days_worked
                    new_payroll.loss_of_pay_days = 30 - days_worked if days_worked < 30 else 0
                    new_payroll.leave_balance_used = 0
                    new_payroll.basic_salary = salary_components['basic']
                    new_payroll.hra = salary_components['hra']
                    new_payroll.conveyance_allowance = salary_components['conveyance_allowance']
                    new_payroll.medical_allowance = salary_components['medical_allowance']
                    new_payroll.special_allowance = salary_components['special_allowance']
                    new_payroll.overtime_amount = salary_components['overtime_amount']
                    new_payroll.expenses = salary_components['expenses']
                    new_payroll.bonus = salary_components['bonus']
                    new_payroll.leave_balance_amount = salary_components['leave_balance_amount']
                    new_payroll.pf_employer = salary_components['employer_pf']
                    new_payroll.pf_employee = salary_components['employee_pf']
                    new_payroll.gross_salary = salary_components['gross_salary']
                    new_payroll.total_deductions = salary_components['total_deductions']
                    new_payroll.net_salary = salary_components['net_salary']
                    new_payroll.hike_amount = 0
                    new_payroll.processed_by = current_user.username  # Track who processed this payroll
                    new_payroll.processed_at = datetime.utcnow()
                    db.session.add(new_payroll)

                success_count += 1
            except Exception as e:
                logging.error(
                    f"Error processing payroll for {row.get('emp_id', 'Unknown')}: {str(e)}"
                )
                error_count += 1
                continue

        db.session.commit()

        flash(
            f'Successfully processed {success_count} payrolls. {error_count} errors.',
            'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Error processing file: {str(e)}', 'error')

    return redirect(url_for('dashboard'))


@app.route('/apply_hike', methods=['POST'])
@role_required('admin')
def apply_hike():
    from models import Employee
    
    try:
        emp_id = request.form['emp_id']
        hike_amount = float(request.form['hike_amount'])
        hike_reason = request.form.get('hike_reason', '')

        # Get employee data
        employee = db.session.query(Employee).filter_by(emp_id=emp_id).first()
        if not employee:
            flash('Employee not found!', 'error')
            return redirect(url_for('dashboard'))

        # Update employee CTC
        employee.ctc_monthly = employee.ctc_monthly + hike_amount
        employee.ctc_annual = employee.ctc_monthly * 12

        db.session.commit()

        flash(
            f'Salary hike of ₹{hike_amount:,.2f} applied successfully for {employee.name}!',
            'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Error applying hike: {str(e)}', 'error')

    return redirect(url_for('dashboard'))


@app.route('/send_payslips', methods=['POST'])
@role_required('admin', 'hr', 'accounts')
def send_payslips():
    try:
        month = request.form['month']
        year = int(request.form['year'])

        from models import Employee, Payroll
        
        # Get payroll records for the specified month/year with data isolation
        if current_user.role == 'admin':
            # Admin sees all payroll records
            payroll_records = db.session.query(Payroll, Employee).join(
                Employee, Payroll.emp_id == Employee.emp_id
            ).filter(Payroll.month == month, Payroll.year == year).all()
        else:
            # HR and Accounts users only see payroll they processed
            payroll_records = db.session.query(Payroll, Employee).join(
                Employee, Payroll.emp_id == Employee.emp_id
            ).filter(
                Payroll.month == month, 
                Payroll.year == year,
                Payroll.processed_by == current_user.username
            ).all()

        if not payroll_records:
            flash(f'No payroll records found for {month} {year}!', 'error')
            return redirect(url_for('dashboard'))

        success_count = 0
        error_count = 0

        for payroll, employee in payroll_records:
            try:
                # Generate payslip PDF with new format data
                record_data = {
                    'emp_id': payroll.emp_id,
                    'name': employee.name,
                    'designation': employee.designation,
                    'department': employee.department,
                    'month': payroll.month,
                    'year': payroll.year,
                    'basic_salary': payroll.basic_salary,
                    'hra': payroll.hra,
                    'special_allowance': payroll.special_allowance,
                    'conveyance_allowance': payroll.conveyance_allowance,
                    'medical_allowance': payroll.medical_allowance,
                    'overtime_amount': payroll.overtime_amount,
                    'expenses': payroll.expenses,
                    'bonus': payroll.bonus,
                    'leave_balance_amount': payroll.leave_balance_amount,
                    'pf_employee': payroll.pf_employee,
                    'gross_salary': payroll.gross_salary,
                    'total_deductions': payroll.total_deductions,
                    'net_salary': payroll.net_salary,
                    'total_days': payroll.total_days,
                    'present_days': payroll.present_days,
                    'leaves_taken': payroll.leaves_taken,
                    'paid_days': payroll.paid_days,
                    'loss_of_pay_days': payroll.loss_of_pay_days
                }

                payslip_pdf = generate_payslip_pdf(record_data, record_data)

                # Send email
                if send_payslip_email(employee.email, employee.name, payslip_pdf, month, year):
                    success_count += 1
                else:
                    error_count += 1

            except Exception as e:
                logging.error(f"Error sending payslip to {employee.name}: {str(e)}")
                error_count += 1
                continue

        flash(
            f'Payslips sent successfully to {success_count} employees. {error_count} failed.',
            'success')
    except Exception as e:
        flash(f'Error sending payslips: {str(e)}', 'error')

    return redirect(url_for('dashboard'))


@app.route('/api/payslip/<emp_id>/<month>/<int:year>')
def api_payslip(emp_id, month, year):
    """API endpoint to get payslip data as HTML"""
    try:
        from models import Employee, Payroll
        
        # Get payroll record with data isolation
        if current_user.role == 'admin':
            # Admin sees all payroll records
            payroll_data = db.session.query(Payroll, Employee).join(
                Employee, Payroll.emp_id == Employee.emp_id
            ).filter(
                Payroll.emp_id == emp_id, 
                Payroll.month == month, 
                Payroll.year == year
            ).first()
        else:
            # HR and Accounts users only see payroll they processed
            payroll_data = db.session.query(Payroll, Employee).join(
                Employee, Payroll.emp_id == Employee.emp_id
            ).filter(
                Payroll.emp_id == emp_id, 
                Payroll.month == month, 
                Payroll.year == year,
                Payroll.processed_by == current_user.username
            ).first()
        
        if not payroll_data:
            return jsonify({
                'success': False,
                'message': 'Payroll record not found'
            })
        
        payroll, employee = payroll_data

        if not payroll:
            return jsonify({
                'success': False,
                'message': 'Payroll record not found'
            })

        # Generate HTML payslip with the 5-column table structure from the email PDF
        payslip_html = f"""
        <div class="payslip-container" style="max-width: 1000px; margin: 0 auto; font-family: Arial, sans-serif;">
            <div class="payslip-header" style="text-align: center; margin-bottom: 30px; padding: 20px; background: linear-gradient(135deg, #2c3e50 0%, #3498db 100%); color: white; border-radius: 8px;">
                <h2 style="margin: 0; font-size: 2rem;">PAYSLIP</h2>
                <p style="margin: 5px 0 0 0; opacity: 0.9;">Salary Statement for {month} {year}</p>
            </div>
            
            <div class="employee-info" style="background: #f8f9fa; padding: 20px; border-radius: 8px; margin-bottom: 20px;">
                <div class="row" style="display: flex; flex-wrap: wrap;">
                    <div style="flex: 1; min-width: 300px; margin-bottom: 10px;">
                        <table style="width: 100%; border-collapse: collapse;">
                            <tr><td style="padding: 8px 0; font-weight: bold; width: 40%;">Employee ID:</td><td style="padding: 8px 0;">{payroll.emp_id}</td></tr>
                            <tr><td style="padding: 8px 0; font-weight: bold;">Name:</td><td style="padding: 8px 0;">{employee.name}</td></tr>
                            <tr><td style="padding: 8px 0; font-weight: bold;">Designation:</td><td style="padding: 8px 0;">{employee.designation or 'N/A'}</td></tr>
                            <tr><td style="padding: 8px 0; font-weight: bold;">Department:</td><td style="padding: 8px 0;">{employee.department or 'N/A'}</td></tr>
                        </table>
                    </div>
                </div>
            </div>
            
            <div class="salary-breakdown" style="background: white; border: 1px solid #ddd; border-radius: 8px; overflow: hidden;">
                <table style="width: 100%; border-collapse: collapse;">
                    <thead>
                        <tr style="background: #f8f9fa;">
                            <th style="padding: 12px; text-align: center; border: 1px solid #000; font-weight: bold;">Rate of Salary/ wages</th>
                            <th style="padding: 12px; text-align: center; border: 1px solid #000; font-weight: bold;">Earning</th>
                            <th style="padding: 12px; text-align: center; border: 1px solid #000; font-weight: bold;">Arrear</th>
                            <th style="padding: 12px; text-align: center; border: 1px solid #000; font-weight: bold;">Deductions</th>
                            <th style="padding: 12px; text-align: center; border: 1px solid #000; font-weight: bold;">Attendance/Leave</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr><td style="padding: 8px; border: 1px solid #000;">BASIC</td><td style="padding: 8px; text-align: right; border: 1px solid #000;">{payroll.basic_salary:,.2f}</td><td style="padding: 8px; text-align: center; border: 1px solid #000;">0</td><td style="padding: 8px; border: 1px solid #000;">PF</td><td style="padding: 8px; border: 1px solid #000;">Total Days: {payroll.total_days}</td></tr>
                        <tr><td style="padding: 8px; border: 1px solid #000;">HRA</td><td style="padding: 8px; text-align: right; border: 1px solid #000;">{payroll.hra:,.2f}</td><td style="padding: 8px; text-align: center; border: 1px solid #000;">0</td><td style="padding: 8px; border: 1px solid #000;">PF(Employer): {payroll.pf_employer:,.2f}</td><td style="padding: 8px; border: 1px solid #000;">Present Days: {payroll.present_days}</td></tr>
                        <tr><td style="padding: 8px; border: 1px solid #000;">Special Allowance</td><td style="padding: 8px; text-align: right; border: 1px solid #000;">{payroll.special_allowance:,.2f}</td><td style="padding: 8px; text-align: center; border: 1px solid #000;">0</td><td style="padding: 8px; border: 1px solid #000;">VPF</td><td style="padding: 8px; border: 1px solid #000;">Leave: {payroll.leaves_taken}</td></tr>
                        <tr><td style="padding: 8px; border: 1px solid #000;">Conveyance Allowance</td><td style="padding: 8px; text-align: right; border: 1px solid #000;">{payroll.conveyance_allowance:,.2f}</td><td style="padding: 8px; text-align: center; border: 1px solid #000;">0</td><td style="padding: 8px; border: 1px solid #000;">PT</td><td style="padding: 8px; border: 1px solid #000;"></td></tr>
                        <tr><td style="padding: 8px; border: 1px solid #000;">Medical Allowance</td><td style="padding: 8px; text-align: right; border: 1px solid #000;">{payroll.medical_allowance:,.2f}</td><td style="padding: 8px; text-align: center; border: 1px solid #000;">0</td><td style="padding: 8px; border: 1px solid #000;">Charity</td><td style="padding: 8px; border: 1px solid #000;">Paid Days: {payroll.paid_days}</td></tr>
                        <tr><td style="padding: 8px; border: 1px solid #000;">Over Time</td><td style="padding: 8px; text-align: right; border: 1px solid #000;">{payroll.overtime_amount:,.2f}</td><td style="padding: 8px; text-align: center; border: 1px solid #000;">0</td><td style="padding: 8px; border: 1px solid #000;">Misc. Deduction: {payroll.pf_employee:,.2f}</td><td style="padding: 8px; border: 1px solid #000;">Loss Of Pay: {payroll.loss_of_pay_days}</td></tr>
                        <tr><td style="padding: 8px; border: 1px solid #000;">Expenses</td><td style="padding: 8px; text-align: right; border: 1px solid #000;">{payroll.expenses:,.2f}</td><td style="padding: 8px; text-align: center; border: 1px solid #000;">0</td><td style="padding: 8px; border: 1px solid #000;"></td><td style="padding: 8px; border: 1px solid #000;"></td></tr>
                        <tr><td style="padding: 8px; border: 1px solid #000;">Bonus</td><td style="padding: 8px; text-align: right; border: 1px solid #000;">{payroll.bonus:,.2f}</td><td style="padding: 8px; text-align: center; border: 1px solid #000;">0</td><td style="padding: 8px; border: 1px solid #000;"></td><td style="padding: 8px; border: 1px solid #000;"></td></tr>
                        <tr><td style="padding: 8px; border: 1px solid #000;">Leave Balance amount</td><td style="padding: 8px; text-align: right; border: 1px solid #000;">{payroll.leave_balance_amount:,.2f}</td><td style="padding: 8px; text-align: center; border: 1px solid #000;"></td><td style="padding: 8px; border: 1px solid #000;"></td><td style="padding: 8px; border: 1px solid #000;"></td></tr>
                        
                        <tr style="background: #e8f5e8; font-weight: bold;">
                            <td style="padding: 10px; border: 1px solid #000;">Earning</td>
                            <td style="padding: 10px; text-align: right; border: 1px solid #000;">{payroll.gross_salary:,.2f}</td>
                            <td style="padding: 10px; text-align: center; border: 1px solid #000;">0</td>
                            <td style="padding: 10px; border: 1px solid #000;">Deduction</td>
                            <td style="padding: 10px; text-align: right; border: 1px solid #000;">{payroll.total_deductions:,.2f}</td>
                        </tr>
                    </tbody>
                </table>
            </div>
            
            <div style="margin-top: 20px; background: white; border: 1px solid #ddd; border-radius: 8px; overflow: hidden;">
                <table style="width: 100%; border-collapse: collapse;">
                    <tr style="background: #f8f9fa; font-weight: bold;">
                        <td style="padding: 12px; border: 1px solid #000;">Net Salary/Wages(In Words)</td>
                        <td style="padding: 12px; text-align: right; border: 1px solid #000;">₹{payroll.net_salary:,.2f}</td>
                    </tr>
                </table>
            </div>
            
            {f'<div style="margin-top: 20px; padding: 15px; background: #fff3cd; border: 1px solid #ffeaa7; border-radius: 8px;"><i class="fas fa-info-circle"></i> <strong>Salary Hike Applied:</strong> ₹{payroll.hike_amount:,.2f} added to monthly CTC</div>' if payroll.hike_amount > 0 else ''}
            
            <div class="footer" style="margin-top: 30px; text-align: center; color: #666; font-size: 0.9em;">
                <p>***This is computer generated salary slip No signature required***</p>
                <p>Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
            </div>
        </div>
        """

        return jsonify({
            'success': True,
            'html': payslip_html,
            'emp_id': emp_id,
            'month': month,
            'year': year
        })

    except Exception as e:
        logging.error(f"Error generating payslip HTML: {str(e)}")
        return jsonify({'success': False, 'message': str(e)})


@app.route('/api/payrolls')
def api_payrolls():
    """API endpoint to get filtered payroll records"""
    try:
        month = request.args.get('month')
        year = request.args.get('year')

        from models import Employee, Payroll
        
        # Build query with filters and data isolation
        if current_user.role == 'admin':
            # Admin sees all payroll records
            query = db.session.query(Payroll, Employee).join(
                Employee, Payroll.emp_id == Employee.emp_id
            )
        else:
            # HR and Accounts users only see payroll they processed
            query = db.session.query(Payroll, Employee).join(
                Employee, Payroll.emp_id == Employee.emp_id
            ).filter(Payroll.processed_by == current_user.username)

        if month:
            query = query.filter(Payroll.month == month)

        if year:
            query = query.filter(Payroll.year == int(year))

        payrolls = query.order_by(Payroll.processed_at.desc()).limit(50).all()

        # Convert to list of dictionaries
        payroll_list = []
        for payroll, employee in payrolls:
            payroll_list.append({
                'emp_id': payroll.emp_id,
                'name': employee.name,
                'month': payroll.month,
                'year': payroll.year,
                'leaves_taken': payroll.leaves_taken,
                'total_days': payroll.total_days,
                'present_days': payroll.present_days,
                'paid_days': payroll.paid_days,
                'loss_of_pay_days': payroll.loss_of_pay_days,
                'gross_salary': payroll.gross_salary,
                'net_salary': payroll.net_salary,
                'basic_salary': payroll.basic_salary,
                'hra': payroll.hra,
                'special_allowance': payroll.special_allowance,
                'conveyance_allowance': payroll.conveyance_allowance,
                'medical_allowance': payroll.medical_allowance,
                'overtime_amount': payroll.overtime_amount,
                'expenses': payroll.expenses,
                'bonus': payroll.bonus,
                'pf_employee': payroll.pf_employee,
                'hike_amount': payroll.hike_amount
            })

        return jsonify({'success': True, 'payrolls': payroll_list})

    except Exception as e:
        logging.error(f"Error fetching payrolls: {str(e)}")
        return jsonify({'success': False, 'message': str(e)})


@app.route('/download_payslip/<emp_id>/<month>/<int:year>')
def download_payslip(emp_id, month, year):
    """Download payslip as PDF"""
    try:
        from models import Employee, Payroll

        # Get employee and payroll data with data isolation
        if current_user.role == 'admin':
            # Admin sees all data
            employee = db.session.query(Employee).filter_by(emp_id=emp_id).first()
            payroll = db.session.query(Payroll).filter_by(
                emp_id=emp_id, month=month, year=year
            ).first()
        else:
            # HR and Accounts users only see their own data
            employee = db.session.query(Employee).filter_by(
                emp_id=emp_id, created_by=current_user.username
            ).first()
            payroll = db.session.query(Payroll).filter_by(
                emp_id=emp_id, month=month, year=year, processed_by=current_user.username
            ).first()

        if not employee or not payroll:
            flash('Employee or payroll record not found!', 'error')
            return redirect(url_for('dashboard'))

        # Convert SQLAlchemy objects to dictionaries for PDF generation
        employee_dict = {
            'emp_id': employee.emp_id,
            'name': employee.name,
            'email': employee.email,
            'designation': employee.designation,
            'department': employee.department,
            'joining_date': employee.joining_date,
            'ctc_monthly': employee.ctc_monthly,
            'ctc_annual': employee.ctc_annual,
            'pf_opted': employee.pf_opted,
            'leave_balance': employee.leave_balance
        }
        
        payroll_dict = {
            'emp_id': payroll.emp_id,
            'month': payroll.month,
            'year': payroll.year,
            'leaves_taken': payroll.leaves_taken,
            'total_days': payroll.total_days,
            'present_days': payroll.present_days,
            'paid_days': payroll.paid_days,
            'loss_of_pay_days': payroll.loss_of_pay_days,
            'basic_salary': payroll.basic_salary,
            'hra': payroll.hra,
            'conveyance_allowance': payroll.conveyance_allowance,
            'medical_allowance': payroll.medical_allowance,
            'special_allowance': payroll.special_allowance,
            'overtime_amount': payroll.overtime_amount,
            'expenses': payroll.expenses,
            'bonus': payroll.bonus,
            'leave_balance_amount': payroll.leave_balance_amount,
            'gross_salary': payroll.gross_salary,
            'pf_employee': payroll.pf_employee,
            'pf_employer': payroll.pf_employer,
            'total_deductions': payroll.total_deductions,
            'net_salary': payroll.net_salary,
            'hike_amount': payroll.hike_amount
        }
        
        # Generate PDF
        payslip_pdf = generate_payslip_pdf(employee_dict, payroll_dict)

        return send_file(payslip_pdf,
                         mimetype='application/pdf',
                         as_attachment=True,
                         download_name=f'payslip_{emp_id}_{month}_{year}.pdf')

    except Exception as e:
        flash(f'Error generating payslip: {str(e)}', 'error')
        return redirect(url_for('dashboard'))


@app.route('/download_report/<month>/<int:year>')
@role_required('admin', 'hr', 'accounts')
def download_report(month, year):
    """Download complete payroll report for selected month/year as Excel"""
    try:
        from models import Employee, Payroll

        # Get payroll records for the month/year with data isolation
        if current_user.role == 'admin':
            # Admin sees all payroll records
            payrolls = db.session.query(Payroll, Employee).join(
                Employee, Payroll.emp_id == Employee.emp_id
            ).filter(Payroll.month == month, Payroll.year == year).order_by(Employee.name).all()
        else:
            # HR and Accounts users only see payroll they processed
            payrolls = db.session.query(Payroll, Employee).join(
                Employee, Payroll.emp_id == Employee.emp_id
            ).filter(
                Payroll.month == month, 
                Payroll.year == year,
                Payroll.processed_by == current_user.username
            ).order_by(Employee.name).all()

        if not payrolls:
            flash(f'No payroll records found for {month} {year}!', 'error')
            return redirect(url_for('dashboard'))

        # Create Excel workbook
        import pandas as pd
        from io import BytesIO

        # Convert payroll data to list of dictionaries
        data = []
        for payroll, employee in payrolls:
            data.append({
                'Employee ID': payroll.emp_id,
                'Name': employee.name,
                'Email': employee.email,
                'Designation': employee.designation or 'N/A',
                'Department': employee.department or 'N/A',
                'Month': payroll.month,
                'Year': payroll.year,
                'Total Days': payroll.total_days,
                'Present Days': payroll.present_days,
                'Leaves Taken': payroll.leaves_taken,
                'Paid Days': payroll.paid_days,
                'Loss of Pay Days': payroll.loss_of_pay_days,
                'Basic Salary': payroll.basic_salary,
                'HRA': payroll.hra,
                'Conveyance Allowance': payroll.conveyance_allowance,
                'Medical Allowance': payroll.medical_allowance,
                'Special Allowance': payroll.special_allowance,
                'Overtime Amount': payroll.overtime_amount,
                'Expenses': payroll.expenses,
                'Bonus': payroll.bonus,
                'Leave Balance Amount': payroll.leave_balance_amount,
                'Gross Salary': payroll.gross_salary,
                'PF Employee': payroll.pf_employee,
                'Total Deductions': payroll.total_deductions,
                'Net Salary': payroll.net_salary,
                'Hike Amount': payroll.hike_amount,
                'Processed Date': payroll.processed_at
            })

        # Create DataFrame
        df = pd.DataFrame(data)

        # Create Excel file in memory
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer,
                        sheet_name=f'{month}_{year}_Payroll',
                        index=False)

            # Add summary sheet
            summary_data = {
                'Metric': [
                    'Total Employees', 'Total Gross Salary',
                    'Total PF Deduction', 'Total Net Salary',
                    'Average Gross Salary', 'Average Net Salary'
                ],
                'Value': [
                    len(payrolls), f"₹{df['Gross Salary'].sum():.2f}",
                    f"₹{df['Total Deductions'].sum():.2f}",
                    f"₹{df['Net Salary'].sum():.2f}",
                    f"₹{df['Gross Salary'].mean():.2f}",
                    f"₹{df['Net Salary'].mean():.2f}"
                ]
            }
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Summary', index=False)

        output.seek(0)

        return send_file(
            output,
            mimetype=
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'payroll_report_{month}_{year}.xlsx')

    except Exception as e:
        logging.error(f"Error generating report: {str(e)}")
        flash(f'Error generating report: {str(e)}', 'error')
        return redirect(url_for('dashboard'))


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
