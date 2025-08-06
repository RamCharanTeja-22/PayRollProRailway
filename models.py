from app import db
from datetime import datetime
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash


class User(UserMixin, db.Model):
    __tablename__ = 'users'
    
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(64), unique=True, nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)
    password_hash = db.Column(db.String(256))
    role = db.Column(db.String(20), nullable=False, default='hr')  # admin, hr, accounts
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    
    def set_password(self, password):
        self.password_hash = generate_password_hash(password)
        
    def check_password(self, password):
        return check_password_hash(self.password_hash, password)
    
    def __repr__(self):
        return f'<User {self.username}>'


class Employee(db.Model):
    __tablename__ = 'employees'
    
    id = db.Column(db.Integer, primary_key=True)
    emp_id = db.Column(db.String(50), unique=True, nullable=False)
    name = db.Column(db.String(100), nullable=False)
    email = db.Column(db.String(120), nullable=False)
    designation = db.Column(db.String(100))
    department = db.Column(db.String(100))
    joining_date = db.Column(db.Date)
    ctc_monthly = db.Column(db.Float, nullable=False)
    ctc_annual = db.Column(db.Float, nullable=False)
    pf_opted = db.Column(db.Boolean, default=True)
    leave_balance = db.Column(db.Float, default=0.0)  # Accumulated paid leave balance
    created_by = db.Column(db.String(64), nullable=False)  # Track which user created this employee
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    
    # Relationship
    payrolls = db.relationship('Payroll', backref='employee', lazy=True)
    leave_transactions = db.relationship('LeaveTransaction', backref='employee', lazy=True)
    
    def __repr__(self):
        return f'<Employee {self.emp_id}: {self.name}>'


class Payroll(db.Model):
    __tablename__ = 'payroll'
    
    id = db.Column(db.Integer, primary_key=True)
    emp_id = db.Column(db.String(50), db.ForeignKey('employees.emp_id'), nullable=False)
    month = db.Column(db.String(20), nullable=False)
    year = db.Column(db.Integer, nullable=False)
    
    # Attendance/Leave Information
    total_days = db.Column(db.Integer, default=30)
    present_days = db.Column(db.Integer, default=30)
    leaves_taken = db.Column(db.Float, default=0.0)
    paid_days = db.Column(db.Integer, default=30)
    loss_of_pay_days = db.Column(db.Float, default=0.0)
    leave_balance_used = db.Column(db.Float, default=0.0)
    
    # Salary Components - Earnings
    basic_salary = db.Column(db.Float)
    hra = db.Column(db.Float)
    special_allowance = db.Column(db.Float)
    conveyance_allowance = db.Column(db.Float)
    medical_allowance = db.Column(db.Float)
    overtime_amount = db.Column(db.Float, default=0.0)
    expenses = db.Column(db.Float, default=0.0)
    bonus = db.Column(db.Float, default=0.0)
    leave_balance_amount = db.Column(db.Float, default=0.0)
    
    # Deductions
    pf_employee = db.Column(db.Float)
    pf_employer = db.Column(db.Float)
    vpf = db.Column(db.Float, default=0.0)
    pt = db.Column(db.Float, default=0.0)
    charity = db.Column(db.Float, default=0.0)
    misc_deduction = db.Column(db.Float, default=0.0)
    additional_deduction = db.Column(db.Float, default=0.0)  # New field for month-specific deductions
    deduction_reason = db.Column(db.String(255), default='')  # Reason for additional deduction
    
    # Totals
    gross_salary = db.Column(db.Float)
    total_deductions = db.Column(db.Float)
    net_salary = db.Column(db.Float)
    
    hike_amount = db.Column(db.Float, default=0)
    processed_by = db.Column(db.String(64), nullable=False)  # Track which user processed this payroll
    processed_at = db.Column(db.DateTime, default=datetime.utcnow)
    
    def __repr__(self):
        return f'<Payroll {self.emp_id} - {self.month} {self.year}>'


class LeaveTransaction(db.Model):
    __tablename__ = 'leave_transactions'
    
    id = db.Column(db.Integer, primary_key=True)
    emp_id = db.Column(db.String(50), db.ForeignKey('employees.emp_id'), nullable=False)
    month = db.Column(db.String(20), nullable=False)
    year = db.Column(db.Integer, nullable=False)
    leaves_allocated = db.Column(db.Float, default=1.5)  # Monthly allocation
    leaves_used = db.Column(db.Float, default=0.0)
    balance_before = db.Column(db.Float, default=0.0)
    balance_after = db.Column(db.Float, default=0.0)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    
    def __repr__(self):
        return f'<LeaveTransaction {self.emp_id} - {self.month} {self.year}>'