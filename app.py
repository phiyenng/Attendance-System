from flask import Flask, render_template, request, jsonify, send_file
from markupsafe import Markup
import pandas as pd
import openpyxl
from datetime import datetime, timedelta, date
import os
import json
from werkzeug.utils import secure_filename
import numpy as np
from io import BytesIO
import re

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your-secret-key-here'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Ensure upload folder exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Global variables to store data
sign_in_out_data = None
apply_data = None
ot_lieu_data = None
attendance_report = None
abnormal_data = None
emp_list = None

# In-memory employee list
EMPLOYEE_LIST_PATH = os.path.join(app.config['UPLOAD_FOLDER'], 'employee_list.csv')

# Load employee list from file if exists
if os.path.exists(EMPLOYEE_LIST_PATH):
    try:
        employee_list_df = pd.read_csv(EMPLOYEE_LIST_PATH)
    except Exception:
        employee_list_df = pd.DataFrame(columns=["STT", "Name", "ID Number", "Dept", "Internship", "14 digits ID"])
else:
    employee_list_df = pd.DataFrame(columns=["STT", "Name", "ID Number", "Dept", "Internship", "14 digits ID"])

# Helper for header translation
APPLY_HEADER_MAP = {
    '申请人工号': 'Emp ID',
    '申请人': 'Emp Name',
    '起始时间': 'Start Date',
    '终止时间': 'End Date',
    '申请说明': 'Note',
    '审批结果': 'Approve Result',
}

# Helper for leave type logic
def map_leave_type(note):
    if pd.isna(note):
        return ''
    note = str(note)
    if re.search(r'事假|Leave|unpaid', note, re.I):
        return 'Unpaid'
    if re.search(r'年休假|Annual', note, re.I):
        return 'Annual'
    if re.search(r'产|婚|育|丧|welfare', note, re.I):
        return 'Welfare'
    if re.search(r'病假|sick', note, re.I):
        return 'Sick'
    return note

def translate_apply_headers(df):
    print('Original columns:', list(df.columns))  # DEBUG
    def normalize(col):
        # Loại bỏ dấu ', ", khoảng trắng, tab, xuống dòng ở đầu/cuối và bên trong tên cột
        return re.sub(r"[ '\"]+", '', str(col)).strip()
    norm_map = {normalize(k): v for k, v in APPLY_HEADER_MAP.items()}
    new_cols = []
    for col in df.columns:
        ncol = normalize(col)
        new_cols.append(norm_map.get(ncol, col))
    print('Normalized columns:', new_cols)  # DEBUG
    df.columns = new_cols
    return df

def filter_apply_employees(df, emp_list):
    if 'Emp Name' not in df.columns:
        return df
    emp_list = emp_list.copy()  # Tránh SettingWithCopyWarning
    emp_list['EmpNameID'] = emp_list['Name'].astype(str) + emp_list['ID Number'].astype(str)
    valid_names = set(emp_list['EmpNameID'])
    return df[df['Emp Name'].astype(str).isin(valid_names)]

def add_apply_columns(df):
    # Type
    if 'Application Type' in df.columns:
        def extract_type(val):
            if not isinstance(val, str):
                return ''
            val_lower = val.lower()
            if 'trip' in val_lower or 'trips' in val_lower:
                return 'Trip'
            if 'leave' in val_lower:
                return 'Leave'
            #if 'supp' in val_lower:
                #return 'Supp'
            if 'replenishment' in val_lower:
                return 'Replenishment'
            return ''
        df['Type'] = df['Application Type'].apply(extract_type)
    else:
        df['Type'] = ''
    # Results
    if 'Approve Result' in df.columns:
        def map_approve_result(x):
            if pd.notna(x):
                x_str = str(x)
                if '通过' in x_str:
                    return 'Approved'
                elif '待审批' in x_str:
                    return 'Pending'
                elif '已撤销' in x_str:
                    return 'Withdrawal'
            return x
        df['Results'] = df['Approve Result'].apply(map_approve_result)
    else:
        df['Results'] = ''
    # Leave Type
    if 'Note' in df.columns:
        df['Leave Type'] = df['Note'].apply(map_leave_type)
    else:
        df['Leave Type'] = ''
    return df

def load_excel_data(file_path, sheet_name=None):
    """Load data from Excel file"""
    try:
        workbook = openpyxl.load_workbook(file_path, data_only=True)
        
        if sheet_name:
            if sheet_name in workbook.sheetnames:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
            else:
                return None
        else:
            # Use first sheet if no specific sheet name
            df = pd.read_excel(file_path, sheet_name=workbook.sheetnames[0])
        
        # Clean the data
        df = df.dropna(how='all')
        df = df.fillna('')
        
        return df
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        return None

def save_excel_data(file_path, data, sheet_name):
    """Save data back to Excel file"""
    try:
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
            data.to_excel(writer, sheet_name=sheet_name, index=False)
        return True
    except Exception as e:
        print(f"Error saving Excel file: {e}")
        return False

def check_apply(employee_name, check_date, shift, apply_type, row, col):
    """Python implementation of CheckApply VBA function"""
    global apply_data
    
    if apply_data is None or apply_data.empty:
        return False
    
    approved = "Approved"
    
    # Adjust check date based on shift
    if shift == 0:  # Morning shift
        check_date = check_date + timedelta(hours=9, minutes=30, seconds=1)
    elif shift == 1:  # Evening shift
        check_date = check_date + timedelta(hours=17, minutes=29, seconds=59)
    
    is_approved = False
    
    for _, row_data in apply_data.iterrows():
        if (employee_name in str(row_data.get('B', '')) and 
            employee_name != "" and employee_name != "0"):
            
            if apply_type in str(row_data.get('K', '')):
                
                if apply_type == "Supp":
                    start_time = row_data.get('F')
                    end_time = row_data.get('G')
                    
                    if (shift == 0 and start_time <= check_date and 
                        start_time.date() == check_date.date()) or \
                       (shift == 1 and start_time.date() == check_date.date() and 
                        check_date <= end_time):
                        
                        if approved in str(row_data.get('L', '')):
                            is_approved = True
                            break
                
                else:  # Trip, Leave
                    start_time = row_data.get('F')
                    end_time = row_data.get('G')
                    
                    if start_time <= check_date <= end_time:
                        if approved in str(row_data.get('L', '')):
                            is_approved = True
                            break
    
    return is_approved

def check_lieu(employee_name, check_date, shift, col, row):
    """Python implementation of checkLieu VBA function"""
    global ot_lieu_data, attendance_report
    
    if ot_lieu_data is None or ot_lieu_data.empty:
        return False
    
    is_approved = False
    
    for _, row_data in ot_lieu_data.iterrows():
        if (employee_name in str(row_data.get('A', '')) and 
            employee_name != "" and employee_name != "0"):
            
            lieu_date = row_data.get('K')
            if lieu_date == check_date:
                
                lieu_from = row_data.get('L')
                lieu_to = row_data.get('M')
                lieu_sum = row_data.get('N', 0)
                
                if shift == 0:  # Morning
                    if pd.isna(lieu_from) or (isinstance(lieu_from, (int, float)) and lieu_from * 24 <= 12):
                        is_approved = True
                else:  # Evening
                    if pd.isna(lieu_to) or (isinstance(lieu_to, (int, float)) and lieu_to * 24 >= 13.5):
                        is_approved = True
                
                break
    
    return is_approved

def is_holiday(check_date):
    """Check if date is a holiday"""
    # This would need to be implemented based on your holiday rules
    # For now, returning False
    return False

def is_special_work_day(check_date):
    """Check if date is a special work day"""
    # This would need to be implemented based on your rules
    # For now, returning False
    return False

def get_day_type(check_date):
    """Get day type (Weekday/Weekend)"""
    if is_special_work_day(check_date):
        return "Weekday"
    
    weekday = check_date.weekday()
    if weekday < 5:  # Monday to Friday
        return "Weekday"
    else:
        return "Weekend"

def calculate_attendance():
    """Main calculation function - equivalent to RecalculateWorkbook VBA macro"""
    global sign_in_out_data, apply_data, ot_lieu_data, attendance_report, abnormal_data
    
    if sign_in_out_data is None or sign_in_out_data.empty:
        return {"error": "No data available"}
    
    # Clear previous calculations
    if abnormal_data is not None:
        abnormal_data = abnormal_data.iloc[0:0]  # Clear all rows
    
    # Initialize abnormal data if needed
    if abnormal_data is None:
        abnormal_data = pd.DataFrame(columns=['Employee', 'Date', 'SignIn', 'SignOut', 'Status', 'LateMinutes'])
    
    # Use new column names
    emp_col = None
    time_col = None
    for col in sign_in_out_data.columns:
        if col.lower() == 'emp_name':
            emp_col = col
        if col.lower() == 'attendance_time':
            time_col = col
    if not emp_col or not time_col:
        return {"error": "Sign In/Out data must have 'emp_name' and 'attendance_time' columns"}
    
    employees = sign_in_out_data[emp_col].unique()
    
    for emp in employees:
        if pd.isna(emp) or emp == "":
            continue
        emp_data = sign_in_out_data[sign_in_out_data[emp_col] == emp]
        for _, row in emp_data.iterrows():
            sign_time = row[time_col]
            if pd.isna(sign_time):
                continue
            # Convert to datetime if needed
            if not isinstance(sign_time, pd.Timestamp):
                try:
                    sign_time = pd.to_datetime(sign_time)
                except:
                    continue
            check_date = sign_time.date() if hasattr(sign_time, 'date') else sign_time
            # Check for late/early
            shift = 0 if sign_time.hour < 12 else 1
            # Check if employee has approved leave/apply
            has_apply = False  # TODO: integrate with apply_data if needed
            has_lieu = False   # TODO: integrate with ot_lieu_data if needed
            # Calculate if late/early
            if not has_apply and not has_lieu:
                day_type = get_day_type(check_date)
                if day_type == "Weekday" and not is_holiday(check_date):
                    if shift == 0:  # Morning
                        expected_time = sign_time.replace(hour=8, minute=30, second=0)
                        if sign_time > expected_time:
                            late_minutes = int((sign_time - expected_time).total_seconds() / 60)
                            abnormal_data = pd.concat([abnormal_data, pd.DataFrame([{
                                'Employee': emp,
                                'Date': check_date,
                                'SignIn': sign_time,
                                'SignOut': '',
                                'Status': 'Late',
                                'LateMinutes': late_minutes
                            }])], ignore_index=True)
                    elif shift == 1:  # Evening
                        expected_time = sign_time.replace(hour=17, minute=30, second=0)
                        if sign_time < expected_time:
                            early_minutes = int((expected_time - sign_time).total_seconds() / 60)
                            abnormal_data = pd.concat([abnormal_data, pd.DataFrame([{
                                'Employee': emp,
                                'Date': check_date,
                                'SignIn': '',
                                'SignOut': sign_time,
                                'Status': 'Early Leave',
                                'LateMinutes': early_minutes
                            }])], ignore_index=True)
    return {"success": True, "message": "Attendance calculation completed"}

@app.route('/')
def index():
    """Main page"""
    return render_template('index.html')

@app.route('/import/signinout', methods=['POST'])
def import_signinout():
    """Import Sign In/Out data (only emp_name and attendance_time columns)"""
    global sign_in_out_data
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    if file and file.filename.lower().endswith(('.xlsx', '.xls', '.csv')):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        # Load data
        if file.filename.lower().endswith('.csv'):
            df = try_read_csv(open(file_path, 'rb').read())
        else:
            df = pd.read_excel(file_path)
        # Only keep emp_name and attendance_time
        keep = [col for col in df.columns if col.lower() in ['emp_name', 'attendance_time']]
        df = df[keep]
        sign_in_out_data = df.reset_index(drop=True)
        return jsonify({
            'success': True,
            'message': f'Sign In/Out data imported successfully. Loaded {len(df)} rows.',
            'rows': len(df)
        })
    else:
        return jsonify({'error': 'Invalid file type'}), 400

@app.route('/import/apply', methods=['POST'])
def import_apply():
    """Import Apply data with header translation, employee filter, and extra columns"""
    global apply_data, employee_list_df
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    if file and file.filename.lower().endswith(('.xlsx', '.xls', '.csv')):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        # Load data
        if file.filename.lower().endswith('.csv'):
            df = try_read_csv(open(file_path, 'rb').read())
        else:
            df = pd.read_excel(file_path)
        df = translate_apply_headers(df)
        df = filter_apply_employees(df, employee_list_df)
        df = add_apply_columns(df)
        apply_data = df.reset_index(drop=True)
        return jsonify({
            'success': True,
            'message': f'Apply data imported successfully. Loaded {len(df)} rows.',
            'rows': len(df)
        })
    else:
        return jsonify({'error': 'Invalid file type'}), 400

@app.route('/import/otlieu', methods=['POST'])
def import_otlieu():
    """Import OT Lieu data"""
    global ot_lieu_data
    
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    if file and file.filename.lower().endswith(('.xlsx', '.xls', '.csv')):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        
        # Load data
        data = load_excel_data(file_path)
        
        if data is not None:
            ot_lieu_data = data
            return jsonify({
                'success': True,
                'message': f'OT Lieu data imported successfully. Loaded {len(data)} rows.',
                'rows': len(data)
            })
        else:
            return jsonify({'error': 'Failed to load file'}), 400
    else:
        return jsonify({'error': 'Invalid file type'}), 400

@app.route('/refresh', methods=['POST'])
def refresh():
    """Refresh/Recalculate all data"""
    result = calculate_attendance()
    return jsonify(result)

@app.route('/calculate_abnormal', methods=['POST'])
def calculate_abnormal():
    """Calculate abnormal attendance"""
    global abnormal_data
    
    if sign_in_out_data is None or sign_in_out_data.empty:
        return jsonify({'error': 'No sign in/out data available'}), 400
    
    # Clear previous abnormal data
    abnormal_data = pd.DataFrame(columns=['Employee', 'Date', 'SignIn', 'SignOut', 'Status', 'LateMinutes'])
    
    # Calculate attendance
    result = calculate_attendance()
    
    if result.get('success'):
        return jsonify({
            'success': True,
            'message': f'Abnormal calculation completed. Found {len(abnormal_data)} records.',
            'abnormal_count': len(abnormal_data)
        })
    else:
        return jsonify(result), 400

@app.route('/get_abnormal_data')
def get_abnormal_data():
    """Get abnormal data for display"""
    global abnormal_data
    
    if abnormal_data is None or abnormal_data.empty:
        return jsonify({'data': [], 'message': 'No abnormal data available'})
    
    # Convert to list for JSON serialization
    data_list = []
    for _, row in abnormal_data.iterrows():
        data_list.append({
            'Employee': str(row['Employee']),
            'Date': str(row['Date']),
            'SignIn': str(row['SignIn']) if pd.notna(row['SignIn']) else '',
            'SignOut': str(row['SignOut']) if pd.notna(row['SignOut']) else '',
            'Status': str(row['Status']),
            'LateMinutes': int(row['LateMinutes']) if pd.notna(row['LateMinutes']) else 0
        })
    
    return jsonify({
        'data': data_list,
        'count': len(data_list)
    })

@app.route('/clear/signinout', methods=['POST'])
def clear_signinout():
    """Clear Sign In/Out data"""
    global sign_in_out_data
    sign_in_out_data = None
    return jsonify({'success': True, 'message': 'Sign In/Out data cleared successfully'})

@app.route('/clear/apply', methods=['POST'])
def clear_apply():
    """Clear Apply data"""
    global apply_data
    apply_data = None
    return jsonify({'success': True, 'message': 'Apply data cleared successfully'})

@app.route('/clear/otlieu', methods=['POST'])
def clear_otlieu():
    """Clear OT Lieu data"""
    global ot_lieu_data
    ot_lieu_data = None
    return jsonify({'success': True, 'message': 'OT Lieu data cleared successfully'})

@app.route('/export', methods=['GET'])
def export():
    """Export processed data as Excel file"""
    global sign_in_out_data, apply_data, ot_lieu_data, abnormal_data
    
    try:
        # Create a new Excel file
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"attendance_report_{timestamp}.xlsx"
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            # Export each data type to different sheets
            if sign_in_out_data is not None and not sign_in_out_data.empty:
                sign_in_out_data.to_excel(writer, sheet_name='Sign In-Out Data', index=False)
            
            if apply_data is not None and not apply_data.empty:
                apply_data.to_excel(writer, sheet_name='Apply Data', index=False)
            
            if ot_lieu_data is not None and not ot_lieu_data.empty:
                ot_lieu_data.to_excel(writer, sheet_name='OT Lieu Data', index=False)
            
            if abnormal_data is not None and not abnormal_data.empty:
                abnormal_data.to_excel(writer, sheet_name='Abnormal Data', index=False)
        
        return send_file(file_path, as_attachment=True, download_name=filename)
    
    except Exception as e:
        return jsonify({'error': f'Failed to export file: {str(e)}'}), 400

@app.route('/get_data_status')
def get_data_status():
    """Get status of loaded data"""
    status = {
        'sign_in_out': len(sign_in_out_data) if sign_in_out_data is not None else 0,
        'apply': len(apply_data) if apply_data is not None else 0,
        'ot_lieu': len(ot_lieu_data) if ot_lieu_data is not None else 0,
        'abnormal': len(abnormal_data) if abnormal_data is not None else 0
    }
    return jsonify(status)

def try_read_csv(file_bytes, **kwargs):
    encodings = ['utf-8', 'utf-8-sig', 'cp1252', 'latin1', 'gbk']
    for enc in encodings:
        try:
            return pd.read_csv(BytesIO(file_bytes), encoding=enc, **kwargs)
        except Exception:
            continue
    raise ValueError('Could not decode CSV file with common encodings.')

@app.route('/preview_upload', methods=['POST'])
def preview_upload():
    """Preview uploaded file: return sheet names and preview data for each sheet"""
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    filename = secure_filename(file.filename)
    ext = os.path.splitext(filename)[1].lower()
    file_bytes = file.read()
    preview = {}
    sheet_names = []
    try:
        if ext in ['.xlsx', '.xls']:
            excel = pd.ExcelFile(BytesIO(file_bytes))
            sheet_names = excel.sheet_names
            for sheet in sheet_names:
                df = pd.read_excel(excel, sheet_name=sheet, nrows=10)
                preview[sheet] = {
                    'columns': df.columns.tolist(),
                    'rows': df.fillna('').astype(str).values.tolist()
                }
        elif ext == '.csv':
            df = try_read_csv(file_bytes, nrows=10)
            preview['CSV'] = {
                'columns': df.columns.tolist(),
                'rows': df.fillna('').astype(str).values.tolist()
            }
            sheet_names = ['CSV']
        else:
            return jsonify({'error': 'Unsupported file type'}), 400
        return jsonify({'success': True, 'sheet_names': sheet_names, 'preview': preview})
    except Exception as e:
        return jsonify({'error': f'Failed to preview file: {str(e)}'}), 400

@app.route('/import_with_sheet', methods=['POST'])
def import_with_sheet():
    """Import a specific sheet from an uploaded file for a given data type"""
    global sign_in_out_data, apply_data, ot_lieu_data, employee_list_df
    data_type = request.form.get('data_type')
    sheet_name = request.form.get('sheet_name')
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    filename = secure_filename(file.filename)
    ext = os.path.splitext(filename)[1].lower()
    file_bytes = file.read()
    try:
        if ext in ['.xlsx', '.xls']:
            df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name)
        elif ext == '.csv':
            df = try_read_csv(file_bytes)
        else:
            return jsonify({'error': 'Unsupported file type'}), 400
        df = df.dropna(how='all').fillna('')
        if data_type == 'signinout':
            sign_in_out_data = df
        elif data_type == 'apply':
            df = translate_apply_headers(df)
            df = filter_apply_employees(df, employee_list_df)
            df = add_apply_columns(df)
            apply_data = df.reset_index(drop=True)
        elif data_type == 'otlieu':
            ot_lieu_data = df
        else:
            return jsonify({'error': 'Invalid data type'}), 400
        return jsonify({'success': True, 'message': f'{data_type.capitalize()} data imported successfully. Loaded {len(df)} rows.', 'rows': len(df)})
    except Exception as e:
        print(str(e))
        return jsonify({'error': f'Failed to import file: {str(e)}'}), 400

@app.route('/employee_list', methods=['GET'])
def employee_list():
    # Return the employee list as an HTML table with delete buttons
    global employee_list_df
    df = employee_list_df.copy()
    if df.empty:
        return Markup('<div class="text-muted">No employees loaded.</div>')
    html = '<table class="table table-bordered table-sm"><thead><tr>'
    for col in df.columns:
        html += f'<th>{col}</th>'
    html += '<th>Action</th></tr></thead><tbody>'
    for idx, row in df.iterrows():
        html += '<tr>'
        for col in df.columns:
            html += f'<td>{row[col]}</td>'
        html += f'<td><button class="btn btn-sm btn-danger" onclick="removeEmployee({idx})">Delete</button></td>'
        html += '</tr>'
    html += '</tbody></table>'
    return Markup(html)

@app.route('/upload_employee_list', methods=['POST'])
def upload_employee_list():
    global employee_list_df
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    ext = os.path.splitext(file.filename)[1].lower()
    file_bytes = file.read()
    try:
        if ext in ['.xlsx', '.xls']:
            # Đọc tất cả các cột là object (chuỗi)
            df = pd.read_excel(BytesIO(file_bytes), dtype=str)
        elif ext == '.csv':
            df = try_read_csv(file_bytes, dtype=str)
        else:
            return jsonify({'error': 'Unsupported file type'}), 400
        # Only keep required columns
        keep = ["STT", "Name", "ID Number", "Dept", "Internship", "14 digits ID"]
        df = df[[col for col in keep if col in df.columns]]
        df = df.fillna('')
        employee_list_df = df.reset_index(drop=True)
        # Save to file
        employee_list_df.to_csv(EMPLOYEE_LIST_PATH, index=False)
        return jsonify({'success': True, 'message': f'Imported {len(df)} employees.'})
    except Exception as e:
        return jsonify({'error': f'Failed to import employee list: {str(e)}'}), 400

@app.route('/add_employee', methods=['POST'])
def add_employee():
    global employee_list_df
    try:
        data = request.json
        columns = list(employee_list_df.columns)
        row = [''] * len(columns)
        for idx, col in enumerate(columns):
            if col == 'Name':
                row[idx] = data.get('Name', '')
            elif col == 'ID Number':
                row[idx] = data.get('ID Number', '')
            elif col == 'Dept':
                row[idx] = data.get('Dept', '')
        employee_list_df.loc[len(employee_list_df)] = row
        # Auto fill STT ascending if exists
        if 'STT' in employee_list_df.columns:
            employee_list_df['STT'] = range(1, len(employee_list_df) + 1)
        employee_list_df.to_csv(EMPLOYEE_LIST_PATH, index=False)
        return jsonify({'success': True, 'message': 'Employee added.'})
    except Exception as e:
        return jsonify({'error': f'Failed to add employee: {str(e)}'}), 400

@app.route('/remove_employee', methods=['POST'])
def remove_employee():
    global employee_list_df
    try:
        idx = int(request.json.get('index'))
        employee_list_df = employee_list_df.drop(idx).reset_index(drop=True)
        # Auto fill STT ascending if exists
        if 'STT' in employee_list_df.columns:
            employee_list_df['STT'] = range(1, len(employee_list_df) + 1)
        employee_list_df.to_csv(EMPLOYEE_LIST_PATH, index=False)
        return jsonify({'success': True, 'message': 'Employee removed.'})
    except Exception as e:
        return jsonify({'error': f'Failed to remove employee: {str(e)}'}), 400

@app.route('/calculate_prep_data', methods=['POST'])
def calculate_prep_data():
    # Placeholder: just return success
    return jsonify({'success': True, 'message': 'Prep data calculation completed (placeholder).'})

@app.route('/ot_lieu_report')
def ot_lieu_report():
    # Placeholder: return a simple OT & Lieu Report table
    html = '''<table class="table table-bordered table-sm">
        <thead><tr><th>Employee</th><th>OT Hours</th><th>Lieu Hours</th></tr></thead>
        <tbody>
            <tr><td>John Doe</td><td>10</td><td>2</td></tr>
            <tr><td>Jane Smith</td><td>8</td><td>1</td></tr>
        </tbody>
    </table>'''
    return Markup(html)

@app.route('/lieu_followup')
def lieu_followup():
    # Placeholder: return a simple Lieu Followup table
    html = '''<table class="table table-bordered table-sm">
        <thead><tr><th>Employee</th><th>Remaining Lieu</th></tr></thead>
        <tbody>
            <tr><td>John Doe</td><td>1</td></tr>
            <tr><td>Jane Smith</td><td>0.5</td></tr>
        </tbody>
    </table>'''
    return Markup(html)

@app.route('/total_attendance_detail')
def total_attendance_detail():
    # Placeholder: return a simple Total Attendance Detail table
    html = '''<table class="table table-bordered table-sm">
        <thead><tr><th>Employee</th><th>Days Present</th><th>Days Absent</th></tr></thead>
        <tbody>
            <tr><td>John Doe</td><td>20</td><td>1</td></tr>
            <tr><td>Jane Smith</td><td>19</td><td>2</td></tr>
        </tbody>
    </table>'''
    return Markup(html)

@app.route('/attendance_report')
def attendance_report():
    # Placeholder: return a simple Attendance Report table
    html = '''<table class="table table-bordered table-sm">
        <thead><tr><th>Employee</th><th>Date</th><th>Status</th></tr></thead>
        <tbody>
            <tr><td>John Doe</td><td>2024-07-01</td><td>Present</td></tr>
            <tr><td>Jane Smith</td><td>2024-07-01</td><td>Absent</td></tr>
        </tbody>
    </table>'''
    return Markup(html)

@app.route('/get_signinout_data')
def get_signinout_data():
    global sign_in_out_data
    if sign_in_out_data is None or sign_in_out_data.empty:
        return jsonify({'columns': [], 'data': []})
    cols = list(sign_in_out_data.columns)
    rows = sign_in_out_data.fillna('').astype(str).values.tolist()
    return jsonify({'columns': cols, 'data': rows})

@app.route('/get_apply_data')
def get_apply_data():
    global apply_data
    if apply_data is None or apply_data.empty:
        return jsonify({'columns': [], 'data': []})
    cols = list(apply_data.columns)
    rows = apply_data.fillna('').astype(str).values.tolist()
    return jsonify({'columns': cols, 'data': rows})

@app.route('/get_otlieu_data')
def get_otlieu_data():
    global ot_lieu_data
    if ot_lieu_data is None or ot_lieu_data.empty:
        return jsonify({'columns': [], 'data': []})
    cols = list(ot_lieu_data.columns)
    rows = ot_lieu_data.fillna('').astype(str).values.tolist()
    return jsonify({'columns': cols, 'data': rows})

@app.route('/update_apply_row', methods=['POST'])
def update_apply_row():
    global apply_data
    try:
        data = request.json
        idx = int(data.get('index'))
        col = data.get('column')
        value = data.get('value')
        if apply_data is not None and 0 <= idx < len(apply_data) and col in apply_data.columns:
            apply_data.at[idx, col] = value
            return jsonify({'success': True})
        else:
            return jsonify({'error': 'Invalid index or column'}), 400
    except Exception as e:
        return jsonify({'error': str(e)}), 400

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000) 