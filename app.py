from flask import Flask, render_template, request, jsonify, send_file
from markupsafe import Markup
import pandas as pd
import openpyxl
import calendar
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
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

# Files path
TEMP_SIGNINOUT_PATH = os.path.join(app.config['UPLOAD_FOLDER'], 'temp_signinout.xlsx')
TEMP_APPLY_PATH = os.path.join(app.config['UPLOAD_FOLDER'], 'temp_apply.xlsx')
TEMP_OTLIEU_PATH = os.path.join(app.config['UPLOAD_FOLDER'], 'temp_otlieu.xlsx')
EMPLOYEE_LIST_PATH = os.path.join(app.config['UPLOAD_FOLDER'], 'employee_list.csv')
RULES_XLSX_PATH = os.path.join(app.config['UPLOAD_FOLDER'], 'rules.xlsx')

# Ensure upload folder exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Global variables to store data
sign_in_out_data = None
apply_data = None
ot_lieu_data = None
attendance_report = None
abnormal_data = None
emp_list = None
rules = None
global lieu_followup_df

# ==========================
# EMPLOYEE LIST
# ==========================
if os.path.exists(EMPLOYEE_LIST_PATH):
    try:
        employee_list_df = pd.read_csv(EMPLOYEE_LIST_PATH, dtype={'ID Number': str})
        if 'Dept' not in employee_list_df.columns:
            employee_list_df['Dept'] = ''
        if 'Internship' not in employee_list_df.columns:
            employee_list_df['Internship'] = ''
    except Exception:
        employee_list_df = pd.DataFrame(columns=["STT", "Name", "ID Number", "Dept", "Internship"])
else:
    employee_list_df = pd.DataFrame(columns=["STT", "Name", "ID Number", "Dept", "Internship"])

# ==========================
# RULES 
# ==========================
if os.path.exists(RULES_XLSX_PATH):
    try:
        rules = pd.read_excel(RULES_XLSX_PATH)
    except Exception as e:
        print(f"Error loading rules: {e}")
        rules = pd.DataFrame()
else:
    rules = pd.DataFrame(columns=['Holiday Date in This Year', 'Special Work Day'])
    try:
        rules.to_excel(RULES_XLSX_PATH, index=False)
        print(f"Created default rules file: {RULES_XLSX_PATH}")
    except Exception as e:
        print(f"Error creating rules file: {e}")
        rules = pd.DataFrame()

# Application Data - Translate Headers
APPLY_HEADER_MAP = {
    '申请人工号': 'Emp ID',
    '申请人': 'Emp Name',
    '提交人工号': 'Submit ID',
    '提交人': 'Submit Name',
    '申请时间': 'Apply Date',
    '起始时间': 'Start Date',
    '申请类型': 'Application Type',
    '终止时间': 'End Date',
    '申请说明': 'Note',
    '审批结果': 'Approve Result',
}

# =======================
# APPLY DATA
# =======================
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
    def normalize(col):
        return re.sub(r"[ '\"]+", '', str(col)).strip()
    norm_map = {normalize(k): v for k, v in APPLY_HEADER_MAP.items()}
    new_cols = []
    for col in df.columns:
        ncol = normalize(col)
        new_cols.append(norm_map.get(ncol, col))
    df.columns = new_cols
    return df

def filter_apply_employees(df, emp_list):
    if 'Emp Name' not in df.columns:
        return df
    emp_list = emp_list.copy()
    valid_names = set(emp_list['Name'].astype(str))
    return df[df['Emp Name'].astype(str).isin(valid_names)]

def add_apply_columns(df):
    # Type
    if 'Application Type' in df.columns:
        def extract_type(val):
            if not isinstance(val, str):
                return ''
            val_lower = val.lower()
            # English/Chinese mapping
            if 'trip' in val_lower or 'trips' in val_lower or '出差' in val_lower:
                return 'Trip'
            if 'leave' in val_lower or '事假' in val_lower or '年休假' in val_lower:
                return 'Leave'
            if 'supp' in val_lower or 'forgot' in val_lower or 'forget' in val_lower or '补单' in val_lower:
                return 'Supp'
            if 'replenishment' in val_lower or '个人补单' in val_lower:
                return 'Replenishment'
            if '病假' in val_lower:
                return 'Sick leave'
            return ''
        df['Type'] = df['Application Type'].apply(extract_type)
    else:
        df['Type'] = ''
    # Bổ sung: Nếu Type vẫn rỗng, kiểm tra Note
    if 'Note' in df.columns:
        def fill_type_from_note(row):
            if row['Type']:
                return row['Type']
            note = str(row['Note']).lower()
            if any(x in note for x in ['supp', 'forgot', 'forget', '补单']):
                return 'Supp'
            if any(x in note for x in ['trip', 'trips', '出差']):
                return 'Trip'
            if any(x in note for x in ['leave', '事假', '年休假']):
                return 'Leave'
            if '病假' in note:
                return 'Sick leave'
            return row['Type']
        df['Type'] = df.apply(fill_type_from_note, axis=1)
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
    def map_leave_type_applytype(val):
        if pd.isna(val):
            return ''
        val_str = str(val)
        val_lower = val_str.lower()
        # English/Chinese mapping
        if 'sick' in val_lower or '病假' in val_lower:
            return 'Sick leave'
        if 'welfare' in val_lower:
            return 'Welfare'
        if 'annual' in val_lower or '年休假' in val_lower:
            return 'Annual leave'
        if 'leave' in val_lower or 'unpaid' in val_lower or '事假' in val_lower:
            return 'Leave'
        if any(x in val_str for x in ['市内出差', '国内出差', '国际/中国港澳台出差']):
            return 'Trip'
        if '补单' in val_lower or 'replenishment' in val_lower:
            return 'Replenishment'
        return val_str
    if 'Application Type' in df.columns:
        df['Leave Type'] = df['Application Type'].apply(map_leave_type_applytype)
    elif 'Note' in df.columns:
        df['Leave Type'] = df['Note'].apply(map_leave_type)
    else:
        df['Leave Type'] = ''
    return df

# =======================
# OT & LIEU
# =======================
def process_ot_lieu_df(df, employee_list_df):

    if 'Title' in df.columns:
        def extract_name_id(title):
            match = re.search(r'_([a-zA-Z\s]+?)[_\s](\d{7,8})', str(title))
            if match:
                emp_id = match.group(2)
                match_row = employee_list_df[employee_list_df['Name'].astype(str).str.endswith(emp_id)]
                if not match_row.empty:
                    return match_row.iloc[0]['Name']
            return None
        df['Name'] = df['Title'].apply(extract_name_id)
        cols = list(df.columns)
        if 'Name' in cols:
            cols.insert(0, cols.pop(cols.index('Name')))
            df = df[cols]

    ot_from_cols = [c for c in df.columns if re.search(r'ot.*from', c, re.I)]
    ot_to_cols = [c for c in df.columns if re.search(r'ot.*to', c, re.I)]
    sum_ot_col = next((c for c in df.columns if re.search(r'ot.*sum', c, re.I)), None)
    lieu_from_cols = [c for c in df.columns if re.search(r'lieu.*from', c, re.I)]
    lieu_to_cols = [c for c in df.columns if re.search(r'lieu.*to', c, re.I)]
    sum_lieu_col = next((c for c in df.columns if re.search(r'lieu.*sum', c, re.I)), None)
    
    def clean_hours(val):
        if pd.isna(val): return ''
        s = str(val).strip().lower()
        s = s.replace('hours', '').replace('hour', '').replace(',', '.').replace(';', '.').replace('；', '.').replace('：', ':')
        s = s.replace('h', ':')
        s = re.sub(r'[^0-9.:]', '', s)

        if ':' in s:
            parts = s.split(':')
            try:
                hour = int(parts[0])
                minute = int(parts[1]) if len(parts) > 1 else 0
                return round(hour + minute/60, 2)
            except:
                return s
        try:
            return float(s)
        except:
            return s
    if sum_ot_col:
        df[sum_ot_col] = df[sum_ot_col].apply(clean_hours)
    if sum_lieu_col:
        df[sum_lieu_col] = df[sum_lieu_col].apply(clean_hours)

    # Mark warnings for invalid time format
    def mark_cell(val, error=False, suggest=None, warning=False):
        if error:
            return {'value': val, 'error': True, 'suggest': suggest}
        if warning:
            return {'value': val, 'warning': True}
        return val

    def parse_time_to_24h(val):
        if pd.isna(val) or not str(val).strip():
            return ''
        s = str(val).strip()
        m = re.match(r'^(0?[1-9]|1[0-2])\s*[:. ]\s*([0-5][0-9])\s*(AM|PM|am|pm)$', s)
        if m:
            hour = int(m.group(1))
            minute = int(m.group(2))
            ampm = m.group(3).upper()
            if ampm == 'PM' and hour != 12:
                hour += 12
            if ampm == 'AM' and hour == 12:
                hour = 0
            return f'{hour:02d}:{minute:02d}'
        m = re.match(r'^([01]?[0-9]|2[0-3])\s*[:. h]\s*([0-5][0-9])$', s)
        if m:
            hour = int(m.group(1))
            minute = int(m.group(2))
            return f'{hour:02d}:{minute:02d}'
        return None

# ---- OT From/To & Lieu From/To change format HH:MM ----
    def norm_time_to_24h(val):
        if str(val).strip() in ['Hour : Minutes AM', 'Hour : Minutes PM', 'Hour h Min']:
            return val
        parsed = parse_time_to_24h(val)
        if parsed is not None and parsed != '':
            return parsed
        elif val and str(val).strip():
            return mark_cell(val, warning=True)
        else:
            return val
    for col in ot_from_cols:
        df[col] = df[col].apply(norm_time_to_24h)
    for col in ot_to_cols:
        df[col] = df[col].apply(norm_time_to_24h)
    for col in lieu_from_cols:
        df[col] = df[col].apply(norm_time_to_24h)
    for col in lieu_to_cols:
        df[col] = df[col].apply(norm_time_to_24h)

    def is_time_ampm(val):
        if pd.isna(val) or not str(val).strip():
            return False
        s = str(val).strip()
        return bool(re.match(r'^(0?[1-9]|1[0-2])\s*[:. ]\s*([0-5][0-9])\s*(AM|PM|am|pm)$', s))

    def ampm_to_24h(s):
        try:
            t = datetime.strptime(s.strip().upper(), '%I:%M %p')
            return t
        except:
            return None
        
    # Convert AM/PM time to decimal hours
    def calc_hours_ampm(from_str, to_str):
        t1 = ampm_to_24h(from_str) if from_str and is_time_ampm(from_str) else None
        t2 = ampm_to_24h(to_str) if to_str and is_time_ampm(to_str) else None
        if t1 and t2:
            diff = (t2 - t1).total_seconds() / 3600
            if diff < 0:
                diff += 24
            # Trừ 1.5h nếu xuyên trưa
            if t1.hour <= 12 < t2.hour or (t1.hour == 12 and t2.hour > 13):
                if t2.hour > 13 or (t2.hour == 13 and t2.minute >= 30):
                    diff -= 1.5
            return round(diff, 2)
        return None


    # OT
    if ot_from_cols and ot_to_cols and sum_ot_col:
        for idx, row in df.iterrows():
            ot_from = row[ot_from_cols[0]] if not isinstance(row[ot_from_cols[0]], dict) else row[ot_from_cols[0]].get('value')
            ot_to = row[ot_to_cols[0]] if not isinstance(row[ot_to_cols[0]], dict) else row[ot_to_cols[0]].get('value')
            user_val = row[sum_ot_col] if not isinstance(row[sum_ot_col], dict) else row[sum_ot_col].get('value')
            # Nếu cả 2 thời gian hợp lệ HH:MM
            if re.match(r'^([01]?\d|2[0-3]):[0-5]\d$', str(ot_from)) and re.match(r'^([01]?\d|2[0-3]):[0-5]\d$', str(ot_to)):
                t1 = datetime.strptime(str(ot_from), '%H:%M')
                t2 = datetime.strptime(str(ot_to), '%H:%M')
                diff = (t2 - t1).total_seconds() / 3600
                if diff < 0:
                    diff += 24
                # Trừ 1.5h nếu xuyên trưa
                if t1.hour <= 12 < t2.hour or (t1.hour == 12 and t2.hour > 13):
                    if t2.hour > 13 or (t2.hour == 13 and t2.minute >= 30):
                        diff -= 1.5
                real = round(diff, 2)
                try:
                    user_val_f = float(user_val)
                except:
                    user_val_f = None
                if user_val_f is None or abs(real - user_val_f) > 0.01:
                    df.at[idx, sum_ot_col] = mark_cell(user_val, error=True, suggest=real)
                else:
                    df.at[idx, sum_ot_col] = real
    # Lieu
    if lieu_from_cols and lieu_to_cols and sum_lieu_col:
        for idx, row in df.iterrows():
            lieu_from = row[lieu_from_cols[0]] if not isinstance(row[lieu_from_cols[0]], dict) else row[lieu_from_cols[0]].get('value')
            lieu_to = row[lieu_to_cols[0]] if not isinstance(row[lieu_to_cols[0]], dict) else row[lieu_to_cols[0]].get('value')
            user_val = row[sum_lieu_col] if not isinstance(row[sum_lieu_col], dict) else row[sum_lieu_col].get('value')
            if re.match(r'^([01]?\d|2[0-3]):[0-5]\d$', str(lieu_from)) and re.match(r'^([01]?\d|2[0-3]):[0-5]\d$', str(lieu_to)):
                t1 = datetime.strptime(str(lieu_from), '%H:%M')
                t2 = datetime.strptime(str(lieu_to), '%H:%M')
                diff = (t2 - t1).total_seconds() / 3600
                if diff < 0:
                    diff += 24
                # Trừ 1.5h nếu xuyên trưa
                if t1.hour <= 12 < t2.hour or (t1.hour == 12 and t2.hour > 13):
                    if t2.hour > 13 or (t2.hour == 13 and t2.minute >= 30):
                        diff -= 1.5
                real = round(diff, 2)
                try:
                    user_val_f = float(user_val)
                except:
                    user_val_f = None
                if user_val_f is None or abs(real - user_val_f) > 0.01:
                    df.at[idx, sum_lieu_col] = mark_cell(user_val, error=True, suggest=real)
                else:
                    df.at[idx, sum_lieu_col] = real

    special_vals = ['Hour : Minutes AM', 'Hour : Minutes PM', 'Hour h Min']
    def is_empty_or_special(val):
        return pd.isna(val) or str(val).strip() == '' or str(val).strip() in special_vals
    if ot_from_cols and ot_to_cols and lieu_from_cols and lieu_to_cols:
        df = df[~(
            df[ot_from_cols[0]].apply(is_empty_or_special) &
            df[ot_to_cols[0]].apply(is_empty_or_special) &
            df[lieu_from_cols[0]].apply(is_empty_or_special) &
            df[lieu_to_cols[0]].apply(is_empty_or_special)
        )].reset_index(drop=True)

    def mark_gray(val):
        return {'value': val, 'gray': True}
    
    ot_day_col = next((c for c in df.columns if re.search(r'ot.*day', c, re.I)), None)
    lieu_date_col = next((c for c in df.columns if re.search(r'lieu.*date', c, re.I)), None)

    # OT From/To
    if ot_from_cols and ot_to_cols:
        for idx, row in df.iterrows():
            ot_from = row[ot_from_cols[0]]
            ot_to = row[ot_to_cols[0]]
            if is_empty_or_special(ot_from) and is_empty_or_special(ot_to):
                df.at[idx, ot_from_cols[0]] = mark_gray(ot_from)
                df.at[idx, ot_to_cols[0]] = mark_gray(ot_to)
                if sum_ot_col:
                    df.at[idx, sum_ot_col] = mark_gray(row[sum_ot_col])
                if ot_day_col:
                    df.at[idx, ot_day_col] = mark_gray(row[ot_day_col])
    # Lieu From/To
    if lieu_from_cols and lieu_to_cols:
        for idx, row in df.iterrows():
            lieu_from = row[lieu_from_cols[0]]
            lieu_to = row[lieu_to_cols[0]]
            if is_empty_or_special(lieu_from) and is_empty_or_special(lieu_to):
                df.at[idx, lieu_from_cols[0]] = mark_gray(lieu_from)
                df.at[idx, lieu_to_cols[0]] = mark_gray(lieu_to)
                if sum_lieu_col:
                    df.at[idx, sum_lieu_col] = mark_gray(row[sum_lieu_col])
                if lieu_date_col:
                    df.at[idx, lieu_date_col] = mark_gray(row[lieu_date_col])
    return df

# ========================
# UPLOAD & SAVE EXCEL
# ========================
def load_excel_data(file_path, sheet_name=None):
    """Load data from Excel file with support for .xls and .xlsx"""
    try:
        file_ext = os.path.splitext(file_path)[1].lower()
        if file_ext == '.xls':
            try:
                if sheet_name:
                    df = pd.read_excel(file_path, sheet_name=sheet_name, engine='xlrd')
                else:
                    df = pd.read_excel(file_path, engine='xlrd')
            except Exception as e:
                print(f"xlrd engine failed for .xls file: {e}")
                try:
                    if sheet_name:
                        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
                    else:
                        df = pd.read_excel(file_path, engine='openpyxl')
                except Exception as e2:
                    print(f"openpyxl engine also failed for .xls file: {e2}")
                    raise Exception(f"Cannot read .xls file. Please convert to .xlsx format. Error: {str(e)}")
        else:
            workbook = openpyxl.load_workbook(file_path, data_only=True)
            
            if sheet_name:
                if sheet_name in workbook.sheetnames:
                    df = pd.read_excel(file_path, sheet_name=sheet_name)
                else:
                    return None
            else:
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

# ========================
# APP ROUTE
# ========================
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
        if file.filename.lower().endswith('.csv'):
            df = try_read_csv(open(file_path, 'rb').read())
        elif file.filename.lower().endswith('.xls'):
            try:
                df = pd.read_excel(file_path, engine='xlrd')
            except Exception as e:
                print(f"xlrd engine failed: {e}")
                try:
                    df = pd.read_excel(file_path, engine='openpyxl')
                except Exception as e2:
                    print(f"openpyxl engine also failed: {e2}")
                    return jsonify({'error': f'Cannot read .xls file. Please convert to .xlsx format. Error: {str(e)}'}), 400
        else:
            df = pd.read_excel(file_path)

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
        if file.filename.lower().endswith('.csv'):
            df = try_read_csv(open(file_path, 'rb').read())
        elif file.filename.lower().endswith('.xls'):
            try:
                df = pd.read_excel(file_path, engine='xlrd')
            except Exception as e:
                print(f"xlrd engine failed: {e}")
                try:
                    df = pd.read_excel(file_path, engine='openpyxl')
                except Exception as e2:
                    print(f"openpyxl engine also failed: {e2}")
                    return jsonify({'error': f'Cannot read .xls file. Please convert to .xlsx format. Error: {str(e)}'}), 400
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
    global ot_lieu_data, employee_list_df
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    if file and file.filename.lower().endswith(('.xlsx', '.xls', '.csv')):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        data = load_excel_data(file_path)
        if data is not None:
            if 'OT From Note: 12AM is midnight' in data.columns:
                data = data.rename(columns={'OT From Note: 12AM is midnight': 'OT From'})
            data = process_ot_lieu_df(data, employee_list_df)
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
    
@app.route('/calculate_abnormal', methods=['POST'])
def calculate_abnormal():
    """Calculate abnormal attendance"""
    global abnormal_data, sign_in_out_data, apply_data, ot_lieu_data
    start_date = request.form.get('start_date', '')
    end_date = request.form.get('end_date', '')
    try:
        if sign_in_out_data is not None and not sign_in_out_data.empty:
            sign_in_out_data.to_excel(os.path.join(app.config['UPLOAD_FOLDER'], 'temp_signinout.xlsx'), index=False)
        
        if apply_data is not None and not apply_data.empty:
            if 'Results' in apply_data.columns:
                approved_apply = apply_data[apply_data['Results'] == 'Approved']
            else:
                approved_apply = apply_data
            approved_apply.to_excel(os.path.join(app.config['UPLOAD_FOLDER'], 'temp_apply.xlsx'), index=False)
        
        if ot_lieu_data is not None and not ot_lieu_data.empty:
            ot_lieu_save = ot_lieu_data.applymap(flatten_cell)
            ot_lieu_save.to_excel(os.path.join(app.config['UPLOAD_FOLDER'], 'temp_otlieu.xlsx'), index=False)
    except Exception as e:
        print(f"Error saving temporary data: {e}")
    
    if sign_in_out_data is None or sign_in_out_data.empty:
        return jsonify({'error': 'No sign in/out data available'}), 400
    abnormal_data = pd.DataFrame(columns=['Employee', 'Date', 'SignIn', 'SignOut', 'Status', 'LateMinutes'])
    date_info = ""
    if start_date and end_date:
        date_info = f" for period {start_date} to {end_date}"
    
    return jsonify({'success': True, 'message': f'Abnormal calculation completed{date_info} (placeholder).'})

@app.route('/get_abnormal_data')
def get_abnormal_data():
    """Get abnormal data for display"""
    global abnormal_data
    if abnormal_data is None or abnormal_data.empty:
        return jsonify({'data': [], 'message': 'No abnormal data available'})
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
    """
    Export processed data as Excel file using existing template.
    Write data to existing sheets, keep all formatting/colors.
    """
    global sign_in_out_data, apply_data, ot_lieu_data, employee_list_df, rules

    try:
        # Đường dẫn template gốc
        template_path = os.path.join(app.config['UPLOAD_FOLDER'], 'AttendanceReport.xlsx')
        if not os.path.exists(template_path):
            return jsonify({'error': 'Template file AttendanceReport.xlsx not found'}), 400

        # Tạo file mới từ template
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"AttendanceReport_{timestamp}.xlsx"
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        import shutil
        shutil.copy2(template_path, file_path)

        # Load workbook với openpyxl (giữ định dạng)
        from openpyxl import load_workbook
        wb = load_workbook(file_path)

        # Helper: Ghi DataFrame vào worksheet, giữ header, chỉ ghi dữ liệu từ dòng 2
        def write_df_to_sheet(ws, df, start_row=2):
            # Xóa dữ liệu cũ (giữ header)
            for row in ws.iter_rows(min_row=start_row, max_row=ws.max_row, max_col=ws.max_column):
                for cell in row:
                    cell.value = None
            # Ghi dữ liệu mới
            for idx, row in df.iterrows():
                for col_idx, value in enumerate(row, 1):
                    ws.cell(row=idx + start_row, column=col_idx, value=value)

        # 1. Employee List
        if employee_list_df is not None and not employee_list_df.empty and 'Emp List' in wb.sheetnames:
            write_df_to_sheet(wb['Emp List'], employee_list_df, start_row=2)

        # 2. Rules
        if rules is not None and not rules.empty and 'Rules' in wb.sheetnames:
            write_df_to_sheet(wb['Rules'], rules, start_row=2)

        # 3. Sign In-Out Data
        if sign_in_out_data is not None and not sign_in_out_data.empty and 'Sign in-out data' in wb.sheetnames:
            write_df_to_sheet(wb['Sign in-out data'], sign_in_out_data, start_row=2)

        # 4. Apply Data
        if apply_data is not None and not apply_data.empty and 'Apply data' in wb.sheetnames:
            write_df_to_sheet(wb['Apply data'], apply_data, start_row=2)

        # 6. OT Lieu Before (calculated)
        if 'OT Lieu Before' in wb.sheetnames:
            try:
                otlieu_before_df = calculate_otlieu_before()
                if otlieu_before_df is not None and not otlieu_before_df.empty:
                    write_df_to_sheet(wb['OT Lieu data'], otlieu_before_df, start_row=3)
            except Exception as e:
                print(f"Error calculating OT Lieu Before: {e}")

        # 7. OT Lieu Report (calculated)
        if 'OT & Lieu Report' in wb.sheetnames:
            try:
                otlieu_report_result = calculate_otlieu_report_for_export()
                if isinstance(otlieu_report_result, dict) and 'columns' in otlieu_report_result and 'rows' in otlieu_report_result:
                    otlieu_report_df = pd.DataFrame(otlieu_report_result['rows'], columns=otlieu_report_result['columns'])
                    if not otlieu_report_df.empty:
                        write_df_to_sheet(wb['OT & Lieu Report'], otlieu_report_df, start_row=9)
            except Exception as e:
                print(f"Error calculating OT Lieu Report: {e}")

        # 8. Total Attendance Detail (calculated)
        if 'Total Attendance detail' in wb.sheetnames:
            try:
                total_attendance_result = calculate_total_attendance_detail_for_export()
                if isinstance(total_attendance_result, dict) and 'columns' in total_attendance_result and 'rows' in total_attendance_result:
                    total_attendance_df = pd.DataFrame(total_attendance_result['rows'], columns=total_attendance_result['columns'])
                    if not total_attendance_df.empty:
                        write_df_to_sheet(wb['Total Attendance detail'], total_attendance_df, start_row=5)
            except Exception as e:
                print(f"Error calculating Total Attendance Detail: {e}")

        # 9. Attendance Report (calculated)
        if 'Attendance Report' in wb.sheetnames:
            try:
                attendance_report_result = calculate_attendance_report_for_export()
                if isinstance(attendance_report_result, dict) and 'columns' in attendance_report_result and 'rows' in attendance_report_result:
                    attendance_report_df = pd.DataFrame(attendance_report_result['rows'], columns=attendance_report_result['columns'])
                    if not attendance_report_df.empty:
                        write_df_to_sheet(wb['Attendance Report'], attendance_report_df, start_row=7)
            except Exception as e:
                print(f"Error calculating Attendance Report: {e}")

        # Lưu file và trả về
        wb.save(file_path)
        wb.close()
        return send_file(file_path, as_attachment=True, download_name=filename)

    except Exception as e:
        return jsonify({'error': f'Failed to export file: {str(e)}'}), 400

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
            if ext == '.xls':
                try:
                    excel = pd.ExcelFile(BytesIO(file_bytes), engine='xlrd')
                except Exception as e:
                    print(f"xlrd engine failed for .xls preview: {e}")
                    try:
                        excel = pd.ExcelFile(BytesIO(file_bytes), engine='openpyxl')
                    except Exception as e2:
                        print(f"openpyxl engine also failed for .xls preview: {e2}")
                        return jsonify({'error': f'Cannot read .xls file. Please convert to .xlsx format. Error: {str(e)}'}), 400
            else:
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
            if ext == '.xls':
                try:
                    df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name, engine='xlrd')
                except Exception as e:
                    print(f"xlrd engine failed for .xls import: {e}")
                    try:
                        df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name, engine='openpyxl')
                    except Exception as e2:
                        print(f"openpyxl engine also failed for .xls import: {e2}")
                        return jsonify({'error': f'Cannot read .xls file. Please convert to .xlsx format. Error: {str(e)}'}), 400
            else:
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
            if 'OT From Note: 12AM is midnight' in df.columns:
                df = df.rename(columns={'OT From Note: 12AM is midnight': 'OT From'})
            df = process_ot_lieu_df(df, employee_list_df)
            ot_lieu_data = df
        else:
            return jsonify({'error': 'Invalid data type'}), 400
        return jsonify({'success': True, 'message': f'{data_type.capitalize()} data imported successfully. Loaded {len(df)} rows.', 'rows': len(df)})
    except Exception as e:
        print(str(e))
        return jsonify({'error': f'Failed to import file: {str(e)}'}), 400

@app.route('/employee_list', methods=['GET'])
def employee_list():
    global employee_list_df
    df = employee_list_df.copy()
    if df.empty:
        return Markup('<div class="text-muted">No employees loaded.</div>')
    html = '<table class="table table-bordered table-sm employee-table"><thead><tr>'
    for col in df.columns:
        html += f'<th>{col}</th>'
    html += '</tr></thead><tbody>'
    for idx, row in df.iterrows():
        html += '<tr>'
        for col in df.columns:
            html += f'<td>{row[col]}</td>'
        html += f'<td><button class="hover-delete-btn" onclick="removeEmployee({idx})" title="Delete employee">×</button></td>'
        html += '</tr>'
    html += '</tbody></table>'
    return Markup(html)

@app.route('/employee_list_filtered', methods=['GET'])
def employee_list_filtered():
    global employee_list_df
    df = employee_list_df.copy()
    
    # Get filter parameters
    dept_filter = request.args.get('dept', '')
    intern_filter = request.args.get('intern', '')
    
    # Apply filters
    if dept_filter:
        df = df[df['Dept'].astype(str).str.contains(dept_filter, case=False, na=False)]
    
    if intern_filter:
        if intern_filter == 'Intern':
            df = df[df['Internship'].astype(str).str.contains('Intern', case=False, na=False)]
        elif intern_filter == 'Regular':
            df = df[~df['Internship'].astype(str).str.contains('Intern', case=False, na=False)]
    
    if df.empty:
        return Markup('<div class="text-muted">No employees match the filter criteria.</div>')
    
    html = '<table class="table table-bordered table-sm employee-table"><thead><tr>'
    for col in df.columns:
        html += f'<th>{col}</th>'
    html += '</tr></thead><tbody>'
    
    for idx, row in df.iterrows():
        html += '<tr>'
        for col in df.columns:
            html += f'<td>{row[col]}</td>'
        html += f'<td><button class="hover-delete-btn" onclick="removeEmployee({idx})" title="Delete employee">×</button></td>'
        html += '</tr>'
    html += '</tbody></table>'
    
    # Add filter summary
    filter_summary = f'<div class="text-muted small mb-2">Showing {len(df)} of {len(employee_list_df)} employees'
    if dept_filter or intern_filter:
        filter_summary += ' (filtered)'
    filter_summary += '</div>'
    
    return Markup(filter_summary + html)

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
            df = pd.read_excel(BytesIO(file_bytes), dtype=str)
        elif ext == '.csv':
            df = try_read_csv(file_bytes, dtype=str)
        else:
            return jsonify({'error': 'Unsupported file type'}), 400
        keep = ["STT", "Name", "ID Number", "Dept", "Internship"]
        for col in keep:
            if col not in df.columns:
                df[col] = ''
        df = df[keep]
        df = df.fillna('')
        if "ID Number" in df.columns:
            df["ID Number"] = df["ID Number"].astype(str)
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
                row[idx] = str(data.get('ID Number', ''))
            elif col == 'Dept':
                row[idx] = data.get('Dept', '')
            elif col == 'Internship':
                row[idx] = data.get('Internship', '')
        employee_list_df.loc[len(employee_list_df)] = row
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
    return jsonify({'success': True, 'message': 'Prepare data calculation completed.'})

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
    rows = ot_lieu_data.fillna('').values.tolist()
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

@app.route('/update_otlieu_row', methods=['POST'])
def update_otlieu_row():
    global ot_lieu_data
    try:
        data = request.json
        idx = int(data.get('index'))
        col = data.get('column')
        value = data.get('value')
        if ot_lieu_data is not None and 0 <= idx < len(ot_lieu_data) and col in ot_lieu_data.columns:
            ot_lieu_data.at[idx, col] = value
            updated = {}
            ot_from_cols = [c for c in ot_lieu_data.columns if re.search(r'ot.*from', c, re.I)]
            ot_to_cols = [c for c in ot_lieu_data.columns if re.search(r'ot.*to', c, re.I)]
            sum_ot_col = next((c for c in ot_lieu_data.columns if re.search(r'ot.*sum', c, re.I)), None)
            lieu_from_cols = [c for c in ot_lieu_data.columns if re.search(r'lieu.*from', c, re.I)]
            lieu_to_cols = [c for c in ot_lieu_data.columns if re.search(r'lieu.*to', c, re.I)]
            sum_lieu_col = next((c for c in ot_lieu_data.columns if re.search(r'lieu.*sum', c, re.I)), None)
            # OT
            if col in ot_from_cols + ot_to_cols and sum_ot_col:
                ot_from = ot_lieu_data.at[idx, ot_from_cols[0]]
                ot_to = ot_lieu_data.at[idx, ot_to_cols[0]]
                if re.match(r'^([01]?\d|2[0-3]):[0-5]\d$', str(ot_from)) and re.match(r'^([01]?\d|2[0-3]):[0-5]\d$', str(ot_to)):
                    t1 = datetime.strptime(str(ot_from), '%H:%M')
                    t2 = datetime.strptime(str(ot_to), '%H:%M')
                    diff = (t2 - t1).total_seconds() / 3600
                    if diff < 0:
                        diff += 24
                    if t1.hour <= 12 < t2.hour or (t1.hour == 12 and t2.hour > 13):
                        if t2.hour > 13 or (t2.hour == 13 and t2.minute >= 30):
                            diff -= 1.5
                    real = round(diff, 2)
                    ot_lieu_data.at[idx, sum_ot_col] = real
                    updated[sum_ot_col] = real

            # Lieu
            if col in lieu_from_cols + lieu_to_cols and sum_lieu_col:
                lieu_from = ot_lieu_data.at[idx, lieu_from_cols[0]]
                lieu_to = ot_lieu_data.at[idx, lieu_to_cols[0]]
                if re.match(r'^([01]?\d|2[0-3]):[0-5]\d$', str(lieu_from)) and re.match(r'^([01]?\d|2[0-3]):[0-5]\d$', str(lieu_to)):
                    t1 = datetime.strptime(str(lieu_from), '%H:%M')
                    t2 = datetime.strptime(str(lieu_to), '%H:%M')
                    diff = (t2 - t1).total_seconds() / 3600
                    if diff < 0:
                        diff += 24
                    if t1.hour <= 12 < t2.hour or (t1.hour == 12 and t2.hour > 13):
                        if t2.hour > 13 or (t2.hour == 13 and t2.minute >= 30):
                            diff -= 1.5
                    real = round(diff, 2)
                    ot_lieu_data.at[idx, sum_lieu_col] = real
                    updated[sum_lieu_col] = real

            # thêm sau khi chỉnh sửa
            temp_path = os.path.join(app.config['UPLOAD_FOLDER'], 'temp_otlieu.xlsx')
            flat_df = ot_lieu_data.applymap(flatten_cell)
            flat_df.to_excel(temp_path, index=False)

            return jsonify({'success': True, 'updated': updated})
        else:
            return jsonify({'error': 'Invalid index or column'}), 400
    except Exception as e:
        return jsonify({'error': str(e)}), 400

@app.route('/get_apply_column_options')
def get_apply_column_options():
    global apply_data
    default_type = ["Trip", "Leave", "Supp", "Replenishment", ""]
    default_leave_type = ["Unpaid", "Sick", "Welfare", "Annual", ""]
    leave_type_values = []
    if apply_data is not None and not apply_data.empty and 'Leave Type' in apply_data.columns:
        leave_type_values = [v for v in apply_data['Leave Type'].unique() if str(v).strip() != '']
    leave_type_all = default_leave_type + [v for v in leave_type_values if v not in default_leave_type]
    
    return jsonify({"Type": default_type, "Leave Type": leave_type_all})

def get_holidays_from_rules():
    global rules
    try:
        if rules is not None and 'Holiday Date in This Year' in rules.columns:
            holidays = []
            for val in rules['Holiday Date in This Year']:
                if pd.notna(val):
                    try:
                        d = pd.to_datetime(val)
                        holidays.append(d.strftime('%m-%d'))
                    except:
                        pass
            return set(holidays)
    except Exception as e:
        print("Error reading holidays from rules:", e)
    return set()

# ----------------------------
# ROUTE - RULES PAGE
# ----------------------------
@app.route('/get_rules_table')
def get_rules_table():
    global rules
    try:
        if rules is not None:
            df = rules.fillna('')
            columns = list(df.columns)
            rows = df.astype(str).values.tolist()
            return jsonify({'columns': columns, 'rows': rows})
        else:
            return jsonify({'columns': [], 'rows': [], 'error': 'No rules loaded'})
    except Exception as e:
        return jsonify({'columns': [], 'rows': [], 'error': str(e)})

@app.route('/update_rule_cell', methods=['POST'])
def update_rule_cell():
    global rules
    try:
        data = request.json
        row = int(data['row'])
        col = int(data['col'])
        value = data['value']
        if rules is not None:
            rules.iat[row, col] = value
            rules.to_excel(RULES_XLSX_PATH, index=False)
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/delete_rule_row', methods=['POST'])
def delete_rule_row():
    global rules
    try:
        data = request.json
        row = int(data['row'])
        if rules is not None:
            rules = rules.drop(rules.index[row]).reset_index(drop=True)
            rules.to_excel(RULES_XLSX_PATH, index=False)
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/add_rule_row', methods=['POST'])
def add_rule_row():
    global rules
    try:
        if rules is not None:
            new_row = ['' for _ in rules.columns]
            rules.loc[len(rules)] = new_row
            rules.to_excel(RULES_XLSX_PATH, index=False)
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/batch_update_rule_cells', methods=['POST'])
def batch_update_rule_cells():
    global rules
    try:
        data = request.json
        edits = data.get('edits', [])
        if rules is not None:
            for edit in edits:
                row = int(edit['row'])
                col = int(edit['col'])
                value = edit['value']
                rules.iat[row, col] = value
            rules.to_excel(RULES_XLSX_PATH, index=False)
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/reload_rules', methods=['POST'])
def reload_rules():
    """Reload rules from file"""
    global rules
    try:
        if os.path.exists(RULES_XLSX_PATH):
            rules = pd.read_excel(RULES_XLSX_PATH)
            return jsonify({'success': True, 'message': 'Rules reloaded successfully'})
        else:
            return jsonify({'success': False, 'error': 'Rules file not found'})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/get_temp_signinout_data')
def get_temp_signinout_data():
    path = os.path.join(app.config['UPLOAD_FOLDER'], 'temp_signinout.xlsx')
    if not os.path.exists(path):
        return jsonify({'columns': [], 'data': []})
    df = pd.read_excel(path)
    cols = list(df.columns)
    rows = df.fillna('').astype(str).values.tolist()
    return jsonify({'columns': cols, 'data': rows})

@app.route('/get_temp_apply_data')
def get_temp_apply_data():
    path = os.path.join(app.config['UPLOAD_FOLDER'], 'temp_apply.xlsx')
    if not os.path.exists(path):
        return jsonify({'columns': [], 'data': []})
    df = pd.read_excel(path)
    cols = list(df.columns)
    rows = df.fillna('').astype(str).values.tolist()
    return jsonify({'columns': cols, 'data': rows})

@app.route('/get_temp_otlieu_data')
def get_temp_otlieu_data():
    path = os.path.join(app.config['UPLOAD_FOLDER'], 'temp_otlieu.xlsx')
    if not os.path.exists(path):
        return jsonify({'columns': [], 'data': []})
    df = pd.read_excel(path)
    cols = list(df.columns)
    rows = df.fillna('').astype(str).values.tolist()
    return jsonify({'columns': cols, 'data': rows})

LIEU_FOLLOWUP_PATH = os.path.join(app.config['UPLOAD_FOLDER'], 'lieu_followup.xlsx')
if os.path.exists(LIEU_FOLLOWUP_PATH):
    lieu_followup_df = pd.read_excel(LIEU_FOLLOWUP_PATH)
else:
    lieu_followup_df = pd.DataFrame(columns=['Name', 'Lieu remain previous month'])

@app.route('/get_lieu_followup')
def get_lieu_followup():
    if not os.path.exists(LIEU_FOLLOWUP_PATH):
        if employee_list_df is not None and not employee_list_df.empty:
            df = employee_list_df[['Name']].copy()
            df['Lieu remain previous month'] = 0
            df.insert(0, 'STT', range(1, len(df) + 1))
            df.to_excel(LIEU_FOLLOWUP_PATH, index=False)
        else:
            df = pd.DataFrame(columns=['STT', 'Name', 'Lieu remain previous month'])
    else:
        df = pd.read_excel(LIEU_FOLLOWUP_PATH)
    cols = list(df.columns)
    rows = df.fillna('').astype(str).values.tolist()
    return jsonify({'columns': cols, 'data': rows})

@app.route('/import_lieu_followup', methods=['POST'])
def import_lieu_followup():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    try:
        df = pd.read_excel(file)
        # Chuẩn hóa cột
        if 'Name' not in df.columns:
            return jsonify({'error': 'Missing Name column'}), 400
        if 'Lieu remain previous month' not in df.columns:
            df['Lieu remain previous month'] = 0
        df = df[['Name', 'Lieu remain previous month']]
        df.insert(0, 'STT', range(1, len(df) + 1))
        df.to_excel(LIEU_FOLLOWUP_PATH, index=False)
        return jsonify({'success': True, 'rows': len(df)})
    except Exception as e:
        return jsonify({'error': str(e)}), 400

@app.route('/add_lieu_followup_row', methods=['POST'])
def add_lieu_followup_row():
    try:
        data = request.json
        name = data.get('Name', '').strip()
        remain = data.get('Lieu remain previous month', 0)
        if not name:
            return jsonify({'error': 'Name required'}), 400
        if os.path.exists(LIEU_FOLLOWUP_PATH):
            df = pd.read_excel(LIEU_FOLLOWUP_PATH)
        else:
            df = pd.DataFrame(columns=['STT', 'Name', 'Lieu remain previous month'])
        df = df.append({'Name': name, 'Lieu remain previous month': remain}, ignore_index=True)
        df['STT'] = range(1, len(df) + 1)
        df.to_excel(LIEU_FOLLOWUP_PATH, index=False)
        return jsonify({'success': True, 'rows': len(df)})
    except Exception as e:
        return jsonify({'error': str(e)}), 400

@app.route('/update_lieu_followup_row', methods=['POST'])
def update_lieu_followup_row():
    try:
        data = request.json
        idx = int(data.get('index'))
        col = data.get('column')
        value = data.get('value')
        if os.path.exists(LIEU_FOLLOWUP_PATH):
            df = pd.read_excel(LIEU_FOLLOWUP_PATH)
        else:
            return jsonify({'error': 'No data'}), 400
        if 0 <= idx < len(df) and col in df.columns:
            df.at[idx, col] = value
            df.to_excel(LIEU_FOLLOWUP_PATH, index=False)
            return jsonify({'success': True})
        else:
            return jsonify({'error': 'Invalid index or column'}), 400
    except Exception as e:
        return jsonify({'error': str(e)}), 400

@app.route('/delete_lieu_followup_row', methods=['POST'])
def delete_lieu_followup_row():
    try:
        data = request.json
        idx = int(data.get('index'))
        if os.path.exists(LIEU_FOLLOWUP_PATH):
            df = pd.read_excel(LIEU_FOLLOWUP_PATH)
        else:
            return jsonify({'error': 'No data'}), 400
        if 0 <= idx < len(df):
            df = df.drop(idx).reset_index(drop=True)
            df['STT'] = range(1, len(df) + 1)
            df.to_excel(LIEU_FOLLOWUP_PATH, index=False)
            return jsonify({'success': True})
        else:
            return jsonify({'error': 'Invalid index'}), 400
    except Exception as e:
        return jsonify({'error': str(e)}), 400

# ==== TÍNH TOÁN OT LIEU BEFORE (HÀM CHUNG) ====
def calculate_otlieu_before():
    global rules
    
    path = os.path.join(app.config['UPLOAD_FOLDER'], 'temp_otlieu.xlsx')
    if not os.path.exists(path):
        return pd.DataFrame()
    df = pd.read_excel(path)

    # Load Lieu followup
    lieu_followup_path = os.path.join(app.config['UPLOAD_FOLDER'], 'lieu_followup.xlsx')
    if os.path.exists(lieu_followup_path):
        lieu_followup_df = pd.read_excel(lieu_followup_path)
    else:
        lieu_followup_df = pd.DataFrame(columns=['Name', 'Lieu remain previous month'])
    
    # Load rules if not loaded
    if rules is None or rules.empty:
        rules_path = os.path.join(app.config['UPLOAD_FOLDER'], 'rules.xlsx')
        if os.path.exists(rules_path):
            rules = pd.read_excel(rules_path)
            print(f"Loaded rules from file: {len(rules)} rows")
        else:
            rules = pd.DataFrame()
            print("Rules file not found, using empty DataFrame")
    else:
        print(f"Using existing rules: {len(rules)} rows")

    # Build remain map: {name: remain}
    lieu_remain_map = {}
    if 'Name' in lieu_followup_df.columns and 'Lieu remain previous month' in lieu_followup_df.columns:
        for _, r in lieu_followup_df.iterrows():
            name = str(r['Name']).strip()
            try:
                remain = float(r['Lieu remain previous month'])
            except:
                remain = 0.0
            lieu_remain_map[name] = remain

    ot_payment_types = [
        'Weekday Rate 150%',
        'Weekday-night Rate 200%',
        'Weekend Rate 200%',
        'Weekend-night Rate 270%',
        'Holiday Rate 300%',
        'Holiday-night Rate 390%',
    ]
    change_in_lieu_types = ot_payment_types.copy()
    for col in ot_payment_types:
        df['OT Payment: ' + col] = 0.0
    for col in change_in_lieu_types:
        df['Change in lieu: ' + col] = 0.0

    # Hàm xác định loại ngày
    def get_day_type(dt, holidays, special_weekends, special_workdays):
        if dt in holidays:
            return 'Holiday'
        if dt in special_workdays:
            return 'Weekday'
        if dt in special_weekends:
            return 'Weekend'
        if dt.weekday() >= 5:
            return 'Weekend'
        return 'Weekday'

    holidays = set()
    special_workdays = set()
    special_weekends = set()
    try:
        if rules is not None:
            if 'Holiday Date in This Year' in rules.columns:
                holidays = set(pd.to_datetime(rules['Holiday Date in This Year'], errors='coerce').dt.date.dropna())
            if 'Special Work Day' in rules.columns:
                special_workdays = set(pd.to_datetime(rules['Special Work Day'], errors='coerce').dt.date.dropna())
            if 'Special Weekend' in rules.columns:
                special_weekends = set(pd.to_datetime(rules['Special Weekend'], errors='coerce').dt.date.dropna())
    except: pass

    # --- BẮT ĐẦU LOGIC TÍNH OT RATES ---
    for idx, row in df.iterrows():
        emp_id = row['Emp ID'] if 'Emp ID' in row else None
        intern = is_intern(emp_id, employee_list_df)
        
        # Ưu tiên lấy từ cột 'Date', sau đó 'OT date', sau đó 'Lieu Date'
        ot_date = None
        if 'Date' in df.columns and pd.notna(row.get('Date', None)):
            try:
                ot_date = pd.to_datetime(row['Date']).date()
            except:
                pass
        elif 'OT date' in df.columns and pd.notna(row.get('OT date', None)):
            try:
                ot_date = pd.to_datetime(row['OT date']).date()
            except:
                pass
        elif 'Lieu Date' in df.columns and pd.notna(row.get('Lieu Date', None)):
            try:
                ot_date = pd.to_datetime(row['Lieu Date']).date()
            except:
                pass
        if ot_date is None:
            continue

        # Lấy OT From và OT To
        ot_from, ot_to = None, None
        for col in df.columns:
            if 'ot' in col.lower() and 'from' in col.lower():
                ot_from = row[col]
            if 'ot' in col.lower() and 'to' in col.lower():
                ot_to = row[col]
        if not ot_from or not ot_to or pd.isna(ot_from) or pd.isna(ot_to):
            continue
        try:
            t1 = pd.to_datetime(str(ot_from), format='%H:%M')
            t2 = pd.to_datetime(str(ot_to), format='%H:%M')
        except:
            continue
        if t2 < t1:
            t2 += pd.Timedelta(days=1)

        # Tính OT theo từng block thời gian
        cur = t1
        while cur < t2:
            hour = cur.hour + cur.minute / 60
            if 6 <= hour < 22:
                block_end = min(cur.replace(hour=22, minute=0), t2)
                block_type = 'day'
            else:
                if hour < 6:
                    next6 = cur.replace(hour=6, minute=0)
                    if next6 <= cur: next6 += pd.Timedelta(days=1)

# ==== TÍNH TOÁN OT LIEU BEFORE (HÀM CHUNG) ====
def calculate_otlieu_before():
    global rules
    
    path = os.path.join(app.config['UPLOAD_FOLDER'], 'temp_otlieu.xlsx')
    if not os.path.exists(path):
        return pd.DataFrame()
    df = pd.read_excel(path)

    # Load Lieu followup
    lieu_followup_path = os.path.join(app.config['UPLOAD_FOLDER'], 'lieu_followup.xlsx')
    if os.path.exists(lieu_followup_path):
        lieu_followup_df = pd.read_excel(lieu_followup_path)
    else:
        lieu_followup_df = pd.DataFrame(columns=['Name', 'Lieu remain previous month'])
    
    # Load rules if not loaded
    if rules is None or rules.empty:
        rules_path = os.path.join(app.config['UPLOAD_FOLDER'], 'rules.xlsx')
        if os.path.exists(rules_path):
            rules = pd.read_excel(rules_path)
            print(f"Loaded rules from file: {len(rules)} rows")
        else:
            rules = pd.DataFrame()
            print("Rules file not found, using empty DataFrame")
    else:
        print(f"Using existing rules: {len(rules)} rows")

    # Build remain map: {name: remain}
    lieu_remain_map = {}
    if 'Name' in lieu_followup_df.columns and 'Lieu remain previous month' in lieu_followup_df.columns:
        for _, r in lieu_followup_df.iterrows():
            name = str(r['Name']).strip()
            try:
                remain = float(r['Lieu remain previous month'])
            except:
                remain = 0.0
            lieu_remain_map[name] = remain

    ot_payment_types = [
        'Weekday Rate 150%',
        'Weekday-night Rate 200%',
        'Weekend Rate 200%',
        'Weekend-night Rate 270%',
        'Holiday Rate 300%',
        'Holiday-night Rate 390%',
    ]
    change_in_lieu_types = ot_payment_types.copy()
    for col in ot_payment_types:
        df['OT Payment: ' + col] = 0.0
    for col in change_in_lieu_types:
        df['Change in lieu: ' + col] = 0.0

    # Hàm xác định loại ngày
    def get_day_type(dt, holidays, special_weekends, special_workdays):
        if dt in holidays:
            return 'Holiday'
        if dt in special_workdays:
            return 'Weekday'
        if dt in special_weekends:
            return 'Weekend'
        if dt.weekday() >= 5:
            return 'Weekend'
        return 'Weekday'

    holidays = set()
    special_workdays = set()
    special_weekends = set()
    try:
        if rules is not None:
            if 'Holiday Date in This Year' in rules.columns:
                holidays = set(pd.to_datetime(rules['Holiday Date in This Year'], errors='coerce').dt.date.dropna())
            if 'Special Work Day' in rules.columns:
                special_workdays = set(pd.to_datetime(rules['Special Work Day'], errors='coerce').dt.date.dropna())
            if 'Special Weekend' in rules.columns:
                special_weekends = set(pd.to_datetime(rules['Special Weekend'], errors='coerce').dt.date.dropna())
    except: pass

    # --- BẮT ĐẦU LOGIC TÍNH OT RATES ---
    for idx, row in df.iterrows():
        emp_id = row['Emp ID'] if 'Emp ID' in row else None
        intern = is_intern(emp_id, employee_list_df)
        
        # Ưu tiên lấy từ cột 'Date', sau đó 'OT date', sau đó 'Lieu Date'
        ot_date = None
        if 'Date' in df.columns and pd.notna(row.get('Date', None)):
            try:
                ot_date = pd.to_datetime(row['Date']).date()
            except:
                pass
        elif 'OT date' in df.columns and pd.notna(row.get('OT date', None)):
            try:
                ot_date = pd.to_datetime(row['OT date']).date()
            except:
                pass
        elif 'Lieu Date' in df.columns and pd.notna(row.get('Lieu Date', None)):
            try:
                ot_date = pd.to_datetime(row['Lieu Date']).date()
            except:
                pass
        if ot_date is None:
            continue

        # Lấy OT From và OT To
        ot_from, ot_to = None, None
        for col in df.columns:
            if 'ot' in col.lower() and 'from' in col.lower():
                ot_from = row[col]
            if 'ot' in col.lower() and 'to' in col.lower():
                ot_to = row[col]
        if not ot_from or not ot_to or pd.isna(ot_from) or pd.isna(ot_to):
            continue
        try:
            t1 = pd.to_datetime(str(ot_from), format='%H:%M')
            t2 = pd.to_datetime(str(ot_to), format='%H:%M')
        except:
            continue
        if t2 < t1:
            t2 += pd.Timedelta(days=1)

        # Tính OT theo từng block thời gian
        cur = t1
        while cur < t2:
            hour = cur.hour + cur.minute / 60
            if 6 <= hour < 22:
                block_end = min(cur.replace(hour=22, minute=0), t2)
                block_type = 'day'
            else:
                if hour < 6:
                    next6 = cur.replace(hour=6, minute=0)
                    if next6 <= cur: next6 += pd.Timedelta(days=1)
                    block_end = min(next6, t2)
                else:
                    next22 = cur.replace(hour=22, minute=0)
                    if next22 <= cur: next22 += pd.Timedelta(days=1)
                    block_end = min(next22, t2)
                block_type = 'night'
            block_date = (ot_date if cur.day == t1.day else ot_date + pd.Timedelta(days=1))
            day_type = get_day_type(block_date, holidays, special_weekends, special_workdays)
            hours = (block_end - cur).total_seconds() / 3600

            if block_type == 'day':
                lunch_start = cur.replace(hour=12, minute=0)
                lunch_end = cur.replace(hour=13, minute=30)
                overlap = max(timedelta(0), min(block_end, lunch_end) - max(cur, lunch_start)).total_seconds() / 3600
                if overlap > 0:
                    hours -= overlap

            target_prefix = 'Change in lieu: ' if intern else 'OT Payment: '
            rate_map = {
                ('Weekday', 'day'): 'Weekday Rate 150%',
                ('Weekday', 'night'): 'Weekday-night Rate 200%',
                ('Weekend', 'day'): 'Weekend Rate 200%',
                ('Weekend', 'night'): 'Weekend-night Rate 270%',
                ('Holiday', 'day'): 'Holiday Rate 300%',
                ('Holiday', 'night'): 'Holiday-night Rate 390%',
            }
            rate_label = rate_map.get((day_type, block_type))
            if rate_label:
                col = target_prefix + rate_label
                df.at[idx, col] += hours
            cur = block_end


    # --- LOGIC TRỪ LIEU ---
    for idx, row in df.iterrows():
        name = row['Name'] if 'Name' in row else None
        if not name:
            continue
            
        # Lấy thông tin Lieu Sum (giờ nghỉ Lieu)
        lieu_sum_col = next((c for c in df.columns if re.search(r'lieu.*sum', c, re.I)), None)
        lieu_sum = 0.0
        if lieu_sum_col:
            try:
                lieu_sum = float(row[lieu_sum_col]) if pd.notna(row[lieu_sum_col]) and str(row[lieu_sum_col]).strip() != '' else 0.0
            except:
                lieu_sum = 0.0
                
        # Nếu có Lieu Sum > 0, trừ vào OT Payment theo thứ tự ưu tiên
        if lieu_sum > 0:
            lieu_remain_old = lieu_remain_map.get(name, 0.0)
            lieu_to_deduct = min(lieu_sum, lieu_remain_old)
            lieu_used = 0.0
            
            if lieu_to_deduct > 0:
                # Thứ tự ưu tiên: Weekday → Night → Weekend → Holiday
                ot_priority = [
                    ('OT Payment: Weekday Rate 150%', 1.5, 'Change in lieu: Weekday Rate 150%'),
                    ('OT Payment: Weekday-night Rate 200%', 2.0, 'Change in lieu: Weekday-night Rate 200%'),
                    ('OT Payment: Weekend Rate 200%', 2.0, 'Change in lieu: Weekend Rate 200%'),
                    ('OT Payment: Weekend-night Rate 270%', 2.7, 'Change in lieu: Weekend-night Rate 270%'),
                    ('OT Payment: Holiday Rate 300%', 3.0, 'Change in lieu: Holiday Rate 300%'),
                    ('OT Payment: Holiday-night Rate 390%', 3.9, 'Change in lieu: Holiday-night Rate 390%'),
                ]
                
                remain = lieu_to_deduct
                for ot_col, ratio, lieu_col in ot_priority:
                    if remain <= 0:
                        break
                        
                    ot_val = df.at[idx, ot_col] if ot_col in df.columns else 0.0
                    try:
                        ot_val = float(ot_val) if pd.notna(ot_val) and str(ot_val).strip() != '' else 0.0
                    except:
                        ot_val = 0.0
                        
                    if ot_val <= 0:
                        continue
                        
                    # Số giờ OT có thể dùng để đổi Lieu ở hệ số này
                    max_lieu_from_this = ot_val * ratio
                    lieu_from_this = min(remain, max_lieu_from_this)
                    
                    # Số giờ OT bị trừ = lieu_from_this / ratio
                    ot_deduct = lieu_from_this / ratio
                    
                    # Cập nhật vào DataFrame
                    df.at[idx, ot_col] = round(ot_val - ot_deduct, 3)
                    df.at[idx, lieu_col] = round(lieu_from_this, 3)
                    
                    remain -= lieu_from_this
                    lieu_used += lieu_from_this
                
                # Cập nhật Lieu remain mới
                lieu_remain_new = round(lieu_remain_old - lieu_used, 3)
                lieu_remain_map[name] = lieu_remain_new
                
                # Ghi nhận vào DataFrame
                df.at[idx, 'Lieu used'] = round(lieu_used, 3)
                df.at[idx, 'Lieu Remain'] = round(lieu_remain_new, 3)
            else:
                df.at[idx, 'Lieu used'] = 0.0
                df.at[idx, 'Lieu Remain'] = round(lieu_remain_old, 3)
        else:
            # Không có Lieu Sum, ghi nhận Lieu Remain cũ
            lieu_remain_old = lieu_remain_map.get(name, 0.0)
            df.at[idx, 'Lieu used'] = 0.0
            df.at[idx, 'Lieu Remain'] = round(lieu_remain_old, 3)

    # Làm tròn 3 số
    for col in ['OT Payment: ' + c for c in ot_payment_types] + ['Change in lieu: ' + c for c in change_in_lieu_types]:
        if col in df.columns:
            df[col] = df[col].apply(lambda x: round(float(x), 3) if str(x).replace('.', '', 1).replace('-', '', 1).isdigit() else x)

    # Ensure required columns exist and reorder
    import numpy as np
    if 'Lieu used' not in df.columns:
        df['Lieu used'] = np.nan
    if 'Lieu Remain' not in df.columns:
        df['Lieu Remain'] = np.nan
    if 'Remark' not in df.columns:
        df['Remark'] = np.nan
        
    # Reorder columns: put Lieu used, Lieu Remain, and Remark at the end
    cols = [c for c in df.columns if c not in ['Lieu used', 'Lieu Remain', 'Remark']] + ['Lieu used', 'Lieu Remain', 'Remark']
    df = df[cols]
    return df

@app.route('/get_otlieu_before')
def get_otlieu_before():
    df = calculate_otlieu_before()
    if df.empty:
        return jsonify({'columns': [], 'data': []})
    cols = list(df.columns)
    rows = df.fillna('').astype(str).values.tolist()
    return jsonify({'columns': cols, 'data': rows})

@app.route('/get_otlieu_report')
def get_otlieu_report():
    emp_path = EMPLOYEE_LIST_PATH
    if os.path.exists(emp_path):
        emp_df = pd.read_csv(emp_path, dtype=str)
    else:
        emp_df = pd.DataFrame(columns=["Name", "ID Number"])
    emp_df = emp_df.fillna('')
    emp_df['No'] = range(1, len(emp_df) + 1)

    ot_df = calculate_otlieu_before()
    ot_cols = [
        'OT Payment: Weekday Rate 150%',
        'OT Payment: Weekday-night Rate 200%',
        'OT Payment: Weekend Rate 200%',
        'OT Payment: Weekend-night Rate 270%',
        'OT Payment: Holiday Rate 300%',
        'OT Payment: Holiday-night Rate 390%',
        'Change in lieu: Weekday Rate 150%',
        'Change in lieu: Weekday-night Rate 200%',
        'Change in lieu: Weekend Rate 200%',
        'Change in lieu: Weekend-night Rate 270%',
        'Change in lieu: Holiday Rate 300%',
        'Change in lieu: Holiday-night Rate 390%',
    ]
    ot_sum = None
    if not ot_df.empty and 'Name' in ot_df.columns:
        ot_df[ot_cols] = ot_df[ot_cols].fillna(0)
        ot_sum = ot_df.groupby('Name')[ot_cols].sum().reset_index()
        
        # Tính các cột mới theo yêu cầu
        used_cols = [c for c in ot_cols if c.startswith('Change in lieu')]
        paid_cols = [c for c in ot_cols if c.startswith('OT Payment')]
        
        # 1. Total used hours in month: Cộng lại các giờ đã dùng
        ot_sum['Total used hours in month'] = ot_sum[used_cols].sum(axis=1)
        
        # 2. Total OT paid: Cộng lại các cột OT Payment
        ot_sum['Total OT paid'] = ot_sum[paid_cols].sum(axis=1)
        
        # 3. Transferred to normal working hours: Logic mới
        def calc_transfer(row):
            try:
                total_ot_paid = float(row['Total OT paid'])
                total_used = float(row['Total used hours in month'])
                
                # Nếu có Lieu Sum (total_used) mà không có OT Payment (total_ot_paid = 0)
                # thì chuyển Lieu Sum vào "Transferred to normal working hours"
                if total_used > 0 and total_ot_paid == 0:
                    return round(total_used, 2)
                # Nếu có OT Payment > 25, thì chuyển phần dư vào normal working hours
                elif total_ot_paid > 25:
                    return round(total_ot_paid - 25, 2)
                else:
                    return ''
            except:
                return ''
        
        ot_sum['Transferred to normal working hours'] = ot_sum.apply(calc_transfer, axis=1)
        
        # 4. Remain unused time off in lieu: Phần Lieu chưa dùng hết
        # Load Lieu followup để lấy Lieu remain
        lieu_followup_path = os.path.join(app.config['UPLOAD_FOLDER'], 'lieu_followup.xlsx')
        lieu_remain_map = {}
        if os.path.exists(lieu_followup_path):
            lieu_followup_df = pd.read_excel(lieu_followup_path)
            if 'Name' in lieu_followup_df.columns and 'Lieu remain previous month' in lieu_followup_df.columns:
                for _, r in lieu_followup_df.iterrows():
                    name = str(r['Name']).strip()
                    try:
                        remain = float(r['Lieu remain previous month'])
                    except:
                        remain = 0.0
                    lieu_remain_map[name] = remain
        
        def calc_remain_unused(row):
            try:
                name = str(row['Name']).strip()
                lieu_remain_old = lieu_remain_map.get(name, 0.0)
                lieu_used = 0.0
                
                # Tính Lieu used từ các cột Change in lieu
                for col in used_cols:
                    try:
                        lieu_used += float(row[col]) if pd.notna(row[col]) else 0.0
                    except:
                        pass
                
                # Remain unused = Lieu remain cũ - Lieu used
                remain_unused = lieu_remain_old - lieu_used
                return round(remain_unused, 2) if remain_unused > 0 else 0.0
            except:
                return 0.0
        
        ot_sum['Remain unused time off in lieu'] = ot_sum.apply(calc_remain_unused, axis=1)
        ot_sum['Date'] = ''

    # Đảm bảo cột Name tồn tại ở cả hai DataFrame trước khi merge
    result = emp_df[['No', 'ID Number', 'Name']].copy()
    result = result.rename(columns={'ID Number': 'Employee ID'})
    # Không đổi tên 'Name' thành 'Employee Name' trước khi merge
    if ot_sum is not None:
        # Ép kiểu và strip để tránh lỗi merge do khác kiểu
        result['Name'] = result['Name'].astype(str).str.strip()
        ot_sum['Name'] = ot_sum['Name'].astype(str).str.strip()
        result = result.merge(ot_sum, on='Name', how='left')
    # Sau khi merge xong mới đổi tên cột cho thân thiện
    result = result.rename(columns={'Name': 'Employee Name'})
    col_rename = {
        'OT Payment: Weekday Rate 150%': 'OT weekday 150%',
        'OT Payment: Weekday-night Rate 200%': 'OT weekday night 200%',
        'OT Payment: Weekend Rate 200%': 'OT weekly holiday 200%',
        'OT Payment: Weekend-night Rate 270%': 'OT weekly holiday night 270%',
        'OT Payment: Holiday Rate 300%': 'OT public holiday 300%',
        'OT Payment: Holiday-night Rate 390%': 'OT public holiday night 390%',
        'Change in lieu: Weekday Rate 150%': 'OT weekday 150% (lieu)',
        'Change in lieu: Weekday-night Rate 200%': 'OT weekday night 200% (lieu)',
        'Change in lieu: Weekend Rate 200%': 'OT weekly holiday 200% (lieu)',
        'Change in lieu: Weekend-night Rate 270%': 'OT weekly holiday night 270% (lieu)',
        'Change in lieu: Holiday Rate 300%': 'OT public holiday 300% (lieu)',
        'Change in lieu: Holiday-night Rate 390%': 'OT public holiday night 390% (lieu)',
    }
    result = result.rename(columns=col_rename)
    col_order = [
        'No', 'Employee ID', 'Employee Name',
        'OT weekday 150%', 'OT weekday night 200%', 'OT weekly holiday 200%',
        'OT weekly holiday night 270%', 'OT public holiday 300%', 'OT public holiday night 390%',
        'OT weekday 150% (lieu)', 'OT weekday night 200% (lieu)', 'OT weekly holiday 200% (lieu)',
        'OT weekly holiday night 270% (lieu)', 'OT public holiday 300% (lieu)', 'OT public holiday night 390% (lieu)',
        'Transferred to normal working hours', 'Date', 'Total used hours in month',
        'Remain unused time off in lieu', 'Total OT paid'
    ]
    result = result[[c for c in col_order if c in result.columns]]
    cols = list(result.columns)
    rows = result.fillna('').astype(str).values.tolist()
    return jsonify({'columns': cols, 'data': rows})

@app.route('/get_total_attendance_detail')
def get_total_attendance_detail():
    emp_path = EMPLOYEE_LIST_PATH
    if os.path.exists(emp_path):
        emp_df = pd.read_csv(emp_path, dtype=str)
    else:
        emp_df = pd.DataFrame(columns=["Name", "ID Number", "Dept"])

    emp_df = emp_df.fillna('')
    emp_df['No'] = range(1, len(emp_df) + 1)

    # Đổi tên các cột theo mẫu hiển thị
    result = emp_df.rename(columns={
        'ID Number': '14 Digits Employee ID',
        'Name': "Employee's name",
        'Dept': 'Group'
    })

    # Sắp xếp theo Group nếu có
    if 'Group' in result.columns:
        result = result.sort_values(by=['Group', 'No'], ascending=[True, True])
        result = result.reset_index(drop=True)
        result['No'] = range(1, len(result) + 1)

    # Khởi tạo các cột tính toán với giá trị 0
    attendance_cols = [
        'Normal working days',
        'Annual leave (100% salary)',
        'Sick leave (50% salary)',
        'Unpaid leave (0% salary)',
        'Welfare leave (100% salary)',
        'Total',
        'Late/Leave early (mins)',
        'Late/Leave early (times)',
        'Forget scanning',
        'Violation',
        'Remark',
        'Attendance for salary payment'
    ]
    for col in attendance_cols:
        result[col] = 0.0

    # Lấy dữ liệu từ các file tạm
    signinout_data = []
    apply_data = []
    otlieu_data = []
    
    if os.path.exists(TEMP_SIGNINOUT_PATH):
        signinout_df = pd.read_excel(TEMP_SIGNINOUT_PATH)
        signinout_data = signinout_df.to_dict('records')
    
    if os.path.exists(TEMP_APPLY_PATH):
        apply_df = pd.read_excel(TEMP_APPLY_PATH)
        apply_data = apply_df.to_dict('records')
    
    if os.path.exists(TEMP_OTLIEU_PATH):
        otlieu_df = pd.read_excel(TEMP_OTLIEU_PATH)
        otlieu_data = otlieu_df.to_dict('records')

    # Lấy thông tin ngày đặc biệt từ rules
    holidays, special_weekends, special_workdays = get_special_days_from_rules(rules)

    # Xác định khoảng thời gian tính toán (20 tháng trước đến 19 tháng này)
    today = datetime.now()
    if today.day >= 20:
        # Nếu hôm nay từ ngày 20 trở đi, tính từ 20 tháng trước đến 19 tháng này
        start_date = today.replace(day=20) - timedelta(days=30)
        end_date = today.replace(day=19)
    else:
        # Nếu hôm nay trước ngày 20, tính từ 20 tháng trước đến 19 tháng trước
        start_date = today.replace(day=20) - timedelta(days=60)
        end_date = today.replace(day=19) - timedelta(days=30)
    
    # Tạo danh sách ngày trong khoảng thời gian
    date_range = pd.date_range(start=start_date, end=end_date, freq='D')

    # Hàm xác định loại ngày
    def get_day_type(dt, holidays, special_weekends, special_workdays):
        dt_date = dt.date()
        if dt_date in holidays:
            return 'Holiday'
        if dt_date in special_workdays:
            return 'Weekday'
        if dt_date in special_weekends:
            return 'Weekend'
        if dt.weekday() >= 5:  # Saturday = 5, Sunday = 6
            return 'Weekend'
        return 'Weekday'

    # Hàm xử lý matching tên nhân viên
    def extract_name_from_emp_name(emp_name):
        # Tách tên nhân viên từ format "Do Thi Thu Trang6970000006"
        import re
        match = re.match(r'^(.+?)(\d{7,10})$', emp_name.strip())
        if match:
            return match.group(1).strip()
        return emp_name.strip()

    # Hàm kiểm tra nhân viên có nghỉ Lieu ngày đó không
    def is_lieu_day(emp_name, check_date, otlieu_data):
        target_name = extract_name_from_emp_name(emp_name)
        for record in otlieu_data:
            name_val = str(record.get('Name', '') or '')
            if name_val.strip() == target_name:
                # Kiểm tra các cột Lieu From, Lieu To
                lieu_cols = ['Lieu From', 'Lieu To', 'Lieu From 2', 'Lieu To 2']
                for col in lieu_cols:
                    if col in record and pd.notna(record[col]) and str(record[col]).strip():
                        try:
                            lieu_date = pd.to_datetime(record[col]).date()
                            if lieu_date == check_date:
                                return True
                        except:
                            continue
        return False

    # Hàm kiểm tra có OT trong ngày không
    def has_ot_on_date(emp_name, check_date, otlieu_data):
        target_name = extract_name_from_emp_name(emp_name)
        for record in otlieu_data:
            name_val = str(record.get('Name', '') or '')
            if name_val.strip() == target_name:
                # Kiểm tra các cột OT Date, Date, OT date
                ot_date_cols = ['OT Date', 'Date', 'OT date']
                for col in ot_date_cols:
                    if col in record and pd.notna(record[col]) and str(record[col]).strip():
                        try:
                            ot_date = pd.to_datetime(record[col]).date()
                            if ot_date == check_date:
                                return True
                        except:
                            continue
        return False

    # Hàm kiểm tra có Apply Leave trong ngày không
    def has_apply_leave_on_date(emp_name, check_date, apply_data):
        target_name = extract_name_from_emp_name(emp_name)
        for record in apply_data:
            if (record.get('Name', '').strip() == target_name and 
                record.get('Type') == 'Leave' and 
                record.get('Results') == 'Approved'):
                
                try:
                    start_date_str = record.get('Start Date', '')
                    end_date_str = record.get('End Date', '')
                    apply_date_str = record.get('Apply Date', '')
                    
                    if start_date_str:
                        start_date = pd.to_datetime(start_date_str).date()
                    elif apply_date_str:
                        start_date = pd.to_datetime(apply_date_str).date()
                    else:
                        continue
                        
                    end_date = pd.to_datetime(end_date_str).date() if end_date_str else start_date
                    
                    # Kiểm tra check_date có trong khoảng start_date đến end_date không
                    if start_date <= check_date <= end_date:
                        return True
                except:
                    continue
        return False

    # Hàm xác định ca làm việc từ signinout data
    def get_shift_info(emp_name, check_date, signinout_data):
        """Trả về thông tin ca làm việc: 'AM', 'PM', 'FULL', hoặc 'NONE'"""
        if not signinout_data:
            return 'NONE'
            
        df_sign = pd.DataFrame(signinout_data)
        if 'emp_name' not in df_sign.columns or 'attendance_time' not in df_sign.columns:
            return 'NONE'
            
        df_sign['attendance_time'] = pd.to_datetime(df_sign['attendance_time'], errors='coerce')
        df_sign['date'] = df_sign['attendance_time'].dt.date
        
        target_name = extract_name_from_emp_name(emp_name)
        
        # Lọc bản ghi của nhân viên trong ngày
        mask = (df_sign['emp_name'].astype(str).str.strip() == target_name) & (df_sign['date'] == check_date)
        day_records = df_sign[mask]
        
        if day_records.empty:
            return 'NONE'
            
        # Xử lý dữ liệu trùng lặp - lấy thời gian sớm nhất cho mỗi loại (sáng/chiều)
        morning_records = day_records[day_records['attendance_time'].dt.hour < 12]
        afternoon_records = day_records[day_records['attendance_time'].dt.hour >= 12]
        
        morning_times = []
        afternoon_times = []
        
        if not morning_records.empty:
            # Lấy thời gian sớm nhất trong buổi sáng
            earliest_morning = morning_records['attendance_time'].min()
            morning_times.append(earliest_morning.time())
            
        if not afternoon_records.empty:
            # Lấy thời gian sớm nhất trong buổi chiều
            earliest_afternoon = afternoon_records['attendance_time'].min()
            afternoon_times.append(earliest_afternoon.time())
        
        if morning_times and afternoon_times:
            return 'FULL'  # Có cả sáng và chiều
        elif morning_times:
            return 'AM'    # Chỉ có sáng
        elif afternoon_times:
            return 'PM'    # Chỉ có chiều
        else:
            return 'NONE'

    # Tính toán cho từng nhân viên

    
    for idx, emp in result.iterrows():
        emp_name = emp["Employee's name"]
        normal_days = 0
        annual_leave = 0
        sick_leave = 0
        unpaid_leave = 0
        welfare_leave = 0

        # Tính Normal working days và các loại leave
        for dt in date_range:
            day_type = get_day_type(dt, holidays, special_weekends, special_workdays)
            dt_date = dt.date()
            
            # Sửa logic: Tính cho ngày làm việc bình thường theo yêu cầu
            # 1. Là ngày trong tuần (Thứ 2-6) VÀ không phải ngày nghỉ lễ
            # 2. Là ngày cuối tuần nhưng được quy định là "ngày làm việc đặc biệt"
            is_normal_workday = (
                (day_type == 'Weekday') or  # Ngày trong tuần không phải lễ
                (day_type == 'Weekend' and dt_date in special_workdays)  # Cuối tuần nhưng là ngày làm việc đặc biệt
            )
            
            if is_normal_workday:
                # Kiểm tra các điều kiện
                has_ot = has_ot_on_date(emp_name, dt_date, otlieu_data)
                has_lieu = is_lieu_day(emp_name, dt_date, otlieu_data)
                has_apply_leave = has_apply_leave_on_date(emp_name, dt_date, apply_data)
                
                # Xác định ca làm việc
                shift_info = get_shift_info(emp_name, dt_date, signinout_data)
                
    
                
                # Tính Normal working days
                if not has_ot and not has_lieu and not has_apply_leave:
                    if shift_info == 'FULL':
                        normal_days += 1.0  # Cả ngày
    
                    elif shift_info == 'AM' or shift_info == 'PM':
                        normal_days += 0.5  # Một ca


        # Tính các loại leave từ apply_data
        target_name = extract_name_from_emp_name(emp_name)
        for record in apply_data:
            if (record.get('Name', '').strip() == target_name and 
                record.get('Type') == 'Leave' and 
                record.get('Results') == 'Approved'):
                
                try:
                    start_date_str = record.get('Start Date', '')
                    end_date_str = record.get('End Date', '')
                    apply_date_str = record.get('Apply Date', '')
                    
                    if start_date_str:
                        start_date = pd.to_datetime(start_date_str).date()
                    elif apply_date_str:
                        start_date = pd.to_datetime(apply_date_str).date()
                    else:
                        continue
                        
                    end_date = pd.to_datetime(end_date_str).date() if end_date_str else start_date
                    
                    # Duyệt qua từng ngày trong khoảng thời gian
                    current_date = start_date
                    while current_date <= end_date:
                        if current_date in date_range:
                            day_type = get_day_type(pd.to_datetime(current_date), holidays, special_weekends, special_workdays)
                            if day_type == 'Weekday':
                                leave_type = str(record.get('Leave Type', '')).lower()
                                note = str(record.get('Note', '')).lower()
                                
                                # Xác định ca nghỉ từ Note
                                if any(keyword in note for keyword in ['morning', 'sáng', '上午', 'am']):
                                    leave_days = 0.5  # Nghỉ buổi sáng
                                elif any(keyword in note for keyword in ['afternoon', 'chiều', '下午', 'pm']):
                                    leave_days = 0.5  # Nghỉ buổi chiều
                                elif start_date == end_date:
                                    leave_days = 1.0  # Nghỉ cả ngày
                                else:
                                    leave_days = 1.0  # Nghỉ nhiều ngày, mỗi ngày = 1.0
                                
                                # Phân loại theo loại leave
                                leave_type_lower = leave_type.lower()
                                if 'annual' in leave_type_lower:
                                    annual_leave += leave_days
                                elif 'sick' in leave_type_lower:
                                    sick_leave += leave_days
                                elif 'unpaid' in leave_type_lower:
                                    unpaid_leave += leave_days
                                elif 'welfare' in leave_type_lower:
                                    welfare_leave += leave_days
                        
                        current_date += timedelta(days=1)
                except:
                    continue

        # Tính các cột violation từ signinout_data
        late_early_mins = 0
        late_early_times = 0
        forget_scanning = 0
        violation = 0

        # Chuẩn bị DataFrame từ signinout_data
        if signinout_data:
            df_sign = pd.DataFrame(signinout_data)
            # Đảm bảo các cột cần thiết
            if 'Name' in df_sign.columns and 'attendance_time' in df_sign.columns:
                df_sign['attendance_time'] = pd.to_datetime(df_sign['attendance_time'], errors='coerce')
                df_sign['date'] = df_sign['attendance_time'].dt.date
            else:
                df_sign = pd.DataFrame(columns=['Name', 'attendance_time', 'date'])
        else:
            df_sign = pd.DataFrame(columns=['Name', 'attendance_time', 'date'])

        for dt in date_range:
            day_type = get_day_type(dt, holidays, special_weekends, special_workdays)
            # Chỉ tính violation cho ngày làm việc bình thường
            is_normal_workday = (
                (day_type == 'Weekday') or  # Ngày trong tuần không phải lễ
                (day_type == 'Weekend' and dt.date() in special_workdays)  # Cuối tuần nhưng là ngày làm việc đặc biệt
            )
            
            if is_normal_workday:
                # Lấy tất cả bản ghi của nhân viên này trong ngày dt
                # Chuyển signinout_data từ list of dicts thành DataFrame
                df_sign = pd.DataFrame(signinout_data)
                if 'emp_name' in df_sign.columns and 'attendance_time' in df_sign.columns:
                    df_sign['attendance_time'] = pd.to_datetime(df_sign['attendance_time'], errors='coerce')
                    df_sign['date'] = df_sign['attendance_time'].dt.date
                    mask = (df_sign['emp_name'].astype(str).str.strip() == target_name) & (df_sign['date'] == dt.date())
                    day_records = df_sign[mask]
                else:
                    day_records = pd.DataFrame()

                if not day_records.empty:
                    # Xử lý dữ liệu trùng lặp - lấy thời gian sớm nhất và muộn nhất
                    morning_records = day_records[day_records['attendance_time'].dt.hour < 12]
                    afternoon_records = day_records[day_records['attendance_time'].dt.hour >= 12]
                    
                    in_time = None
                    out_time = None
                    
                    if not morning_records.empty:
                        in_time = morning_records['attendance_time'].min()  # Thời gian sớm nhất buổi sáng
                    
                    if not afternoon_records.empty:
                        out_time = afternoon_records['attendance_time'].max()  # Thời gian muộn nhất buổi chiều
                    
                    # Nếu không có dữ liệu sáng/chiều rõ ràng, lấy min/max của cả ngày
                    if in_time is None and out_time is None:
                        in_time = day_records['attendance_time'].min()
                        out_time = day_records['attendance_time'].max()
                    elif in_time is None:
                        in_time = out_time
                    elif out_time is None:
                        out_time = in_time
                    
                    # Đi muộn
                    if in_time.hour > 8 or (in_time.hour == 8 and in_time.minute > 30):
                        late_minutes = (in_time.hour - 8) * 60 + (in_time.minute - 30)
                        late_early_mins += late_minutes
                        late_early_times += 1
                    # Về sớm
                    if out_time.hour < 17 or (out_time.hour == 17 and out_time.minute < 30):
                        early_minutes = (17 - out_time.hour) * 60 + (30 - out_time.minute)
                        late_early_mins += early_minutes
                        late_early_times += 1
                else:
                    # Không có bản ghi nào trong ngày => quên quẹt
                    forget_scanning += 1

        # Cập nhật kết quả
        result.at[idx, 'Normal working days'] = normal_days
        result.at[idx, 'Annual leave (100% salary)'] = annual_leave
        result.at[idx, 'Sick leave (50% salary)'] = sick_leave
        result.at[idx, 'Unpaid leave (0% salary)'] = unpaid_leave
        result.at[idx, 'Welfare leave (100% salary)'] = welfare_leave
        result.at[idx, 'Late/Leave early (mins)'] = late_early_mins
        result.at[idx, 'Late/Leave early (times)'] = late_early_times
        result.at[idx, 'Forget scanning'] = forget_scanning
        result.at[idx, 'Violation'] = violation
        

        
        # Tính Total
        total_leave = annual_leave + sick_leave + unpaid_leave + welfare_leave
        result.at[idx, 'Total'] = total_leave
        
        # Tính Attendance for salary payment (tổng ngày làm việc + nghỉ phép)
        attendance_for_salary = normal_days + total_leave
        result.at[idx, 'Attendance for salary payment'] = attendance_for_salary

    # Xác định lại thứ tự cột giống giao diện
    col_order = [
        'No',
        '14 Digits Employee ID',
        "Employee's name",
        'Group',
        'Normal working days',
        'Annual leave (100% salary)',
        'Sick leave (50% salary)',
        'Unpaid leave (0% salary)',
        'Welfare leave (100% salary)',
        'Total',
        'Late/Leave early (mins)',
        'Late/Leave early (times)',
        'Forget scanning',
        'Violation',
        'Remark',
        'Attendance for salary payment'
    ]
    result = result[[c for c in col_order if c in result.columns]]

    # Chuyển dữ liệu về dạng list để trả ra frontend
    cols = list(result.columns)
    rows = result.fillna('').astype(str).values.tolist()

    return jsonify({'columns': cols, 'data': rows})

@app.route('/get_attendance_report')
def get_attendance_report():
    # Lấy dữ liệu từ các file tạm trước
    signinout_data,apply_data,otlieu_data = [], [], []
    
    if os.path.exists(TEMP_SIGNINOUT_PATH):
        signinout_df = pd.read_excel(TEMP_SIGNINOUT_PATH)
        signinout_data = signinout_df.to_dict('records')
    
    if os.path.exists(TEMP_APPLY_PATH):
        apply_df = pd.read_excel(TEMP_APPLY_PATH)
        apply_data = apply_df.to_dict('records')
    
    if os.path.exists(TEMP_OTLIEU_PATH):
        otlieu_df = pd.read_excel(TEMP_OTLIEU_PATH)
        otlieu_data = otlieu_df.to_dict('records')

    month, year = 7, 2024
    if signinout_data:
        dates = [pd.to_datetime(r['attendance_time']) for r in signinout_data if pd.notna(r.get('attendance_time'))]
        if dates:
            month_counts = {}
            for date in dates:
                month_key = (date.month, date.year)
                month_counts[month_key] = month_counts.get(month_key, 0) + 1
            most_common_month = max(month_counts.items(), key=lambda x: x[1])[0]
            month, year = most_common_month

    # Load employee list
    emp_df = pd.read_csv(EMPLOYEE_LIST_PATH, dtype=str) if os.path.exists(EMPLOYEE_LIST_PATH) else pd.DataFrame(columns=["Dept", "Name"])
    if 'Dept' in emp_df.columns:
        emp_df = emp_df.sort_values(by=['Dept', 'Name']).reset_index(drop=True)

    start_date = pd.Timestamp(year, month, 1) - pd.DateOffset(months=1, day=20)
    end_date = pd.Timestamp(year, month, 19)
    days = pd.date_range(start=start_date, end=end_date, freq='D')
    day_cols = [d.strftime('%Y-%m-%d') for d in days]
    
    holidays, special_weekends, special_workdays = get_special_days_from_rules(rules)
    
    def normalize_name(name_field):
        import re
        if not isinstance(name_field, str):
            return ""
        # Xóa ID ở cuối nếu có (ví dụ: 10046198)
        name_only = re.sub(r'\d{8,}$', '', name_field).strip()
        # Chuyển thành chữ thường và xóa khoảng trắng thừa
        return name_only.lower()


    def get_day_type(dt, holidays, special_weekends, special_workdays):
        dt_date = dt.date()
        if dt_date in holidays:
            return 'Holiday'
        if dt_date in special_workdays:
            return 'Weekday'
        if dt_date in special_weekends:
            return 'Weekend'
        if dt.weekday() >= 5:  # Saturday = 5, Sunday = 6
            return 'Weekend'
        return 'Weekday'

    def extract_name_from_emp_name(emp_name):
        import re
        match = re.match(r'^(.+?)(\d{7,10})$', emp_name.strip())
        if match:
            return match.group(1).strip()
        return emp_name.strip()

    def is_lieu_day(emp_name, check_date, otlieu_data):
        # emp_name is already normalized
        target_name = emp_name.lower().strip()
        for record in otlieu_data:
            name_val = str(record.get('Name', '') or '')
            normalized_name_val = normalize_name(name_val)
            if normalized_name_val == target_name:
                lieu_cols = ['Lieu From', 'Lieu To', 'Lieu From 2', 'Lieu To 2']
                for col in lieu_cols:
                    if col in record and pd.notna(record[col]) and str(record[col]).strip():
                        try:
                            lieu_date = pd.to_datetime(record[col]).date()
                            if lieu_date == check_date:
                                return True
                        except:
                            continue
        return False

    def get_apply_info_for_date(emp_name, check_date, apply_data):
        """Trả về thông tin apply cho ngày: type, leave_type, is_approved"""
        # emp_name is already normalized
        target_name = emp_name.lower().strip()
        for record in apply_data:
            name_val = str(record.get('Name', '') or '')
            normalized_name_val = normalize_name(name_val)
            if (normalized_name_val == target_name and 
                record.get('Results') == 'Approved'):
                
                try:
                    start_date_str = record.get('Start Date', '')
                    end_date_str = record.get('End Date', '')
                    apply_date_str = record.get('Apply Date', '')
                    
                    if start_date_str:
                        start_date = pd.to_datetime(start_date_str).date()
                    elif apply_date_str:
                        start_date = pd.to_datetime(apply_date_str).date()
                    else:
                        continue
                        
                    end_date = pd.to_datetime(end_date_str).date() if end_date_str else start_date
                    
                    if start_date <= check_date <= end_date:
                        return {
                            'type': record.get('Type', ''),
                            'leave_type': record.get('Leave Type', ''),
                            'is_approved': record.get('Results') == 'Approved'
                        }
                except:
                    continue
        return None

    # --- PHẦN 2: CHUẨN HÓA VÀ XÂY DỰNG BẢNG DỮ LIỆU GỐC ---

    # Bước 1: Chuẩn hóa file danh sách nhân viên
    if os.path.exists(EMPLOYEE_LIST_PATH):
        emp_df = pd.read_csv(EMPLOYEE_LIST_PATH, dtype=str)
        emp_df.dropna(subset=['Name'], inplace=True)
        emp_df['NormalizedName'] = emp_df['Name'].apply(normalize_name)
        # Giữ lại tên gốc để hiển thị
        emp_df['DisplayName'] = emp_df['Name'].apply(lambda x: re.sub(r'\d{8,}$', '', x).strip())
    else:
        return jsonify({'error': 'Không tìm thấy file danh sách nhân viên.'})

    # Bước 2: Chuẩn hóa file chấm công
    if not signinout_data:
        return jsonify({'error': 'Không có dữ liệu chấm công.'})
    
    df_sign = pd.DataFrame(signinout_data)
    df_sign.dropna(subset=['attendance_time', 'emp_name'], inplace=True)
    df_sign['attendance_time'] = pd.to_datetime(df_sign['attendance_time'], errors='coerce')
    df_sign['Date'] = df_sign['attendance_time'].dt.date
    df_sign['NormalizedName'] = df_sign['emp_name'].apply(normalize_name)

    # Bước 3: Tổng hợp giờ vào/ra bằng tên đã chuẩn hóa
    daily_times = df_sign.groupby(['NormalizedName', 'Date'])['attendance_time'].agg(
        SignIn=('min'),
        SignOut=('max')
    ).reset_index()

    # Bước 4: Tạo bảng master và hợp nhất bằng khóa chuẩn hóa
    days = pd.date_range(start=(pd.Timestamp(year, month, 1) - pd.DateOffset(months=1)).replace(day=20), 
                         end=pd.Timestamp(year, month, 19), freq='D')
    
    master_df = pd.DataFrame([
        (row['DisplayName'], row['NormalizedName'], day.date())
        for _, row in emp_df.iterrows()
        for day in days
    ], columns=['DisplayName', 'NormalizedName', 'Date'])

    master_df = pd.merge(master_df, emp_df[['NormalizedName', 'Dept']], on='NormalizedName', how='left')
    master_df = pd.merge(master_df, daily_times, on=['NormalizedName', 'Date'], how='left')

    # (Chuẩn hóa các file khác nếu cần dùng, ví dụ Apply Data)
    if apply_data:
        apply_df = pd.DataFrame(apply_data)
        apply_df['NormalizedName'] = apply_df['Emp Name'].apply(normalize_name)

    processed_records = []
    abnormal_report_data = []

    for _, row in master_df.iterrows():
        record = row.to_dict()
        emp_name, day_date = record['DisplayName'], record['Date']
        day_pd_ts = pd.Timestamp(day_date)

        record['DayType'] = get_day_type(day_pd_ts, holidays, special_weekends, special_workdays)
        record['IsLieu'] = is_lieu_day(record['NormalizedName'], day_date, otlieu_data)
        record['ApplyInfo'] = get_apply_info_for_date(record['NormalizedName'], day_date, apply_data)
        
        status, late_minutes = '', 0

        if record['DayType'] != 'Weekday' or record['IsLieu']:
            status = 'Lieu' if record['IsLieu'] else 'Off'
        elif record['ApplyInfo'] and record['ApplyInfo']['is_approved']:
            apply_type = record['ApplyInfo']['type']
            if apply_type == 'Supplement':
                # Nếu có đơn Supplement, kiểm tra chấm công
                if pd.isna(record.get('SignIn')) and pd.isna(record.get('SignOut')):
                    status = 'Supplement'  # Không có chấm công, làm bổ sung
                elif pd.isna(record.get('SignIn')) or pd.isna(record.get('SignOut')):
                    status = 'Supplement'  # Thiếu chấm công, làm bổ sung
                else:
                    # Có đầy đủ chấm công, xử lý theo logic bình thường
                    status = 'Normal'  # Sẽ được xử lý tiếp ở phần else
            else:
                # Các loại đơn khác (Trip, Leave)
                status = apply_type
        elif pd.isna(record.get('SignIn')) and pd.isna(record.get('SignOut')):
            status = 'Miss' # Không có chấm công nào
        elif pd.isna(record.get('SignIn')) or pd.isna(record.get('SignOut')):
            status = 'Miss' # Thiếu SignIn hoặc SignOut
        else:
            status = 'Normal'
            check_in_time, check_out_time = record['SignIn'], record['SignOut']
            
            if pd.notna(check_in_time) and pd.notna(check_out_time):
                work_seconds = (check_out_time - check_in_time).total_seconds()
                
                lunch_start = check_in_time.replace(hour=12, minute=0, second=0)
                lunch_end = check_in_time.replace(hour=13, minute=30, second=0)
                if check_in_time < lunch_end and check_out_time > lunch_start:
                    work_seconds -= 1.5 * 3600

                work_hours = work_seconds / 3600
                
                if work_hours < 8.0:
                    late_minutes = int((8.0 - work_hours) * 60)
                    status = 'Late/Soon' 
                    
                    abnormal_report_data.append({
                        'Department': record.get('Dept', ''), 'Name': emp_name, 'Date': day_date,
                        'SignIn': check_in_time.strftime('%H:%M'), 'SignOut': check_out_time.strftime('%H:%M'),
                        'Reason': 'Late/Soon', 'Minutes': late_minutes
                    })

        record['Status'] = status
        record['LateMinutes'] = late_minutes
        # Ensure we have the normalized name for later processing
        record['NormalizedName'] = record.get('NormalizedName', '')
        processed_records.append(record)

        processed_df = pd.DataFrame(processed_records)

    result_rows = []
    summary_keys = ['Normal', 'Leave', 'Trip', 'Miss', 'Late/Soon', 'Lieu', 'Supplement', 'TotalLateMinutes']

    for _, emp in emp_df.iterrows():
        emp_name_display = emp['DisplayName']
        emp_name_normalized = emp['NormalizedName']
        emp_data = processed_df[processed_df['NormalizedName'] == emp_name_normalized]
        
        morning_row = {'Department': emp.get('Dept', ''), 'Name': emp_name_display, 'Shift': 'Morning shift'}
        afternoon_row = {'Department': emp.get('Dept', ''), 'Name': emp_name_display, 'Shift': 'Afternoon shift'}
        
        summary = {key: 0 for key in summary_keys}

        for day in days:
            day_str = day.strftime('%Y-%m-%d')
            day_record_df = emp_data[emp_data['Date'] == day.date()]

            if not day_record_df.empty:
                record = day_record_df.iloc[0]
                status = record['Status']
                
                if status in ['Trip', 'Leave', 'Supplement', 'Miss']:
                    morning_row[day_str], afternoon_row[day_str] = status, status
                    summary[status] += 0.5
                elif status == 'Lieu':
                    morning_row[day_str], afternoon_row[day_str] = 'Lieu', 'Lieu'
                    summary['Lieu'] += 0.5
                elif status == 'Off':
                    morning_row[day_str], afternoon_row[day_str] = '', ''
                elif status in ['Normal', 'Late/Soon']:
                    # Trả về thời gian với prefix để frontend phân biệt
                    morning_time = pd.to_datetime(record['SignIn']).strftime('%H:%M') if pd.notna(record['SignIn']) else '0'
                    afternoon_time = pd.to_datetime(record['SignOut']).strftime('%H:%M') if pd.notna(record['SignOut']) else '0'
                    
                    if status == 'Late/Soon':
                        morning_row[day_str] = f"LATE:{morning_time}"
                        afternoon_row[day_str] = f"LATE:{afternoon_time}"
                    else:
                        morning_row[day_str] = f"NORMAL:{morning_time}"
                        afternoon_row[day_str] = f"NORMAL:{afternoon_time}"
                    
                    summary[status] += 0.5
                
                summary['TotalLateMinutes'] += record['LateMinutes']

        morning_row.update(summary)
        
        result_rows.append(morning_row)
        result_rows.append(afternoon_row)

    result = pd.DataFrame(result_rows)
    columns = ['Department', 'Name', 'Shift'] + day_cols + list(summary.keys())
    for col in columns:
        if col not in result.columns:
            result[col] = ''
    result = result[columns]
    
    cols = list(result.columns)
    rows = result.fillna('').astype(str).values.tolist()

    abnormal_cols = list(abnormal_report_data[0].keys()) if abnormal_report_data else []
    abnormal_rows = [list(d.values()) for d in abnormal_report_data]

    return jsonify({
        'columns': cols, 'rows': rows,
        'abnormal_columns': abnormal_cols, 'abnormal_rows': abnormal_rows
    })

def flatten_cell(cell):
    if isinstance(cell, dict) and 'value' in cell:
        return cell['value']
    return cell

def get_special_days_from_rules(rules):
    holidays = set()
    special_weekends = set()
    special_workdays = set()
    if rules is not None:
        if 'Holiday Date in This Year' in rules.columns:
            holidays = set(pd.to_datetime(rules['Holiday Date in This Year'], errors='coerce').dt.date.dropna())
        if 'Special Weekend' in rules.columns:
            special_weekends = set(pd.to_datetime(rules['Special Weekend'], errors='coerce').dt.date.dropna())
        if 'Special Work Day' in rules.columns:
            special_workdays = set(pd.to_datetime(rules['Special Work Day'], errors='coerce').dt.date.dropna())
    return holidays, special_weekends, special_workdays

def is_intern(emp_id, emp_list_df):
    if emp_id in emp_list_df['ID Number'].values:
        emp_row = emp_list_df[emp_list_df['ID Number'] == emp_id].iloc[0]
        return emp_row.get('Internship', '') == 'Intern'
    return False

# ========================
# MULTI-FILE IMPORT ROUTES
# ========================
@app.route('/import_multiple_files', methods=['POST'])
def import_multiple_files():
    """Import multiple files for a specific data type with append functionality"""
    global sign_in_out_data, apply_data, ot_lieu_data, employee_list_df
    
    data_type = request.form.get('data_type')
    if not data_type:
        return jsonify({'error': 'Data type is required'}), 400
    
    if 'files[]' not in request.files:
        return jsonify({'error': 'No files uploaded'}), 400
    
    files = request.files.getlist('files[]')
    if not files or all(f.filename == '' for f in files):
        return jsonify({'error': 'No files selected'}), 400
    
    total_rows = 0
    imported_files = []
    
    try:
        for file in files:
            if file.filename == '':
                continue
                
            if not file.filename.lower().endswith(('.xlsx', '.xls', '.csv')):
                continue
                
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)
            
            # Load data based on file type
            if file.filename.lower().endswith('.csv'):
                df = try_read_csv(open(file_path, 'rb').read())
            elif file.filename.lower().endswith('.xls'):
                try:
                    # Thử với xlrd engine trước
                    df = pd.read_excel(file_path, engine='xlrd')
                except Exception as e:
                    print(f"xlrd engine failed for .xls multiple import: {e}")
                    try:
                        # Fallback: thử với openpyxl engine
                        df = pd.read_excel(file_path, engine='openpyxl')
                    except Exception as e2:
                        print(f"openpyxl engine also failed for .xls multiple import: {e2}")
                        return jsonify({'error': f'Cannot read .xls file {filename}. Please convert to .xlsx format. Error: {str(e)}'}), 400
            else:
                df = pd.read_excel(file_path)
            
            df = df.dropna(how='all').fillna('')
            
            # Process data based on type
            if data_type == 'signinout':
                keep = [col for col in df.columns if col.lower() in ['emp_name', 'attendance_time']]
                df = df[keep]
                if sign_in_out_data is None or sign_in_out_data.empty:
                    sign_in_out_data = df
                else:
                    sign_in_out_data = pd.concat([sign_in_out_data, df], ignore_index=True)
                total_rows = len(sign_in_out_data)
                
            elif data_type == 'apply':
                df = translate_apply_headers(df)
                df = filter_apply_employees(df, employee_list_df)
                df = add_apply_columns(df)
                if apply_data is None or apply_data.empty:
                    apply_data = df
                else:
                    apply_data = pd.concat([apply_data, df], ignore_index=True)
                total_rows = len(apply_data)
                
            elif data_type == 'otlieu':
                if 'OT From Note: 12AM is midnight' in df.columns:
                    df = df.rename(columns={'OT From Note: 12AM is midnight': 'OT From'})
                df = process_ot_lieu_df(df, employee_list_df)
                if ot_lieu_data is None or ot_lieu_data.empty:
                    ot_lieu_data = df
                else:
                    ot_lieu_data = pd.concat([ot_lieu_data, df], ignore_index=True)
                total_rows = len(ot_lieu_data)
            
            imported_files.append(filename)
        
        return jsonify({
            'success': True,
            'message': f'Successfully imported {len(imported_files)} files. Total rows: {total_rows}',
            'imported_files': imported_files,
            'total_rows': total_rows
        })
        
    except Exception as e:
        return jsonify({'error': f'Error importing files: {str(e)}'}), 400

@app.route('/remove_file_from_data', methods=['POST'])
def remove_file_from_data():
    """Remove specific file data from the current dataset"""
    global sign_in_out_data, apply_data, ot_lieu_data
    
    data_type = request.form.get('data_type')
    file_index = request.form.get('file_index')
    
    if not data_type or file_index is None:
        return jsonify({'error': 'Data type and file index are required'}), 400
    
    try:
        file_index = int(file_index)
        if data_type == 'signinout':
            sign_in_out_data = None
        elif data_type == 'apply':
            apply_data = None
        elif data_type == 'otlieu':
            ot_lieu_data = None
        
        return jsonify({
            'success': True,
            'message': f'{data_type.capitalize()} data cleared successfully'
        })
        
    except Exception as e:
        return jsonify({'error': f'Error removing file: {str(e)}'}), 400

@app.route('/get_current_data_info', methods=['GET'])
def get_current_data_info():
    """Get information about currently loaded data"""
    data_type = request.args.get('data_type')
    
    if data_type == 'signinout':
        count = len(sign_in_out_data) if sign_in_out_data is not None else 0
        data = sign_in_out_data
    elif data_type == 'apply':
        count = len(apply_data) if apply_data is not None else 0
        data = apply_data
    elif data_type == 'otlieu':
        count = len(ot_lieu_data) if ot_lieu_data is not None else 0
        data = ot_lieu_data
    else:
        return jsonify({'error': 'Invalid data type'}), 400
    
    columns = list(data.columns) if data is not None else []
    
    return jsonify({
        'data_type': data_type,
        'row_count': count,
        'has_data': count > 0,
        'columns': columns
    })

@app.route('/save_apply_changes', methods=['POST'])
def save_apply_changes():
    """Save changes to Apply data"""
    global apply_data
    try:
        data = request.json.get('data', [])
        if data:
            df = pd.DataFrame(data)
            apply_data = df
            apply_data.to_excel(TEMP_APPLY_PATH, index=False)
            return jsonify({'success': True, 'message': 'Apply data saved successfully'})
        else:
            return jsonify({'success': False, 'error': 'No data provided'})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/save_otlieu_changes', methods=['POST'])
def save_otlieu_changes():
    """Save changes to OT Lieu data"""
    global ot_lieu_data
    try:
        data = request.json.get('data', [])
        if data:
            df = pd.DataFrame(data)
            ot_lieu_data = df
            ot_lieu_save = ot_lieu_data.applymap(flatten_cell)
            ot_lieu_save.to_excel(TEMP_OTLIEU_PATH, index=False)
            return jsonify({'success': True, 'message': 'OT Lieu data saved successfully'})
        else:
            return jsonify({'success': False, 'error': 'No data provided'})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/save_signinout_changes', methods=['POST'])
def save_signinout_changes():
    """Save changes to Sign In/Out data"""
    global sign_in_out_data
    try:
        data = request.json.get('data', [])
        if data:
            df = pd.DataFrame(data)
            sign_in_out_data = df
            sign_in_out_data.to_excel(TEMP_SIGNINOUT_PATH, index=False)
            return jsonify({'success': True, 'message': 'Sign In/Out data saved successfully'})
        else:
            return jsonify({'success': False, 'error': 'No data provided'})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

# ========================
# EXPORT ROUTES
# ========================
def calculate_otlieu_report_for_export():
    """Calculate OT Lieu Report data for export (without request context)"""
    global ot_lieu_data, employee_list_df, rules
    
    if ot_lieu_data is None or ot_lieu_data.empty:
        temp_otlieu_path = os.path.join(app.config['UPLOAD_FOLDER'], 'temp_otlieu.xlsx')
        if os.path.exists(temp_otlieu_path):
            ot_lieu_data = pd.read_excel(temp_otlieu_path)
        else:
            return {'columns': [], 'rows': []}
    
    if employee_list_df is None or employee_list_df.empty:
        emp_path = os.path.join(app.config['UPLOAD_FOLDER'], 'employee_list.csv')
        if os.path.exists(emp_path):
            employee_list_df = pd.read_csv(emp_path, dtype=str)
        else:
            return {'columns': [], 'rows': []}
    
    if rules is None or rules.empty:
        rules_path = os.path.join(app.config['UPLOAD_FOLDER'], 'rules.xlsx')
        if os.path.exists(rules_path):
            rules = pd.read_excel(rules_path)
        else:
            return {'columns': [], 'rows': []}
    
    try:
        # Get lieu followup data
        lieu_followup_path = os.path.join(app.config['UPLOAD_FOLDER'], 'lieu_followup.xlsx')
        lieu_followup_df = pd.DataFrame()
        if os.path.exists(lieu_followup_path):
            lieu_followup_df = pd.read_excel(lieu_followup_path)
        
        # Calculate OT Lieu Before first
        df = calculate_otlieu_before()
        
        if df is None or df.empty:
            return {'columns': [], 'rows': []}
        
        # Group by employee
        result_rows = []
        
        for name in df['Name'].unique():
            if pd.isna(name) or str(name).strip() == '':
                continue
                
            emp_data = df[df['Name'] == name]
            
            # Get lieu remain from followup
            lieu_remain_prev = 0.0
            if not lieu_followup_df.empty and 'Name' in lieu_followup_df.columns:
                followup_row = lieu_followup_df[lieu_followup_df['Name'].astype(str).str.strip() == str(name).strip()]
                if not followup_row.empty:
                    lieu_remain_prev = float(followup_row.iloc[0].get('Lieu remain', 0.0))
            
            # Calculate totals
            total_ot_paid = 0.0
            total_used_hours = 0.0
            
            # Sum OT Payment columns
            ot_payment_cols = [col for col in df.columns if col.startswith('OT Payment:')]
            for col in ot_payment_cols:
                total_ot_paid += emp_data[col].sum()
            
            # Sum Change in lieu columns
            lieu_cols = [col for col in df.columns if col.startswith('Change in lieu:')]
            for col in lieu_cols:
                total_used_hours += emp_data[col].sum()
            
            # Calculate transferred hours
            transferred_hours = 0.0
            if total_used_hours > 0 and total_ot_paid == 0:
                transferred_hours = total_used_hours
            elif total_ot_paid > 25:
                transferred_hours = total_ot_paid - 25
            
            # Calculate remain unused
            remain_unused = lieu_remain_prev - total_used_hours
            
            row = {
                'Name': name,
                'Lieu remain previous month': round(lieu_remain_prev, 3),
                'Total used hours in month': round(total_used_hours, 3),
                'Remain unused time off in lieu': round(remain_unused, 3),
                'Total OT paid': round(total_ot_paid, 3),
                'Transferred to normal working hours': round(transferred_hours, 3)
            }
            
            # Add individual OT Payment columns
            for col in ot_payment_cols:
                row[col] = round(emp_data[col].sum(), 3)
            
            # Add individual Change in lieu columns
            for col in lieu_cols:
                row[col] = round(emp_data[col].sum(), 3)
            
            result_rows.append(row)
        
        if not result_rows:
            return {'columns': [], 'rows': []}
        
        result_df = pd.DataFrame(result_rows)
        cols = list(result_df.columns)
        rows = result_df.fillna('').astype(str).values.tolist()
        
        return {'columns': cols, 'rows': rows}
        
    except Exception as e:
        print(f"Error in calculate_otlieu_report_for_export: {e}")
        return {'columns': [], 'rows': []}

def calculate_total_attendance_detail_for_export():
    """Calculate Total Attendance Detail data for export (without request context)"""
    global sign_in_out_data, apply_data, ot_lieu_data, employee_list_df, rules
    
    # Load data from temp files if global variables are empty
    if sign_in_out_data is None or sign_in_out_data.empty:
        temp_signinout_path = os.path.join(app.config['UPLOAD_FOLDER'], 'temp_signinout.xlsx')
        if os.path.exists(temp_signinout_path):
            sign_in_out_data = pd.read_excel(temp_signinout_path)
        else:
            return {'columns': [], 'rows': []}
    
    if apply_data is None or apply_data.empty:
        temp_apply_path = os.path.join(app.config['UPLOAD_FOLDER'], 'temp_apply.xlsx')
        if os.path.exists(temp_apply_path):
            apply_data = pd.read_excel(temp_apply_path)
        else:
            return {'columns': [], 'rows': []}
    
    if ot_lieu_data is None or ot_lieu_data.empty:
        temp_otlieu_path = os.path.join(app.config['UPLOAD_FOLDER'], 'temp_otlieu.xlsx')
        if os.path.exists(temp_otlieu_path):
            ot_lieu_data = pd.read_excel(temp_otlieu_path)
        else:
            return {'columns': [], 'rows': []}
    
    if employee_list_df is None or employee_list_df.empty:
        emp_path = os.path.join(app.config['UPLOAD_FOLDER'], 'employee_list.csv')
        if os.path.exists(emp_path):
            employee_list_df = pd.read_csv(emp_path, dtype=str)
        else:
            return {'columns': [], 'rows': []}
    
    if rules is None or rules.empty:
        rules_path = os.path.join(app.config['UPLOAD_FOLDER'], 'rules.xlsx')
        if os.path.exists(rules_path):
            rules = pd.read_excel(rules_path)
        else:
            return {'columns': [], 'rows': []}
    
    try:
        # Get current month/year
        current_date = datetime.now()
        month = current_date.month
        year = current_date.year
        
        # Calculate date range: 20th of previous month to 19th of current month
        if month == 1:
            prev_month = 12
            prev_year = year - 1
        else:
            prev_month = month - 1
            prev_year = year
        
        start_date = pd.Timestamp(prev_year, prev_month, 20)
        end_date = pd.Timestamp(year, month, 19)
        date_range = pd.date_range(start=start_date, end=end_date, freq='D')
        
        # Get special days from rules
        holidays, special_weekends, special_workdays = get_special_days_from_rules(rules)
        
        # Helper functions
        def get_day_type(dt, holidays, special_weekends, special_workdays):
            dt_date = dt.date()
            if dt_date in holidays:
                return 'Holiday'
            if dt_date in special_workdays:
                return 'Weekday'
            if dt_date in special_weekends:
                return 'Weekend'
            if dt.weekday() >= 5:  # Saturday = 5, Sunday = 6
                return 'Weekend'
            return 'Weekday'
        
        def extract_name_from_emp_name(emp_name):
            # Tách tên nhân viên từ format "Do Thi Thu Trang6970000006"
            import re
            match = re.match(r'^(.+?)(\d{7,10})$', emp_name.strip())
            if match:
                return match.group(1).strip()
            return emp_name.strip()
        
        def is_lieu_day(emp_name, check_date, otlieu_data):
            if otlieu_data is None or otlieu_data.empty:
                return False
            target_name = extract_name_from_emp_name(emp_name)
            for _, record in otlieu_data.iterrows():
                name_val = str(record.get('Name', '') or '')
                if name_val.strip() == target_name:
                    lieu_cols = ['Lieu From', 'Lieu To', 'Lieu From 2', 'Lieu To 2']
                    for col in lieu_cols:
                        if col in record and pd.notna(record[col]) and str(record[col]).strip():
                            try:
                                lieu_date = pd.to_datetime(record[col]).date()
                                if lieu_date == check_date:
                                    return True
                            except:
                                continue
            return False
        
        def has_ot_on_date(emp_name, check_date, otlieu_data):
            if otlieu_data is None or otlieu_data.empty:
                return False
            target_name = extract_name_from_emp_name(emp_name)
            for _, record in otlieu_data.iterrows():
                name_val = str(record.get('Name', '') or '')
                if name_val.strip() == target_name:
                    ot_date_cols = ['OT Date', 'Date', 'OT date']
                    for col in ot_date_cols:
                        if col in record and pd.notna(record[col]) and str(record[col]).strip():
                            try:
                                ot_date = pd.to_datetime(record[col]).date()
                                if ot_date == check_date:
                                    return True
                            except:
                                continue
            return False
        
        def has_apply_leave_on_date(emp_name, check_date, apply_data):
            if apply_data is None or apply_data.empty:
                return False
            target_name = extract_name_from_emp_name(emp_name)
            for _, record in apply_data.iterrows():
                if (record.get('Name', '').strip() == target_name and 
                    record.get('Results') == 'Approved' and
                    record.get('Type') == 'Leave'):
                    
                    try:
                        start_date_str = record.get('Start Date', '')
                        end_date_str = record.get('End Date', '')
                        apply_date_str = record.get('Apply Date', '')
                        
                        if start_date_str:
                            start_date = pd.to_datetime(start_date_str).date()
                        elif apply_date_str:
                            start_date = pd.to_datetime(apply_date_str).date()
                        else:
                            continue
                            
                        end_date = pd.to_datetime(end_date_str).date() if end_date_str else start_date
                        
                        if start_date <= check_date <= end_date:
                            return True
                    except:
                        continue
            return False
        
        def get_shift_info(emp_name, check_date, sign_in_out_data):
            if sign_in_out_data is None or sign_in_out_data.empty:
                return 'NONE'
            
            target_name = extract_name_from_emp_name(emp_name)
            day_records = sign_in_out_data[
                (sign_in_out_data['emp_name'].astype(str).str.strip() == target_name) &
                (pd.to_datetime(sign_in_out_data['attendance_time'], errors='coerce').dt.date == check_date)
            ]
            
            if day_records.empty:
                return 'NONE'
            
            times = pd.to_datetime(day_records['attendance_time'], errors='coerce')
            times = times.dropna()
            
            if len(times) == 0:
                return 'NONE'
            
            # Check for morning shift (before 12:00)
            morning_times = times[times.dt.hour < 12]
            # Check for afternoon shift (after 12:00)
            afternoon_times = times[times.dt.hour >= 12]
            
            if len(morning_times) > 0 and len(afternoon_times) > 0:
                return 'FULL'
            elif len(morning_times) > 0:
                return 'AM'
            elif len(afternoon_times) > 0:
                return 'PM'
            else:
                return 'NONE'
        
        # Calculate for each employee
        result_rows = []
        row_no = 1
        
        for _, emp in employee_list_df.iterrows():
            emp_name = emp['Name']
            emp_id = emp.get('ID Number', '')
            emp_dept = emp.get('Dept', '')
            
            # Initialize counters
            normal_days = 0.0
            annual_leave = 0.0
            sick_leave = 0.0
            unpaid_leave = 0.0
            welfare_leave = 0.0
            
            # Calculate normal working days
            for dt in date_range:
                day_type = get_day_type(dt, holidays, special_weekends, special_workdays)
                day_date = dt.date()
                
                # Sửa logic: Tính cho ngày làm việc bình thường theo yêu cầu
                # 1. Là ngày trong tuần (Thứ 2-6) VÀ không phải ngày nghỉ lễ
                # 2. Là ngày cuối tuần nhưng được quy định là "ngày làm việc đặc biệt"
                is_normal_workday = (
                    (day_type == 'Weekday') or  # Ngày trong tuần không phải lễ
                    (day_type == 'Weekend' and day_date in special_workdays)  # Cuối tuần nhưng là ngày làm việc đặc biệt
                )
                
                if (is_normal_workday and 
                    not is_lieu_day(emp_name, day_date, ot_lieu_data) and
                    not has_ot_on_date(emp_name, day_date, ot_lieu_data) and
                    not has_apply_leave_on_date(emp_name, day_date, apply_data)):
                    
                    shift_info = get_shift_info(emp_name, day_date, sign_in_out_data)
                    if shift_info == 'FULL':
                        normal_days += 1.0
                    elif shift_info in ['AM', 'PM']:
                        normal_days += 0.5
            
            # Calculate leave days from apply data
            if apply_data is not None and not apply_data.empty:
                target_name = extract_name_from_emp_name(emp_name)
                emp_apply_data = apply_data[
                    (apply_data['Emp Name'].astype(str).str.strip() == target_name) &
                    (apply_data['Results'] == 'Approved') &
                    (apply_data['Type'] == 'Leave')
                ]
                
                for _, apply_record in emp_apply_data.iterrows():
                    try:
                        start_date_str = apply_record.get('Start Date', '')
                        end_date_str = apply_record.get('End Date', '')
                        apply_date_str = apply_record.get('Apply Date', '')
                        leave_type = apply_record.get('Leave Type', '').lower()
                        note = apply_record.get('Note', '').lower()
                        
                        if start_date_str:
                            start_date = pd.to_datetime(start_date_str).date()
                        elif apply_date_str:
                            start_date = pd.to_datetime(apply_date_str).date()
                        else:
                            continue
                            
                        end_date = pd.to_datetime(end_date_str).date() if end_date_str else start_date
                        
                        # Calculate leave days within date range
                        for dt in date_range:
                            day_date = dt.date()
                            if start_date <= day_date <= end_date:
                                day_type = get_day_type(dt, holidays, special_weekends, special_workdays)
                                
                                if day_type == 'Weekday':
                                    # Determine if it's half day or full day
                                    leave_days = 1.0
                                    if 'morning' in note or 'afternoon' in note:
                                        leave_days = 0.5
                                    elif (end_date - start_date).days > 0:
                                        # Multiple days, check if it's the same day
                                        if start_date == end_date:
                                            leave_days = 0.5
                                    
                                    # Add to appropriate leave type
                                    if 'annual' in leave_type:
                                        annual_leave += leave_days
                                    elif 'sick' in leave_type:
                                        sick_leave += leave_days
                                    elif 'unpaid' in leave_type:
                                        unpaid_leave += leave_days
                                    elif 'welfare' in leave_type:
                                        welfare_leave += leave_days
                    except:
                        continue
            
            # Calculate total
            total = normal_days + annual_leave + sick_leave + unpaid_leave + welfare_leave
            
            row = {
                'No': row_no,
                '14 Digits Employee ID': emp_id,
                "Employee's name": emp_name,
                'Group': emp_dept,
                'Normal working days': round(normal_days, 1),
                'Annual leave (100% salary)': round(annual_leave, 1),
                'Sick leave (50% salary)': round(sick_leave, 1),
                'Unpaid leave (0% salary)': round(unpaid_leave, 1),
                'Welfare leave (100% salary)': round(welfare_leave, 1),
                'Total': round(total, 1),
                'Late/Leave early (mins)': 0,  # Placeholder
                'Late/Leave early (times)': 0,  # Placeholder
                'Forget scanning': 0,  # Placeholder
                'Violation': 0,  # Placeholder
                'Remark': '',
                'Attendance for salary payment': round(total, 1)
            }
            
            result_rows.append(row)
            row_no += 1
        
        if not result_rows:
            return {'columns': [], 'rows': []}
        
        result_df = pd.DataFrame(result_rows)
        cols = list(result_df.columns)
        rows = result_df.fillna('').astype(str).values.tolist()
        
        return {'columns': cols, 'rows': rows}
        
    except Exception as e:
        print(f"Error in calculate_total_attendance_detail_for_export: {e}")
        import traceback
        traceback.print_exc()
        return {'columns': [], 'rows': []}

def calculate_attendance_report_for_export():
    """Calculate Attendance Report data for export (without request context)"""
    global sign_in_out_data, apply_data, ot_lieu_data, employee_list_df, rules
    
    # Load data from temp files if global variables are empty
    if sign_in_out_data is None or sign_in_out_data.empty:
        temp_signinout_path = os.path.join(app.config['UPLOAD_FOLDER'], 'temp_signinout.xlsx')
        if os.path.exists(temp_signinout_path):
            sign_in_out_data = pd.read_excel(temp_signinout_path)
        else:
            return {'columns': [], 'rows': []}
    
    if apply_data is None or apply_data.empty:
        temp_apply_path = os.path.join(app.config['UPLOAD_FOLDER'], 'temp_apply.xlsx')
        if os.path.exists(temp_apply_path):
            apply_data = pd.read_excel(temp_apply_path)
        else:
            return {'columns': [], 'rows': []}
    
    if ot_lieu_data is None or ot_lieu_data.empty:
        temp_otlieu_path = os.path.join(app.config['UPLOAD_FOLDER'], 'temp_otlieu.xlsx')
        if os.path.exists(temp_otlieu_path):
            ot_lieu_data = pd.read_excel(temp_otlieu_path)
        else:
            return {'columns': [], 'rows': []}
    
    if employee_list_df is None or employee_list_df.empty:
        emp_path = os.path.join(app.config['UPLOAD_FOLDER'], 'employee_list.csv')
        if os.path.exists(emp_path):
            employee_list_df = pd.read_csv(emp_path, dtype=str)
        else:
            return {'columns': [], 'rows': []}
    
    if rules is None or rules.empty:
        rules_path = os.path.join(app.config['UPLOAD_FOLDER'], 'rules.xlsx')
        if os.path.exists(rules_path):
            rules = pd.read_excel(rules_path)
        else:
            return {'columns': [], 'rows': []}
    
    try:
        # Tự động xác định tháng/năm từ dữ liệu signin/out
        month = 7  # default
        year = 2024  # default
        
        if not sign_in_out_data.empty:
            # Tìm ngày đầu tiên và cuối cùng trong dữ liệu signin/out
            dates = []
            for _, record in sign_in_out_data.iterrows():
                if 'attendance_time' in record and pd.notna(record['attendance_time']):
                    try:
                        date_val = pd.to_datetime(record['attendance_time'])
                        dates.append(date_val)
                    except:
                        continue
            
            if dates:
                min_date = min(dates)
                max_date = max(dates)
                
                # Tìm tháng phổ biến nhất trong dữ liệu
                month_counts = {}
                for date in dates:
                    month_key = (date.month, date.year)
                    month_counts[month_key] = month_counts.get(month_key, 0) + 1
                
                # Lấy tháng/năm có nhiều dữ liệu nhất
                most_common_month = max(month_counts.items(), key=lambda x: x[1])[0]
                month, year = most_common_month
                
                print(f"Auto-detected month/year from signin/out data (export): {month}/{year}")
                print(f"Date range in data: {min_date.date()} to {max_date.date()}")
                print(f"Month distribution: {month_counts}")
        
        # Calculate date range: 20th of previous month to 19th of current month
        if month == 1:
            prev_month = 12
            prev_year = year - 1
        else:
            prev_month = month - 1
            prev_year = year
        
        start_date = pd.Timestamp(prev_year, prev_month, 20)
        end_date = pd.Timestamp(year, month, 19)
        days = pd.date_range(start=start_date, end=end_date, freq='D')
        day_cols = [d.strftime('%Y-%m-%d') for d in days]
        
        # Get special days from rules
        holidays, special_weekends, special_workdays = get_special_days_from_rules(rules)
        
        # Helper functions
        def get_day_type(dt, holidays, special_weekends, special_workdays):
            dt_date = dt.date()
            if dt_date in holidays:
                return 'Holiday'
            if dt_date in special_workdays:
                return 'Weekday'
            if dt_date in special_weekends:
                return 'Weekend'
            if dt.weekday() >= 5:  # Saturday = 5, Sunday = 6
                return 'Weekend'
            return 'Weekday'
        
        def extract_name_from_emp_name(emp_name):
            # Tách tên nhân viên từ format "Do Thi Thu Trang6970000006"
            import re
            match = re.match(r'^(.+?)(\d{7,10})$', emp_name.strip())
            if match:
                return match.group(1).strip()
            return emp_name.strip()
        
        def is_lieu_day(emp_name, check_date, ot_lieu_data):
            if ot_lieu_data is None or ot_lieu_data.empty:
                return False
            target_name = extract_name_from_emp_name(emp_name)
            for _, record in ot_lieu_data.iterrows():
                name_val = str(record.get('Name', '') or '')
                if name_val.strip() == target_name:
                    lieu_cols = ['Lieu From', 'Lieu To', 'Lieu From 2', 'Lieu To 2']
                    for col in lieu_cols:
                        if col in record and pd.notna(record[col]) and str(record[col]).strip():
                            try:
                                lieu_date = pd.to_datetime(record[col]).date()
                                if lieu_date == check_date:
                                    return True
                            except:
                                continue
            return False

        def has_ot_on_date(emp_name, check_date, ot_lieu_data):
            if ot_lieu_data is None or ot_lieu_data.empty:
                return False
            target_name = extract_name_from_emp_name(emp_name)
            for _, record in ot_lieu_data.iterrows():
                name_val = str(record.get('Name', '') or '')
                if name_val.strip() == target_name:
                    ot_date_cols = ['OT Date', 'Date', 'OT date']
                    for col in ot_date_cols:
                        if col in record and pd.notna(record[col]) and str(record[col]).strip():
                            try:
                                ot_date = pd.to_datetime(record[col]).date()
                                if ot_date == check_date:
                                    return True
                            except:
                                continue
            return False
        
        def calculate_actual_hours(check_in_time, check_out_time):
            if not check_in_time or not check_out_time:
                return 0, 0
            
            if isinstance(check_in_time, str):
                try:
                    check_in_time = pd.to_datetime(check_in_time)
                except:
                    return 0, 0
            if isinstance(check_out_time, str):
                try:
                    check_out_time = pd.to_datetime(check_out_time)
                except:
                    return 0, 0
            
            if check_in_time.hour <= 12 and check_out_time.hour >= 13:
                morning_hours = min(12 - check_in_time.hour - check_in_time.minute/60, 4)
                afternoon_hours = min(check_out_time.hour + check_out_time.minute/60 - 13.5, 4.5)
                total_hours = morning_hours + afternoon_hours
            else:
                total_hours = (check_out_time - check_in_time).total_seconds() / 3600
            
            if total_hours < 8:
                late_minutes = (8 - total_hours) * 60
            else:
                late_minutes = 0
                
            return total_hours, late_minutes
        
        def get_apply_info_for_date(emp_name, check_date, apply_data):
            if apply_data is None or apply_data.empty:
                return None
            target_name = extract_name_from_emp_name(emp_name)
            for _, record in apply_data.iterrows():
                if (record.get('Name', '').strip() == target_name and 
                    record.get('Results') == 'Approved'):
                    
                    try:
                        start_date_str = record.get('Start Date', '')
                        end_date_str = record.get('End Date', '')
                        apply_date_str = record.get('Apply Date', '')
                        
                        if start_date_str:
                            start_date = pd.to_datetime(start_date_str).date()
                        elif apply_date_str:
                            start_date = pd.to_datetime(apply_date_str).date()
                        else:
                            continue
                            
                        end_date = pd.to_datetime(end_date_str).date() if end_date_str else start_date
                        
                        if start_date <= check_date <= end_date:
                            return {
                                'type': record.get('Type', ''),
                                'leave_type': record.get('Leave Type', ''),
                                'is_approved': record.get('Results') == 'Approved'
                            }
                    except:
                        continue
            return None
        
        # Create result rows
        result_rows = []
        row_no = 1
        
        for _, emp in employee_list_df.iterrows():
            emp_name = emp['Name']
            emp_dept = emp.get('Dept', '')
            
            # Create 2 rows for each employee: Morning shift & Afternoon shift
            for shift in ['Morning shift', 'Afternoon shift']:
                row = {
                    'Department': emp_dept,
                    'Name': emp_name,
                    'Shift': shift
                }
                
                # Add day columns
                for day in days:
                    day_str = day.strftime('%Y-%m-%d')
                    day_type = get_day_type(day, holidays, special_weekends, special_workdays)
                    day_date = day.date()
                    
                    # Process weekday (Monday to Friday, no Lieu)
                    if day_type == 'Weekday' and not is_lieu_day(emp_name, day_date, otlieu_data):
                        
                        # Get signinout data for this day
                        day_signinout = []
                        if sign_in_out_data is not None and not sign_in_out_data.empty:
                            df_sign = sign_in_out_data.copy()
                            if 'emp_name' in df_sign.columns and 'attendance_time' in df_sign.columns:
                                df_sign['attendance_time'] = pd.to_datetime(df_sign['attendance_time'], errors='coerce')
                                df_sign['date'] = df_sign['attendance_time'].dt.date
                                target_name = extract_name_from_emp_name(emp_name)
                                mask = (df_sign['emp_name'].astype(str).str.strip() == target_name) & (df_sign['date'] == day_date)
                                day_signinout = df_sign[mask]['attendance_time'].tolist()
                        
                        # Get apply info for this day
                        apply_info = get_apply_info_for_date(emp_name, day_date, apply_data)
                        
                        # Process based on apply type
                        if apply_info:
                            apply_type = apply_info['type']
                            leave_type = apply_info['leave_type'].lower()
                            
                            if apply_type == 'Leave':
                                row[day_str] = 'L'  # Leave
                            elif apply_type == 'Supplement':
                                if day_signinout:
                                    # Xử lý dữ liệu trùng lặp - lấy thời gian sớm nhất và muộn nhất
                                    morning_times = [t for t in day_signinout if t.hour < 12]
                                    afternoon_times = [t for t in day_signinout if t.hour >= 12]
                                    
                                    if morning_times and afternoon_times:
                                        check_in = min(morning_times)  # Thời gian sớm nhất buổi sáng
                                        check_out = max(afternoon_times)  # Thời gian muộn nhất buổi chiều
                                    else:
                                        check_in = min(day_signinout)
                                        check_out = max(day_signinout)
                                    
                                    actual_hours, late_minutes = calculate_actual_hours(check_in, check_out)
                                    
                                    if actual_hours >= 8:
                                        row[day_str] = 'N'  # Normal
                                    else:
                                        row[day_str] = 'LS'  # Late/Soon
                                else:
                                    row[day_str] = 'S'  # Supplement
                            elif apply_type == 'Trip':
                                row[day_str] = 'T'  # Trip
                        else:
                            # No apply, calculate normal working
                            if day_signinout:
                                check_in = min(day_signinout)
                                check_out = max(day_signinout)
                                actual_hours, late_minutes = calculate_actual_hours(check_in, check_out)
                                
                                if actual_hours >= 8:
                                    row[day_str] = 'N'  # Normal
                                else:
                                    row[day_str] = 'LS'  # Late/Soon
                                    
                                    if actual_hours < 4:
                                        row[day_str] = 'M'  # Miss
                            else:
                                row[day_str] = 'M'  # Miss
                    else:
                        # Not a weekday or has Lieu
                        if day_type == 'Holiday':
                            row[day_str] = 'H'  # Holiday
                        elif day_type == 'Weekend':
                            row[day_str] = 'W'  # Weekend
                        elif is_lieu_day(emp_name, day_date, ot_lieu_data):
                            row[day_str] = 'LE'  # Lieu
                        elif has_ot_on_date(emp_name, day_date, ot_lieu_data):
                            row[day_str] = 'OT'  # OT
                        else:
                            row[day_str] = ''  # Empty
                
                # Calculate summary
                summary = {'Normal': 0, 'Leave': 0, 'Trip': 0, 'Miss': 0, 'Late/Soon': 0, 'Lieu': 0, 'OT': 0, 'Supplement': 0}
                for day in days:
                    day_str = day.strftime('%Y-%m-%d')
                    val = row.get(day_str, '')
                    if val == 'N': summary['Normal'] += 1
                    elif val == 'L': summary['Leave'] += 1
                    elif val == 'T': summary['Trip'] += 1
                    elif val == 'M': summary['Miss'] += 1
                    elif val == 'LS': summary['Late/Soon'] += 1
                    elif val == 'LE': summary['Lieu'] += 1
                    elif val == 'OT': summary['OT'] += 1
                    elif val == 'S': summary['Supplement'] += 1
                
                row.update(summary)
                result_rows.append(row)
                row_no += 1
        
        if not result_rows:
            return {'columns': [], 'rows': []}
        
        # Create DataFrame
        result = pd.DataFrame(result_rows)
        
        # Define column order
        columns = ['Department', 'Name', 'Shift'] + day_cols + [
            'Normal', 'Leave', 'Trip', 'Miss', 'Late/Soon', 'Lieu', 'OT', 'Supplement'
        ]
        
        # Ensure all columns exist
        for col in columns:
            if col not in result.columns:
                result[col] = ''
        
        # Sort by column order
        result = result[columns]
        
        cols = list(result.columns)
        rows = result.fillna('').astype(str).values.tolist()
        
        return {'columns': cols, 'rows': rows}
        
    except Exception as e:
        print(f"Error in calculate_attendance_report_for_export: {e}")
        return {'columns': [], 'rows': []}


# ==========================
# Styling Functions
# ==========================
def apply_employee_list_styling(worksheet):
    """Apply styling to Employee List sheet"""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    
    # Header styling
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center")
    
    # Apply header styling
    for cell in worksheet[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
    
    # Column widths
    worksheet.column_dimensions['A'].width = 5   # STT
    worksheet.column_dimensions['B'].width = 25  # Name
    worksheet.column_dimensions['C'].width = 20  # ID Number
    worksheet.column_dimensions['D'].width = 15  # Dept
    worksheet.column_dimensions['E'].width = 12  # Internship
    worksheet.column_dimensions['F'].width = 8   # Delete button
    
    # Border styling
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    for row in worksheet.iter_rows():
        for cell in row:
            cell.border = thin_border

def apply_rules_styling(worksheet):
    """Apply styling to Rules sheet"""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    
    # Header styling
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center")
    
    # Apply header styling
    for cell in worksheet[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
    
    # Auto-adjust column widths
    for column in worksheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        worksheet.column_dimensions[column_letter].width = adjusted_width
    
    # Border styling
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    for row in worksheet.iter_rows():
        for cell in row:
            cell.border = thin_border

def apply_signinout_styling(worksheet):
    """Apply styling to Sign In-Out Data sheet"""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    
    # Header styling
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center")
    
    # Apply header styling
    for cell in worksheet[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
    
    # Auto-adjust column widths
    for column in worksheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        worksheet.column_dimensions[column_letter].width = adjusted_width
    
    # Border styling
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    for row in worksheet.iter_rows():
        for cell in row:
            cell.border = thin_border

def apply_apply_styling(worksheet):
    """Apply styling to Apply Data sheet"""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    
    # Header styling
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center")
    
    # Apply header styling
    for cell in worksheet[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
    
    # Auto-adjust column widths
    for column in worksheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        worksheet.column_dimensions[column_letter].width = adjusted_width
    
    # Border styling
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    for row in worksheet.iter_rows():
        for cell in row:
            cell.border = thin_border

def apply_otlieu_styling(worksheet):
    """Apply styling to OT Lieu Data sheet"""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    
    # Header styling
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="C5504B", end_color="C5504B", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center")
    
    # Apply header styling
    for cell in worksheet[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
    
    # Auto-adjust column widths
    for column in worksheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        worksheet.column_dimensions[column_letter].width = adjusted_width
    
    # Border styling
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Apply styling to all cells including error/warning colors
    for row_idx, row in enumerate(worksheet.iter_rows(), 1):
        for col_idx, cell in enumerate(row, 1):
            cell.border = thin_border
            
            # Apply error/warning styling based on cell content
            if row_idx > 1:  # Skip header row
                cell_value = str(cell.value) if cell.value else ""
                
                # Error cells (red background)
                if "error" in cell_value.lower() or "invalid" in cell_value.lower():
                    cell.fill = PatternFill(start_color="FFCDD2", end_color="FFCDD2", fill_type="solid")
                    cell.font = Font(color="D32F2F", bold=True)
                
                # Warning cells (yellow background)
                elif "warning" in cell_value.lower() or "suggest" in cell_value.lower():
                    cell.fill = PatternFill(start_color="FFF9C4", end_color="FFF9C4", fill_type="solid")
                    cell.font = Font(color="F57F17", bold=True)
                
                # Gray cells (inactive)
                elif "gray" in cell_value.lower() or "inactive" in cell_value.lower():
                    cell.fill = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")
                    cell.font = Font(color="757575")

def apply_otlieu_before_styling(worksheet):
    """Apply styling to OT Lieu Before sheet"""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    
    # Header styling
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="9C5700", end_color="9C5700", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center")
    
    # Apply header styling
    for cell in worksheet[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
    
    # Auto-adjust column widths
    for column in worksheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        worksheet.column_dimensions[column_letter].width = adjusted_width
    
    # Border styling
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    for row in worksheet.iter_rows():
        for cell in row:
            cell.border = thin_border

def apply_otlieu_report_styling(worksheet):
    """Apply styling to OT Lieu Report sheet"""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    
    # Header styling
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="7030A0", end_color="7030A0", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center")
    
    # Apply header styling
    for cell in worksheet[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
    
    # Auto-adjust column widths
    for column in worksheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        worksheet.column_dimensions[column_letter].width = adjusted_width
    
    # Border styling
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Apply styling to all cells including highlight for "Total OT paid" > 25
    for row_idx, row in enumerate(worksheet.iter_rows(), 1):
        for col_idx, cell in enumerate(row, 1):
            cell.border = thin_border
            
            # Highlight "Total OT paid" column when value > 25
            if row_idx > 1:  # Skip header row
                header_cell = worksheet.cell(row=1, column=col_idx)
                if header_cell.value == "Total OT paid":
                    try:
                        ot_value = float(cell.value) if cell.value and str(cell.value) != '-' else 0
                        if ot_value > 25:
                            cell.fill = PatternFill(start_color="FFEBEE", end_color="FFEBEE", fill_type="solid")
                            cell.font = Font(color="D32F2F", bold=True)
                    except:
                        pass

def apply_total_attendance_styling(worksheet):
    """Apply styling to Total Attendance Detail sheet"""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    
    # Header styling with different colors for column groups
    header_font = Font(bold=True, color="FFFFFF")
    header_alignment = Alignment(horizontal="center", vertical="center")
    
    # Define column group colors
    base_cols = ["No", "14 Digits Employee ID", "Employee's name", "Group"]
    attendance_cols = [
        "Normal working days", "Annual leave (100% salary)", "Sick leave (50% salary)",
        "Unpaid leave (0% salary)", "Welfare leave (100% salary)", "Total"
    ]
    violation_cols = [
        "Late/Leave early (mins)", "Late/Leave early (times)", "Forget scanning", "Violation"
    ]
    remain_cols = ["Remark", "Attendance for salary payment"]
    
    # Apply header styling with different colors
    for col_idx, cell in enumerate(worksheet[1], 1):
        cell.font = header_font
        cell.alignment = header_alignment
        
        # Determine column group and apply color
        if cell.value in base_cols:
            cell.fill = PatternFill(start_color="31869B", end_color="31869B", fill_type="solid")
        elif cell.value in attendance_cols:
            cell.fill = PatternFill(start_color="E0F7FA", end_color="E0F7FA", fill_type="solid")
            cell.font = Font(bold=True, color="000000")  # Black text for light background
        elif cell.value in violation_cols:
            cell.fill = PatternFill(start_color="FFEAEA", end_color="FFEAEA", fill_type="solid")
            cell.font = Font(bold=True, color="000000")  # Black text for light background
        elif cell.value in remain_cols:
            cell.fill = PatternFill(start_color="31869B", end_color="31869B", fill_type="solid")
        else:
            cell.fill = PatternFill(start_color="31869B", end_color="31869B", fill_type="solid")
    
    # Auto-adjust column widths
    for column in worksheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        worksheet.column_dimensions[column_letter].width = adjusted_width
    
    # Border styling
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    for row in worksheet.iter_rows():
        for cell in row:
            cell.border = thin_border

def apply_attendance_report_styling(worksheet):
    """
    Áp dụng màu sắc và định dạng cho sheet 'Attendance Report'.
    Hàm này sẽ tự suy luận trạng thái từ giá trị của ô.
    """
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    
    # --- STYLE ---
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="255E91", end_color="255E91", fill_type="solid")
    name_body_alignment = Alignment(horizontal="left", vertical="center")
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    
    # --- STYLE HEADER ---
    for cell in worksheet[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    
    for col_idx, column in enumerate(worksheet.columns, 1):
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
            adjusted_width = min(max_length + 2, 50)
        worksheet.column_dimensions[column_letter].width = adjusted_width
    
    # --- LOGIC ---

    # 'TotalLateMinutes'
    late_minutes_col_idx = None
    summary_start_col_idx = None
    for col_idx, cell in enumerate(worksheet[1], 1):
        if cell.value == 'TotalLateMinutes':
            late_minutes_col_idx = col_idx
        if cell.value == 'Normal': 
             summary_start_col_idx = col_idx

    for row_idx, row in enumerate(worksheet.iter_rows(min_row=2), 2):
        has_late_soon_violation = False
        if late_minutes_col_idx:
            try:
                minutes = float(worksheet.cell(row=row_idx, column=late_minutes_col_idx).value)
                if minutes > 0:
                    has_late_soon_violation = True
            except (ValueError, TypeError):
                pass

        for cell in row:
            cell.border = thin_border
            
            if summary_start_col_idx and cell.column >= summary_start_col_idx:
                continue

            cell_value = str(cell.value) if cell.value is not None else ""

            if cell_value == 'Trip':
                cell.fill = PatternFill(start_color="BBDEFB", fill_type="solid")
            elif cell_value == 'Leave':
                cell.fill = PatternFill(start_color="FFE0B2", fill_type="solid")
            elif cell_value == 'Supplement':
                cell.fill = PatternFill(start_color="C5CAE9", fill_type="solid")
            elif cell_value == 'Lieu':
                cell.fill = PatternFill(start_color="E1BEE7", fill_type="solid")
            elif cell_value == '0' or cell_value.lower() == 'miss':
                cell.fill = PatternFill(start_color="FFCDD2", fill_type="solid")  # Nền đỏ
                cell.font = Font(color="D32F2F", bold=True)
            elif ':' in cell_value:
                if has_late_soon_violation:
                    cell.fill = PatternFill(start_color="FFF9C4", fill_type="solid")  # Nền vàng
                    cell.font = Font(color="D32F2F", bold=True)
                else:
                    cell.fill = PatternFill(start_color="C8E6C9", fill_type="solid")  # Nền xanh lá
            elif cell_value == '': 
                 cell.fill = PatternFill(start_color="E0E0E0", fill_type="solid")

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000) 