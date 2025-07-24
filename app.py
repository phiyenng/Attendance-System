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
EMPLOYEE_LIST_PATH = os.path.join(app.config['UPLOAD_FOLDER'], 'employee_list.csv')
# Load employee list from file if exists
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
RULES_XLSX_PATH = os.path.join(app.config['UPLOAD_FOLDER'], 'rules.xlsx')
if os.path.exists(RULES_XLSX_PATH):
    try:
        rules = pd.read_excel(RULES_XLSX_PATH)
    except Exception as e:
        print(f"Error loading rules: {e}")
        rules = pd.DataFrame()
else:
    # Create default rules file
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

    # Add "Name" columns
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

    # Xác định tất cả các cột ngày/giờ/tổng OT/Lieu (chứa cả 'ot' và 'from', v.v.)
    ot_from_cols = [c for c in df.columns if re.search(r'ot.*from', c, re.I)]
    ot_to_cols = [c for c in df.columns if re.search(r'ot.*to', c, re.I)]
    sum_ot_col = next((c for c in df.columns if re.search(r'ot.*sum', c, re.I)), None)
    lieu_from_cols = [c for c in df.columns if re.search(r'lieu.*from', c, re.I)]
    lieu_to_cols = [c for c in df.columns if re.search(r'lieu.*to', c, re.I)]
    sum_lieu_col = next((c for c in df.columns if re.search(r'lieu.*sum', c, re.I)), None)
    
    # Chuẩn hóa số giờ OT/Lieu
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

    # Mark warnings for invalid time format for all relevant columns
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
        # # Nếu là 'Hour : Minutes AM', 'Hour : Minutes PM', hoặc 'Hour : Minutes AM/PM' (bất kể hoa thường, khoảng trắng)
        # #s_no_space = re.sub(r'\s+', '', s).lower()
        # if s in ['Hour : Minutes AM', 'Hour : Minutes PM', 'Hour h Min']:
        #     return ''
        
        # 1. Dạng 12h AM/PM
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
        # 2. Dạng 24h: 21:00, 21.00, 21 00, 21h00
        m = re.match(r'^([01]?[0-9]|2[0-3])\s*[:. h]\s*([0-5][0-9])$', s)
        if m:
            hour = int(m.group(1))
            minute = int(m.group(2))
            return f'{hour:02d}:{minute:02d}'
        return None

# ---- OT From/To & Lieu From/To change format HH:MM ----
    def norm_time_to_24h(val):
        # Bỏ qua các giá trị đặc biệt
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

    # Helper: check if string is in 12h AM/PM format
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
    """Load data from Excel file"""
    try:
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
            df = pd.read_excel(file_path, engine='xlrd')
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
            df = pd.read_excel(file_path, engine='xlrd')
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

        # Load data
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
    
    # Save temporary reviewed/edited data for report
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
    
    # Clear previous abnormal data
    abnormal_data = pd.DataFrame(columns=['Employee', 'Date', 'SignIn', 'SignOut', 'Status', 'LateMinutes'])
    
    # Calculate attendance
    # This function is no longer needed as check_apply and check_lieu are removed.
    # The logic for calculating abnormal attendance needs to be re-evaluated based on the new data structures.
    # For now, we'll just return a placeholder message.
    return jsonify({'success': True, 'message': 'Abnormal calculation completed (placeholder).'})

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
                excel = pd.ExcelFile(BytesIO(file_bytes), engine='xlrd')
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
                df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name, engine='xlrd')
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
    # Placeholder: just return success
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
    # Không ép kiểu str, chỉ fillna('') để giữ dict lỗi
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
            # Luôn cập nhật giá trị mới vào DataFrame, không trả về object warning/error nữa
            ot_lieu_data.at[idx, col] = value
            updated = {}
            # Nếu là cột OT/Lieu Sum, có thể tính lại giá trị thực tế và trả về nếu cần
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

    for idx, row in df.iterrows():
        emp_id = row['Emp ID'] if 'Emp ID' in row else None
        intern = is_intern(emp_id, employee_list_df)

        ot_date = None
        for col in df.columns:
            if 'ot' in col.lower() and 'day' in col.lower() and pd.notna(row[col]):
                try:
                    ot_date = pd.to_datetime(row[col]).date()
                    break
                except:
                    pass
        if ot_date is None:
            for col in df.columns:
                if 'lieu' in col.lower() and 'date' in col.lower() and pd.notna(row[col]):
                    try:
                        ot_date = pd.to_datetime(row[col]).date()
                        break
                    except:
                        pass
        if ot_date is None:
            continue

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

    # Làm tròn 3 số
    for col in ['OT Payment: ' + c for c in ot_payment_types] + ['Change in lieu: ' + c for c in change_in_lieu_types]:
        if col in df.columns:
            df[col] = df[col].apply(lambda x: round(float(x), 3) if str(x).replace('.', '', 1).replace('-', '', 1).isdigit() else x)

    # Tính Lieu used
    import numpy as np
    for idx, row in df.iterrows():
        name = row['Name'] if 'Name' in row else None
        lieu_needed = 0.0
        lieu_from_col = next((c for c in df.columns if 'lieu' in c.lower() and 'from' in c.lower()), None)
        lieu_to_col = next((c for c in df.columns if 'lieu' in c.lower() and 'to' in c.lower()), None)
        if lieu_from_col and lieu_to_col and name:
            try:
                t1 = pd.to_datetime(str(row[lieu_from_col]), format='%H:%M')
                t2 = pd.to_datetime(str(row[lieu_to_col]), format='%H:%M')
                if t2 < t1: t2 += pd.Timedelta(days=1)
                hours = (t2 - t1).total_seconds() / 3600
                if t1.hour <= 12 < t2.hour or (t1.hour == 12 and t2.hour >= 13):
                    if t2.hour > 13 or (t2.hour == 13 and t2.minute >= 30):
                        hours -= 1.5
                lieu_needed = round(hours, 2)
            except:
                pass

        used = 0.0
        if lieu_needed > 0 and name in lieu_remain_map:
            remain = lieu_remain_map[name]
            if remain >= lieu_needed:
                lieu_remain_map[name] -= lieu_needed
                used = lieu_needed
            else:
                used = remain
                lieu_remain_map[name] = 0.0
        df.at[idx, 'Lieu used'] = round(used, 2) if used else np.nan

    # Trừ từ OT Payment theo hệ số ưu tiên
    priority_order = [
        ('OT Payment: Weekday Rate 150%', 1.5),
        ('OT Payment: Weekday-night Rate 200%', 2.0),
        ('OT Payment: Weekend Rate 200%', 2.0),
        ('OT Payment: Weekend-night Rate 270%', 2.7),
        ('OT Payment: Holiday Rate 300%', 3.0),
        ('OT Payment: Holiday-night Rate 390%', 3.9),
    ]
    for idx, row in df.iterrows():
        used = row.get('Lieu used', 0)
        if not pd.isna(used) and used > 0:
            remaining_lieu = used
            for col, ratio in priority_order:
                available = df.at[idx, col] if col in df.columns else 0.0
                try:
                    available = float(available)
                except:
                    continue
                deduct = round(remaining_lieu / ratio, 3)
                if available >= deduct:
                    df.at[idx, col] = round(available - deduct, 3)
                    break
                else:
                    remaining_lieu = round(remaining_lieu - available * ratio, 3)
                    df.at[idx, col] = 0.0

    # Ensure 'Lieu used' and 'Remark' columns exist before reordering
    if 'Lieu used' not in df.columns:
        df['Lieu used'] = np.nan
    if 'Remark' not in df.columns:
        df['Remark'] = np.nan
    cols = [c for c in df.columns if c not in ['Lieu used', 'Remark']] + ['Lieu used', 'Remark']
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
        used_cols = [c for c in ot_cols if c.startswith('Change in lieu')]
        paid_cols = [c for c in ot_cols if c.startswith('OT Payment')]
        ot_sum['Total used hours in month'] = ot_sum[used_cols].sum(axis=1)
        ot_sum['Total OT paid'] = ot_sum[paid_cols].sum(axis=1)
        def calc_transfer(row):
            try:
                val = float(row['Total OT paid'])
                return round(val - 25, 2) if val > 25 else ''
            except:
                return ''
        ot_sum['Transferred to normal working hours'] = ot_sum.apply(calc_transfer, axis=1)
        ot_sum['Date'] = ''
        ot_sum['Remain unused time off in lieu'] = ''

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

    # Thêm các cột Attendance & Violation mặc định là rỗng (hoặc 0)
    for col in [
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
    ]:
        result[col] = ''

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

import random
@app.route('/get_attendance_report')
def get_attendance_report():
    # Lấy tháng/năm từ query string
    month = int(request.args.get('month', 7))
    year = int(request.args.get('year', 2024))

    # Lấy danh sách nhân viên từ file employee_list.csv
    emp_path = EMPLOYEE_LIST_PATH
    if os.path.exists(emp_path):
        emp_df = pd.read_csv(emp_path, dtype=str)
    else:
        emp_df = pd.DataFrame(columns=["Dept", "Name"])

    # Sắp xếp theo Dept nếu có
    if 'Dept' in emp_df.columns:
        emp_df = emp_df.sort_values(by=['Dept', 'Name'], ascending=[True, True])
        emp_df = emp_df.reset_index(drop=True)

    # Tạo dãy ngày từ 20 tháng trước đến 19 tháng này
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

    # Tạo columns
    columns = ['Department', 'Name'] + day_cols + [
        'Normal', 'Leave', 'Trip', 'Miss', 'Late/Soon', 'Lieu', 'OT', 'Supplement'
    ]

    # Tạo dữ liệu mẫu cho từng nhân viên
    rows = []
    for _, emp in emp_df.iterrows():
        row = [emp.get('Dept', ''), emp.get('Name', '')]
        # Mỗi ngày là 2 ca: random trạng thái
        day_statuses = []
        for d in days:
            # Ví dụ: random trạng thái cho demo, bạn thay bằng tính toán thực tế
            ca_sang = random.choice(['N', 'L', 'T', 'M', 'LS', 'LE', 'OT', 'S', ''])
            ca_chieu = random.choice(['N', 'L', 'T', 'M', 'LS', 'LE', 'OT', 'S', ''])
            day_statuses.append([ca_sang, ca_chieu])
        row.extend(day_statuses)
        # Tổng hợp: đếm số lần xuất hiện từng loại
        summary = {'Normal': 0, 'Leave': 0, 'Trip': 0, 'Miss': 0, 'Late/Soon': 0, 'Lieu': 0, 'OT': 0, 'Supplement': 0}
        for ca in day_statuses:
            for val in ca:
                if val == 'N': summary['Normal'] += 1
                if val == 'L': summary['Leave'] += 1
                if val == 'T': summary['Trip'] += 1
                if val == 'M': summary['Miss'] += 1
                if val == 'LS': summary['Late/Soon'] += 1
                if val == 'LE': summary['Lieu'] += 1
                if val == 'OT': summary['OT'] += 1
                if val == 'S': summary['Supplement'] += 1
        row.extend([summary['Normal'], summary['Leave'], summary['Trip'], summary['Miss'],
                    summary['Late/Soon'], summary['Lieu'], summary['OT'], summary['Supplement']])
        rows.append(row)

    return jsonify({'columns': columns, 'rows': rows})

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

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000) 