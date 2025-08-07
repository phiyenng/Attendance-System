#!/usr/bin/env python3
"""
Test script for multiple sign-in/sign-out handling
"""
import pandas as pd
import os
import sys

# Add current directory to path to import app
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def test_multiple_signin_logic():
    """Test the logic for handling multiple sign-in/sign-out records"""
    
    # Create test data similar to user's example
    test_data = [
        {'emp_name': 'BUI BA UY12345678', 'attendance_time': '2025-06-11 17:45:00'},
        {'emp_name': 'BUI BA UY12345678', 'attendance_time': '2025-06-11 08:14:00'},
        {'emp_name': 'BUI BA UY12345678', 'attendance_time': '2025-05-20 17:52:00'},
        {'emp_name': 'BUI BA UY12345678', 'attendance_time': '2025-05-20 08:17:00'},
        {'emp_name': 'BUI BA UY12345678', 'attendance_time': '2025-05-20 08:16:00'},  # Multiple morning check-ins
    ]
    
    df = pd.DataFrame(test_data)
    df['attendance_time'] = pd.to_datetime(df['attendance_time'])
    
    print("=== TEST DATA ===")
    print(df.to_string(index=False))
    print()
    
    # Simulate the logic from process_daily_attendance_with_constraints
    def extract_name_from_emp_name(emp_name):
        """Extract name from emp_name field (remove employee ID)"""
        import re
        if not isinstance(emp_name, str):
            return ""
        match = re.match(r'^(.+?)(\d{7,10})$', emp_name.strip())
        if match:
            return match.group(1).strip()
        return emp_name.strip()
    
    df['date'] = df['attendance_time'].dt.date
    df['NormalizedName'] = df['emp_name'].apply(extract_name_from_emp_name).str.lower()
    
    print("=== PROCESSED DATA ===")
    print(df[['NormalizedName', 'date', 'attendance_time']].to_string(index=False))
    print()
    
    # Group by employee and date
    result_data = []
    for (name, date), group in df.groupby(['NormalizedName', 'date']):
        times = group['attendance_time'].dropna()
        
        # Apply time constraints: 8:00 AM - 6:00 PM working hours
        valid_times = times[(times.dt.hour >= 8) & (times.dt.hour <= 18)]
        
        if valid_times.empty:
            continue
            
        # Morning shift: Find MIN time, exclude if > 4:00 PM
        morning_times = valid_times[valid_times.dt.hour < 16]  # Before 4 PM
        morning_time = morning_times.min() if not morning_times.empty else None
        
        # Afternoon shift: Find MAX time, exclude if < 10:30 AM  
        afternoon_times = valid_times[valid_times.dt.time >= pd.Timestamp('10:30').time()]
        afternoon_time = afternoon_times.max() if not afternoon_times.empty else None
        
        # Overall SignIn/SignOut for work hours calculation
        sign_in = valid_times.min()
        sign_out = valid_times.max()
        
        result_data.append({
            'Employee': name.upper(),
            'Date': date,
            'OriginalRecords': len(group),
            'ValidRecords': len(valid_times),
            'MorningTime': morning_time.strftime('%H:%M') if morning_time else 'None',
            'AfternoonTime': afternoon_time.strftime('%H:%M') if afternoon_time else 'None',
            'SignIn': sign_in.strftime('%H:%M') if sign_in else 'None',
            'SignOut': sign_out.strftime('%H:%M') if sign_out else 'None',
            'AllTimes': [t.strftime('%H:%M') for t in times.sort_values()]
        })
    
    result_df = pd.DataFrame(result_data)
    print("=== FINAL RESULT ===")
    for _, row in result_df.iterrows():
        print(f"Employee: {row['Employee']}")
        print(f"Date: {row['Date']}")
        print(f"Original Records: {row['OriginalRecords']}")
        print(f"Valid Records: {row['ValidRecords']}")
        print(f"All Times: {row['AllTimes']}")
        print(f"Morning Time (MIN, exclude >4PM): {row['MorningTime']}")
        print(f"Afternoon Time (MAX, exclude <10:30AM): {row['AfternoonTime']}")
        print(f"SignIn (earliest): {row['SignIn']}")
        print(f"SignOut (latest): {row['SignOut']}")
        print("-" * 50)

if __name__ == "__main__":
    test_multiple_signin_logic()
