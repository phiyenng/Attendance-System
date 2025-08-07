#!/usr/bin/env python3
"""
Test script for flexible timing logic
Testing late arrival compensation rule: sign in 8:00-8:30 with extended work time
"""

import pandas as pd
from datetime import datetime, timedelta

def test_flexible_timing():
    """Test the new flexible timing logic"""
    
    # Test cases based on user requirements
    test_cases = [
        {
            'name': 'Normal case - 8:00 to 17:30',
            'sign_in': '08:00',
            'sign_out': '17:30',
            'expected': 'Normal (8 hours exactly)'
        },
        {
            'name': 'Late arrival case - 8:24 to 18:09', 
            'sign_in': '08:24',
            'sign_out': '18:09',
            'expected': 'Normal (24 min late, 39 min extended = 15 min extra)'
        },
        {
            'name': 'Late arrival insufficient compensation - 8:24 to 17:50',
            'sign_in': '08:24', 
            'sign_out': '17:50',
            'expected': 'Late/Soon (24 min late, only 20 min extended = 4 min short)'
        },
        {
            'name': 'Late arrival exact compensation - 8:15 to 17:45',
            'sign_in': '08:15',
            'sign_out': '17:45', 
            'expected': 'Normal (15 min late, exactly 15 min extended)'
        },
        {
            'name': 'Too late arrival - 8:35 to 18:05',
            'sign_in': '08:35',
            'sign_out': '18:05',
            'expected': 'Standard logic (outside 8:00-8:30 window)'
        }
    ]
    
    print("=== Testing Flexible Timing Logic ===\n")
    
    for test_case in test_cases:
        print(f"Test: {test_case['name']}")
        print(f"Sign In: {test_case['sign_in']}, Sign Out: {test_case['sign_out']}")
        
        # Parse times
        test_date = datetime(2024, 7, 15)  # Monday
        check_in_time = datetime.strptime(f"2024-07-15 {test_case['sign_in']}", "%Y-%m-%d %H:%M")
        check_out_time = datetime.strptime(f"2024-07-15 {test_case['sign_out']}", "%Y-%m-%d %H:%M")
        
        # Apply logic
        work_seconds = (check_out_time - check_in_time).total_seconds()
        
        # Check if spans lunch
        lunch_start = check_in_time.replace(hour=12, minute=0, second=0)
        lunch_end = check_in_time.replace(hour=13, minute=30, second=0)
        spans_lunch = check_in_time < lunch_end and check_out_time > lunch_start
        
        if spans_lunch:
            work_seconds -= 1.5 * 3600  # Subtract lunch break
            required_hours = 8.0
        else:
            required_hours = 8.0
        
        work_hours = work_seconds / 3600
        
        # Time range validation with flexible compensation
        standard_start = check_in_time.replace(hour=8, minute=0, second=0)
        late_arrival_cutoff = check_in_time.replace(hour=8, minute=30, second=0)
        standard_end_time = check_in_time.replace(hour=18, minute=0, second=0)
        
        # Check for late arrival within acceptable range (8:00-8:30)
        late_arrival_minutes = 0
        max_allowed_checkout = standard_end_time
        
        if check_in_time > standard_start and check_in_time <= late_arrival_cutoff:
            late_arrival_minutes = int((check_in_time - standard_start).total_seconds() / 60)
            # Allow extended checkout time for late arrival compensation
            max_allowed_checkout = standard_end_time + pd.Timedelta(minutes=late_arrival_minutes)
        
        time_range_valid = (
            check_in_time.hour >= 8 and 
            check_out_time <= max_allowed_checkout
        )
        
        if not time_range_valid:
            status = 'Late/Soon'
            late_minutes = 30
            if check_out_time > max_allowed_checkout:
                if late_arrival_minutes > 0:
                    reason = f"Late departure beyond compensation limit (max allowed: {max_allowed_checkout.strftime('%H:%M')})"
                else:
                    reason = "Late departure after 6:00 PM"
            else:
                reason = "Time outside valid range"
        else:
            # NEW FLEXIBLE LOGIC
            standard_work_start = check_in_time.replace(hour=8, minute=0, second=0)
            standard_work_end = check_in_time.replace(hour=17, minute=30, second=0)
            
            if check_in_time > standard_work_start and check_in_time <= late_arrival_cutoff:
                required_end_time = standard_work_end + pd.Timedelta(minutes=late_arrival_minutes)
                
                if check_out_time >= required_end_time:
                    status = 'Normal'
                    late_minutes = 0
                    reason = f"Late arrival compensated ({late_arrival_minutes} min late, sufficient extension)"
                else:
                    shortage_minutes = int((required_end_time - check_out_time).total_seconds() / 60)
                    status = 'Late/Soon' 
                    late_minutes = shortage_minutes
                    reason = f"Late arrival not fully compensated ({late_arrival_minutes} min late, {shortage_minutes} min short)"
            else:
                # Standard check
                if work_hours < required_hours:
                    shortage_hours = required_hours - work_hours
                    late_minutes = int(shortage_hours * 60)
                    status = 'Late/Soon'
                    reason = f"Insufficient work hours ({work_hours:.2f}/{required_hours})"
                else:
                    status = 'Normal'
                    late_minutes = 0
                    reason = f"Sufficient work hours ({work_hours:.2f})"
        
        print(f"Result: {status} ({late_minutes} penalty minutes)")
        print(f"Reason: {reason}")
        print(f"Expected: {test_case['expected']}")
        print(f"Work Hours: {work_hours:.2f} (spans lunch: {spans_lunch})")
        
        if late_arrival_minutes > 0:
            required_end = standard_work_end + pd.Timedelta(minutes=late_arrival_minutes)
            print(f"Late arrival: {late_arrival_minutes} min, Required end: {required_end.strftime('%H:%M')}")
            print(f"Max allowed checkout: {max_allowed_checkout.strftime('%H:%M')}")
        
        print("-" * 60)

if __name__ == "__main__":
    test_flexible_timing()
