import streamlit as st
import pandas as pd
import numpy as np
from ortools.sat.python import cp_model
from rapidfuzz import process, fuzz
from collections import defaultdict
from openpyxl import load_workbook
from datetime import datetime, timedelta
import re
from dateutil.parser import parse
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="Exam Timetabling System", layout="wide")

st.title("Exam Timetabling System")
st.markdown("""
This app helps create an optimal exam timetable while considering various constraints like:
- Student exam conflicts
- Module leader preferences
- Special accommodations for students with extra time
- Fixed exam dates
- Bank holidays
- Room assignments
- AEA requirements
""")

# File upload section
st.header("Upload Required Files")
col1, col2, col3 = st.columns(3)

with col1:
    student_file = st.file_uploader("Upload Student List", type=['xlsx'])
with col2:
    module_file = st.file_uploader("Upload Module List", type=['xlsx'])
with col3:
    dates_file = st.file_uploader("Upload Useful Dates", type=['xlsx'])

# Parameters section
st.header("Timetabling Parameters")
col1, col2 = st.columns(2)

with col1:
    num_days = st.number_input("Number of Days for Exam Period", min_value=1, max_value=30, value=20)
    max_exams_2days = st.number_input("Maximum Exams in 2-Day Window", min_value=1, max_value=5, value=3)
    max_exams_5days = st.number_input("Maximum Exams in 5-Day Window", min_value=1, max_value=10, value=4)

with col2:
    week3_penalty = st.slider("Week 3 Penalty Weight", min_value=0, max_value=10, value=5)
    spread_penalty = st.slider("Exam Spacing Penalty Weight", min_value=0, max_value=10, value=5)
    extra_time_penalty = st.slider("Extra Time Student Penalty Weight", min_value=0, max_value=10, value=5)

# Fixed modules dictionary
Fixed_modules = {
    "BUSI60039 Business Strategy": [1,1],
    "BUSI60046 Project Management": [2,1],
    "ME-ELEC70098 Optimisation": [3,0],
    "MECH70001 Nuclear Thermal Hydraulics": [3,0],
    "BUSI60040/BUSI60043 Corporate Finance Online/Finance & Financial Management": [3,1],
    "MECH60004/MECH70042 Introduction to Nuclear Energy A/B": [4,0],
    "ME-ELEC70022 Modelling and Control of Multi-body Mechanical Systems": [4,0],
    "MATE97022 Nuclear Materials 1": [4,0],
    "ME-MATE70029 Nuclear Fusion": [9,0],
    "MECH70002 Nuclear Reactor Physics": [10,0],
    "ME-ELEC70076 Sustainable Electrical Systems": [10,0],
    "ME ELEC70066 Applied Advanced Optimisation": [10,0],
    "MECH70020 Combustion, Safety and Fire Dynamics": [11,0],
    "BIOE70016 Human Neuromechanical Control and Learning": [11,0],
    "CENG60013 Nuclear Chemical Engineering": [11,0],
    "MECH70008 Mechanical Transmissions Technology": [17,1],
    "MECH70006 Metal Processing Technology": [17,1],
    "MECH70021Aircraft Engine Technology": [17,1],
    "MECH70003 Future Clean Transport Technology": [17,1],
    "MECH60015/70030 PEN3/AME": [18,1]
}

# Core modules list
Core_modules = [
    "MECH70001 Nuclear Thermal Hydraulics",
    "MECH60004/MECH70042 Introduction to Nuclear Energy A/B",
    "MECH70002 Nuclear Reactor Physics",
    "MECH70008 Mechanical Transmissions Technology",
    "MECH70006 Metal Processing Technology",
    "MECH70021Aircraft Engine Technology",
    "MECH70003 Future Clean Transport Technology",
    "MECH60015/70030 PEN3/AME"
]

# Room dictionary with capacities and features
rooms = {
    'CAGB 203': [["Computer", "SEQ","AEA"], 65],
    'CAGB 309': [["SEQ"], 54],
    'CAGB 659-652': [["SEQ"], 75],
    'CAGB 747-748': [["SEQ"], 36],
    'CAGB 749-752': [["SEQ"], 75],
    'CAGB 761': [["Computer"], 25],
    'CAGB 762': [["Computer"], 25],
    'SKEM 208': [["Computer"], 35],
    'SKEM 317': [["Computer"], 20],
    'CAGB 320-321': [["AEA"], 10],
    'CAGB 305': [["AEA"], 4],
    'CAGB 349': [["AEA"], 2],
    'CAGB 311': [["AEA"], 1],
    'CAGB 765': [["AEA"], 10],
    'CAGB 527': [["AEA"], 2]
}

def ordinal(n):
    if 11 <= (n % 100) <= 13:
        return f"{n}th"
    else:
        return f"{n}{['th','st','nd','rd','th','th','th','th','th','th'][n % 10]}"

def validate_student_list(df):
    """Validate the student list Excel file format and content."""
    errors = []
    
    # Check if file has enough rows (header + at least one student)
    if len(df) < 3:
        errors.append("Student list must have at least 3 rows (header + students)")
        return errors
    
    # Check if required columns exist (CID and AEA status)
    if df.iloc[0, 0] != "CID" or df.iloc[0, 3] != "AEA":
        errors.append("Student list must have 'CID' in column A and 'AEA' in column D")
        return errors
    
    # Check if exam columns exist (starting from column J)
    exam_columns = df.iloc[0, 9:].dropna()
    if len(exam_columns) == 0:
        errors.append("No exam columns found starting from column J")
        return errors
    
    # Check for valid exam indicators (x, a, b)
    student_rows = df.iloc[2:, :]
    for idx, row in student_rows.iterrows():
        cid = row[0]
        if pd.isna(cid):
            errors.append(f"Missing CID in row {idx + 3}")
            continue
            
        # Check exam indicators
        for col_idx, exam_name in enumerate(exam_columns, start=9):
            value = str(row[col_idx]).strip().lower()
            if value not in ['x', 'a', 'b', 'nan']:
                errors.append(f"Invalid exam indicator '{value}' for student {cid} in exam {exam_name}")
    
    return errors

def validate_module_list(df):
    """Validate the module list Excel file format and content."""
    errors = []
    
    # Check if file has enough rows
    if len(df) < 2:
        errors.append("Module list must have at least 2 rows")
        return errors
    
    # Check required columns
    required_cols = ['Banner Code (New CR)', 'Module Name', 'Module Leader (lecturer 1)']
    for col in required_cols:
        if col not in df.columns:
            errors.append(f"Missing required column: {col}")
    
    # Check for missing values in required columns
    for col in required_cols:
        missing = df[col].isna().sum()
        if missing > 0:
            errors.append(f"Found {missing} missing values in column {col}")
    
    return errors

def validate_useful_dates(wb):
    """Validate the useful dates Excel file format and content."""
    errors = []
    
    if not wb:
        errors.append("Could not open useful dates file")
        return errors
    
    ws = wb.active
    
    # Check if file has bank holidays section
    found_bank_holidays = False
    for row in range(1, 10):
        cell_value = ws[f"F{row}"].value
        if cell_value and "Bank Holiday" in str(cell_value):
            found_bank_holidays = True
            break
    
    if not found_bank_holidays:
        errors.append("Could not find bank holidays section in useful dates file")
    
    # Check if file has summer term dates
    found_summer_term = False
    for row in range(1, ws.max_row):
        cell_value = ws[f"F{row}"].value
        if cell_value and "Summer Term" in str(cell_value):
            found_summer_term = True
            break
    
    if not found_summer_term:
        errors.append("Could not find summer term dates in useful dates file")
    
    return errors

def process_files():
    """Process uploaded files and validate their contents."""
    if not all([student_file, module_file, dates_file]):
        st.error("Please upload all required files")
        return None, None, None, None, None, None, None, None
    
    # Validate student list
    try:
        students_df = pd.read_excel(student_file, header=None)
        student_errors = validate_student_list(students_df)
        if student_errors:
            st.error("Student list validation errors:")
            for error in student_errors:
                st.error(f"- {error}")
            return None, None, None, None, None, None, None, None
    except Exception as e:
        st.error(f"Error reading student list: {str(e)}")
        return None, None, None, None, None, None, None, None
    
    # Validate module list
    try:
        leaders_df = pd.read_excel(module_file, sheet_name=1, header=1)
        module_errors = validate_module_list(leaders_df)
        if module_errors:
            st.error("Module list validation errors:")
            for error in module_errors:
                st.error(f"- {error}")
            return None, None, None, None, None, None, None, None
    except Exception as e:
        st.error(f"Error reading module list: {str(e)}")
        return None, None, None, None, None, None, None, None
    
    # Validate useful dates
    try:
        wb = load_workbook(dates_file)
        date_errors = validate_useful_dates(wb)
        if date_errors:
            st.error("Useful dates validation errors:")
            for error in date_errors:
                st.error(f"- {error}")
            return None, None, None, None, None, None, None, None
    except Exception as e:
        st.error(f"Error reading useful dates: {str(e)}")
        return None, None, None, None, None, None, None, None
    
    # If all validations pass, process the files
    try:
        # Extract exam names
        exams = students_df.iloc[0, 9:].dropna().tolist()
        
        # Create student exams dictionary
        student_rows = students_df.iloc[2:, :]
        student_exams = {}
        for _, row in student_rows.iterrows():
            cid = row[0]
            exams_taken = []
            for col_idx, exam_name in enumerate(exams, start=9):
                if str(row[col_idx]).strip().lower() in ['x', 'a', 'b']:
                    exams_taken.append(exam_name)
            student_exams[cid] = exams_taken
        
        # Extract AEA students
        valid_aea_mask = (
            student_rows.iloc[:, 3].notna() &
            (student_rows.iloc[:, 3].astype(str).str.strip() != "#N/A")
        )
        AEA = student_rows.loc[valid_aea_mask, student_rows.columns[0]].tolist()
        
        # Extract module leaders
        leader_courses = defaultdict(list)
        standardized_names = exams
        for _, row in leaders_df.iterrows():
            leader = row['Module Leader (lecturer 1)']
            name = row['Module Name']
            code = row['Banner Code (New CR)']
            
            if pd.isna(code) or pd.isna(name) or pd.isna(leader) or leader == "n/a":
                continue
                
            combined_name = f"{code} {name}"
            best_match, score, _ = process.extractOne(
                combined_name, standardized_names, scorer=fuzz.token_sort_ratio
            )
            
            if score >= 70:
                if best_match not in leader_courses[leader]:
                    leader_courses[leader].append(best_match)
        
        # Extract extra time students
        extra_time_students_25 = students_df[students_df.iloc[:, 3].astype(str).str.startswith(("15min/hour", "25% extra time"))].iloc[:, 0].tolist()
        extra_time_students_50 = students_df[students_df.iloc[:, 3].astype(str).str.startswith(("30min/hour", "50% extra time"))].iloc[:, 0].tolist()
        
        # Process useful dates
        ws = wb.active
        bank_holidays = []
        row = 5
        while True:
            name = ws[f"F{row}"].value
            date_cell = ws[f"G{row}"].value
            if name is None or "Term Dates" in str(name):
                break
            if isinstance(date_cell, datetime):
                bank_holidays.append((str(name).strip(), date_cell.date()))
            row += 1
        
        summer_start = None
        while row < ws.max_row:
            cell_value = ws[f"F{row}"].value
            if cell_value and "Summer Term" in str(cell_value):
                term_range = ws[f"F{row + 1}"].value
                if term_range:
                    try:
                        start_part = term_range.split("to")[0].strip()
                        start_str = re.sub(r"^\w+\s+", "", start_part)
                        year_match = re.search(r"\b\d{4}\b", term_range)
                        if year_match:
                            start_str += f" {year_match.group(0)}"
                        else:
                            raise ValueError("Year not found in date range.")
                        summer_start = parse(start_str, dayfirst=True).date()
                    except Exception as e:
                        st.error(f"Could not parse Summer Term start: {term_range}")
                        return None, None, None, None, None, None, None, None
                break
            row += 1
        
        if not summer_start:
            st.error("Summer Term start date not found")
            return None, None, None, None, None, None, None, None
        
        first_monday = summer_start
        while first_monday.weekday() != 0:
            first_monday += timedelta(days=1)
        
        no_exam_dates = [[5,0],[5,1],[6,0],[6,1],[12,0],[12,1],[13,0],[13,1],[18,0],[19,0],[19,1],[20,0],[20,1]]
        for name, bh_date in bank_holidays:
            delta = (bh_date - first_monday).days
            if 0 <= delta <= 20:
                no_exam_dates.append([delta, 0])
                no_exam_dates.append([delta, 1])
        
        days = []
        for i in range(21):
            date = first_monday + timedelta(days=i)
            day_str = date.strftime("%A ") + ordinal(date.day) + date.strftime(" %B")
            days.append(day_str)
        
        return exams, student_exams, AEA, leader_courses, extra_time_students_25, extra_time_students_50, no_exam_dates, days
        
    except Exception as e:
        st.error(f"Error processing files: {str(e)}")
        return None, None, None, None, None, None, None, None

def create_timetable(exams, student_exams, leader_courses, no_exam_dates, extra_time_students_25, extra_time_students_50, AEA, exam_counts, days):
    model = cp_model.CpModel()
    slots = [0, 1]  # 0 for morning, 1 for afternoon
    num_slots = len(slots)

    # Variables
    exam_day = {}
    exam_slot = {}
    for exam in exams:
        exam_day[exam] = model.NewIntVar(0, num_days - 1, f'{exam}_day')
        exam_slot[exam] = model.NewIntVar(0, num_slots - 1, f'{exam}_slot')

    # Room assignment variables
    exam_room = {}
    for exam in exams:
        for room in rooms:
            exam_room[(exam, room)] = model.NewBoolVar(f'{exam}_in_{room.replace(" ", "_")}')

    # Constraint 1: Core modules can't be on the same day as other modules
    for student, exs in student_exams.items():
        core_mods = [exam for exam in exs if exam in Core_modules]
        other_mods = [exam for exam in exs if exam not in Core_modules]
        for exam in core_mods:
            for other in other_mods:
                model.Add(exam_day[exam] != exam_day[other])

    # Constraint 2: Fixed module dates
    for exam, (day_fixed, slot_fixed) in Fixed_modules.items():
        if exam in exams:
            model.Add(exam_day[exam] == day_fixed)
            model.Add(exam_slot[exam] == slot_fixed)

    # Constraint 3: Forbidden exam day-slot assignments
    for exam in exams:
        for day, slot in no_exam_dates:
            model.AddForbiddenAssignments([exam_day[exam], exam_slot[exam]], [(day, slot)])

    # Constraint 4: Max 3 exams in any 2-day window per student
    for student, exs in student_exams.items():
        for d in range(num_days - 1):
            exams_in_2_days = []
            for exam in exs:
                is_on_d = model.NewBoolVar(f'{student}_{exam}_on_day_{d}')
                is_on_d1 = model.NewBoolVar(f'{student}_{exam}_on_day_{d+1}')
                is_on_either = model.NewBoolVar(f'{student}_{exam}_on_day_{d}_or_{d+1}')

                model.Add(exam_day[exam] == d).OnlyEnforceIf(is_on_d)
                model.Add(exam_day[exam] != d).OnlyEnforceIf(is_on_d.Not())
                model.Add(exam_day[exam] == d + 1).OnlyEnforceIf(is_on_d1)
                model.Add(exam_day[exam] != d + 1).OnlyEnforceIf(is_on_d1.Not())
                model.AddBoolOr([is_on_d, is_on_d1]).OnlyEnforceIf(is_on_either)
                model.AddBoolAnd([is_on_d.Not(), is_on_d1.Not()]).OnlyEnforceIf(is_on_either.Not())

                exams_in_2_days.append(is_on_either)

            model.Add(sum(exams_in_2_days) <= max_exams_2days)

    # Constraint 5: Max 4 exams in any 5-day window per student
    for student, exs in student_exams.items():
        for start_day in range(num_days - 4):
            exams_in_window = []
            for exam in exs:
                in_window = model.NewBoolVar(f'{student}_{exam}_in_day_{start_day}_to_{start_day + 4}')
                model.AddLinearConstraint(exam_day[exam], start_day, start_day + 4).OnlyEnforceIf(in_window)
                before_window = model.NewBoolVar(f'{student}_{exam}_before_{start_day}')
                after_window = model.NewBoolVar(f'{student}_{exam}_after_{start_day + 4}')
                model.Add(exam_day[exam] < start_day).OnlyEnforceIf(before_window)
                model.Add(exam_day[exam] >= start_day).OnlyEnforceIf(before_window.Not())
                model.Add(exam_day[exam] > start_day + 4).OnlyEnforceIf(after_window)
                model.Add(exam_day[exam] <= start_day + 4).OnlyEnforceIf(after_window.Not())
                model.AddBoolOr([before_window, after_window]).OnlyEnforceIf(in_window.Not())
                exams_in_window.append(in_window)
            model.Add(sum(exams_in_window) <= max_exams_5days)

    # Constraint 6: At most 1 exam in week 3 per module leader
    for leader, leader_exams in leader_courses.items():
        exams_in_week3 = []
        for exam in leader_exams:
            in_week3 = model.NewBoolVar(f'{exam}_in_week3')
            model.AddLinearConstraint(exam_day[exam], 13, 20).OnlyEnforceIf(in_week3)
            before_week3 = model.NewBoolVar(f'{exam}_before_week3')
            after_week3 = model.NewBoolVar(f'{exam}_after_week3')
            model.Add(exam_day[exam] < 13).OnlyEnforceIf(before_week3)
            model.Add(exam_day[exam] >= 13).OnlyEnforceIf(before_week3.Not())
            model.Add(exam_day[exam] > 20).OnlyEnforceIf(after_week3)
            model.Add(exam_day[exam] <= 20).OnlyEnforceIf(after_week3.Not())
            model.AddBoolOr([before_week3, after_week3]).OnlyEnforceIf(in_week3.Not())
            exams_in_week3.append(in_week3)
        model.Add(sum(exams_in_week3) <= 1)

    # Constraint 7: Extra time 50% students: max 1 exam per day
    for student in extra_time_students_50:
        for day in range(num_days):
            exams_on_day = []
            for exam in student_exams[student]:
                is_on_day = model.NewBoolVar(f'{student}_{exam}_on_day_{day}')
                model.Add(exam_day[exam] == day).OnlyEnforceIf(is_on_day)
                model.Add(exam_day[exam] != day).OnlyEnforceIf(is_on_day.Not())
                exams_on_day.append(is_on_day)
            model.Add(sum(exams_on_day) <= 1)

    # Constraint 8: Room assignments
    for exam in exams:
        # Each exam must be assigned to at least one room
        model.Add(sum(exam_room[(exam, room)] for room in rooms) >= 1)
        
        # Room capacity constraints
        for room, (features, capacity) in rooms.items():
            # Get number of AEA and non-AEA students for this exam
            aea_count, non_aea_count = exam_counts[exam]
            
            # If room has AEA feature, it can take AEA students
            if "AEA" in features:
                model.Add(aea_count * exam_room[(exam, room)] <= capacity)
            else:
                # If room doesn't have AEA feature, it can't be used for AEA students
                model.Add(aea_count * exam_room[(exam, room)] == 0)
            
            # Total capacity constraint
            model.Add((aea_count + non_aea_count) * exam_room[(exam, room)] <= capacity)

    # Soft constraints and penalties
    soft_penalties = []
    for student in extra_time_students_25:
        for day in range(num_days):
            exams_on_day = []
            for exam in student_exams[student]:
                is_on_day = model.NewBoolVar(f'{student}_{exam}_on_day_{day}')
                model.Add(exam_day[exam] == day).OnlyEnforceIf(is_on_day)
                model.Add(exam_day[exam] != day).OnlyEnforceIf(is_on_day.Not())
                exams_on_day.append(is_on_day)

            num_exams = model.NewIntVar(0, len(exams_on_day), f'{student}_num_exams_day_{day}')
            model.Add(num_exams == sum(exams_on_day))

            has_multiple_exams = model.NewBoolVar(f'{student}_more_than_one_exam_day_{day}')
            model.Add(num_exams >= 2).OnlyEnforceIf(has_multiple_exams)
            model.Add(num_exams < 2).OnlyEnforceIf(has_multiple_exams.Not())

            penalty = model.NewIntVar(0, 1, f'{student}_penalty_day_{day}')
            model.Add(penalty == 1).OnlyEnforceIf(has_multiple_exams)
            model.Add(penalty == 0).OnlyEnforceIf(has_multiple_exams.Not())

            soft_penalties.append(extra_time_penalty * penalty)

    # Spread penalties for module leaders
    spread_penalties = []
    for leader in leader_courses:
        mods = leader_courses[leader]
        for i in range(len(mods)):
            for j in range(i+1, len(mods)):
                m1 = mods[i]
                m2 = mods[j]
                diff = model.NewIntVar(-21, 21, f'{m1}_{m2}_diff')
                abs_diff = model.NewIntVar(0, 21, f'{m1}_{m2}_abs_diff')
                model.Add(diff == exam_day[m1] - exam_day[m2])
                model.AddAbsEquality(abs_diff, diff)

                close_penalty = model.NewIntVar(0, 5, f'{m1}_{m2}_penalty')
                is_gap_3 = model.NewBoolVar(f'{m1}_{m2}_gap3')
                is_gap_2 = model.NewBoolVar(f'{m1}_{m2}_gap2')
                is_gap_1 = model.NewBoolVar(f'{m1}_{m2}_gap1')
                is_gap_0 = model.NewBoolVar(f'{m1}_{m2}_gap0')

                model.Add(abs_diff == 3).OnlyEnforceIf(is_gap_3)
                model.Add(abs_diff != 3).OnlyEnforceIf(is_gap_3.Not())
                model.Add(abs_diff == 2).OnlyEnforceIf(is_gap_2)
                model.Add(abs_diff != 2).OnlyEnforceIf(is_gap_2.Not())
                model.Add(abs_diff == 1).OnlyEnforceIf(is_gap_1)
                model.Add(abs_diff != 1).OnlyEnforceIf(is_gap_1.Not())
                model.Add(abs_diff == 0).OnlyEnforceIf(is_gap_0)
                model.Add(abs_diff != 0).OnlyEnforceIf(is_gap_0.Not())

                model.Add(close_penalty == 1).OnlyEnforceIf(is_gap_3)
                model.Add(close_penalty == 3).OnlyEnforceIf(is_gap_2)
                model.Add(close_penalty == 4).OnlyEnforceIf(is_gap_1)
                model.Add(close_penalty == 5).OnlyEnforceIf(is_gap_0)
                model.Add(close_penalty == 0).OnlyEnforceIf(
                    is_gap_3.Not(), is_gap_2.Not(), is_gap_1.Not(), is_gap_0.Not()
                )
                spread_penalties.append(spread_penalty * close_penalty)

    # Minimize total penalties
    model.Minimize(sum(spread_penalties) + sum(soft_penalties))

    # Solve
    solver = cp_model.CpSolver()
    status = solver.Solve(model)

    if status == cp_model.FEASIBLE or status == cp_model.OPTIMAL:
        timetable = {}
        room_assignments = {}
        for exam in exams:
            d = solver.Value(exam_day[exam])
            s = solver.Value(exam_slot[exam])
            timetable[exam] = (d, s)
            
            # Get room assignments
            assigned_rooms = [room for room in rooms if solver.Value(exam_room[(exam, room)])]
            room_assignments[exam] = assigned_rooms
            
        return timetable, room_assignments
    else:
        return None, None

def visualize_timetable(timetable, room_assignments, days):
    if not timetable:
        return None

    # Create a DataFrame for visualization
    data = []
    for exam, (day, slot) in timetable.items():
        rooms_str = ", ".join(room_assignments[exam])
        data.append({
            'Exam': exam,
            'Day': days[day],
            'Slot': 'Morning' if slot == 0 else 'Afternoon',
            'Rooms': rooms_str
        })
    
    df = pd.DataFrame(data)
    
    # Create a heatmap
    pivot = pd.crosstab(df['Day'], df['Slot'])
    
    fig = go.Figure(data=go.Heatmap(
        z=pivot.values,
        x=pivot.columns,
        y=pivot.index,
        colorscale='Viridis'
    ))
    
    fig.update_layout(
        title='Exam Timetable Heatmap',
        xaxis_title='Time Slot',
        yaxis_title='Day'
    )
    
    return fig, df

def check_exam_constraints(student_exams, exams_timetabled, Fixed_modules, Core_modules, module_leaders, extra_time_students_50, AEA):
    """Check if the generated timetable satisfies all exam constraints."""
    violations = []
    schedule = get_full_schedule(exams_timetabled, Fixed_modules)

    # 1. No more than 3 exams in any 2 consecutive days (per student)
    for student, exams in student_exams.items():
        day_count = defaultdict(int)
        for exam in exams:
            if exam in schedule:
                day = schedule[exam][0]
                day_count[day] += 1

        days = sorted(day_count.keys())
        for day in days:
            next_day = day + 1
            if next_day in day_count:
                total = day_count[day] + day_count[next_day]
                if total > 3:
                    violations.append(
                        f"❌ Student {student} has more than 3 exams across days {day} and {next_day}"
                    )

    # 2. No more than 4 exams in any 5 consecutive weekdays
    for student, exams in student_exams.items():
        day_count = defaultdict(int)
        for exam in exams:
            if exam in schedule:
                day = schedule[exam][0]
                day_count[day] += 1

        all_days = sorted(day_count.keys())
        if all_days:
            min_day, max_day = all_days[0], all_days[-1]
            for start_day in range(min_day, max_day - 4 + 1):
                total = sum(day_count.get(day, 0) for day in range(start_day, start_day + 5))
                if total > 4:
                    violations.append(
                        f"❌ Student {student} has more than 4 exams from day {start_day} to {start_day + 4}"
                    )

    # 3. Core modules can't be on same day as other modules
    for student, exams in student_exams.items():
        core_mods = [exam for exam in exams if exam in Core_modules]
        other_mods = [exam for exam in exams if exam not in Core_modules]

        for core_exam in core_mods:
            if core_exam in exams_timetabled:
                core_day = exams_timetabled[core_exam][0]
                for other_exam in other_mods:
                    if other_exam in exams_timetabled:
                        other_day = exams_timetabled[other_exam][0]
                        if core_day == other_day:
                            violations.append(
                                f"❌ Student {student} has core exam '{core_exam}' and non-core exam '{other_exam}' on the same day ({core_day})"
                            )

    # 4. Module leaders can't have more than one exam in week 3
    week3_days = set(range(15, 21))
    for leader, mods in module_leaders.items():
        exams_in_week3 = [exam for exam in mods if exam in schedule and schedule[exam][0] in week3_days]
        if len(exams_in_week3) > 1:
            violations.append(f"❌ Module leader {leader} has more than one exam in week 3: {exams_in_week3}")

    # 5. Students with >50% extra time can't have more than one exam per day
    for student in extra_time_students_50:
        if student not in student_exams:
            continue
        day_count = defaultdict(int)
        for exam in student_exams[student]:
            if exam in schedule:
                day = schedule[exam][0]
                day_count[day] += 1
        for day, count in day_count.items():
            if count > 1:
                violations.append(f"❌ Student {student} with >50% extra time has {count} exams on day {day}")

    # 6. Students with 25% extra time (soft constraint)
    for student in AEA:
        if student not in extra_time_students_50:
            day_count = defaultdict(int)
            for exam in student_exams[student]:
                if exam in schedule:
                    day = schedule[exam][0]
                    day_count[day] += 1
            for day, count in day_count.items():
                if count > 1:
                    violations.append(f"⚠️soft warning Student {student} with <=25% extra time has {count} exams on day {day}")

    return violations

def check_room_constraints(exams_timetabled, exam_counts, room_dict):
    """Check if the room assignments satisfy all constraints."""
    violations = []

    # 1. No room double-booked at same day & slot
    room_schedule = defaultdict(list)  # key=(day, slot, room), value=list of exams
    for exam, (day, slot, rooms_) in exams_timetabled.items():
        for room in rooms_:
            room_schedule[(day, slot, room)].append(exam)

    for (day, slot, room), exams_in_room in room_schedule.items():
        if len(exams_in_room) > 1:
            violations.append(
                f"❌ Room '{room}' double-booked on day {day}, slot {slot} for exams: {exams_in_room}"
            )

    # 2. Check every exam assigned at least one room
    for exam, (day, slot, rooms) in exams_timetabled.items():
        if not rooms:
            violations.append(f"❌ Exam '{exam}' has no assigned room!")

    # 3. Check room capacity sufficiency per exam
    for exam, (day, slot, rooms) in exams_timetabled.items():
        if exam not in exam_counts:
            violations.append(f"⚠️ No student count for exam '{exam}', skipping capacity check")
            continue

        AEA_students, SEQ_students = exam_counts[exam]
        AEA_capacity = sum(room_dict[r][1] for r in rooms if "AEA" in room_dict[r][0])
        SEQ_capacity = sum(room_dict[r][1] for r in rooms if "SEQ" in room_dict[r][0])
        
        if AEA_capacity < AEA_students:
            violations.append(
                f"❌ Exam '{exam}' has insufficient AEA capacity: needed {AEA_students}, assigned {AEA_capacity}"
            )
        if SEQ_capacity < SEQ_students:
            violations.append(
                f"❌ Exam '{exam}' has insufficient SEQ capacity: needed {SEQ_students}, assigned {SEQ_capacity}"
            )

    return violations

def get_full_schedule(exams_timetabled, Fixed_modules):
    """Combine fixed modules with dynamically assigned ones."""
    full_schedule = Fixed_modules.copy()
    full_schedule.update(exams_timetabled)
    return full_schedule

def validate_timetable(timetable, room_assignments, student_exams, leader_courses, extra_time_students_50, AEA, exam_counts):
    """Validate the generated timetable against all constraints."""
    if not timetable or not room_assignments:
        return ["❌ No timetable generated"]
        
    violations = []
    
    # Check exam constraints
    exam_violations = check_exam_constraints(
        student_exams=student_exams,
        exams_timetabled=timetable,
        Fixed_modules=Fixed_modules,
        Core_modules=Core_modules,
        module_leaders=leader_courses,
        extra_time_students_50=extra_time_students_50,
        AEA=AEA
    )
    violations.extend(exam_violations)
    
    # Check room constraints
    room_violations = check_room_constraints(
        exams_timetabled={exam: (day, slot, room_assignments[exam]) for exam, (day, slot) in timetable.items()},
        exam_counts=exam_counts,
        room_dict=rooms
    )
    violations.extend(room_violations)
    
    return violations

if st.button("Generate Timetable"):
    with st.spinner("Processing files and generating timetable..."):
        result = process_files()
        if all(result):
            exams, student_exams, AEA, leader_courses, extra_time_students_25, extra_time_students_50, no_exam_dates, days = result
            
            # Calculate exam counts for AEA and non-AEA students
            exam_counts = defaultdict(lambda: [0, 0])
            for cid, exams_taken in student_exams.items():
                if cid in AEA:
                    for exam in exams_taken:
                        exam_counts[exam][0] += 1
                else:
                    for exam in exams_taken:
                        exam_counts[exam][1] += 1
            
            timetable, room_assignments = create_timetable(exams, student_exams, leader_courses, no_exam_dates, extra_time_students_25, extra_time_students_50, AEA, exam_counts, days)
            
            if timetable and room_assignments:
                # Validate the timetable
                violations = validate_timetable(timetable, room_assignments, student_exams, leader_courses, extra_time_students_50, AEA, exam_counts)
                
                if violations:
                    st.warning("⚠️ Timetable generated with some issues:")
                    for violation in violations:
                        st.write(violation)
                else:
                    st.success("✅ Timetable generated successfully with no constraint violations!")
                
                # Display the timetable
                st.header("Generated Timetable")
                fig, df = visualize_timetable(timetable, room_assignments, days)
                if fig:
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # Display detailed timetable
                    st.subheader("Detailed Timetable")
                    st.dataframe(df)
                
                # Export option
                if st.button("Export Timetable"):
                    csv = df.to_csv(index=False)
                    st.download_button(
                        "Download CSV",
                        csv,
                        "exam_timetable.csv",
                        "text/csv",
                        key='download-csv'
                    )
            else:
                st.error("❌ Failed to generate a valid timetable. Please check your input data and constraints.")
        else:
            st.error("❌ Failed to process input files. Please check the file formats and try again.") 