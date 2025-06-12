# Exam Timetabling System

A Streamlit application that helps create optimal exam timetables while considering various constraints and preferences.

## Features

- Interactive web interface for easy timetable generation
- Handles student exam conflicts
- Considers module leader preferences
- Accommodates students with extra time requirements
- Respects fixed exam dates and bank holidays
- Visualizes the generated timetable
- Exports timetable to CSV format

## Installation

1. Clone this repository
2. Install the required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

1. Run the Streamlit app:
```bash
streamlit run app.py
```

2. Upload the required Excel files:
   - Student List (containing student IDs and their exams)
   - Module List (containing module information and leaders)
   - Useful Dates (containing term dates and bank holidays)

3. Adjust the timetabling parameters as needed:
   - Number of days for the exam period
   - Maximum exams in 2-day window
   - Maximum exams in 5-day window
   - Week 3 penalty weight
   - Exam spacing penalty weight

4. Click "Generate Timetable" to create the optimal schedule

5. View the generated timetable and export it to CSV if needed

## Input File Formats

### Student List
- Column A: Student CID
- Column D: Extra time accommodations
- Columns J onwards: Exam selections (marked with 'x')

### Module List
- Sheet 2 (index 1)
- Headers start from row 2
- Required columns:
  - Module Leader (lecturer 1)
  - Module Name
  - Banner Code (New CR)

### Useful Dates
- Contains term dates and bank holidays
- Summer Term start date is used to calculate the exam period
- Bank holidays are automatically excluded from the timetable

## Constraints

The system considers the following constraints:
1. No more than 3 exams in a 2-day window
2. No more than 4 exams in a 5-day window
3. Core modules can't be on the same day as other modules
4. Module leaders can have at most one exam in week 3
5. Students with 50% extra time can't have more than one exam per day
6. No exams on weekends or bank holidays
7. Soft penalties for students with 25% extra time having multiple exams per day
8. Penalties for exams being too close together

## Output

The generated timetable includes:
- Exam name
- Day of the week
- Time slot (Morning/Afternoon)
- Visual heatmap representation
- CSV export option 