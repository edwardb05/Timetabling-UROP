# Handling data

Owner: Edward Brady

For this project the three input excel files need to be turned into data that can be accessed by the model. The way this is done can be seen in this notebook.

https://github.com/edwardb05/Timetabling-UROP/blob/main/Data_handling.ipynb

## What is needed from the excel file:

- List of exams named `‘exams’`
- Integer of the number of days for exam period named `‘num_days’`
- Dictionary which has students CID and then a list of there exams

```python
student_exams ={
	001 : ["Maths","Physics"]
	002 : ["Chemistry","Maths"]
	}
```

- List of core modules/exams named `“Core_modules”`
- Dictionary of fixed exams and their days called `“fixed_modules”` this will have the day it is on and the slot, morning or afternoon
- Dictionary of rooms and their uses ( AEA, SEQ or computer) and capacity in the format `{room: [List of uses],[Capacity]}` called `rooms` this can be updated to reflect changes
- List of days and times that can’t have exams such as weekends or bank holidys called `“no_exam_dates”`
- List called `No_exam_dates_soft` for dates that shouldn't have an exam on such as Tuesday and Wednesday morning of week 3
- Dictionary of Module leaders and their modules called `“leader_courses”`
- List of students with 25% Extra time called `“extra_time_students_25”`
- List of students with 50% Extra time called `“extra_time_students_50”`
- List of the exam dates called `days`
- List of any student with AEA called `AEA`
- Dictionary called `exam _counts` in the format `{exam:[AEA students taking it,SEQ students taking it]}`
- Dictionary called `exam_types` in the format `{exam : type}` where type can be Standard (flat floor) or PC requiring a computer room

## Issues:

- The exam modules must all have the same name so when taking the names from the 2 different files it is important to only select one.
    
    
    ```python
    # Extract official module names from row 0, columns J onwards (i.e., column 9 onward, 0-indexed)
    standardized_names = students_df.iloc[0, 9:].dropna().astype(str).str.strip().tolist()
    
    # Prepare module-leader dictionary
    leader_courses = defaultdict(list)
    
    # Loop through rows in the module list
    for _, row in leaders_df.iterrows():
        leader = row['Module Leader (lecturer 1)']
        name = row['Module Name']
        code = row['Banner Code (New CR)']   # module leader
    
        # Skip if any required field is missing
        if pd.isna(code) or pd.isna(name) or pd.isna(leader):
            continue
    
        if leader == "n/a":
            continue
    
        # Combine code and name
        combined_name = f"{code} {name}"
    
        # Fuzzy match to standardized names
        best_match, score, _ = process.extractOne(
            combined_name, standardized_names, scorer=fuzz.token_sort_ratio
        )
    
        if score >= 70:
            if best_match not in leader_courses[leader]:
                leader_courses[leader].append(best_match)
            else:
                print(f"⚠️ Duplicate match skipped for '{combined_name}': '{best_match}' is already listed for {leader}.")
        else:
            print(f"⚠️ Low confidence match for '{combined_name}' (best: '{best_match}', score: {score}).")
    
    # Convert to normal dict if desired
    leader_courses = dict(leader_courses)
    
    print("Module leaders and their courses:")
    print(leader_courses)
    ```
    
    This code simply adds the module code and name together from the module leader spreadsheet and then check it against the list of names from the student spreadsheet. It then checks for the closest one and ensures there score is higher than 70. Sometimes modules were being duplicated due to there being A and B and it producing the module name with A/B so that is removed in the duplicate section.
    
- Extracting the date of the start of exam period from the excel file
    
    ```python
    row = 5
    while True:
        name = ws[f"F{row}"].value
        date_cell = ws[f"G{row}"].value
        if name is None or "Term Dates" in str(name):
            break
        if isinstance(date_cell, datetime):
            bank_holidays.append((str(name).strip(), date_cell.date()))
        row += 1
    
    # --- Step 2: Find Summer Term start date from section below ---
    summer_start = None
    while row < ws.max_row:
        cell_value = ws[f"F{row}"].value
        if cell_value and "Summer Term" in str(cell_value):
            term_range = ws[f"F{row + 1}"].value
            if term_range:
                try:
                    # Extract left side of range before "to", drop weekday (e.g., "Fri"), and append year
                    start_part = term_range.split("to")[0].strip()
                    start_str = re.sub(r"^\w+\s+", "", start_part)  # Removes "Fri", leaves "24 Apr"
                    # Try extracting year from second part if present
                    year_match = re.search(r"\b\d{4}\b", term_range)
                    if year_match:
                        start_str += f" {year_match.group(0)}"
                    else:
                        raise ValueError("Year not found in date range.")
                    summer_start = parse(start_str, dayfirst=True).date()
                except Exception as e:
                    raise ValueError(f"Could not parse Summer Term start: {term_range}") from e
            else:
                raise ValueError("Summer Term range cell is empty.")
            break
        row += 1
    ```
    
    Searches for bank holidays and makes a list of them and their dates 
    
    Then to find the start date of summer term it finds the cell containing summer term and then moves to the cell below. It extracts the day and month from the first part and then the year from the second part
    
- Which rooms are eligible for AEA students and how to prioritize (UNSOLVED)
    
    On the excel spreadsheet it says the purpose of each room but i can see from last year that 203 was used for PEN AEA students. Can we also use aero rooms. Once solved edit the dictionary accordingly.