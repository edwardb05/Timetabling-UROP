# %% [markdown]
# # Exam Timetabling with Morning and Afternoon Slots using OR-Tools
# 
# This notebook models exams scheduled over multiple days with morning and afternoon slots.
# 
# Constraints:
# - No student can have two exams in the same slot (same day & same slot)
# - No more than 2 exams per slot (day + morning/afternoon)
# 
# Outputs the exam schedule showing day and slot assignment.

# %%
# Install OR-Tools if you haven't yet
#!pip install ortools

# %%
from ortools.sat.python import cp_model

# %%
# Sample data
exams = ['Math', 'Physics', 'Chemistry', 'History', 'Geography', 'PE', 'Music']

# Students and their exams
students = {
    'Alice': ['Math', 'Physics', 'Geography'],
    'Bob': ['Physics', 'Chemistry', 'Music'],
    'Charlie': ['Math', 'History', 'PE'],
    'Diana': ['Chemistry', 'History', 'Music'],
}

# Module leaders for info
module_leaders = {
    'Math': 'Dr. Smith',
    'Physics': 'Dr. Johnson',
    'Chemistry': 'Dr. Lee',
    'History': 'Dr. Patel',
    'Geography': 'Dr. Brown',
    'PE': 'Coach Carter',
    'Music': 'Ms. Green',
}

# Days and slots
days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday','Friday',"saturday","sunday"]
slots = ['Morning', 'Afternoon']

# %%
model = cp_model.CpModel()

num_days = len(days)
num_slots = len(slots)

# Variables: exam_day and exam_slot
exam_day = {}
exam_slot = {}
for exam in exams:
    exam_day[exam] = model.NewIntVar(0, num_days - 1, f'{exam}_day')
    exam_slot[exam] = model.NewIntVar(0, num_slots - 1, f'{exam}_slot')

# %%
# Constraint 1: No student can have two exams at the same time (same day and same slot)
for student, stu_exams in students.items():
    for i in range(len(stu_exams)):
        for j in range(i + 1, len(stu_exams)):
          # Create boolean variables for the conditions
          diff_day = model.NewBoolVar('diff_day')
          diff_slot = model.NewBoolVar('diff_slot')

          # Link them to the actual constraints
          model.Add(exam_day[stu_exams[i]] != exam_day[stu_exams[j]]).OnlyEnforceIf(diff_day)
          model.Add(exam_day[stu_exams[i]] == exam_day[stu_exams[j]]).OnlyEnforceIf(diff_day.Not())

          model.Add(exam_slot[stu_exams[i]] != exam_slot[stu_exams[j]]).OnlyEnforceIf(diff_slot)
          model.Add(exam_slot[stu_exams[i]] == exam_slot[stu_exams[j]]).OnlyEnforceIf(diff_slot.Not())

          # Enforce that at least one of them is true
          model.AddBoolOr([diff_day, diff_slot])


# %%
for d in range(num_days):
    for s in range(num_slots):
        exams_in_slot = []
        for exam in exams:
            is_day = model.NewBoolVar(f'{exam}_is_day{d}')
            is_slot = model.NewBoolVar(f'{exam}_is_slot{s}')
            is_in_slot = model.NewBoolVar(f'{exam}_in_day{d}_slot{s}')

            # Link day and slot indicators
            model.Add(exam_day[exam] == d).OnlyEnforceIf(is_day)
            model.Add(exam_day[exam] != d).OnlyEnforceIf(is_day.Not())

            model.Add(exam_slot[exam] == s).OnlyEnforceIf(is_slot)
            model.Add(exam_slot[exam] != s).OnlyEnforceIf(is_slot.Not())

            # is_in_slot = is_day AND is_slot
            model.AddBoolAnd([is_day, is_slot]).OnlyEnforceIf(is_in_slot)
            model.AddBoolOr([is_day.Not(), is_slot.Not()]).OnlyEnforceIf(is_in_slot.Not())

            exams_in_slot.append(is_in_slot)

        # Enforce max 2 exams per slot
        model.Add(sum(exams_in_slot) <= 2)


# %%
# Solve the model
solver = cp_model.CpSolver()
status = solver.Solve(model)

if status == cp_model.FEASIBLE or status == cp_model.OPTIMAL:
    print("Exam Schedule:")
    for exam in exams:
        d = solver.Value(exam_day[exam])
        s = solver.Value(exam_slot[exam])
        leader = module_leaders.get(exam, 'Unknown')
        print(f" - {exam} (Leader: {leader}): {days[d]} {slots[s]}")
else:
    print("No solution found.")


