# Exam Timetabling System

A website that helps optimize the timetabling of exams for a set of soft and hard constraints

## Notion page

There's documentation and a collation of research on the [notion page](https://tin-fog-24f.notion.site/209713b1f81d80d1aae3f5726b91a131?v=20b713b1f81d803ea7f3000ccb59e68d). 

## Constraints currently considered

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
- CSV export option 
