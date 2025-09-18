# Displaying the timetable

Owner: Edward Brady

Once the solver has found a solution for the timetable it will have a day and a slot for each exam. This needs to then be converted into something readable, I think the easiest way to do this will to be to use the template from last years exams and update the dates and exam slots. This is done in two steps:

1. Create an excel spreadsheet with the data
2. Format the spreadsheet to ensure that it has;
    1. Merged days the length of the morning and afternoon cells
    2. Merged Morning and afternoon cells to ensure they’re as long as their exams
    3. Alternate fill each day green and blue
    4. Fill core modules red and fixed modules yellow
    5. An empty row between each slot for clearer reading