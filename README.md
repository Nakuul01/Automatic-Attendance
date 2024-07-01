Student Attendance Management Script
Overview
This script is designed to manage student attendance for both offline and online classes. It reads student data from a CSV file, allows manual attendance marking, and saves the results to an Excel file. This tool is useful for educators and administrators to efficiently track and record student attendance.

                                                                Student Attendance Management Script

Overview
This script is designed to manage student attendance for both offline and online classes. It reads student data from a CSV file, allows manual attendance marking, and saves the results to an Excel file. This tool is useful for educators and administrators to efficiently track and record student attendance.

Features
Read Student Data: Reads student names and their attendance categories (offline/online) from a CSV file.
Interactive Attendance Marking: Allows manual marking of attendance via interactive prompts.
Excel Export: Saves the attendance data to an Excel file with automatic naming based on the current date and attendance type.

Prerequisites
Python 3.x
pip install xlsxwriter

Usage:
Prepare the CSV File:

Create a CSV file with student names and their attendance categories (offline/online).
Save the file at /Users//Downloads/.csv or modify the file_path variable in the script to match your CSV file location.


Mark Attendance:

The script will prompt you to choose the attendance type (offline or online).
Enter student names or parts of their names to mark them as present.
Use the 4-digit code 0000 to stop marking attendance.

View Attendance Record:

The file will be named based on the current date and attendance type, e.g., attendance_offline_YYYY-MM-DD.xlsx.


Functions:

read_student_names(file_path)
Reads student names and categories from a CSV file.
Organizes names into offline and online categories.
flatten_student_names(student_names)
Flattens the nested dictionary of student names for easier access.
mark_attendance(student_names, attendance_type)
Prompts the user to mark attendance interactively.
Handles full and partial name matches and allows for disambiguation.
save_attendance_excel(attendance, attendance_type, directory)
Saves the attendance data to an Excel file.
main()
Coordinates the script's workflow: reading data, choosing attendance type, marking attendance, and saving the record.


Acknowledgements:

xlsxwriter
Python Standard Library
