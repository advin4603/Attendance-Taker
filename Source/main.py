"""
main.py - Compares a list of present students with total strength to determine absentees and unrecognized students.
Author: Ayan Datta
"""
import time
import pyperclip
from openpyxl import load_workbook
import os
import sys
from datetime import datetime


def get_data(excel_file_path: str, sheet: str = "Main Sheet"):
    """
        Gets the first column of the spreadsheet in order and returns a dictionary with key as row number and the value of the cell.
    """
    return {i: x[0].value for i, x in enumerate(load_workbook(excel_file_path)[sheet].rows, start=1) if x[0].value}


def main():
    print("\n"+"-"*20 + "\n")
    print("--Attendance-Taker--")

    # Check if logs folder exists. If not then create one
    if not os.path.isdir("Logs"):
        os.mkdir("Logs")

    if not os.path.isdir("Attendances"):
        os.mkdir("Attendances")

    if not os.path.isdir("Tracebacks"):
        os.mkdir("Tracebacks")

    # Check if data file exists. If not then prompt user to create one.
    if not os.path.isfile("data.xlsx"):
        input(
            f"data.xlsx not found. Create an excel workbook with a spreadsheet Main Sheet and store student names in order in column A1 in {os.getcwd()}.\nPress Enter to Quit.\n{'-'*20}\n")
        sys.exit()

    # Get Dictionary containing roll numbers as keys and names as values from data.xlsx
    student_data = get_data("data.xlsx")

    # Make a backup of present students in a logs folder with the time. Generate a name for the file using the current time.
    name = f"Logs\\Attendance;{time.ctime()}.txt".replace(
        " ", ";").replace(":", "-")
    with open(name, "w") as f:
        # Get present students from clipboard.
        s = pyperclip.paste()
        print(s, file=f)

    # Open the file of all present students to read all present students.
    with open(name, "r") as file:
        # Separate all students in a list line by line.
        lines = [i.strip("\n") for i in file.readlines()]
        # Make a list to store the indices of the lines containing names of recognized students that need to be deleted from the list of lines so that only unrecognized names remain.
        remove_line = []

        # Loop over every line.
        for line in lines:
            # Start looking for the name in the database of students.
            for key, val in student_data.items():
                # Check if name matches the one in database.
                if val in line:
                    # If match found then remove the name from the database so that it is not checked for again and add the line to remove lines.
                    del student_data[key]
                    remove_line.append(line)
                    break

    # Remove all the lines with recognized names.
    for line in remove_line:
        lines.remove(line)

    print("\n"+"-"*20 + "\n")

    # Print out all the Absentees that remain in the database.
    if student_data:
        print("Absentees:", *
              [f"{i} : {student_data[i].strip()}" for i in student_data], sep="\n ")
    else:
        print("No Absentees.")

    print("\n"+"-"*20 + "\n")

    # Print out all the unrecognized names.
    print("Unrecognized Students:", *[i.strip()
                                      for i in list(filter(lambda n: n, lines))], sep="\n ")

    print("\n"+"-"*20 + "\n")

    # Ask user if Absentees should be logged in Attendance file.
    subject = input("Log under subject>")
    if subject:

        file_name = f"Attendances\\Attendance {datetime.now().date()}.txt"
        # Log the Absentees in Attendance{date}.txt under the Attendances directory
        if not os.path.isfile(file_name):
            # If file does not exist then add heading
            with open(file_name, "w") as f:
                print(f"Absentees - {datetime.now().date()}", file=f)

        with open(file_name, "a") as f:
            # Get a list of Absentees.
            absentees = list(student_data.values())
            if len(absentees) == 0:
                # No Absentees.
                print(f"{subject} : No Absentees", file=f)
            elif len(absentees) == 1:
                # Just 1 Absentee.
                print(f"{subject} : {absentees[0].strip()}", file=f)
            else:
                # Multiple Absentees.
                print(
                    f"{subject} : {', '.join([i.strip() for i in absentees[:-1]])}, and {absentees[-1].strip()}", file=f)

    print("\n"+"-"*20 + "\n")

    input("<Done>\nPress Enter to Quit.")
    print("\n"+"-"*20 + "\n")


if __name__ == "__main__":
    try:
        main()
    except:
        # If an error is encountered, then store error in a traceback file.
        import traceback
        error_file_name = f"Tracebacks\\Traceback{datetime.now()}".replace(
            ":", ";").replace(".", ",")+".txt"
        print("\n"+"-"*20 + "\n")
        print(
            f"- Something went wrong -\n\nCheck {os.getcwd()}\\{error_file_name} for details.")
        print("\n"+"-"*20 + "\n")
        with open(error_file_name, "w") as error_file:
            print(traceback.format_exc(), file=error_file)
        input("Press Enter to Quit.")
        print("\n"+"-"*20 + "\n")
