"""
attpy - Compares a list of present students with total strength to determine absentees and unrecognized students.
Author: Ayan Datta
"""

import pprint
import time
import pyperclip

# Dictionary containing roll numbers and names.
student_data = {1: "XII B 160340ABEY JACOB JOHN",
                2: " ABISHEK ",
                3: "JUHI",
                4: " AIDEN ",
                5: "XII B 4135 AKHIL LAKKARAJU",
                6: "ANANYA BANERJEE",
                7: " ASHMITA ",
                8: " AYAN ",
                9: "XII B 190254 AYUSH NISTALA",
                10: "ARYAN",
                11: " BHAVANA ",
                12: "XII B 4075 CHEEDELLA RAMCHARAN",
                13: " SURYA ",
                14: "KAVYA",
                15: " ESHWAR ",
                16: " GANESH ",
                17: "HARSHA",
                18: " HARSHIT ",
                19: "KATCHARALA ANANYA",
                20: "VARSHITHA",
                21: " KSHYATI ",
                22: "LAASYA",
                23: " ADITHI ",
                24: " MEGHANA ",
                25: "XII B 3953 MERUGU ADITHYA",
                26: " ABHINAV ",
                27: " NIHAAN ",
                28: "XII B 4225 PARANDKAR HARSH",
                29: " POORNA ",
                30: " PRIYANGA ",
                31: " PRIYATAM ",
                32: "XII B 190016 RISHIVEER YADAV ANGIREKULA",
                33: " PAVAN ",
                34: "XII B 190027 SANKARGAL SUBHAN NASIRA BANU",
                35: " GAYATRI ",
                36: " AASIMA ",
                37: " SHRAVANI ",
                38: "XII B 190314 SIRIVELLA NAGA PARINITHA",
                39: " JISHNU ",
                40: "XII B 190045 VELUGULETI JAANVI",
                41: " KOUSHIK ",
                42: " YASHASWINI "}


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

# Print out all the Absentees that remain in the database.
if student_data:
    print("Absentees:", *
          [f"{i} : {student_data[i].strip()}" for i in student_data], sep="\n ")
else:
    print("No Absentees.")
    
# Print out all the unrecognized names.
print("Unrecognized Students:", *list(filter(lambda n: n, lines)), sep="\n ")

# Ask user if Absentees should be logged in todayAttendance file.
a = input("Log under subject>")
if a:
    # Log the Absentees in todayAttendance.txt
    with open("todayAttendance.txt", "a") as f:
        absentees = list(student_data.values())
        if len(absentees) == 0:
            print(f"{a}:No Absentees", file=f)
        elif len(absentees) == 1:
            print(f"{a}:{absentees[0]}", file=f)
        else:
            print(
                f"{a}:{','.join(absentees[:-1])}, and {absentees[-1]}", file=f)

input("<Done>\nPress Enter to Quit")
