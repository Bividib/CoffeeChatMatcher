KBC Coffee Chat Matcher

What is This?

This project is a small, automated system for creating random coffee chat pairings for a club or organization.

It reads a list of all members from an Excel file, applies a set of matching rules, and generates a new list of pairs. It's smart enough to remember old matches so you never get a repeat pairing.

Features

Avoids Repeats: Never pairs two people who have been paired before.

Cross-Department: Never pairs two people from the same department.

Hierarchy-Friendly: Never pairs an "Executive" with another "Executive" (Associates can be paired with anyone).

Safe & Reversible: Creates a new, versioned history file (e.g., past_matches_1.csv, past_matches_2.csv) every time it runs. If you make a mistake, just delete the newest file to "roll back" the history.

Clean Output: Generates a simple current_matches.txt file with the new pairs for easy reading, copying, and pasting.

Smart: Cleans up data from the Excel file, like extra spaces in names or column headers.

Strict: The script will automatically stop and warn you if it finds duplicate names in your Excel file, preventing bad data.

What You Need

Python 3 (3.6 or newer): The script is written in Python. The run scripts will check if you have it.

A Member List: An Excel (.xlsx) file.

This file must have a sheet named All Exe & Ass.

This sheet must have three columns with the exact (case-sensitive) names:

Full Name

Department

Exec or Assoc

Project Files

create_coffee_pairs.py: The main Python script that contains all the matching logic.

run.sh: The one-click script for macOS and Linux (WSL) users.

run.bat: The one-click script for Windows users.

requirements.txt: An auto-generated file listing the Python libraries needed (pandas, openpyxl).

current_matches.txt: A clean, readable text file showing the most recent pairs. This file is overwritten every time you run the script.

past_matches_X.csv: The history files (e.g., past_matches_1.csv). The script uses these as its "memory." Do not edit these unless you want to roll back history.

README.md: This file.

How to Use (For Non-Technical Users)

Follow the steps for your operating system.

➡️ For Windows Users

Place your Excel member list, create_coffee_pairs.py, and run.bat in the same folder.

Double-click the run.bat file.

A command window will pop up. The first time you run it, it will set up an environment and install the required libraries (this may take a minute).

The script will run, and you will see the results in the window.

When it's done, you will find a new current_matches.txt file in the folder with your pairs!

To use a file in a different folder: You can drag your Excel file onto the run.bat icon. The script will automatically use that file.

➡️ For macOS / Linux (WSL) Users

Place your Excel member list, create_coffee_pairs.py, and run.sh in the same folder.

Open your Terminal.

Change to the project's directory. (The easiest way is to type cd  (with a space) and then drag the folder from Finder/File Explorer into the Terminal window). Press Enter.

The very first time you run this, you must make the script executable. Type this and press Enter:

chmod +x run.sh


Now, just run the script by typing this and pressing Enter:

./run.sh


The script will set up the environment, run the pairing, and you will see the results in the terminal.

You will find a new current_matches.txt file in the folder with your pairs!

To use a file in a different folder: You can pass the path as an argument:

./run.sh /path/to/my/other/file.xlsx
