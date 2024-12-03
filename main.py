# Import packages
from os.path import exists
import os
import sys
import sqlite3
import smtplib
import ssl
import random
import json
import tkinter as tk
from tkinter import ttk
import time
import subprocess  # Import subprocess for LaTeX compilation

# Declare the version number
version = "v1.0.0"

# Declare widely used paths and files
fileDirectory = './whiteElephantNumberDistributor/'
dbfile = './whiteElephantNumberDistributor/data.db'
configfile = './whiteElephantNumberDistributor/config.json'

# Check if fileDirectory exists, and create it if it doesn't
if exists(fileDirectory):
    print("Directory found")
else:
    os.mkdir(fileDirectory)

# Check if dbfile exists, and create it if it doesn't
if exists(dbfile):
    print("Database file found")
else:
    print("Database file not found")
    print("Creating new database file")

    open(dbfile, "x")

    con = sqlite3.connect(dbfile)
    cur = con.cursor()
    data = cur.execute("""CREATE TABLE IF NOT EXISTS whiteElephantData (
    name TEXT,
    primaryEmail TEXT,
    number INTEGER
    );""")
    con.commit()
    con.close()

# Check if configfile exists, and create it if it doesn't
if exists(configfile):
    print("Configuration file found")
else:
    print("Configuration file not found")
    print("Creating new configuration file")

    with open(configfile, 'w') as config:
        config.writelines('{"smtpServer": "","smtpPort": "465","smtpPassword": "","fromAddress": ""}')
    print("\nPlease completely fill config.json before running again. Program will exit in 10 seconds.")
    time.sleep(10)
    sys.exit()

# Load config file as json
config = json.load(open(configfile))

# Stop program if any config values are empty
if config['smtpServer'] == "" or config['smtpPort'] == "" or config['smtpPassword'] == "" or config[
    'fromAddress'] == "":
    print("\nPlease completely fill config.json before running again. Program will exit in 10 seconds.")
    time.sleep(10)
    sys.exit()

# Save config as variables
smtpServer = config['smtpServer']
smtpPort = config['smtpPort']
fromAddress = config['fromAddress']
smtpPassword = config['smtpPassword']


# Functions
def clearDB():
    con = sqlite3.connect(dbfile)
    cur = con.cursor()
    cur.execute("DELETE FROM whiteElephantData")
    con.commit()
    con.close()
    print("Database cleared")


def assignNumbers():
    con = sqlite3.connect(dbfile)
    cur = con.cursor()
    try:  # Checks if the data exists
        data = cur.execute("SELECT * FROM whiteElephantData").fetchall()
        name, primaryEmail, number = zip(*data)
    except ValueError:  # Handles exception for no rows in the table
        print("No participants found")
    else:
        # Generate numbers
        assignedNumbers = random.sample(range(len(name)), len(name))
        assignedNumbers = [x + 1 for x in assignedNumbers]

        # Save numbers to data.db
        for i, name in enumerate(name):
            cur.execute("UPDATE whiteElephantData SET number = ? WHERE name = ?;", (assignedNumbers[i], name))

        con.commit()
        con.close()

        print("Numbers assigned")


def emailNumbers():
    con = sqlite3.connect(dbfile)
    cur = con.cursor()
    try:
        data = cur.execute("SELECT * FROM whiteElephantData").fetchall()
        name, primaryEmail, number = zip(*data)  # Fetch numbers from database
    except ValueError:  # Handles exception for no rows in the table
        con.close()
        print("No participants found")
    else:
        try:  # Checks to make sure numbers have been distributed
            orderedList = list(zip(name, number))
            orderedList.sort(key=lambda x: x[1], reverse=False)  # Numerically orders list according to number
            orderedName, orderedNumber = zip(*orderedList)
            con.close()
        except TypeError:
            assignNumbers()  # Assign numbers
            data = cur.execute("SELECT * FROM whiteElephantData").fetchall()
            name, primaryEmail, number = zip(*data)  # Fetch numbers from database
            con.close()
            orderedList = list(zip(name, number))
            orderedList.sort(key=lambda x: x[1], reverse=False)  # Numerically orders list according to number
            orderedName, orderedNumber = zip(*orderedList)
        fullNumberList = """"""
        for i in range(len(name)):
            fullNumberList = fullNumberList + ("\n" + str(orderedNumber[i]) + ": " + orderedName[i])
        print(fullNumberList)  # Prints ordered list to console

        message = """From: noreply@tylerdavis.net\nTo: {email}\nSubject: White Elephant Number\n\n{name}, your number for White Elephant is {number}!"""
        messageFullNumberList = """From: noreply@tylerdavis.net\nTo: {email}\nSubject: Full list of White Elephant Numbers\n\nSee below for the full list of numbers:\n{fullNumberList}"""

        context = ssl.create_default_context()  # Set up SMTP

        with smtplib.SMTP_SSL(smtpServer, smtpPort, context=context) as server:
            server.login(fromAddress, smtpPassword)
            server.sendmail(
                fromAddress,
                primaryEmail[0],
                messageFullNumberList.format(email=primaryEmail[0], fullNumberList=fullNumberList)
                # Sends the full number list
            )
            print("Full Number List sent to " + name[0] + " at " + primaryEmail[0])
            for i in range(len(name)):  # Sends all individual emails
                try:
                    server.sendmail(
                        fromAddress,
                        primaryEmail[i],
                        message.format(name=name[i], email=primaryEmail[i], number=number[i])
                    )
                    print("Email sent to " + name[i] + " at " + primaryEmail[i])
                except Exception as e:
                    print("Email failed to send to " + name[i] + " at " + (primaryEmail[i] if primaryEmail[i] else "") + " with error " + str(e))


def listCurrentParticipants():
    con = sqlite3.connect(dbfile)
    cur = con.cursor()
    try:
        data = cur.execute("SELECT * FROM whiteElephantData").fetchall()
        name, primaryEmail, number = zip(*data)
    except ValueError:  # Handles exception for no rows in the table
        print("No participants found")
    else:
        con.close()
        print("Name and email of all participants")
        for i in range(len(name)):
            print(name[i] + " | " + (primaryEmail[i] if primaryEmail[i] else ""))  # Prints the name and email of all participants in data.db


def openWindow():
    # Remove blur and open main window, decides whether to deiconify the window or to start it for the first time
    if rootWindow.state() == "withdrawn":
        rootWindow.deiconify()
        nameBox.focus()
        print("Main window opened")
    else:
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
        except:
            print("Not a Windows machine, skipping ctypes")
        finally:
            rootWindow.mainloop()
            nameBox.focus()
            print("Main window opened")


def onClosing():
    rootWindow.withdraw()
    print("Main window closed")


def closeWindow(event):
    onClosing()


def submitValues():
    name = nameBox.get()
    primaryEmail = primaryEmailBox.get()

    con = sqlite3.connect(dbfile)
    cur = con.cursor()
    sqliteCommand = "INSERT INTO whiteElephantData (name, primaryEmail) VALUES (?, ?);"
    cur.execute(sqliteCommand, (name, primaryEmail))
    con.commit()
    con.close()

    print("Added " + name + " with email " + primaryEmail + " to the database")

    nameBox.delete(0, tk.END)
    primaryEmailBox.delete(0, tk.END)
    nameBox.focus()


def submitValesReturn(event):
    submitValues()


def deleteParticipant():
    con = sqlite3.connect(dbfile)
    cur = con.cursor()
    while True:
        try:
            data = cur.execute("SELECT * FROM whiteElephantData").fetchall()
            name, primaryEmail, number = zip(*data)
        except ValueError:  # Handles exception for no rows in the table
            print("No participants found")
            break
        else:
            print("Name and email of all participants")
            for i in range(len(name)):
                print(name[i] + " | " + primaryEmail[i])
        deleteName = input("Enter the name of who you want to remove as it appears, or press enter to exit:")
        if deleteName == "":
            break
        else:
            sqliteCommand = "DELETE FROM whiteElephantData WHERE name=?;"
            cur.execute(sqliteCommand, (deleteName,))
            print("\nDeleted " + deleteName + " from the database\n")  # Deletes values matching the inputted name
    con.commit()
    con.close()


def pruneParticipants():
    con = sqlite3.connect(dbfile)
    cur = con.cursor()
    try:
        data = cur.execute("SELECT * FROM whiteElephantData").fetchall()
        name, primaryEmail, number = zip(*data)
    except ValueError:  # Handles exception for no rows in the table
        print("No participants found")
    else:  # Deletes empty rows
        cur.execute("DELETE FROM whiteElephantData WHERE name=''")
        cur.execute("DELETE FROM whiteElephantData WHERE name IS NULL")
        print("Participants pruned")
    con.commit()
    con.close()

def createBeamerPresentation():
    # Connect to the database and fetch all participant data
    con = sqlite3.connect(dbfile)
    cur = con.cursor()

    # Get the ordered participants by number
    participants = cur.execute("SELECT number, name FROM whiteElephantData ORDER BY number").fetchall()
    con.close()

    if len(participants) < 1:
        print("Not enough participants to create a presentation.")
        return

    total_participants = len(participants)

    # LaTeX content for the Beamer presentation with a theme and custom footer
    latex_content = f"""
    \\documentclass{{beamer}}
    \\usepackage[utf8]{{inputenc}}

    % Use a modern theme for the presentation
    \\usetheme{{Madrid}}

    \\title{{White Elephant Participants}}
    \\date{{\\today}}
    
    \\setbeamertemplate{{navigation symbols}}{{}} % Remove navigation symbols

    \\begin{{document}}

    % Title Page
    \\begin{{frame}}[plain]
      \\titlepage
    \\end{{frame}}
    """

    # Generate a slide for each participant
    for i in range(len(participants)):
        current_person = participants[i][1]  # Access the name of the current person
        next_person = participants[(i + 1) % len(participants)][1]  # Wrap around to the first participant

        slide_content = f"""
        \\begin{{frame}}[plain]
          \\begin{{block}}{{Turn Information}}
              \\vspace{{1cm}}
              \\centering
              \\Huge{{\\textbf{{Current Person: {current_person}}}}} \\par
              \\vspace{{1cm}}
              \\Large{{Next Person: {next_person}}}
          \\end{{block}}
        \\end{{frame}}
        """
        latex_content += slide_content

    # Final slide: Only the first person, with consistent block formatting
    first_person = participants[0][1]
    final_slide = f"""
    \\begin{{frame}}[plain]
      \\begin{{block}}{{Turn Information}}
          \\vspace{{1cm}}
          \\centering
          \\Huge{{\\textbf{{Return to Start}}}} \\\\
          \\vspace{{0.5cm}} 
          \\Huge{{\\textbf{{First Person to Go Again:}}}} \\\\
          \\vspace{{0.5cm}}
          \\Huge{{\\textbf{{{first_person}}}}}
      \\end{{block}}
    \\end{{frame}}
    """

    latex_content += final_slide

    latex_content += "\\end{document}"
    # Write LaTeX content to a file
    with open('presentation.tex', 'w') as f:
        f.write(latex_content)

    # Compile the LaTeX file to PDF using pdflatex
    try:
        subprocess.run(['pdflatex', '-interaction=nonstopmode', 'presentation.tex'], check=True)
        print("Presentation created successfully.")
    except subprocess.CalledProcessError as e:
        print(f"An error occurred while creating the presentation: {e}")
        return

    # Remove auxiliary LaTeX-generated files, keeping only the PDF
    aux_files = [
        'presentation.aux', 'presentation.log', 'presentation.nav',
        'presentation.out', 'presentation.snm', 'presentation.toc',
        'presentation.tex'
    ]
    for file in aux_files:
        if os.path.exists(file):
            os.remove(file)

# Main window definition
rootWindow = tk.Tk()

rootWindow.attributes('-fullscreen', True)
rootWindow.protocol("WM_DELETE_WINDOW", onClosing)
rootWindow.bind('<Escape>', closeWindow)
rootWindow.bind('<Return>', submitValesReturn)

window_width = 400
window_height = 300

# Get the screen dimension
screen_width = rootWindow.winfo_screenwidth()
screen_height = rootWindow.winfo_screenheight()

# Find the center point
center_x = int(screen_width / 2 - window_width / 2)
center_y = int(screen_height / 2 - window_height / 2)

rootWindow.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

name = tk.StringVar()
primaryEmail = tk.StringVar()

style = ttk.Style()
style.configure('TButton', font=("Times New Roman", 20))
style.configure('TEntry', font=("Times New Roman", 20))
style.configure('TLabel', font=("Times New Roman", 20))

rootWindow.title("White Elephant Number Distributor")

label = ttk.Label(rootWindow, text="Welcome to the White Elephant Game!")
label.pack()

nameLabel = ttk.Label(rootWindow, text="Please enter your name below")
nameLabel.pack(padx=10, pady=10, fill='x', expand=False)
nameBox = ttk.Entry(rootWindow, textvariable=name)
nameBox.pack(padx=10, pady=10, fill='x', expand=False)
nameBox.focus()

primaryEmailLabel = ttk.Label(rootWindow, text="Please enter your email below")
primaryEmailLabel.pack(padx=10, pady=10, fill='x', expand=False)
primaryEmailBox = ttk.Entry(rootWindow, textvariable=primaryEmail)
primaryEmailBox.pack(padx=10, pady=10, fill='x', expand=False)

submitButton = ttk.Button(
    rootWindow,
    text="Submit",
    command=submitValues
)
submitButton.pack(padx=10, pady=10, fill='x', expand=False)

rootWindow.withdraw()


def printMenu():
    menu = {}
    menu['1'] = "Gather new participants"
    menu['2'] = "List current participants"
    menu['3'] = "Assign numbers"
    menu['4'] = "Email numbers"
    menu['5'] = "Delete participant"
    menu['6'] = "Prune empty entries"
    menu['7'] = "Clear DB"
    menu['8'] = "Create Beamer Presentation"
    menu['0'] = "Print menu to console"
    menu['-'] = "Exit"

    options = menu.keys()
    for entry in options:
        print(entry, menu[entry])


print()
print(f"You are using whiteElephantNumberDistributor version {version}")
print()
print()
printMenu()

while True:
    if rootWindow.state() == "normal":
        input("Please press enter when the data collection window is closed and you are ready to continue\n\n")
    else:
        selection = input("Please select an option: ")
        if selection == '1':
            print()
            openWindow()
            print()
        elif selection == '2':
            print()
            listCurrentParticipants()
            print()
        elif selection == '3':
            print()
            assignNumbers()
            print()
        elif selection == '4':
            print()
            emailNumbers()
            print()
        elif selection == '5':
            print()
            deleteParticipant()
            print()
        elif selection == '6':
            print()
            pruneParticipants()
            print()
        elif selection == '7':
            print()
            clearDB()
            print()
        elif selection == '8':
            print()
            createBeamerPresentation()
            print()
        elif selection == '0':
            print()
            printMenu()
        elif selection == '-':
            break
        else:
            print("\nInvalid selection\n")
