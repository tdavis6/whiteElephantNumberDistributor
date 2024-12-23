import csv
import os
import random
import smtplib
import ssl
import sys
import time
import json
import subprocess  # For LaTeX compilation

# -------------------
# Global Config
# -------------------
version = "v1.0.0"

fileDirectory = './whiteElephantNumberDistributor/'
csvfile = os.path.join(fileDirectory, 'data.csv')
configfile = os.path.join(fileDirectory, 'config.json')

# Create required directories/files if needed
if not os.path.exists(fileDirectory):
    os.mkdir(fileDirectory)
    print(f"Created directory: {fileDirectory}")
else:
    print("Directory found")

if not os.path.exists(csvfile):
    print("CSV file not found, creating a new one with headers [name, email, number].")
    with open(csvfile, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f, quoting=csv.QUOTE_NONE, escapechar='\\')
        writer.writerow(["name", "email", "number"])  # Initialize headers
else:
    print("CSV file found")

if not os.path.exists(configfile):
    print("Configuration file not found, creating new configuration file...")
    with open(configfile, 'w') as config:
        config.write(json.dumps({
            "smtpServer": "",
            "smtpPort": "465",
            "smtpPassword": "",
            "fromAddress": ""
        }, indent=2))
    print("\nPlease fill in the config.json before running again. Program will exit in 10 seconds.")
    time.sleep(10)
    sys.exit()

# Load config and verify
with open(configfile, 'r') as cfg:
    config = json.load(cfg)

if (not config.get('smtpServer') or not config.get('smtpPort')
        or not config.get('smtpPassword') or not config.get('fromAddress')):
    print("\nPlease completely fill config.json before running again. Program will exit in 10 seconds.")
    time.sleep(10)
    sys.exit()

smtpServer = config['smtpServer']
smtpPort = config['smtpPort']
smtpPassword = config['smtpPassword']
fromAddress = config['fromAddress']  # actual email from config


# -------------------
# Helper Functions
# -------------------

def read_csv_data() -> list:
    """
    Reads and returns CSV data as a list of dicts:
    [{'name': ..., 'email': ..., 'number': ...}, ...].

    Also removes any leftover quote characters within each field.
    """
    rows = []
    with open(csvfile, mode='r', newline='', encoding='utf-8') as f:
        # Use QUOTE_MINIMAL or none. If your CSV is in messy format,
        # you might need to experiment with how it is read.
        reader = csv.DictReader(f, escapechar='\\', quoting=csv.QUOTE_MINIMAL)
        for row in reader:
            # Remove all quote characters from each field just in case
            for key in row:
                if row[key]:
                    # Strip both leading/trailing quotes and replace any internal quotes
                    row[key] = row[key].replace('"', '').strip()
            rows.append(row)
    return rows


def write_csv_data(data: list):
    """
    Writes the given list of dicts back to the CSV file
    with columns name, email, number, without adding quotes.
    """
    with open(csvfile, mode='w', newline='', encoding='utf-8') as f:
        fieldnames = ["name", "email", "number"]
        writer = csv.DictWriter(f, fieldnames=fieldnames, quoting=csv.QUOTE_NONE, escapechar='\\')
        writer.writeheader()
        writer.writerows(data)


def clearCSV():
    """
    Clears the CSV file except for the header row.
    """
    print("Clearing CSV data...")
    with open(csvfile, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f, quoting=csv.QUOTE_NONE, escapechar='\\')
        writer.writerow(["name", "email", "number"])
    print("CSV cleared.")


def listCurrentParticipants():
    data = read_csv_data()
    if not data:
        print("No participants found.")
    else:
        print("\nCurrent participants:")
        for row in data:
            name = row.get('name', '')
            email = row.get('email', '')
            print(f"  {name} | {email}")
    print()


def gatherNewParticipants():
    """
    Repeatedly ask for name and email, store them in CSV.
    Stop if name is empty.
    """
    data = read_csv_data()
    while True:
        name = input("Enter participant name (or press Enter to finish): ").strip()
        if not name:
            break
        email = input("Enter participant email: ").strip()
        data.append({"name": name, "email": email, "number": ""})
        print(f"Added {name} with email {email}\n")
    write_csv_data(data)


def assignNumbers():
    """
    Assign a random unique number to each participant and save back to CSV.
    """
    data = read_csv_data()
    if not data:
        print("No participants to assign numbers to.")
        return

    indices = list(range(1, len(data) + 1))
    random.shuffle(indices)

    for i, row in enumerate(data):
        row['number'] = str(indices[i])

    write_csv_data(data)
    print("Numbers assigned successfully.\n")


def emailNumbers():
    """
    Email participants their assigned White Elephant numbers,
    plus send the full list of participants (by number) to the first participant.
    """
    data = read_csv_data()
    if not data:
        print("No participants found in CSV.\n")
        return

    # Ensure we have assigned numbers; if any row has empty 'number', assign them
    need_assign = any(not row.get('number') for row in data)
    if need_assign:
        print("Some participants have no numbers. Assigning numbers first...")
        assignNumbers()
        data = read_csv_data()

    # Sort by number ascending
    try:
        data.sort(key=lambda x: int(x.get('number', 0)))
    except ValueError:
        print("Invalid number in CSV; reassigning numbers...\n")
        assignNumbers()
        data = read_csv_data()
        data.sort(key=lambda x: int(x.get('number', 0)))

    fullNumberList = "\n".join(f"{row['number']}: {row['name']}" for row in data)

    # Remove extra quotes around the "From" field
    from_field = f"White Elephant Numbers <{fromAddress}>"

    message_individual = """From: {from_field}
To: {email}
Subject: White Elephant Number

{p_name}, your number for White Elephant is {p_num}!
"""

    message_full_list = """From: {from_field}
To: {email}
Subject: Full list of White Elephant Numbers

See below for the full list of numbers:
{list_content}
"""

    context = ssl.create_default_context()
    try:
        with smtplib.SMTP_SSL(smtpServer, smtpPort, context=context) as server:
            server.login(fromAddress, smtpPassword)

            # Send full list to the first participant in the sorted data
            first_email = data[0]['email']
            first_name = data[0]['name']
            if first_email:
                server.sendmail(
                    fromAddress,
                    first_email,
                    message_full_list.format(
                        from_field=from_field,
                        email=first_email,
                        list_content=fullNumberList
                    )
                )
                print(f"Full Number List sent to {first_name} at {first_email}")

            # Send individual emails to everyone
            for row in data:
                name = row['name']
                email = row['email']
                number = row['number']
                try:
                    server.sendmail(
                        fromAddress,
                        email,
                        message_individual.format(
                            from_field=from_field,
                            p_name=name,
                            email=email,
                            p_num=number
                        )
                    )
                    print(f"Email sent to {name} at {email}")
                except Exception as e:
                    print(f"Failed to send email to {name} at {email}: {e}")
        print()
    except Exception as e:
        print(f"SMTP error occurred: {e}\n")


def deleteParticipant():
    data = read_csv_data()
    if not data:
        print("No participants found.\n")
        return

    listCurrentParticipants()
    while True:
        deleteName = input("Enter the exact name to remove (or press Enter to cancel): ").strip()
        if not deleteName:
            break
        new_data = [row for row in data if row['name'] != deleteName]
        if len(new_data) == len(data):
            print(f"No participant found with the name '{deleteName}'.")
        else:
            print(f"Deleted '{deleteName}' from the list.")
            data = new_data
    write_csv_data(data)
    print()


def pruneParticipants():
    """
    Removes entries from the CSV that have empty or null name fields.
    """
    data = read_csv_data()
    if not data:
        print("No participants found.\n")
        return

    before_count = len(data)
    data = [row for row in data if row.get('name', '').strip()]
    after_count = len(data)
    write_csv_data(data)
    print(f"Pruned {before_count - after_count} empty entries.\n")


def createBeamerPresentation():
    """
    Creates a LaTeX Beamer presentation showing the turn order.
    """
    data = read_csv_data()
    if not data:
        print("Not enough participants to create a presentation.\n")
        return

    # Ensure numbers are assigned
    try:
        data.sort(key=lambda x: int(x.get('number', 0)))
    except ValueError:
        print("Invalid 'number' field; reassigning.\n")
        assignNumbers()
        data = read_csv_data()
        data.sort(key=lambda x: int(x.get('number', 0)))

    if len(data) < 1:
        print("Not enough participants to create a presentation.\n")
        return

    # Build LaTeX content
    latex_content = r"""
\documentclass{beamer}
\usepackage[utf8]{inputenc}
\usetheme{Madrid}
\title{White Elephant Participants}
\date{\today}
\setbeamertemplate{navigation symbols}{}

\begin{document}

\begin{frame}[plain]
  \titlepage
\end{frame}
"""

    for i in range(len(data)):
        current_person = data[i]['name']
        next_person = data[(i + 1) % len(data)]['name']  # wrap around
        slide = f"""
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
        latex_content += slide

    first_person = data[0]['name']
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
    latex_content += "\n\\end{document}\n"

    with open("presentation.tex", "w", encoding="utf-8") as f:
        f.write(latex_content)

    try:
        subprocess.run(["pdflatex", "-interaction=nonstopmode", "presentation.tex"], check=True)
        print("Presentation created successfully.\n")
    except subprocess.CalledProcessError as e:
        print(f"An error occurred while creating the presentation: {e}\n")
        return

    aux_files = [
        'presentation.aux', 'presentation.log', 'presentation.nav',
        'presentation.out', 'presentation.snm', 'presentation.toc',
        'presentation.tex'
    ]
    for file in aux_files:
        if os.path.exists(file):
            os.remove(file)


# -------------------
# Menu
# -------------------
def printMenu():
    menu = {
        '1': "Gather new participants",
        '2': "List current participants",
        '3': "Assign numbers",
        '4': "Email numbers",
        '5': "Delete participant",
        '6': "Prune empty entries",
        '7': "Clear CSV",
        '8': "Create Beamer Presentation",
        '0': "Print menu again",
        '-': "Exit"
    }
    print("\n--- White Elephant Number Distributor ---")
    for k, v in menu.items():
        print(k, v)
    print()


# -------------------
# Main Loop
# -------------------
if __name__ == "__main__":
    print(f"\nYou are using White Elephant Number Distributor version {version}\n")
    printMenu()

    while True:
        selection = input("Please select an option: ").strip()
        if selection == '1':
            gatherNewParticipants()
        elif selection == '2':
            listCurrentParticipants()
        elif selection == '3':
            assignNumbers()
        elif selection == '4':
            emailNumbers()
        elif selection == '5':
            deleteParticipant()
        elif selection == '6':
            pruneParticipants()
        elif selection == '7':
            clearCSV()
        elif selection == '8':
            createBeamerPresentation()
        elif selection == '0':
            printMenu()
        elif selection == '-':
            print("Goodbye!")
            break
        else:
            print("Invalid selection. Please try again.\n")
