# Author: Joshua Church
# Purpose: Automate the input process for NO LOCS

from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
import os
import pyautogui
import time

# Upload an Excel file
def upload():

    # Configure your settings here.
    title = "Select your file"
    filetypes = (("Excel Workbook", "*.xlsx"), ("Legacy Excel Worksheets", "*.xls"),          
        ("Excel Macro-enabled Workbook", "*.xlsm"), ("Comma Separated Values", "*.csv"))

    # Set the initial starting directory
    initialdir = os.path.join(os.path.join(os.environ["USERPROFILE"]), "Desktop")

    # Store the file
    global filename
    filename = filedialog.askopenfilename(initialdir=initialdir, title=title, filetypes=filetypes)

    # If the user doesn't click 'Cancel', sort the columns.
    if filename:
        sort_columns()
    else:
        filename = None

# Verify that the Excel file is valid
def valid_excel_file(columns):

    try:
        # Make sure the file has the column 'CDSTYL'
        if not any("CDSTYL" in i.upper() for i in columns):
            message = "You must provide a valid Excel file that contains 'CDSTYL'."
            error(message)
            return False

        # Make sure the file has the column 'CHCASN'
        if not any("CHCASN" in i.upper() for i in columns):
            message = "You must provide a valid Excel file that contains 'CHCASN'."
            error(message)
            return False

    except:
        message = "Invalid Excel file."
        error(message)
        return False

    # Valid file
    return True

# Sort the columns by the 'CDSTYL' column
def sort_columns():
    global filename
    workbook = pd.read_excel(filename)
    columns = list(workbook)

    # If the file is valid, sort and extract data
    if valid_excel_file(columns):
        workbook = workbook.sort_values("CDSTYL")
        extract_data(workbook)


# Extract data from the column
def extract_data(workbook):

    global data
    data = workbook["CHCASN"].values
    messagebox.showinfo("Success", "File has been uploaded and sorted!")


# Move cursor to appropriate position
def move_cursor(x, y):

    # Verify all conditions are met.
    if ready():

        # Move mouse to position & click into program
        pyautogui.moveRel(x, y, duration=pyautogui.MINIMUM_DURATION)
        pyautogui.click()

        # Execute the macro commands
        commands()
        

# Macro commands
def commands():

    global data
 
    counter = 0 

    # Loop and execute all commands
    for datum in data:

        # Pause after each set of 100 
        # entries to give the program
        # time to catch up
        if (counter > 0) and (counter % 100) == 0:
            time.sleep(3)

        # Write values to screen
        pyautogui.typewrite(str(datum))

        # Press 'enter'
        pyautogui.keyDown("enter")
        pyautogui.keyUp("enter")

        # Press '9'
        pyautogui.keyDown("9")
        pyautogui.keyUp("9")

        # Press 'enter'
        pyautogui.keyDown("enter")
        pyautogui.keyUp("enter")

        # Press 'f19'
        pyautogui.keyDown("f12")
        pyautogui.keyUp("f12")

        counter += 1



# Helper function to check if all conditions are met
# before using the program.
def ready():

    global filename
    global data

    # Make sure a file has been uploaded.
    if filename is None or data is None:
        messagebox.showinfo("Error", "Please upload a file.")
        return False

    # All conditions have been met.
    else:
        return True



def how_to_use():

    step1 = "Step 1: Upload an Excel file.\n\n"
    step2 = "Step 2: (OPTIONAL) Set delay between commands.\n\n"
    step3 = "Step 3: Place the macro program really close to the PKMS program.\n\n"
    step4 = "Step 4: Choose the direction the cursor needs to move.\n\n"
    step5 = "Step 5: Click the button with the proper direction.\n\n"

    step_opt_a = "** Do NOT touch the computer during this process.\n\n"
    step_opt_b = "** If you want to stop the process, quickly move the mouse to the top-left corner.\n\n"

    message = step1 + step2 + step3 + step4 + step5
    message += step_opt_a + step_opt_b
    messagebox.showinfo("How to Use", message)


# Popup configuration to set delay between commands
def get_delay():

    # Create popup
    popup = Toplevel()

    # Create variable to store user input
    seconds = StringVar()

    # Set default value to '0'
    seconds.set("0")

    Label(popup, text="\nSet a delay for the program (in seconds)").pack()
    Entry(popup, textvariable=seconds).pack()
    Label(popup, text="\n")

    # Onclick, set the value set by the user.
    Button(popup, text="Submit", command=lambda: set_delay(seconds.get(), popup)).pack()


# Set the delay value
def set_delay(seconds, popup):

    try:
        # Check if value is a string
        seconds = float(seconds)

        # Catch negative values
        if (seconds < 0):
            message = "Value must be greater than or equal to 0"
            error(message)
            return

        # Catch '-' character
        elif ('-' in str(seconds)):
            message = "Cannot use '-' character."
            error(message)
            return

    # If a string, throw error
    except:
        message = "Please enter a valid number."
        error(message)
        return

    # Set the delay
    pyautogui.PAUSE = seconds

    # Alert success and close popup
    messagebox.showinfo("Success", "Delay between commands: " +  str(seconds) + " seconds")

    # Close the popup after success
    popup.destroy()

# Error helper function
def error(message):
    messagebox.showerror("Error", message)


# Store the path and filename of the uploaded file
global filename
filename = None

# Store the extracted data from the column
global data
data = None

# Delay between each pyautogui call
pyautogui.PAUSE = 0

# Quickly move the mouse to the top-left of the screen to stop the program.
pyautogui.FAILSAFE = True

# Graphical User Interface Settings
root = Tk()
root.title("PKMS Automation")
root.geometry("250x250")
root.resizable(False, False)

# How to Use
Button(root, text="How to Use", command=how_to_use).place(x=75, y=10)

# Upload
Button(root, text="Upload", command=upload).place(x=30, y=40)

# Set Delay
Button(root, text="Set Delay", command=get_delay).place(x=120, y=40)

# Move Up
Button(root, text="Move Up", command=lambda: move_cursor(0, -150)).place(x=75, y=90)

# Move Down
Button(root, text="Move Down", command=lambda: move_cursor(0, 150)).place(x=70, y=150)

# Move Left
Button(root, text="Move Left", command=lambda: move_cursor(-110, 0)).place(x=10, y=120)

# Move Right
Button(root, text="Move Right", command=lambda: move_cursor(110, 0)).place(x=140, y=120)

# Close
Button(root, text="Close", command=root.quit).place(x=90, y=210)

# Start GUI
root.mainloop()
