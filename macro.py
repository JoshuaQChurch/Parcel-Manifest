# Author: Joshua Church
# Purpose: Tool that will allow the user to manifest open cartons OR reprint cartons for Williams Sonoma. 

from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
import os
import pyautogui
import time

# Upload an Excel file
def upload():
    filename = None
    data = None

    # Configure your settings here.
    title = "Select your file"
    filetypes = (("Excel Workbook", "*.xlsx"), ("Legacy Excel Worksheets", "*.xls"),          
        ("Excel Macro-enabled Workbook", "*.xlsm"), ("Comma Separated Values", "*.csv"))

    # Set the initial starting directory
    initialdir = os.path.join(os.path.join(os.environ["USERPROFILE"]), "Desktop")

    filename = filedialog.askopenfilename(initialdir=initialdir, title=title, filetypes=filetypes)
    
    if filename:
        # Handler for csv files
        if filename.endswith(".csv"):
            data = pd.read_csv(filename)
            option_select(data)

        # Handler for Excel-specific files
        elif filename.rsplit(".")[1] in ["xlsx", "xls", "xlxsm"]:
            data = pd.read_excel(filename)
            option_select(data)

def option_select(data):

    # Remove previous widgets
    for widget in root.winfo_children():
        widget.destroy()

    # Change size of frame
    root.geometry("320x100")
    Label(root, text="Select Your Choice", font=("Arial", 15, "bold")).place(x=60, y=20)
    Label(root, text="OR", font=("Arial", 15)).place(x=110,y=55)

    # Reprint Selection
    Button(root, text="Reprint", command=lambda: reprint(data), font=("Arial", 15)).place(x=20, y=50)
    
    # Carton Manifest Selection
    Button(root, text="Carton Manifest", command=lambda: carton_manifest(data), font=("Arial", 15)).place(x=155, y=50)


def reprint(data):

    columns = data.columns.values.tolist()
    
    if "CDSTYL" not in columns:
        error("You must provide a valid Excel file that contains 'CDSTYL'. Exiting program now.")

    if "CHCASN" not in columns:
        error("You must provide a valid Excel file that contains 'CHCASN'. Exiting program now.")
       
    workbook = data.sort_values("CDSTYL")
    data = workbook["CHCASN"].values
    messagebox.showinfo("Success", "File has been uploaded and sorted!")
    set_position(1, data)   

def set_position(macro, data):

    # Remove previous widgets
    for widget in root.winfo_children():
        widget.destroy()

    root.geometry("200x120")
    root.resizable(False, False)

    # Move Up
    Button(root, text="Move Up", command=lambda: move_cursor(0, -150, macro, data)).place(x=70, y=20)

    # Move Down
    Button(root, text="Move Down", command=lambda: move_cursor(0, 150, macro, data)).place(x=60, y=80)

    # Move Left
    Button(root, text="Move Left", command=lambda: move_cursor(-150, 0, macro, data)).place(x=25, y=50)

    # Move Right
    Button(root, text="Move Right", command=lambda: move_cursor(150, 0, macro, data)).place(x=110, y=50)


# Move cursor to appropriate position
def move_cursor(x, y, macro, data):

    # Move mouse to position & click into program
    pyautogui.moveRel(x, y, duration=pyautogui.MINIMUM_DURATION)
    pyautogui.click()
     
    if macro == 1:
        reprint_macro(data)
    
    else:
        carton_manifest_macro(data)
        
        
# Macro commands
def reprint_macro(data):

    # Loop and execute all commands
    counter = 0 
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

    messagebox.showinfo("Complete", "The macro is finished. The program will now close.")
    root.quit()


def carton_manifest(data):
    # Remove previous widgets
    for widget in root.winfo_children():
        widget.destroy()

    root.geometry("170x120")

    Label(root, text="Does this file have\ncolumn headers?", font=("Arial", 12, "bold")).place(x=15, y=5)
    Button(root, text="Yes", command=lambda: carton_manifest_with_headers(data), font=("Arial", 15)).place(x=20, y=55)
    Button(root, text="No", command=lambda: carton_manifest_without_headers(data), font=("Arial", 15)).place(x=100, y=55)


def carton_manifest_with_headers(data):
    # Remove previous widgets
    for widget in root.winfo_children():
        widget.destroy()

    root.geometry("225x100")

    columns = data.columns
    options = [i for i in columns]
    variable = StringVar(root)
    variable.set(options[0])

    menu = OptionMenu(root, variable, *options)
    Label(root, text="Select the column header", font=("Arial", 12, "bold")).place(x=10, y=10)
    menu.place(x=20, y=40)
    Button(root, text="Submit", command=lambda: set_position(2, data[variable.get()].values.tolist())).place(x=120, y=42)


def carton_manifest_without_headers(data):
    # Remove previous widgets
    for widget in root.winfo_children():
        widget.destroy()

    root.geometry("200x100")

    columns = data.columns
    options = [str(i+1) for i in range(len(columns))]
    variable = StringVar(root)
    variable.set(options[0])

    menu = OptionMenu(root, variable, *options)
    Label(root, text="Select the column number", font=("Arial", 10, "bold")).place(x=10, y=10)
    menu.place(x=20, y=40)
    Button(root, text="Submit", command=lambda: set_position(2, data.iloc[:,int(variable.get())-1].values.tolist())).place(x=100, y=42)

def carton_manifest_macro(data):

    # Loop and execute all commands
    counter = 0 
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

        counter += 1
    
    messagebox.showinfo("Complete", "The macro is finished. The program will now close.")
    root.quit()


# Error helper function
def error(message):
    messagebox.showerror("Error", message)
    root.quit()
    exit()

    
if __name__ == "__main__": 

    # Delay between each pyautogui call
    pyautogui.PAUSE = 0

    # Quickly move the mouse to the top-left of the screen to stop the program.
    pyautogui.FAILSAFE = True

    # Create the GUI
    root = Tk()
    root.title("PKMS Automation")
    root.geometry("300x300")
    root.resizable(True, True)

    btn = Button(root, text="Upload File", command=lambda: upload())
    btn.config(height=50, width=50, font=("Arial", 18, "bold"))
    btn.pack()

    root.mainloop()
