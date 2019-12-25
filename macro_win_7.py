import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd 
import os 
import pyautogui
import time 
from pathlib import Path 
import copy


class MainApplication:
    """ 
    This object creates a tkinter graphical
    user interface that allows the user to
    automate the following processes:
    
    1: Reprint 
    2. Carton Manifest 
    3. Salesman Code Maintenance 
    """

    def __init__(self, root):
        self.root = root
        self.root.title("PKMS Automation Tool")
        
        # Start at main menu.
        self.__main_menu()


    def __main_menu(self):
        """ 
        This is the application main menu. When the user starts OR
        returns to the main menu, all default values are
        reset.

        Here, the user selects either the Reprint or 
        Carton Manifest Macro.
        """

        # Data extracted from file.
        self.data = None 

        # User-selected macro choice.
        self.macro = None

        # Reactive variable to display selected file.
        self.uploaded_file = tk.StringVar()
        self.uploaded_file.set("None")

        # Number of carton entries to enter.
        self.carton_entry_count = 30

        # Range selection for data. 
        self.start_index = 0
        self.end_index = 0
        self.previous_start_index = 0
        self.previous_end_index = 0

        # Display remaining carton count to user.
        self.remaining_cartons = 0

        # Keep up with the previous page the user was on.
        # Used for carton manifest "Next" and "Retry" functionality.
        self.previous_page = None

        self.__set_geometry(width = 300, height = 300)
        tk.Label(self.root, text = "Select Macro", font = ("Arial", 16, "bold")).place(x = 80, y = 10)

        # Reprint Selection
        reprint_button = tk.Button(self.root, text = "Reprint", command = lambda: self.__config_reprint())
        reprint_button.config(height = 4, width = 9, font = ("Arial", 16))
        reprint_button.place(x = 20, y = 60)
        
        # Carton Manifest Selection
        carton_manifest_button = tk.Button(self.root, text = "Carton\nManifest", command = lambda: self.__config_carton_manifest())
        carton_manifest_button.config(height = 4, width = 9, font = ("Arial", 16))
        carton_manifest_button.place(x = 155, y = 60)

        # Salesman Code Maintenance
        salesman_code_button = tk.Button(self.root, text = "Salesman Code\n Maintenance", command = lambda: self.__config_salesman_code())
        salesman_code_button.config(height = 3, width = 20, font = ("Arial", 16))
        salesman_code_button.place(x = 25, y = 190)


    """ Clear the window frame, and set the frame dimensions """
    def __set_geometry(self, width, height, resizable=False):
        for widget in self.root.winfo_children():
            widget.destroy()

        self.root.geometry("{}x{}".format(width, height))
        self.root.resizable(resizable, resizable)


    """ Verify user wants to return to main menu """
    def __confirm_main_menu(self):
        choice = messagebox.askokcancel("Confirm", "Are you sure you want to return to the main menu?")
        if choice:
            self.__main_menu()  


    """ Configuration settings for the reprint macro """
    def __config_reprint(self):

        self.macro = "reprint"
        self.__set_geometry(width = 400, height = 380, resizable = True)

        label = tk.Label(self.root, text = "Reprint Macro Instructions")
        label.config(font = ("Arial", 18, "bold"))
        label.place(x = 60, y = 10)

        label = tk.Label(self.root, text = '-'*65)
        label.config(font = ("Arial", 16))
        label.place(x = 0, y = 45)

        label = tk.Label(self.root, text = "1. Click the \"Select File\" button.")
        label.config(font = ("Arial", 12))
        label.place(x = 25, y = 70)

        label = tk.Label(self.root, text = "2. Select data file.")
        label.config(font = ("Arial", 12))
        label.place(x = 25, y = 100)

        label = tk.Label(self.root, text = "a. This file must contain the following columns: ")
        label.config(font = ("Arial", 12))
        label.place(x = 40, y = 130)

        label = tk.Label(self.root, text = "i. CDSTYL")
        label.config(font = ("Arial", 12))

        label.place(x = 55, y = 160)

        label = tk.Label(self.root, text = "ii. CHCASN")
        label.config(font = ("Arial", 12))
        label.place(x = 55, y = 190)

        label = tk.Label(self.root, text = "3. Click the \"Submit\" button.")
        label.config(font = ("Arial", 12))
        label.place(x = 25, y = 220)

        label = tk.Label(self.root, text = "Selected File: ")
        label.config(font = ("Arial", 12, "bold"))
        label.place(x = 25, y = 270)

        label = tk.Label(self.root, textvariable = self.uploaded_file)
        label.config(font = ("Arial", 12, "italic"))
        label.place(x = 140, y = 270)

        button = tk.Button(self.root, text = "Select\nFile", command = lambda: self.__upload())
        button.config(height = 3, width = 10, font = ("Arial", 12, "bold"))
        button.place(x = 20, y = 300)

        button = tk.Button(self.root, text = "Submit", command = lambda: self.__verify_reprint())
        button.config(height = 3, width = 10, font = ("Arial", 12, "bold"))
        button.place(x = 145, y = 300)

        button = tk.Button(self.root, text = "Return to\nMain Menu", command = lambda: self.__confirm_main_menu())
        button.config(height = 3, width = 10, font = ("Arial", 12, "bold"))
        button.place(x = 270, y = 300)


    """ Verify the requirements for the reprint macro. """ 
    def __verify_reprint(self):

        if self.data is None:
            self.__error(message = "Please select a file to proceed.")
            return()

        columns = self.data.columns.values.tolist()
        
        if "CDSTYL" not in columns or "CHCASN" not in columns:
            self.__error("You must provide a valid file that contains the following columns: CDSTYL and CHCASN")
            return()
        
        workbook = self.data.sort_values("CDSTYL")
        self.data = workbook["CHCASN"].values
        self.__set_position_instructions()  


    """ Upload the data file. """
    def __upload(self):

        title = "Select your file."
        filetypes = (("Legacy Excel Worksheets", "*.xls"), ("Excel Workbook", "*.xlsx"),          
            ("Excel Macro-enabled Workbook", "*.xlsm"), ("Comma Separated Values", "*.csv"))

        # Set the initial starting directory
        home = str(Path.home())
        directories = os.listdir(home)

        if "Desktop" in directories: 
            initialdir = os.path.join(home, "Desktop")

        elif "Downloads" in directories:
            initialdir = os.path.join(home, "Downloads")
        
        else:
            initialdir = home 
    
        filename = filedialog.askopenfilename(initialdir=initialdir, title=title, filetypes=filetypes)
        
        if filename:
            extension = filename.split('.')[-1].lower()

            # Handler for csv files
            if extension == "csv":
                self.data = pd.read_csv(filename)
                self.remaining_cartons = len(self.data)
                self.uploaded_file.set(os.path.basename(filename))

            # Handler for Excel-specific files
            elif extension in ["xlsx", "xls", "xlsm"]:
                self.data = pd.read_excel(filename)
                self.remaining_cartons = len(self.data)
                self.uploaded_file.set(os.path.basename(filename))

            else:
                self.__error(message = "Please select a valid file.")


    """ Show users how to set the macro position. """
    def __set_position_instructions(self):

        self.__set_geometry(width = 650, height = 350, resizable = True)

        label = tk.Label(self.root, text = "Macro Placement Instructions")
        label.config(font = ("Arial", 18, "bold"))
        label.place(x = 160, y = 10)

        label = tk.Label(self.root, text = '-'*150)
        label.config(font = ("Arial", 16))
        label.place(x = 0, y = 40)

        label = tk.Label(self.root, text = "1. Click the \"Set Position\" button.")
        label.config(font = ("Arial", 12))
        label.place(x = 25, y = 70)

        label = tk.Label(self.root, text = "2. Move this application next to the PKMS system.")
        label.config(font = ("Arial", 12))
        label.place(x = 25, y = 100)

        label = tk.Label(self.root, text = "a. Essentially, place this on the PKMS system.")
        label.config(font = ("Arial", 12))
        label.place(x = 55, y = 130)

        label = tk.Label(self.root, text = "3. Click the direction where the mouse cursor should go to click into the PKMS system.")
        label.config(font = ("Arial", 12))
        label.place(x = 25, y = 160)

        label = tk.Label(self.root, text = "4. Wait until the application finishes or another prompt displays a new set of instructions.")
        label.config(font = ("Arial", 12))
        label.place(x = 25, y = 190)

        label = tk.Label(self.root, text = "** NOTE: To stop the macro, quickly move the mouse cursor to the top left corner of screen.")
        label.config(font = ("Arial", 10, "bold"), fg = "red")
        label.place(x = 25, y = 230)

        button = tk.Button(self.root, text = "Set Positition", command = lambda: self.__set_position())
        button.config(height = 3, width = 25, font = ("Arial", 12, "bold"))
        button.place(x = 30, y = 270)

        button = tk.Button(self.root, text = "Return to\nMain Menu", command = lambda: self.__confirm_main_menu())
        button.config(height = 3, width = 25, font = ("Arial", 12, "bold"))
        button.place(x = 330, y = 270)


    """ Set the mouse cursor position """
    def __set_position(self):

        self.__set_geometry(350, 350)

        label = tk.Label(self.root, text = "Set Position and Mouse Direction")
        label.config(font = ("Arial", 14, "bold"))
        label.place(x = 25, y = 10)

        label = tk.Label(self.root, text = '-'*100)
        label.config(font = ("Arial", 14))
        label.place(x = 0, y = 40)

        # Move Up
        button = tk.Button(self.root, text = "Move Up", command = lambda: self.__move_cursor(0, -200))
        button.config(height = 2, width = 9, font = ("Arial", 12, "bold"))
        button.place(x = 125, y = 75)

        # Move Down
        button = tk.Button(self.root, text = "Move Down", command = lambda: self.__move_cursor(0, 250))
        button.config(height = 2, width = 9, font = ("Arial", 12, "bold"))
        button.place(x = 125, y = 150)

        # Move Left
        button = tk.Button(self.root, text = "Move Left", command = lambda: self.__move_cursor(-175, 0))
        button.config(height = 2, width = 9, font = ("Arial", 12, "bold"))
        button.place(x = 15, y = 110)

        # Move Right
        button = tk.Button(self.root, text = "Move Right", command = lambda: self.__move_cursor(175, 0))
        button.config(height = 2, width = 9, font = ("Arial", 12, "bold"))
        button.place(x = 235, y = 110)

        if self.macro == "reprint" or self.macro == "salesman_code":

            button = tk.Button(self.root, text = "Return to\nMain Menu", command = lambda: self.__confirm_main_menu())
            button.config(height = 3, width = 20, font=("Arial", 12, "bold"))
            button.place(x = 75, y = 260)

        else:
            button = tk.Button(self.root, text = "Return to\nMain Menu", command = lambda: self.__confirm_main_menu())
            button.config(height = 3, width = 12, font=("Arial", 12, "bold"))
            button.place(x = 25, y = 260)

            button = tk.Button(self.root, text = "Previous Page", command = lambda: self.__previous_page())
            button.config(height = 3, width = 14, font=("Arial", 12, "bold"))
            button.place(x = 170, y = 260)


    """ Move cursor to appropriate position """
    def __move_cursor(self, x, y):

        # Move mouse to position & click into program
        pyautogui.moveRel(x, y, duration=pyautogui.MINIMUM_DURATION)
        pyautogui.click()

        if self.macro == "reprint":
            self.__reprint_macro()

        elif self.macro == "carton_manifest":
            self.__carton_manifest_macro()

        elif self.macro == "salesman_code":
            self.__salesman_code_macro()


    """ Allow user to go back to previous page. """ 
    def __previous_page(self):

        if self.previous_page == "__set_carton_count":
            self.remaining_cartons += len(self.data) 
            self.start_index = self.previous_start_index
            self.end_index = self.previous_end_index
            self.__set_carton_count()

        elif self.previous_page == "__remove_problem_cells":
            self.data = copy.deepcopy(self.data_backup)
            self.data = self.data[self.start_index:self.end_index]
            self.__remove_problem_cells()
            
            
    """ Reprint macro commands """
    def __reprint_macro(self):

        # Catch pyautogui.FAILSAFE.
        try:
            for i in range(len(self.data)):

                # Pause after each set of 100 entries to give the program time to catch up
                if (i > 0) and (i % 100) == 0:
                    time.sleep(3)

                # Write values to screen
                pyautogui.typewrite(str(self.data[i]))

                # Press 'enter'
                pyautogui.keyDown("enter")
                pyautogui.keyUp("enter")

                # Press '9'
                pyautogui.keyDown("9")
                pyautogui.keyUp("9")

                # Press 'enter'
                pyautogui.keyDown("enter")
                pyautogui.keyUp("enter")

                # Press 'f12'
                pyautogui.keyDown("f12")
                pyautogui.keyUp("f12")

            messagebox.showinfo("Complete", "The macro is finished. Returning to main menu.")
            self.__main_menu()

        except:
            messagebox.showinfo("Stop", "Failsafe activated. The macro has stopped.")


    """ Configuration settings for the reprint macro """
    def __config_carton_manifest(self):

        self.macro = "carton_manifest"
        self.__set_geometry(width = 400, height = 350, resizable = True)

        label = tk.Label(self.root, text = "Carton Manifest Macro Instructions")
        label.config(font = ("Arial", 14, "bold"))
        label.place(x = 30, y = 10)

        label = tk.Label(self.root, text = '-'*100)
        label.config(font = ("Arial", 16))
        label.place(x = 0, y = 40)

        label = tk.Label(self.root, text = "1. Click the \"Select File\" button.")
        label.config(font = ("Arial", 12))
        label.place(x = 25, y = 70)

        label = tk.Label(self.root, text = "2. Select data file.")
        label.config(font = ("Arial", 12))
        label.place(x = 25, y = 100)

        label = tk.Label(self.root, text = "a. This file must contain the following column: ")
        label.config(font = ("Arial", 12))
        label.place(x = 40, y = 130)

        label = tk.Label(self.root, text = "i. CHCASN")
        label.config(font = ("Arial", 12))
        label.place(x = 55, y = 160)

        label = tk.Label(self.root, text = "3. Click the \"Submit\" button.")
        label.config(font = ("Arial", 12))
        label.place(x = 25, y = 190)

        label = tk.Label(self.root, text = "Selected File: ")
        label.config(font = ("Arial", 12, "bold"))
        label.place(x = 25, y = 240)

        label = tk.Label(self.root, textvariable = self.uploaded_file)
        label.config(font = ("Arial", 12, "italic"))
        label.place(x = 140, y = 240)

        button = tk.Button(self.root, text = "Select\nFile", command = lambda: self.__upload())
        button.config(height = 3, width = 10, font = ("Arial", 12, "bold"))
        button.place(x = 20, y = 270)

        button = tk.Button(self.root, text = "Submit", command = lambda: self.__verify_carton_manifest())
        button.config(height = 3, width = 10, font = ("Arial", 12, "bold"))
        button.place(x = 145, y = 270)

        button = tk.Button(self.root, text = "Return to\nMain Menu", command = lambda: self.__confirm_main_menu())
        button.config(height = 3, width = 10, font = ("Arial", 12, "bold"))
        button.place(x = 270, y = 270)


    """ Verify valid file for carton manifest. """ 
    def __verify_carton_manifest(self):

        if self.data is None:
            self.__error(message = "Please select a file to proceed.")
            return()

        columns = self.data.columns.values.tolist()
        
        if "CHCASN" not in columns:
            self.__error("You must provide a valid file that contains the following column: CHCASN")
            return()

        self.data = self.data["CHCASN"].values.tolist()
        self.data_backup = copy.deepcopy(self.data)
        self.__set_carton_count()


    """ Allow user to set number of carton entries. """ 
    def __set_carton_count(self):

        self.__set_geometry(width = 390, height = 200)

        self.previous_page = "__set_carton_count"

        # Display remaining cartons.
        remaining_cartons = tk.StringVar()

        # Alert user if an error occurs with input data. 
        carton_entry_error = tk.StringVar()

        # Reactive value to get input from user. 
        # Update 'Remaining Cartons' on each character entry. 
        carton_number = tk.StringVar()
        carton_number.trace(mode = 'w', callback = lambda name, index, mode, 
            carton_number = carton_number: self.__remaining_cartons(carton_number = carton_number.get(),
            carton_entry_error = carton_entry_error, remaining_cartons = remaining_cartons)
        )

        label = tk.Label(self.root, text = "Remaining Cartons:")
        label.config(font = ("Arial", 12, "bold"))
        label.place(x = 25, y = 10)

        label = tk.Label(self.root, textvariable = remaining_cartons)
        label.config(font = ("Arial", 12))
        label.place(x = 190, y = 10)

        label = tk.Label(self.root, text = "Carton Count:")
        label.config(font = ("Arial", 12, "bold"))
        label.place(x = 25, y = 40)

        label = tk.Label(self.root, textvariable = carton_entry_error)
        label.config(font = ("Arial", 9, "italic"), fg = "red")
        label.place(x = 25, y = 70)
        
        entry = tk.Entry(self.root, textvariable = carton_number)
        entry.insert(index = 0, string = str(self.carton_entry_count))
        entry.focus_set()
        entry.place(x = 150, y = 45)

        button = tk.Button(self.root, text = "Return to\nMain Menu", command = lambda: self.__confirm_main_menu())
        button.config(height = 3, width = 10, font = ("Arial", 12, "bold"))
        button.place(x = 25, y = 110)        

        button = tk.Button(self.root, text = "Submit Carton Count", 
            command = lambda: self.__next(count = carton_number.get()))
        button.config(height = 3, width = 20, font = ("Arial", 12, "bold"))
        button.place(x = 150, y = 110)


    """ Update remaining cartons. """
    def __remaining_cartons(self, carton_number, carton_entry_error, remaining_cartons):

        try:
            carton_number = int(carton_number)

            if (carton_number <= 0):
                carton_entry_error.set("ERROR: Carton Count must be greater than 0.")
                remaining_cartons.set(str(self.remaining_cartons))

            else:
                carton_entry_error.set("")
                remaining = self.remaining_cartons - carton_number
                
                if remaining <= 0:
                    remaining = 0

                remaining_cartons.set(str(remaining))

        except ValueError:
            carton_entry_error.set("ERROR: Carton Count must be an integer value (i.e., 5, 10, etc.)")
            remaining_cartons.set(str(self.remaining_cartons))


    """ Verify there are more cartons to process. """ 
    def __are_cartons_remaining(self):
        if self.remaining_cartons <= 0:
            messagebox.showinfo("Complete", "The macro is finished. Returning to main menu.")
            self.__main_menu()

        else:
            self.__set_carton_count()
            

    """ Ask user if any cartons caused issues during entry. """ 
    def __any_issues(self):
        self.__set_geometry(width = 300, height = 180)
        tk.Label(self.root, text = "Were there any issues?", font = ("Arial", 14, "bold")).place(x = 40, y = 10)

        # Issues with entered values. 
        reprint_button = tk.Button(self.root, text = "Yes", command = lambda: self.__remove_problem_cells())
        reprint_button.config(height = 4, width = 10, font = ("Arial", 14))
        reprint_button.place(x = 20, y = 50)
        
        # No issues. 
        carton_manifest_button = tk.Button(self.root, text = "No", command = lambda: self.__are_cartons_remaining())
        carton_manifest_button.config(height = 4, width = 10, font = ("Arial", 14))
        carton_manifest_button.place(x = 155, y = 50)


    """ Allow user to remove entries causing processing issues. """ 
    def __remove_problem_cells(self):
        self.__set_geometry(width = 420, height = 500)

        self.previous_page = "__remove_problem_cells"

        label = tk.Label(self.root, text = "Search Results")
        label.config(font = ("Arial", 14))
        label.place(x = 20, y = 70)

        search_results = tk.Listbox(self.root)
        search_results.place(x = 25, y = 100)

        label = tk.Label(self.root, text = "Values to Ignore")
        label.config(font = ("Arial", 14))
        label.place(x = 235, y = 70)

        ignore_values = tk.Listbox(self.root)
        ignore_values.place(x = 240, y = 100)

        label = tk.Label(self.root, text = "Search")
        label.config(font = ("Arial", 14))
        label.place(x = 25, y = 10)

        # Reactive value for searched value.  
        search_value = tk.StringVar()
        search_value.trace(mode = 'w', callback = lambda name, index, mode, 
            search_value = search_value: self.__filter_columns(query = search_value.get(), 
            search_results = search_results, ignore_values = ignore_values))

        # Allow user to search for value. 
        entry = tk.Entry(self.root, textvariable = search_value)
        entry.insert(index = 0, string = '')
        entry.focus_set()
        entry.place(x = 20, y = 40)

        button = tk.Button(self.root, text = "Add", 
            command = lambda: self.__add_value(search_results = search_results, ignore_values = ignore_values))
        button.config(height = 2, width = 10, font = ("Arial", 12))
        button.place(x = 35, y = 280)

        button = tk.Button(self.root, text = "Remove",
            command = lambda: self.__remove_value(ignore_values = ignore_values, search_results = search_results))
        button.config(height = 2, width = 10, font = ("Arial", 12))
        button.place(x = 255, y = 280)

        button = tk.Button(self.root, text = "Return to\nMain Menu", command = lambda: self.__confirm_main_menu())
        button.config(height = 3, width = 10, font = ("Arial", 12, "bold"))
        button.place(x = 25, y = 400)

        button = tk.Button(self.root, text = "Retry Values", 
            command = lambda: self.__retry(ignore_values = ignore_values))
        button.config(height = 3, width = 20, font = ("Arial", 12, "bold"))
        button.place(x = 190, y = 400)

        # Call an empty search query to show all values initially. 
        self.__filter_columns(query = '', search_results = search_results, ignore_values = ignore_values)


    """ Filter search results """
    def __filter_columns(self, query, search_results, ignore_values):
        search_results.delete(0, tk.END)
        ignore_values = ignore_values.get(0, tk.END)

        query = query.strip()

        # Show matching query results. 
        if query != '':
            results = [data for data in set(self.data) if query in str(data)]
            for result in sorted(results):
                if result not in ignore_values:
                    search_results.insert(tk.END, result)

        # Show all results. 
        else:
            for data in sorted(set(self.data)):
                if data not in ignore_values:
                    search_results.insert(tk.END, data)


    """ Add value to "Ignore Values" list. """
    def __add_value(self, search_results, ignore_values):
        
        try:
            index = search_results.curselection()
            selected = search_results.get(index)

            if selected not in ignore_values.get(0, tk.END):
                ignore_values.insert(tk.END, selected)
                search_results.delete(index)

        except:
            pass


    """ Remove value from "Ignore Values" list. """
    def __remove_value(self, ignore_values, search_results):

        try:
            index = ignore_values.curselection()
            selected = ignore_values.get(index)

            _search_results = list(search_results.get(0, tk.END))
            _search_results.append(selected)
            search_results.delete(0, tk.END)
            
            for value in sorted(_search_results):
                search_results.insert(tk.END, value)
            
            ignore_values.delete(index)

        except:
            pass
    

    """ Retry previous values minus skip values. """
    def __retry(self, ignore_values):
        
        if self.end_index != 0:
            ignore_values = ignore_values.get(0, tk.END)
            self.data = [i for i in self.data if i not in ignore_values]
            self.__set_position()

    
    """ Try next set of values based on user-supplied carton count. """ 
    def __next(self, count):
        try:
            count = int(count)

            if (count <= 0):
                self.__error(message = "Carton Count must be greater than 0.")
                return()

            else:
                self.carton_entry_count = count

        except ValueError:
            self.__error(message = "Carton Count must be an integer value (i.e., 5, 10, etc.)")
            return()

        self.data = copy.deepcopy(self.data_backup)

        # Update previos start and end indices.
        # This allows the user to go back to the previous page
        # with the correct data.  
        self.previous_start_index = self.start_index
        self.previous_end_index = self.end_index 

        # First time running. 
        if self.end_index == 0:
            self.end_index = self.carton_entry_count
            self.data = self.data[self.start_index:self.end_index]

        else:
            self.start_index = self.end_index
            self.end_index += self.carton_entry_count
            self.data = self.data[self.start_index:self.end_index]

        # Update remaining cartons
        self.remaining_cartons -= len(self.data)  
        self.__set_position()


    """ Perform the carton manifest macro. """ 
    def __carton_manifest_macro(self):
        
        try:
            for i in range(len(self.data)):

                # Pause after each set of 100 entries to give the program time to catch up
                if (i > 0) and (i % 100) == 0:
                    time.sleep(3)

                # Write values to screen
                pyautogui.typewrite(str(self.data[i]))

                # Press 'enter'
                pyautogui.keyDown("enter")
                pyautogui.keyUp("enter")
            
            self.__any_issues()

        except:
            self.__remove_problem_cells()


    """ Configuration settings for the salesman code maintenance macro """
    def __config_salesman_code(self):

        self.macro = "salesman_code"
        self.__set_geometry(width = 400, height = 420, resizable = True)

        label = tk.Label(self.root, text = "Salesman Code Maintenance Macro Instructions")
        label.config(font = ("Arial", 12, "bold"))
        label.place(x = 10, y = 10)

        label = tk.Label(self.root, text = '-'*65)
        label.config(font = ("Arial", 16))
        label.place(x = 0, y = 45)

        label = tk.Label(self.root, text = "1. Click the \"Select File\" button.")
        label.config(font = ("Arial", 12))
        label.place(x = 25, y = 70)

        label = tk.Label(self.root, text = "2. Select data file.")
        label.config(font = ("Arial", 12))
        label.place(x = 25, y = 100)

        label = tk.Label(self.root, text = "a. This file must contain 3 columns: ")
        label.config(font = ("Arial", 12))
        label.place(x = 40, y = 130)

        label = tk.Label(self.root, text = "(1) Must be the \"code\" values.")
        label.config(font = ("Arial", 12))
        label.place(x = 55, y = 160)

        label = tk.Label(self.root, text = "(2) Must be the \"style\" values.")
        label.config(font = ("Arial", 12))
        label.place(x = 55, y = 190)

        label = tk.Label(self.root, text = "(3) Must be the \"suffix\" values.")
        label.config(font = ("Arial", 12))
        label.place(x = 55, y = 220)

        label = tk.Label(self.root, text = "3. Click the \"Submit\" button.")
        label.config(font = ("Arial", 12))
        label.place(x = 25, y = 250)

        label = tk.Label(self.root, text = "Selected File: ")
        label.config(font = ("Arial", 12, "bold"))
        label.place(x = 25, y = 300)

        label = tk.Label(self.root, textvariable = self.uploaded_file)
        label.config(font = ("Arial", 12, "italic"))
        label.place(x = 140, y = 300)

        button = tk.Button(self.root, text = "Select\nFile", command = lambda: self.__upload())
        button.config(height = 3, width = 10, font = ("Arial", 12, "bold"))
        button.place(x = 20, y = 330)

        button = tk.Button(self.root, text = "Submit", command = lambda: self.__verify_salesman_code())
        button.config(height = 3, width = 10, font = ("Arial", 12, "bold"))
        button.place(x = 145, y = 330)

        button = tk.Button(self.root, text = "Return to\nMain Menu", command = lambda: self.__confirm_main_menu())
        button.config(height = 3, width = 10, font = ("Arial", 12, "bold"))
        button.place(x = 270, y = 330)


    """ Verify the requirements for the salesman code maintenance macro. """ 
    def __verify_salesman_code(self):

        if self.data is None:
            self.__error(message = "Please select a file to proceed.")
            return()

        columns = self.data.columns.values.tolist()
        
        if len(columns) != 3:
            self.__error("This file must contain 3 columns.")
            return()

        for i in self.data.iloc[:, 0]:
            if len(str(i)) != 3:
                self.__error("Each value in column 1 must be 3 characters.")
                return()

        for i in self.data.iloc[:, 1]:
            if len(str(i)) > 7:
                self.__error("Each value in column 2 must be between 5-7 characters.")
                return()

        for i in self.data.iloc[:, 2]:
            if len(str(i)) != 2:
                self.__error("Each value in column 3 must be 2 characters.")
                return()

        self.__set_position_instructions()  


    """ Salesman code maintenance macro commands """
    def __salesman_code_macro(self):

        code = self.data.iloc[:, 0]
        style = self.data.iloc[:, 1]
        suffix = self.data.iloc[:, 2]

        # Catch pyautogui.FAILSAFE.
        try:
            for i in range(len(self.data)):

                # Pause after each set of 100 entries to give the program time to catch up
                if (i > 0) and (i % 100) == 0:
                    time.sleep(3)

                # Press 'f6'
                pyautogui.keyDown("f6")
                pyautogui.keyUp("f6")

                # Enter 'code'
                pyautogui.typewrite(str(code[i]).upper())

                # Enter 'style'
                pyautogui.typewrite(str(style[i]))

                # Enter 'tab'
                pyautogui.keyDown("tab")
                pyautogui.keyUp("tab")

                # Enter 'suffix'
                pyautogui.typewrite(str(suffix[i]).upper())

                # Enter 'enter'
                pyautogui.keyDown("enter")
                pyautogui.keyUp("enter")

                # Enter 'f16' = 'shift' + 'f4'
                pyautogui.keyDown("shift")
                pyautogui.keyDown("f4")
                pyautogui.keyUp("f4")
                pyautogui.keyUp("shift")

            messagebox.showinfo("Complete", "The macro is finished. Returning to main menu.")
            self.__main_menu()

        except:
            messagebox.showinfo("Stop", "Failsafe activated. The macro has stopped.")


    """ Error helper function """
    def __error(self, message):
        messagebox.showerror("Error", message)


""" Start the application """
if __name__ == "__main__":
    
    # Delay between each pyautogui call
    pyautogui.PAUSE = 0.0

    # Quickly move the mouse to the top-left of the screen to stop the program.
    pyautogui.FAILSAFE = True

    root = tk.Tk()
    app = MainApplication(root)
    root.mainloop()


