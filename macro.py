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
    def __init__(self, root):
        self.root = root
        self.root.title("PKMS Automation Tool")
        
        # Start at main menu.
        self.__main_menu()


    # Application main menu. Allows user to select macro choice. 
    def __main_menu(self):
        self.data = None 
        self.uploaded_file = tk.StringVar()
        self.uploaded_file.set("None")
        self.macro = None

        self.carton_entry_count = 10
        self.start_index = 0
        self.end_index = 0
        self.previous_start_index = 0
        self.previous_end_index = 0
        self.remaining_cartons = 0

        self.previous_page = None

        self.__set_geometry(width = 300, height = 180)
        tk.Label(self.root, text = "Select Macro", font = ("Arial", 18, "bold")).place(x = 90, y = 10)

        # Reprint Selection
        reprint_button = tk.Button(self.root, text="Reprint", command=lambda: self.__config_reprint())
        reprint_button.config(height = 5, width = 10, font=("Arial", 18))
        reprint_button.place(x = 20, y = 50)
        
        # Carton Manifest Selection
        carton_manifest_button = tk.Button(self.root, text="Carton\nManifest", command=lambda: self.__config_carton_manifest())
        carton_manifest_button.config(height = 5, width = 10, font=("Arial", 18))
        carton_manifest_button.place(x=155, y=50)


    # Clear the window frame, and set the frame dimensions 
    def __set_geometry(self, width, height, resizable=False):
        for widget in self.root.winfo_children():
            widget.destroy()

        self.root.geometry("%sx%s" % (width, height))
        self.root.resizable(resizable, resizable)    


    # Configuration settings for the reprint macro
    def __config_reprint(self):

        self.macro = "reprint"
        self.__set_geometry(width = 400, height = 375, resizable = True)

        label = tk.Label(self.root, text = "Reprint Macro Instructions")
        label.config(font = ("Arial", 18, "bold"))
        label.place(x = 65, y = 10)

        label = tk.Label(self.root, text = '-'*65)
        label.config(font = ("Arial", 18, "bold"))
        label.place(x = 0, y = 35)

        label = tk.Label(self.root, text = "1. Click the \"Upload\" button.")
        label.config(font = ("Arial", 16))
        label.place(x = 25, y = 70)

        label = tk.Label(self.root, text = "2. Select data file.")
        label.config(font = ("Arial", 16))
        label.place(x = 25, y = 100)

        label = tk.Label(self.root, text = "a. This file must contain the following columns: ")
        label.config(font = ("Arial", 16))
        label.place(x = 40, y = 130)

        label = tk.Label(self.root, text = "i. CDSTYL")
        label.config(font = ("Arial", 16))
        label.place(x = 55, y = 160)

        label = tk.Label(self.root, text = "ii. CHCASN")
        label.config(font = ("Arial", 16))
        label.place(x = 55, y = 190)

        label = tk.Label(self.root, text = "3. Click the \"Submit\" button.")
        label.config(font = ("Arial", 16))
        label.place(x = 25, y = 220)

        label = tk.Label(self.root, text = "Selected File: ")
        label.config(font = ("Arial", 16, "bold"))
        label.place(x = 25, y = 270)

        label = tk.Label(self.root, textvariable = self.uploaded_file)
        label.config(font = ("Arial", 16, "italic"))
        label.place(x = 125, y = 270)

        button = tk.Button(self.root, text = "Upload", command = lambda: self.__upload())
        button.config(height = 3, width = 12, font=("Arial", 16, "bold"))
        button.place(x = 20, y = 300)

        button = tk.Button(self.root, text = "Submit", command = lambda: self.__verify_reprint())
        button.config(height = 3, width = 12, font=("Arial", 16, "bold"))
        button.place(x = 145, y = 300)

        button = tk.Button(self.root, text="Return to\nMain Menu", command=lambda: self.__main_menu())
        button.config(height = 3, width = 12, font=("Arial", 16, "bold"))
        button.place(x = 270, y = 300)


    # Verify the requirements for the reprint macro. 
    def __verify_reprint(self):

        if self.data is None:
            self.error(message = "Please upload a file to proceed.")
            return()

        columns = self.data.columns.values.tolist()
        
        if "CDSTYL" not in columns or "CHCASN" not in columns:
            self.error("You must provide a valid file that contains the following columns: CDSTYL and CHCASN")
            return()
        
        workbook = self.data.sort_values("CDSTYL")
        self.data = workbook["CHCASN"].values
        self.__set_position_instructions()  


    # Upload the data file.
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
                self.error(message = "Please select a valid file.")


    # Show users how to use the position setter.
    def __set_position_instructions(self):

        self.__set_geometry(width = 640, height = 350)

        label = tk.Label(self.root, text = "Macro Placement Instructions")
        label.config(font = ("Arial", 18, "bold"))
        label.place(x = 160, y = 10)

        label = tk.Label(self.root, text = '-'*150)
        label.config(font = ("Arial", 18, "bold"))
        label.place(x = 0, y = 35)

        label = tk.Label(self.root, text = "1. Click the \"Set Position\" button.")
        label.config(font = ("Arial", 16))
        label.place(x = 25, y = 70)

        label = tk.Label(self.root, text = "2. Move this application next to the PKMS system.")
        label.config(font = ("Arial", 16))
        label.place(x = 25, y = 100)

        label = tk.Label(self.root, text = "a. Essentially, place this on the PKMS system.")
        label.config(font = ("Arial", 16))
        label.place(x = 55, y = 130)

        label = tk.Label(self.root, text = "3. Click the direction where the mouse cursor should go to click into the PKMS system.")
        label.config(font = ("Arial", 16))
        label.place(x = 25, y = 160)

        label = tk.Label(self.root, text = "4. Wait until the application finishes or another prompt displays a new set of instructions.")
        label.config(font = ("Arial", 16))
        label.place(x = 25, y = 190)

        label = tk.Label(self.root, text = "** NOTE: To stop the macro, quickly move the mouse cursor to the top left corner of screen.")
        label.config(font = ("Arial", 13, "bold"), fg = "red")
        label.place(x = 25, y = 230)

        button = tk.Button(self.root, text = "Set Positition", command = lambda: self.__set_position())
        button.config(height = 3, width = 30, font=("Arial", 16, "bold"))
        button.place(x = 30, y = 270)

        button = tk.Button(self.root, text="Return to\nMain Menu", command=lambda: self.__main_menu())
        button.config(height = 3, width = 30, font=("Arial", 16, "bold"))
        button.place(x = 330, y = 270)


    # Set the mouse cursor position
    def __set_position(self):

        self.__set_geometry(350, 350)

        label = tk.Label(self.root, text = "Set Position and Mouse Direction")
        label.config(font = ("Arial", 18, "bold"))
        label.place(x = 25, y = 10)

        label = tk.Label(self.root, text = '-'*100)
        label.config(font = ("Arial", 16, "bold"))
        label.place(x = 0, y = 35)

        # Move Up
        button = tk.Button(self.root, text="Move Up", command=lambda: self.__move_cursor(0, -200))
        button.config(height = 3, width = 10, font=("Arial", 16, "bold"))
        button.place(x = 125, y = 75)

        # Move Down
        button = tk.Button(self.root, text="Move Down", command=lambda: self.__move_cursor(0, 250))
        button.config(height = 3, width = 10, font=("Arial", 16, "bold"))
        button.place(x = 125, y = 150)

        # Move Left
        button = tk.Button(self.root, text="Move Left", command=lambda: self.__move_cursor(-175, 0))
        button.config(height = 3, width = 10, font=("Arial", 16, "bold"))
        button.place(x = 15, y = 110)

        # Move Right
        button = tk.Button(self.root, text="Move Right", command=lambda: self.__move_cursor(175, 0))
        button.config(height = 3, width = 10, font=("Arial", 16, "bold"))
        button.place(x = 235, y = 110)

        if self.macro == "reprint":

            button = tk.Button(self.root, text="Return to\nMain Menu", command=lambda: self.__main_menu())
            button.config(height = 3, width = 20, font=("Arial", 16, "bold"))
            button.place(x = 75, y = 260)

        else:
            button = tk.Button(self.root, text="Return to\nMain Menu", command=lambda: self.__main_menu())
            button.config(height = 3, width = 12, font=("Arial", 16, "bold"))
            button.place(x = 25, y = 260)

            button = tk.Button(self.root, text="Previous Page", command=lambda: self.__previous_page())
            button.config(height = 3, width = 18, font=("Arial", 16, "bold"))
            button.place(x = 160, y = 260)


    # Allow user to go back to previous page. 
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


    # Move cursor to appropriate position
    def __move_cursor(self, x, y):

        # Move mouse to position & click into program
        pyautogui.moveRel(x, y, duration=pyautogui.MINIMUM_DURATION)
        pyautogui.click()

        if self.macro == "reprint":
            self.__reprint_macro()

        elif self.macro == "carton_manifest":
            self.__carton_manifest_macro()
            
            
    # Macro commands
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

                # Press 'f19'
                pyautogui.keyDown("f12")
                pyautogui.keyUp("f12")

            messagebox.showinfo("Complete", "The macro is finished. Returning to main menu.")
            self.__main_menu()

        except:
            messagebox.showinfo("Stop", "Failsafe activated. The macro has stopped.")


    # Configuration settings for the reprint macro
    def __config_carton_manifest(self):

        self.macro = "carton_manifest"
        self.__set_geometry(width = 400, height = 350, resizable = True)

        label = tk.Label(self.root, text = "Carton Manifest Macro Instructions")
        label.config(font = ("Arial", 18, "bold"))
        label.place(x = 30, y = 10)

        label = tk.Label(self.root, text = '-'*100)
        label.config(font = ("Arial", 18, "bold"))
        label.place(x = 0, y = 35)

        label = tk.Label(self.root, text = "1. Click the \"Upload\" button.")
        label.config(font = ("Arial", 16))
        label.place(x = 25, y = 70)

        label = tk.Label(self.root, text = "2. Select data file.")
        label.config(font = ("Arial", 16))
        label.place(x = 25, y = 100)

        label = tk.Label(self.root, text = "a. This file must contain the following column: ")
        label.config(font = ("Arial", 16))
        label.place(x = 40, y = 130)

        label = tk.Label(self.root, text = "i. CHCASN")
        label.config(font = ("Arial", 16))
        label.place(x = 55, y = 160)

        label = tk.Label(self.root, text = "3. Click the \"Submit\" button.")
        label.config(font = ("Arial", 16))
        label.place(x = 25, y = 190)

        label = tk.Label(self.root, text = "Selected File: ")
        label.config(font = ("Arial", 16, "bold"))
        label.place(x = 25, y = 240)

        label = tk.Label(self.root, textvariable = self.uploaded_file)
        label.config(font = ("Arial", 16, "italic"))
        label.place(x = 125, y = 240)

        button = tk.Button(self.root, text = "Upload", command = lambda: self.__upload())
        button.config(height = 3, width = 12, font=("Arial", 16, "bold"))
        button.place(x = 20, y = 270)

        button = tk.Button(self.root, text = "Submit", command = lambda: self.__verify_carton_manifest())
        button.config(height = 3, width = 12, font=("Arial", 16, "bold"))
        button.place(x = 145, y = 270)

        button = tk.Button(self.root, text="Return to\nMain Menu", command=lambda: self.__main_menu())
        button.config(height = 3, width = 12, font=("Arial", 16, "bold"))
        button.place(x = 270, y = 270)


    # Verify valid file for carton manifest. 
    def __verify_carton_manifest(self):
        if self.data is None:
            self.error(message = "Please upload a file to proceed.")
            return()

        columns = self.data.columns.values.tolist()
        
        if "CHCASN" not in columns:
            self.error("You must provide a valid file that contains the following column: CHCASN")
            return()

        self.data = self.data["CHCASN"].values.tolist()
        self.data_backup = copy.deepcopy(self.data)
        self.__set_carton_count()


    # Allow user to set number of carton entries. 
    def __set_carton_count(self):

        self.__set_geometry(width = 450, height = 200)

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
        label.config(font = ("Arial", 16, "bold"))
        label.place(x = 25, y = 10)

        label = tk.Label(self.root, textvariable = remaining_cartons)
        label.config(font = ("Arial", 16))
        label.place(x = 180, y = 10)

        label = tk.Label(self.root, text = "Carton Count: ")
        label.config(font = ("Arial", 16, "bold"))
        label.place(x = 25, y = 40)

        label = tk.Label(self.root, textvariable = carton_entry_error)
        label.config(font = ("Arial", 14, "italic"), fg = "red")
        label.place(x = 25, y = 70)
        
        entry = tk.Entry(self.root, textvariable = carton_number)
        entry.insert(index = 0, string = str(self.carton_entry_count))
        entry.focus_set()
        entry.place(x = 140, y = 40)

        button = tk.Button(self.root, text = "Return to\nMain Menu", command = lambda: self.__main_menu())
        button.config(height = 3, width = 12, font = ("Arial", 16, "bold"))
        button.place(x = 25, y = 120)        

        button = tk.Button(self.root, text = "Submit Carton Count", 
            command = lambda: self.__next(count = carton_number.get()))
        button.config(height = 3, width = 30, font = ("Arial", 16, "bold"))
        button.place(x = 150, y = 120)


    # Update remaining cartons. 
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


    # Verify there are more cartons to process. 
    def __are_cartons_remaining(self):
        if self.remaining_cartons <= 0:
            messagebox.showinfo("Complete", "The macro is finished. Returning to main menu.")
            self.__main_menu()

        else:
            self.__set_carton_count()
            

    # Ask user if any cartons caused issues during entry. 
    def __any_issues(self):
        self.__set_geometry(width = 300, height = 180)
        tk.Label(self.root, text = "Were there any issues?", font = ("Arial", 18, "bold")).place(x = 40, y = 10)

        # Issues with entered values. 
        reprint_button = tk.Button(self.root, text="Yes", command=lambda: self.__remove_problem_cells())
        reprint_button.config(height = 5, width = 10, font=("Arial", 18))
        reprint_button.place(x = 20, y = 50)
        
        # No issues. 
        carton_manifest_button = tk.Button(self.root, text="No", command=lambda: self.__are_cartons_remaining())
        carton_manifest_button.config(height = 5, width = 10, font=("Arial", 18))
        carton_manifest_button.place(x=155, y=50)


    # Allow user to remove entries causing processing issues. 
    def __remove_problem_cells(self):
        self.__set_geometry(width = 450, height = 500)

        self.previous_page = "__remove_problem_cells"

        label = tk.Label(self.root, text = "Search Results")
        label.config(font = ("Arial", 16))
        label.place(x = 25, y = 70)

        search_results = tk.Listbox(self.root)
        search_results.place(x = 25, y = 100)

        label = tk.Label(self.root, text = "Values to Ignore")
        label.config(font = ("Arial", 16))
        label.place(x = 240, y = 70)

        ignore_values = tk.Listbox(self.root)
        ignore_values.place(x = 240, y = 100)

        label = tk.Label(self.root, text = "Search")
        label.config(font = ("Arial", 16))
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
        entry.place(x = 20, y = 30)

        button = tk.Button(self.root, text = "Add", 
            command = lambda: self.__add_value(search_results = search_results, ignore_values = ignore_values))
        button.config(height = 2, width = 20, font = ("Arial", 15))
        button.place(x = 25, y = 280)

        button = tk.Button(self.root, text = "Remove",
            command = lambda: self.__remove_value(ignore_values = ignore_values, search_results = search_results))
        button.config(height = 2, width = 20, font = ("Arial", 15))
        button.place(x = 240, y = 280)

        button = tk.Button(self.root, text = "Return to\nMain Menu", command = lambda: self.__main_menu())
        button.config(height = 3, width = 14, font = ("Arial", 16, "bold"))
        button.place(x = 25, y = 400)

        button = tk.Button(self.root, text = "Retry Values", 
            command = lambda: self.__retry(ignore_values = ignore_values))
        button.config(height = 3, width = 25, font = ("Arial", 16, "bold"))
        button.place(x = 190, y = 400)

        # Call an empty search query to show all values initially. 
        self.__filter_columns(query = '', search_results = search_results, ignore_values = ignore_values)


    # Filter search results
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


    # Add value to "Ignore Values" list.
    def __add_value(self, search_results, ignore_values):
        
        try:
            index = search_results.curselection()
            selected = search_results.get(index)

            if selected not in ignore_values.get(0, tk.END):
                ignore_values.insert(tk.END, selected)
                search_results.delete(index)

        except:
            pass


    # Remove value from "Ignore Values" list.
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
    

    # Retry previous values minus skip values.
    def __retry(self, ignore_values):
        
        if self.end_index != 0:
            ignore_values = ignore_values.get(0, tk.END)
            self.data = [i for i in self.data if i not in ignore_values]
            self.__set_position()

    
    # Try next set of values based on user-supplied carton count. 
    def __next(self, count):
        try:
            count = int(count)

            if (count <= 0):
                self.error(message = "Carton Count must be greater than 0.")
                return()

            else:
                self.carton_entry_count = count

        except ValueError:
            self.error(message = "Carton Count must be an integer value (i.e., 5, 10, etc.)")
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


    # Perform the carton manifest macro. 
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


    # Error helper function
    def error(self, message):
        messagebox.showerror("Error", message)


if __name__ == "__main__":
    
    # Delay between each pyautogui call
    pyautogui.PAUSE = 0.1

    # Quickly move the mouse to the top-left of the screen to stop the program.
    pyautogui.FAILSAFE = True

    root = tk.Tk()
    app = MainApplication(root)
    root.mainloop()


