#To export "pyinstaller --onefile --noconsole KitchenDuty.py" noconsole suppresses the console window
import os
import random
import pickle
import json
import calendar
import sys
import customtkinter
import threading
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.http import BatchHttpRequest


def get_days_in_month(year, month):
    _, num_days = calendar.monthrange(year, month)
    return num_days

#Function to get all possible pairs in a list
def getPairs(lst):
    pairs = []
    for i in range(len(lst)):
        # For each element, iterate over all subsequent elements
        for j in range(i+1, len(lst)):
            # Create a pair as a tuple
            pair = (lst[i], lst[j])
            # Add the pair to the list of pairs
            pairs.append(pair)
    return pairs

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar']

def callback(request_id, response, exception):
    if exception is not None:
        print(exception)
    else:
        print("Event created: " + response['summary'])

def generateSchedule(year,month):
    #Verification stuff
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    service = build('calendar', 'v3', credentials=creds)
    # Call the Calendar API
    # open the file in read mode
    with open('details.json', 'r') as file:
        data = json.load(file)

    #Create a string list of Room No + Name
    namesList = []
    for item in data:
        color = 1 if item["Enrollment"].lower() == "sp" else 6
        namesList.append("[" + str(item["Room No"]) + "] "+ item["Name"]+ ","+ str(color))

    #Get days in a month to determine number of iterations
    days = get_days_in_month(year, month)
    #List for storing all unused pairs
    Pairs = getPairs(namesList)
    #List for storing USABLE pairs
    ValidPairs = getPairs(namesList)
    #Dictionary for counting how many shifts are done per month
    countDict = {value: 0 for value in namesList}
    # Select a random pair for each day of the month and add it to calendar, +1 here because in range excludes last value

    # Create batch request
    batch = service.new_batch_http_request(callback=callback)

    for i in range(1, days+1):
        #Random pair is seleced here using the random function
        Selected = ValidPairs[random.randint(0,len(ValidPairs)-1)]
        #Change tuple into list
        Pair = list(Selected)

        residentAssistant1 = Pair[0].split(",")
        residentAssistant2 = Pair[1].split(",")
        events = [
            {
        'summary': residentAssistant1[0],
        'description': 'Kitchen Duty',
        'colorId' : int(residentAssistant1[1]),
        'start': {'date': str(year)+'-'+ str(month) +'-'+str(i),},
        'end': {'date': str(year)+'-'+ str(month) +'-'+str(i),},
        },
        {
        'summary': residentAssistant2[0],
        'description': 'Kitchen Duty',
        'colorId' : int(residentAssistant2[1]),
        'start': {'date': str(year)+'-'+ str(month) +'-'+str(i),},
        #If you see this line hemlo, don't tell ;> [P.s. GOD BLESS TANEZAKI ATSUMIII]
        'end': {'date': str(year)+'-'+ str(month) +'-'+str(i),}
        }
        ]
        for event in events:
            batch.add(service.events().insert(calendarId='primary', body=event))

        #Add count to number of shifts
        countDict[Pair[0]] += 1
        countDict[Pair[1]] += 1
        Pairs.remove(Selected)
        ValidPairs.remove(Selected)
        #IMPORTANT!!! If any person has done 3 shifts, their entries have to be removed from valid pairs
        if countDict[Pair[0]] == 3:
            ValidPairs = [tup for tup in ValidPairs if Pair[0] not in tup]
        if countDict[Pair[1]] == 3:
            ValidPairs = [tup for tup in ValidPairs if Pair[1] not in tup]
        if not ValidPairs:
            break

    batch.execute()
    
    #Write remaing pairs left into remaining pairs file along with number of shifts each person has done
    with open('remainingPairs.txt', 'w') as f:
        for item in Pairs:
            f.write(str(item) + "\n")
            
    return countDict

def save_to_json(treeview):
    # Get all records from the treeview
    records = treeview.get_children()
    
    # List to hold data dictionaries
    data = []
    
    for record in records:
        # Get values from each record
        values = treeview.item(record, 'values')

        # Create a dictionary from the record values and append it to the data list
        data.append({
            "Name": values[0],
            "Room No": values[1],
            "Enrollment": values[2]
        })

    # Save the data list to the JSON file
    with open("details.json", "w") as file:
        json.dump(data, file)

def validate(input):
    # check if input is a digit
    if input.isdigit():
        return True
    elif input == "":  # Allow empty input
        return True
    else:
        return False

def styleTree():
        # Create the style
        style = ttk.Style()
        style.theme_use("clam")

        # Configure the Treeview style
        style.configure("Treeview",
                        background="black",
                        fieldbackground="black",
                        foreground="white")

        # Configure the Treeview heading style
        style.configure("Treeview.Heading",
                        background="black",
                        foreground="white")
        #Hover style
        style.map('Treeview.Heading',
                    background=[('active', 'grey')],
                    foreground=[('active', 'white')])
    
def month_to_int(month):
        month_dict = {"Jan / 1 月": 1, "Feb / 2 月": 2, "Mar / 3 月": 3, "Apr / 4 月": 4, "May / 5 月": 5, "Jun / 6 月": 6, "Jul / 7 月": 7, "Aug / 8 月": 8, "Sep / 9 月": 9, "Oct / 10 月": 10, "Nov / 11 月": 11, "Dec / 12 月": 12}
        return month_dict[month]

def openDir():
        # Ask for confirmation before changing the user
        response = messagebox.askyesno("Confirmation",
                                    "Are you sure you want to change the user?\nユーザー変更てもよろしいですか？")
        if response == True:
            # Delete token, doing this requires user to log in again
            if os.path.exists('token.pickle'):
                os.remove('token.pickle')
            # Get the directory where the script is located
            if getattr(sys, 'frozen', False):
                script_dir = os.path.dirname(sys.executable)  # Executable location
            else:
                script_dir = os.path.dirname(os.path.abspath(__file__))  # Script location
            # Open the file explorer at the script's directory
            os.startfile(script_dir)
        else:
            # If the user clicked 'No', do nothing
            return
        
def treeview_sort_column(tv, col, reverse):
        l = [(tv.set(k, col), k) for k in tv.get_children('')]
        try:
            l.sort(key = lambda x: int(x[0]), reverse=reverse)
        except ValueError:
            l.sort(reverse=reverse)

        # rearrange items in sorted positions
        for index, (val, k) in enumerate(l):
            tv.move(k, '', index)
        # reverse sort next time
        tv.heading(col, command=lambda: treeview_sort_column(tv, col, not reverse))

# global variable to keep track of editing state
is_editing = False
selected_id=None

def main():

    def button_callback():
        month = optionmenu_1.get()
        year = optionmenu_2.get()
        if month == "Month / 月" or year == "Year / 年":
            error_message.configure(text="Please select a month and a year\n日付を入力してくだだい （#^ω^）", text_color="#fa3c58")
        elif len(recTree.get_children()) < 3:
            error_message.configure(text="Please add at least 3 records\n最低20名を記入してっください", text_color="#fa3c58")
        else:
            button_1.configure(state="disabled")
            error_message.configure(text_color="#00FFD3",text=f"Generating Schedule, Please wait... \n スケジュール作成中… (´・ω・`)")
            
            # Define a new function that includes the generateSchedule function and re-enables the button
            def generate_and_reenable():
                shiftDict = generateSchedule(int(year),month_to_int(month))
                #Destroy original label
                label2.destroy()
                #Display shifts done in Shifts tab of gui
                styleTree()
                # Create the treeview
                tree = ttk.Treeview(tabView.tab("Shifts Done\nKD回数"), show='headings')
                # define columns
                tree["columns"]=("Name / 名前","Shifts Done / KD回数")
                # format columns
                tree.column("Name / 名前", anchor="center", width=100)
                tree.column("Shifts Done / KD回数", anchor="center", width=100)
                # set heading
                tree.heading("Name / 名前", text="Name / 名前", anchor="center", command=lambda: treeview_sort_column(tree, "Name / 名前", False))
                tree.heading("Shifts Done / KD回数", text="Shifts Done / KD回数", anchor="center", command=lambda: treeview_sort_column(tree, "Shifts Done / KD回数", False))
                # add data
                for key, value in shiftDict.items():
                    lis = key.split(',')
                    tree.insert("", tk.END, values=(lis[0], value))
                # Create the scrollbar
                scrollbar = ttk.Scrollbar(tabView.tab("Shifts Done\nKD回数"), orient="vertical", command=tree.yview)
                tree.configure(yscrollcommand=scrollbar.set)

                # Create Header for shifts done
                label3 = customtkinter.CTkLabel(tabView.tab("Shifts Done\nKD回数"), text= month +' '+year , font = ("Italic",20), padx=0, pady=0)
                label3.grid(row=0, column=0, columnspan=2, pady=0, padx=0, sticky="nsew")

                # Then the tree and scrollbar (ensure these are grid'ed on row 1, so they appear below the label)
                tree.grid(row=1, column=0, sticky='nsew', padx=0, pady=0)
                scrollbar.grid(row=1, column=1, sticky='nsew')
                
                # Configure the grid to take up available space
                tabView.tab("Shifts Done\nKD回数").grid_rowconfigure(0, weight=1)
                tabView.tab("Shifts Done\nKD回数").grid_rowconfigure(1, weight=2)
                tabView.tab("Shifts Done\nKD回数").grid_columnconfigure(0, weight=1)



                error_message.configure(text_color="#40ff06",text=f"Generation Complete!\nスケジュール完了！☜( Φ∀Φ )ﾖｼｯ")

                #Readjust the width to fix a bug where the bottom corners of the button disappear after reenabling it
                bWid = button_1.cget("width")
                button_1.configure(state="normal",width=bWid)
            
            # Run the new function on a separate thread
            thread = threading.Thread(target=generate_and_reenable)
            thread.start()

    customtkinter.set_appearance_mode("dark")
    customtkinter.set_default_color_theme("dark-blue") #Blue,green,darkblue

    root = customtkinter.CTk()
    root.geometry("500x420")
    root.resizable(0,0)
    root.title("Schedule Generator")

    # The following function will be used to center the window.
    def center_window():
        # Update the window to ensure it has been initialized and has the correct size.
        root.update()

        # Calculate the position to center the window.
        window_width = root.winfo_width()
        window_height = root.winfo_height()
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()

        position_top = int(screen_height / 2 - window_height / 2)
        position_right = int(screen_width / 2 - window_width / 2)

        root.geometry(f'{window_width}x{window_height}+{position_right}+{position_top}')

    def addRec():
        new_name = nameEntry.get()  # Get value from name entry
        new_room = roomEntry.get()  # Get value from room entry
        new_enro = enroOpt.get()  # Get value from option menu
        
        # Check if the fields are not empty
        if new_name and new_room and new_enro:
            # Insert a new record to treeview
            recTree.insert('', 'end', values=(new_name, new_room, new_enro))
            
            # Add this data to the json file as well
            save_to_json(recTree)
            nameEntry.delete(0, 'end')
            roomEntry.delete(0, 'end')
        else:
            messagebox.showerror("Input Error", "All fields must be filled out.\nすべて記入しなければなりません")

    
    def editRec():
        global is_editing
        global selected_id
        is_editing = not is_editing
        
        #Only when edit button is pressed and a row is selected, then it'll go into edit mode
        if is_editing and recTree.selection():
            addButton.configure(state="disabled")
            deleteButton.configure(state="disabled")
            editButton.configure(text="Save / 保存",fg_color="#1df8bf",hover_color="#1ccd9f")
            selected_id = recTree.selection()[0]  # Get the id of the selected item
            selected_values = recTree.item(selected_id, 'values')  # Get the values of the selected item
            nameEntry.delete(0, 'end')  # Clear the current text
            nameEntry.insert(0, selected_values[0])  # Insert the new text
            roomEntry.delete(0, 'end')
            roomEntry.insert(0, selected_values[1])
            enroOpt.set(selected_values[2])

        elif not is_editing and recTree.selection():
            # Here, get the updated values from the Entry boxes and OptionMenu
            updated_name = nameEntry.get()
            updated_room = roomEntry.get()
            updated_enro = enroOpt.get()
            if updated_name and updated_room and updated_enro:
                # Update the selected item in the TreeView with the new values
                recTree.item(selected_id, values=(updated_name, updated_room, updated_enro))
                addButton.configure(state="normal")
                deleteButton.configure(state="normal")
                nameEntry.delete(0, 'end')
                roomEntry.delete(0, 'end')
                editButton.configure(text="Edit / 編集",fg_color="#d7a71b",hover_color="#b4890d")
                selected_id=None
                #Save to Json file
                save_to_json(recTree)
            else:
                messagebox.showerror("Input Error", "All fields must be filled out.\nすべて記入しなければなりません")
                is_editing = not is_editing
        else:
            messagebox.showinfo("No Selection", "No record selected.\n記録を選択してください")
            is_editing = not is_editing

    def deleteRec():
        selected_id = recTree.selection()  # Get the id of the selected item
        if selected_id:
            # Confirm if the user really wants to delete this record
            response = messagebox.askquestion("Delete Record", "Are you sure you want to delete this record?\nこの記録を削除しますか？")
            if response == 'yes':
                recTree.delete(selected_id)
                
                # Remove this data from the json file as well
                save_to_json(recTree)
        else:
            messagebox.showinfo("No Selection", "No record selected.\n記録を選択してください")



    # Call the center_window function.
    center_window()
    tabView = customtkinter.CTkTabview(root)
    tabView.pack(pady=10, padx=20, fill="both", expand = True)
    
    tabView.add("Schedule\nスケジュール")
    tabView.add("Records\n記録")
    tabView.add("Shifts Done\nKD回数")
    #Start Window
    tabView.set("Schedule\nスケジュール")

    #Schedule Tab
    label = customtkinter.CTkLabel(tabView.tab("Schedule\nスケジュール"), text="KD Schedule Generator\n自動 KD 割り当て システム", font = ("Italic",24))
    label.pack(pady=10, padx=10)

    optionmenu_1 = customtkinter.CTkOptionMenu(tabView.tab("Schedule\nスケジュール"),anchor="center", values=["Jan / 1 月", "Feb / 2 月", "Mar / 3 月", "Apr / 4 月", "May / 5 月", "Jun / 6 月", "Jul / 7 月", "Aug / 8 月", "Sep / 9 月", "Oct / 10 月", "Nov / 11 月", "Dec / 12 月"])
    optionmenu_1.pack(padx=10, pady=10)
    optionmenu_1.set("Month / 月")

    optionmenu_2 = customtkinter.CTkOptionMenu(tabView.tab("Schedule\nスケジュール"),anchor="center", values=[str(year) for year in range(2023, 2034)])  # Range of years you want
    optionmenu_2.pack(padx=10, pady=10)
    optionmenu_2.set("Year / 年")

    error_message = customtkinter.CTkLabel(tabView.tab("Schedule\nスケジュール"), text="")  # Label to display error message
    error_message.pack(padx=10)

    button_1 = customtkinter.CTkButton(tabView.tab("Schedule\nスケジュール"), text="Generate Schedule\n   開始   ", command=button_callback)
    button_1.pack(pady=10, padx=50)

    button_2 = customtkinter.CTkButton(tabView.tab("Schedule\nスケジュール"), text="Change user\nユーザー変更", command=openDir,fg_color="#d7a71b",text_color="black",hover_color="#ce7e00")
    button_2.pack(pady=10, padx=10)

    smolLabel = customtkinter.CTkLabel(tabView.tab("Schedule\nスケジュール"), text="Developed By: Chun and Jojo",text_color="black",font = ("Calibri",11))
    smolLabel.pack(pady=0, padx=1)

    #Records Tab
    styleTree()
    recTree = ttk.Treeview(tabView.tab("Records\n記録"), show='headings')
    # define columns
    recTree["columns"]=("Name / 名前","Room No","Enrollment")
    # format columns
    recTree.column("Name / 名前", anchor="center", width=100)
    recTree.column("Room No", anchor="center", width=100)
    recTree.column("Enrollment", anchor="center", width=100)
    # set heading
    recTree.heading("Name / 名前", text="Name / 名前", anchor="center",command=lambda: treeview_sort_column(recTree, "Name / 名前", False))
    recTree.heading("Room No", text="Room No / 部屋番号", anchor="center",command=lambda: treeview_sort_column(recTree, "Room No", False))
    recTree.heading("Enrollment", text="Enrollment / 学期", anchor="center",command=lambda: treeview_sort_column(recTree, "Enrollment", False))
    # add data
    with open("details.json") as file:
            data = json.load(file)
    for row in data:
        recTree.insert('', 'end', values=(row["Name"], row["Room No"], row["Enrollment"]))
    # Create the scrollbar
    recScrollbar = ttk.Scrollbar(tabView.tab("Records\n記録"), orient="vertical", command=recTree.yview)
    recTree.configure(yscrollcommand=recScrollbar.set)
    # Then the tree and scrollbar (ensure these are grid'ed on row 1, so they appear below the label)
    recTree.grid(row=0, column=0, sticky='nsew', columnspan=3)  # the treeview spans 3 columns
    recScrollbar.grid(row=0, column=3, sticky='ns')
    #Labels
    namelabel = customtkinter.CTkLabel(tabView.tab("Records\n記録"), text="Name/名前", font = ("Italic",16))
    namelabel.grid(row=1, column=0, sticky='ew', padx=10)  # make it fill the grid cell
    roomlabel = customtkinter.CTkLabel(tabView.tab("Records\n記録"), text="Room No/部屋番号", font = ("Italic",16))
    roomlabel.grid(row=1, column=1, sticky='ew', padx=10)  # make it fill the grid cell
    enrolabel = customtkinter.CTkLabel(tabView.tab("Records\n記録"), text="Enrollment/学期", font = ("Italic",16))
    enrolabel.grid(row=1, column=2, sticky='ew', padx=10)  # make it fill the grid cell

    #Entry and Dropdown box
    nameEntry = customtkinter.CTkEntry(tabView.tab("Records\n記録"), width=100, corner_radius=5)
    nameEntry.grid(row=2, column=0, sticky='ew',padx=10)  # make it fill the grid cell
    # Create a Tkinter callback variable linked to our validation function
    vcmd = tabView.register(validate)
    roomEntry = customtkinter.CTkEntry(tabView.tab("Records\n記録"), width=100, corner_radius=5,validate="key", validatecommand=(vcmd, '%P'))
    roomEntry.grid(row=2, column=1, sticky='ew',padx=50)  # make it fill the grid cell
    enroOpt = customtkinter.CTkOptionMenu(tabView.tab("Records\n記録"), width=100, values=["SP","FA"],anchor="center",fg_color="#403F45",button_color="black")
    enroOpt.grid(row=2, column=2, sticky='ew',padx=10)  # make it fill the grid cell
    enroOpt.set("SP")
    #Buttons Add, Edit, Delete
    addButton = customtkinter.CTkButton(tabView.tab("Records\n記録"), text="Add / 追加", command=addRec)
    addButton.grid(row=3, column=0, sticky='ew',pady=10,padx=10)
    editButton = customtkinter.CTkButton(tabView.tab("Records\n記録"), text="Edit / 編集", command=editRec,fg_color="#d7a71b",text_color="black",hover_color="#b4890d")
    editButton.grid(row=3, column=1, sticky='ew',pady=10,padx=25)
    deleteButton = customtkinter.CTkButton(tabView.tab("Records\n記録"), text="Delete / 削除", command=deleteRec,fg_color="#c40909",text_color="white",hover_color="#86020d")
    deleteButton.grid(row=3, column=2, sticky='ew',pady=10,padx=10)

    records_tab = tabView.tab("Records\n記録")
    records_tab.grid_rowconfigure(0, weight=1)  # treeview will expand
    records_tab.grid_rowconfigure(1, weight=0)  # entry widgets won't expand
    records_tab.grid_rowconfigure(2, weight=0)  # entry widgets won't expand
    records_tab.grid_columnconfigure((0, 1, 2), weight=1)  # all columns have the same weight


    #Shifts Done Tab
    label2 = customtkinter.CTkLabel(tabView.tab("Shifts Done\nKD回数"), text="Please generate a schedule first\nスケジュールを作成してください", font = ("Italic",30),text_color="yellow")
    label2.pack(padx=10,pady=120)

    root.mainloop()

if __name__ == '__main__':
    main()