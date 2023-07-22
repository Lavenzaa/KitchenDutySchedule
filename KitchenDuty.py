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
#declare Constants
CREDENTIALS_PATH = 'token.pickle'
CLIENT_SECRET_PATH = 'client_secret.json'
SCOPES = ['https://www.googleapis.com/auth/calendar']


#Function to get all possible pairs in a list
def get_pairs(lst):
    pairs = [(lst[i], lst[j]) for i in range(len(lst)) for j in range(i+1, len(lst))]
    return pairs

#Function for creating events, parameters needed are Resident Assistant 1 and 2, year, month and day
def event_dict(RA1,color,year,month,day):
    event = {
        'summary': RA1,
        'description': 'Kitchen Duty',
        'colorId' : int(color),
        'start': {'date': str(year)+'-'+ str(month) +'-'+str(day),},
        'end': {'date': str(year)+'-'+ str(month) +'-'+str(day),},
    }
    return event

#Used to save student records to the details.json file
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

#Check that number being inputted is a digit
def validate(input):
    # check if input is a digit
    if input.isdigit():
        return True
    elif input == "":  # Allow empty input
        return True
    else:
        return False

#Style the tree/table
def style_tree():
        # Create the style
        style = ttk.Style()
        style.theme_use("clam")

        # Configure the Treeview style
        style.configure("Treeview",
                        background="black",
                        fieldbackground="black",
                        foreground="white")



#If you see this line hemlo, don't tell ;> [P.s. GOD BLESS TANEZAKI ATSUMIII]
        # Configure the Treeview heading style
        style.configure("Treeview.Heading",
                        background="black",
                        foreground="white")
        #Hover style
        style.map('Treeview.Heading',
                    background=[('active', 'grey')],
                    foreground=[('active', 'white')])

#Used to allow sorting of tree by clicking on heading
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

#Convert the selected month to int
def month_to_int(month):
        month_dict = {"Jan / 1 月": 1, "Feb / 2 月": 2, "Mar / 3 月": 3, "Apr / 4 月": 4, "May / 5 月": 5, "Jun / 6 月": 6, "Jul / 7 月": 7, "Aug / 8 月": 8, "Sep / 9 月": 9, "Oct / 10 月": 10, "Nov / 11 月": 11, "Dec / 12 月": 12}
        return month_dict[month]

#Opens folder where program is located to allow user to change google account
def open_dir():
        # Ask for confirmation before changing the user
        response = messagebox.askyesno("Confirmation",
                                    "Are you sure you want to change the user?\nユーザー変更てもよろしいですか？")
        if response == True:
            # Delete token, doing this requires user to log in again
            if os.path.exists(CREDENTIALS_PATH):
                os.remove(CREDENTIALS_PATH)
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

#Used to check if response is received
def callback(request_id, response, exception):
    if exception is not None:
        print(exception)
    else:
        print("Event created: " + response['summary'])  

# Used to center the window.
def center_window(root):
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

#adds inputted record and saves it to json
def add_rec(NE,RE,EO,RT):
        new_name = NE.get()  # Get value from name entry
        new_room = RE.get()  # Get value from room entry
        new_enro = EO.get()  # Get value from option menu
        
        # Check if the fields are not empty
        if new_name and new_room and new_enro:
            # Insert a new record to treeview
            RT.insert('', 'end', values=(new_name, new_room, new_enro))
            
            # Add this data to the json file as well
            save_to_json(RT)
            NE.delete(0, 'end')
            RE.delete(0, 'end')
        else:
            messagebox.showerror("Input Error", "All fields must be filled out.\nすべて記入しなければなりません")

#Turn on editing mode
def edit_mode_on(addBut,delBut,editBut,RT,NE,RE,EO):
    global selected_id
    addBut.configure(state="disabled")
    delBut.configure(state="disabled")
    editBut.configure(text="Save / 保存",fg_color="#1df8bf",hover_color="#1ccd9f")
    selected_id = RT.selection()[0]  # Get the id of the selected item
    selected_values = RT.item(selected_id, 'values')  # Get the values of the selected item
    NE.delete(0, 'end')  # Clear the current text
    NE.insert(0, selected_values[0])  # Insert the new text
    RE.delete(0, 'end')
    RE.insert(0, selected_values[1])
    EO.set(selected_values[2])

#Turn off editing mode
def edit_mode_off(addBut,delBut,editBut,RT,NE,RE,EO):
    global selected_id
    # Here, get the updated values from the Entry boxes and OptionMenu
    updated_name = NE.get()
    updated_room = RE.get()
    updated_enro = EO.get()
    if updated_name and updated_room and updated_enro:
        # Update the selected item in the TreeView with the new values
        RT.item(selected_id, values=(updated_name, updated_room, updated_enro))
        addBut.configure(state="normal")
        delBut.configure(state="normal")
        NE.delete(0, 'end')
        RE.delete(0, 'end')
        editBut.configure(text="Edit / 編集",fg_color="#d7a71b",hover_color="#b4890d")
        selected_id=None
        #Save to Json file
        save_to_json(RT)
    else:
        messagebox.showerror("Input Error", "All fields must be filled out.\nすべて記入しなければなりません")
        
#For editing records in records tab
def edit_rec(addBut,delBut,editBut,RT,NE,RE,EO):
    global is_editing
    is_editing = not is_editing
    
    #Only when edit button is pressed and a row is selected, then it'll go into edit mode
    if is_editing and RT.selection():
        edit_mode_on(addBut,delBut,editBut,RT,NE,RE,EO)
    elif not is_editing and RT.selection():
        edit_mode_off(addBut,delBut,editBut,RT,NE,RE,EO)
    else:
        messagebox.showinfo("No Selection", "No record selected.\n記録を選択してください")
        is_editing = not is_editing

#For deleteing a specific record
def delete_rec(RT):
        selected_id = RT.selection()  # Get the id of the selected item
        if selected_id:
            # Confirm if the user really wants to delete this record
            response = messagebox.askquestion("Delete Record", "Are you sure you want to delete this record?\nこの記録を削除しますか？")
            if response == 'yes':
                RT.delete(selected_id)
                
                # Remove this data from the json file as well
                save_to_json(RT)
        else:
            messagebox.showinfo("No Selection", "No record selected.\n記録を選択してください")

# global variable to keep track of editing state
is_editing = False
selected_id=None

#Generates schedule
def gen_schedule(year,month):
    #Verification stuff
    creds = None
    if os.path.exists(CREDENTIALS_PATH):
        with open(CREDENTIALS_PATH, 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                CLIENT_SECRET_PATH, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(CREDENTIALS_PATH, 'wb') as token:
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

    #Get days in a month to determine number of iterations, function returns a tuple so [1] needed
    days = calendar.monthrange(year, month)[1]
    #List for storing all unused pairs
    Pairs = get_pairs(namesList)
    #List for storing USABLE pairs
    ValidPairs = get_pairs(namesList)
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
            event_dict(residentAssistant1[0],residentAssistant1[1],year,month,i),
            event_dict(residentAssistant2[0],residentAssistant2[1],year,month,i)
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

# Actual GenerateSchedule function and re-enables the button
def generate_and_reenable(month,year,message,sBut,tLab,tabView):
    shiftDict = gen_schedule(int(year),month_to_int(month))
    #Destroy original label
    tLab.destroy()
    #Display shifts done in Shifts tab of gui
    style_tree()
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

    message.configure(text_color="#40ff06",text=f"Generation Complete!\nスケジュール完了！☜( Φ∀Φ )ﾖｼｯ")

    #Readjust the width to fix a bug where the bottom corners of the button disappear after reenabling it
    bWid = sBut.cget("width")
    sBut.configure(state="normal",width=bWid)

#For error handling
def generate_duty(month,year,RT,message,sBut,tLab,tabView):
        if month == "Month / 月" or year == "Year / 年":
            message.configure(text="Please select a month and a year\n日付を入力してくだだい （#^ω^）", text_color="#fa3c58")
        elif len(RT.get_children()) < 3:
            message.configure(text="Please add at least 3 records\n最低20名を記入してっください", text_color="#fa3c58")
        else:
            sBut.configure(state="disabled")
            message.configure(text_color="#00FFD3",text=f"Generating Schedule, Please wait... \n スケジュール作成中… (´・ω・`)")
    
            # Run the new function on a separate thread
            thread = threading.Thread(target=lambda:generate_and_reenable(month,year,message,sBut,tLab,tabView))
            thread.start()


def main():
    customtkinter.set_appearance_mode("dark")
    customtkinter.set_default_color_theme("dark-blue") #Blue,green,darkblue
    root = customtkinter.CTk()
    root.geometry("500x420")
    root.resizable(0,0)
    root.title("Schedule Generator")

    # Call the center_window function.
    center_window(root)
    tab_view = customtkinter.CTkTabview(root)
    tab_view.pack(pady=10, padx=20, fill="both", expand = True)
    tab_view.add("Schedule\nスケジュール")
    tab_view.add("Records\n記録")
    tab_view.add("Shifts Done\nKD回数")
    tab_view.set("Schedule\nスケジュール")

    #Schedule Tab components
    title_label = customtkinter.CTkLabel(tab_view.tab("Schedule\nスケジュール"), text="KD Schedule Generator\n自動 KD 割り当て システム", font = ("Italic",24))
    title_label.pack(pady=10, padx=10)

    month_option = customtkinter.CTkOptionMenu(tab_view.tab("Schedule\nスケジュール"),anchor="center", values=["Jan / 1 月", "Feb / 2 月", "Mar / 3 月", "Apr / 4 月", "May / 5 月", "Jun / 6 月", "Jul / 7 月", "Aug / 8 月", "Sep / 9 月", "Oct / 10 月", "Nov / 11 月", "Dec / 12 月"])
    month_option.pack(padx=10, pady=10)
    month_option.set("Month / 月")

    year_option = customtkinter.CTkOptionMenu(tab_view.tab("Schedule\nスケジュール"),anchor="center", values=[str(year) for year in range(2023, 2034)])  # Range of years you want
    year_option.pack(padx=10, pady=10)
    year_option.set("Year / 年")

    info_message = customtkinter.CTkLabel(tab_view.tab("Schedule\nスケジュール"), text="")  # Label to display info messages
    info_message.pack(padx=10)

    sched_button = customtkinter.CTkButton(tab_view.tab("Schedule\nスケジュール"), text="Generate Schedule\n   開始   ", command=lambda: generate_duty(month_option.get(),year_option.get(),record_tree,info_message,sched_button,tab_label,tab_view))
    sched_button.pack(pady=10, padx=50)

    change_button = customtkinter.CTkButton(tab_view.tab("Schedule\nスケジュール"), text="Change user\nユーザー変更", command=open_dir,fg_color="#d7a71b",text_color="black",hover_color="#ce7e00")
    change_button.pack(pady=10, padx=10)

    smol_label = customtkinter.CTkLabel(tab_view.tab("Schedule\nスケジュール"), text="Developed By: Chun and Jojo",text_color="black",font = ("Calibri",11))
    smol_label.pack(pady=0, padx=1)


    #Records Tab components
    style_tree()
    record_tree = ttk.Treeview(tab_view.tab("Records\n記録"), show='headings')
    # define columns
    record_tree["columns"]=("Name / 名前","Room No","Enrollment")
    # format columns
    record_tree.column("Name / 名前", anchor="center", width=100)
    record_tree.column("Room No", anchor="center", width=100)
    record_tree.column("Enrollment", anchor="center", width=100)
    # set heading
    record_tree.heading("Name / 名前", text="Name / 名前", anchor="center",command=lambda: treeview_sort_column(record_tree, "Name / 名前", False))
    record_tree.heading("Room No", text="Room No / 部屋番号", anchor="center",command=lambda: treeview_sort_column(record_tree, "Room No", False))
    record_tree.heading("Enrollment", text="Enrollment / 学期", anchor="center",command=lambda: treeview_sort_column(record_tree, "Enrollment", False))
    # add data
    with open("details.json") as file:
            data = json.load(file)
    for row in data:
        record_tree.insert('', 'end', values=(row["Name"], row["Room No"], row["Enrollment"]))
    # Create the scrollbar
    rec_scrollbar = ttk.Scrollbar(tab_view.tab("Records\n記録"), orient="vertical", command=record_tree.yview)
    record_tree.configure(yscrollcommand=rec_scrollbar.set)
    # Tree and scrollbar (ensure these are grid'ed on row 1, so they appear below the label)
    record_tree.grid(row=0, column=0, sticky='nsew', columnspan=3)  # the treeview spans 3 columns
    rec_scrollbar.grid(row=0, column=3, sticky='ns')
    #Labels
    name_label = customtkinter.CTkLabel(tab_view.tab("Records\n記録"), text="Name/名前", font = ("Italic",16))
    name_label.grid(row=1, column=0, sticky='ew', padx=10)  # make it fill the grid cell
    room_label = customtkinter.CTkLabel(tab_view.tab("Records\n記録"), text="Room No/部屋番号", font = ("Italic",16))
    room_label.grid(row=1, column=1, sticky='ew', padx=10)  # make it fill the grid cell
    enro_label = customtkinter.CTkLabel(tab_view.tab("Records\n記録"), text="Enrollment/学期", font = ("Italic",16))
    enro_label.grid(row=1, column=2, sticky='ew', padx=10)  # make it fill the grid cell
    #Entry and Dropdown box
    name_entry = customtkinter.CTkEntry(tab_view.tab("Records\n記録"), width=100, corner_radius=5)
    name_entry.grid(row=2, column=0, sticky='ew',padx=10)  # make it fill the grid cell
    # Create a Tkinter callback variable linked to our validation function
    vcmd = tab_view.register(validate)
    room_entry = customtkinter.CTkEntry(tab_view.tab("Records\n記録"), width=100, corner_radius=5,validate="key", validatecommand=(vcmd, '%P'))
    room_entry.grid(row=2, column=1, sticky='ew',padx=50)  # make it fill the grid cell
    enro_option = customtkinter.CTkOptionMenu(tab_view.tab("Records\n記録"), width=100, values=["SP","FA"],anchor="center",fg_color="#403F45",button_color="black")
    enro_option.grid(row=2, column=2, sticky='ew',padx=10)  # make it fill the grid cell
    enro_option.set("SP")
    #Buttons Add, Edit, Delete
    add_button = customtkinter.CTkButton(tab_view.tab("Records\n記録"), text="Add / 追加", command=lambda: add_rec(name_entry,room_entry,enro_option,record_tree))
    add_button.grid(row=3, column=0, sticky='ew',pady=10,padx=10)
    edit_button = customtkinter.CTkButton(tab_view.tab("Records\n記録"), text="Edit / 編集", command=lambda: edit_rec(add_button,del_button,edit_button,record_tree,name_entry,room_entry,enro_option),fg_color="#d7a71b",text_color="black",hover_color="#b4890d")
    edit_button.grid(row=3, column=1, sticky='ew',pady=10,padx=25)
    del_button = customtkinter.CTkButton(tab_view.tab("Records\n記録"), text="Delete / 削除", command=lambda: delete_rec(record_tree),fg_color="#c40909",text_color="white",hover_color="#86020d")
    del_button.grid(row=3, column=2, sticky='ew',pady=10,padx=10)
    #Aligning components in records tab
    records_tab = tab_view.tab("Records\n記録")
    records_tab.grid_rowconfigure(0, weight=1)  # treeview will expand
    records_tab.grid_rowconfigure(1, weight=0)  # entry widgets won't expand
    records_tab.grid_rowconfigure(2, weight=0)  # entry widgets won't expand
    records_tab.grid_columnconfigure((0, 1, 2), weight=1)  # all columns have the same weight


    #Shifts Done Tab components
    tab_label = customtkinter.CTkLabel(tab_view.tab("Shifts Done\nKD回数"), text="Please generate a schedule first\nスケジュールを作成してください", font = ("Italic",30),text_color="yellow")
    tab_label.pack(padx=10,pady=120)

    root.mainloop()

if __name__ == '__main__':
    main()