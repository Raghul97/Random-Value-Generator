from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import random
import pandas as pd


def generate_excel(source_file_path, destination_folder_path, record_count):
    filename = source_file_path
    df = pd.read_excel(filename)
    names = df['Name:'].tolist()
    emails = df['Email ID:'].tolist()
    data = list(zip(names, emails))
    data = random.sample(data, record_count)
    new_data = []
    i = 1
    for name, email in data:
        new_data.append((i, name, email))
        i += 1
    pd.DataFrame(new_data).to_excel(destination_folder_path + '/output.xlsx', header=["Number:", "Name:", "Email ID:"], index=False)


def get_source_file_path():
    source_file_path = filedialog.askopenfilename(initialdir = "/home", title = "Select a File", filetypes = [('Excel', ('*.xls', '*.xlsx'))])
    if source_file_path and ( source_file_path.endswith(".xlsx") or source_file_path.endswith(".xls") ):
        source_file.set(source_file_path)
        source_entry_error_label.config(text="")
    else:
        source_entry_error_label.config(text="select a valid excel file.")

def get_destination_folder_path():
    folder_selected_path = filedialog.askdirectory()
    if folder_selected_path:
        destination_folder.set(folder_selected_path)
        destination_entry_error_label.config(text="")
    else:
        destination_entry_error_label.config(text="select a valid folder.")

def start_process():
    try:
        record_count = int(record_entry.get())
        record_error_label.config(text="")
    except ValueError:
        record_error_label.config(text="Enter a valid number.")
    source_file_path = source_file.get()
    destination_folder_path = destination_folder.get()
    if not source_file_path:
        source_entry_error_label.config(text="select a valid excel file.")
    elif not destination_folder_path:
        destination_entry_error_label.config(text="select a valid folder.")
    else:
        destination_entry_error_label.config(text="")
        source_entry_error_label.config(text="")
        try:
            generate_excel(source_file_path, destination_folder_path, record_count)
            process_info_label.config(text="Success: file generated.")
        except:
            process_error_label.config(text="Failed: something went wrong.")


window = Tk()
window.geometry("465x200")
window.title("Randomizer")

window.minsize(465, 200)
window.maxsize(465, 200)

source_file= StringVar()
destination_folder = StringVar()

heading = Label(window, text="Generate Random values", font=("Helvetica", 16))
heading.place(x=95, y=10)

source_label = Label(window ,text="Source: ")
source_label.place(x=10, y=60)

source_entry = Entry(window, textvariable = source_file, width=30)
source_entry.place(x=95, y=60)

source_entry_error_label = Label(window, text="", fg='red', font=("Helvetica", 10))
source_entry_error_label.place(x=95, y=85)

source_button = ttk.Button(window, text="Browse Folder", command=get_source_file_path)
source_button.place(x=350, y=60)

destination_label = Label(window ,text="Destination: ")
destination_label.place(x=10, y=100)

destination_entry = Entry(window, textvariable = destination_folder, width=30)
destination_entry.place(x=95, y=100)

destination_entry_error_label = Label(window, text="", fg='red', font=("Helvetica", 10))
destination_entry_error_label.place(x=95, y=125)

destination_button = ttk.Button(window, text="Browse Folder", command=get_destination_folder_path)
destination_button.place(x=350, y=100)

record_label = Label(window ,text="Records: ")
record_label.place(x=10, y=140)

record_entry = Entry(window, width=5)
record_entry.place(x=95, y=140)

record_error_label = Label(window, text="", fg='red', font=("Helvetica", 10))
record_error_label.place(x=10, y=175)

start_button = ttk.Button(window ,text="Start", width=35, command=start_process)
start_button.place(x=162, y=140)

process_info_label = Label(window, text="", fg='green', font=("Helvetica", 10))
process_info_label.place(x=162, y=175)

process_error_label = Label(window, text="", fg='red', font=("Helvetica", 10))
process_error_label.place(x=162, y=175)

window.mainloop()