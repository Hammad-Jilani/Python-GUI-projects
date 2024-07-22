import tkinter as tk
from tkinter import ttk
import openpyxl

window = tk.Tk()

style = ttk.Style(window)

def loadData():
    path = "C:\\Users\CoreCom\PycharmProjects\excelSheet\main\people.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active

    list_values = list(sheet.values)
    for col_name in list_values[0]:
        treeview.heading(col_name,text=col_name)

    for value_tuple in list_values[1:]:
        treeview.insert('',tk.END,values=value_tuple)


window.tk.call("source", "forest-light.tcl")
window.tk.call("source", "forest-dark.tcl")
style.theme_use("forest-dark")

frame = ttk.Frame(window)
frame.pack()

widgets_frame = ttk.LabelFrame(frame, text="Insert Row")
widgets_frame.grid(row=0, column=0, padx=20, pady=10)

name_entry = ttk.Entry(widgets_frame, width=30)
name_entry.insert(0, "Name")
name_entry.bind("<FocusIn>", lambda e: name_entry.delete('0', 'end'))
name_entry.grid(row=0, column=0, sticky="ew", padx=5, pady=(0, 5))

age_spinbox = ttk.Spinbox(widgets_frame, from_=18, to=100, width=30)
age_spinbox.insert(0, "Age")
age_spinbox.bind("<FocusIn>", lambda e: age_spinbox.delete('0', 'end'))
age_spinbox.grid(row=1, column=0, sticky="ew", padx=5, pady=(0, 5))

# ComboBox
comboList = ["Subscribed", "Not Subscribed", "Other"]
subscribe = ttk.Combobox(widgets_frame, width=30, value=comboList)
subscribe.current(0)
subscribe.grid(row=2, column=0, sticky="ew", padx=5, pady=(0, 5))

# Check Box
a = tk.BooleanVar()
checkButton = ttk.Checkbutton(widgets_frame, text="Employed", width=30, variable=a)
checkButton.grid(row=3, column=0, sticky="nsew", padx=5, pady=(0, 5))

def insert_row():
    name=name_entry.get()
    age=int(age_spinbox.get())
    subscribe_status=subscribe.get()
    employeeStatus = "Employed" if a.get() else "Unemployed"
    print(name,age,subscribe_status,employeeStatus)

    # Insert row into Excel sheet
    path = "C:\\Users\CoreCom\PycharmProjects\excelSheet\main\people.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    row_values = [name,age,subscribe_status,employeeStatus]
    sheet.append(row_values)
    workbook.save(path)

    # Insert row into treeview
    treeview.insert('',tk.END,values=row_values)

    name_entry.delete(0,"end")
    name_entry.insert(0,"Name")
    age_spinbox.delete(0,"end")
    age_spinbox.insert(0,"Age")
    subscribe.set(comboList[0])
    checkButton.state(["!selected"])


insertButton = ttk.Button(widgets_frame, text="Insert", command=insert_row)
insertButton.grid(row=4, column=0, sticky="ew", padx=5, pady=(0, 5))

separator = ttk.Separator(widgets_frame)
separator.grid(row=5, column=0, padx=(20, 10), pady=10, sticky="ew")


def toggle_mode():
    if mode_switch.instate(["selected"]):
        style.theme_use("forest-light")
    else:
        style.theme_use("forest-dark")


mode_switch = ttk.Checkbutton(widgets_frame, text="Mode", style='Switch', command=toggle_mode)
mode_switch.grid(row=6, column=0, padx=(20, 10), pady=10, sticky="ew")

treeFrame = ttk.Frame(frame)
treeFrame.grid(row=0, column=1, pady=10)

treeScroll = ttk.Scrollbar(treeFrame)
treeScroll.pack(side="right", fill='y')

cols = ("Name", "Age", "Subscription", "Employment")
treeview = ttk.Treeview(treeFrame, show="headings", columns=cols, height=13, yscrollcommand=treeScroll.set)
treeview.column("Name", width=100)
treeview.column("Age", width=50)
treeview.column("Subscription", width=100)
treeview.column("Employment", width=100)
treeview.pack()
treeScroll.config(command=treeview.yview)

loadData()

window.mainloop()
