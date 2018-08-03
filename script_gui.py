from tkinter import *
from tkinter import messagebox
import Script as script
import openpyxl
import datetime
import os

#TODO
# Write error messages
# Sheet name doesn't exist(i.e wrong name)
# Naming convention(has doubles. Is not unique. In that case number them)
# More categories than necessary OR Less Categories than necessary
# Try integer inputs

# Text File
# Too few ###, or too many ###
# What if they want signature, or other stuff

# Excel spreadsheet
# Some lines are empty. 
# Wrong email or missing email
# Future improvements
# Middle column is empty
# Column name is empty

def create_participant_dictionaries():
    # Get all the data input
    dict_data = get_entry_fields()
    # Insert the title of the spreadsheet here
    wb = openpyxl.load_workbook(dict_data["sp_title"])
    # Specify which sheet of the excel spreadsheet you need
    sheet = wb.get_sheet_by_name(dict_data["sheet_title"])
    # Specify the text file which you want to read
    file = open(dict_data["file_to_read"], "r")
    # Specify naming convention
    fields = dict_data["fields"]
    # has to be that way because the functions in Script are looking for the
    # folder in fields[0]
    # fields[0] is name of the folder i want to store the files in
    fields = ["Custom Emails"] + fields
    #fields = [file_name] + fields
    subject = dict_data["subject"]
    # list of dictionaries store the data for each participant
    list_of_dictionaries = script.dictList(sheet)
    return (file, list_of_dictionaries, fields, subject)

    


def email_checks():
    '''Input: list of dict of (subject, email, email body)'''
    (file, list_of_dictionaries, fields,
     subject) = create_participant_dictionaries()
    # Get a list which contains a specific amount of email samples
    listOfEmails = script.first_last_file(
        file, list_of_dictionaries, fields, subject)
    # onClick(listOfEmails)
    message = "First email:\n"
    message += str(listOfEmails[0]["subject"])
    message += "\n" + "To: " + str(listOfEmails[0]["email"])
    message += "\n" + str(listOfEmails[0]["email_body"]) + "\n"
    message += "-----------------------------------\nLast email:\n"
    message += str(listOfEmails[1]["subject"])
    message += "\n" + "To: " + str(listOfEmails[1]["email"])
    message += "\n" + str(listOfEmails[1]["email_body"]) + "\n"
    master2 = Tk()
    master2.title("Email examples")
    msg = Message(master2, text=message)
    msg.config(bg = "#43ABC9", fg = "#f2f2f2", font="Alice")
    msg.pack()


def send_out_emails():
    result = messagebox.askquestion(
        "Send all Emails", "Are You Sure?", icon='warning')
    if result == 'yes':
        (file, list_of_dictionaries, fields,
         subject) = create_participant_dictionaries()
        script.send_text_files(file, list_of_dictionaries, fields, subject)
        master2 = Tk()
        message = "Congratulations!\nAll emails were sent successfully\n\n\n"
        msg = Message(master2, text=message)
        msg.config(pady=5)
        msg.pack()


def save_file():
    file = open("email_distribution.txt", "w+")
    file.write("sp_title: " + str(sp_title.get()) + "\n")
    file.write("sheet_title: " + str(sheet_title.get()) + "\n")
    file.write("file_to_read: " + str(file_to_read.get()) + "\n")
    file.write("fields: " + str(fields.get()) + "\n")
    file.write("subject: " + str(subject.get()) + "\n")
    file.close()


def load_file():
    file = open("email_distribution.txt", "r")
    # Create a dictionary where we'll store the type of entry as the key
    # and its text as the value
    text = {}
    for line in file:
        index_column = line.find(":")
        key = line[:index_column]
        value = line[index_column + 2:].rstrip()
        text[key] = value
    entry_types = [sp_title, sheet_title, file_to_read, fields, subject]
    value_types = ["sp_title", "sheet_title", "file_to_read", "fields", "subject"]
    counter = 0
    for e in entry_types:
        e.delete(0, END)
        e.insert(0, text[value_types[counter]])
        counter += 1


def get_entry_fields():
    dict_data = {}
    dict_data["sp_title"] = str(sp_title.get()) + ".xlsx"
    dict_data["sheet_title"] = str(sheet_title.get())
    dict_data["file_to_read"] = str(file_to_read.get()) + ".txt"
    dict_data["subject"] = str(subject.get())
    # Convert input from fields to list
    fields_str = fields.get()
    fields_list = fields_str.split(", ")
    dict_data["fields"] = fields_list
    return dict_data

def focus_set(event):
    if type(event.widget) is not Button:
        event.widget.focus()
master = Tk()
master.title("Email Distribution")
master.configure(background='#3E50B4')
#master.resizable(width=False, height=False)
master.bind_all("<1>", focus_set)

# Create labels
l1 = Label(master, text="Excel file(*xlsx):")
l1.grid(row=0)
l2 = Label(master, text="Sheet name:")
l2.grid(row=1)
l3 = Label(master, text="E-mail body(*txt):")
l3.grid(row=2)
l4 = Label(master, text="Categories to fill:")
l4.grid(row=3)
l5 = Label(master, text="E-mail Subject:")
l5.grid(row=4)

# Edit the labels.
label_list = [l1, l2, l3, l4, l5]
for label in label_list:
    label.grid(sticky=E)
    label.config(bg='#3E50B4', fg="#FCFCFC", font="Raven, 9")



def track_change_to_text(event):
    event.widget.config(background="#FFFF66")

def track_focus_out(event):
    event.widget.config(background="white")

# create entries
sp_title = Entry(master)
sheet_title = Entry(master)
file_to_read = Entry(master)
fields = Entry(master)
subject = Entry(master)

# Edit the entries
entry_types = [sp_title, sheet_title,
               file_to_read, fields, subject]
for entry in entry_types:
    entry.config(font = "Raven, 9")
    entry.grid(pady=2, padx=5)
    #entry.grid_columnconfigure(0,weight=1)
    entry.bind('<FocusIn>', track_change_to_text)
    entry.bind('<FocusOut>', track_focus_out) 

# Assign entries to labels
sp_title.grid(row=0, column=1)
sheet_title.grid(row=1, column=1)
file_to_read.grid(row=2, column=1)
fields.grid(row=3, column=1)
subject.grid(row=4, column=1)

# Create the buttons
#b1 = Button(master, text="Create files", command=just_files)
#b1.grid(row=9, column=0)
b1 = Button(master, text="See examples",
            command=email_checks)
b1.grid(row=9, column=0)
b2 = Button(master, text="Send all Emails",
            command=send_out_emails)
b2.grid(row=10, column=0)
b3 = Button(master, text='Quit', command=master.quit,width = 8)
b3.grid(row=12, column=1, pady=5, sticky = E)
b4 = Button(master, text="Save form", command=save_file)
b4.grid(row=11, column=0)
b5 = Button(master, text="Load form", command=load_file)
b5.grid(row=12, column=0)

# Make button appear pressed, and then return to normal
def button_press(event):
    event.widget.config(relief=SUNKEN, activebackground="#A9206D")


def button_release(event):
    event.widget.config(relief=RAISED, activebackground="#A9206D")
# Darken/return to normal when user moves over button


def track_change_to_button(event):
    event.widget.config(background="#811853")


def button_focus_out(event):
    event.widget.config(background="#A9206D")
# Change colour when focus on them, return to normal when focus out
buttons_list = [b1, b2, b3, b4, b5]
button_list_no_quit= [b1, b2, b4, b5]
for button in buttons_list:
    button.config(bg = "#A9206D", fg = "#FCFCFC", font = "Raven, 9")
    button.grid(padx = 5, pady=2)
    button.bind('<Enter>', track_change_to_button)
    button.bind('<Leave>', button_focus_out)
    button.bind("<ButtonPress-1>", button_press)
    button.bind("<ButtonRelease-1>", button_release)
# Don't want to stretch the "Quit" Button in the whole cell
for button in button_list_no_quit:
    button.grid(sticky=N+S+E+W, padx = 5, pady=2)
mainloop()
