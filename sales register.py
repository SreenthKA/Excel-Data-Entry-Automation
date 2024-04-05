#!/usr/bin/env python
# coding: utf-8

# In[43]:


import numpy as np
import pandas as pd
import openpyxl
import os
from tkinter import messagebox
import tkinter as tk
from datetime import datetime
from tkinter import simpledialog


# In[44]:


dataini = {"date":[],"Bill No.":[], "Place":[],"Party Name":[],"GST Number":[],"Taxable Value":[], "CGST 9%":[],"SGST 9%":[], "Total Amount":[]}
dataframe = pd.DataFrame(dataini)
filepath = r"D:\SNU Chennai\projects\krl office\sales register.xlsx"
if not os.path.exists(filepath):
    dataframe.to_excel(r"D:\SNU Chennai\projects\krl office\sales register.xlsx")


# In[45]:


def create_excel_sheet_creator():
    def createNewSheet():
        try:
            # Read existing Excel file into a DataFrame
            excel_file_path = r"D:\SNU Chennai\projects\krl office\sales register.xlsx"
            existing_df = pd.read_excel(excel_file_path)

            # Get the sheet name from user input
            sheetname = entry_month.get().lower()

            # Create a new DataFrame with the specified columns
            datanew = {"date": [], "Bill No.": [], "Place": [], "Party Name": [], "GST Number": [],
                       "Taxable Value": [], "CGST 9%": [], "SGST 9%": [], "Total Amount": []}

            # Create a new DataFrame
            dataframenew = pd.DataFrame(datanew)

            # Write the new DataFrame to a new sheet in the existing Excel file
            with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a') as writer:
                dataframenew.to_excel(writer, sheet_name=sheetname, index=False)

            messagebox.showinfo("Success", f"Sheet '{sheetname}' created successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    # Create the main window
    window = tk.Tk()
    window.title("Excel Sheet Creator")

    # Create a StringVar to store the entered month
    var_month = tk.StringVar()
    window.geometry("400x200")

    # Create a label and pack it
    label_month = tk.Label(window, text="Enter the month:")
    label_month.pack()

    # Create an entry (text box) for user input
    entry_month = tk.Entry(window, textvariable=var_month)
    entry_month.pack()

    # Create a button to trigger sheet creation
    button_create_sheet = tk.Button(window, text="Create Sheet", command=createNewSheet)
    button_create_sheet.pack()

    # Start the GUI event loop
    window.mainloop()

# Call the function to run the GUI



# In[46]:


df = pd.DataFrame()  # Declare df as a global variable

def confirmMonth(entry_month):
    global m
    m = entry_month.get().lower()
    messagebox.showinfo("Success", f"Month '{m}' confirmed!")

def confirmNumberOfData(entry_entries):
    global n
    n = int(entry_entries.get())
    messagebox.showinfo("Success", f"Number of data entries '{n}' confirmed!")

def submitData(window, entry_date, entry_bill, entry_place, entry_party, entry_gst, entry_taxable):
    try:
        global df, m  # Reference the global df and m variables

        d = entry_date.get()
        d1 = datetime.strptime(d, "%d-%m-%Y")
        b = int(entry_bill.get())
        p = entry_place.get()
        pn = entry_party.get()
        gn = entry_gst.get()
        tv = float(entry_taxable.get())

        new_data = {'date': [d1],
                    'Bill No.': [b],
                    'Place': [p],
                    'Party Name': [pn],
                    'GST Number': [gn],
                    'Taxable Value': [tv],
                    'CGST 9%': [0.09 * tv],  
                    'SGST 9%': [0.09 * tv],  
                    'Total Amount': [tv + 2 * (0.09 * tv)]}

        new_df = pd.DataFrame(new_data)
        df = pd.concat([df, new_df], ignore_index=True)

        with pd.ExcelWriter(r"D:\SNU Chennai\projects\krl office\sales register.xlsx", engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=m, index=False)
        
        messagebox.showinfo("Success", f"Data added to sheet '{m}' successfully!")
        window.destroy()
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

def addData():
    try:
        global df, m, n  # Reference the global df, m, and n variables

        excel_file_path = r"D:\SNU Chennai\projects\krl office\sales register.xlsx"

        try:
            # Try reading the existing sheet from the Excel file
            df = pd.read_excel(excel_file_path, sheet_name=m)
        except pd.errors.EmptyDataError:
            # If the sheet does not exist, create a new DataFrame
            df = pd.DataFrame()

        for i in range(n):
            window2 = tk.Tk()
            window2.title("Data Entry")
            window2.geometry("600x600")

            label_date = tk.Label(window2, text="Date (DD-MM-YYYY):")
            label_date.pack()

            entry_date = tk.Entry(window2)
            entry_date.pack()

            label_bill = tk.Label(window2, text="Bill No.:")
            label_bill.pack()

            entry_bill = tk.Entry(window2)
            entry_bill.pack()

            label_place = tk.Label(window2, text="Place:")
            label_place.pack()

            entry_place = tk.Entry(window2)
            entry_place.pack()

            label_party = tk.Label(window2, text="Party Name:")
            label_party.pack()

            entry_party = tk.Entry(window2)
            entry_party.pack()

            label_gst = tk.Label(window2, text="GST Number:")
            label_gst.pack()

            entry_gst = tk.Entry(window2)
            entry_gst.pack()

            label_taxable = tk.Label(window2, text="Taxable value:")
            label_taxable.pack()

            entry_taxable = tk.Entry(window2)
            entry_taxable.pack()

            # Create a button to trigger data entry
            button_add_data = tk.Button(window2, text="Add Data", command=lambda: submitData(window2, entry_date, entry_bill, entry_place, entry_party, entry_gst, entry_taxable))
            button_add_data.pack()
            
            # Start the GUI event loop
            window2.mainloop()
            

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

def runDataEntryApp():
    # Create the main window
    window1 = tk.Tk()
    window1.title("Data Entry")
    window1.geometry("600x400")

    label_month = tk.Label(window1, text="Enter the month:")
    label_month.pack()

    entry_month = tk.Entry(window1)
    entry_month.pack()

    button_confirm_month = tk.Button(window1, text="Confirm month", command=lambda: confirmMonth(entry_month))
    button_confirm_month.pack()

    label_entries = tk.Label(window1, text="Number of data entries:")
    label_entries.pack()

    entry_entries = tk.Entry(window1)
    entry_entries.pack()

    button_confirm_data_entries = tk.Button(window1, text="Confirm number of data", command=lambda: confirmNumberOfData(entry_entries))
    button_confirm_data_entries.pack()

    button_open_data_entry = tk.Button(window1, text="Open Data Entry", command=addData)
    button_open_data_entry.pack()

    window1.mainloop()


# In[47]:


def close_window():
    window3.destroy()

window3 = tk.Tk()
window3.title("Enter Choice")
window3.geometry("400x200")

ch1 = tk.Button(window3, text="Create new month", command=lambda: create_excel_sheet_creator())
ch1.pack()

ch2 = tk.Button(window3, text="Enter Data", command=lambda: runDataEntryApp())
ch2.pack()

ch3 = tk.Button(window3, text="Exit", command=close_window)
ch3.pack()

window3.mainloop()


# In[ ]:




