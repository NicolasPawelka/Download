import win32com.client as win32
from win32com.client import constants as pjconstants
import pandas as pd
import pytz
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
import math
from openpyxl import load_workbook
import re
import sys
import os
def find_TASK_NAME(df,TASK_NAME):
    num_columns = df.shape[1]
    for i in range(num_columns):
        if df.iloc[0,i] == TASK_NAME:
            return i
    return -1

def find_ID(df):
    num_columns = df.shape[1]
    for i in range(num_columns):
        if df.iloc[0,i] == "ID":
            return i
    return -1

def find_START(df):
    num_columns = df.shape[1]
    for i in range(num_columns):
        if df.iloc[1,i] == "Startdatum":
            return i
    return -1

def find_BUDGET(df):
    num_columns = df.shape[1]
    for i in range(num_columns):
        if df.iloc[1,i] == "Geplant":
            return i
    return -1    
 
    
def find_RESOURCE(df):
    num_columns = df.shape[1]
    for i in range(num_columns):
        if df.iloc[0,i] == "Bearbeiter":
            return i
    return -1
    
def choose_excel_file():
    root = tk.Tk()
    root.withdraw() # Vertsecke das Hauptfenster
    
    file_path = filedialog.askopenfilename(filetypes=[("Excel-Dateien", "*.xlsx *.xls *.xlsm")])
    
    return file_path
    
    
def choose_excel_sheet(file_path):
    if file_path:
        xls = pd.ExcelFile(file_path)
        sheet_names = xls.sheet_names
        
        #Popup Dialog Fenster erstellen um das sheet auszuwählen
        options = ["Load all sheets"]
        options.extend([f"Load sheet {i + 1}: {sheet_name}" for i, sheet_name in enumerate(sheet_names)])
        
        choice = simpledialog.askinteger("Auswahl der Arbeitsmappe", "Waählen Sie eine Arbeitsmappe die geladen werden soll: ",
                                         initialvalue=0, minvalue=0, maxvalue=len(options)-1)
        
        choice += 1
        selected_sheets = []
        if choice == 0:
            selected_sheets = sheet_names
        elif 1 <= choice < len(options):
            selected_sheets = [sheet_names[choice-1]]
            
        return selected_sheets
    return 0

    
def calculate_depth(text):
    numbers = re.findall(r'\.', text)
    return len(numbers)

def extract_budget(text):
    pattern = r'\*(\d+)'
    match = re.search(pattern,text)
    
    if match:
        extrcted_number = int(match.group(1))
        return extrcted_number
    return -1

def add_Task(TASKS,name,depth,date,budget,vorgänger):
    task = TASKS.Add()
    task.Manual = False
    task.Name = name
    task.Start = date
    task.OutlineLevel = depth
    task.Cost = budget
    if vorgänger != -1:
        task.Predecessors = vorgänger
    return 1

def add_Summary(TASKS, name,depth,date):
    task = TASKS.Add()
    task.Manual = False
    task.Name = name
    task.Start = date
    task.OutlineLevel = depth
    return 1

def add_resource(list, resources):
    for resource_name in list:
        resources.Add(resource_name)
