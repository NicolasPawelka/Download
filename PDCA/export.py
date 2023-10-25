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
from pyxll import xl_func, xl_macro
############################
DEMO = False
TASK_NAME = "Arbeitsergebnis"
START_DATE = None
END_DATE = "Zieldatum"
DATA_FRAME = None
EXCEL_FILE_PATH = None
PROJECT = None
ACTIVE_PROJECT = None
SELECTED_SHEET = []
TASKS = None
PROJECT_FILE_PATH = None
WAS_SUMMARY = False
RESOURCES = None
ID = []
############################
def is_summary(current_depth, saved_depth, next_depth):
    if current_depth > saved_depth and current_depth > next_depth:
        return True
    elif current_depth < next_depth:
        return True
    elif current_depth > saved_depth and current_depth == next_depth:
        return True
    return False

def find_existing_task_by_name(name, TASKS):
    for task in TASKS:
        if task.Name == name:
            return task
    return None



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
        
        #Popup Dialog Fenster öffnen um das sheet auszuwählen
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
    if isinstance(text,int):
        return text
    elif text is None:
        return 0
    
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
        pass
    return 1

def add_Summary(TASKS, name,depth,date,budget):
    task = TASKS.Add()
    task.Manual = False
    task.Name = name
    task.Start = date
    task.OutlineLevel = depth
    task.Cost = budget
    return 1

def add_resource(list, resources):
    for resource_name in list:
        resources.Add(resource_name)

  
def init(project_file_path):
    global DATA_FRAME
    global EXCEL_FILE_PATH
    global PROJECT
    global SELECTED_SHEET
    global ACTIVE_PROJECT
    global TASKS
    global PROJECT_FILE_PATH
    global ID
    
    EXCEL_FILE_PATH = choose_excel_file()
    SELECTED_SHEET = choose_excel_sheet(EXCEL_FILE_PATH)
    
    PROJECT = win32.Dispatch("MSProject.Application")
    PROJECT_FILE_PATH = project_file_path
    PROJECT.FileOpen(project_file_path)
    
    workbook = load_workbook(EXCEL_FILE_PATH)
    worksheet = workbook[SELECTED_SHEET[0]]
    
    DATA_FRAME = pd.DataFrame(worksheet.values)
    DATA_FRAME = DATA_FRAME.reset_index()
    
    ACTIVE_PROJECT = PROJECT.ActiveProject
    TASKS = ACTIVE_PROJECT.Tasks
    
    ID_index = find_ID(DATA_FRAME)
    last_value = None  # Initialize last_value to None
    for _, row in DATA_FRAME.iterrows():
        if row.iloc[ID_index] is not None:
            last_value = row.iloc[ID_index]
            ID.append(last_value)

    if last_value is not None:
        ID.append(last_value)

               
        
    main()

def update():
    messagebox.showinfo("Hallo")
    global DATA_FRAME
    global EXCEL_FILE_PATH
    global PROJECT
    global ACTIVE_PROJECT
    global TASKS

    if not DATA_FRAME or not EXCEL_FILE_PATH or not PROJECT or not ACTIVE_PROJECT:
        messagebox.showerror("Error", "Data is not properly initialized. Please run 'init' first.")
        return

    Task_Name_index = find_TASK_NAME(DATA_FRAME, TASK_NAME)
    ID_index = find_ID(DATA_FRAME)
    START_index = find_START(DATA_FRAME)
    BUDGET_index = find_BUDGET(DATA_FRAME)
    RESOURCE_index = find_RESOURCE(DATA_FRAME)

    if Task_Name_index == -1 or ID_index == -1 or START_index == -1 or BUDGET_index == -1 or RESOURCE_index == -1:
        messagebox.showerror("Error", "Could not find required columns in the Excel data.")
        return

    current_index = 1
    for _, row in DATA_FRAME.iterrows():
        current_name = row.iloc[Task_Name_index]
        current_id = row.iloc[ID_index]
        current_budget = row.iloc[BUDGET_index]
        if current_name is None or current_name == TASK_NAME:
            continue
        else:
            current_depth = calculate_depth(current_id)
            date = row.iloc[START_index]
            if date is None:
                date = datetime.now().strftime("%d.%m.%Y")
            task = find_existing_task_by_name(current_name, TASKS)
            if task:
                task.Start = date
                task.OutlineLevel = current_depth
                task.Cost = extract_budget(current_budget)
            else:
                messagebox.showwarning("Warning", f"Task '{current_name}' not found in the existing project. Skipping.")
            current_index += 1

    messagebox.showinfo("Update Complete", "Existing tasks have been updated.")


def main():
    global ID
    current_index = 1
    saved_depth = 0
    task_number = 0
    first_summary = True
    Task_Name_index = find_TASK_NAME(DATA_FRAME, TASK_NAME)
    ID_index = find_ID(DATA_FRAME)
    START_index = find_START(DATA_FRAME)
    BUDGET_index = find_BUDGET(DATA_FRAME)
    RESOURCE_index = find_RESOURCE(DATA_FRAME)
    if Task_Name_index == -1 or ID_index == -1 or START_index == -1 or BUDGET_index == -1 or RESOURCE_index == -1:
        messagebox.showerror("Error", "Could not find Column")
        sys.exit()
    
        
    for _,row in DATA_FRAME.iterrows():
        global TASKS
        global WAS_SUMMARY

        current_name = row.iloc[Task_Name_index]
        current_id = row.iloc[ID_index]
        current_budget = row.iloc[BUDGET_index]
        if current_name is None or current_name == TASK_NAME:
            continue
        else:
            current_depth = calculate_depth(current_id)
            next_depth = calculate_depth(ID[current_index + 1])
            current_index += 1
            if first_summary:
                task_number += 1
                START_DATE = simpledialog.askstring("Start festlegen",f"Bitte legen sie ein Start für {current_name} fest")
                first_summary = False
                add_Summary(TASKS,current_name,current_depth,START_DATE,extract_budget(current_budget))
                WAS_SUMMARY = True
                saved_depth = current_depth
            else:
                date = row.iloc[START_index]
                if date == None:
                    date = datetime.now().strftime("%d.%m.%Y")
                if is_summary(current_depth,saved_depth,next_depth) :
                    add_Summary(TASKS,current_name,current_depth,date,extract_budget(current_budget))
                    task_number += 1
                    WAS_SUMMARY = True
                    saved_depth = current_depth
                else:
              
                        add_Task(TASKS,current_name,current_depth,date,extract_budget(current_budget),task_number)
                        task_number += 1
                        saved_depth = current_depth
            
                     
    #project.FileSave()
    TASKS = None
    messagebox.showinfo("Completed!","Import der Daten erfolgreich")

if __name__ == "__main__":
    if DEMO is False:
        if len(sys.argv) != 2:
            messagebox.showerror("Error","Fehler bei der Ermittlung des Commands")
            sys.exit()
        else:
            if sys.argv[1] == "update":
                update()
            else:
                mpp_file_path = sys.argv[1]
                init(mpp_file_path)
    else:
        init(r"C:\Users\npawelka\Desktop\Beispiel.mpp")