import win32com.client as win32
from win32com.client import constants as pjconstants
import pandas as pd
import pytz
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
import math
from openpyxl import load_workbook
from openpyxl.styles import Font,Color
import re
import sys
import os
from pyxll import xl_func, xl_macro
##########################
EXCEL = None
PROJECT = None
EXCEL_FILE_PATH = None
ACTIVE_PROJECT = None
TASKS = None
WORKBOOK = None
WORKSHEET = None
SHEET_NAME = None
OFFSET = None
COLOR_CODE = "D7E6F4"
#########################
def ersetze_uebereinstimmung(matchobj):
    global OFFSET
    global COLOR_CODE
    zahl = int(matchobj.group(2)) + OFFSET
    return matchobj.group(1) + str(zahl)


def add_task(task,position):
    global WORKSHEET
    global OFFSET
    prefix = get_positiion()
    name_index = find_ID("Arbeitsergebnis")
    id_index = find_ID("ID")
    start_index = find_Index("Startdatum")
    ziel_index = find_Index("Zieldatum")
    OFFSET = position - 8
    
    
    WORKSHEET.Cells(position,name_index).Value = task.Name
    WORKSHEET.Cells(position, id_index).Value = f"{prefix}.{task.OutlineNumber}"
    WORKSHEET.Cells(position,start_index).Value = task.Start
    WORKSHEET.Cells(position, ziel_index).ColumnWidth = len(str(task.Finish))
    WORKSHEET.Cells(position, ziel_index).Value = task.Finish 
    
    #Extrahieren der Formel zur Berechnung des Fortschritts
    formel = WORKSHEET.Cells(position - OFFSET, ziel_index-1).formula
    formel_aktualisiert = re.sub(r'([A-Z]+)(\d+)',ersetze_uebereinstimmung, formel)
    WORKSHEET.Cells(position, ziel_index-1).formula = formel_aktualisiert
    
    
    #Extrahieren der Formel zur Berechnung der Restzeit
    OFFSET += 5
    formel = WORKSHEET.Cells(position - OFFSET, ziel_index + 1).formula
    formel_aktualisiert = re.sub(r'([A-Z]+)(\d+)',ersetze_uebereinstimmung, formel)
    WORKSHEET.Cells(position, ziel_index + 1).formula = formel_aktualisiert
    
    
def get_task(id):
    global TASKS
    prefix = get_positiion()
    
    for task in TASKS:
        if f"{prefix}.{task.OutlineNumber}" == id:
            return task
    return None

def extract_parent_id(children_id):
    parts = children_id.split('.')
    last_digit = int(parts[-1]) - 1
    
    if last_digit == 0:
        return ".".join(parts[:-1])
    else:
        parts[-1] = str(last_digit)
        id = ".".join(parts)
    return id

def get_children_postion(parent_id):
    global WORKSHEET
    
    excel_data = WORKSHEET.UsedRange.Value
    index = 2
    
    for row in excel_data[1:]:
        if parent_id == row[3]:
            return index + 1
        index += 1
    return None


def update_values():
    parent_id = get_children_postion()
    


def extract_id(input):
    if input == None:
        return None
    result = re.sub(r'^\d+\.', '', input)
    return result


def choose_project_file():
    root = tk.Tk()
    root.withdraw()
    
    file_path = filedialog.askopenfilename(filetypes=[("Project-Dateien", "*.mpp")])
    return file_path

def get_positiion():
    global EXCEL
    active_sheet = EXCEL.ActiveSheet
    
    for i, sheet in enumerate(EXCEL.Worksheets, 1):
        if sheet.Name == active_sheet.Name:
            return i - 1
    return None


def find_ID(value):
    sheet_name = EXCEL.ActiveWorkbook.ActiveSheet.Name
    excel_worksheet = EXCEL.ActiveWorkbook.Sheets(sheet_name)
    
    
    headers = excel_worksheet.Rows(1).value[0]
    
    for col_index, header in enumerate(headers, start=1):
        if header == value:
            return col_index
    return -1

def find_Index(value):
    sheet_name = EXCEL.ActiveWorkbook.ActiveSheet.Name
    excel_worksheet = EXCEL.ActiveWorkbook.Sheets(sheet_name)
    
    
    headers = excel_worksheet.Rows(2).value[0]
    
    for col_index, header in enumerate(headers, start=1):
        if header == value:
            return col_index
    return -1


def search_id_column(excel_worksheet, id_column_index):
    excel_data = excel_worksheet.UsedRange.Value
    result = []
    
    for row in excel_data[1:]:
        if row[id_column_index-1] is None:
            continue
        cell_id = row[id_column_index-1]
        result.append(cell_id)
    return result


def init(excel_file_path):
    global EXCEL
    global EXCEL_FILE_PATH
    global PROJECT
    global ACTIVE_PROJECT
    global TASKS
    global WORKSHEET
    
    project_file_path = choose_project_file()
    
    PROJECT = win32.Dispatch("MSProject.Application")
    EXCEL = win32.Dispatch("Excel.Application")
    EXCEL_FILE_PATH = excel_file_path
    # WORKBOOK = EXCEL.Workbooks.Open(EXCEL_FILE_PATH)
    WORKSHEET = EXCEL.ActiveWorkbook.Sheets("1. Software Entwicklungssuppor")
    PROJECT.FileOpen(project_file_path)
    
    ACTIVE_PROJECT = PROJECT.ActiveProject
    TASKS = ACTIVE_PROJECT.Tasks
    main()

def sorting_key(item):
    parts = item.split('.')
    return [int(part) if part.isdigit() else part for part in parts]
    
def main():
    global EXCEL
    global EXCEL_FILE_PATH
    global PROJECT
    global ACTIVE_PROJECT
    global TASKS
    global WORKSHEET
    
    if TASKS is None:
        messagebox.showerror("Error", "Keine Project Daten erhalten")
        sys.exit(1)
        
    id_index = find_ID("ID")
    if id_index == -1:
        messagebox.showerror("Error", "ID konnte in Excel nicht gefunden werden")
        sys.exit() 
    tasks_id_unsorted = {f"{get_positiion()}.{task.OutlineNumber}"for task in TASKS}
    tasks_id = sorted(tasks_id_unsorted, key=sorting_key)
    
    for id in tasks_id:
        excel_id = search_id_column(WORKSHEET,id_index)
        if id not in excel_id:
            position = get_children_postion(extract_parent_id(id))
            if position is not None:
                task = get_task(id)
                add_task(task,position)
                
            
        
        
if __name__ == "__main__":
    init(r"C:\Users\npawelka\Desktop\PDCA\ETO.EEVACTUATOR.Entw.016.xlsm")