from export_func import *
############################
DEMO = True
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
############################
def init(project_file_path):
    global DATA_FRAME
    global EXCEL_FILE_PATH
    global PROJECT
    global SELECTED_SHEET
    global ACTIVE_PROJECT
    global TASKS
    global PROJECT_FILE_PATH
    
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
    main()
    

def main():
    saved_depth = 0
    task_number = 0
    first_summary = True
    Task_Name_index = find_TASK_NAME(DATA_FRAME, TASK_NAME)
    ID_index = find_ID(DATA_FRAME)
    START_index = find_START(DATA_FRAME)
    BUDGET_index = find_BUDGET(DATA_FRAME)
    RESOURCE_index = find_RESOURCE(DATA_FRAME)
    if Task_Name_index == -1 or ID_index == -1 or START_index == -1:
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
            if first_summary:
                task_number += 1
                START_DATE = simpledialog.askstring("Start festlegen",f"Bitte legen sie ein Start für {current_name} fest")
                first_summary = False
                add_Summary(TASKS,current_name,current_depth,START_DATE)
                WAS_SUMMARY = True
                saved_depth = current_depth
            else:
                date = row.iloc[START_index]
                if date == None:
                    date = datetime.now().strftime("%d.%m.%Y")
                if current_depth < saved_depth :
                    add_Summary(TASKS,current_name,current_depth,date)
                    task_number += 1
                    WAS_SUMMARY = True
                    saved_depth = current_depth
                else:
                    if WAS_SUMMARY:
                        add_Task(TASKS,current_name,current_depth,date,extract_budget(current_budget),-1)
                        task_number += 1
                        WAS_SUMMARY = False
                        saved_depth = current_depth
                        # resource_list = row[RESOURCE_index].ListEntries
                        # resource = ACTIVE_PROJECT.Resources
                        # add_resource(resource_list,resource)
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
            messagebox.showerror("Error","Fehler bei der Ermittlung des Pfades")
            sys.exit()
        else:
            mpp_file_path = sys.argv[1]
            init(mpp_file_path)
    else:
        init(r"C:\Users\npawelka\Desktop\Beispiel.mpp")
