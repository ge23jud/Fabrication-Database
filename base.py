import argparse
import pickle
import os
from datetime import datetime
import sys
import ctypes
import subprocess
import tkinter as tk
from tkinter import filedialog
import shutil
import re
#import openpyxl
from openpyxl import load_workbook
import win32com.client
from pathlib import Path

# Enable ANSI escape codes on Windows
kernel32 = ctypes.windll.kernel32
kernel32.SetConsoleMode(kernel32.GetStdHandle(-11), 7)

# Define ANSI color codes
RED = '\033[91m'
GREEN = '\033[92m'
BLUE = '\033[94m'
RESET = '\033[0m'
YELLOW = '\033[93m'
MAGENTA = '\033[95m'

IDbase_dir = r"I:\e24\SQN\Researchers\Haubmann Benjamin\01_PhD\IDbase"
SampleOverview_dir = r"I:\e24\SQN\Researchers\Haubmann Benjamin\01_PhD\Sample Overview.xlsx"
sheet_name = "Tabelle1"

IDdir_dic = {"sem": r"I:\e24\SQN\Researchers\Haubmann Benjamin\01_PhD\16_SEM",
             "plm": r"I:\e24\SQN\Researchers\Haubmann Benjamin\01_PhD\13_PL",
             "epi": r"I:\e24\SQN\Researchers\Haubmann Benjamin\01_PhD\15_Growth",
             "elx": r"I:\e24\SQN\Researchers\Haubmann Benjamin\01_PhD\12_Elionix",
             "mic": r"I:\e24\SQN\Researchers\Haubmann Benjamin\01_PhD\14_Microscope",
             "xrd": r"I:\e24\SQN\Researchers\Haubmann Benjamin\01_PhD\20_XRD",
             "tem": r"I:\e24\SQN\Researchers\Haubmann Benjamin\01_PhD\21_TEM", 
             "mla": r"I:\e24\SQN\Researchers\Haubmann Benjamin\01_PhD\23_MLA"}
               
Sampledir_dic = {"sem": "SEM", "plm": "PL", "epi": "MBE", "elx": "Elionix", "mic": "Microscope", "xrd": "XRD", "tem": "TEM", "mla": "MLA"}

SampleOverview_column_dic = {"sem": "R", "plm": "S", "epi": "O", "elx": "H", "mic": "Q", "xrd": "T", "tem": "U", "mla": "I"}


def extract_ID_from_path(path):
    ID = path.split("\\")[-1][:16]
    if os.path.isfile(path):
        path = "\\".join(path.split("\\")[:-1])
        print(path)
    return ID, path    
    

def ID_exists(ID, base):
    if ID in base.keys():
        return True
    else:
        return False
        



class entry:
    
    def __init__(self, path):

        if not os.path.exists(path):
            print(f"{RED}Path does not exist{RESET}")
            return
        
        #self.path = path
        self.ID, self.path = extract_ID_from_path(path)
        
        with open(IDbase_dir, 'rb') as file:
            base = pickle.load(file)
        
        print(f"{BLUE}Please provide a short description about the entry you want to add{RESET}")
        info = input()
        separator = "#"*70
        
        with open(f"{self.path}\\{self.ID}_readme.txt", "a") as readme:
            readme.write(f"{info}\n\n{separator}\n\n")
            
        base[self.ID] = {"path": self.path, "info": info, "comments": ""}
        
        
        with open(IDbase_dir, 'wb') as file:
            pickle.dump(base, file)
         
        print(f"{GREEN}Entry \"{self.ID}\" has been added{RESET}")
         

def new_sample(spl_name):
    
    basepath = r"I:\e24\SQN\Researchers\Haubmann Benjamin\01_PhD\11_Samples"
    current_date = datetime.now()
    date = current_date.strftime('%Y%m%d')
    
    with open(IDbase_dir, 'rb') as file:
        base = pickle.load(file)
    
    samples = sorted([key for key in base.keys() if "spl" in key])
    
    if any([True for x in samples if spl_name in x]):
        print(f"{RED}Sample already exists.{RESET}")
        return
        
    else:
        path = basepath + "\\" + date + "-" + spl_name
        os.makedirs(path)
        for key in Sampledir_dic:
            subfolder_name = Sampledir_dic[key]
            os.makedirs(path + "\\" + subfolder_name)
     
        add(path)


def convert_date_format(date_str):

    if len(date_str) != 8 or not date_str.isdigit():
        raise ValueError("Input must be a valid date in 'YYYYMMDD' format.")
    
    year = date_str[:4]
    month = date_str[4:6]
    day = date_str[6:8]
    
    return f"{day}.{month}.{year}"


def write_to_cell(file_name, sheet_name, cell_address, value):
 
    workbook = load_workbook(file_name)
    sheet = workbook[sheet_name]
    
    existing_value = sheet[cell_address].value
    if existing_value == None:
        sheet[cell_address] = value
    else:
        sheet[cell_address] = (str(sheet[cell_address].value) or "") + "\n\n" + value
    workbook.save(file_name)
    print(f"Value '{value}' written to {cell_address} in sheet '{sheet_name}' of '{file_name}'.")

def save_close_excel(file_path):

    xl = win32com.client.GetActiveObject("Excel.Application")
    
    # Look for the workbook in open workbooks
    for wb in xl.Workbooks:
        if wb.FullName == file_path:
            print(f"Sample Overview active")
            wb.Save()  # Save before closing
            wb.Close()
            print(f"Closed Sample Overview")
            break

        
        
def reopen_excel(file_path):

    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True  # Open Excel in visible mode
        excel.Workbooks.Open(file_path)
        print(f"Opened Sample Overview")
    except Exception as e:
        print(f"Error while opening Sample Overview")
    

def get_sample_index(spl_name):
    
    with open(IDbase_dir, 'rb') as file:
        base = pickle.load(file)
    
    samples = sorted([key for key in base.keys() if "spl" in key])
    
    index = 0
    for i, sample in enumerate(samples):
        if spl_name in sample:
            return index
        else:
            index += 1
    print("No matching sample found")


def update_SampleOverview(ID, spl):
    
    date = convert_date_format(ID[:8])
    print(get_sample_index(spl))
    row = str(get_sample_index(spl) + 2)
    print(row)
    column = get_column(ID)
    cell = column + row
    value = f"{date}, {ID}"
    
    

    save_close_excel(SampleOverview_dir)
    
    write_to_cell(SampleOverview_dir, sheet_name, cell, value)
    
    # In case of Elionix process: Write design-ID in column "R"
    if "elx" in ID:
        design_ID = input("Design-ID: ")
        design_cell = "R" + row
        write_to_cell(SampleOverview_dir, sheet_name, design_cell, design_ID)
    
    reopen_excel(SampleOverview_dir)


def get_column(ID):
    for key in IDdir_dic.keys():
        if key in ID:
            column = SampleOverview_column_dic[key]
    return column


def create(new_name, initial_path = r"I:\e24\SQN\Researchers\Haubmann Benjamin\01_PhD"):
    
    root = tk.Tk()
    root.withdraw()  

    folder_path = filedialog.askdirectory(title="Select a Folder", initialdir=initial_path)

    parent_dir = os.path.dirname(folder_path)

    renamed_folder_path = os.path.join(parent_dir, new_name)
    os.rename(folder_path, renamed_folder_path)
    
    ID = new_name[:16]
    
    if len(new_name) > 16:
        description = new_name[16:]
    
    for key in IDdir_dic.keys():
        if key in ID:
            process = Sampledir_dic[key]
            new_parent_dir = IDdir_dic[key]
      
    new_path = shutil.move(renamed_folder_path, new_parent_dir)
    
    add(new_path)
    
    valid_pattern = r'spl\d{4}'
    valid_matches = re.findall(valid_pattern, new_name)
    basepath = r"I:\e24\SQN\Researchers\Haubmann Benjamin\01_PhD\11_Samples"
    dirs = os.listdir(basepath)
    
    result = []
    for item in dirs:
        if any(match in item for match in valid_matches):
            result.append(basepath + "\\" + item)
    
    for x in result:
        shutil.copytree(new_path, x + "\\" + process + "\\" + new_name)
    
    for match in valid_matches:
        update_SampleOverview(ID, match)
    
    
        
def add(path):
    entry(path)

    

def goto(ID):

    with open(IDbase_dir, 'rb') as file:
        base = pickle.load(file)
    
    if not ID_exists(ID, base):
        print(f"{RED}Invalid ID{RESET}")
        return
    
    path = base[ID]["path"]
    os.startfile(path)
   
   
def delete(ID):

    with open(IDbase_dir, 'rb') as file:
        base = pickle.load(file)
    
    if not ID_exists(ID, base):
        print(f"{RED}Invalid ID{RESET}")
        return
    
    print(f"{BLUE}Do you really want to delete entry \"{ID}\"? {GREEN}y{BLUE}/{RED}n{RESET}")
    choice = input()
    if choice == "y":

        del base[ID]
        with open(IDbase_dir, 'wb') as file:
            pickle.dump(base, file)
        print(f"{GREEN}Entry \"{ID}\" has been deleted{RESET}")
            
    elif choice == "n":
        return
        
    else:
        print(f"{RED}Invalid Entry{RESET}")
        delete(ID)
        
    

def ls():
    with open(IDbase_dir, 'rb') as file:
        base = pickle.load(file)
    l = []
    for ID in base.keys():
        l.append(ID)
    
    for ID in sorted(l):
        print(f"{MAGENTA}{ID}{RESET}")
        

def path_is_valid(ID):

    with open(IDbase_dir, 'rb') as file:
        base = pickle.load(file)
        
    if not os.path.exists(base[ID]["path"]):
        return False
    else:
        return True
        
 
def checkall():
    
    with open(IDbase_dir, 'rb') as file:
        base = pickle.load(file)
      
    invalid_list = [] 
     
    for ID in base.keys():
        if not path_is_valid(ID):
            invalid_list.append(ID)
    
    if len(invalid_list) == 0:
        print(f"{GREEN}Everythings seems up to date{RESET}")
    else:
        print(f"{RED}Invalid path found for the following entries:{RESET}")
        for ID in invalid_list:
            print(f"{RED}{ID}{RESET}")
            
            
def update(path):
    
    if not os.path.exists(path):
        print(f"{RED}Path does not exist{RESET}")
        return
    
    ID, path = extract_ID_from_path(path)
    
    with open(IDbase_dir, 'rb') as file:
        base = pickle.load(file)
        
    if not ID in base.keys():
        print(f"{RED}Path exists, but ID has not been added yet{RESET}")
        return
        
    oldpath = base[ID]["path"]
    print(f"{BLUE}Do you really want to change the path of entry \"{ID}\" from \"{oldpath}\" to \"{path}\"? {GREEN}y{BLUE}\{RED}n{RESET}")
 
    choice = input()
    if choice == "y":

        base[ID]["path"] = path
        with open(IDbase_dir, 'wb') as file:
            pickle.dump(base, file)
        print(f"{GREEN}Path has been updated{RESET}")
            
    elif choice == "n":
        return
        
    else:
        print(f"{RED}Invalid Entry{RESET}")
        update(path)
        
 
def update_readme_single(ID):

    with open(IDbase_dir, 'rb') as file:
        base = pickle.load(file)

    with open(base[ID]["path"]+"\\"+ID+"_readme.txt", 'r') as readme:
        content = readme.read()
    
    info_new, comments_new = [y.strip() for y in content.split("#"*70)]
    base[ID]["info"] = info_new
    base[ID]["comments"] = comments_new  

    with open(IDbase_dir, 'wb') as file:
        pickle.dump(base, file)
      
        
   
def update_readme():

    with open(IDbase_dir, 'rb') as file:
        base = pickle.load(file)
       
    for ID in base.keys():
        update_readme_single(ID)
    
    print(f"{GREEN}Updated sucessfully{RESET}")


def find(ID):

    with open(IDbase_dir, 'rb') as file:
        base = pickle.load(file)
        
    if not ID in base.keys():
        print(f"{RED}Invalid ID{RESET}")
        return
    
    print(f"{GREEN}Searching for path of entry \"{ID}\" in PhD folder...{RESET}")
    for root, dirs, files in os.walk(r"I:\e24\SQN\Researchers\Haubmann Benjamin\01_PhD"):
        continue
        

def display(type_):

    with open(IDbase_dir, 'rb') as file:
        base = pickle.load(file)
    
    l = []
    for ID in base.keys():
        l.append(ID)
       
    for ID in sorted(l):
        if type_ in ID or type_ == "all":
            info = base[ID]["info"]
            print(f"{MAGENTA}{ID}{RESET}\n{info}\n")
        

def comment(ID):

    with open(IDbase_dir, 'rb') as file:
        base = pickle.load(file)
        
    if not ID in base.keys():
        print(f"{RED}Invalid ID{RESET}")
        return
        
    print(f"{BLUE}Type your comment{RESET}")
    
    comment = input()
    now = datetime.now()
    full_date = now.strftime("%A, %B %d, %Y")
    comment = f"{full_date}:\n{comment}\n\n"
    
    with open(base[ID]["path"]+"\\"+ID+"_readme.txt", 'a') as readme:
        readme.write(comment)
      
    print(f"{GREEN}Comment added{RESET}")
    

def inspect(ID):

    with open(IDbase_dir, 'rb') as file:
        base = pickle.load(file)
        
    if not ID in base.keys():
        print(f"{RED}Invalid ID{RESET}")
        return
        
    with open(base[ID]["path"]+"\\"+ID+"_readme.txt", 'r') as readme:
        content = readme.read()
        
    print(content)
 

def runpy(ID): # not operational

    with open(IDbase_dir, 'rb') as file:
        base = pickle.load(file)
        
    if not ID in base.keys():
        print(f"{RED}Invalid ID{RESET}")
        return
    
    folder_path = base[ID]["path"]
    all_items = os.listdir(folder_path)
    
    for item in all_items:
        file_path = f"{folder_path}\\{item}"
        if ID in item and ".py" in item:
            print(file_path)
            subprocess.run(['python', file_path])
            return
        
    
    
  
    

def parse_arguments():
    parser = argparse.ArgumentParser(description='Call individual functions from the command line.')
    subparsers = parser.add_subparsers(dest='function', required=True)

    # Subparser for add
    parser_add = subparsers.add_parser('add', help='Call add')
    parser_add.add_argument('path', type=str, help='folder path of the entry to add')
    
    # Subparser for goto
    parser_goto = subparsers.add_parser("goto", help="Call goto")
    parser_goto.add_argument("ID", type=str, help="ID which should be opened")

    # Subparser for ls
    parser_ls = subparsers.add_parser('ls', help='Call ls')
    
    # Subparser for delete
    parser_delete = subparsers.add_parser("delete", help="Call delete")
    parser_delete.add_argument("ID", type=str, help="ID which should be deleted")
    
    # Subparser for checkall
    parser_checkall = subparsers.add_parser('checkall', help='Call checkall')
    
    # Subparser for delete
    parser_update = subparsers.add_parser("update", help="Call update")
    parser_update.add_argument("path", type=str, help="path which should be updated")
    
    # Subparser for display
    parser_display = subparsers.add_parser("display", help="Call display")
    parser_display.add_argument("type_", type=str, help="type of entries to be displayed")
    
    # Subparser for update_readme
    parser_update_readme = subparsers.add_parser('update_readme', help='Call update_readme')
    
    # Subparser for comment
    parser_comment = subparsers.add_parser("comment", help="Call comment")
    parser_comment.add_argument("ID", type=str, help="Comment to add")
    
    # Subparser for inspect
    parser_inspect = subparsers.add_parser("inspect", help="Call inspect")
    parser_inspect.add_argument("ID", type=str, help="ID whose readme is to show")
    
    # Subparser for runpy
    parser_runpy = subparsers.add_parser("runpy", help="Call runpy")
    parser_runpy.add_argument("ID", type=str, help="ID to run script")
    
    # Subparser for create
    parser_create = subparsers.add_parser("create", help="Create ID dir from folder and sort automatically")
    parser_create.add_argument("new_name", type=str, help="name of the ID folder")
   
    # Subparser for new_sample
    parser_new_sample = subparsers.add_parser("new_sample", help="Create new sample")
    parser_new_sample.add_argument("spl_name", type=str, help="Sample name")
    
    # Subparser for reopen_excel
    parser_reopen_excel = subparsers.add_parser("reopen_excel", help="Open Sample Overview")
    parser_reopen_excel.add_argument("file_path", type=str, help="Path to Sample Overview")
    
    # Subparser for write_to_cell
    parser_write_to_cell = subparsers.add_parser("write_to_cell", help="Write value to excel cell")
    parser_write_to_cell.add_argument("file_name", type=str, help="Name of excel file")
    parser_write_to_cell.add_argument("sheet_name", type=str, help="Name of excel sheet")
    parser_write_to_cell.add_argument("cell_address", type=str, help="Which cell to write to")
    parser_write_to_cell.add_argument("value", type=str, help="What to write to cell")
    
    # Subparser for save_close_excel
    parser_save_close_excel = subparsers.add_parser("save_close_excel", help="Save and close Sample Overview")
    parser_save_close_excel.add_argument("file_path", type=str, help="Path to Sample Overview")

    return parser.parse_args()    
       
if __name__ == "__main__":

    args = parse_arguments()

    if args.function == 'add':
        add(args.path)
    elif args.function == 'goto':
        goto(args.ID)
    elif args.function == "delete":
        delete(args.ID)
    elif args.function == "ls":
        ls()
    elif args.function == "checkall":
        checkall()
    elif args.function == "update":
        update(args.path)
    elif args.function == "display":
        display(args.type_)
    elif args.function == "update_readme":
        update_readme()
    elif args.function == "comment":
        comment(args.ID)
    elif args.function == "inspect":
        inspect(args.ID)
    elif args.function == "runpy":
        runpy(args.ID)
    elif args.function == "create":
        create(args.new_name)
    elif args.function == "new_sample":
        new_sample(args.spl_name)
    elif args.function == "reopen_excel":
        reopen_excel(args.file_path)
    elif args.function == "write_to_cell":
        write_to_cell(args.file_name, args.sheet_name, args.cell_address, args.value)
    elif args.function == "save_close_excel":
        save_close_excel(args.file_path)   

#things to add

# edit path; check if path exists; search for Entry automatically
# When ID path is updated, readme file should move as well

