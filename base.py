# todo:

#make path both sample path and process path


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
             "mla": r"I:\e24\SQN\Researchers\Haubmann Benjamin\01_PhD\23_MLA",
             "rie": r"I:\e24\SQN\Researchers\Haubmann Benjamin\01_PhD\24_RIE"}

Sampledir_dic = {"sem": "SEM", "plm": "PL", "epi": "MBE", "elx": "Elionix", "mic": "Microscope", "xrd": "XRD", "tem": "TEM", "mla": "MLA", "rie": "RIE"}

SampleOverview_column_dic = {"sem": "S", "plm": "T", "epi": "P", "elx": "H", "mic": "R", "xrd": "U", "tem": "V", "mla": "I", "rie": "M"}


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


def get_process_subdir(ID):
    for key in IDdir_dic.keys():
        if key in ID:
            return Sampledir_dic[key]
    return None


def get_sample_path(spl_name, base):
    for key in base.keys():
        if "spl" in key and spl_name in key:
            return base[key]["path"]
    return None


def sync_folder(ID):
    with open(IDbase_dir, 'rb') as file:
        base = pickle.load(file)

    if not ID_exists(ID, base):
        print(f"{RED}Invalid ID{RESET}")
        return

    tag_dict = base[ID].get("tags", {})
    if not tag_dict:
        return

    source = base[ID]["path"]
    count = 0
    for spl_name, copy_path in tag_dict.items():
        if os.path.exists(copy_path) and not os.path.isdir(copy_path):
            print(f"{YELLOW}Skipping \"{spl_name}\": copy path is a file, not a folder — please fix tag manually{RESET}")
            continue
        if os.path.isdir(copy_path):
            shutil.rmtree(copy_path)
        shutil.copytree(source, copy_path)
        count += 1

    print(f"{GREEN}Synced \"{ID}\" to {count} sample folder(s){RESET}")


def sync_all():
    with open(IDbase_dir, 'rb') as file:
        base = pickle.load(file)

    synced = 0
    for ID in base.keys():
        tag_dict = base[ID].get("tags", {})
        if tag_dict:
            source = base[ID]["path"]
            for spl_name, copy_path in tag_dict.items():
                if os.path.exists(copy_path) and not os.path.isdir(copy_path):
                    print(f"{YELLOW}Skipping \"{spl_name}\" for \"{ID}\": copy path is a file, not a folder — please fix tag manually{RESET}")
                    continue
                if os.path.isdir(copy_path):
                    shutil.rmtree(copy_path)
                shutil.copytree(source, copy_path)
                synced += 1

    basepath = r"I:\e24\SQN\Researchers\Haubmann Benjamin\01_PhD\11_Samples"
    cleaned = 0
    if os.path.exists(basepath):
        for sample_folder in os.listdir(basepath):
            sample_folder_path = os.path.join(basepath, sample_folder)
            if not os.path.isdir(sample_folder_path):
                continue
            for process_folder in os.listdir(sample_folder_path):
                process_folder_path = os.path.join(sample_folder_path, process_folder)
                if not os.path.isdir(process_folder_path):
                    continue
                for item in os.listdir(process_folder_path):
                    item_path = os.path.join(process_folder_path, item)
                    if not os.path.isfile(item_path):
                        continue
                    ID = item[:16]
                    if not ID_exists(ID, base):
                        continue
                    for spl_name, copy_path in base[ID].get("tags", {}).items():
                        if os.path.dirname(copy_path) == process_folder_path and os.path.isdir(copy_path):
                            os.remove(item_path)
                            cleaned += 1
                            break

    print(f"{GREEN}Synced {synced} copy/copies, removed {cleaned} loose file(s){RESET}")


def tag(ID, spl_name):
    with open(IDbase_dir, 'rb') as file:
        base = pickle.load(file)

    if not ID_exists(ID, base):
        print(f"{RED}Invalid ID{RESET}")
        return

    if "tags" not in base[ID]:
        base[ID]["tags"] = {}

    if spl_name in base[ID]["tags"]:
        print(f"{YELLOW}Tag \"{spl_name}\" already exists for \"{ID}\"{RESET}")
        return

    sample_path = get_sample_path(spl_name, base)
    if sample_path is None:
        print(f"{RED}Sample \"{spl_name}\" not found in database{RESET}")
        return

    process = get_process_subdir(ID)
    if process is None:
        print(f"{RED}Could not determine process type for \"{ID}\"{RESET}")
        return

    folder_name = os.path.basename(base[ID]["path"])
    copy_path = os.path.join(sample_path, process, folder_name)

    if os.path.exists(copy_path):
        print(f"{YELLOW}Copy already exists at \"{copy_path}\", registering without re-copying{RESET}")
    else:
        shutil.copytree(base[ID]["path"], copy_path)

    base[ID]["tags"][spl_name] = copy_path

    with open(IDbase_dir, 'wb') as file:
        pickle.dump(base, file)

    print(f"{GREEN}Tagged \"{ID}\" with \"{spl_name}\" → {copy_path}{RESET}")


def untag(ID, spl_name):
    with open(IDbase_dir, 'rb') as file:
        base = pickle.load(file)

    if not ID_exists(ID, base):
        print(f"{RED}Invalid ID{RESET}")
        return

    tag_dict = base[ID].get("tags", {})
    if spl_name not in tag_dict:
        print(f"{RED}Tag \"{spl_name}\" not found for \"{ID}\"{RESET}")
        return

    copy_path = tag_dict[spl_name]
    print(f"{BLUE}Delete the copy at \"{copy_path}\"? {GREEN}y{BLUE}/{RED}n{RESET}")
    choice = input()
    if choice == "y":
        if os.path.isdir(copy_path):
            shutil.rmtree(copy_path)
            print(f"{GREEN}Deleted copy{RESET}")
        elif os.path.isfile(copy_path):
            os.remove(copy_path)
            print(f"{GREEN}Deleted file{RESET}")
        else:
            print(f"{YELLOW}Copy path does not exist, skipping deletion{RESET}")
    elif choice == "n":
        pass
    else:
        print(f"{RED}Invalid input, tag not removed{RESET}")
        return

    del base[ID]["tags"][spl_name]

    with open(IDbase_dir, 'wb') as file:
        pickle.dump(base, file)

    print(f"{GREEN}Tag \"{spl_name}\" removed from \"{ID}\"{RESET}")



def untagged():
    with open(IDbase_dir, 'rb') as file:
        base = pickle.load(file)

    exclude = {"sim", "scr", "ana", "spl", "des"}
    result = []
    for ID in base.keys():
        if any(ex in ID for ex in exclude):
            continue
        if not base[ID].get("tags"):
            result.append(ID)

    if not result:
        print(f"{GREEN}All IDs have at least one tag{RESET}")
    else:
        for ID in sorted(result):
            print(f"{MAGENTA}{ID}{RESET}")



def list_tags(ID):
    with open(IDbase_dir, 'rb') as file:
        base = pickle.load(file)

    if not ID_exists(ID, base):
        print(f"{RED}Invalid ID{RESET}")
        return

    tag_dict = base[ID].get("tags", {})
    if not tag_dict:
        print(f"{YELLOW}No tags for \"{ID}\"{RESET}")
    else:
        for spl_name, copy_path in tag_dict.items():
            status = f"{GREEN}ok{RESET}" if os.path.exists(copy_path) else f"{RED}missing{RESET}"
            print(f"  {MAGENTA}{spl_name}{RESET} → {copy_path} [{status}]")


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

        base[self.ID] = {"path": self.path, "info": info, "comments": "", "tags": {}}


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
            new_parent_dir = IDdir_dic[key]

    new_path = shutil.move(renamed_folder_path, new_parent_dir)

    add(new_path)

    print(f"{BLUE}Which samples are involved? Enter sample names separated by commas (e.g. spl01,spl02), or press Enter to skip:{RESET}")
    sample_input = input().strip()
    if sample_input:
        for spl_name in [s.strip() for s in sample_input.split(",") if s.strip()]:
            tag(ID, spl_name)



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

    tag_dict = base[ID].get("tags", {})

    print(f"{BLUE}Do you really want to delete entry \"{ID}\"? {GREEN}y{BLUE}/{RED}n{RESET}")
    choice = input()
    if choice == "y":

        if tag_dict:
            print(f"{BLUE}Also delete {len(tag_dict)} sample folder copy/copies? {GREEN}y{BLUE}/{RED}n{RESET}")
            copy_choice = input()
            if copy_choice == "y":
                for spl_name, copy_path in tag_dict.items():
                    if os.path.exists(copy_path):
                        shutil.rmtree(copy_path)
                        print(f"{GREEN}Deleted copy for \"{spl_name}\"{RESET}")

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



def checkall():

    with open(IDbase_dir, 'rb') as file:
        base = pickle.load(file)

    invalid_list = [ID for ID in base.keys() if not os.path.exists(base[ID]["path"])]

    if len(invalid_list) == 0:
        print(f"{GREEN}Everything seems up to date{RESET}")
        return

    print(f"{RED}Invalid path found for {len(invalid_list)} entrie(s). Searching...{RESET}\n")

    for ID in invalid_list:
        expected_dir = None
        for key in IDdir_dic.keys():
            if key in ID:
                expected_dir = IDdir_dic[key]
                break

        if expected_dir is None or not os.path.exists(expected_dir):
            print(f"{RED}{ID}{RESET} — could not determine expected directory, skipping")
            continue

        matches = [f for f in os.listdir(expected_dir) if f[:16] == ID and os.path.isdir(os.path.join(expected_dir, f))]

        if not matches:
            print(f"{RED}{ID}{RESET} — not found in {expected_dir}, skipping")
            continue

        for match in matches:
            new_path = os.path.join(expected_dir, match)
            print(f"{YELLOW}{ID}{RESET} — found at: {new_path}")
            print(f"{BLUE}Update path? {GREEN}y{BLUE}/{RED}n{RESET}")
            choice = input()
            if choice == "y":
                base[ID]["path"] = new_path
                with open(IDbase_dir, 'wb') as file:
                    pickle.dump(base, file)
                with open(IDbase_dir, 'rb') as file:
                    base = pickle.load(file)
                print(f"{GREEN}Path updated{RESET}")
            else:
                print(f"{YELLOW}Skipped{RESET}")


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
    print(f"{BLUE}Do you really want to change the path of entry \"{ID}\" from \"{oldpath}\" to \"{path}\"? {GREEN}y{BLUE}{RED}n{RESET}")

    choice = input()
    if choice == "y":

        base[ID]["path"] = path
        with open(IDbase_dir, 'wb') as file:
            pickle.dump(base, file)
        print(f"{GREEN}Path has been updated{RESET}")
        sync_folder(ID)

    elif choice == "n":
        return

    else:
        print(f"{RED}Invalid Entry{RESET}")
        update(path)


def update_readme_single(ID):

    with open(IDbase_dir, 'rb') as file:
        base = pickle.load(file)

    readme_path = base[ID]["path"] + "\\" + ID + "_readme.txt"

    if not os.path.exists(readme_path):
        print(f"{YELLOW}Skipping \"{ID}\": readme not found{RESET}")
        return False

    with open(readme_path, 'r') as readme:
        content = readme.read()

    parts = [y.strip() for y in content.split("#"*70, 1)]
    if len(parts) != 2:
        print(f"{YELLOW}Skipping \"{ID}\": readme has no separator, cannot parse{RESET}")
        return False

    info_new, comments_new = parts
    base[ID]["info"] = info_new
    base[ID]["comments"] = comments_new

    with open(IDbase_dir, 'wb') as file:
        pickle.dump(base, file)

    sync_folder(ID)
    return True



def update_readme():

    with open(IDbase_dir, 'rb') as file:
        base = pickle.load(file)

    ok = 0
    skipped = 0
    for ID in base.keys():
        if update_readme_single(ID):
            ok += 1
        else:
            skipped += 1

    print(f"{GREEN}Updated {ok} readme(s){RESET}" + (f", skipped {skipped}" if skipped else ""))



def display(type_):

    with open(IDbase_dir, 'rb') as file:
        base = pickle.load(file)

    l = []
    for ID in base.keys():
        l.append(ID)

    for ID in sorted(l):
        if type_ in ID or type_ == "all":
            info = base[ID]["info"]
            tag_dict = base[ID].get("tags", {})
            tag_str = f" {YELLOW}[{', '.join(tag_dict.keys())}]{RESET}" if tag_dict else ""
            print(f"{MAGENTA}{ID}{RESET}{tag_str}\n{info}\n")


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

    update_readme_single(ID)

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


def edit_readme(ID):

    with open(IDbase_dir, 'rb') as file:
        base = pickle.load(file)

    if not ID_exists(ID, base):
        print(f"{RED}Invalid ID{RESET}")
        return

    readme_path = base[ID]["path"] + "\\" + ID + "_readme.txt"

    if not os.path.exists(readme_path):
        print(f"{RED}Readme not found: {readme_path}{RESET}")
        return

    subprocess.run(['notepad', readme_path])

    update_readme_single(ID)
    print(f"{GREEN}Readme saved and synced{RESET}")





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

    # Subparser for update
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

    # Subparser for edit_readme
    parser_edit_readme = subparsers.add_parser("edit_readme", help="Open readme in text editor and sync on close")
    parser_edit_readme.add_argument("ID", type=str, help="ID whose readme to edit")

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

    # Subparser for update_SampleOverview
    parser_update_SampleOverview = subparsers.add_parser("update_SampleOverview", help="Update SampleOverview")
    parser_update_SampleOverview.add_argument("ID", type=str, help="process ID")
    parser_update_SampleOverview.add_argument("spl", type=str, help="sample number")

    # Subparser for tag
    parser_tag = subparsers.add_parser("tag", help="Tag an ID with a sample")
    parser_tag.add_argument("ID", type=str, help="Process ID to tag")
    parser_tag.add_argument("spl_name", type=str, help="Sample name to tag with")

    # Subparser for untag
    parser_untag = subparsers.add_parser("untag", help="Remove a sample tag from an ID")
    parser_untag.add_argument("ID", type=str, help="Process ID to untag")
    parser_untag.add_argument("spl_name", type=str, help="Sample name to remove")

    # Subparser for sync
    parser_sync = subparsers.add_parser("sync", help="Sync ID folder to all tagged sample copies")
    parser_sync.add_argument("ID", type=str, help="ID to sync")

    # Subparser for sync_all
    parser_sync_all = subparsers.add_parser("sync_all", help="Sync all tagged ID folders to their sample copies")

    # Subparser for tags
    parser_tags = subparsers.add_parser("tags", help="List tags for an ID")
    parser_tags.add_argument("ID", type=str, help="ID whose tags to list")

    # Subparser for untagged
    parser_untagged = subparsers.add_parser("untagged", help="List all IDs without tags (excludes sim, scr, ana)")

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
    elif args.function == "edit_readme":
        edit_readme(args.ID)
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
    elif args.function == "update_SampleOverview":
        update_SampleOverview(args.ID, args.spl)
    elif args.function == "tag":
        tag(args.ID, args.spl_name)
    elif args.function == "untag":
        untag(args.ID, args.spl_name)
    elif args.function == "sync":
        sync_folder(args.ID)
    elif args.function == "sync_all":
        sync_all()
    elif args.function == "tags":
        list_tags(args.ID)
    elif args.function == "untagged":
        untagged()



    #things to add

# edit path; check if path exists; search for Entry automatically
# When ID path is updated, readme file should move as well
