import os
import requests
import tkinter
from tkinter import filedialog
from tkinter import *
import tkinter as tk

root = tk.Tk()
frame = tk.Frame(root)
frame.pack()




api_key = input("Please Enter Your API Key- ")
url = "https://8oiyjy8w63.execute-api.us-west-2.amazonaws.com/Prod/api/rebuild/file"
file_extenstions = [".xlsm", ".pdf", ".jpg", ".gif", ".png", ".emf", ".tiff", ".bmp", ".doc", ".doc", ".xls", ".xlt", ".ppt", ".pot", ".docx", ".docm", ".dotx", ".dotm", ".xlsx", ".xlam", ".xltx", ".xltm", ".pptx", ".potx", ".potx", ".potm", ".pptm", ".ppsx", ".ppam", ".ppsm"]


def list_excel_files(input_folder_path):
    # Returns a list of all excel files in the given folder path.
    excel_files = []

    # Insert code to loop through input folder, and add each file to the list.
    for file in os.listdir(input_folder_path):
            if any(file.endswith(fe) for fe in file_extenstions):
                print(os.path.join(input_folder_path, file))
                excel_files.append(file)
    return excel_files
    
    
def rebuild_files(excel_file_paths, input_folder_path, output_folder_path): 
    
    # Loops through the list passed in, rebuilds each file and stores each one in a new folder.
    for file in (excel_file_paths):
            input_file_path = os.path.join(input_folder_path, file)
            output_file_path = os.path.join(output_folder_path, file)
            # Insert code to send 'file' through Rebuild.
            with open(input_file_path, "rb") as f:
                response = requests.post(
                    url=url,
                    files=[("file", f)],
                    headers={
                        "x-api-key": api_key,
                        "accept": "application/octet-stream"
                    }
                )
            if response.status_code == 200 and response.content:
                os.makedirs(os.path.dirname(output_file_path), exist_ok=True)
                with open(output_file_path, "wb") as f:
                    f.write(response.content)
                print("Successfully wrote file to:",
                    os.path.abspath(output_file_path))
            else:
                response.raise_for_status()


print("Select Input Folder")
input_folder_path = tkinter.filedialog.askdirectory(initialdir = "/",title = "Select input folder")
print("Select Output Folder")
output_folder_path = tkinter.filedialog.askdirectory(initialdir = "/",title = "Select output folder")
excel_file_list = list_excel_files(input_folder_path)

rebuild_files(excel_file_list, input_folder_path, output_folder_path)
