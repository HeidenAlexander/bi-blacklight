import subprocess
import os
from tkinter import messagebox


def publish_multiple_reports(folder_path, workspace_id):
    current_dir = os.getcwd()
    p = subprocess.Popen(["pwsh.exe", f"{current_dir}\\Modules\\PublishMultipleToWorkspace.ps1", folder_path,
                          workspace_id], stdout=subprocess.PIPE, text=True)
    while p.poll() is None:
        line = p.stdout.readline()
        print(line)
    print(p.stdout.read())
    if p.returncode != 0:
        messagebox.showinfo('Login Failed', 'Process ended due to login error.', )
    else:
        messagebox.showinfo('Process Completed', 'Process completed successfully.', )


def publish_single_report(file_path, workspace_id):
    current_dir = os.getcwd()
    p = subprocess.Popen(["pwsh.exe", f"{current_dir}\\Modules\\PublishSingleToWorkspace.ps1", file_path, workspace_id],
                         stdout=subprocess.PIPE, text=True)
    while p.poll() is None:
        line = p.stdout.readline()
        print(line)
    print(p.stdout.read())
    if p.returncode != 0:
        messagebox.showinfo('Login Failed', 'Process ended due to login error.', )
    else:
        messagebox.showinfo('Process Completed', 'Process completed successfully.', )
