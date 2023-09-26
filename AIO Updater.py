import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import sys
from win32com.client import Dispatch  # Import Dispatch from win32com.client

# Get the directory where the script is located
if getattr(sys, 'frozen', False):
    # If the code is frozen (running as an EXE), use the sys._MEIPASS variable to get the base directory
    base_dir = sys._MEIPASS
else:
    # If the code is running as a script, use the current directory
    base_dir = os.path.abspath(os.path.dirname(__file__))

# Determine the path to the 'players' folder in the same directory as the script
players_folder_path = os.path.join(base_dir, 'players')

def replace_discord_sdk(target_folder):
    source_dll_path = os.path.join(base_dir, 'discord_game_sdk.dll')
    target_path = os.path.join(target_folder, 'discord_game_sdk.dll').replace('/', '\\')

    print(f"Source DLL path: {source_dll_path}")
    print(f"Target DLL path: {target_path}")

    if os.path.exists(source_dll_path):
        try:
            if os.path.exists(target_path):
                os.remove(target_path)

            shutil.copy2(source_dll_path, target_path)
            print("discord_game_sdk.dll copied successfully")

            copied_dll_path = os.path.join(target_folder, 'copied_discord_game_sdk.dll').replace('/', '\\')
            if os.path.exists(copied_dll_path):
                os.remove(copied_dll_path)

        except Exception as e:
            print(f"Error copying discord_game_sdk.dll: {e}")
            messagebox.showerror("Error", f"Error copying discord_game_sdk.dll: {e}")
            return False
    else:
        print("Step 2: discord_game_sdk.dll not found in the program directory.")
        messagebox.showerror("Error", "discord_game_sdk.dll not found in the program directory.")
        return False

    return True

def create_desktop_shortcut(executable_path):
    desktop_path = os.path.expanduser('~\\Desktop')
    shortcut_path = os.path.join(desktop_path, 'game_dx12_ship_replay.lnk')

    try:
        os.remove(shortcut_path)
    except OSError:
        pass

    # Create a shortcut file on the desktop using win32com
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(shortcut_path)
    shortcut.Targetpath = executable_path
    shortcut.save()

    print("Desktop shortcut created successfully")

def select_folder():
    mw19_folder = filedialog.askdirectory(title="Select Call of Duty MW19 Folder")
    if mw19_folder:
        folder_entry.delete(0, tk.END)
        folder_entry.insert(0, mw19_folder)

def process_folders():
    mw19_folder = folder_entry.get()
    if not mw19_folder:
        messagebox.showerror("Error", "Please select the Call of Duty MW19 folder.")
        return

    if not replace_discord_sdk(mw19_folder):
        print("Step 2: Failed to replace discord_game_sdk.dll")
        return

    game_exe_path = os.path.join(mw19_folder, 'game_dx12_ship_replay.exe')
    if os.path.exists(game_exe_path):
        create_desktop_shortcut(game_exe_path)
        print("Step 3: Desktop shortcut created successfully")
    else:
        print("Step 3: game_dx12_ship_replay.exe not found in the folder")

    # Use the 'players' folder in the same directory as the script
    source_folder = players_folder_path
    target_folder = os.path.join(os.path.expanduser('~\\Documents\\Call of Duty Modern Warfare'), 'players')

    print(f"Source Folder: {source_folder}")
    print(f"Target Folder: {target_folder}")

    if not os.path.exists(target_folder):
        os.makedirs(target_folder)
        print("Step 4: Created target folder")

    backup_folder = os.path.join(target_folder, 'BACKUP')
    if not os.path.exists(backup_folder):
        os.makedirs(backup_folder)
        print("Step 4: Created BACKUP folder")

    # Copy the contents of the 'players' folder to the target location
    for root, dirs, files in os.walk(source_folder):
        for dir in dirs:
            source_subdir = os.path.join(root, dir)
            target_subdir = source_subdir.replace(source_folder, target_folder)
            if not os.path.exists(target_subdir):
                os.makedirs(target_subdir)
                print(f"Step 4: Created target subfolder: {target_subdir}")
        for file in files:
            source_file = os.path.join(root, file)
            target_file = source_file.replace(source_folder, target_folder)
            try:
                shutil.copy2(source_file, target_file)
                print(f"Step 4: Copied {file}")
            except PermissionError as e:
                print(f"Step 4: Skipped copying {file} due to PermissionError: {e}")

    messagebox.showinfo("Success", "IHACK MW19 AIO Tool has completed the tasks.\nThanks to codUPLOADER, -coldwarbilly")

root = tk.Tk()
root.title("IHACK MW19 AIO Tool")  # Change the title
root.geometry("800x600")  # Set window size to 800x600

# Update the background image path using a relative path
background_image = Image.open(os.path.join(base_dir, "Capture.jpg"))
background_image = background_image.resize((800, 600), Image.LANCZOS)  # Use Image.LANCZOS
background_photo = ImageTk.PhotoImage(background_image)

background_label = tk.Label(root, image=background_photo)
background_label.place(relwidth=1, relheight=1)

frame = tk.Frame(root)
frame.place(relx=0.5, rely=1, anchor='s')  # Move to the bottom center

select_button = tk.Button(frame, text="Select Call of Duty MW19 Folder", command=select_folder)
select_button.grid(row=0, column=0, padx=10)

folder_entry = tk.Entry(frame, width=50)
folder_entry.grid(row=0, column=1, padx=10)

process_button = tk.Button(frame, text="Process Folders", command=process_folders)
process_button.grid(row=0, column=2, padx=10)

root.mainloop()
