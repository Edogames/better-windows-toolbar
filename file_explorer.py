import os
import tkinter as tk
from tkinter import messagebox
import win32com.client
import subprocess
import psutil
import ctypes
from PIL import Image, ImageTk
import win32gui
import win32con
import argparse
from screeninfo import get_monitors

# Function to check if the file is a media file
def is_media_file(file_path):
    media_extensions = ('.jpg', '.jpeg', '.png', '.gif', '.mp4', '.avi', '.mkv', '.mov')
    file_extension = os.path.splitext(file_path)[1].lower()
    return file_extension in media_extensions

# Function to get all .lnk, .exe, and media files in the folder
def get_files_in_folder(folder_path):
    files = []
    for file in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file)
        file_extension = os.path.splitext(file)[1].lower()

        # Check for .lnk shortcut files
        if os.path.isfile(file_path) and file_extension == '.lnk':
            target, icon_location = get_target_from_shortcut(file_path)
            if target:
                files.append({'type': 'link', 'name': file, 'target': target, 'icon_location': icon_location})

        # Check for .exe files
        elif os.path.isfile(file_path) and file_extension == '.exe':
            files.append({'type': 'app', 'name': file, 'path': file_path})

        # Check for media files
        elif os.path.isfile(file_path) and is_media_file(file_path):
            files.append({'type': 'media', 'name': file, 'path': file_path})

    return files

# Function to get the target path from a shortcut (.lnk) file
def get_target_from_shortcut(lnk_file_path):
    try:
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(lnk_file_path)
        return shortcut.TargetPath, shortcut.IconLocation
    except Exception as e:
        print(f"Error extracting target from shortcut {lnk_file_path}: {e}")
        return None, None

# Function to check if a process is running based on the executable path
def is_process_running(executable_path):
    for proc in psutil.process_iter(attrs=['pid', 'name']):
        try:
            if executable_path.lower() in proc.info['name'].lower():
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    return False

# Function to open the shortcut target when clicked
def open_shortcut_target(target_path):
    try:
        subprocess.Popen(target_path)
        print(f"Launching: {target_path}")
    except Exception as e:
        print(f"Error launching {target_path}: {e}")

# Function to open media files when clicked
def open_media_file(media_path):
    try:
        subprocess.Popen(media_path, shell=True)
        print(f"Opening media file: {media_path}")
    except Exception as e:
        print(f"Error opening media file {media_path}: {e}")

# Function to open executable files when clicked
def open_exe_file(exe_path):
    try:
        subprocess.Popen(exe_path)
        print(f"Opening executable file: {exe_path}")
    except Exception as e:
        print(f"Error opening executable file {exe_path}: {e}")

# Function to apply hover and click effects without smooth transition
def apply_hover_and_click_effect(button, initial_size, hover_size, click_size, initial_color, hover_color, click_color):
    def on_hover(event):
        button.configure(width=hover_size[0], height=hover_size[1], bg=hover_color)

    def on_click(event):
        button.configure(width=click_size[0], height=click_size[1], bg=click_color)

    def on_leave(event):
        button.configure(width=initial_size[0], height=initial_size[1], bg=initial_color)

    button.bind("<Enter>", on_hover)
    button.bind("<Leave>", on_leave)
    button.bind("<ButtonPress>", on_click)

# Function to get the system theme (dark or light)
def get_system_theme():
    try:
        registry_key = r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        reg_value = ctypes.windll.user32.GetSysColor(30)  # Get the color value for the theme
        if reg_value == 0:
            return "light"
        else:
            return "dark"
    except Exception as e:
        print(f"Error detecting system theme: {e}")
        return "light"

# Function to apply dark or light theme to the UI
def apply_theme(root, frame, buttons, theme):
    if theme == "dark":
        root.configure(bg="#2C2F38")
        frame.configure(bg="#2C2F38")
        for button in buttons:
            button.configure(bg="#3D4758", fg="white", relief="flat", font=("Segoe UI", 12))
    else:
        root.configure(bg="white")
        frame.configure(bg="white")
        for button in buttons:
            button.configure(bg="#E0E0E0", fg="black", relief="flat", font=("Segoe UI", 12))

# Function to extract the icon from the shortcut
def get_icon_image(icon_location):
    try:
        icon_path, icon_index = icon_location.split(',')
        icon_index = int(icon_index)
        hicon = win32gui.ExtractIcon(0, icon_path, icon_index)
        if hicon:
            icon = Image.frombytes("RGBA", (32, 32), win32gui.GetIconInfo(hicon))
            return icon
        else:
            print(f"Error extracting icon: Invalid handle for {icon_path}")
            return None
    except Exception as e:
        print(f"Error extracting icon: {e}")
        return None

def move_window_to_cursor(root):
    root.update_idletasks()  # Essential: Update window dimensions

    window_width = root.winfo_width()
    window_height = root.winfo_height()

    # Get primary monitor info
    monitors = get_monitors()
    if monitors:
        primary_monitor = monitors[0]  # Assuming the first monitor is primary
        screen_width = primary_monitor.width
        screen_height = primary_monitor.height
        x = root.winfo_pointerx() - window_width // 2
        y = root.winfo_pointery() - window_height // 2

        padding = 10

        # Ensure window stays within screen bounds
        x = max(padding, min(x, screen_width - window_width - padding))
        y = max(padding, min(y, screen_height - window_height - padding))

        root.geometry(f"+{x}+{y}")
    else:
        print("No monitors found.")

# Main function to create the GUI
def create_gui(folder_path):
    root = tk.Tk()
    root.title("File Explorer")
    root.geometry("400x500")  # Fixed window size
    root.resizable(False, False)

    theme = get_system_theme()

    frame = tk.Frame(root)
    frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

    files = get_files_in_folder(folder_path)

    if not files:
        messagebox.showerror("Error", "No files found in the folder.")
        return

    buttons = []

    for file in files:
        name_without_extension = os.path.splitext(file['name'])[0]

        # Prepare label with type information
        if file['type'] == 'media':
            label = f"(video) {name_without_extension}" if file['path'].endswith(('.mp4', '.avi', '.mkv')) else f"(audio) {name_without_extension}"
        elif file['type'] == 'app':
            label = f"(app) {name_without_extension}"
        else:
            label = f"(link) {name_without_extension}"

        if file['type'] == 'link':
            target_path = file['target']
            icon_location = file['icon_location']
            icon = get_icon_image(icon_location)
            if icon:
                icon = icon.resize((32, 32), Image.Resampling.LANCZOS)
                icon = ImageTk.PhotoImage(icon)
                button = tk.Button(frame, text=label, image=icon, width=30, height=2, compound="left", command=lambda target=target_path: open_shortcut_target(target))
                button.image = icon
            else:
                button = tk.Button(frame, text=label, width=30, height=2, command=lambda target=target_path: open_shortcut_target(target))

        elif file['type'] == 'media':
            media_path = file['path']
            button = tk.Button(frame, text=label, width=30, height=2, command=lambda media=media_path: open_media_file(media))

        elif file['type'] == 'app':
            exe_path = file['path']
            button = tk.Button(frame, text=label, width=30, height=2, command=lambda exe=exe_path: open_exe_file(exe))

        initial_size = (30, 2)
        hover_size = (35, 3)
        click_size = (28, 2)
        initial_color = "#E0E0E0"
        hover_color = "#D3D3D3"
        click_color = "#B0B0B0"

        if theme == "dark":
            initial_color = "#3D4758"
            hover_color = "#4d596e"
            click_color = "#2c3340"

        apply_hover_and_click_effect(button, initial_size, hover_size, click_size, initial_color, hover_color, click_color)

        button.pack(fill=tk.X, pady=5)
        buttons.append(button)

    apply_theme(root, frame, buttons, theme)

    # Move the window based on cursor position and keep it within bounds
    move_window_to_cursor(root)

    root.mainloop()

# Run the script
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Open shortcut files (.lnk) from a given folder.")
    parser.add_argument("folder_path", type=str, help="Path to the folder containing .lnk shortcut files.")
    args = parser.parse_args()

    create_gui(args.folder_path)
