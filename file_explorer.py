import os
import tkinter as tk
from tkinter import ttk, messagebox
import argparse
import subprocess
import psutil
from PIL import Image, ImageTk
import configparser
import pythoncom
from win32com.client import Dispatch

# --- Helper Functions ---
def is_media_file(file_path):
    media_extensions = ('.jpg', '.jpeg', '.png', '.gif', '.mp4', '.avi', '.mkv', '.mov', '.mp3', '.wav')
    return os.path.splitext(file_path)[1].lower() in media_extensions

def get_files_recursively(folder_path):
    apps = []
    media = []
    print(f"Scanning folder: {folder_path}")
    for root, _, file_names in os.walk(folder_path):
        for file in file_names:
            file_path = os.path.join(root, file)
            file_extension = os.path.splitext(file)[1].lower()

            if os.path.isfile(file_path):
                if file_extension in ('.exe', '.lnk'):  # Include .lnk files as apps
                    apps.append({'type': 'app', 'name': file, 'path': file_path})
                elif is_media_file(file_path):
                    media.append({'type': 'media', 'name': file, 'path': file_path})
    
    print(f"Found {len(apps)} apps and {len(media)} media files.")
    return apps, media

def is_process_running(executable_path):
    for proc in psutil.process_iter(attrs=['pid', 'name']):
        if executable_path.lower() in proc.info['name'].lower():
            return True
    return False

def resolve_lnk(lnk_path):
    """Resolve the target path and properties of a .lnk file."""
    try:
        shell = Dispatch("WScript.Shell")
        shortcut = shell.CreateShortcut(lnk_path)
        target_path = shortcut.TargetPath
        arguments = shortcut.Arguments
        working_directory = shortcut.WorkingDirectory
        description = shortcut.Description  # May contain AppUserModelID for Store apps
        return target_path, arguments, working_directory, description
    except Exception as e:
        print(f"Error resolving .lnk file: {e}")
        return None, None, None, None


def is_store_app(description):
    """Check if a shortcut description indicates a Microsoft Store app."""
    return "!" in description if description else False


def open_file(file_path):
    """Open app or media based on file path."""
    try:
        if file_path.lower().endswith('.lnk'):
            # Resolve .lnk to its target path
            target_path, arguments, working_directory, description = resolve_lnk(file_path)
            if target_path or description:
                if is_store_app(description):
                    # Exclude Microsoft Store app shortcuts
                    print(f"Skipping Microsoft Store app shortcut: {file_path}")
                elif os.path.exists(target_path):
                    # Handle traditional shortcuts
                    print(f"Launching shortcut target: {target_path}")
                    subprocess.Popen(f'"{target_path}" {arguments}', cwd=working_directory, shell=True)
                else:
                    print(f"Invalid shortcut or target not found: {file_path}")
            else:
                print(f"Could not resolve .lnk file: {file_path}")
        elif file_path.lower().endswith('.exe'):
            # Run executable (launch app)
            subprocess.Popen(file_path, shell=True)
            print(f"Launching app: {file_path}")
        else:
            # Fallback for other file types
            os.startfile(file_path)
            print(f"Opening file: {file_path}")
    except Exception as e:
        print(f"Error opening {file_path}: {e}")

def get_file_icon(file_path):
    """Get the app's or media file's icon (or placeholder)"""
    try:
        if file_path.lower().endswith(('jpg', 'jpeg', 'png', 'gif')):
            image = Image.open(file_path)
            image.thumbnail((30, 30))  # Resize for preview
            return ImageTk.PhotoImage(image)
        elif file_path.lower().endswith(('.exe', '.lnk')):
            # Placeholder for executable file (we can use an app icon)
            return ImageTk.PhotoImage(Image.new('RGB', (30, 30), color=(50, 50, 50)))  # Dark square as placeholder
    except Exception as e:
        print(f"Error getting icon for {file_path}: {e}")
        return ImageTk.PhotoImage(Image.new('RGB', (30, 30), color=(50, 50, 50)))  # Dark square as placeholder

# --- Style Parsing ---
def load_style_config(file_path):
    # Check if style_config.ini exists, otherwise fallback to hardcoded values
    if not os.path.exists(file_path):
        print("INI file not found, using hardcoded style values.")
        return {
            "general": {
                "background_color": "#6917d4",
                "sidebar_color": "#46128a",
                "button_color": "#46128a",
                "text_color": "#000000",
                "launcher_title": "My Launcher",
                "title_font_size": 16,
                "button_roundness": 1
            },
            "treeview": {
                "background_color": "#46128a",
                "foreground_color": "#52199c",
                "row_height": 30,
                "field_background_color": "#1a0440",
                "selected_background_color": "#070121",
                "font_size": 12,
                "normal_font_color": "#000000",
                "border_color": "#000000",
            },
            "column": {
                "column_title_background": "#46128a"
            },
            "sidebar": {
                "font_size": 12
            },
            "padding": {
                "inner_padding": 10
            }
        }

    # If INI exists, load the config
    config = configparser.ConfigParser()
    config.read(file_path)

    style = {
        "general": {
            "background_color": config.get("general", "background_color", fallback="#FFFFFF"),
            "sidebar_color": config.get("general", "sidebar_color", fallback="#333333"),
            "button_color": config.get("general", "button_color", fallback="#444444"),
            "text_color": config.get("general", "text_color", fallback="#FFFFFF"),
            "launcher_title": config.get("general", "launcher_title", fallback="All-in-One Launcher"),
            "title_font_size": config.getint("general", "title_font_size", fallback=16),
            "button_roundness": config.getfloat("general", "button_roundness", fallback=1)
        },
        "treeview": {
            "background_color": config.get("treeview", "background_color", fallback="#333333"),
            "foreground_color": config.get("treeview", "foreground_color", fallback="#FFFFFF"),
            "row_height": config.getint("treeview", "row_height", fallback=30),
            "field_background_color": config.get("treeview", "field_background_color", fallback="#2E2E2E"),
            "selected_background_color": config.get("treeview", "selected_background_color", fallback="#555555"),
            "font_size": config.getint("treeview", "font_size", fallback=12),
            "normal_font_color": config.get("treeview", "normal_font_color", fallback="#000000"),
            "border_color": config.get("treeview", "border_color", fallback="#000000"),
        },
        "column": {
            "column_title_background": config.get("column", "column_title_background", fallback="#46128a")
        },
        "sidebar": {
            "font_size": config.getint("sidebar", "font_size", fallback=12)
        },
        "padding": {
            "inner_padding": config.getint("padding", "inner_padding", fallback=10)
        }
    }

    return style

# --- Launcher UI ---
def apply_filter(filter_func, apps, media, file_tree, filters):
    global filtered_files  # Declare as global to update it across functions
    file_tree.delete(*file_tree.get_children())  # Clear the current treeview items
    
    # Filter files based on the selected filter function
    filtered_files = apps + media if filter_func == filters["All"] else (
        [file for file in apps if filter_func(file)] if filter_func == filters["Apps"] else 
        [file for file in media if filter_func(file)]
    )

    # List to store icons and file paths for the double-click action
    cached_files = {}

    # Add the filtered files to the Treeview
    for file in filtered_files:
        status = "N/A"  # Set a simple status instead of checking for running processes
        icon = get_file_icon(file['path'])  # Get icon

        # If no valid icon is returned, set it to a placeholder
        if icon is None:
            icon = ImageTk.PhotoImage(Image.new('RGB', (30, 30), color=(50, 50, 50)))  # Dark square placeholder
        
        # Remove the file extension from the file name
        file_name_without_extension = os.path.splitext(file['name'])[0]
        
        # Insert the filtered file data into the Treeview (without the status)
        item = file_tree.insert("", "end", values=(file['type'], file_name_without_extension), tags=(file['type'],), image=icon)

        # Store the file name and path in the cache (key-value pair)
        cached_files[item] = file['path']  # Use item ID as key

    # Bind double-click event for launching or opening apps/media
    def on_treeview_double_click(event):
        # Get the selected item from the treeview
        item = file_tree.selection()
        if item:
            file_path = cached_files.get(item[0])  # Retrieve the file path using item ID
            if file_path:
                open_file(file_path)

    # Bind the double-click event to the treeview
    file_tree.bind("<Double-1>", on_treeview_double_click)

# --- Design (UI) Part ---
def calculate_luminance(color_hex):
    """Calculate luminance of the background color."""
    color_hex = color_hex.lstrip("#")  # Remove '#' if present
    r, g, b = int(color_hex[0:2], 16), int(color_hex[2:4], 16), int(color_hex[4:6], 16)
    
    # Apply luminance formula (using standard luminance coefficients for RGB)
    luminance = 0.2126 * r + 0.7152 * g + 0.0722 * b
    return luminance

def get_contrasting_text_color(background_color):
    """Return black or white text color based on the background color luminance."""
    luminance = calculate_luminance(background_color)
    if luminance < 128:  # Dark background (luminance < 128)
        return "#FFFFFF"  # White text
    else:  # Light background
        return "#000000"  # Black text

def create_steam_like_launcher(folder_path, style, apps, media):
    root = tk.Tk()
    root.title(style['general']['launcher_title'])
    root.geometry("1000x600")
    root.configure(bg=style['general']['background_color'])

    style_ = ttk.Style()
    style_.theme_use("clam")
    
    # Get contrasting text color for Treeview
    treeview_text_color = get_contrasting_text_color(style['treeview']['background_color'])

    style_.configure("Treeview", 
                    background=style['treeview']['background_color'],
                    foreground=treeview_text_color,
                    rowheight=style['treeview']['row_height'],
                    fieldbackground=style['treeview']['field_background_color'])
    style_.map("Treeview", background=[("selected", style['treeview']['selected_background_color'])])

    # --- Sidebar Frame --- 
    sidebar = tk.Frame(root, bg=style['general']['sidebar_color'], width=200)
    sidebar.pack(side=tk.LEFT, fill=tk.Y)

    # --- Content Area --- 
    content_area = tk.Frame(root, bg=style['general']['background_color'])
    content_area.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

    # --- Treeview with custom columns --- 
    def create_treeview_with_custom_columns(content_area, style):
        file_tree = ttk.Treeview(content_area, columns=("Type", "Name"), show="headings")
        file_tree.heading("Type", text="Type")
        file_tree.heading("Name", text="Name")

        file_tree.column("Type", width=100)
        file_tree.column("Name", width=600)  # Adjusted width for Name column

        style_.configure("Treeview",
                        background=style['treeview']['background_color'],
                        foreground=treeview_text_color,  # Use dynamic contrasting text color
                        rowheight=style['treeview']['row_height'],
                        fieldbackground=style['treeview']['field_background_color'],
                        font=("Helvetica", style['treeview']['font_size']),
                        borderwidth=2,  # Treeview border width
                        relief="solid",
                        bordercolor=style['treeview']['border_color'])  # Treeview border color

        style_.map("Treeview", background=[("selected", style['treeview']['selected_background_color'])])

        # Apply styles to the column headers (titles)
        style_.configure("Treeview.Heading",
                        font=("Helvetica", style['treeview']['font_size'], 'bold'),
                        background=style['column']['column_title_background'],
                        foreground=treeview_text_color)  # Column title text color

        # Disable any hover or click effects on column headers
        style_.map("Treeview.Heading", background=[('active', style['column']['column_title_background'])])

        # Create a tag for the default row text color
        file_tree.tag_configure("normal", foreground=treeview_text_color)  # Row text color

        # Add rows to the Treeview
        file_tree.pack(fill=tk.BOTH, expand=True, padx=style['padding']['inner_padding'], pady=style['padding']['inner_padding'])

        return file_tree

    file_tree = create_treeview_with_custom_columns(content_area, style)

    # --- Sidebar Buttons --- 
    filters = {"All": lambda x: True, "Apps": lambda x: x['type'] == 'app', "Media": lambda x: x['type'] == 'media'}
    
    def create_sidebar_buttons(apps, media, file_tree):
        # Clear existing buttons
        for widget in sidebar.winfo_children():
            widget.destroy()

        # Add sidebar buttons for filters (All, Apps, Media)
        filters = {"All": lambda x: True, "Apps": lambda x: x['type'] == 'app', "Media": lambda x: x['type'] == 'media'}

        for filter_name, filter_func in filters.items():
            button_text_color = get_contrasting_text_color(style['general']['sidebar_color'])

            button = tk.Button(sidebar, 
                            text=filter_name, 
                            command=lambda f=filter_func: apply_filter(f, apps, media, file_tree, filters), 
                            font=(style['sidebar']['font_size']),
                            fg=button_text_color,  # Use dynamic contrasting text color
                            bg=style['general']['sidebar_color'])
            button.pack(fill=tk.X)

        # Add the actual files (app/media) to the Treeview
        for file in apps + media:
            # Add all apps/media to the Treeview as separate rows with appropriate icons and file details
            file_tree.insert("", "end", values=(file['type'], file['name'], "Idle" if file['type'] == 'app' else "N/A"),
                            tags=(file['type'],))
            
        # Step 4: Apply initial filter to display all files
        apply_filter(filters["All"], apps, media, file_tree, filters)


    # Pass the file_tree to the create_sidebar_buttons function
    create_sidebar_buttons(apps, media, file_tree)

    # Step 4: Apply initial filter to display all files
    apply_filter(filters["All"], apps, media, file_tree, filters)

    root.mainloop()

# Run the launcher with the scanned apps and media
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="File Explorer Launcher")
    parser.add_argument("folder", type=str, help="Folder path to scan for applications and media files")
    args = parser.parse_args()

    # Step 1: Scan the folder for apps and media files
    apps, media = get_files_recursively(args.folder)

    # Step 2: Load the style configuration
    style_config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'style_config.ini')
    style = load_style_config(style_config_path)

    # Step 3: Create and launch the UI
    create_steam_like_launcher(args.folder, style, apps, media)
