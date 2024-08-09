import os
import tkinter as tk
from tkinter import ttk, PhotoImage
from PIL import Image, ImageTk
import win32com.client

# Folder containing the .lnk shortcut files
SHORTCUTS_FOLDER = r"C:\Users\James\Downloads\shortcuts"

class KioskApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Program Kiosk")
        self.attributes('-fullscreen', True)  # Enable full-screen mode

        self.bind("<Home>", self.exit_fullscreen)  # Bind the Escape key to exit full-screen

        self.programs = self.load_programs()
        self.display_programs()

    def exit_fullscreen(self, event=None):
        self.attributes('-fullscreen', False)

    def load_programs(self):
        programs = []
        shell = win32com.client.Dispatch("WScript.Shell")

        for shortcut_file in os.listdir(SHORTCUTS_FOLDER):
            if shortcut_file.endswith(".lnk"):
                shortcut_path = os.path.join(SHORTCUTS_FOLDER, shortcut_file)
                shortcut = shell.CreateShortCut(shortcut_path)
                program_path = shortcut.Targetpath

                # Use the shortcut's name (without the .lnk extension)
                shortcut_name = os.path.splitext(shortcut_file)[0]

                program_folder = os.path.dirname(program_path)
                cover_image_path = os.path.join(program_folder, "cover.jpg")
                description_path = os.path.join(program_folder, "description.txt")

                if os.path.exists(cover_image_path):
                    cover_image = Image.open(cover_image_path)
                else:
                    cover_image = None

                if os.path.exists(description_path):
                    with open(description_path, "r") as desc_file:
                        description = desc_file.read()
                else:
                    description = "No description available."

                programs.append({
                    "name": shortcut_name,  # Use the shortcut name instead of the program name
                    "path": program_path,
                    "cover_image": cover_image,
                    "description": description,
                })

        return programs

    def display_programs(self):
        frame = ttk.Frame(self)
        frame.pack(fill=tk.BOTH, expand=True)

        for program in self.programs:
            self.create_program_widget(frame, program)

    def create_program_widget(self, parent, program):
        widget_frame = ttk.Frame(parent)
        widget_frame.pack(pady=10, padx=10, fill=tk.X)

        if program["cover_image"]:
            img = program["cover_image"].resize((100, 100), Image.LANCZOS)
            img = ImageTk.PhotoImage(img)
            cover_label = ttk.Label(widget_frame, image=img)
            cover_label.image = img  # Keep a reference
            cover_label.pack(side=tk.LEFT, padx=5)

        text_frame = ttk.Frame(widget_frame)
        text_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        name_label = ttk.Label(text_frame, text=program["name"], font=("Arial", 16))
        name_label.pack(anchor=tk.W)

        desc_label = ttk.Label(text_frame, text=program["description"], font=("Arial", 10), wraplength=400)
        desc_label.pack(anchor=tk.W)

        open_button = ttk.Button(widget_frame, text="Open", command=lambda: self.open_program(program["path"]))
        open_button.pack(side=tk.RIGHT)

    def open_program(self, program_path):
        print(f"Attempting to open: {program_path}")  # Debugging line
        try:
            os.startfile(program_path)
        except Exception as e:
            print(f"Failed to open {program_path}: {e}")

if __name__ == "__main__":
    app = KioskApp()
    app.mainloop()
