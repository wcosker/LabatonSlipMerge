# Third party
import tkinter as tk
from tkinter import scrolledtext
from PIL import Image, ImageTk
# Locally created packages
from merging import run_merges
from validate import validate_paths_and_explain
from utils import get_excel_file_path,read_excel_sheets,check_version,resource_path
from constants import CURRENT_VERSION, ICON_PATH, BG_COLOR, BUTTON_COLOR, TEXT_COLOR, HEADER_COLOR, API_LINK, REPO_LINK

def main_gui():
    def on_run_merge():
        file_path = get_excel_file_path()
        if file_path:
            config, filelist = read_excel_sheets(file_path)
            scrub_metadata = scrub_var.get()

            # Run the merge process
            run_merges(config, filelist, scrub_metadata, output_text)

    def on_validate():
        file_path = get_excel_file_path()
        if file_path:
            config, filelist = read_excel_sheets(file_path)
            validate_paths_and_explain(config, filelist, output_text)
    
    # GUI setup
    root = tk.Tk()
    root.title("Labaton Slip Merge")
    root.geometry("800x600")
    root.configure(bg=BG_COLOR)

    # Set window icon
    icon_image = Image.open(resource_path(ICON_PATH))
    icon_size = (32, 32)
    icon_image.thumbnail(icon_size, Image.LANCZOS)
    icon_photo = ImageTk.PhotoImage(icon_image)
    root.iconphoto(False, icon_photo)

    # Frame for buttons and options
    frame = tk.Frame(root, padx=20, pady=20, bg=BG_COLOR)
    frame.pack(fill=tk.X)

    # Buttons and checkboxes
    run_button = tk.Button(frame, text="Run Merge", command=on_run_merge, width=20, bg=BUTTON_COLOR, fg="white", font=('Helvetica', 12, 'bold'))
    run_button.grid(row=0, column=0, padx=10, pady=10)

    scrub_var = tk.BooleanVar()
    scrub_checkbox = tk.Checkbutton(frame, text="Scrub PDF Metadata", variable=scrub_var, font=('Helvetica', 12), bg=BG_COLOR, fg=TEXT_COLOR)
    scrub_checkbox.grid(row=1, column=0, padx=10, pady=10, sticky='w')  # Positioned below the Run Merge button

    validate_button = tk.Button(frame, text="Validate Config", command=on_validate, width=20, bg=BUTTON_COLOR, fg="white", font=('Helvetica', 12, 'bold'))
    validate_button.grid(row=0, column=1, padx=10, pady=10)

    # Scrolled Text box for output display
    output_text = scrolledtext.ScrolledText(root, width=95, height=25, wrap=tk.WORD, font=('Helvetica', 12), bg=HEADER_COLOR, fg=TEXT_COLOR)
    output_text.pack(padx=20, pady=10)

    check_version(CURRENT_VERSION,API_LINK,REPO_LINK)
    root.mainloop()