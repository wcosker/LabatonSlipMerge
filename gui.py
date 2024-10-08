# Third party
import tkinter as tk
from tkinter import scrolledtext
from PIL import Image, ImageTk
# Locally created packages
from merging import run_merges
from validate import validate_paths_and_explain
from utils import get_excel_file_path,read_excel_sheets,check_version,resource_path
from constants import CURRENT_VERSION, ICON_PATH, LOGO_PATH, BG_COLOR, BUTTON_COLOR, HEADER_COLOR, API_LINK, REPO_LINK

def main_gui():
    def on_run_merge():
        file_path = get_excel_file_path()
        if file_path:
            config, filelist = read_excel_sheets(file_path)
            preserve_metadata = preserve_var.get()

            # Run the merge process
            run_merges(config, filelist, preserve_metadata, output_text)

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

    preserve_var = tk.BooleanVar()
    preserve_checkbox = tk.Checkbutton(frame, text="Preserve PDF Metadata", variable=preserve_var, font=('Helvetica', 12), bg=BG_COLOR, fg=BUTTON_COLOR)
    preserve_checkbox.grid(row=1, column=0, padx=10, pady=10, sticky='w')  # Positioned below the Run Merge button

    validate_button = tk.Button(frame, text="Validate Config", command=on_validate, width=20, bg=BUTTON_COLOR, fg="white", font=('Helvetica', 12, 'bold'))
    validate_button.grid(row=0, column=1, padx=10, pady=10)

    # Add a space for the logo image to the right of the buttons
    logo_image = Image.open(resource_path(LOGO_PATH))
    logo_size = (200, 200)  # Adjust the size of the logo as needed
    logo_image.thumbnail(logo_size, Image.LANCZOS)
    logo_photo = ImageTk.PhotoImage(logo_image)

    logo_label = tk.Label(frame, image=logo_photo, bg=BG_COLOR)
    logo_label.image = logo_photo  # Keep a reference to the image
    logo_label.grid(row=0, column=2, rowspan=2, padx=20, pady=10)  # Positioned next to the buttons

    # Label for the Output Log
    output_label = tk.Label(root, text="Output Log", font=('Helvetica', 14, 'bold'), bg=BG_COLOR, fg=BUTTON_COLOR)
    output_label.pack(padx=20, pady=(0, 0), anchor='w')  # Positioned right above the scrolled text box


    # Scrolled Text box for output display
    output_text = scrolledtext.ScrolledText(root, width=95, height=25, wrap=tk.WORD, font=('Helvetica', 12), bg=HEADER_COLOR, fg=BUTTON_COLOR)
    output_text.pack(padx=20, pady=10)

    # check_version(CURRENT_VERSION,API_LINK,REPO_LINK)
    root.mainloop()