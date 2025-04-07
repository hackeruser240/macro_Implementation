import tkinter as tk
from tkinter import filedialog
import os
import win32com.client
import time
import logging
import sys
#import argparse as ag #you may remove this

def create_pdf_folder(base_path, folder_name="PDFs"):
    # ... (rest of your code, but replace print with logging)
    folder_path = os.path.join(base_path, folder_name)
    
    # If folder already exists, find the next available name
    counter = 0
    while os.path.exists(folder_path):
        folder_path = os.path.join(base_path, f"{folder_name}({counter})")
        counter += 1

    # Create the new folder
    os.makedirs(folder_path)
    logging.info(f"PDF folder created: {folder_path}")
    log_message(f"PDF folder created: {folder_path}")
    return folder_path

def convert_docm_to_pdf(input_path, output_path):
    # ... (rest of your code, but replace print with logging)
    time.sleep(3)
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False  # Run in the background

    try:
        # Open the Word document
        doc = word.Documents.Open(os.path.abspath(input_path))
        
        # Save as PDF (FileFormat=17 means PDF)
        doc.SaveAs(os.path.abspath(output_path), FileFormat=17)
      
        # Close document
        doc.Close(SaveChanges=False)
        logging.info("PDF created!")
        log_message("PDF created!")
    except Exception as e:
        logging.error(f"Error: {e}")
    #finally:
        # Quit Word application
        word.Quit()

def getting_filenames(base_directory):
    # ... (rest of your code, but replace print with logging)
    filenames=[]
    names=[]
    for files in os.listdir(base_directory):
        if files.endswith('.docm'):
            logging.info (files)
            filenames.append(files)
            names.append(files.strip('.docm'))
    if not filenames:
        logging.info("No files in the directory!")
    return filenames

def apply_formatting(index, filepath):
    # ... (rest of your code, but replace print with logging)
    # Configure logging
    logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
    
    # File paths
    word_file = rf"{filepath}"
    print(f"location of word file: {word_file}")   

    try:
        macro_file_path = os.path.join(os.path.dirname(filepath), "TXT files", "Macro_DirectCertifyQA.txt")
        with open(macro_file_path, "r", encoding="utf-8") as file:
            macro_code = file.read()
            #logging.info(f"Macro file successfully loaded from: {macro_file_path}")
            #return macro_code
    except FileNotFoundError:
        logging.error(f"Macro file not found at: {macro_file_path}")
        return None
    except Exception as e:
        logging.error(f"Error loading macro file: {e}")
        return None

    # Generate a unique macro name
    unique_macro_name = f"DirectCertifyQA_{os.path.splitext(os.path.basename(filepath))[0]}_{index}"  # Include index for uniqueness
    logging.info(unique_macro_name)

    # Replace the macro name inside the VBA code. Crucial step.
    macro_code = macro_code.replace("DirectCertifyQA", unique_macro_name)

    # Open Word
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True  # Show Word for debugging

    # Open the Word document
    doc = word.Documents.Open(word_file)

    # Flag to track if the macro has been injected
    macro_injected = False

    retries = 2
    for attempt in range(retries):
        try:
            logging.info(f"Processing file: {filepath}, attempt: {attempt + 1}")  # Log file and attempt
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = True
            doc = word.Documents.Open(word_file)
            vb_component = word.VBE.ActiveVBProject.VBComponents("ThisDocument")
            vb_component.CodeModule.AddFromString(macro_code)
            word.Run(unique_macro_name)
            doc.Save()
            logging.info(f"Successfully applied formatting: {unique_macro_name}") #log success
            log_message(f"Successfully applied formatting: {unique_macro_name}")
            word.Quit()
            break
        except Exception as e:
            logging.error(f"Error processing {filepath} (Attempt {attempt + 1}/{retries}): {e}") #log error
            if attempt < retries - 1:
                time.sleep(5)
            else:
                logging.error(f"Failed to process {filepath} after {retries} attempts.")
                word.Quit() #quit word in finally
                return  # Add a return here to stop processing the file
        finally:
            try:
                word.Quit()
            except Exception as e:
                logging.error(f"Error quitting Word: {e}") # Log if word.quit fails.

def start_conversion():
    base_directory = path_entry.get()  # Get path from the entry
    if not base_directory:
        log_message("Error: Please select a directory.", "error") #use log_message
        return

    path = rf"{base_directory}\TXT files\Macro_DirectCertifyQA.txt"

    if os.path.exists(path):
        print("The file exists!")
    else:
        print("The macro file does NOT exist, or its folder does not exist!")
        sys.exit()
    
    #Check the files in the selected directory
    filenames = getting_filenames(base_directory)

    for index, value in enumerate(filenames):
        log_message("=" * 10)
        log_message(f"Index: {index}")
        log_message(f"Processing: {value}")
        src_path = os.path.join(base_directory, f'{value}')
        apply_formatting(index, src_path)

    pdf_folder = create_pdf_folder(base_directory)
    for file in filenames:
        log_message("=" * 10)
        log_message(file)
        src_path = os.path.join(base_directory, f"{file}")
        dst_path = os.path.join(pdf_folder, f"{file.split('.')[0]}")
        convert_docm_to_pdf(rf"{src_path}", rf"{dst_path}")

    log_message(f"Succesfully processed {len(filenames)} file(s)!")

def browse_directory():
    directory = filedialog.askdirectory()

    if directory:  # Check if a directory was actually selected
        directory = directory.replace("/", "\\")  # Replace forward slashes with backslashes
        path_entry.delete(0, tk.END)
        path_entry.insert(0, directory)

    path_entry.delete(0, tk.END)
    path_entry.insert(0, directory)

def log_message(message, level="info"):
    # Log the message and display it in the text widget
    if level == "error":
        logging.error(message)
        output_text.tag_config("error", foreground="red")
        output_text.insert(tk.END, message + "\n", "error")
    else:
        logging.info(message)
        output_text.insert(tk.END, message + "\n")
    output_text.see(tk.END)  # Scroll to the end

#=====================Main GUI=====================


print(f"__name__:{__name__}")

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

window = tk.Tk()
window.title("Word to PDF Converter")

# Path Entry
path_label = tk.Label(window, text="Enter Directory Path:")
path_label.pack(pady=10)
path_entry = tk.Entry(window, width=50)
path_entry.pack(pady=5)

# Browse Button
browse_button = tk.Button(window, text="Browse", command=browse_directory)
browse_button.pack(pady=5)

# Start Conversion Button
start_button = tk.Button(window, text="Start Conversion", command=start_conversion)
start_button.pack(pady=10)

# Output Text Widget
output_text = tk.Text(window, height=20, width=80)
output_text.pack(pady=10)

# Main loop
window.mainloop()


#pyinstaller --onefile --hidden-import=_tkinter --add-binary "D:\ProgramFiles\Anaconda\pkgs\tk-8.6.14-h0416ee5_0\Library\bin\tcl86t.dll;." --add-binary "D:\ProgramFiles\Anaconda\pkgs\tk-8.6.14-h0416ee5_0\Library\bin\tk86t.dll;." CodeGUI_EXE.py