import os
import win32com.client
import time
import argparse as ag
import logging

#====================UDF's====================

def createPDFfolder(base_path, folder_name="PDFs"):
    # Full path for the initial folder
    folder_path = os.path.join(base_path, folder_name)
    
    # If folder already exists, find the next available name
    counter = 0
    while os.path.exists(folder_path):
        folder_path = os.path.join(base_path, f"{folder_name}({counter})")
        counter += 1

    # Create the new folder
    os.makedirs(folder_path)
    print("="*10)
    print(f"PDF folder created: {folder_path}")
    return folder_path

def convertDocmToPDF(input_path, output_path):
    # Initialize Word application
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
        print("PDF created!")
    except Exception as e:
        print(f"Error: {e}")
    #finally:
        # Quit Word application
        word.Quit()

def gettingFilenames(base_directory):
    #collecting the names of '.docm' files in the folder
    filenames=[]
    names=[]
    for files in os.listdir(base_directory):
        if files.endswith('.docm'):
            print (files)
            filenames.append(files)
            names.append(files.strip('.docm'))
    if not filenames:
        print("No files in the directory!")
    return filenames

import logging

# Configure logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s', datefmt="%Y-%m-%d")

def applyFormatting(index, filepath):
    # Configure logging
    logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')    
    
    # File paths
    word_file = rf"{filepath}"   

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
    print(unique_macro_name)

    # Replace the macro name inside the VBA code.  Crucial step.
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


#=====================Main=====================

import sys 

def mymain(path):

    #base_directory=r'C:\Users\HP\OneDrive\Documents\Word Macro Prob'
    print(f'in here! {path}')
    print(f"__name__:{__name__}")
    #sys.exit()
    
    base_directory=rf'{path}'
    filenames=gettingFilenames(base_directory)

    for index,value in enumerate(filenames):
        print("="*10)
        print(f"Index:{index}")
        print(f"Processing: {value}")
        #print(item= os.path.join(base_directory,item) )
        src_path=os.path.join(base_directory,f'{value}')    
        applyFormatting(index,src_path)

    PDFfolder=createPDFfolder(base_directory)
    for file in filenames:
        print("="*10)
        print(file)
        src_path=os.path.join(base_directory,f"{file}")
        dst_path=os.path.join(PDFfolder,f"{file.split('.')[0]}")
        convertDocmToPDF( rf"{src_path}", rf"{dst_path}" )

#pyinstaller --onefile --hidden-import=_tkinter --add-binary "D:\ProgramFiles\Anaconda\pkgs\tk-8.6.14-h0416ee5_0\Library\bin\tcl86t.dll;." --add-binary "D:\ProgramFiles\Anaconda\pkgs\tk-8.6.14-h0416ee5_0\Library\bin\tk86t.dll;." CodeGUI-v1.py