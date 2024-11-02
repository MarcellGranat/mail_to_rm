import subprocess
from tqdm import tqdm
import os

def doc_to_pdf(input_file):
    # based on https://github.com/jeongwhanchoi/convert-ppt-to-pdf
    applescript = f'''
    on run
        set theOutput to ""
        set inputFile to "{input_file}"
        
        tell application "Microsoft Word" -- work on version 15.15 or newer
            launch
            set t to inputFile as string
            if t ends with ".doc" or t ends with ".docx" then
                set pdfPath to my makeNewPath(inputFile)
                try
                    open POSIX file inputFile
                    set activeDoc to active document
                    -- Export the active document to PDF
                    save as activeDoc file name (POSIX file pdfPath) file format format PDF
                    close active document saving no
                    set theOutput to pdfPath
                on error errMsg
                    -- handle errors here if needed
                end try
            end if
        end tell
        tell application "Microsoft Word"
            quit
        end tell
        
        return theOutput
    end run

    on makeNewPath(f)
        set t to f as string
        if t ends with ".docx" then
            return (POSIX path of (text 1 thru -6 of t)) & ".pdf"
        else
            return (POSIX path of (text 1 thru -5 of t)) & ".pdf"
        end if
    end makeNewPath
    '''
    subprocess.run(["osascript", "-e", applescript])

def ppt_to_pdf(input_file):
    # based on https://github.com/jeongwhanchoi/convert-ppt-to-pdf
    applescript = f'''
    on run
        set theOutput to ""
        set inputFile to "{input_file}"
        
        tell application "Microsoft PowerPoint" -- work on version 15.15 or newer
            launch
            set t to inputFile as string
            if t ends with ".ppt" or t ends with ".pptx" then
                set pdfPath to my makeNewPath(inputFile)
                try
                    open POSIX file inputFile
                    set activePres to active presentation
                    -- Export the active presentation to PDF
                    save activePres in (POSIX file pdfPath) as save as PDF
                    close active presentation saving no
                    set theOutput to pdfPath
                on error errMsg
                    -- handle errors here if needed
                end try
            end if
        end tell
        tell application "Microsoft PowerPoint"
            quit
        end tell
        
        return theOutput
    end run

    on makeNewPath(f)
        set t to f as string
        if t ends with ".pptx" then
            return (POSIX path of (text 1 thru -6 of t)) & ".pdf"
        else
            return (POSIX path of (text 1 thru -5 of t)) & ".pdf"
        end if
    end makeNewPath
    '''
    # Run the AppleScript from Python
    subprocess.run(["osascript", "-e", applescript])

def convert_folder(folder = "attachments"):
    for file in tqdm(os.listdir(folder), desc=f"Converting folder: {folder}"):
        if file.endswith(".docx") or file.endswith(".doc"):
            doc_to_pdf(input_file=os.path.abspath(f"{folder}/{file}"))
        elif file.endswith(".pptx") or file.endswith(".ppt"):
            ppt_to_pdf(input_file=os.path.abspath(f"{folder}/{file}"))

if __name__ == "__main__":
    input_file = os.path.abspath("attachments/spederb_EPU_manuscript_20240202_MB.docx")
    doc_to_pdf(input_file)
    input_file = os.path.abspath("attachments/granatm_MonPol - NMG seminar - 03 - Célok, rezsimek, reakciófüggvények.pptx")
    ppt_to_pdf(input_file)
