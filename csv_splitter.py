import pandas
import os
import tkinter.messagebox

def user_cancel_overwrite(filePath):
    """
    If file exists prompts user if thwy want to overwrite

    Returns false either when the file does not exist or the user decides to overwrite

    Returns true if the user does not want to overwrite
    """
    if os.path.exists(filePath): 
        return not tkinter.messagebox.askyesno("Overwrite File",
            f"{filePath} eksisterer allerede, vil du overskrive?"
        )

    return False

def read_csv_file(fileName):
    """
    Leser csv fil og reurnerer som 2 dimensional liste
    """
    values = []

    with open(fileName, "r") as f:
        for line in f.readlines():
            values.append([str(i).replace('"', '') for i in line.split(";")])
    
    for row in values:
        row.pop()

    return values

def reformat_into_billed(saveDir, fileName):
    if user_cancel_overwrite(saveDir): return
    
    csvFile = read_csv_file(fileName)

    print(csvFile)