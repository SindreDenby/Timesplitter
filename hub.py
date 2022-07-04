import os
import tkinter as tk
import tkinter.messagebox
import configurator
import csv_splitter
from tkinter import filedialog

appdata_dir = (os.getenv('APPDATA')).replace("\\", "/") + "/Timesplitter/config/"

"""
Use: py -3.9-64 -m PyInstaller --distpath ./pyinstoutput/dist --workpath ./pyinstoutput/build --clean -w hub.py
"""

class Hub_UI: 
    def __init__(self):
        root = tk.Tk()
        
        mainFrame = tk.Frame(root, padx=30, pady=30)
        mainFrame.grid()

        curCol = 0

        # Ansatte Btn
        tk.Button(mainFrame,
            text=".csv -> Ansatte timer.xlsx",
            command= lambda: csv_splitter.reformat_into_employees(self.get_file_save_dir(), self.get_file_dir())
        ).grid(row=0, column= curCol)

        # Prosjekter til timer Btn
        tk.Button(mainFrame,
            text=".csv -> Prosjekt timer.xlsx",
            command= lambda: csv_splitter.reformat_into_projects(self.get_file_save_dir(), self.get_file_dir())
        ).grid(row=1, column= curCol)

        # Prosjekter til fakturert
        tk.Button(mainFrame,
            text=".csv -> Kunder fakturert.xlsx",
            command= lambda: csv_splitter.reformat_into_company_billed(self.get_file_save_dir(), self.get_file_dir())
        ).grid(row=2, column= curCol)

        tk.Button(mainFrame,
            text=".csv -> Snitt pris.xlsx",
            command= lambda: csv_splitter.reformat_into_average_hourly(self.get_file_save_dir(), self.get_file_dir())
        ).grid(row=3, column= curCol)

        # Config Btn
        tk.Button(mainFrame,
            text="Config",
            command=configurator.main,
            bg="#4287f5"
        ).grid(row=4, column=curCol)

        curCol += 1

        self.setFileBtn = tk.Button(mainFrame,
            text="Set file",
            command=self.set_file,
            bg= "#03fca9"
        )
        self.setFileBtn.grid(row=0, column=curCol)

        self.setDirBtn = tk.Button(mainFrame,
            text="Set Save Directory",
            command=self.set_save_dir,
            bg= "#03fc6b"
        )
        self.setDirBtn.grid(row=1, column= curCol)

        # Filename frame
        fileNameFrame = tk.Frame(mainFrame)
        fileNameFrame.grid(row=2, column=curCol)
        fileNameFrame.grid()

        tk.Label(fileNameFrame, text="Filnavn:").grid(row=0, column=0)

        self.fileNameEntry = tk.Entry(fileNameFrame, justify=tk.RIGHT)
        self.fileNameEntry.grid(row=0, column=1)

        tk.Label(fileNameFrame, text=".xlsx").grid(row=0, column=2)

        root.mainloop()

    def set_save_dir(self):
        dir = filedialog.askdirectory()
        if dir != "":
            self.saveDir = dir
            self.setDirBtn.config(text=self.saveDir)

    def set_file(self):
        file = filedialog.askopenfilename()
        if file != "":
            self.fileDir = file
            self.setFileBtn.config(text=self.fileDir)

    def get_file_save_dir(self):
        try:
            return f'{self.saveDir}/{self.fileNameEntry.get()}.xlsx'
        except AttributeError: 
            tkinter.messagebox.showinfo("File error", "Fil er ikke valgt")

    def get_file_dir(self):
        try:
            return self.fileDir
        except AttributeError: 
            tkinter.messagebox.showinfo("File error", "Lagrinspunkt er ikke valgt")

project_types = ['projects.json', 'admin.json', 'ekstern.json', 'intern.json', 'fastpris.json']

def first_time_setup():
    # print("First time setup")
    os.makedirs(appdata_dir, exist_ok=True)

    for type in project_types:
        with open(appdata_dir + type, "w+") as f:
            f.write("[]")

def main():
    if not os.path.exists(appdata_dir): 
        first_time_setup()
        
    Hub_UI()
    pass

if __name__ == '__main__':
    main()