import os
import tkinter as tk
import tkinter.messagebox
import configurator
import csv_splitter
import adigo_icon
import tooltip
from tkinter import filedialog, ttk

appdata_dir = (os.getenv('APPDATA')).replace("\\", "/") + "/Timesplitter/config/"

class Hub_UI: 
    def __init__(self):
        root = tk.Tk()
        
        root.title("Adigo Financial Analysis System")

        adigo_icon.set_base64_icon(adigo_icon.icon, root)

        style = ttk.Style(root)

        style.theme_use('vista')

        mainFrame = tk.Frame(root, padx=30, pady=30)
        mainFrame.grid()

        # Set file frame
        setFileFrame = tk.Frame(mainFrame)
        setFileFrame.grid(row=0, column=0, columnspan= 2)
        setFileFrame.grid()

        ttk.Label(setFileFrame, text=".csv fil:").grid(row=0, column=0)

        self.csvFileInput = ttk.Entry(setFileFrame, justify=tk.RIGHT, width=40)
        self.csvFileInput.grid(row=0, column=1)

        ttk.Button(setFileFrame,
            text="...",
            command=self.set_file,
            width=6
        ).grid(row=0, column=2)

        ttk.Label(setFileFrame, text="Lagres som:").grid(row=1, column=0)

        self.saveDirInput = ttk.Entry(setFileFrame, justify=tk.RIGHT, width=40)
        self.saveDirInput.grid(row=1, column=1)

        ttk.Button(setFileFrame,
            text="...",
            command=self.set_save_dir,
            width=6
        ).grid(row=1, column=2)

        # Config Btn
        

        ttk.Label(mainFrame,
            text="Timeoversikt.csv"
        ).grid(row= 1, column=0)

        curRow = 2

        # Export knapper
        btns = []
        for exportKey in csv_splitter.timeoversikt_exports :
            btns.append(
                ttk.Button(mainFrame,
                    text=csv_splitter.timeoversikt_exports[exportKey]['name'],
                    command= lambda i = csv_splitter.timeoversikt_exports[exportKey]: 
                        csv_splitter.reformat_timeoversikt(self.get_file_save_dir(), self.get_csv_file(), i),
                    width=17
                )
            )
            btns[len(btns) - 1].grid(row=curRow, column=0)
            tooltip.create(btns[len(btns) - 1], csv_splitter.timeoversikt_exports[exportKey]['description'])
            curRow += 1


        curRow = 2
        ttk.Label(mainFrame,
            text="Prosjektstatus.csv"
        ).grid(row= 1, column=1)

        for exportKey in csv_splitter.prosjektstatus_exports :
            btns.append(
                ttk.Button(mainFrame,
                    text=csv_splitter.prosjektstatus_exports[exportKey]['name'],
                    command= lambda i = csv_splitter.prosjektstatus_exports[exportKey]: 
                        csv_splitter.reformat_prosjektstatus(self.get_file_save_dir(), self.get_csv_file(), i),
                    width=17
                )
            )
            btns[len(btns) - 1].grid(row=curRow, column=1)
            tooltip.create(btns[len(btns) - 1], csv_splitter.prosjektstatus_exports[exportKey]['description'])
            curRow += 1

        ttk.Button(mainFrame,
            text="Velg prosjekt typer",
            command=configurator.main,
            width=30,
            # bg="#4287f5"
        ).grid(row=100, column=0, columnspan=2)

        root.mainloop()

    def set_save_dir(self):
        dir = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=(('Excel file', '*.xlsx'), ("All Files", "*.*")))
        if dir != "":
            self.saveDirInput.delete(0, tk.END)
            self.saveDirInput.insert(0, dir)

    def set_file(self):
        file = filedialog.askopenfilename(defaultextension=".csv", filetypes=(('Comma seperated file', '*.csv'), ("All Files", "*.*")))
        if file != "":
            self.csvFileInput.delete(0, tk.END)
            self.csvFileInput.insert(0, file)

    def get_file_save_dir(self):
        entryVal = self.saveDirInput.get()
        if entryVal == "":
            tkinter.messagebox.showinfo("File error", "Filens lagringspunkt er ikke valgt")
            return

        return entryVal

    def execute_all(self):
        pass

    def get_csv_file(self):
        entryVal = self.csvFileInput.get()
        if entryVal == "":
            tkinter.messagebox.showinfo("File error", "Csv fil er ikke valgt")
            return
        
        return entryVal

project_types = [
    'projects.json',
    'admin.json',
    'løpende.json',
    'intern.json',
    'fastpris.json',
    'salg.json',
    'bedriftsutvikling.json'
]

def first_time_setup():

    os.makedirs(appdata_dir, exist_ok=True)

    files = os.listdir(appdata_dir)

    for type in project_types:
        if type not in files:
            with open(appdata_dir + type, "w+") as f:
                f.write("[]")

def main():
    
    first_time_setup()
        
    Hub_UI()

if __name__ == '__main__':
    main()