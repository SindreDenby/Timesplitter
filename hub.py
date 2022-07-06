import os
import tkinter as tk
import tkinter.messagebox
import configurator
import csv_splitter
import adigo_icon
import tooltip
from tkinter import filedialog

appdata_dir = (os.getenv('APPDATA')).replace("\\", "/") + "/Timesplitter/config/"

class Hub_UI: 
    def __init__(self):
        root = tk.Tk()
        
        root.title("Adigo Financial Analysis System")

        adigo_icon.set_base64_icon(adigo_icon.icon, root)

        mainFrame = tk.Frame(root, padx=30, pady=30)
        mainFrame.grid()

        curCol = 0

        tk.Button(mainFrame,
            text="Kjør alle",
            command= lambda: self.execute_all(),
            bg = "#0aff74"
        ).grid(row=0, column= curCol)

        # Ansatte Btn
        curRow = 1


        # Export knapper
        btns = []
        for exportKey in csv_splitter.export_types :
            btns.append(
                tk.Button(mainFrame,
                    text=csv_splitter.export_types[exportKey]['name'],
                    command= lambda i = csv_splitter.export_types[exportKey]: 
                        csv_splitter.reformat(self.get_file_save_dir(), self.get_csv_file(), i),
                )
            )
            btns[len(btns) - 1].grid(row=curRow, column=curCol)
            tooltip.CreateToolTip(btns[len(btns) - 1], csv_splitter.export_types[exportKey]['description'])
            curRow += 1

        # Config Btn
        tk.Button(mainFrame,
            text="Config",
            command=configurator.main,
            bg="#4287f5"
        ).grid(row=curRow, column=curCol)

        curCol += 1

        # Set file frame
        setFileFrame = tk.Frame(mainFrame)
        setFileFrame.grid(row=0, column=curCol)
        setFileFrame.grid()

        tk.Label(setFileFrame, text=".csv fil:").grid(row=0, column=0)

        self.csvFileInput = tk.Entry(setFileFrame, justify=tk.RIGHT, width=40)
        self.csvFileInput.grid(row=0, column=1)

        tk.Button(setFileFrame,
            text="...",
            command=self.set_file,
        ).grid(row=0, column=2)

        # Set directory frame

        tk.Label(setFileFrame, text="Lagres som:").grid(row=1, column=0)

        self.saveDirInput = tk.Entry(setFileFrame, justify=tk.RIGHT, width=40)
        self.saveDirInput.grid(row=1, column=1)

        tk.Button(setFileFrame,
            text="...",
            command=self.set_save_dir,
        ).grid(row=1, column=2)

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