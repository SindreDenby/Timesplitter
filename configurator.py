import json
import tkinter as tk
import os

project_types = ['projects.json', 'admin.json', 'løpende.json', 'intern.json', 'fastpris.json', 'salg.json', 'bedriftsutvikling.json']
appdata_dir = (os.getenv('APPDATA')).replace("\\", "/") + "/Timesplitter/config/"

def create_listBox(box, row, col):
    listboxFrame = tk.Frame(box)
    listboxFrame.grid(row=row, column=col)

    listbox = tk.Listbox(listboxFrame, width=30, height=30, selectmode='extended')

    listbox.pack(side = tk.LEFT, fill = tk.BOTH)

    scrollbar = tk.Scrollbar(listboxFrame)

    scrollbar.pack(side = tk.RIGHT, fill = tk.BOTH)
        
    listbox.config(yscrollcommand = scrollbar.set)

    scrollbar.config(command = listbox.yview)

    return listbox

def insert_into_listbox(listbox: tk.Listbox, list):
    listbox.delete(0, tk.END)

    for i in list:
        listbox.insert(tk.END, i)

class config_ui:
    def __init__(self):

        self.projects = self.read_files()

        self.projects[0]['projects'] = self.get_unselected_projects()

        root = tk.Tk()
        
        root.title("Adigo Financial Analysis System Configuration Tool")
        
        curCol = 0

        mainFrame = tk.Frame(root, padx=30, pady=30)
        mainFrame.grid()

        self.project_list_boxes = []
        for i, project in enumerate(self.projects):
            tk.Label(mainFrame,
                text= project['name'].split('.')[0]
            ).grid(row=0, column=curCol)

            self.project_list_boxes.append(create_listBox(mainFrame, 1, curCol))

            tk.Button(mainFrame, 
                command= lambda i = i: self.move_selected(i),
                text="Flytt hit"
            ).grid(row=2, column=curCol)
            
            curCol += 1

        self.update_lists()
        root.mainloop()

    def read_files(self):
        """
        Leser av posjekt filene og returnerer dem som en dict
        """
        loaded_project_types = []

        for i in project_types:
            f = open(appdata_dir + i, 'r')
            loaded_project_types.append({
                'name': i,
                'projects': json.loads(f.read())
            })
            f.close()
        return loaded_project_types

    def get_selected_elements(self):
        """
        Returnerer: (listboxNr: int, list of selected indexes: (1, 3, 6))
        """

        for listboxNr, listbox in enumerate(self.project_list_boxes) :
            for selected in listbox.curselection():
                return (listboxNr, listbox.curselection())

    def move_selected(self, box_index):
        """
        Moves selcted to box_index
        """
        selectedElements = self.get_selected_elements()

        insertBoxIndex = box_index
        oldBoxIndex = selectedElements[0]

        selectedProjectNames = [self.projects[oldBoxIndex]['projects'][i] for i in selectedElements[1]]

        for projectName in selectedProjectNames:
            self.projects[insertBoxIndex]['projects'].append(projectName)
            self.projects[oldBoxIndex]['projects'].remove(projectName)

        self.save_config()

        self.update_lists()

    def update_lists(self):
        for i, listbox in enumerate(self.project_list_boxes):
            insert_into_listbox(listbox, self.projects[i]['projects'])
        
    def save_config(self):
        for projectNr in range(len(self.projects) - 1):
            f = open(appdata_dir + self.projects[projectNr + 1]['name'], 'w')
            f.write(json.dumps(list(dict.fromkeys(self.projects[projectNr + 1]['projects'])), indent=2))
            f.close()

        self.projects = self.read_files()

        self.projects[0]['projects'] = self.get_unselected_projects()

        self.update_lists()
        # tkinter.messagebox.showinfo("Lagret", "Config er lagret")

    def get_unselected_projects(self):
        unselected_projects = self.projects[0]['projects'].copy()

        selected_projects = []

        for i in range(len(project_types) - 1):
            selected_projects.extend(self.projects[i + 1]['projects'])
       
        for project in self.projects[0]['projects']:
   
            if project in selected_projects:
                unselected_projects.remove(project)
                    
        return unselected_projects


def main():
    config_ui()

if __name__ == '__main__':
    main()