import json
import tkinter.messagebox
import pandas
import os

appdata_dir = (os.getenv('APPDATA')).replace("\\", "/") + "/Timesplitter/config/"

def get_projects(excelFile, includeType = False):
    """
    Returnerer en liste med prosjekter som dicts
    """
    availableProjects = get_available_projects(excelFile)
    save_as(appdata_dir + "projects.json", availableProjects)
    projects = []
    projectTypes = get_project_types()
    for curProject in availableProjects:
        projectIndex = find_project(excelFile, curProject)

        i = 0
        notFound = True
        while notFound:
            i+= 1
            # if "Sum" in excelFile[projectIndex + i][1]:
            if is_sum(excelFile[projectIndex + i][1]):
                sum = excelFile[projectIndex + i][2]
                notFound = False

        projectType = get_project_type(excelFile[projectIndex][1], projectTypes)

        if includeType: 
            projects.append({
                'name': excelFile[projectIndex][1],
                projectType: float(sum.replace(",", ".")),
                'type': projectType
            })
        else:
            projects.append({
                'name': excelFile[projectIndex][1],
                projectType: float(sum.replace(",", "."))
            })

    return (projects)


project_types = ['admin.json', 'ekstern.json', 'intern.json', 'fastpris.json']
def get_project_types():
    """
    Returnerer: 
    [
        {
            'projectType': "admin",
            'projects': ["økonomi", "Ledelse", ...]
        },
        ...
    ]
    """
    types = []

    for projectType in project_types:
        
        f = open(appdata_dir + projectType, "r")
        types.append ( {
            'projectType': projectType.split('.')[0],
            'projects': json.loads(f.read())
        })
        f.close
    return types

def get_project_type(projectName, types):
    """
    Tar inn prosjekt navnet og prosjekt type listen ( get_project_types() )

    Returnerer prosjekt typen
    """
    for type in types:
        if projectName in type['projects']:
            return type['projectType']
    return "Ikke valgt"

def find_project(list, project):
    """
    Returnerer indeksen av prosjektet
    """
    for i, row in enumerate(list):
        if row[1] == project:
            return i
    print("Not found:", project)

def get_available_projects(list):
    """
    Henter ut alle prosjektnavnene i filen
    """
    availableProjects = []

    for i, row in enumerate(list):
        if row[1] == "Aktivitetssammendrag":
            availableProjects.append(list[i - 2][1])

    return availableProjects

def read_excel_file(fileName):
    """
    Returnerer et excel ark som en 2 dimensional liste
    """
    try: 
        excelFile = pandas.read_excel(fileName)
    except ValueError:
        tkinter.messagebox.showerror("Error", "Feil filformat. Må være .xlsx fil.")

    values = []

    for i, row in enumerate(excelFile.values):
        list = row.tolist()

        for j, col in enumerate(list):
            list[j] = str(col)

        values.append(list)

    return values

def save_as(filename, data):
    os.makedirs(os.path.dirname(filename), exist_ok=True)
    f = open(filename, "w+")
    f.write(json.dumps(data, indent=2))
    f.close()

def get_project_by_name(name, projects):

    for project in projects:
        if project['name'] == name:
            return project

    print(f"Project {name} not in list")
    return None

def get_employee_names(excelFile):
    employees = []

    for i, row in enumerate(excelFile) :
        if is_sum(row[1]): 
            if excelFile[i + 1][1] != 'nan':
                employees.append(excelFile[i + 1][1])

    return list(dict.fromkeys(employees))

def create_employees_shell(employeNames):
    """
    Returnerer en liste med ansatte uten sumerte timer
    """
    employees = []

    for name in employeNames:
        employees.append({
            'name': name,
            'admin': 0,
            'ekstern': 0,
            'fastpris': 0,
            'intern': 0,
            'Ikke valgt': 0
        })

    return employees

def get_employees_data(names, excelFile, projects):
    """
    Returnerer en liste med ansatte og deres timer
    """
    employees_data = create_employees_shell(names)
    projectTypes = get_project_types()

    for nameNr, name in enumerate(names):
        for rowIndex, row in enumerate(excelFile):
            if row[1] == name:
                projectName = get_line_project(rowIndex ,excelFile, projects)
                projectType = get_project_type(projectName, projectTypes)
                employees_data[nameNr][projectType] += get_employee_instance_sum(rowIndex, excelFile)

    return employees_data

def get_employee_instance_sum(instanceRowNr, excelFile):

    i = instanceRowNr + 1

    while (not is_sum(excelFile[i][1])):
        i += 1
    
    # print(f"{excelFile[instanceRowNr][1]} {excelFile[i][1]} {instanceRowNr} at {i} ")

    return float(excelFile[i][2].replace(",", "."))

def get_line_project(rowIndex, excelFile, projects):
    """
    Finner prosjektnavnet til linjen av index
    """
    projectNames = [i['name'] for i in projects]

    i = rowIndex
    searching = True
    while searching:
        i-= 1
        if excelFile[i][1] in projectNames:
            searching = False
            return excelFile[i][1]

def is_sum(val):
    return  val == "Sum NOK" or val == "Sum "

def user_cancel_overwrite(filePath):
    """
    If file exists prompts user if thwy want to overwrite

    Returns false either when the file does not exist or the user decides to overwrite

    Returns true if the user does not want to overwrite
    """
    if os.path.exists(filePath): 
        return not tkinter.messagebox.askyesno("Overwrite File", f"{filePath} eksisterer allerede, vil du overskrive?")

    return False

def save_dataframe_as_excel(dataframe, saveDir):
    """
    Returns True if completes without error
    """

    try: 
        dataframe.to_excel(saveDir, index=False)
        return True
    except PermissionError: 
        tkinter.messagebox.showerror("Error", "Permission Error: Filen er åpen i et annet program eller er utiljengelig.")
        return False
    except ValueError: 
        tkinter.messagebox.showerror("File error", "Filnavn er ikke valgt")
        return False

def reformat_into_projects(saveDir, fileName):
    if user_cancel_overwrite(saveDir): return

    excelFile = read_excel_file(fileName)
    
    # with open('out.json', 'w') as f:
    #     f.write(json.dumps(excelFile, indent=4))

    projects = get_projects(excelFile)

    df = pandas.DataFrame(projects)

    if not save_dataframe_as_excel(df, saveDir): return

    tkinter.messagebox.showinfo("Konvertert", "Filen er lagret i " + saveDir)

def reformat_into_employees(saveDir, fileName):
    if user_cancel_overwrite(saveDir): return

    excelFile = read_excel_file(fileName)

    projects = get_projects(excelFile, includeType=True)

    employee_names = get_employee_names(excelFile)

    employees = get_employees_data(employee_names, excelFile, projects)

    df = pandas.DataFrame(employees)

    if not save_dataframe_as_excel(df, saveDir): return

    tkinter.messagebox.showinfo("Konvertert", "Filen er lagret i " + saveDir)

def main():
    pass

if __name__ == '__main__':
    main()