import json
import pandas
import os
import csv
import tkinter.messagebox

appdata_dir = (os.getenv('APPDATA')).replace("\\", "/") + "/Timesplitter/config/"

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

def get_company_names(csvFile):
    names = list(dict.fromkeys([i[1] for i in csvFile]))

    names.remove('Kunde')
    names.remove('')

    return names

def create_companies_shell(companyNames):
    companies = []

    for name in companyNames:
        companies.append({
            'name': name,
            'timer': 0,
            'fakt. timer': 0
        })
    
    return companies
    
def get_company_index_by_name(companyName, companies):
    """
    Return index of the given company name
    """
    for companyIndex, company in enumerate(companies) :
        if company['name'] == companyName:
            return companyIndex

    return None

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

def get_companies(csvFile):
    companyNames = get_company_names(csvFile)

    companies = create_companies_shell(companyNames)

    for rowIndex, row in enumerate(csvFile) :
        if row[1] in companyNames:
            companyName = row[1]
            companyIndex = get_company_index_by_name(companyName, companies)

            # print(f'RowNr: {rowIndex}, timer: {float(row[16].replace(",", "."))} ')
            companies[companyIndex]['timer'] += float(row[16].replace(",", ".")) 
            companies[companyIndex]['fakt. timer'] += float(row[19].replace(",", ".")) 
        

    return companies

def save_as(filename, data):
    # os.makedirs(os.path.dirname(filename), exist_ok=True)
    f = open(filename, "w+")
    f.write(json.dumps(data, indent=2))
    f.close()

def create_projects_shell(projectNames):
    projects = []

    for name in projectNames:
        projects.append({
            'name': name,
            'admin': 0,
            'ekstern': 0,
            'intern': 0,
            'fastpris': 0,
            'ikke valgt': 0
        })
    
    return projects

def get_available_projects(csvFile):
    projects = []

    for row in csvFile:
        projects.append(row[6])

    projects = list(dict.fromkeys(projects))
    projects.remove('Prosjektnavn')

    return projects

def get_project_type(projectName, types):
    """
    Tar inn prosjekt navnet og prosjekt type listen ( get_project_types() )

    Returnerer prosjekt typen
    """
    for type in types:
        if projectName in type['projects']:
            return type['projectType']
    return "ikke valgt"    

def get_project_index_by_name(projectName, projects):
    """
    Return index of the given project name
    """
    for projectIndex, project in enumerate(projects) :
        if project['name'] == projectName:
            return projectIndex

    return None

def get_projects(csvFile):
    """
    Returnerer en liste med prosjekter som dicts
    """
    projectNames = get_available_projects(csvFile)
    save_as(appdata_dir + "projects.json", projectNames)
    projects = create_projects_shell(projectNames)
    projectTypes = get_project_types()

    run = False
    for row in csvFile:
        if run:
            if row[6] in projectNames:
                projectIndex = get_project_index_by_name(row[6], projects)
                projectType = get_project_type(row[6], projectTypes)
                if len(row) > 25: print(row)

                projects[projectIndex][projectType] += float(row[16].replace(",", '.'))

        run = True

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

def read_csv_file(fileName):
    """
    Leser csv fil og reurnerer som 2 dimensional liste
    """
    with open(fileName, newline='') as f:
        reader = csv.reader(f, delimiter=";")
        data = [tuple(row) for row in reader]

    # save_as("test.json", data)

    return data

def reformat_into_company_billed(saveDir, fileName):
    if user_cancel_overwrite(saveDir): return
    
    csvFile = read_csv_file(fileName)

    companies = get_companies(csvFile)
    df = pandas.DataFrame(companies)

    if not save_dataframe_as_excel(df, saveDir): return

    tkinter.messagebox.showinfo("Konvertert", "Filen er lagret i " + saveDir)

def reformat_into_projects(saveDir, fileName):
    if user_cancel_overwrite(saveDir): return

    csvFile = read_csv_file(fileName)

    projects = get_projects(csvFile)
    df = pandas.DataFrame(projects)

    if not save_dataframe_as_excel(df, saveDir): return

    tkinter.messagebox.showinfo("Konvertert", "Filen er lagret i " + saveDir)

