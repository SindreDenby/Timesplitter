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

def get_object_index_by_name(projectName, projects):
    """
    Return index of the given project name
    """
    for objectIndex, project in enumerate(projects) :
        if project['name'] == projectName:
            return objectIndex

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
                projectIndex = get_object_index_by_name(row[6], projects)
                projectType = get_project_type(row[6], projectTypes)

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

def get_employee_names(csvFile):
    return list(dict.fromkeys([i[13] for i in csvFile[1:]]))


def get_employees_data(names, csvFile):
    employees = create_projects_shell(names)
    projectTypes = get_project_types()

    for row in csvFile[1:]:
        curNameIndex = get_object_index_by_name(row[13], employees)
        projectType = get_project_type(row[6], projectTypes)

        employees[curNameIndex][projectType] += float(row[16].replace(",", "."))

    return employees    

def check_document_invalid(excelFile):
    """
    Returns true if document is invalid
    """
    if excelFile[0][0] == "Kundenummer": 
        return False

    tkinter.messagebox.showerror("Invalid", "Filen som leses av er feil eller korrupt")
    return True

def format_month_year(date):
    """
    Reformats date to month and year

    Ex: "2022-04-22" returns "2022-04"
    """
    return "-".join(date.split("-")[:2]) 

def get_monthly_hour_average(csvFile):
    months = []
    projectTypes = get_project_types()

    for row in csvFile[1:]:
        if get_project_type(row[6], projectTypes) == 'ekstern':
            date = format_month_year(row[15])

            if date not in [i['name'] for i in months]:
                months.append({
                    'name': date,
                    'timer': 0,
                    'snitt': 0,
                    'timeprisSum': 0
                }) 
            
            dateIndex = get_object_index_by_name(date, months)
            months[dateIndex]['timer'] += float(row[16].replace(",", "."))
            months[dateIndex]['timeprisSum'] += float(row[20].replace(",", ".")) * float(row[16].replace(",", "."))

    for i, month in enumerate(months):
        months[i]['snitt'] = month['timeprisSum'] / month['timer']
        months[i].pop('timeprisSum')

    return months

def read_csv_file(fileName):
    """
    Leser csv fil og reurnerer som 2 dimensional liste
    """
    try:

     with open(fileName, newline='') as f:
        reader = csv.reader(f, delimiter=";")
        data = [tuple(row) for row in reader]

    except UnicodeDecodeError:
        tkinter.messagebox.showerror("Invalid", "Filen som leses av er feil eller korrupt")

    except PermissionError:
        tkinter.messagebox.showerror("Permission Error", "Fil er ikke tiljengelig/åpen i et annet program")

    # save_as("test.json", data)

    return data

def reformat_into_company_billed(saveDir, fileName):
    if user_cancel_overwrite(saveDir): return
    
    csvFile = read_csv_file(fileName)

    if check_document_invalid(csvFile): return

    companies = get_companies(csvFile)
    df = pandas.DataFrame(companies)

    if not save_dataframe_as_excel(df, saveDir): return

    tkinter.messagebox.showinfo("Konvertert", "Filen er lagret i " + saveDir)

def reformat_into_employees(saveDir, fileName):
    if user_cancel_overwrite(saveDir): return

    csvFile = read_csv_file(fileName)

    if check_document_invalid(csvFile): return

    projects = get_projects(csvFile)
    employee_names = get_employee_names(csvFile)

    employees = get_employees_data(employee_names, csvFile)
    df = pandas.DataFrame(employees)

    if not save_dataframe_as_excel(df, saveDir): return

    tkinter.messagebox.showinfo("Konvertert", "Filen er lagret i " + saveDir)

def reformat_into_projects(saveDir, fileName):
    if user_cancel_overwrite(saveDir): return

    csvFile = read_csv_file(fileName)

    if check_document_invalid(csvFile): return

    projects = get_projects(csvFile)
    df = pandas.DataFrame(projects)

    if not save_dataframe_as_excel(df, saveDir): return

    tkinter.messagebox.showinfo("Konvertert", "Filen er lagret i " + saveDir)

def reformat_into_average_hourly(saveDir, fileName):
    if user_cancel_overwrite(saveDir): return

    csvFile = read_csv_file(fileName)

    if check_document_invalid(csvFile): return

    get_projects(csvFile)
    hourlyAverage = get_monthly_hour_average(csvFile)
    df = pandas.DataFrame(hourlyAverage)

    if not save_dataframe_as_excel(df, saveDir): return

    tkinter.messagebox.showinfo("Konvertert", "Filen er lagret i " + saveDir)

