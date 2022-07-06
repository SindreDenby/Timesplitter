import json
import pandas
import os
import csv
import datetime
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
            'løpende': 0,
            'intern': 0,
            'fastpris': 0,
            'salg': 0,
            'bedriftsutvikling': 0,
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
    projectTypes = read_project_types()

    run = False
    for row in csvFile:
        if run:
            if row[6] in projectNames:
                projectIndex = get_object_index_by_name(row[6], projects)
                projectType = get_project_type(row[6], projectTypes)

                projects[projectIndex][projectType] += float(row[16].replace(",", '.'))

        run = True

    return (projects)

project_types = [
    'admin.json',
    'løpende.json',
    'intern.json',
    'fastpris.json',
    'salg.json',
    'bedriftsutvikling.json'
]

def read_project_types():
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
        f.close()
    return types

def get_employee_names(csvFile):
    return list(dict.fromkeys([i[13] for i in csvFile[1:]]))


def get_employees_data(csvFile):
    names = get_employee_names(csvFile)
    employees = create_projects_shell(names)
    projectTypes = read_project_types()

    for row in csvFile[1:]:
        curNameIndex = get_object_index_by_name(row[13], employees)
        projectType = get_project_type(row[6], projectTypes)

        employees[curNameIndex][projectType] += float(row[16].replace(",", "."))

    return employees    

def check_document_invalid(excelFile, value):
    """
    Returns true if document is invalid
    """
    if excelFile[0][0] == value: 
        return False

    tkinter.messagebox.showerror("Invalid", "Filen som leses av er feil eller korrupt")
    return True

def format_month_year(date):
    """
    Reformats date to month and year

    Ex: "2022-04-22" returns "2022-04"
    """
    return "-".join(date.split("-")[:2]) 

def get_weeks_in_file(csvFile):
    out = list(dict.fromkeys([get_week_number(i[15]) for i in csvFile[1:]]))
    out.sort()
    return out

def get_weeks_hours(csvFile):
    availableWeeks = get_weeks_in_file(csvFile)

    weeks = create_projects_shell(availableWeeks)
    projectTypes = read_project_types()

    for row in csvFile[1:]:
        weekNr = get_week_number(row[15])
        projectType = get_project_type(row[6], projectTypes)
        weekIndex = get_object_index_by_name(weekNr, weeks)

        weeks [weekIndex][projectType] += float(row[16].replace(",","."))

    return weeks

def sum_all_keys(dictionary: dict):
    sum = 0
    for i in list(dictionary.items()):
        try:
            sum += float(i[1])
        except ValueError: pass

    return sum

def get_division_names(csvFile):
    return list(dict.fromkeys([row[11] for row in csvFile[1:]]))

def get_divisions_percentage(csvFile):
    divisionNames = get_division_names(csvFile)
    divisionNames.append("total")
    divisions = create_projects_shell(divisionNames)

    projectTypes = read_project_types()

    totalIndex = get_object_index_by_name('total', divisions)
    for row in csvFile[1:]:
        divisionIndex = get_object_index_by_name(row[11], divisions)
        projectType = get_project_type(row[6], projectTypes)

        divisions[divisionIndex][projectType] += float(row[16].replace(",", "."))
        divisions[totalIndex][projectType] += float(row[16].replace(",", "."))

    for i, division in enumerate(divisions):
        sum = sum_all_keys(division)

        for key in list(division.items())[1:]:
            divisions[i][key[0]] = divisions[i][key[0]] / sum

    return divisions

def get_week_number(date):
    """
    Returns week number of date string

    Ex: "2022-07-06" returns 27
    """
    return datetime.date(*[int(i) for i in date.split("-")]).isocalendar()[1]

def get_monthly_hour_average(csvFile):
    months = []
    projectTypes = read_project_types()

    for row in csvFile[1:]:
        if get_project_type(row[6], projectTypes) == 'løpende':
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

    return data

def reformat(saveDir, fileName, exportType):
    if user_cancel_overwrite(saveDir): return

    csvFile = read_csv_file(fileName)

    if check_document_invalid(csvFile, 'Kundenummer'): return

    get_projects(csvFile)
    data = exportType['function'](csvFile)
    df = pandas.DataFrame(data)

    if "format" in exportType:
        df.rename(columns=exportType['format'], inplace=True)

    if not save_dataframe_as_excel(df, saveDir): return

    tkinter.messagebox.showinfo("Konvertert", "Filen er lagret i " + saveDir)

export_types = {
    'employee_hours':{
        'name': 'Ansatte timer',
        'input': 'timeoversikt',
        'description': "Deller opp ansatte i timer brukt på forskjellige prosjekt typer.",
        'function': get_employees_data
    },
    'project_hours':{
        'name': 'Prosjekt timer',
        'input': 'timeoversikt',
        'description': "Deler opp i timer brukt på prosjekter.",
        'function': get_projects
    },
    'kunder_fakturert':{
        'name': 'Kunder fakturert',
        'input': 'timeoversikt',
        'description': "Deler opp kunder i timer og fakturert timer.",
        'function': get_companies
    },
    'snitt_pris':{
        'name': 'Snitt timepris',
        'input': 'timeoversikt',
        'description': "Deler opp i måndeder med timer brukt og gjennomsnitlig time lønn.",
        'function': get_monthly_hour_average
    },
    'avdeling_fordeling':{
        'name': 'Avdeling timer',
        'input': 'timeoversikt',
        'description': "Deler opp avdelinger i prosent av tid brukt på forskjellige prosjekt typer.",
        'function': get_divisions_percentage
    },
    'time_per_uke':{
        'name': 'Timer uke basis',
        'input': 'timeoversikt',
        'description': "Deler opp i timer brukt på ukentlig basis.",
        'function': get_weeks_hours,
        'format': {'name': 'Uke Nr'}
    },
}