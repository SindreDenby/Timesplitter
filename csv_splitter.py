from unicodedata import name
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

def reformat_into_company_billed(saveDir, fileName):
    if user_cancel_overwrite(saveDir): return
    
    csvFile = read_csv_file(fileName)
    companies = get_companies(csvFile)

    df = pandas.DataFrame(companies)

    if not save_dataframe_as_excel(df, saveDir): return

    tkinter.messagebox.showinfo("Konvertert", "Filen er lagret i " + saveDir)