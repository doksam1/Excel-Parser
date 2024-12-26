import os
import csv
from methods import data_retrieval

#directory for excel files
directory = "C:/Users/cameronshaw/Documents/Affordable Research/All Applications Since 2022/first_round_2024"
#gets the excel files
files = os.listdir(directory)
#turns them into usable filepaths for method
excel_paths = [directory + "/" + i for i in files]

#gets usable CTCHC data from excel spreadsheet

data = [data_retrieval.get_CTCHC_data(i) for i in excel_paths]

#writes excel spreadsheet data into a csv file
folder_name = directory.split('/')[len(directory.split('/')) - 1]
with open(
        folder_name + '_consultant.csv',
        'w',
        encoding='utf-8',
        newline='',
) as file:
    writer = csv.writer(file)
    app = [
        'Units', 'NRSF', 'Prevailing Wage', 'Parking SF', 'Stories',
        "Buildings", 'Architect', 'Parking Spaces', 'Subterranean?'
    ]
    budg = ['land', 'hard', 'soft', 'arch', 'finance', 'dev']

    agent = [
        'CPA', 'Address', 'City, State, Zip', 'Contact Person', 'Phone',
        'Email'
    ]

    GC2 = [
        'Project Number', 'Site Work', 'Structures', 'Requirements',
        'Overhead', 'Profit', 'Prevailing Wages', 'Insurance',
        'Third Party Management', 'Constr Costs'
    ]

    expenses_per_unit = ['Units', 'Total Annual Expenses']

    misc_income = [
        'Annual Laundry Income', 'Annual Vending Machine Income',
        'Annual Interest Income', 'Other - Specified', 'Other Income',
        'Total Misc Income', 'Total Annual Potential Gross Income'
    ]

    operating_expenses = [
        'admin', 'management', 'utilities', 'payroll & payroll taxes',
        'insurance', 'maintenance', 'Other-specify', 'Other'
    ]

    attorneys = [
        'Attorney', 'Address', 'City, State, Zip', 'Contact', 'Phone', 'Email'
    ]

    consultants = [
        'Consultant', 'Address', 'City, State, Zip', 'Contact', 'Phone',
        'Email'
    ]

    GC = [
        'General Contractor', 'Address', 'City, State, Zip', 'Contact',
        'Phone', 'Email'
    ]

    tax_professional = [
        'Tax Professional', 'Address', 'City, State, Zip', 'Contact', 'Phone',
        'Email'
    ]

    energy_consultant = [
        'Energy_Consultant', 'Address', 'City, State, Zip', 'Contact', 'Phone',
        'Email'
    ]

    investor = [
        'Investor', 'Address', 'City, State, Zip', 'Contact', 'Phone', 'Email'
    ]

    market_analyst = [
        'Analyst', 'Address', 'City, State, Zip', 'Contact', 'Phone', 'Email'
    ]

    appraiser = [
        'Appraiser', 'Address', 'City, State, Zip', 'Contact', 'Phone', 'Email'
    ]

    CNA_consultant = [
        'Consultant', 'Address', 'City, State, Zip', 'Contact', 'Phone',
        'Email'
    ]

    fields = [
        'App Number'
    ] + tax_professional + GC + energy_consultant + investor + market_analyst + appraiser + CNA_consultant

    writer.writerow(fields)
    # for writing a list of lists
    # for project in data:
    # writer.writerows(project)
    # for writing a single table
    writer.writerows(data)
