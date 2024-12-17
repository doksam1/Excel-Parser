import os
import openpyxl
import csv
from methods import data_retrieval

#directory for excel files
directory = "C:/Users/cameronshaw/Documents/Affordable Research/All Applications Since 2022/second_round_2024"
#gets the excel files
files = os.listdir(directory)
#turns them into usable filepaths for method
excel_paths = [directory + "/" + i for i in files]

#gets usable CTCHC data from excel spreadsheet

data = [data_retrieval.get_CTCHC_data(i) for i in excel_paths]

#writes excel spreadsheet data into a csv file
with open('round2_app_data.csv', 'w', newline='') as file:
    writer = csv.writer(file)
    app = [
        'App Number', 'Units', 'NRSF', 'Prevailing Wage', 'Parking SF',
        'Stories', "Buildings", 'Architect', 'Parking Spaces', 'Subterranean?'
    ]
    budg = ['land', 'hard', 'soft', 'arch', 'finance', 'dev']

    agent = [
        'App Number', 'CPA', 'Address', 'City, State, Zip', 'Contact Person',
        'Phone', 'Email'
    ]

    GC2 = [
        'Project Number', 'Site Work', 'Structures', 'Requirements',
        'Overhead', 'Profit', 'Prevailing Wages', 'Insurance',
        'Third Party Management', 'Constr Costs'
    ]

    expenses_per_unit = ['App Number', 'Units', 'Total Annual Expenses']

    misc_income = [
        'App Number', 'Annual Laundry Income', 'Annual Vending Machine Income',
        'Annual Interest Income', 'Other - Specified', 'Other Income',
        'Total Misc Income', 'Total Annual Potential Gross Income'
    ]

    fields = misc_income

    writer.writerow(fields)
    # for writing a list of lists
    # for project in data:
    # writer.writerows(project)
    # for writing a single table
    writer.writerows(data)
