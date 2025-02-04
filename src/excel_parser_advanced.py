import os
import csv
import tkinter as tk
from tkinter import filedialog
from methods import data_retrieval as dr

# begin prompting user for relevant values

# get user input for the sheet
sheet = input("Input the name of the excel sheet you wish to search:")

# get user input for cells
cells = input("Enter the cells you want to search, separated by commas: ")

#convert input into a list
cells = [value.strip() for value in cells.split(',')]

# get user input for directory

loading_directory = dr.select_load_directory()
saving_directory = dr.select_save_directory()
#gets the excel files
files = [
    f for f in os.listdir(loading_directory)
    if f.endswith('.xlsx') or f.endswith('.xls')
]
#turns them into usable filepaths for method
excel_paths = [loading_directory + "/" + i for i in files]

dr = dr()

#gets usable CTCHC data from excel spreadsheet

data = [dr.get_CTCHC_data(i, cells, sheet) for i in excel_paths]

#change this to change the name of the file
file_type = 'misc'

#writes excel spreadsheet data into a csv file
folder_name = loading_directory.split('/')[len(loading_directory.split('/')) -
                                           1]
with open(
        saving_directory + folder_name + '_' + file_type + '.csv',
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

    financing = ['lender', 'source', 'term', 'interest', 'funds']

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

    contact_person_during_apps = [
        'Company Name', 'Street Address', 'City', 'State', 'Zip',
        'Contact Person', 'Phone', 'Email'
    ]

    GP = [
        'GP Name', 'GP Role', 'Street Address', 'City', 'State', 'Zip',
        'Contact Person', 'Ownership Interest', 'Phone', 'Email',
        'Profit/Nonprofit'
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

    developer = [
        'Developer', 'Address', 'City, State, Zip', 'Contact', 'Phone', 'Email'
    ]

    architect = [
        'Architect', 'Address', 'City, State, Zip', 'Contact', 'Phone', 'Email'
    ]

    CPA = ['CPA', 'Address', 'City, State, Zip', 'Contact', 'Phone', 'Email']

    bond_issuer = [
        'Bond_issuer', 'Address', 'City, State, Zip', 'Contact', 'Phone',
        'Email'
    ]

    prop_mgmt = [
        'Prop Mgmt', 'Address', 'City, State, Zip', 'Contact', 'Phone', 'Email'
    ]

    subsidy_info = [
        'Approval Date_1', 'Source_1', 'If Section 8_1:', 'Percentage_1',
        'Units Subsidized_1', 'Amount Per Year_1', 'Total Subsidy_1',
        'Term (in years)_1', 'Approval Date_2', 'Source_2', 'If Section 8_2:',
        'Percentage_2', 'Units Subsidized_2', 'Amount Per Year_2',
        'Total Subsidy_2', 'Term (in years)_2'
    ]

    pre_existing_subsidies = [
        'Sec221(d)(3) BMIR', 'HUD Sec 236', 'If Section 236, IRP?', 'RHS 538',
        'HUD Section 8', 'If Section 8', 'HUD SHP',
        'Will the subsidy continue?', 'If yes, amount', 'RHS 514', 'RHS 515',
        'RHS 521', 'State/Local', 'Rent Sup / RAP', 'Other (specify)', 'Amount'
    ]

    soft_cost = [
        'Land Cost or Value', 'Demolition', 'Legal',
        'Land Lease Rent Prepayment', 'Total Land Cost or Value',
        'Off-Site Improvements', 'Total Acquisition Cost',
        'Predevelopment Interest/Holding Cost',
        'Assumed, Accrued Interest on Existing Debt',
        'Excess Purchase Price Over Appraisal', ''
    ]

    service_amenities = ['Service Amenities']

    num_beds = ['Num Beds']

    hard_cost_contingency = ["Hard Cost Contingency"]

    management_fee = ['Management Fee']

    fields = ['App Number'] + financing
    writer.writerow(fields)
    # for writing a list of lists
    for project in data:
        writer.writerows(project)
    # for writing a single table
    # writer.writerows(data)
print("Processing complete!")
