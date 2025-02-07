import os
import csv
from methods import data_retrieval

#directory for excel files

directory = "C:/Users/cameronshaw/Documents/Affordable Research/All Applications Since 2022/2024_Applications"
saving_directory = "C:/Users/cameronshaw/Documents/Affordable Research/Stats and Documents/Data/raw/"
dr = data_retrieval()
#gets the excel files
files = os.listdir(directory)
#turns them into usable filepaths for method
excel_paths = [directory + "/" + i for i in files]

#gets usable CTCHC data from excel spreadsheet

data = [dr.get_CTCHC_data_simple(i) for i in excel_paths]

#change this to change the name of the file
file_type = 'permanent_financial'

#writes excel spreadsheet data into a csv file
folder_name = directory.split('/')[len(directory.split('/')) - 1]
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

    unit_mix = ['0 Beds', '1 Bed', '2 Beds', '3 Beds', '4 Beds']

    fields = ['App Number'] + unit_mix
    writer.writerow(fields)
    # for writing a list of lists
    for project in data:
        writer.writerows(project)
    # for writing a single table
    # writer.writerows(data)
print("Processing complete!")
