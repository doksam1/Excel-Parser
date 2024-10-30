import os
import openpyxl
import csv

#directory for excel files
directory = "first_round_2024"
#gets the excel files
files = os.listdir(directory)
#turns them into usable filepaths for method
excel_paths = [directory + "/" + i for i in files]


#gets usable CTCHC data from excel spreadsheet
def get_CTCHC_data(file_path):
    # name = 'H18'
    # address = 'I185'
    # project_type = 'D211'
    # consultant_company = 'H319'
    # consultant_contact = 'H322'
    # attorney_firm = 'H295'
    # attorney_contact = 'H298'
    # tax_company = 'H303'
    # tax_contact = 'H306'
    # cpa = 'H311'
    # cpa_contact = 'H314'
    # appraiser = 'H327'
    # appraiser_contact = 'H330'
    # property_manager = 'AA335'
    # property_manager_contact = 'AA338'
    # cells_to_search = [
    #     name, address, project_type, consultant_company, consultant_contact,
    #     attorney_firm, attorney_contact, tax_company, tax_contact, cpa,
    #     cpa_contact, appraiser, appraiser_contact, property_manager,
    #     property_manager_contact
    # ]

    application_data = ['AG437', 'AG442', 'O773', 'T773', 'AG993', 'AG449']

    budget_and_sources = [
        'B12', 'B26', 'B38', 'B42', 'B43', 'B54', 'B62', 'B70', 'B77', 'B79',
        'B80', 'B96', 'B104', 'B13', 'B14', 'B27'
    ]

    cells_to_search = budget_and_sources

    #load excel file
    workbook = openpyxl.load_workbook(file_path, data_only=True)

    #Select the sheet
    app = "Application"
    budg = 'Sources and Uses Budget'
    sheet = workbook[budg]

    #get the values
    cell_data = [sheet[i].value for i in cells_to_search]

    #removes null values
    for j in range(len(cell_data)):
        if cell_data[j] is None:
            cell_data[j] = 'None'

    #per value changes needed
    if isinstance(cell_data[1], str) or isinstance(
            cell_data[2], str) or isinstance(cell_data[3], str):
        cell_data[1] = 0
        cell_data[2] = 0
        cell_data[3] = 0
        print(f'skipped numbers for this:{file_path}')

    land_cost = cell_data[0] + cell_data[13] + cell_data[14]
    hard_cost = cell_data[1] + cell_data[2] + cell_data[9] + cell_data[15]
    soft_cost = cell_data[7] + cell_data[8] + cell_data[10] + cell_data[11]
    architectural_cost = cell_data[3] + cell_data[4]
    finance_cost = cell_data[5] + cell_data[6]
    developer_fee = cell_data[12]

    #per value changes needed

    print(f'returning started:{file_path}')
    to_return = [
        land_cost, hard_cost, soft_cost, architectural_cost, finance_cost,
        developer_fee
        # cell_data[0], cell_data[1] + cell_data[2] * cell_data[3], cell_data[4],
        # cell_data[5]

        # cell_data[0], cell_data[1], cell_data[2],
        # cell_data[3] + ' - ' + cell_data[4],
        # cell_data[5] + ' - ' + cell_data[6],
        # cell_data[7] + ' - ' + cell_data[8],
        # cell_data[9] + ' - ' + cell_data[10],
        # cell_data[11] + ' - ' + cell_data[12],
        # cell_data[13] + ' - ' + cell_data[14]
    ]
    print(f'data returned:{file_path}')
    return to_return


data = [get_CTCHC_data(i) for i in excel_paths]

#writes excel spreadsheet data into a csv file
with open('CTCHC2.csv', 'w', newline='') as file:
    writer = csv.writer(file)
    app = ['Units', 'NRSF', 'Prevailing Wage', 'Parking SF']
    budg = ['land', 'hard', 'soft', 'arch', 'finance', 'dev']
    fields = budg

    writer.writerow(fields)
    writer.writerows(data)
