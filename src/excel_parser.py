import os
import openpyxl
import csv

#directory for excel files
directory = "C:/Users/cameronshaw/Documents/Affordable Research/All Applications Since 2022/Accepted_2024_construction_only"
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

    app_number = "CA-" + file_path.split('/')[7].split('.')[0]

    application_data = [
        'AG437', 'AG442', 'O773', 'T773', 'AG993', 'AG449', 'AD411', 'AE424',
        'AA287', 'Q494', 'S413'
    ]

    address = ['I185']

    budget_and_sources = [
        'B12', 'B26', 'B38', 'B42', 'B43', 'B54', 'B62', 'B70', 'B77', 'B79',
        'B80', 'B96', 'B104', 'B13', 'B14', 'B27'
    ]

    construction_financing = [['C554', 'M554', 'Q554', 'W554', 'AM554'],
                              ['C555', 'M555', 'Q555', 'W555', 'AM555'],
                              ['C556', 'M556', 'Q556', 'W556', 'AM556'],
                              ['C557', 'M557', 'Q557', 'W557', 'AM557'],
                              ['C558', 'M558', 'Q558', 'W558', 'AM558'],
                              ['C559', 'M559', 'Q559', 'W559', 'AM559'],
                              ['C560', 'M560', 'Q560', 'W560', 'AM560'],
                              ['C561', 'M561', 'Q561', 'W561', 'AM561'],
                              ['C562', 'M562', 'Q562', 'W562', 'AM562'],
                              ['C563', 'M563', 'Q563', 'W563', 'AM563'],
                              ['C564', 'M564', 'Q564', 'W564', 'AM564'],
                              ['C565', 'M565', 'Q565', 'W565', 'AM565']]

    permanent_financing = [['C627', 'M627', 'Q627', 'W627', 'AO627'],
                           ['C628', 'M628', 'Q628', 'W628', 'AO628'],
                           ['C629', 'M629', 'Q629', 'W629', 'AO629'],
                           ['C630', 'M630', 'Q630', 'W630', 'AO630'],
                           ['C631', 'M631', 'Q631', 'W631', 'AO631'],
                           ['C632', 'M632', 'Q632', 'W632', 'AO632'],
                           ['C633', 'M633', 'Q633', 'W633', 'AO633'],
                           ['C634', 'M634', 'Q634', 'W634', 'AO634'],
                           ['C635', 'M635', 'Q635', 'W635', 'AO635'],
                           ['C636', 'M636', 'Q636', 'W636', 'AO636'],
                           ['C637', 'M637', 'Q637', 'W637', 'AO637'],
                           ['C638', 'M638', 'Q638', 'W638', 'AO638']]

    parking_spaces = ['Q494', 'M495', 'AH495']

    GC_Fees = [
        'B19', 'B20', 'B21', 'B23', 'B26', 'B31', 'B32', 'B33', 'B35', 'B38'
    ]

    cells_to_search = GC_Fees

    #load excel file
    workbook = openpyxl.load_workbook(file_path, data_only=True)

    #Select the sheet
    app = "Application"
    budg = 'Sources and Uses Budget'
    sheet = workbook[budg]

    #get the values
    cell_data = [sheet[i].value for i in cells_to_search]

    #for getting financial data
    # cell_data = []
    # for row in cells_to_search:
    #     lender = sheet[row[0]].value
    #     source = sheet[row[1]].value
    #     term = sheet[row[2]].value
    #     interest = sheet[row[3]].value
    #     funds = sheet[row[4]].value
    #     cell_data.append([app_number, lender, source, term, interest, funds])

    #removes null values
    # for j in range(len(cell_data)):
    #     if cell_data[j] is None:
    #         cell_data[j] = 'None'

    # for j in range(len(cell_data) - 1, -1, -1):
    #     if cell_data[j][5] is None:
    #         cell_data.pop(j)

    #per value changes needed
    # if isinstance(cell_data[1], str) or isinstance(
    #         cell_data[2], str) or isinstance(cell_data[3], str):
    #     cell_data[1] = 0
    #     cell_data[2] = 0
    #     cell_data[3] = 0
    #     print(f'skipped numbers for this:{file_path}')

    # land_cost = cell_data[0] + cell_data[13] + cell_data[14]
    # hard_cost = cell_data[1] + cell_data[2] + cell_data[9] + cell_data[15]
    # soft_cost = cell_data[7] + cell_data[8] + cell_data[10] + cell_data[11]
    # architectural_cost = cell_data[3] + cell_data[4]
    # finance_cost = cell_data[5] + cell_data[6]
    # developer_fee = cell_data[12]

    #per value changes needed

    print(f'returning started:{file_path}')
    to_return = [
        app_number,
        #     # land_cost, hard_cost, soft_cost, architectural_cost, finance_cost,
        #     # developer_fee
        #     # cell_data[0],
        #     # cell_data[1] + cell_data[2] * cell_data[3],
        #     # cell_data[4],
        #     # cell_data[5],
        #     # cell_data[6],
        #     # cell_data[7],
        #     # cell_data[8],
        #     # cell_data[9],
        #     # cell_data[10]

        #   various budget data
        #     # cell_data[0], cell_data[1], cell_data[2],
        #     # cell_data[3] + ' - ' + cell_data[4],
        #     # cell_data[5] + ' - ' + cell_data[6],
        #     # cell_data[7] + ' - ' + cell_data[8],
        #     # cell_data[9] + ' - ' + cell_data[10],
        #     # cell_data[11] + ' - ' + cell_data[12],
        #     # cell_data[13] + ' - ' + cell_data[14]
        #     cell_data[0]

        # parking data
        # cell_data[0],
        # cell_data[1],
        # cell_data[2]

        #single data point
        cell_data[0],
        cell_data[1],
        cell_data[2],
        cell_data[3],
        cell_data[4],
        cell_data[5],
        cell_data[6],
        cell_data[7],
        cell_data[8],
        cell_data[9]
    ]

    print(f'data returned:{file_path}')
    return to_return


data = [get_CTCHC_data(i) for i in excel_paths]

#writes excel spreadsheet data into a csv file
with open('stories.csv', 'w', newline='') as file:
    writer = csv.writer(file)
    app = [
        'Units', 'NRSF', 'Prevailing Wage', 'Parking SF', 'Stories',
        "Buildings", 'Architect', 'Parking Spaces', 'Subterranean?'
    ]
    budg = ['land', 'hard', 'soft', 'arch', 'finance', 'dev']

    GC = [
        'Requirements - Rehab', 'Overhead - Rehab', 'Profit - Rehab',
        'Insurance - Rehab', 'Constr Costs - Rehab', 'Requirements - New',
        'Overhead - New', 'Profit - New', 'Insurance - New',
        'Constr Costs - New'
    ]

    fields = GC

    writer.writerow(fields)
    # for writing a list of lists
    # for project in data:
    #     writer.writerows(project)
    # for writing a single table
    writer.writerows(data)
