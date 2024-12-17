import openpyxl
import os
import shutil
import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin


class data_retrieval:

    def __init__(self):
        pass

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
            'AG437', 'AG442', 'O773', 'T773', 'AG993', 'AG449', 'AD411',
            'AE424', 'AA287', 'Q494', 'S413'
        ]

        expenses_per_unit = ['AG437', 'AC883']

        address = ['I185']

        budget_and_sources = [
            'B12', 'B26', 'B38', 'B42', 'B43', 'B54', 'B62', 'B70', 'B77',
            'B79', 'B80', 'B96', 'B104', 'B13', 'B14', 'B27'
        ]

        CPA = ['H311', 'H312', 'H313', 'H314', 'H315', 'H317']
        # because the ones from 2023 are slightly different
        CPA2 = ['H310', 'H311', 'H312', 'H313', 'H314', 'H316']

        expenses = ['AC885']

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
            'B17', 'B18', 'B19', 'B20', 'B21', 'B22', 'B23', 'B24', 'B26',
            'B29', 'B30', 'B31', 'B32', 'B33', 'B34', 'B35', 'B36', 'B38'
        ]

        Miscellaneous_income = [
            'Z813', 'Z814', 'Z815', 'P816', 'Z816', 'Z817', 'Z818'
        ]

        cells_to_search = Miscellaneous_income

        #load excel file
        workbook = openpyxl.load_workbook(file_path, data_only=True)

        #Select the sheet
        app = "Application"
        budg = 'Sources and Uses Budget'
        sheet = workbook[app]

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
        # for finding rehab or new construction
        # if cell_data[1] != 0:
        #     to_return = [
        #         app_number,
        #         cell_data[0],
        #         cell_data[1],
        #         cell_data[2],
        #         cell_data[3],
        #         cell_data[4],
        #         cell_data[5],
        #         cell_data[6],
        #         cell_data[7],
        #         cell_data[8],
        #     ]

        # else:
        #     to_return = [
        #         app_number, cell_data[9], cell_data[10], cell_data[11],
        #         cell_data[12], cell_data[13], cell_data[14], cell_data[15],
        #         cell_data[16], cell_data[17]
        #     ]
        # default
        to_return = [app_number] + cell_data
        # land_cost, hard_cost, soft_cost, architectural_cost, finance_cost,
        # developer_fee
        # cell_data[0],
        # cell_data[1] + cell_data[2] * cell_data[3],
        # cell_data[4],
        # cell_data[5],
        # cell_data[6],
        # cell_data[7],
        # cell_data[8],
        # cell_data[9],
        # cell_data[10]

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

        # CPA
        # cell_data[0],
        # cell_data[1],
        # cell_data[2],
        # cell_data[3],
        # cell_data[4],
        # cell_data[5]

        #expenses_per_unit
        # cell_data[0],
        # cell_data[1]

        # misc income

        print(f'data returned:{file_path}')
        return to_return

    def find_and_move_files_by_name(source_folder, destination_folder,
                                    search_string):
        # Create the destination folder if it doesn't exist
        if not os.path.exists(destination_folder):
            os.makedirs(destination_folder)

        # Loop through each file in the source folder
        for filename in os.listdir(source_folder):
            file_path = os.path.join(source_folder, filename)

            # Check if the filename contains the search string
            if search_string in filename:
                shutil.move(file_path,
                            os.path.join(destination_folder, filename))
                print(f"Moved: {filename}")

    def download_file(file_url, download_folder):
        # Get the filename from the URL
        filename = os.path.join(download_folder, file_url.split('/')[-1])

        # Download the file content
        response = requests.get(file_url, stream=True)
        response.raise_for_status()  # Check for request errors

        # Save the file to the download folder
        with open(filename, 'wb') as file:
            for chunk in response.iter_content(chunk_size=8192):
                file.write(chunk)

        print(f"Downloaded: {filename}")

    # Function to download xlsx files from a webpage
    def download_xlsx_files(self, url, download_folder="xlsx_downloads"):
        # Create the download folder if it doesn't exist
        if not os.path.exists(download_folder):
            os.makedirs(download_folder)

        # Send a request to the webpage
        response = requests.get(url)
        response.raise_for_status()  # Check for request errors

        # Parse the webpage content
        soup = BeautifulSoup(response.text, 'html.parser')

        # Find all <a> tags with .xlsx links
        for link in soup.find_all('a', href=True):
            file_url = link['href']
            if file_url.endswith('.xlsx'):
                # Create the full URL if it's a relative link
                full_url = urljoin(url, file_url)

                # Download the .xlsx file
                self.download_file(full_url, download_folder)
