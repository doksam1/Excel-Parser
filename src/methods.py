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
        excel_name = file_path.split('/')[len(file_path.split('/')) - 1]

        app_number = "CA-" + excel_name.split('.')[0]

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

        Misc_income_24 = [
            'Z813', 'Z814', 'Z815', 'P816', 'Z816', 'Z817', 'Z818'
        ]

        Misc_income_23 = [
            'Z797', 'Z798', 'Z799', 'P800', 'Z800', 'Z801', 'Z802'
        ]

        Misc_income_22 = [
            'Z811', 'Z812', 'Z813', 'P814', 'Z814', 'Z815', 'Z816'
        ]

        operating_expenses_proforma = [
            'E15', 'E16', 'E17', 'E18', 'E19', 'E20', 'A21', 'E21'
        ]

        operating_expenses_proforma_22 = [
            'E14', 'E15', 'E16', 'E17', 'E18', 'E19', 'A20', 'E20'
        ]

        GP1 = [
            'M243', 'AF243', 'M244', 'M245', 'W245', 'AC245', 'M246', 'AH246',
            'M247', 'M248', 'M249'
        ]

        GP2 = [
            'M251', 'AF251', 'M252', 'M253', 'W253', 'AC253', 'M254', 'AH254',
            'M255', 'M256', 'M257'
        ]

        GP3 = [
            'M259', 'AF259', 'M260', 'M261', 'W261', 'AC261', 'M262', 'AH262',
            'M263', 'M264', 'M265'
        ]

        contact_person_during_apps = [
            'L274', 'L275', 'L276', 'V276', 'AB276', 'L277', 'L278', 'L279'
        ]

        developer = ['H287', 'H288', 'H289', 'H290', 'H291', 'H293']

        architect = ['AA287', 'AA288', 'AA289', 'AA290', 'AA291', 'AA293']

        attorneys = ['H295', 'H296', 'H297', 'H298', 'H299', 'H301']

        GC = ['AA295', 'AA296', 'AA297', 'AA298', 'AA299', 'AA301']

        tax_professional = ['H303', 'H304', 'H305', 'H306', 'H307', 'H309']

        energy_consultant = [
            'AA303', 'AA304', 'AA305', 'AA306', 'AA307', 'AA309'
        ]

        consultants = ['H319', 'H320', 'H321', 'H322', 'H323', 'H325']

        investor = ['AA311', 'AA312', 'AA313', 'AA314', 'AA315', 'AA317']

        market_analyst = ['AA319', 'AA320', 'AA321', 'AA322', 'AA323', 'AA325']

        appraiser = ['H327', 'H328', 'H329', 'H330', 'H331', 'H333']

        CNA_consultant = ['AA327', 'AA328', 'AA329', 'AA330', 'AA331', 'AA333']

        Bond_issuer = ['H335', 'H336', 'H337', 'H338', 'H339', 'H341']

        prop_mgmt = ['AA335', 'AA336', 'AAA337', 'AA338', 'AA339', 'AA341']

        subsidy_info = [
            'M954', 'M955', 'J956', 'M957', 'M958', 'M959', 'M960', 'M961',
            'AE954', 'AE955', 'AB956', 'AE957', 'AE958', 'AE959', 'AE960',
            'AE961'
        ]

        pre_existing_subsidies = [
            'M966', 'M967', 'M968', 'M969', 'M970', 'M971', 'M972', 'O973',
            'M974', 'AA966', 'AA967', 'AA968', 'AA969', 'AA970', 'AA973',
            'AA974'
        ]

        service_amenities = ['E27']

        num_beds = [
            'O756', 'J773', 'O773', 'J774', 'O774', 'J775', 'O775', 'J776',
            'O776'
        ]

        hard_cost_contingency = ['B79']

        cells_to_search = hard_cost_contingency

        #load excel file
        workbook = openpyxl.load_workbook(file_path, data_only=True)

        #Select the sheet
        app = "Application"
        budg = 'Sources and Uses Budget'
        proforma = '15 Year Pro Forma'
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

    def get_cell_value(file_path, sheet_name, cell):

        #load excel file
        workbook = openpyxl.load_workbook(file_path)

        #Select the sheet
        sheet = workbook[sheet_name]

        #get the value
        value = sheet[cell].value

        return value
