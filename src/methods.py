import openpyxl
import os
import shutil
import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin


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
        'B12', 'B27', 'B38', 'B42', 'B43', 'B54', 'B62', 'B70', 'B77', 'B79',
        'B80', 'B96', 'B104'
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

    land_cost = cell_data[0]
    hard_cost = cell_data[1] + cell_data[2] + cell_data[9]
    soft_cost = cell_data[7] + cell_data[8] + cell_data[10] + cell_data[11]
    architectural_cost = cell_data[3] + cell_data[4]
    finance_cost = cell_data[5] + cell_data[6]
    developer_fee = cell_data[12]

    #per value changes needed

    print(f'returning started:{file_path}')
    to_return = {
        "land": land_cost,
        "hard": hard_cost,
        "soft": soft_cost,
        'arch': architectural_cost,
        'finance': finance_cost,
        'dev': developer_fee
        # cell_data[0], cell_data[1] + cell_data[2] * cell_data[3], cell_data[4],
        # cell_data[5]

        # cell_data[0], cell_data[1], cell_data[2],
        # cell_data[3] + ' - ' + cell_data[4],
        # cell_data[5] + ' - ' + cell_data[6],
        # cell_data[7] + ' - ' + cell_data[8],
        # cell_data[9] + ' - ' + cell_data[10],
        # cell_data[11] + ' - ' + cell_data[12],
        # cell_data[13] + ' - ' + cell_data[14]
    }
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
            shutil.move(file_path, os.path.join(destination_folder, filename))
            print(f"Moved: {filename}")


# Function to download xlsx files from a webpage
def download_xlsx_files(url, download_folder="xlsx_downloads"):
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
            download_file(full_url, download_folder)


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
