import os
import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin


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


# Example usage
webpage_urls = [
    'https://www.treasurer.ca.gov/ctcac/2023/firstround/applications/index.asp',
    'https://www.treasurer.ca.gov/ctcac/2023/secondround/applications/index.asp',
    'https://www.treasurer.ca.gov/ctcac/2023/thirdround/applications/index.asp',
    'https://www.treasurer.ca.gov/ctcac/2022/firstround/applications/index.asp',
    'https://www.treasurer.ca.gov/ctcac/2022/secondround/applications/index.asp'
]  # Replace with the target webpage URL

for link in webpage_urls:
    file_name = link.split('/')[5] + '_' + link.split('/')[4]
    download_xlsx_files(link, file_name)
