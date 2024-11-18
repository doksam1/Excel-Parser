import os
import shutil


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


# Example usage
source_folder = 'C:/Users/cameronshaw/Documents/Affordable Research/All Applications Since 2022/first_round_2024'  # Replace with the path to the source folder
destination_folder = 'C:/Users/cameronshaw/Documents/Affordable Research/All Applications Since 2022/Accepted_2024_2'  # Replace with the path to the destination folder  # Replace with the string to search for in file names
Accepted_new_construction_only = [
    '24-414', '24-424', '24-426', '24-427', '24-428', '24-433', '24-434',
    '24-435', '24-441', '24-443', '24-455', '24-459', '24-460', '24-467',
    '24-469', '24-471', '24-472', '24-473', '24-474', '24-476', '24-477',
    '24-478', '24-481', '24-482', '24-483', '24-485', '24-489', '24-490',
    '24-492', '24-493', '24-494', '24-497', '24-500', '24-502', '24-503',
    '24-504', '24-509', '24-511', '24-515', '24-516', '24-521', '24-522',
    '24-525', '24-527', '24-528', '24-535', '24-539', '24-541', '24-545',
    '24-552', '24-553', '24-554', '24-564'
]

for i in Accepted_new_construction_only:
    find_and_move_files_by_name(source_folder, destination_folder, i)
