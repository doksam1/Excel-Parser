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
source_folder = 'first_round_2024'  # Replace with the path to the source folder
destination_folder = 'Accepted_2024'  # Replace with the path to the destination folder  # Replace with the string to search for in file names
