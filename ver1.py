import zipfile
import os
from tabulate import tabulate

def count_file_types(zip_ref, folder_path):
    file_types_count = {}

    for file_info in zip_ref.infolist():
        # Extract file path from the zip folder
        file_path = file_info.filename

        # If the file is within the specified folder
        if file_path.startswith(folder_path):
            # Determine file type (folder, xls, file, etc.)
            file_type = 'folder' if file_info.is_dir() else file_path.split('.')[-1].lower()

            # Update the count in the dictionary
            file_types_count[file_type] = file_types_count.get(file_type, 0) + 1

    return file_types_count

def main():
    zip_path = 'C:/Users/aiden/Downloads/WindowsProject(3912).zip'

    if os.path.exists(zip_path):
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            descendants_count = count_file_types(zip_ref, '')

        # Convert the dictionary to a list of lists
        data = [[file_type, count] for file_type, count in descendants_count.items()]

        # Print the table using tabulate
        table = tabulate(data, headers=["File Type", "Count"], tablefmt="pretty")
        print(table)
    else:
        print(f"Error: The specified path '{zip_path}' does not exist.")

if __name__ == "__main__":
    main()
