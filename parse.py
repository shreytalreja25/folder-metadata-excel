import os
from datetime import datetime
from openpyxl import Workbook

def get_file_info(path):
    file_info = []
    for entry in os.listdir(path):
        entry_path = os.path.join(path, entry)
        if os.path.isfile(entry_path):
            stat_info = os.stat(entry_path)
            file_info.append({
                'Name': entry,
                'Date Last Modified': datetime.fromtimestamp(stat_info.st_mtime),
                'Size (Bytes)': stat_info.st_size,
                'Type': 'File'
            })
        elif os.path.isdir(entry_path):
            stat_info = os.stat(entry_path)
            file_info.append({
                'Name': entry,
                'Date Last Modified': datetime.fromtimestamp(stat_info.st_mtime),
                'Size (Bytes)': stat_info.st_size,
                'Type': 'Folder'
            })
    return file_info

def main():
    print("Running...")
    folder_path = input("Enter the path of the folder: ")
    file_info = get_file_info(folder_path)

    # Create an Excel workbook and sheet
    wb = Workbook()
    ws = wb.active
    ws.append(['Name', 'Date Last Modified', 'Size (Bytes)', 'Type'])

    for info in file_info:
        ws.append([info['Name'], info['Date Last Modified'], info['Size (Bytes)'], info['Type']])

    excel_file_path = os.path.join(os.path.dirname(__file__), 'folder_info.xlsx')
    wb.save(excel_file_path)
    print(f"Excel file saved at: {excel_file_path}")

if __name__ == "__main__":
    main()
