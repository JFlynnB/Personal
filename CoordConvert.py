# -*- coding: utf-8 -*-
import xlrd
import time
import csv
import os


def find_folder_in_root(folder_name):
    # Get a list of all drive letters on the system
    drives = ['{}:\\'.format(d) for d in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ' if os.path.exists('{}:'.format(d))]

    # Iterate over each drive and check for the folder
    for drive in drives:
        folder_path = os.path.join(drive, folder_name)
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            return folder_path  # Return the path to the folder if found

    return None  # Return None if the folder is not found in any root directory


# Example usage:
# folder_im_looking_for = "Part_Coordinates"
# folder_path = find_folder_in_root(folder_im_looking_for)
# if folder_path:
#    print("Folder '{}' found at: {}".format(folder_im_looking_for, folder_path))
# else:
#    print("Folder '{}' not found in any root directory.".format(folder_im_looking_for))


def read_columns_from_excel(file_path):
    xnums = []
    ynums = []
    znums = []

    # Load the Excel file
    wb = xlrd.open_workbook(file_path)

    # Select the first sheet
    sheet = wb.sheet_by_index(0)

    # Iterate over rows starting from row 2 (skipping header)
    for row_index in range(sheet.nrows):
        # Append the value from column B to the xnum array
        if isinstance(sheet.cell(row_index, 1).value, str):
            continue
        else:
            xnums.append(sheet.cell(row_index, 1).value)
        # Append the value from column C to the ynum array
        if isinstance(sheet.cell(row_index, 2).value, str):
            continue
        else:
            ynums.append(sheet.cell(row_index, 2).value)
            # Append the value from column D to the znum array
        if isinstance(sheet.cell(row_index, 2).value, str):
            continue
        else:
            znums.append(sheet.cell(row_index, 3).value)

    return xnums, ynums, znums


def write_to_csv(folderPath, fileName, xNums, yNums, zNums):
    base_name, old_extension = os.path.splitext(fileName)
    base_name_a = base_name[len(folderPath):]
    csv_file_name = "corrected_" + base_name_a[1:] + ".csv"
    new_path = folderPath + "\\" + csv_file_name
    #print(xNums) #debug
    with open(new_path, 'w', newline='') as file:
        writer = csv.writer(file)
        for i in range(300):
            if i <= 99:
                if i < len(xNums):
                    data = [xNums[i]]
                    writer.writerow(data)
                else:
                    writer.writerow('')
            if 100 <= i <= 199:
                if i - 100 < len(yNums):
                    data = [yNums[i-100]]
                    writer.writerow(data)
                else:
                    writer.writerow('')
            if 200 <= i <= 299:
                if i - 200 < len(yNums):
                    data = [zNums[i-200]]
                    writer.writerow(data)
                else:
                    writer.writerow('')

    return csv_file_name


def get_files_sorted_by_creation_time(folder_path):
    files = []
    # Iterate through all files in the folder
    try:
        for file_name in os.listdir(folder_path):
            file_path = os.path.join(folder_path, file_name)
            # Check if the path is a file
            if os.path.isfile(file_path):
                # Get file creation time
                creation_time = os.path.getctime(file_path)
                # Append file path and creation time to the list
                files.append((file_path, creation_time))
    except TypeError:
        return None

    # Sort files by creation time in descending order
    sorted_files = sorted(files, key=lambda x: x[1], reverse=True)

    # Extract file paths from sorted list
    sorted_file_paths = [file[0] for file in sorted_files]

    return sorted_file_paths


def get_file_to_correct(data_files):
    for file in data_files:
        if file is not None:
            if file[0:9] == "corrected_":
                continue
            else:
                name, extension = os.path.splitext(file)
                if extension == ".xls":
                    already_corrected = False
                    for file_two in data_files:
                        name_two, ext_two = os.path.splitext(file_two)
                        if "corrected_" + file == name_two:
                            already_corrected = True
                            break
                    if already_corrected:
                        continue
                    else:
                        return file
                else:
                    continue
        else:
            return None
    return None

def search_and_convert():
    folder_im_looking_for = "Part_Coordinates"
    found_folder = find_folder_in_root(folder_im_looking_for)
    sorted_data_files = get_files_sorted_by_creation_time(found_folder)
    if not sorted_data_files:
        return None
    else:
        file_string = get_file_to_correct(sorted_data_files)
        if file_string is None:
            return None
        else:
            #full_path = found_folder + file_string
            resultx, resulty, resultz = read_columns_from_excel(file_string)
            try:
                return write_to_csv(found_folder, file_string, resultx, resulty, resultz)
            except:
                return 1


def main():
    while True:
        search_and_convert()
        time.sleep(15)


if __name__ == "__main__":
    main()
# Example usage:
# folder_path = "/path/to/your/folder"
# sorted_data_files = get_files_sorted_by_creation_time(found_folder)
# print(sorted_data_files)

# Example usage:
# folder_im_looking_for = "Part_Coordinates"
# found_file = "\\example.xls"
# found_folder = find_folder_in_root(folder_im_looking_for)

# sorted_data_files = get_files_sorted_by_creation_time(found_folder)
# print(sorted_data_files)
# file_string = get_file_to_correct(sorted_data_files)
# print(file_string is not None)

# resultX, resultY, resultZ = read_columns_from_excel()
# print(resultX)
# print(resultY)

# write_to_csv(found_folder, found_file, resultX, resultY, resultZ)

print("done")



