#XCDE Map Build Process
#This process parses an excel file and produces a json map for consumption by XCDE

acceptable_filenames = [
    "PDXC_CIs_Attr_Rel_Mapping.xlsx",
    "PDXC_ITAM_CIs_Attr_Rel_Mapping.xlsx",
    "CIs_Attr_Rel_Mapping_ES_UCMDB_ESL.xls"
]
file_directory = 'processqueue'
print("File to processed must be in the "+file_directory+" directory.\nSelect one of the below file names for processing: ")
file_number = 1
for this_file in acceptable_filenames:
    print(str(file_number) + ") " + this_file)
    file_number += 1
print(str(file_number) + ") Other")
print(str(99) + ") Quit")

file_location = ''
file_name = ''
while file_name == '':
    selected_number = 0
    try:
        selected_number = int(input("Enter numeric selection: "))
        if selected_number == 99:
            print('Exiting program')
            quit(200)
        if selected_number < 1 or selected_number > file_number:
            print("Invalid input. Must be between 1 and " + str(file_number))
        elif selected_number == file_number:
            file_name = input("Enter file name: ")
        else:
            file_name = acceptable_filenames[selected_number-1]
    except ValueError:
        print("Invalid input. Must be an integer.")

    if file_name != '':
        file_location = file_directory+"/"+file_name
        print("Opening "+file_location)
        try:
            workbook = xlrd.open_workbook(file_location)
        except FileNotFoundError:
            print("File not found.  Ensure "+file_name+" is in the "+file_directory+" directory.")
            file_location = ''
            file_name = ''
