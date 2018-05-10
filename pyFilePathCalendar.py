'''
    Title:  Recursive Length of File 0.7.1
    Author:  Joe Friedrich
    License:  MIT
'''

import openpyxl
import os
import datetime

def get_excel_output_location():
	'''
		Write!
	'''
	print('\nThis will save a file in .xlsx format to your desktop.')
	print("It's name will be the current time.xlsx.")
	user_folder = os.environ['USERPROFILE']
	time = datetime.datetime.now()
	time = time.strftime("%Y-%m-%d.%H.%M.%S")
	file_location = user_folder + "\desktop\py" + time + ".xlsx"
	return file_location
	
def get_starting_directory():
	'''
		Write!
	'''
    print('\nPlease enter the top level of your path.')
    starting_directory = input() #get starting directory from user
    return starting_directory
	
def load_directory_tree(start_walking):
	'''
		Write!
	'''
    directory_tree = os.walk(start_walking)
    return directory_tree

#------------------Begin Program------------------------
	
print("\nWelcome to the Recursive 'Length of file-path' program.")

file_location = get_excel_output_location()

excel_file = openpyxl.Workbook() #create a new excel workbook

starting_directory = get_starting_directory()
directory_tree = load_directory_tree(starting_directory)

worksheet = excel_file.active #use the first sheet of the workbook
worksheet.append(['directory', 'file', 'length', starting_directory]) #sets title line

for directory in directory_tree: #traverses directories
    for file in directory[2]:  #for all files in the directory
        length_of_path = len(directory[0]) + 1 + len(file) - len(starting_directory) #calculate length
        worksheet.append([directory[0],file,length_of_path]) #write the directory/file/length
        
excel_file.save(file_location)

print('\nExcel file saved: ' + file_location)
print('\nThank you and have a nice day.\n')