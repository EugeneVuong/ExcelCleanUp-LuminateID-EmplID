import xlrd
import xlwt
from xlutils.copy import copy
import shutil

import os

# Starting the Workbook/Worksheet
# excel_workbook = xlrd.open_workbook("Alumni donors 2011 - 2021.xls")
# excel_sheet = excel_workbook.sheet_by_index(0)



def find_cellvalue_by_text(text: str):
    for col in range(unclean_sheet.ncols):
        for row in range(unclean_sheet.nrows):
            if text == unclean_sheet.cell_value(col, row):
                return [col, row]

    return None

def is_valid_EmplID(text: str):
    return (len(text.strip()) == 9) and (text.isnumeric())

def is_valid_LuminateOnlineID(text: str):
    return (len(text.strip()) == 7) and (text.isnumeric())

if __name__ == "__main__":

    parent_directory = os.getcwd()

    uncleaned_directory = 'Uncleaned'
    cleaned_directory = 'Cleaned'
    uncleaned_newpath = os.path.join(parent_directory, uncleaned_directory)  
    cleaned_newpath = os.path.join(parent_directory, cleaned_directory)



    
    
    if not os.path.exists(uncleaned_newpath):
        os.makedirs(uncleaned_newpath)
    
    if not os.path.exists(cleaned_newpath):
        os.makedirs(cleaned_newpath)

    
    

    total_counter = 0

    
    
    for xls_files in os.scandir(uncleaned_directory):
        if xls_files.is_file() and xls_files.path.lower().endswith(('.XLS', '.xls')):   
            uncleaned_xls_filename = xls_files.path # File Name
            uncleaned_workbook = xlrd.open_workbook(uncleaned_xls_filename) # Reading It
            unclean_sheet = uncleaned_workbook.sheet_by_index(0)
            clean_workbook = copy(uncleaned_workbook) # Writing on Copied Verison
            clean_sheet = clean_workbook.get_sheet(0)
            cleaned_xls_filename = 'Cleaned - ' + os.path.basename(uncleaned_xls_filename)
            clean_workbook.save(cleaned_xls_filename)



            LuminateHeader = find_cellvalue_by_text('CnAls_1_01_Alias_Type')
            LuminateID = find_cellvalue_by_text("Constituent ID")
            EmplIDHeader = find_cellvalue_by_text("CnAls_1_02_Alias_Type")
            EmplID = find_cellvalue_by_text("Constituent ID_1")


            error_fixed = 0
            for row in range(unclean_sheet.nrows):
                if row != 0:
                    LHeader_Row = unclean_sheet.cell_value(row, LuminateHeader[1])
                    LID_Row = unclean_sheet.cell_value(row, LuminateID[1])
                    EHeader_Row = unclean_sheet.cell_value(row, EmplIDHeader[1])
                    EID_Row = unclean_sheet.cell_value(row, EmplID[1])

                    if(LHeader_Row == 'EmplID'): # If EmplID is on LuminateOnline
                        error_fixed += 1
                        total_counter += 1
                        clean_sheet.write(row, LuminateHeader[1], '')
                        clean_sheet.write(row, LuminateID[1], '')
                        clean_sheet.write(row, EmplIDHeader[1], LHeader_Row)
                        if(EID_Row == ''):
                            print('Changed (Wrong Placement of EmplID [Insert]):', 'Row', row+1)
                            clean_sheet.write(row, EmplID[1], LID_Row)
                        else:
                            print('Changed (Wrong Placement of EmplID [Added]):', 'Row', row+1)
                            More_EmplID = (EID_Row + ', ' + LID_Row)
                            clean_sheet.write(row, EmplID[1], More_EmplID)

                    elif(LID_Row.strip() == EID_Row.strip() and (LID_Row.isnumeric()) and (EID_Row.isnumeric())): # Duplicate
                        error_fixed += 1
                        total_counter += 1
                        if(is_valid_EmplID(EID_Row)):
                            print('Changed (Duplicate EmplID):', 'Row', row+1)
                            clean_sheet.write(row, LuminateHeader[1], '')
                            clean_sheet.write(row, LuminateID[1], '')
                            clean_sheet.write(row, EmplIDHeader[1], EHeader_Row)
                            clean_sheet.write(row, EmplID[1], EID_Row)
                        else:
                            print('Changed (Duplicate LuminateID):', 'Row', row+1)
                            clean_sheet.write(row, LuminateHeader[1], LHeader_Row)
                            clean_sheet.write(row, LuminateID[1], LID_Row)
                            clean_sheet.write(row, EmplIDHeader[1], '')
                            clean_sheet.write(row, EmplID[1], '')
                    
                    elif(EHeader_Row == 'Luminate Online ID'): # If LuminateOnline is on EmplID
                        error_fixed += 1
                        total_counter += 1
                        clean_sheet.write(row, LuminateHeader[1], EHeader_Row)
                        clean_sheet.write(row, EmplID[1], '')
                        clean_sheet.write(row, EmplIDHeader[1], '')
                        if(EID_Row == ''):
                            print('Changed (Wrong Placement of LuminateOnlineID [Insert]):', 'Row', row+1)
                            clean_sheet.write(row, LuminateID[1], EID_Row)
                        else:
                            print('Changed (Wrong Placement of LuminateOnlineID [Added Duplicate]):', 'Row', row+1)
                            More_LID = (EID_Row + ', ' + LID_Row)
                            clean_sheet.write(row, LuminateID[1], More_LID)

                    
            clean_workbook.save(cleaned_xls_filename)
            shutil.move(parent_directory+"\\"+cleaned_xls_filename, parent_directory+"\\Cleaned\\"+cleaned_xls_filename)
            
            print('Error Fixed:', error_fixed)
        
    print('Total Error:', total_counter)
    print(f'Stored in: {parent_directory}\\Cleaned\\')

   
    


    