import re
import shutil
import os
from openpyxl import load_workbook

#setup workbook
wb = load_workbook(filename= 'summary_of_participants.xlsx')
ws = wb.active


# function takes name of participant, and topic number/index
# moves to desired location
def cert_finder(name_of_participant,topic_number):

    # directory of source of files and desired destination
    main_dir = 'C:\\Users\\justi\\Documents\\certs'
    folder_dest = 'C:\\Users\\justi\\Documents\\certs_landing'
    
        
    dir_of_cert = main_dir + f'\\Topic {str(topic_number)}' 
    # sets directory of topic
    
    certs = os.listdir(dir_of_cert)
    # lists the certs inside the topic dir

    for certificate in certs:
        # loops through all certificates inside respective topic

        new_name_certificate = certificate[:-4]
        # print(certificate[:-4])

        if new_name_certificate == name_of_participant:
            # if name of cert matches name of participant
            # copy the file

            source = os.path.join(dir_of_cert,new_name_certificate + '.png')
            dest = os.path.join(folder_dest,new_name_certificate)
            shutil.copy(source,dest)
            file_name = os.path.join(dest,new_name_certificate)
            
            # rename so that to avoid replacing 
            os.rename(file_name+'.png',file_name+f'{topic_number}.png')

            print(source)
            print(dest)

            print('sucessfully copied!')            
        
 
# function gets array with name and all bool values from excel file

def get_row(row_number):
    row = ws[row_number]
    array = []
    for cell in row:
        array.append(cell.value)
        
    return array

# SCRIPT

max_rows = 27

for row_number in range(1,max_rows):
    # array of names and bool
    array = get_row(row_number)

    # first element is name
    name_of_participant = array[0]

    for i in range(1,9):
        cert_finder(name_of_participant,i)
        
        # for all 8 topics, search for the participants' name
        # copy if it matches
        


