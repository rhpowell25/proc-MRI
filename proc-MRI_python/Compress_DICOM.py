#%% Import basic packages

import os
import win32com.client
import shutil

subject_name = 'Naser'
STU_num = 'STU00215984'
base_path = 'Z:/MRI/'

#%% Indexing to the MRI DICOM files
# Files & folders
dir_path = os.listdir(base_path)
# Folders containg subject name
folders_in_path = list(filter(lambda ii: subject_name in ii, dir_path))

# Selected folder
selec_folder = folders_in_path[0]
subject_shortcut = base_path + selec_folder

# Convert the shortcut into the actual path
shell = win32com.client.Dispatch("WScript.Shell")
shortcut = shell.CreateShortCut(subject_shortcut)
subject_path = shortcut.Targetpath

# Find the MRI
subject_dir = os.listdir(subject_path)
# Files and folders containing the STU number
STU_in_path = list(filter(lambda ii: STU_num in ii, subject_dir))
# Find folders only
STU_folder_name = []
for ii in range(len(STU_in_path)):
    if os.path.isdir(subject_path + '/' + STU_in_path[ii]):
        if 'NIFTI' in STU_in_path[ii] or 'Zip' in STU_in_path[ii]:
            continue
        else:
            STU_folder_name.append(STU_in_path[ii])

#%% Zip the DICOM files

# Loop through all the STU folders
for ii in range(len(STU_folder_name)):
    print(STU_folder_name[ii])
    
    # Define the input & ouput folders
    DICOM_dir = subject_path + '/' + STU_folder_name[ii]
    
    shutil.make_archive(DICOM_dir, 'zip', DICOM_dir)
        
        
        
        

