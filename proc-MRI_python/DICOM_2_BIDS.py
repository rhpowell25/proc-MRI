#%% Import basic packages
import dicom2nifti

import os
import win32com.client 

subject_name = 'Khalil'
STU_num = 'STU00215984'
base_path = 'Z:/MRI/'

#import dicom2nifti.settings as settings
#settings. disable_validate_slice_increment()

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
        if 'NIFTI' in STU_in_path[ii] or 'BIDS' in STU_in_path[ii]:
            continue
        else:
            STU_folder_name.append(STU_in_path[ii])

#%% Convert the DICOM files

# Loop through all the STU folders
for ii in range(len(STU_folder_name)):
    print(STU_folder_name[ii])
    # Define the input & ouput folders
    DICOM_dir = subject_path + '/' + STU_folder_name[ii]
    NIFTI_dir = DICOM_dir + '_NIFTI'
    # Create the output folder
    if not os.path.exists(NIFTI_dir):
        os.mkdir(NIFTI_dir)
    
    # Convert the selected sequences
    for ss in range(len(sequence_names)):
        DICOM_path = os.listdir(DICOM_dir)
        sequences_in_path = list(filter(lambda pp: sequence_names[ss] in pp, DICOM_path))
        for pp in range(len(sequences_in_path)):
            print(sequences_in_path[pp])
            # Define the input & ouput sequence folders
            DICOM_seq_dir = DICOM_dir + '/' + sequences_in_path[pp]
            NIFTI_seq_dir = NIFTI_dir + '/' + sequences_in_path[pp]
            # Create the output sequence folder
            if not os.path.exists(NIFTI_seq_dir):
                os.mkdir(NIFTI_seq_dir)
            # Convert to NIFTI
            dicom2nifti.convert_directory(DICOM_seq_dir, NIFTI_seq_dir)









