#%% Import basic packages
import nibabel as nib
import numpy as np
import os
import win32com.client 
import matplotlib.pyplot as plt

subject_name = 'Naser'
STU_num = 'STU00215984'
base_path = 'Z:/MRI/'
sequence_name = 'T1'

#%% Indexing to the MRI NIFTI files
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
        if 'NIFTI' in STU_in_path[ii]:
            STU_folder_name.append(STU_in_path[ii])
        else:
            continue

#%% Read the NIFTI files

# Loop through all the STU folders
for ii in range(len(STU_folder_name)):
    print(STU_folder_name[ii])
    # Define the input folder
    NIFTI_dir = subject_path + '/' + STU_folder_name[ii]
    NIFTI_path = os.listdir(NIFTI_dir)
    sequences_in_path = list(filter(lambda pp: sequence_name in pp, NIFTI_path))
    
    # Loop through each sequence folder
    for pp in range(len(sequences_in_path)):
        print(sequences_in_path[pp])
        # Redefine the input sequence folder
        NIFTI_seq_dir = NIFTI_dir + '/' + sequences_in_path[pp]
        # Find the NIFTI files in the folder
        NIFTI_files = os.listdir(NIFTI_seq_dir)

        # Read the NIFTI
        t1_img = nib.load(NIFTI_seq_dir + '/' + NIFTI_files[0])
        
        t1_hdr = t1_img.header
        print(t1_hdr)
        t1_hdr.keys()
        
        t1_data = t1_img.get_fdata()
        t1_data
        type(t1_data)
        t1_data.dtype
        
        print(np.min(t1_data))
        print(np.max(t1_data))
        t1_data[9, 19, 2]
        x_slice = t1_data[9, :, :]
        y_slice = t1_data[:, 19, :]
        z_slice = t1_data[:, :, 2]
        
        slices = [x_slice, y_slice, z_slice]
        
        fig, axes = plt.subplots(1, len(slices))
        for i, slice in enumerate(slices):
            axes[i].imshow(slice.T, cmap="gray", origin="lower")
            
        t1_affine = t1_img.affine
        t1_affine






