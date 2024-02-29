
%% Load the MRI Files
clc
clear

Sampling_Params = struct( ...
    'Scan', 'mt', ... % Scan Modality ('dmri', 'mt')
    'CSV_Tag', 'mtr'); % Tag on the CSV file ('mtr', 'mtstat')

Save_Excel = 1;

MRI_Excel_Extract(Sampling_Params, Save_Excel);

MRI_Excel_Merge




