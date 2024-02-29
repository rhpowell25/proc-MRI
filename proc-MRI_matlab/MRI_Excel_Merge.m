function MRI_Excel_Merge

%% Import the excel spreadsheets of the selected parameters 

% Define where the excel spreadsheets are saved
Base_Path = 'Z:\Lab Members\Henry\4AP MRI\Extracted_MRI\';
Save_Path = 'Z:\Lab Members\Henry\4AP MRI\Merged_MRI\';

%% Find the number of subjects
Excel_Files = struct2table(dir(strcat(Base_Path, '*.xlsx')));

[Subject, Group] = MRI_File_Details;

%% Loop twice for pre and post measures
for pp = 1:2

    %% Define the excel matrix headers
    matrix_headers = {'Subject', 'Group', 'WM left lateral corticospinal tract', 'WM right lateral corticospinal tract', ...
        'WM left ventral corticospinal tract', 'WM right ventral corticospinal tract', ...
        'WM left lateral reticulospinal tract', 'WM right lateral reticulospinal tract', ...
        'WM left ventrolateral reticulospinal tract', 'WM right ventrolateral reticulospinal tract', ...
        'WM left ventral reticulospinal tract', 'WM right ventral reticulospinal tract', ...
        'WM left medial reticulospinal tract', 'WM right medial reticulospinal tract', ...
        'GM left ventral horn', 'GM right ventral horn', 'GM left dorsal horn', 'GM right dorsal horn', ...
        'GM left intermediate zone', 'GM right intermediate zone', 'white matter', 'gray matter'};
    
    % Create the excel matrix
    MRI_Table = array2table(NaN(length(Excel_Files.name), length(matrix_headers)));
    
    MRI_Table.Properties.VariableNames = matrix_headers;
    
    %% Loop through each subject
    Subjects = strings;
    for ii = 1:length(Excel_Files.name)

        Excel_Path = strcat(Base_Path, Excel_Files.name(ii));
        Excel_Table = readtable(char(Excel_Path));
        Subjects{ii,1} = char(extractBefore(Excel_Files.name(ii), '_mt'));
        Subject_idx = strcmp(Subject, Subjects{ii,1});
        try
            MRI_Table.Group(ii,1) = Group(Subject_idx);
        catch
            continue
        end
   
        Tracts = matrix_headers(3:end);
        for tt = 1:length(Tracts)
    
            Tract_idx = find(contains(Excel_Table.Label, Tracts{tt}));
            MRI_Table_idx = find(contains(MRI_Table.Properties.VariableNames, Tracts{tt}));
            if isequal(pp,1)
                MRI_Table{ii, MRI_Table_idx} = Excel_Table.MAP_Pre(Tract_idx);
            else
                MRI_Table{ii, MRI_Table_idx} = Excel_Table.MAP_Post(Tract_idx);
            end
        end
    end

    MRI_Table.Subject = Subjects;

    %% Save to Excel

    % Define the file name
    if isequal(pp,1)
        filename = 'Merged_Excel_Pre';
    else
        filename = 'Merged_Excel_Post';
    end
    

    % Save the file
    if ~exist(Save_Path, 'dir')
        mkdir(Save_Path);
    end
    writetable(MRI_Table, strcat(Save_Path, filename, '.xlsx'))

end

