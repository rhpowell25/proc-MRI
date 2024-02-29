function MRI_Excel_Extract(Sampling_Params, Save_Excel)

%% Import the excel spreadsheets of the selected parameters 

% Define where the excel spreadsheets are saved
Base_Path = 'Z:\Lab Members\Henry\4AP MRI\SRAlab_SCT_outputs\';
Save_Path = 'Z:\Lab Members\Henry\4AP MRI\Extracted_MRI\';

%% Find the number of subjects
Base_Directory = struct2table(dir(Base_Path));

Pre_Subjects = sort(Base_Directory.name(contains(Base_Directory.name, 'A')));
Post_Subjects = sort(Base_Directory.name(contains(Base_Directory.name, 'B')));

Matching_Subjects = intersect(erase(Pre_Subjects, 'A'), erase(Post_Subjects, 'B'));

%% Loop through each subject

for ii = 1:length(Matching_Subjects)
    Pre_CSV_Path = strcat(Base_Path, Matching_Subjects{ii}, 'A\', Sampling_Params.Scan, '\');
    Post_CSV_Path = strcat(Base_Path, Matching_Subjects{ii}, 'B\', Sampling_Params.Scan, '\');

    % Find the corresponding Pre-.csv files
    Pre_CSV_Files = struct2table(dir(strcat(Pre_CSV_Path, '*.csv')));
    if isempty(Pre_CSV_Files)
        disp('No Pre-Table Found!')
        continue
    end
    Pre_CSV_idx = contains(Pre_CSV_Files.name, Sampling_Params.CSV_Tag);
    if ~any(Pre_CSV_idx)
        disp('No Pre-Table Found!')
        continue
    end
    Pre_CSV_path = strcat(Pre_CSV_Path, Pre_CSV_Files.name(Pre_CSV_idx));
    Pre_CSV_table = readtable(char(Pre_CSV_path));

    % Find the corresponding Post-.csv files
    Post_CSV_Files = struct2table(dir(strcat(Post_CSV_Path, '*.csv')));
    if isempty(Post_CSV_Files)
        disp('No Post-Table Found!')
        continue
    end
    Post_CSV_idx = contains(Post_CSV_Files.name, Sampling_Params.CSV_Tag);
    if ~any(Post_CSV_idx)
        disp('No Post-Table Found!')
        continue
    end
    Post_CSV_path = strcat(Post_CSV_Path, Post_CSV_Files.name(Post_CSV_idx));
    Post_CSV_table = readtable(char(Post_CSV_path));

    % Generate the combined table
    matrix_headers = {'Label', 'MAP_Pre', 'STD_Pre', 'MAP_Post', 'STD_Post'};
    matrix_length = height(Pre_CSV_table);
    CSV_table = array2table(zeros(matrix_length, length(matrix_headers)));
    CSV_table.Properties.VariableNames = matrix_headers;

    % Assign the values to the new CSV
    CSV_table.Label = Pre_CSV_table.Label;
    CSV_table.MAP_Pre = Pre_CSV_table.MAP__;
    CSV_table.STD_Pre = Pre_CSV_table.STD__;
    for tt = 1:height(CSV_table)
        Label_idx = contains(Post_CSV_table.Label, CSV_table.Label{tt});
        CSV_table.MAP_Post(tt) = Post_CSV_table.MAP__(Label_idx);
        CSV_table.STD_Post(tt) = Post_CSV_table.STD__(Label_idx);
    end

    %% Save to Excel

    if isequal(Save_Excel, 1)

        % Define the file name
        filename = char(strcat(Matching_Subjects{ii}, '_', Sampling_Params.Scan, '_', ...
            Sampling_Params.CSV_Tag));

        % Save the file
        if ~exist(Save_Path, 'dir')
            mkdir(Save_Path);
        end
        writetable(CSV_table, strcat(Save_Path, filename, '.xlsx'))

    end

end

disp('Done')