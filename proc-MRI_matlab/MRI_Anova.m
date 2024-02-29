function MRI_Anova

%% Import the excel spreadsheets of the selected parameters 

% Define where the excel spreadsheets are saved
Base_Path = 'Z:\Lab Members\Henry\4AP MRI\Extracted_MRI\';

Label = 'corticospinal';

%% Find the number of subjects
Excel_Files = struct2table(dir(strcat(Base_Path, '*.xlsx')));

[Subjects, Groups] = MRI_File_Details;

%% Loop through each subject

for ii = 1:length(Excel_Files.name)

    Excel_Path = strcat(Base_Path, Excel_Files.name(ii));
    Excel_Table = readtable(char(Excel_Path));

    Subject = char(extractBefore(Excel_Files.name(ii), '_mt'));
    Subject_idx = strcmp(Subjects, Subject);
    Group_num = Groups(Subject_idx);

    Label_idx = find(contains(Excel_Table.Label, Label));

    % Assign the temporary variables
    temp_MAP = cat(1, Excel_Table.MAP_Pre(Label_idx), Excel_Table.MAP_Post(Label_idx));
    temp_Time = strings;
    temp_Time(1:length(Excel_Table.MAP_Pre(Label_idx)), 1) = 'Pre';
    temp_Time(length(Excel_Table.MAP_Pre(Label_idx)) + 1:length(temp_MAP), 1) = 'Post';
    temp_Label = strings;
    temp_Label(1:length(Excel_Table.MAP_Pre(Label_idx)), 1) = Excel_Table.Label(Label_idx);
    temp_Label(length(Excel_Table.MAP_Pre(Label_idx)) + 1:length(temp_MAP), 1) = Excel_Table.Label(Label_idx);
    temp_Group = ones(length(temp_MAP), 1)*Group_num;

    if isequal(ii,1)
        MAP = temp_MAP;
        Time = temp_Time;
        Label = temp_Label;
        Group = temp_Group;
    else
        MAP = cat(1, MAP, temp_MAP);
        Time = cat(1, Time, temp_Time);
        Label = cat(1, Label, temp_Label);
        Group = cat(1, Group, temp_Group);
    end
end

%% Do the statistics


[p, ~, stats] = anovan(MAP, {Time, Group, Label}, ...
    'varnames', {'Pre_vs_Post', 'Group', 'Tract'});

figure
hold on
[results, ~, ~, gnames] = multcompare(stats, "Dimension", [1 ]);




