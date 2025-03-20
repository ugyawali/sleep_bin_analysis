function analyzeSleepDataByTimeBins()
    % Main function to analyze sleep data across multiple conditions with specific time bins
    %
    % This function allows the user to:
    % 1. Select a time bin resolution (1, 3, or 6 hours)
    % 2. Select baseline and extinction day folders
    % 3. Process sleep data within chosen time bins
    % 4. Generate summaries and comparisons for each time bin
    % 5. Save all results to Excel
    %
    % Written for sleep analysis with time bins, building on analyzeSleepData function
    
    % Define treatment groups based on serial numbers
    suvoSerials = {'164a', '164b', '1676'};
    
    % Ask user which time bin resolution to use
    binSizeOptions = {'1-hour bins', '3-hour bins', '6-hour bins', '12-hour bins'};
    [binSizeIdx, binOk] = listdlg('PromptString', 'Select time bin resolution:', ...
                         'SelectionMode', 'single', ...
                         'ListString', binSizeOptions);
    
    if ~binOk || isempty(binSizeIdx)
        error('No time bin resolution selected. Analysis aborted.');
    end
    
    % Calculate the bin size in hours
    binSizeHours = [1, 3, 6, 12];
    selectedBinSize = binSizeHours(binSizeIdx);
    
    % Calculate number of bins in a 12-hour recording
    numBins = 12 / selectedBinSize;
    
    fprintf('Selected time resolution: %d-hour bins (%d bins total)\n', selectedBinSize, numBins);
    
    % Ask user to select the baseline folder
    disp('Select the baseline folder');
    baselineFolder = uigetdir('', 'Select the baseline folder');
    
    % Ask user which extinction days to analyze
    extDayOptions = {'Ext_D1', 'Ext_D2', 'Ext_D3', 'Ext_D4', 'Ext_D5'};
    [extDaysIndices, ok] = listdlg('PromptString', 'Select extinction days to analyze:', ...
                          'SelectionMode', 'multiple', ...
                          'ListString', extDayOptions);
    
    if ~ok || isempty(extDaysIndices)
        error('No extinction days selected. Analysis aborted.');
    end
    
    selectedExtDays = extDayOptions(extDaysIndices);
    
    % Create a cell array to store folder paths
    extFolders = cell(length(selectedExtDays), 1);
    
    % Get folder paths for selected extinction days
    for i = 1:length(selectedExtDays)
        prompt = sprintf('Select folder for %s', selectedExtDays{i});
        extFolders{i} = uigetdir('', prompt);
    end
    
    % Process baseline data with time bins
    baselineData = processFolder(baselineFolder, suvoSerials, selectedBinSize, numBins);
    
    % Get baseline metrics by treatment
    baselineVehData = baselineData(strcmp({baselineData.Treatment}, 'Vehicle'));
    baselineSuvoData = baselineData(strcmp({baselineData.Treatment}, 'Suvorexant'));
    
    fprintf('Baseline data: %d total subjects\n', length(baselineData));
    fprintf('- Vehicle group: %d subjects\n', length(baselineVehData));
    fprintf('- Suvorexant group: %d subjects\n', length(baselineSuvoData));

    % Process all extinction day data
    extDayData = cell(length(selectedExtDays), 1);
    for i = 1:length(selectedExtDays)
        fprintf('Processing %s data...\n', selectedExtDays{i});
        extDayData{i} = processFolder(extFolders{i}, suvoSerials, selectedBinSize, numBins);
        
        % Add debug info
        extVehData = extDayData{i}(strcmp({extDayData{i}.Treatment}, 'Vehicle'));
        extSuvoData = extDayData{i}(strcmp({extDayData{i}.Treatment}, 'Suvorexant'));
        fprintf('%s data: %d total subjects\n', selectedExtDays{i}, length(extDayData{i}));
        fprintf('  - Vehicle group: %d subjects\n', length(extVehData));
        fprintf('  - Suvorexant group: %d subjects\n', length(extSuvoData));
    end
    
    % Create results tables for different comparisons for each time bin
    
    % 1. Baseline vs Vehicle (for all extinction days, all time bins)
    baselineVsVehTables = cell(length(selectedExtDays), numBins);
    for i = 1:length(selectedExtDays)
        extVehData = extDayData{i}(strcmp({extDayData{i}.Treatment}, 'Vehicle'));
        
        for bin = 1:numBins
            binLabel = sprintf('Bin%d', bin);
            fprintf('Creating Baseline vs Vehicle table for %s, %s...\n', selectedExtDays{i}, binLabel);
            fprintf('  - Baseline group: %d subjects\n', length(baselineVehData));
            fprintf('  - Vehicle group: %d subjects\n', length(extVehData));
            
            % Create comparison table with the right group names
            tempTable = compareGroupsByBin(baselineVehData, extVehData, 'Baseline', 'Vehicle', selectedExtDays{i}, bin);
            
            % Verify table structure
            fprintf('  - Table size: %d rows, %d columns\n', height(tempTable), width(tempTable));
            fprintf('  - Variable names: %s\n', strjoin(tempTable.Properties.VariableNames, ', '));
            
            baselineVsVehTables{i, bin} = tempTable;
        end
    end
    
    % 2. Suvorexant vs Vehicle (for each extinction day, all time bins)
    suvoVsVehTables = cell(length(selectedExtDays), numBins);
    for i = 1:length(selectedExtDays)
        extSuvoData = extDayData{i}(strcmp({extDayData{i}.Treatment}, 'Suvorexant'));
        extVehData = extDayData{i}(strcmp({extDayData{i}.Treatment}, 'Vehicle'));
        
        for bin = 1:numBins
            suvoVsVehTables{i, bin} = compareGroupsByBin(extSuvoData, extVehData, 'Suvorexant', 'Vehicle', selectedExtDays{i}, bin);
        end
    end
    
    % Display results
    displayResultsByTimeBin(baselineVsVehTables, suvoVsVehTables, selectedExtDays, numBins);
    
    % Ask if user wants to save the results
    answer = questdlg('Do you want to save the results to Excel?', 'Save Results', 'Yes', 'No', 'Yes');
    if strcmp(answer, 'Yes')
        [fileName, pathName] = uiputfile('*.xlsx', 'Save Results As');
        if fileName ~= 0
            % Pass baselineData and extDayData to save individual subject data
            saveTimeBinnedResultsToExcel(baselineVsVehTables, suvoVsVehTables, ...
                selectedExtDays, fullfile(pathName, fileName), ...
                baselineData, extDayData, selectedBinSize, numBins);
            disp(['Results saved to ' fullfile(pathName, fileName)]);
        end
    end
end