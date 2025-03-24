function saveTimeBinnedResultsToExcel(baselineVsVehTables, suvoVsVehTables, conditionNames, filePath, baselineData, extDayData, binSizeHours, numBins)
    % Save time-binned results to Excel file with multiple sheets
    % This function combines the original functionality with enhanced Prism-friendly output
    %
    % Parameters:
    %   baselineVsVehTables - Cell array of comparison tables for Baseline vs Vehicle
    %   suvoVsVehTables - Cell array of comparison tables for Suvorexant vs Vehicle
    %   conditionNames - Cell array of condition names
    %   filePath - Full path to Excel file to create
    %   baselineData - Array of baseline subject data structures
    %   extDayData - Cell array of extinction day subject data structures
    %   binSizeHours - Size of time bins in hours
    %   numBins - Number of time bins
    
    % Single try-catch block for all Excel operations
    try
        % First check if file exists and delete it if it does
        if exist(filePath, 'file')
            delete(filePath);
        end
        
        % Create a more detailed README with actual condition names and bin numbers
        readme = {'This Excel file contains time-binned sleep analysis data from a cocaine CPP study with suvorexant treatment.'};
        readme = [readme; {''}];
        readme = [readme; {sprintf('TIME BINS: Data analyzed in %d-hour bins (%d bins total)', binSizeHours, numBins)}];
        readme = [readme; {''}];
        readme = [readme; {'SHEETS:'}];
        readme = [readme; {'- README: This information'}];
        readme = [readme; {'- BL_vs_Veh_Summary: Effect of cocaine across all time bins'}];
        readme = [readme; {'- Suvo_vs_Veh_Summary: Effect of suvorexant across all time bins'}];
        readme = [readme; {''}];

        % Add condition and bin specific descriptions
        for i = 1:length(conditionNames)
            for bin = 1:numBins
                readme = [readme; {sprintf('- %s_Bin%d_BL_vs_Veh: Effect of cocaine on %s during bin %d', conditionNames{i}, bin, conditionNames{i}, bin)}];
                readme = [readme; {sprintf('- %s_Bin%d_Suvo_vs_Veh: Effect of suvorexant on %s during bin %d', conditionNames{i}, bin, conditionNames{i}, bin)}];
            end
        end

        % Add information about transition sheets
        readme = [readme; {''}];
        for bin = 1:numBins
            readme = [readme; {sprintf('- Transitions_Bin%d: Stage transitions data with p-values for bin %d', bin, bin)}];
        end

        % Add information about Prism-formatted sheets
        readme = [readme; {''}];
        readme = [readme; {'PRISM-FORMATTED SHEETS:'}];
        metrics = {
            {'PercentTime', 'Wake', 'Wake_Time_Pct'},
            {'PercentTime', 'SWS', 'SWS_Time_Pct'},
            {'PercentTime', 'REM', 'REM_Time_Pct'},
            {'Bouts', 'Wake', 'Wake_Bouts'},
            {'Bouts', 'SWS', 'SWS_Bouts'},
            {'Bouts', 'REM', 'REM_Bouts'},
            {'BoutLength', 'Wake', 'Wake_Bout_Length'},
            {'BoutLength', 'SWS', 'SWS_Bout_Length'},
            {'BoutLength', 'REM', 'REM_Bout_Length'},
            {'Transitions', 'Count', 'Transition_Count'}
        };
        
        for m = 1:length(metrics)
            readme = [readme; {sprintf('- %s: Individual data for %s', metrics{m}{3}, metrics{m}{3})}];
        end
        
        readme = [readme; {'- Prism_Combined_Data: Means for all metrics by group and bin'}];

        % Add metrics information
        readme = [readme; {''}];
        readme = [readme; {'METRICS:'}];
        readme = [readme; {'- TotalTime_min: Total time spent in each sleep stage (minutes)'}];
        readme = [readme; {'- PercentTime: Percentage of recording time in each stage'}];
        readme = [readme; {'- Bouts: Number of episodes of each stage'}];
        readme = [readme; {'- BoutLength_min: Average duration of bouts for each stage (minutes)'}];
        readme = [readme; {'- Trans_Count: Number of transitions between stages'}];
        readme = [readme; {'- Latencies_min: Time to first occurrence of each stage (minutes)'}];
        readme = [readme; {'- Latency2Min_min: Time to first 2-minute bout of each stage (minutes)'}];
        readme = [readme; {''}];
        readme = [readme; {sprintf('GENERATED: %s', string(datetime('now')))}];
        
        % Write the README sheet
        xlswrite(filePath, readme, 'README');
        disp('Created README sheet.');
        
        % Create summary sheets
        disp('Creating summary sheets...');
        createTimeBinSummarySheet(baselineVsVehTables, suvoVsVehTables, conditionNames, numBins, filePath);
        
        % For each bin, create:
        % 1. Enhanced transition sheets with p-values
        % 2. The original comparison tables
        for bin = 1:numBins
            % Create enhanced transition sheets with p-values
            createEnhancedTransitionSheet(baselineData, extDayData, conditionNames, filePath, bin);
            
            % For each condition, write the original comparison tables
            for i = 1:length(conditionNames)
                % Baseline vs Vehicle tables for this condition and bin
                sheetName = sprintf('%s_Bin%d_BL_vs_Veh', conditionNames{i}, bin);
                sheetName = strrep(sheetName, ' ', '_'); % Replace spaces
                sheetName = validateSheetName(sheetName);
                
                disp(['Writing sheet: ' sheetName]);
                % Make sure the table is not empty
                if ~isempty(baselineVsVehTables{i, bin})
                    try
                        writetable(baselineVsVehTables{i, bin}, filePath, 'Sheet', sheetName);
                    catch e1
                        warning('Error writing BL vs Veh table: %s', e1.message);
                        try
                            % Fallback method
                            headers = baselineVsVehTables{i, bin}.Properties.VariableNames;
                            data = table2cell(baselineVsVehTables{i, bin});
                            xlswrite(filePath, headers, sheetName, 'A1');
                            xlswrite(filePath, data, sheetName, 'A2');
                        catch e2
                            warning('Fallback method also failed: %s', e2.message);
                        end
                    end
                else
                    warning('Baseline vs Vehicle table for %s, Bin %d is empty', conditionNames{i}, bin);
                end
                
                % Suvorexant vs Vehicle tables for this condition and bin
                sheetName = sprintf('%s_Bin%d_Suvo_vs_Veh', conditionNames{i}, bin);
                sheetName = strrep(sheetName, ' ', '_'); % Replace spaces
                sheetName = validateSheetName(sheetName);
                
                disp(['Writing sheet: ' sheetName]);
                % Make sure the table is not empty
                if ~isempty(suvoVsVehTables{i, bin})
                    try
                        writetable(suvoVsVehTables{i, bin}, filePath, 'Sheet', sheetName);
                    catch e1
                        warning('Error writing Suvo vs Veh table: %s', e1.message);
                        try
                            % Fallback method
                            headers = suvoVsVehTables{i, bin}.Properties.VariableNames;
                            data = table2cell(suvoVsVehTables{i, bin});
                            xlswrite(filePath, headers, sheetName, 'A1');
                            xlswrite(filePath, data, sheetName, 'A2');
                        catch e2
                            warning('Fallback method also failed: %s', e2.message);
                        end
                    end
                else
                    warning('Suvorexant vs Vehicle table for %s, Bin %d is empty', conditionNames{i}, bin);
                end
            end
        end
        
        % Create Prism-formatted sheets
        disp('Creating Prism-formatted sheets...');
        createPrismFormattedSheets(baselineData, extDayData, conditionNames, filePath, binSizeHours, numBins);
        
        disp('Excel file writing completed successfully!');
        
    catch mainError
        warning('Error saving Excel file: %s', mainError.message);
        disp('Trying alternative method to save data as separate CSV files...');
        
        % Alternative method: Save as separate CSV files
        [pathName, nameOnly, ~] = fileparts(filePath);
        baseFileName = fullfile(pathName, nameOnly);
        
        % Save README
        writecell(readme, [baseFileName '_README.csv']);
        
        % Save Summary (create it first)
        summaryTable = createTimeBinSummaryTableOnly(baselineVsVehTables, suvoVsVehTables, conditionNames, numBins);
        writetable(summaryTable, [baseFileName '_Summary.csv']);
        
        % For each bin, save separate tables
        for bin = 1:numBins
            % For each condition, save separate tables
            for i = 1:length(conditionNames)
                % Save comparison tables
                baseName = sprintf('%s_Bin%d_BL_vs_Veh.csv', conditionNames{i}, bin);
                if ~isempty(baselineVsVehTables{i, bin})
                    writetable(baselineVsVehTables{i, bin}, [baseFileName '_' baseName]);
                end
                
                suvoName = sprintf('%s_Bin%d_Suvo_vs_Veh.csv', conditionNames{i}, bin);
                if ~isempty(suvoVsVehTables{i, bin})
                    writetable(suvoVsVehTables{i, bin}, [baseFileName '_' suvoName]);
                end
            end
        end
        disp(['Data saved as separate CSV files in ' pathName]);
    end
end

function createTimeBinSummarySheet(baselineVsVehTables, suvoVsVehTables, conditionNames, numBins, filePath)
    % Create a simplified summary sheet that will work reliably
    
    % Create the baseline vs vehicle summary
    blSummary = {'Condition', 'Bin', 'Metric', 'Baseline Mean', 'Vehicle Mean', 'Percent Change', 'P-Value', 'Significant'};
    blRows = {};
    
    % Create the suvorexant vs vehicle summary
    suvoSummary = {'Condition', 'Bin', 'Metric', 'Suvorexant Mean', 'Vehicle Mean', 'Percent Change', 'P-Value', 'Significant'};
    suvoRows = {};
    
    % Define key metrics to display
    keyMetrics = {
        {'BoutLength', 'BoutLength_min', 'Wake', 'Wake bout length (min)'};
        {'BoutLength', 'BoutLength_min', 'SWS', 'SWS bout length (min)'};
        {'BoutLength', 'BoutLength_min', 'REM', 'REM bout length (min)'};
        {'PercentTime', 'PercentTime', 'Wake', 'Wake time %'};
        {'PercentTime', 'PercentTime', 'SWS', 'SWS time %'};
        {'PercentTime', 'PercentTime', 'REM', 'REM time %'};
        {'Bouts', 'Bouts', 'Wake', 'Wake bouts'};
        {'Bouts', 'Bouts', 'SWS', 'SWS bouts'};
        {'Bouts', 'Bouts', 'REM', 'REM bouts'};
        {'Transitions', 'Trans_Count', 'All', 'Transitions count'};
    };
    
    % Loop through each condition and bin
    for c = 1:length(conditionNames)
        for bin = 1:numBins
            % Process baseline vs vehicle
            if ~isempty(baselineVsVehTables{c, bin}) && height(baselineVsVehTables{c, bin}) > 0
                for m = 1:length(keyMetrics)
                    % Find the row for this metric
                    idx = find(strcmp(baselineVsVehTables{c, bin}.Category, keyMetrics{m}{1}) & ...
                             strcmp(baselineVsVehTables{c, bin}.Metric, keyMetrics{m}{2}) & ...
                             strcmp(baselineVsVehTables{c, bin}.Stage, keyMetrics{m}{3}));
                    
                    if ~isempty(idx)
                        blMean = baselineVsVehTables{c, bin}.Baseline_Mean(idx);
                        vehMean = baselineVsVehTables{c, bin}.Vehicle_Mean(idx);
                        pctChange = baselineVsVehTables{c, bin}.Percent_Change(idx);
                        pValue = baselineVsVehTables{c, bin}.P_Value(idx);
                        
                        % Determine significance
                        if pValue < 0.05
                            sig = '*';
                        elseif pValue < 0.1
                            sig = '+';
                        else
                            sig = '';
                        end
                        
                        % Add to summary rows
                        blRows{end+1} = {conditionNames{c}, num2str(bin), keyMetrics{m}{4}, blMean, vehMean, pctChange, pValue, sig};
                    end
                end
            end
            
            % Process suvorexant vs vehicle
            if ~isempty(suvoVsVehTables{c, bin}) && height(suvoVsVehTables{c, bin}) > 0
                for m = 1:length(keyMetrics)
                    % Find the row for this metric
                    idx = find(strcmp(suvoVsVehTables{c, bin}.Category, keyMetrics{m}{1}) & ...
                             strcmp(suvoVsVehTables{c, bin}.Metric, keyMetrics{m}{2}) & ...
                             strcmp(suvoVsVehTables{c, bin}.Stage, keyMetrics{m}{3}));
                    
                    if ~isempty(idx)
                        suvoMean = suvoVsVehTables{c, bin}.Suvorexant_Mean(idx);
                        vehMean = suvoVsVehTables{c, bin}.Vehicle_Mean(idx);
                        pctChange = suvoVsVehTables{c, bin}.Percent_Change(idx);
                        pValue = suvoVsVehTables{c, bin}.P_Value(idx);
                        
                        % Determine significance
                        if pValue < 0.05
                            sig = '*';
                        elseif pValue < 0.1
                            sig = '+';
                        else
                            sig = '';
                        end
                        
                        % Add to summary rows
                        suvoRows{end+1} = {conditionNames{c}, num2str(bin), keyMetrics{m}{4}, suvoMean, vehMean, pctChange, pValue, sig};
                    end
                end
            end
        end
    end
    
    % Write summaries to Excel
    try
        % Write headers
        writecell(blSummary, filePath, 'Sheet', 'BL_vs_Veh_Summary', 'Range', 'A1');
        
        % Convert cell arrays to tables for better Excel compatibility
        if ~isempty(blRows)
            blTable = cell2table(blRows, 'VariableNames', blSummary);
            writetable(blTable, filePath, 'Sheet', 'BL_vs_Veh_Summary', 'Range', 'A2', 'WriteMode', 'append');
        end
        
        writecell(suvoSummary, filePath, 'Sheet', 'Suvo_vs_Veh_Summary', 'Range', 'A1');
        if ~isempty(suvoRows)
            suvoTable = cell2table(suvoRows, 'VariableNames', suvoSummary);
            writetable(suvoTable, filePath, 'Sheet', 'Suvo_vs_Veh_Summary', 'Range', 'A2', 'WriteMode', 'append');
        end
        
        % Write a legend
        writecell({'* p < 0.05, + p < 0.1', ...
                   '', ...
                   'Summary Tables:', ...
                   '- BL_vs_Veh_Summary: Effect of cocaine across all time bins', ...
                   '- Suvo_vs_Veh_Summary: Effect of suvorexant treatment in this time bin'}, ...
                   filePath, 'Sheet', 'Summary', 'Range', 'A1');
        
        disp('Summary tables created successfully.');
    catch e
        warning('Error writing summary tables: %s', e.message);
        
        % Fallback method using xlswrite
        try
            % Write BL vs Veh summary
            xlswrite(filePath, blSummary, 'BL_vs_Veh_Summary', 'A1');
            if ~isempty(blRows)
                xlswrite(filePath, blRows, 'BL_vs_Veh_Summary', 'A2');
            end
            
            % Write Suvo vs Veh summary
            xlswrite(filePath, suvoSummary, 'Suvo_vs_Veh_Summary', 'A1');
            if ~isempty(suvoRows)
                xlswrite(filePath, suvoRows, 'Suvo_vs_Veh_Summary', 'A2');
            end
            
            disp('Summary tables created using fallback method.');
        catch e2
            warning('Fallback method for summaries also failed: %s', e2.message);
        end
    end
end

function createEnhancedTransitionSheet(baselineData, extDayData, conditionNames, filePath, binIndex)
    % Create an enhanced transitions sheet with individual rat data and p-values
    %
    % Parameters:
    %   baselineData - Array of baseline subject data structures
    %   extDayData - Cell array of extinction day subject data structures
    %   conditionNames - Cell array of condition names
    %   filePath - Path to Excel file
    %   binIndex - Index of the time bin
    
    % Bin field name in the structure
    binField = sprintf('Bin%d', binIndex);
    
    % Create sheet name
    sheetName = sprintf('Transitions_Bin%d', binIndex);
    sheetName = validateSheetName(sheetName);
    
    % Define the transitions of interest
    transitions = {
        'Wake_to_SWS', 
        'Wake_to_REM', 
        'SWS_to_Wake', 
        'SWS_to_REM', 
        'REM_to_Wake', 
        'REM_to_SWS'
    };
    
    % Get the data for each group
    baselineVeh = baselineData(strcmp({baselineData.Treatment}, 'Vehicle'));
    baselineSuvo = baselineData(strcmp({baselineData.Treatment}, 'Suvorexant'));
    
    % Initialize the output data structure
    outputData = {};
    
    % Add header row
    headerRow = {'Group', 'Subject'};
    for t = 1:length(transitions)
        headerRow{end+1} = transitions{t};
    end
    outputData{1} = headerRow;
    
    % Process each condition
    rowIdx = 2;
    
    % First, add baseline data
    % Add baseline vehicle data
    for s = 1:length(baselineVeh)
        if isfield(baselineVeh(s), binField) && isfield(baselineVeh(s).(binField), 'DetailedTransitions')
            rowData = {'Baseline Vehicle', baselineVeh(s).SerialNumber};
            
            % Add transition data
            for t = 1:length(transitions)
                transField = transitions{t};
                if isfield(baselineVeh(s).(binField).DetailedTransitions, transField)
                    rowData{end+1} = baselineVeh(s).(binField).DetailedTransitions.(transField);
                else
                    rowData{end+1} = NaN;
                end
            end
            
            outputData{rowIdx} = rowData;
            rowIdx = rowIdx + 1;
        end
    end
    
    % Add baseline suvorexant data
    for s = 1:length(baselineSuvo)
        if isfield(baselineSuvo(s), binField) && isfield(baselineSuvo(s).(binField), 'DetailedTransitions')
            rowData = {'Baseline Suvorexant', baselineSuvo(s).SerialNumber};
            
            % Add transition data
            for t = 1:length(transitions)
                transField = transitions{t};
                if isfield(baselineSuvo(s).(binField).DetailedTransitions, transField)
                    rowData{end+1} = baselineSuvo(s).(binField).DetailedTransitions.(transField);
                else
                    rowData{end+1} = NaN;
                end
            end
            
            outputData{rowIdx} = rowData;
            rowIdx = rowIdx + 1;
        end
    end
    
    % For each condition, add data
    for c = 1:length(conditionNames)
        % Vehicle group
        vehData = extDayData{c}(strcmp({extDayData{c}.Treatment}, 'Vehicle'));
        
        for s = 1:length(vehData)
            if isfield(vehData(s), binField) && isfield(vehData(s).(binField), 'DetailedTransitions')
                rowData = {[conditionNames{c} ' Vehicle'], vehData(s).SerialNumber};
                
                % Add transition data
                for t = 1:length(transitions)
                    transField = transitions{t};
                    if isfield(vehData(s).(binField).DetailedTransitions, transField)
                        rowData{end+1} = vehData(s).(binField).DetailedTransitions.(transField);
                    else
                        rowData{end+1} = NaN;
                    end
                end
                
                outputData{rowIdx} = rowData;
                rowIdx = rowIdx + 1;
            end
        end
        
        % Suvorexant group
        suvoData = extDayData{c}(strcmp({extDayData{c}.Treatment}, 'Suvorexant'));
        
        for s = 1:length(suvoData)
            if isfield(suvoData(s), binField) && isfield(suvoData(s).(binField), 'DetailedTransitions')
                rowData = {[conditionNames{c} ' Suvorexant'], suvoData(s).SerialNumber};
                
                % Add transition data
                for t = 1:length(transitions)
                    transField = transitions{t};
                    if isfield(suvoData(s).(binField).DetailedTransitions, transField)
                        rowData{end+1} = suvoData(s).(binField).DetailedTransitions.(transField);
                    else
                        rowData{end+1} = NaN;
                    end
                end
                
                outputData{rowIdx} = rowData;
                rowIdx = rowIdx + 1;
            end
        end
        
        % Add empty row between conditions
        outputData{rowIdx} = cell(1, length(headerRow));
        rowIdx = rowIdx + 1;
    end
    % Add an empty row before p-values
    outputData{rowIdx} = cell(1, length(headerRow));
    rowIdx = rowIdx + 1;
    
    % Add p-value section
    outputData{rowIdx} = {'Statistical Comparisons', ''};
    rowIdx = rowIdx + 1;
    
    % Add headers for p-value section
    outputData{rowIdx} = {'Comparison', ''};
    for t = 1:length(transitions)
        outputData{rowIdx}{end+1} = transitions{t};
    end
    rowIdx = rowIdx + 1;
    
    % Calculate p-values for each condition and transition
    for c = 1:length(conditionNames)
        % Baseline vs Vehicle comparison
        baselineVehValues = extractTransitionValues(baselineVeh, transitions, binField);
        extVehData = extDayData{c}(strcmp({extDayData{c}.Treatment}, 'Vehicle'));
        extVehValues = extractTransitionValues(extVehData, transitions, binField);
        
        % Row for Baseline vs Vehicle
        pvalRow = {[conditionNames{c} ' BL vs Veh'], ''};
        for t = 1:length(transitions)
            % Extract values for this transition
            baseVals = baselineVehValues{t};
            extVals = extVehValues{t};
            
            % Calculate p-value
            if length(baseVals) >= 2 && length(extVals) >= 2
                [~, pVal] = ttest2(baseVals, extVals);
                
                % Format the p-value with significance markers
                if pVal < 0.05
                    pvalRow{end+1} = sprintf('%.3f *', pVal);
                elseif pVal < 0.1
                    pvalRow{end+1} = sprintf('%.3f +', pVal);
                else
                    pvalRow{end+1} = sprintf('%.3f', pVal);
                end
            else
                pvalRow{end+1} = 'N/A';
            end
        end
        outputData{rowIdx} = pvalRow;
        rowIdx = rowIdx + 1;
        
        % Suvorexant vs Vehicle comparison
        extSuvoData = extDayData{c}(strcmp({extDayData{c}.Treatment}, 'Suvorexant'));
        extSuvoValues = extractTransitionValues(extSuvoData, transitions, binField);
        
        % Row for Suvorexant vs Vehicle
        pvalRow = {[conditionNames{c} ' Suvo vs Veh'], ''};
        for t = 1:length(transitions)
            % Extract values for this transition
            vehVals = extVehValues{t};
            suvoVals = extSuvoValues{t};
            
            % Calculate p-value
            if length(vehVals) >= 2 && length(suvoVals) >= 2
                [~, pVal] = ttest2(vehVals, suvoVals);
                
                % Format the p-value with significance markers
                if pVal < 0.05
                    pvalRow{end+1} = sprintf('%.3f *', pVal);
                elseif pVal < 0.1
                    pvalRow{end+1} = sprintf('%.3f +', pVal);
                else
                    pvalRow{end+1} = sprintf('%.3f', pVal);
                end
            else
                pvalRow{end+1} = 'N/A';
            end
        end
        outputData{rowIdx} = pvalRow;
        rowIdx = rowIdx + 1;
    end
    
    % Add a legend for p-values
    rowIdx = rowIdx + 2;
    outputData{rowIdx} = {'* p < 0.05', ''};
    rowIdx = rowIdx + 1;
    outputData{rowIdx} = {'+ p < 0.1', ''};
    
    % Write to Excel
    try
        % Convert cell array to format suitable for xlswrite
        writeData = cell(length(outputData), length(headerRow));
        for r = 1:length(outputData)
            for c = 1:min(length(outputData{r}), length(headerRow))
                writeData{r,c} = outputData{r}{c};
            end
        end
        
        % Write data
        xlswrite(filePath, writeData, sheetName);
        disp(['Created enhanced transition sheet for bin ' num2str(binIndex)]);
    catch e
        warning('Error writing enhanced transition sheet for bin %d: %s', binIndex, e.message);
    end
end

function transValues = extractTransitionValues(subjectData, transitions, binField)
    % Extract transition values for a group of subjects
    %
    % Parameters:
    %   subjectData - Array of subject data structures
    %   transitions - Cell array of transition names
    %   binField - Name of the bin field (e.g., 'Bin1')
    %
    % Returns:
    %   transValues - Cell array of transition values for each transition
    
    % Initialize output
    transValues = cell(length(transitions), 1);
    for t = 1:length(transitions)
        transValues{t} = [];
    end
    
    % Extract data for each subject
    for s = 1:length(subjectData)
        if isfield(subjectData(s), binField) && isfield(subjectData(s).(binField), 'DetailedTransitions')
            for t = 1:length(transitions)
                transField = transitions{t};
                if isfield(subjectData(s).(binField).DetailedTransitions, transField)
                    transValues{t}(end+1) = subjectData(s).(binField).DetailedTransitions.(transField);
                end
            end
        end
    end
end

function createPrismFormattedSheets(baselineData, extDayData, conditionNames, filePath, binSizeHours, numBins)
    % Define metrics to summarize
    metrics = {
        {'PercentTime', 'Wake', 'Wake_Time_Pct'},
        {'PercentTime', 'SWS', 'SWS_Time_Pct'},
        {'PercentTime', 'REM', 'REM_Time_Pct'},
        {'Bouts', 'Wake', 'Wake_Bouts'},
        {'Bouts', 'SWS', 'SWS_Bouts'},
        {'Bouts', 'REM', 'REM_Bouts'},
        {'BoutLength', 'Wake', 'Wake_Bout_Length'},
        {'BoutLength', 'SWS', 'SWS_Bout_Length'},
        {'BoutLength', 'REM', 'REM_Bout_Length'},
        {'Transitions', 'Count', 'Transition_Count'}
    };
    
    % Create each metric sheet
    for m = 1:length(metrics)
        metricName = metrics{m}{1};
        stageName = metrics{m}{2};
        displayName = metrics{m}{3};
        
        % Prepare data structure
        allData = {};
        currentRow = 1;
        
        % Create header row
        header = {'Bin (Hours)', 'Condition', 'Treatment', 'Serial Number'};
        for bin = 1:numBins
            header{end+1} = sprintf('Bin %d', bin);
        end
        allData{currentRow} = header;
        currentRow = currentRow + 1;
        
        % First process all baseline data together
        % Note: No treatment during baseline
        baselineVeh = baselineData(strcmp({baselineData.Treatment}, 'Vehicle'));
        for s = 1:length(baselineVeh)
            rowData = cell(1, length(header));
            rowData{1} = binSizeHours;
            rowData{2} = 'Baseline';
            rowData{3} = 'Vehicle';  %// Since these are vehicle subjects
            rowData{4} = baselineVeh(s).SerialNumber;
            
            for bin = 1:numBins
                binField = sprintf('Bin%d', bin);
                
                % Extract individual metric value
                if isfield(baselineVeh(s), binField)
                    binData = baselineVeh(s).(binField);
                    
                    % Special case for transition count
                    if strcmp(metricName, 'Transitions') && strcmp(stageName, 'Count')
                        if isfield(binData.Transitions, 'Count')
                            rowData{4+bin} = binData.Transitions.Count;
                        else
                            rowData{4+bin} = NaN;
                        end
                    else
                        % Normal case for other metrics
                        if isfield(binData, metricName) && isfield(binData.(metricName), stageName)
                            rowData{4+bin} = binData.(metricName).(stageName);
                        else
                            rowData{4+bin} = NaN;
                        end
                    end
                else
                    rowData{4+bin} = NaN;
                end
            end
            
            allData{currentRow} = rowData;
            currentRow = currentRow + 1;
        end
        
        % Then process extinction day data
        for c = 1:length(conditionNames)
            % First process all vehicle data
            extVehicle = extDayData{c}(strcmp({extDayData{c}.Treatment}, 'Vehicle'));
            for s = 1:length(extVehicle)
                rowData = cell(1, length(header));
                rowData{1} = binSizeHours;
                rowData{2} = conditionNames{c};
                rowData{3} = 'Vehicle';
                rowData{4} = extVehicle(s).SerialNumber;
                
                for bin = 1:numBins
                    binField = sprintf('Bin%d', bin);
                    
                    % Extract individual metric value
                    if isfield(extVehicle(s), binField)
                        binData = extVehicle(s).(binField);
                        
                        % Special case for transition count
                        if strcmp(metricName, 'Transitions') && strcmp(stageName, 'Count')
                            if isfield(binData.Transitions, 'Count')
                                rowData{4+bin} = binData.Transitions.Count;
                            else
                                rowData{4+bin} = NaN;
                            end
                        else
                            % Normal case for other metrics
                            if isfield(binData, metricName) && isfield(binData.(metricName), stageName)
                                rowData{4+bin} = binData.(metricName).(stageName);
                            else
                                rowData{4+bin} = NaN;
                            end
                        end
                    else
                        rowData{4+bin} = NaN;
                    end
                end
                
                allData{currentRow} = rowData;
                currentRow = currentRow + 1;
            end
            
            % Then process suvorexant data
            extSuvo = extDayData{c}(strcmp({extDayData{c}.Treatment}, 'Suvorexant'));
            for s = 1:length(extSuvo)
                rowData = cell(1, length(header));
                rowData{1} = binSizeHours;
                rowData{2} = conditionNames{c};
                rowData{3} = 'Suvorexant';
                rowData{4} = extSuvo(s).SerialNumber;
                
                for bin = 1:numBins
                    binField = sprintf('Bin%d', bin);
                    
                    % Extract individual metric value
                    if isfield(extSuvo(s), binField)
                        binData = extSuvo(s).(binField);
                        
                        % Special case for transition count
                        if strcmp(metricName, 'Transitions') && strcmp(stageName, 'Count')
                            if isfield(binData.Transitions, 'Count')
                                rowData{4+bin} = binData.Transitions.Count;
                            else
                                rowData{4+bin} = NaN;
                            end
                        else
                            % Normal case for other metrics
                            if isfield(binData, metricName) && isfield(binData.(metricName), stageName)
                                rowData{4+bin} = binData.(metricName).(stageName);
                            else
                                rowData{4+bin} = NaN;
                            end
                        end
                    else
                        rowData{4+bin} = NaN;
                    end
                end
                
                allData{currentRow} = rowData;
                currentRow = currentRow + 1;
            end
        end
        
        % Write to Excel
        try
            % Convert cell array to format suitable for xlswrite
            writeData = cell(length(allData), length(allData{1}));
            for r = 1:length(allData)
                for c = 1:length(allData{r})
                    writeData(r,c) = allData{r}(c);
                end
            end
            
            % Validate sheet name
            sheetName = validateSheetName(displayName);
            
            % Write to Excel
            xlswrite(filePath, writeData, sheetName);
            disp(['Created sheet for ' displayName]);
        catch e
            warning('Error writing sheet for %s: %s', displayName, e.message);
            try
                % Fallback method - write direct CSV
                [path, name, ~] = fileparts(filePath);
                csvPath = fullfile(path, [name '_' displayName '.csv']);
                
                % Open file
                fid = fopen(csvPath, 'w');
                
                % Write header
                for i = 1:length(header)-1
                    fprintf(fid, '%s,', header{i});
                end
                fprintf(fid, '%s\n', header{end});
                
                % Write data rows
                for r = 2:length(allData)  % Skip header
                    row = allData{r};
                    for c = 1:length(row)-1
                        if isnumeric(row{c})
                            fprintf(fid, '%.6f,', row{c});
                        else
                            fprintf(fid, '%s,', row{c});
                        end
                    end
                    
                    % Last column
                    if isnumeric(row{end})
                        fprintf(fid, '%.6f\n', row{end});
                    else
                        fprintf(fid, '%s\n', row{end});
                    end
                end
                
                % Close file
                fclose(fid);
                disp(['Created CSV for ' displayName ' at: ' csvPath]);
            catch fe
                warning('Fallback method also failed: %s', fe.message);
            end
        end
    end
    
    % Additionally create a combined Prism-formatted sheet
    createPrismCombinedSheet(baselineData, extDayData, conditionNames, filePath, binSizeHours, numBins);
end

function createPrismCombinedSheet(baselineData, extDayData, conditionNames, filePath, binSizeHours, numBins)
    % Define metrics to summarize
    metrics = {
        {'PercentTime', 'Wake', 'Wake Time %'},
        {'PercentTime', 'SWS', 'SWS Time %'},
        {'PercentTime', 'REM', 'REM Time %'},
        {'Bouts', 'Wake', 'Wake Bouts'},
        {'Bouts', 'SWS', 'SWS Bouts'},
        {'Bouts', 'REM', 'REM Bouts'},
        {'BoutLength', 'Wake', 'Wake Bout Length'},
        {'BoutLength', 'SWS', 'SWS Bout Length'},
        {'BoutLength', 'REM', 'REM Bout Length'},
        {'Transitions', 'Count', 'Transition Count'}
    };
    
    % Prepare combined sheet
    combinedSheetName = 'Prism_Combined_Data';
    combinedData = {};
    
    % For each metric, create a section
    currentRow = 1;
    for m = 1:length(metrics)
        metricName = metrics{m}{1};
        stageName = metrics{m}{2};
        displayName = metrics{m}{3};
        
        % Add metric name as header
        combinedData{currentRow, 1} = displayName;
        currentRow = currentRow + 1;
        
        % Create headers for different conditions
        headers = {'Bin'};
        conditionGroups = {};
        
        % Add baseline
        headers{end+1} = 'Baseline';
        conditionGroups{end+1} = baselineData;
        
        % Add each condition's Vehicle and Suvorexant groups
        for c = 1:length(conditionNames)
            extVehicle = extDayData{c}(strcmp({extDayData{c}.Treatment}, 'Vehicle'));
            extSuvo = extDayData{c}(strcmp({extDayData{c}.Treatment}, 'Suvorexant'));
            
            headers{end+1} = [conditionNames{c} ' Vehicle'];
            conditionGroups{end+1} = extVehicle;
            
            headers{end+1} = [conditionNames{c} ' Suvorexant'];
            conditionGroups{end+1} = extSuvo;
        end
        
        % Write headers
        for h = 1:length(headers)
            combinedData{currentRow, h} = headers{h};
        end
        currentRow = currentRow + 1;
        
        % Process each bin
        for bin = 1:numBins
            rowData = cell(1, length(headers));
            rowData{1} = bin * binSizeHours;  % Bin hours
            
            % Process each condition group
            for g = 1:length(conditionGroups)
                groupData = conditionGroups{g};
                groupValues = [];
                
                % Collect values for this group and metric
                for s = 1:length(groupData)
                    binField = sprintf('Bin%d', bin);
                    
                    if isfield(groupData(s), binField)
                        binData = groupData(s).(binField);
                        
                        % Special case for transition count
                        if strcmp(metricName, 'Transitions') && strcmp(stageName, 'Count')
                            if isfield(binData.Transitions, 'Count')
                                groupValues(end+1) = binData.Transitions.Count;
                            end
                        else
                            % Normal case for other metrics
                            if isfield(binData, metricName) && isfield(binData.(metricName), stageName)
                                groupValues(end+1) = binData.(metricName).(stageName);
                            end
                        end
                    end
                end
                
                % Add group data to row
                if ~isempty(groupValues)
                    rowData{g+1} = mean(groupValues);
                else
                    rowData{g+1} = NaN;
                end
            end
            
            % Write row to combined data
            for h = 1:length(rowData)
                combinedData{currentRow, h} = rowData{h};
            end
            currentRow = currentRow + 1;
        end
        
        % Add blank row between metrics
        currentRow = currentRow + 1;
    end
    
    % Write combined sheet
    try
        xlswrite(filePath, combinedData, combinedSheetName);
        disp('Created combined Prism-friendly sheet');
    catch e
        warning('Error writing combined sheet: %s', e.message);
        
        % Try alternative method
        try
            % Convert empty cells to NaN
            for r = 1:size(combinedData, 1)
                for c = 1:size(combinedData, 2)
                    if isempty(combinedData{r, c})
                        combinedData{r, c} = NaN;
                    end
                end
            end
            
            % Write using writecell
            writecell(combinedData, filePath, 'Sheet', combinedSheetName);
            disp('Created combined Prism-friendly sheet using alternative method');
        catch e2
            warning('Alternative method also failed: %s', e2.message);
        end
    end
end

function summaryTable = createTimeBinSummaryTableOnly(baselineVsVehTables, suvoVsVehTables, conditionNames, numBins)
    % Create a summary table without writing to Excel
    % This is used as a backup method
    %
    % Parameters:
    %   baselineVsVehTables - Cell array of comparison tables for Baseline vs Vehicle
    %   suvoVsVehTables - Cell array of comparison tables for Suvorexant vs Vehicle
    %   conditionNames - Cell array of condition names
    %   numBins - Number of time bins
    
    % Define key metrics for summary
    keyMetrics = {
        {'BoutLength', 'BoutLength_min', 'Wake', 'Wake bout length (min)'};
        {'BoutLength', 'BoutLength_min', 'SWS', 'SWS bout length (min)'};
        {'BoutLength', 'BoutLength_min', 'REM', 'REM bout length (min)'};
        {'PercentTime', 'PercentTime', 'Wake', 'Wake time %'};
        {'PercentTime', 'PercentTime', 'SWS', 'SWS time %'};
        {'PercentTime', 'PercentTime', 'REM', 'REM time %'};
        {'Bouts', 'Bouts', 'Wake', 'Wake bouts'};
        {'Bouts', 'Bouts', 'SWS', 'SWS bouts'};
        {'Bouts', 'Bouts', 'REM', 'REM bouts'};
        {'Transitions', 'Trans_Count', 'All', 'Transitions count'};
    };
    
    % Create headers for the summary table
    headers = {'Metric', 'Stage', 'Condition', 'Comparison'};
    
    % Add bin columns
    for bin = 1:numBins
        headers{end+1} = sprintf('Bin%d_Group1_Mean', bin);
        headers{end+1} = sprintf('Bin%d_Group2_Mean', bin);
        headers{end+1} = sprintf('Bin%d_Pct_Change', bin);
        headers{end+1} = sprintf('Bin%d_P_Value', bin);
    end
    
    % Create data matrix
    summaryData = {};
    
    % Process each condition
    for c = 1:length(conditionNames)
        % For each metric...
        for m = 1:length(keyMetrics)
            category = keyMetrics{m}{1};
            metric = keyMetrics{m}{2};
            stage = keyMetrics{m}{3};
            
            % Create rows for each comparison type
            blRow = cell(1, length(headers));
            blRow{1} = metric;
            blRow{2} = stage;
            blRow{3} = conditionNames{c};
            blRow{4} = 'Baseline_vs_Vehicle';
            
            suvoRow = cell(1, length(headers));
            suvoRow{1} = metric;
            suvoRow{2} = stage;
            suvoRow{3} = conditionNames{c};
            suvoRow{4} = 'Suvorexant_vs_Vehicle';
            
            % For each bin...
            for bin = 1:numBins
                % First check if data exists for this bin
                if ~isempty(baselineVsVehTables{c, bin}) && height(baselineVsVehTables{c, bin}) > 0
                    % Find the row with this metric
                    idx = find(strcmp(baselineVsVehTables{c, bin}.Category, category) & ...
                             strcmp(baselineVsVehTables{c, bin}.Metric, metric) & ...
                             strcmp(baselineVsVehTables{c, bin}.Stage, stage));
                    
                    if ~isempty(idx)
                        % Get column indices for this bin
                        colIdx = 4 + (bin-1)*4 + 1;
                        blRow{colIdx} = baselineVsVehTables{c, bin}.Baseline_Mean(idx);
                        blRow{colIdx+1} = baselineVsVehTables{c, bin}.Vehicle_Mean(idx);
                        blRow{colIdx+2} = baselineVsVehTables{c, bin}.Percent_Change(idx);
                        blRow{colIdx+3} = baselineVsVehTables{c, bin}.P_Value(idx);
                    end
                end
                
                % Do the same for Suvorexant vs Vehicle
                if ~isempty(suvoVsVehTables{c, bin}) && height(suvoVsVehTables{c, bin}) > 0
                    % Find the row with this metric
                    idx = find(strcmp(suvoVsVehTables{c, bin}.Category, category) & ...
                             strcmp(suvoVsVehTables{c, bin}.Metric, metric) & ...
                             strcmp(suvoVsVehTables{c, bin}.Stage, stage));
                    
                    if ~isempty(idx)
                        % Get column indices for this bin
                        colIdx = 4 + (bin-1)*4 + 1;
                        suvoRow{colIdx} = suvoVsVehTables{c, bin}.Suvorexant_Mean(idx);
                        suvoRow{colIdx+1} = suvoVsVehTables{c, bin}.Vehicle_Mean(idx);
                        suvoRow{colIdx+2} = suvoVsVehTables{c, bin}.Percent_Change(idx);
                        suvoRow{colIdx+3} = suvoVsVehTables{c, bin}.P_Value(idx);
                    end
                end
            end
            
            % Add rows to the summary data
            summaryData{end+1} = blRow;
            summaryData{end+1} = suvoRow;
        end
    end
    
    % Convert to table
    summaryTable = cell2table(vertcat(summaryData{:}), 'VariableNames', headers);
end

function validName = validateSheetName(sheetName)
    % Ensure sheet name is valid for Excel
    
    % Replace invalid characters
    invalid = {':', '\\', '/', '?', '*', '[', ']'};
    replacement = {'_', '_', '_', '_', '_', '(', ')'};
    
    validName = sheetName;
    for i = 1:length(invalid)
        validName = strrep(validName, invalid{i}, replacement{i});
    end
    
    % Truncate to 31 characters (Excel limit)
    if length(validName) > 31
        validName = validName(1:31);
    end
end
