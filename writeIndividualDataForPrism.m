function writeIndividualDataForPrism(baselineData, extDayData, conditionNames, filePath, binSizeHours, numBins)
    % Write individual subject data in a format suitable for GraphPad Prism
    %
    % Parameters:
    %   baselineData - Array of baseline subject data structures
    %   extDayData - Cell array of extinction day subject data
    %   conditionNames - Cell array of condition names
    %   filePath - Path to Excel file
    %   binSizeHours - Size of time bins in hours
    %   numBins - Number of time bins
    
    % Define important metrics to extract
    metrics = {{'BoutLength', 'Wake', 'Wake_Bout_Length_min'},
{'BoutLength', 'SWS', 'SWS_Bout_Length_min'},
{'BoutLength', 'REM', 'REM_Bout_Length_min'},
        {'PercentTime', 'Wake', 'Wake Time (%)'},
        {'PercentTime', 'SWS', 'SWS Time (%)'},
        {'PercentTime', 'REM', 'REM Time (%)'},
        {'Bouts', 'Wake', 'Wake Bouts'},
        {'Bouts', 'SWS', 'SWS Bouts'},
        {'Bouts', 'REM', 'REM Bouts'},
        {'Transitions', 'Count', 'Transition Count'}
    };
    
    % Process each metric
    for m = 1:length(metrics)
        metricName = metrics{m}{1};
        stageName = metrics{m}{2};
        displayName = metrics{m}{3};
        
        % Create sheet name (limited to 31 chars for Excel)
        sheetName = sprintf('Indiv_%s', displayName);
        sheetName = strrep(sheetName, ' ', '_');
        sheetName = strrep(sheetName, '(%)', 'Pct');
        sheetName = validateSheetName(sheetName);
        
        % Create headers based on conditions and bins
        headers = {'Subject', 'Treatment', 'Condition'};
        
        % Add bin columns
        for bin = 1:numBins
            headers{end+1} = sprintf('Bin%d', bin);
        end
        
        % Create data rows
        dataRows = {};
        
        % Process baseline data first
        for s = 1:length(baselineData)
            row = {baselineData(s).SerialNumber, baselineData(s).Treatment, 'Baseline'};
            
            % Add data for each bin
            for bin = 1:numBins
                binField = sprintf('Bin%d', bin);
                
                if isfield(baselineData(s), binField)
                    if strcmp(metricName, 'Transitions') && strcmp(stageName, 'Count')
                        % Special case for transition count
                        if isfield(baselineData(s).(binField).Transitions, 'Count')
                            row{end+1} = baselineData(s).(binField).Transitions.Count;
                        else
                            row{end+1} = NaN;
                        end
                    else
                        % Normal case for other metrics
                        if isfield(baselineData(s).(binField), metricName) && ...
                           isfield(baselineData(s).(binField).(metricName), stageName)
                            row{end+1} = baselineData(s).(binField).(metricName).(stageName);
                        else
                            row{end+1} = NaN;
                        end
                    end
                else
                    row{end+1} = NaN;
                end
            end
            
            dataRows{end+1} = row;
        end
        
        % Process extinction day data
        for c = 1:length(conditionNames)
            for s = 1:length(extDayData{c})
                row = {extDayData{c}(s).SerialNumber, extDayData{c}(s).Treatment, conditionNames{c}};
                
                % Add data for each bin
                for bin = 1:numBins
                    binField = sprintf('Bin%d', bin);
                    
                    if isfield(extDayData{c}(s), binField)
                        if strcmp(metricName, 'Transitions') && strcmp(stageName, 'Count')
                            % Special case for transition count
                            if isfield(extDayData{c}(s).(binField).Transitions, 'Count')
                                row{end+1} = extDayData{c}(s).(binField).Transitions.Count;
                            else
                                row{end+1} = NaN;
                            end
                        else
                            % Normal case for other metrics
                            if isfield(extDayData{c}(s).(binField), metricName) && ...
                               isfield(extDayData{c}(s).(binField).(metricName), stageName)
                                row{end+1} = extDayData{c}(s).(binField).(metricName).(stageName);
                            else
                                row{end+1} = NaN;
                            end
                        end
                    else
                        row{end+1} = NaN;
                    end
                end
                
                dataRows{end+1} = row;
            end
        end
        
        % Write to Excel
        try
            % Write headers
            writecell(headers, filePath, 'Sheet', sheetName, 'Range', 'A1');
            
            % Write data
            if ~isempty(dataRows)
                writecell(dataRows, filePath, 'Sheet', sheetName, 'Range', 'A2');
            end
            
            disp(['Created individual data sheet for ' displayName]);
        catch e
            warning('Error writing individual data for %s: %s', displayName, e.message);
        end
    end
    
    % Create an additional sheet with a more Prism-friendly format (grouped by treatment/condition)
    createPrismFormattedSheet(baselineData, extDayData, conditionNames, filePath, numBins);
end

function createPrismFormattedSheet(baselineData, extDayData, conditionNames, filePath, numBins)
    % Create a sheet with data formatted specifically for easy import into GraphPad Prism
    % This organizes data in columns by treatment/condition groups
    
    % Define important metrics
    metrics = {
        {'PercentTime', 'Wake', 'Wake_Time_Pct'},
        {'PercentTime', 'SWS', 'SWS_Time_Pct'},
        {'PercentTime', 'REM', 'REM_Time_Pct'},
        {'Bouts', 'Wake', 'Wake_Bouts'},
        {'Bouts', 'SWS', 'SWS_Bouts'},
        {'Bouts', 'REM', 'REM_Bouts'},
        {'Transitions', 'Count', 'Transition_Count'}
    };
    
    % Process each bin separately
    for bin = 1:numBins
        binField = sprintf('Bin%d', bin);
        sheetName = sprintf('Prism_Bin%d', bin);
        
        % For each metric, create a separate section in the sheet
        dataToWrite = {};
        currentRow = 1;
        
        for m = 1:length(metrics)
            metricName = metrics{m}{1};
            stageName = metrics{m}{2};
            displayName = metrics{m}{3};
            
            % Add a title for this metric
            dataToWrite{currentRow, 1} = displayName;
            currentRow = currentRow + 1;
            
            % Collect all unique treatment/condition combinations
            groups = {};
            
            % Add baseline groups
            % Add baseline as a single group (no treatment distinction)
if ~isempty(baselineData)
    groups{end+1} = {'Baseline', 'All'};
end
            
            % Add extinction day groups
            for c = 1:length(conditionNames)
                extVehicle = extDayData{c}(strcmp({extDayData{c}.Treatment}, 'Vehicle'));
                extSuvo = extDayData{c}(strcmp({extDayData{c}.Treatment}, 'Suvorexant'));
                
                if ~isempty(extVehicle)
                    groups{end+1} = {conditionNames{c}, 'Vehicle'};
                end
                
                if ~isempty(extSuvo)
                    groups{end+1} = {conditionNames{c}, 'Suvorexant'};
                end
            end
            
            % Create column headers
            for g = 1:length(groups)
                dataToWrite{currentRow, g} = sprintf('%s_%s', groups{g}{1}, groups{g}{2});
            end
            currentRow = currentRow + 1;
            
            % Find maximum number of subjects in any group
            maxSubjects = 0;
            for g = 1:length(groups)
                condition = groups{g}{1};
                treatment = groups{g}{2};
                
                if strcmp(condition, 'Baseline')
                    subjectData = baselineData(strcmp({baselineData.Treatment}, treatment));
                else
                    condIndex = find(strcmp(conditionNames, condition));
                    if ~isempty(condIndex)
                        subjectData = extDayData{condIndex}(strcmp({extDayData{condIndex}.Treatment}, treatment));
                    else
                        subjectData = [];
                    end
                end
                
                maxSubjects = max(maxSubjects, length(subjectData));
            end
            
            % Add data for each group in columns
            for subjectIdx = 1:maxSubjects
                for g = 1:length(groups)
                    condition = groups{g}{1};
                    treatment = groups{g}{2};
                    
                   % When retrieving subject data:
if strcmp(condition, 'Baseline')
    if strcmp(treatment, 'All')
        subjectData = baselineData; % Use all baseline subjects
    else
        subjectData = baselineData(strcmp({baselineData.Treatment}, treatment));
    end
else
    condIndex = find(strcmp(conditionNames, condition));
    if ~isempty(condIndex)
        subjectData = extDayData{condIndex}(strcmp({extDayData{condIndex}.Treatment}, treatment));
    else
        subjectData = [];
    end
end
                    
                    % Get value for this subject
                    if subjectIdx <= length(subjectData) && isfield(subjectData(subjectIdx), binField)
                        if strcmp(metricName, 'Transitions') && strcmp(stageName, 'Count')
                            % Special case for transition count
                            if isfield(subjectData(subjectIdx).(binField).Transitions, 'Count')
                                dataToWrite{currentRow + subjectIdx - 1, g} = subjectData(subjectIdx).(binField).Transitions.Count;
                            end
                        else
                            % Normal case for other metrics
                            if isfield(subjectData(subjectIdx).(binField), metricName) && ...
                               isfield(subjectData(subjectIdx).(binField).(metricName), stageName)
                                dataToWrite{currentRow + subjectIdx - 1, g} = subjectData(subjectIdx).(binField).(metricName).(stageName);
                            end
                        end
                    end
                end
            end
            
            % Move to next section (add a blank row)
            currentRow = currentRow + maxSubjects + 2;
        end
        
        % Write to Excel
        try
            % Handle empty cells
            for r = 1:size(dataToWrite, 1)
                for c = 1:size(dataToWrite, 2)
                    if isempty(dataToWrite{r, c})
                        dataToWrite{r, c} = NaN;
                    end
                end
            end
            
            writecell(dataToWrite, filePath, 'Sheet', sheetName);
            disp(['Created Prism-formatted data sheet for Bin ' num2str(bin)]);
        catch e
            warning('Error writing Prism-formatted data for Bin %d: %s', bin, e.message);
        end
    end
end