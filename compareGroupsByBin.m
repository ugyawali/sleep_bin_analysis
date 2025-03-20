function resultTable = compareGroupsByBin(group1Data, group2Data, group1Name, group2Name, condition, binIndex)
    % Compare metrics between two groups for a specific time bin
    %
    % Parameters:
    %   group1Data - Array of subject data structures for first group
    %   group2Data - Array of subject data structures for second group
    %   group1Name - Name of first group (e.g., 'Baseline')
    %   group2Name - Name of second group (e.g., 'Vehicle')
    %   condition - Name of the condition being analyzed
    %   binIndex - Index of the time bin to compare
    
    if nargin < 6
        error('Bin index is required for compareGroupsByBin');
    end
    
    % Bin field name in the structure
    binField = sprintf('Bin%d', binIndex);
    
    % Check if data is empty
    if isempty(group1Data) || isempty(group2Data)
        warning('One or both groups are empty. Group1 (%s): %d subjects, Group2 (%s): %d subjects', ...
                group1Name, length(group1Data), group2Name, length(group2Data));
        
        % Create a minimal table with the correct column names
        resultTable = table('Size', [1, 9], ...
            'VariableTypes', {'string', 'string', 'string', 'double', 'double', 'double', 'double', 'double', 'double'}, ...
            'VariableNames', {'Category', 'Metric', 'Stage', ...
            [group1Name '_Mean'], [group1Name '_SEM'], ...
            [group2Name '_Mean'], [group2Name '_SEM'], ...
            'P_Value', 'Percent_Change'});
        
        % Fill with placeholder data to avoid errors
        resultTable.Category(1) = "Placeholder";
        resultTable.Metric(1) = "Placeholder";
        resultTable.Stage(1) = "Placeholder";
        resultTable.([group1Name '_Mean'])(1) = NaN;
        resultTable.([group1Name '_SEM'])(1) = NaN;
        resultTable.([group2Name '_Mean'])(1) = NaN;
        resultTable.([group2Name '_SEM'])(1) = NaN;
        resultTable.P_Value(1) = NaN;
        resultTable.Percent_Change(1) = NaN;
        
        return;
    end
    
    % Check if all subjects have data for this bin
    validGroup1 = true;
    validGroup2 = true;
    
    for i = 1:length(group1Data)
        if ~isfield(group1Data(i), binField)
            warning('Subject %d in group 1 missing data for %s', i, binField);
            validGroup1 = false;
            break;
        end
    end
    
    for i = 1:length(group2Data)
        if ~isfield(group2Data(i), binField)
            warning('Subject %d in group 2 missing data for %s', i, binField);
            validGroup2 = false;
            break;
        end
    end
    
    if ~validGroup1 || ~validGroup2
        warning('One or both groups missing data for bin %d', binIndex);
        
        % Create placeholder table
        resultTable = table('Size', [1, 9], ...
            'VariableTypes', {'string', 'string', 'string', 'double', 'double', 'double', 'double', 'double', 'double'}, ...
            'VariableNames', {'Category', 'Metric', 'Stage', ...
            [group1Name '_Mean'], [group1Name '_SEM'], ...
            [group2Name '_Mean'], [group2Name '_SEM'], ...
            'P_Value', 'Percent_Change'});
        
        % Fill with placeholder data
        resultTable.Category(1) = "Missing_Data";
        resultTable.Metric(1) = "Missing_Data";
        resultTable.Stage(1) = "Missing_Data";
        resultTable.([group1Name '_Mean'])(1) = NaN;
        resultTable.([group1Name '_SEM'])(1) = NaN;
        resultTable.([group2Name '_Mean'])(1) = NaN;
        resultTable.([group2Name '_SEM'])(1) = NaN;
        resultTable.P_Value(1) = NaN;
        resultTable.Percent_Change(1) = NaN;
        
        return;
    end
    
    % List of metrics to compare
    metrics = {'TotalTime', 'PercentTime', 'Bouts', 'BoutLength', 'Latencies', 'Latency2Min', 'Transitions'};
    stages = {'Wake', 'SWS', 'REM'};
    
    % Calculate number of rows needed
    totalRows = 0;
    for i = 1:length(metrics)
        if strcmp(metrics{i}, 'Transitions')
            totalRows = totalRows + 1; % Only one row for transitions
        else
            totalRows = totalRows + length(stages); % One row per stage
        end
    end
    
    % Initialize result arrays
    category = cell(totalRows, 1);
    metric = cell(totalRows, 1);
    stage = cell(totalRows, 1);
    group1Mean = zeros(totalRows, 1);
    group1SEM = zeros(totalRows, 1);
    group2Mean = zeros(totalRows, 1);
    group2SEM = zeros(totalRows, 1);
    pValue = zeros(totalRows, 1);
    percentChange = zeros(totalRows, 1);
    
    % Fill comparison data
    rowIdx = 1;
    for i = 1:length(metrics)
        metricName = metrics{i};
        
        if strcmp(metricName, 'Transitions')
            % Special case for transitions
            category{rowIdx} = 'Transitions';
            metric{rowIdx} = 'Trans_Count';
            stage{rowIdx} = 'All';
            
            % Extract values - safer approach
            group1Values = zeros(1, length(group1Data));
            for k = 1:length(group1Data)
                if isfield(group1Data(k).(binField).Transitions, 'Count')
                    group1Values(k) = group1Data(k).(binField).Transitions.Count;
                else
                    warning('Count field not found in Transitions for subject %d in group 1', k);
                    group1Values(k) = NaN;
                end
            end
            
            group2Values = zeros(1, length(group2Data));
            for k = 1:length(group2Data)
                if isfield(group2Data(k).(binField).Transitions, 'Count')
                    group2Values(k) = group2Data(k).(binField).Transitions.Count;
                else
                    warning('Count field not found in Transitions for subject %d in group 2', k);
                    group2Values(k) = NaN;
                end
            end
            
            % Remove NaN values
            group1Values = group1Values(~isnan(group1Values));
            group2Values = group2Values(~isnan(group2Values));
            
            % Calculate statistics
            if ~isempty(group1Values)
                group1Mean(rowIdx) = mean(group1Values);
                group1SEM(rowIdx) = std(group1Values) / sqrt(length(group1Values));
            else
                group1Mean(rowIdx) = NaN;
                group1SEM(rowIdx) = NaN;
            end
            
            if ~isempty(group2Values)
                group2Mean(rowIdx) = mean(group2Values);
                group2SEM(rowIdx) = std(group2Values) / sqrt(length(group2Values));
            else
                group2Mean(rowIdx) = NaN;
                group2SEM(rowIdx) = NaN;
            end
            
            % Perform t-test if enough data
            if length(group1Values) > 1 && length(group2Values) > 1
                [~, pValue(rowIdx)] = ttest2(group1Values, group2Values);
            else
                pValue(rowIdx) = NaN;
            end
            
            % Calculate percent change
            if ~isnan(group1Mean(rowIdx)) && ~isnan(group2Mean(rowIdx)) && group1Mean(rowIdx) ~= 0
                percentChange(rowIdx) = ((group2Mean(rowIdx) - group1Mean(rowIdx)) / group1Mean(rowIdx)) * 100;
            else
                percentChange(rowIdx) = NaN;
            end
            
            rowIdx = rowIdx + 1;
        else
            % Process each stage
            for j = 1:length(stages)
                stageName = stages{j};
                
                category{rowIdx} = metricName;
                
                % Handle the case where the metric field might not exist
                if isfield(group1Data(1).(binField), metricName) && isfield(group1Data(1).(binField).(metricName), 'Metric')
                    metric{rowIdx} = group1Data(1).(binField).(metricName).Metric;
                else
                    % Default metrics based on category
                    switch metricName
                        case 'TotalTime'
                            metric{rowIdx} = 'TotalTime_min';
                        case 'PercentTime'
                            metric{rowIdx} = 'PercentTime';
                        case 'Bouts'
                            metric{rowIdx} = 'Bouts';
                        case 'Latencies'
                            metric{rowIdx} = 'Latencies_min';
                        case 'Latency2Min'
                            metric{rowIdx} = 'Latency2Min_min';
                        otherwise
                            metric{rowIdx} = metricName;
                    end
                end
                
                stage{rowIdx} = stageName;
                
                % Extract values for this metric and stage - safer approach
                group1Values = zeros(1, length(group1Data));
                for k = 1:length(group1Data)
                    if isfield(group1Data(k).(binField), metricName) && isfield(group1Data(k).(binField).(metricName), stageName)
                        group1Values(k) = group1Data(k).(binField).(metricName).(stageName);
                    else
                        warning('Field %s or stage %s not found for subject %d in group 1, bin %d', metricName, stageName, k, binIndex);
                        group1Values(k) = NaN;
                    end
                end
                
                group2Values = zeros(1, length(group2Data));
                for k = 1:length(group2Data)
                    if isfield(group2Data(k).(binField), metricName) && isfield(group2Data(k).(binField).(metricName), stageName)
                        group2Values(k) = group2Data(k).(binField).(metricName).(stageName);
                    else
                        warning('Field %s or stage %s not found for subject %d in group 2, bin %d', metricName, stageName, k, binIndex);
                        group2Values(k) = NaN;
                    end
                end
                
                % Remove NaN values
                group1Values = group1Values(~isnan(group1Values));
                group2Values = group2Values(~isnan(group2Values));
                
                % Calculate statistics
                if ~isempty(group1Values)
                    group1Mean(rowIdx) = mean(group1Values);
                    group1SEM(rowIdx) = std(group1Values) / sqrt(length(group1Values));
                else
                    group1Mean(rowIdx) = NaN;
                    group1SEM(rowIdx) = NaN;
                end
                
                if ~isempty(group2Values)
                    group2Mean(rowIdx) = mean(group2Values);
                    group2SEM(rowIdx) = std(group2Values) / sqrt(length(group2Values));
                else
                    group2Mean(rowIdx) = NaN;
                    group2SEM(rowIdx) = NaN;
                end
                
                % Perform t-test if enough data
                if length(group1Values) > 1 && length(group2Values) > 1
                    [~, pValue(rowIdx)] = ttest2(group1Values, group2Values);
                else
                    pValue(rowIdx) = NaN;
                end
                
                % Calculate percent change
                if ~isnan(group1Mean(rowIdx)) && ~isnan(group2Mean(rowIdx)) && group1Mean(rowIdx) ~= 0
                    percentChange(rowIdx) = ((group2Mean(rowIdx) - group1Mean(rowIdx)) / group1Mean(rowIdx)) * 100;
                else
                    percentChange(rowIdx) = NaN;
                end
                
                rowIdx = rowIdx + 1;
            end
        end
    end
    
    % Create the results table with consistent column names
    resultTable = table(category, metric, stage, ...
        group1Mean, group1SEM, group2Mean, group2SEM, ...
        pValue, percentChange, ...
        'VariableNames', {
            'Category', 'Metric', 'Stage', ...
            [group1Name '_Mean'], [group1Name '_SEM'], ...
            [group2Name '_Mean'], [group2Name '_SEM'], ...
            'P_Value', 'Percent_Change'
        });
    
    % Add group counts and bin info
    resultTable.Properties.UserData.Group1Count = length(group1Data);
    resultTable.Properties.UserData.Group2Count = length(group2Data);
    resultTable.Properties.UserData.Group1Name = group1Name;
    resultTable.Properties.UserData.Group2Name = group2Name;
    resultTable.Properties.UserData.Condition = condition;
    resultTable.Properties.UserData.BinIndex = binIndex;
end
    
   