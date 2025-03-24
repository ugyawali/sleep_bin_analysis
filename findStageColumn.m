function colIdx = findStageColumn(data)
    % Find the column that contains stage information
    
    varNames = data.Properties.VariableNames;
    fprintf('Available column names: %s\n', strjoin(varNames, ', '));
    
    % Look for columns with specific names
    possibleNames = {'Stage', 'State', 'SleepStage', 'Sleep_Stage', 'Sleep', 'Score'};
    
    for i = 1:length(possibleNames)
        idx = find(strcmpi(varNames, possibleNames{i}));
        if ~isempty(idx)
            fprintf('Found stage column "%s" at index %d\n', varNames{idx}, idx);
            colIdx = idx;
            return;
        end
    end
    
    % If we haven't found it by name, try to find a column with expected values
    stages = {'Wake', 'SWS', 'REM', 'NREM', 'W', 'N', 'R'};
    
    for i = 1:width(data)
        % Get the column data
        columnData = data{:, i};
        
        % If it's a cell array, check for stage names
        if iscell(columnData)
            uniqueVals = unique(columnData);
            if iscell(uniqueVals)
                fprintf('Checking column %d (%s) - unique values: %s\n', ...
                    i, varNames{i}, strjoin(uniqueVals, ', '));
                
                matches = 0;
                for j = 1:length(stages)
                    if any(strcmpi(uniqueVals, stages{j}))
                        matches = matches + 1;
                    end
                end
                
                if matches >= 2  % At least 2 stage types found
                    fprintf('Found likely stage column at %d based on matching %d stage labels\n', i, matches);
                    colIdx = i;
                    return;
                end
            end
            
        % If it's numeric, check for values in the range of stage codes
        elseif isnumeric(columnData)
            uniqueVals = unique(columnData);
            if length(uniqueVals) <= 5 && all(uniqueVals >= 0 & uniqueVals <= 5)
                fprintf('Found likely numeric stage column at %d with values: %s\n', ...
                    i, num2str(uniqueVals'));
                colIdx = i;
                return;
            end
        end
    end
    
    % If we still can't find it, use the first column
    warning('Could not identify stage column. Using column 1 as default.');
    colIdx = 1;
end
