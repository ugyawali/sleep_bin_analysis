function displayTable(table, keyMetrics, group1Name, group2Name)
    % Display selected metrics from a comparison table
    
    % Check if the table is empty or doesn't have the right columns
    if isempty(table) || height(table) == 0
        fprintf('Table is empty. No data to display.\n');
        return;
    end
    
    % Get variable names from the table
    varNames = table.Properties.VariableNames;
    fprintf('Variable names in table: %s\n', strjoin(varNames, ', '));
    
    % Check if required columns exist
    group1Col = [group1Name '_Mean'];
    group2Col = [group2Name '_Mean'];
    
    if ~ismember(group1Col, varNames)
        fprintf('Error: Column "%s" not found in table\n', group1Col);
        return;
    end
    
    if ~ismember(group2Col, varNames)
        fprintf('Error: Column "%s" not found in table\n', group2Col);
        return;
    end
    
    % Print header
    fprintf('%-25s | %-12s | %-12s | %-10s | %-10s\n', 'Metric', group1Name, group2Name, 'Change %', 'P-value');
    fprintf('-------------------------------------------------------------------------\n');
    
    % Print each metric
    for i = 1:length(keyMetrics)
        % Find the row for this metric
        idx = find(strcmp(table.Category, keyMetrics{i}{1}) & ...
                 strcmp(table.Metric, keyMetrics{i}{2}) & ...
                 strcmp(table.Stage, keyMetrics{i}{3}));
        
        if ~isempty(idx)
            % Get the values
            group1Val = table{idx, group1Col};
            group2Val = table{idx, group2Col};
            
            if ismember('Percent_Change', varNames)
                pctChange = table{idx, 'Percent_Change'};
            else
                pctChange = NaN;
            end
            
            if ismember('P_Value', varNames)
                pVal = table{idx, 'P_Value'};
            else
                pVal = NaN;
            end
            
            % Format the p-value display
            if isnan(pVal)
                pValStr = 'N/A';
            elseif pVal < 0.05
                pValStr = sprintf('%.3f *', pVal);
            elseif pVal < 0.1
                pValStr = sprintf('%.3f +', pVal);
            else
                pValStr = sprintf('%.3f', pVal);
            end
            
            % Print the row
            fprintf('%-25s | %12.2f | %12.2f | %+10.2f | %10s\n', ...
                keyMetrics{i}{4}, group1Val, group2Val, pctChange, pValStr);
        else
            fprintf('%-25s | %12s | %12s | %10s | %10s\n', ...
                keyMetrics{i}{4}, 'N/A', 'N/A', 'N/A', 'N/A');
        end
    end
    
    fprintf('* p < 0.05, + p < 0.10\n');
end
