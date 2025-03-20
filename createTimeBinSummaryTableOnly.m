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
    keyMetrics = {{'BoutLength', 'BoutLength_min', 'Wake', 'Wake bout length (min)'};
{'BoutLength', 'BoutLength_min', 'SWS', 'SWS bout length (min)'};
{'BoutLength', 'BoutLength_min', 'REM', 'REM bout length (min)'};
        {'PercentTime', 'PercentTime', 'Wake'};
        {'PercentTime', 'PercentTime', 'SWS'};
        {'PercentTime', 'PercentTime', 'REM'};
        {'Bouts', 'Bouts', 'Wake'};
        {'Bouts', 'Bouts', 'SWS'};
        {'Bouts', 'Bouts', 'REM'};
        {'Transitions', 'Trans_Count', 'All'};
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