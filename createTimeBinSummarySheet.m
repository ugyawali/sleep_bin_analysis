function createTimeBinSummarySheet(baselineVsVehTables, suvoVsVehTables, conditionNames, numBins, filePath)
    % Create a simplified summary sheet that will work reliably
    
    % Create the baseline vs vehicle summary
    blSummary = {'Condition', 'Bin', 'Metric', 'Baseline Mean', 'Vehicle Mean', 'Percent Change', 'P-Value', 'Significant'};
    blRows = {};
    
    % Create the suvorexant vs vehicle summary
    suvoSummary = {'Condition', 'Bin', 'Metric', 'Suvorexant Mean', 'Vehicle Mean', 'Percent Change', 'P-Value', 'Significant'};
    suvoRows = {};
    
    % Define key metrics to display
    keyMetrics = {{'BoutLength', 'BoutLength_min', 'Wake', 'Wake bout length (min)'};
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
        writeTableToExcel(blTable, filePath, 'Sheet', 'BL_vs_Veh_Summary', 'WriteMode', 'append');
    end
    
    writecell(suvoSummary, filePath, 'Sheet', 'Suvo_vs_Veh_Summary', 'Range', 'A1');
    if ~isempty(suvoRows)
        suvoTable = cell2table(suvoRows, 'VariableNames', suvoSummary);
        writeTableToExcel(suvoTable, filePath, 'Sheet', 'Suvo_vs_Veh_Summary', 'WriteMode', 'append');
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
end