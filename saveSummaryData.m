function saveSummaryData(baselineVsVehTables, suvoVsVehTables, conditionNames, numBins, filePath)
    % New separate function to save summary data
    % This should be called at the end of analyzeSleepDataByTimeBins
    % after all other processing is complete
    
    disp('Starting separate summary data export...');
    
    % Get the directory path for output files
    [folderPath, fileName, ~] = fileparts(filePath);
    
    % Create output file paths
    blSummaryPath = fullfile(folderPath, [fileName '_BL_vs_Veh_Summary.xlsx']);
    suvoSummaryPath = fullfile(folderPath, [fileName '_Suvo_vs_Veh_Summary.xlsx']);
    
    % Define key metrics to include in summaries
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
    
    % ------ EXPORT BASELINE VS VEHICLE SUMMARY ------
    
    try
        % Collect all data into a matrix
        blData = cell(1, 8); % Headers
        blData{1, 1} = 'Condition';
        blData{1, 2} = 'Bin';
        blData{1, 3} = 'Metric';
        blData{1, 4} = 'Baseline Mean';
        blData{1, 5} = 'Vehicle Mean';
        blData{1, 6} = 'Percent Change';
        blData{1, 7} = 'P-Value';
        blData{1, 8} = 'Significant';
        
        % Track significant findings
        significantFindings = {};
        
        % Current row index
        row = 2;
        
        % Process each condition, bin, and metric
        for c = 1:length(conditionNames)
            for bin = 1:numBins
                % Only process if the table exists and has data
                if ~isempty(baselineVsVehTables{c, bin}) && height(baselineVsVehTables{c, bin}) > 0
                    for m = 1:length(keyMetrics)
                        % Find the row for this metric
                        idx = find(strcmp(baselineVsVehTables{c, bin}.Category, keyMetrics{m}{1}) & ...
                                 strcmp(baselineVsVehTables{c, bin}.Metric, keyMetrics{m}{2}) & ...
                                 strcmp(baselineVsVehTables{c, bin}.Stage, keyMetrics{m}{3}));
                        
                        if ~isempty(idx)
                            % Extract values
                            blMean = baselineVsVehTables{c, bin}.Baseline_Mean(idx(1));
                            vehMean = baselineVsVehTables{c, bin}.Vehicle_Mean(idx(1));
                            pctChange = baselineVsVehTables{c, bin}.Percent_Change(idx(1));
                            pVal = baselineVsVehTables{c, bin}.P_Value(idx(1));
                            
                            % Determine significance
                            if pVal < 0.05
                                sig = '*';
                                significantFindings{end+1} = sprintf('%s, Bin %d: %s (p=%.3f, change: %.1f%%)', ...
                                    conditionNames{c}, bin, keyMetrics{m}{4}, pVal, pctChange);
                            elseif pVal < 0.1
                                sig = '+';
                            else
                                sig = '';
                            end
                            
                            % Add to data matrix
                            blData{row, 1} = conditionNames{c};
                            blData{row, 2} = bin;
                            blData{row, 3} = keyMetrics{m}{4};
                            blData{row, 4} = blMean;
                            blData{row, 5} = vehMean;
                            blData{row, 6} = pctChange;
                            blData{row, 7} = pVal;
                            blData{row, 8} = sig;
                            
                            row = row + 1;
                        end
                    end
                end
            end
        end
        
        % Add significant findings section
        blData{row, 1} = '';
        row = row + 1;
        blData{row, 1} = 'SIGNIFICANT FINDINGS (p < 0.05):';
        row = row + 1;
        
        if isempty(significantFindings)
            blData{row, 1} = 'No significant findings.';
            row = row + 1;
        else
            for i = 1:length(significantFindings)
                blData{row, 1} = significantFindings{i};
                row = row + 1;
            end
        end
        
        % Save using writematrix (available in MATLAB 2022b)
        writematrix(blData, blSummaryPath);
        disp(['Baseline vs Vehicle summary saved to: ' blSummaryPath]);
        
    catch e
        disp(['Error saving BL vs Veh summary: ' e.message]);
        
        % Try CSV fallback
        try
            csvPath = fullfile(folderPath, [fileName '_BL_vs_Veh_Summary.csv']);
            fid = fopen(csvPath, 'w');
            
            % Write headers
            fprintf(fid, 'Condition,Bin,Metric,Baseline Mean,Vehicle Mean,Percent Change,P-Value,Significant\n');
            
            % Write data (skipping the first row which is headers)
            for i = 2:(row-1)
                if i <= size(blData, 1)
                    % Handle different data types
                    if ischar(blData{i, 1}) || isstring(blData{i, 1})
                        fprintf(fid, '%s,', blData{i, 1});
                    else
                        fprintf(fid, ',');
                    end
                    
                    if isnumeric(blData{i, 2})
                        fprintf(fid, '%d,', blData{i, 2});
                    else
                        fprintf(fid, ',');
                    end
                    
                    if ischar(blData{i, 3}) || isstring(blData{i, 3})
                        fprintf(fid, '"%s",', blData{i, 3});
                    else
                        fprintf(fid, ',');
                    end
                    
                    % Numeric columns
                    for j = 4:7
                        if isnumeric(blData{i, j})
                            fprintf(fid, '%.6f,', blData{i, j});
                        else
                            fprintf(fid, ',');
                        end
                    }
                    
                    % Last column (significance)
                    if ischar(blData{i, 8}) || isstring(blData{i, 8})
                        fprintf(fid, '%s\n', blData{i, 8});
                    else
                        fprintf(fid, '\n');
                    end
                end
            end
            
            fclose(fid);
            disp(['Baseline vs Vehicle summary saved as CSV: ' csvPath]);
        catch e2
            disp(['CSV fallback also failed: ' e2.message]);
        end
    end
    
    % ------ EXPORT SUVOREXANT VS VEHICLE SUMMARY ------
    
    try
        % Collect all data into a matrix
        suvoData = cell(1, 8); % Headers
        suvoData{1, 1} = 'Condition';
        suvoData{1, 2} = 'Bin';
        suvoData{1, 3} = 'Metric';
        suvoData{1, 4} = 'Suvorexant Mean';
        suvoData{1, 5} = 'Vehicle Mean';
        suvoData{1, 6} = 'Percent Change';
        suvoData{1, 7} = 'P-Value';
        suvoData{1, 8} = 'Significant';
        
        % Track significant findings
        significantFindings = {};
        
        % Current row index
        row = 2;
        
        % Process each condition, bin, and metric
        for c = 1:length(conditionNames)
            for bin = 1:numBins
                % Only process if the table exists and has data
                if ~isempty(suvoVsVehTables{c, bin}) && height(suvoVsVehTables{c, bin}) > 0
                    for m = 1:length(keyMetrics)
                        % Find the row for this metric
                        idx = find(strcmp(suvoVsVehTables{c, bin}.Category, keyMetrics{m}{1}) & ...
                                 strcmp(suvoVsVehTables{c, bin}.Metric, keyMetrics{m}{2}) & ...
                                 strcmp(suvoVsVehTables{c, bin}.Stage, keyMetrics{m}{3}));
                        
                        if ~isempty(idx)
                            % Extract values
                            suvoMean = suvoVsVehTables{c, bin}.Suvorexant_Mean(idx(1));
                            vehMean = suvoVsVehTables{c, bin}.Vehicle_Mean(idx(1));
                            pctChange = suvoVsVehTables{c, bin}.Percent_Change(idx(1));
                            pVal = suvoVsVehTables{c, bin}.P_Value(idx(1));
                            
                            % Determine significance
                            if pVal < 0.05
                                sig = '*';
                                significantFindings{end+1} = sprintf('%s, Bin %d: %s (p=%.3f, change: %.1f%%)', ...
                                    conditionNames{c}, bin, keyMetrics{m}{4}, pVal, pctChange);
                            elseif pVal < 0.1
                                sig = '+';
                            else
                                sig = '';
                            end
                            
                            % Add to data matrix
                            suvoData{row, 1} = conditionNames{c};
                            suvoData{row, 2} = bin;
                            suvoData{row, 3} = keyMetrics{m}{4};
                            suvoData{row, 4} = suvoMean;
                            suvoData{row, 5} = vehMean;
                            suvoData{row, 6} = pctChange;
                            suvoData{row, 7} = pVal;
                            suvoData{row, 8} = sig;
                            
                            row = row + 1;
                        end
                    end
                end
            end
        end
        
        % Add significant findings section
        suvoData{row, 1} = '';
        row = row + 1;
        suvoData{row, 1} = 'SIGNIFICANT FINDINGS (p < 0.05):';
        row = row + 1;
        
        if isempty(significantFindings)
            suvoData{row, 1} = 'No significant findings.';
            row = row + 1;
        else
            for i = 1:length(significantFindings)
                suvoData{row, 1} = significantFindings{i};
                row = row + 1;
            end
        end
        
        % Save using writematrix (available in MATLAB 2022b)
        writematrix(suvoData, suvoSummaryPath);
        disp(['Suvorexant vs Vehicle summary saved to: ' suvoSummaryPath]);
        
    catch e
        disp(['Error saving Suvo vs Veh summary: ' e.message]);
        
        % Try CSV fallback
        try
            csvPath = fullfile(folderPath, [fileName '_Suvo_vs_Veh_Summary.csv']);
            fid = fopen(csvPath, 'w');
            
            % Write headers
            fprintf(fid, 'Condition,Bin,Metric,Suvorexant Mean,Vehicle Mean,Percent Change,P-Value,Significant\n');
            
            % Write data (skipping the first row which is headers)
            for i = 2:(row-1)
                if i <= size(suvoData, 1)
                    % Handle different data types
                    if ischar(suvoData{i, 1}) || isstring(suvoData{i, 1})
                        fprintf(fid, '%s,', suvoData{i, 1});
                    else
                        fprintf(fid, ',');
                    end
                    
                    if isnumeric(suvoData{i, 2})
                        fprintf(fid, '%d,', suvoData{i, 2});
                    else
                        fprintf(fid, ',');
                    end
                    
                    if ischar(suvoData{i, 3}) || isstring(suvoData{i, 3})
                        fprintf(fid, '"%s",', suvoData{i, 3});
                    else
                        fprintf(fid, ',');
                    end
                    
                    % Numeric columns
                    for j = 4:7
                        if isnumeric(suvoData{i, j})
                            fprintf(fid, '%.6f,', suvoData{i, j});
                        else
                            fprintf(fid, ',');
                        end
                    end
                    
                    % Last column (significance)
                    if ischar(suvoData{i, 8}) || isstring(suvoData{i, 8})
                        fprintf(fid, '%s\n', suvoData{i, 8});
                    else
                        fprintf(fid, '\n');
                    end
                end
            end
            
            fclose(fid);
            disp(['Suvorexant vs Vehicle summary saved as CSV: ' csvPath]);
        catch e2
            disp(['CSV fallback also failed: ' e2.message]);
        end
    end
    
    % Add references to the main Excel file
    try
        % Create a simple message with file locations
        message = {
            'SUMMARY DATA SAVED TO SEPARATE FILES:', 
            '', 
            ['Baseline vs Vehicle: ' blSummaryPath], 
            '', 
            ['Suvorexant vs Vehicle: ' suvoSummaryPath]
        };
        
        % Write this message to the Excel file
        xlswrite(filePath, message, 'Summary');
        disp('Added references to main Excel file');
    catch
        disp('Could not add references to main Excel file');
    end
    
    disp('Summary data export complete.');
end