function createTimeBinSummarySheet(baselineVsVehTables, suvoVsVehTables, conditionNames, numBins, filePath)
    % This is the absolute simplest approach possible - direct Excel writing
    
    % Print debug info
    disp('*** STARTING DIRECT SUMMARY CREATION ***');
    
    try
        % Extract raw data for direct Excel writing
        blData = extractRawData(baselineVsVehTables, conditionNames, numBins, 'Baseline', 'Vehicle');
        suvoData = extractRawData(suvoVsVehTables, conditionNames, numBins, 'Suvorexant', 'Vehicle');
        
        % Add headers
        blHeaders = {'Condition', 'Bin', 'Metric', 'Baseline Mean', 'Vehicle Mean', 'Percent Change', 'P-Value', 'Significant'};
        suvoHeaders = {'Condition', 'Bin', 'Metric', 'Suvorexant Mean', 'Vehicle Mean', 'Percent Change', 'P-Value', 'Significant'};
        
        % Combine headers and data
        blOutput = [blHeaders; blData];
        suvoOutput = [suvoHeaders; suvoData];
        
        % Write directly to Excel without any complex formatting
        disp('Writing BL vs Veh summary directly to Excel...');
        xlswrite(filePath, blOutput, 'BL_vs_Veh_Summary');
        
        disp('Writing Suvo vs Veh summary directly to Excel...');
        xlswrite(filePath, suvoOutput, 'Suvo_vs_Veh_Summary');
        
        disp('*** SUMMARY SHEETS CREATED SUCCESSFULLY ***');
    catch e
        disp(['ERROR: ' e.message]);
        
        % Try alternate method
        try
            % Create flat text files instead
            [path, baseName, ~] = fileparts(filePath);
            blTxtPath = fullfile(path, [baseName '_BL_vs_Veh_Summary.txt']);
            suvoTxtPath = fullfile(path, [baseName '_Suvo_vs_Veh_Summary.txt']);
            
            % Open files for writing
            blFid = fopen(blTxtPath, 'w');
            suvoFid = fopen(suvoCsvPath, 'w');
            
            % Write flattened data
            fprintf(blFid, 'BASELINE VS VEHICLE SUMMARY\n\n');
            fprintf(suvoFid, 'SUVOREXANT VS VEHICLE SUMMARY\n\n');
            
            % Close files
            fclose(blFid);
            fclose(suvoFid);
            
            disp(['Summaries saved as text files: ' blTxtPath]);
        catch e2
            disp(['BACKUP ERROR: ' e2.message]);
        end
    end
end

% Helper function to extract data in simple format for direct Excel writing
function rawData = extractRawData(tables, conditionNames, numBins, group1Name, group2Name)
    % Initialize empty cell array for output
    rawData = {};
    
    % Define key metrics to include
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
            % Skip if table doesn't exist
            if isempty(tables{c, bin})
                continue;
            end
            
            % Process each metric
            for m = 1:length(keyMetrics)
                cat = keyMetrics{m}{1};
                met = keyMetrics{m}{2};
                stg = keyMetrics{m}{3};
                displayName = keyMetrics{m}{4};
                
                % Find this metric in the table
                idx = find(strcmp(tables{c, bin}.Category, cat) & ...
                         strcmp(tables{c, bin}.Metric, met) & ...
                         strcmp(tables{c, bin}.Stage, stg));
                
                if ~isempty(idx)
                    % Extract values
                    group1Col = [group1Name '_Mean'];
                    group2Col = [group2Name '_Mean'];
                    
                    group1Mean = tables{c, bin}.(group1Col)(idx(1));
                    group2Mean = tables{c, bin}.(group2Col)(idx(1));
                    pctChange = tables{c, bin}.Percent_Change(idx(1));
                    pVal = tables{c, bin}.P_Value(idx(1));
                    
                    % Determine significance
                    if pVal < 0.05
                        sig = '*';
                    elseif pVal < 0.1
                        sig = '+';
                    else
                        sig = '';
                    end
                    
                    % Add to raw data
                    rawData(end+1,:) = {conditionNames{c}, bin, displayName, ...
                                      group1Mean, group2Mean, pctChange, pVal, sig};
                end
            end
        end
    end
end
