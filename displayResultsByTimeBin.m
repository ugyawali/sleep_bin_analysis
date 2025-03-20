function displayResultsByTimeBin(baselineVsVehTables, suvoVsVehTables, conditionNames, numBins)
    % Display results to console in a formatted way, organized by time bin
    
    % Define key metrics to display
    keyMetrics = {{'BoutLength', 'BoutLength_min', 'Wake', 'Wake bout length (min)'};
{'BoutLength', 'BoutLength_min', 'SWS', 'SWS bout length (min)'};
{'BoutLength', 'BoutLength_min', 'REM', 'REM bout length (min)'};
        {'PercentTime', 'PercentTime', 'Wake', 'Wake time %'};
        {'PercentTime', 'PercentTime', 'SWS', 'SWS time %'};
        {'PercentTime', 'PercentTime', 'REM', 'REM time %'};
        {'Latencies', 'Latencies_min', 'SWS', 'SWS latency (min)'};
        {'Latencies', 'Latencies_min', 'REM', 'REM latency (min)'};
        {'Transitions', 'Trans_Count', 'All', 'Transitions count'};
    };
    
    % Display header
    fprintf('\n===== TIME-BINNED SLEEP ANALYSIS RESULTS =====\n\n');
    
    % Display results for each time bin
    for bin = 1:numBins
        fprintf('\n===== TIME BIN %d =====\n\n', bin);
        
        % Display key metrics for each condition within this bin
        for i = 1:length(conditionNames)
            fprintf('--- %s (Bin %d) ---\n\n', conditionNames{i}, bin);
            
            % Check if tables exist for this bin
            if isempty(baselineVsVehTables{i, bin}) || height(baselineVsVehTables{i, bin}) == 0
                fprintf('No data available for Baseline vs Vehicle comparison\n\n');
            else
                % Display Baseline vs Vehicle comparisons
                fprintf('Baseline vs Vehicle (Effect of Cocaine):\n');
                try
                    displayTable(baselineVsVehTables{i, bin}, keyMetrics, 'Baseline', 'Vehicle');
                catch e
                    fprintf('Error displaying table: %s\n', e.message);
                    fprintf('Table variable names: %s\n', strjoin(baselineVsVehTables{i, bin}.Properties.VariableNames, ', '));
                end
            end
            
            if isempty(suvoVsVehTables{i, bin}) || height(suvoVsVehTables{i, bin}) == 0
                fprintf('No data available for Suvorexant vs Vehicle comparison\n\n');
            else
                % Display Suvorexant vs Vehicle comparisons
                fprintf('\nSuvorexant vs Vehicle (Effect of Treatment):\n');
                try
                    displayTable(suvoVsVehTables{i, bin}, keyMetrics, 'Suvorexant', 'Vehicle');
                catch e
                    fprintf('Error displaying table: %s\n', e.message);
                    fprintf('Table variable names: %s\n', strjoin(suvoVsVehTables{i, bin}.Properties.VariableNames, ', '));
                end
            end
            
            fprintf('\n');
        end
    end
    
    % Print a summary of significant findings across all bins
    printTimeBinSummary(baselineVsVehTables, suvoVsVehTables, conditionNames, numBins);
end

function printTimeBinSummary(baselineVsVehTables, suvoVsVehTables, conditionNames, numBins)
    % Print a summary of significant findings across all time bins
    
    fprintf('\n===== SIGNIFICANT FINDINGS SUMMARY =====\n');
    
    % Define key metrics to check
    keyMetrics = {
        {'PercentTime', 'PercentTime', 'Wake', 'Wake time %'};
        {'PercentTime', 'PercentTime', 'SWS', 'SWS time %'};
        {'PercentTime', 'PercentTime', 'REM', 'REM time %'};
        {'Bouts', 'Bouts', 'Wake', 'Wake bouts'};
        {'Bouts', 'Bouts', 'SWS', 'SWS bouts'};
        {'Bouts', 'Bouts', 'REM', 'REM bouts'};
        {'Transitions', 'Trans_Count', 'All', 'Transitions count'};
    };
    
    % 1. Check Baseline vs Vehicle (Effect of Cocaine)
    fprintf('\n*** Effects of Cocaine (Baseline vs Vehicle) ***\n');
    
    for i = 1:length(conditionNames)
        fprintf('\n- %s:\n', conditionNames{i});
        hasSigFindings = false;
        
        for bin = 1:numBins
            if ~isempty(baselineVsVehTables{i, bin}) && height(baselineVsVehTables{i, bin}) > 0
                % Check if any metrics are significant
                sigFindings = findSignificantMetrics(baselineVsVehTables{i, bin}, keyMetrics, 0.05);
                
                if ~isempty(sigFindings)
                    hasSigFindings = true;
                    fprintf('  Time Bin %d:\n', bin);
                    for s = 1:length(sigFindings)
                        fprintf('    - %s (p=%.3f, change: %.1f%%)\n', ...
                            sigFindings{s}.MetricName, ...
                            sigFindings{s}.PValue, ...
                            sigFindings{s}.PercentChange);
                    end
                end
            end
        end
        
        if ~hasSigFindings
            fprintf('  No significant findings\n');
        end
    end
    
    % 2. Check Suvorexant vs Vehicle (Effect of Treatment)
    fprintf('\n*** Effects of Suvorexant (Suvorexant vs Vehicle) ***\n');
    
    for i = 1:length(conditionNames)
        fprintf('\n- %s:\n', conditionNames{i});
        hasSigFindings = false;
        
        for bin = 1:numBins
            if ~isempty(suvoVsVehTables{i, bin}) && height(suvoVsVehTables{i, bin}) > 0
                % Check if any metrics are significant
                sigFindings = findSignificantMetrics(suvoVsVehTables{i, bin}, keyMetrics, 0.05);
                
                if ~isempty(sigFindings)
                    hasSigFindings = true;
                    fprintf('  Time Bin %d:\n', bin);
                    for s = 1:length(sigFindings)
                        fprintf('    - %s (p=%.3f, change: %.1f%%)\n', ...
                            sigFindings{s}.MetricName, ...
                            sigFindings{s}.PValue, ...
                            sigFindings{s}.PercentChange);
                    end
                end
            end
        end
        
        if ~hasSigFindings
            fprintf('  No significant findings\n');
        end
    end
end