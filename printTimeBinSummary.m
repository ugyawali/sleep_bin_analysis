function printTimeBinSummary(baselineVsVehTables, suvoVsVehTables, conditionNames, numBins)
    % Print a summary of significant findings across all time bins
    
    fprintf('\n===== SIGNIFICANT FINDINGS SUMMARY =====\n');
    
    % Define key metrics to check
    keyMetrics = {{'BoutLength', 'Wake', 'Wake_Bout_Length_min'},
{'BoutLength', 'SWS', 'SWS_Bout_Length_min'},
{'BoutLength', 'REM', 'REM_Bout_Length_min'},
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