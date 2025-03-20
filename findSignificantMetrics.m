function sigFindings = findSignificantMetrics(table, keyMetrics, alpha)
    % Find all metrics in the table that are statistically significant
    
    sigFindings = {};
    
    for i = 1:length(keyMetrics)
        % Find the row for this metric
        idx = find(strcmp(table.Category, keyMetrics{i}{1}) & ...
                 strcmp(table.Metric, keyMetrics{i}{2}) & ...
                 strcmp(table.Stage, keyMetrics{i}{3}));
        
        if ~isempty(idx) && table.P_Value(idx) <= alpha
            sigFindings{end+1} = struct(...
                'MetricName', keyMetrics{i}{4}, ...
                'PValue', table.P_Value(idx), ...
                'PercentChange', table.Percent_Change(idx));
        end
    end
end