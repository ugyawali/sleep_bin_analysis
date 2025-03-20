function writeIndividualDataByBin(subjectData, conditionName, filePath, binIndex, csvMode)
    % Write individual subject data to a separate sheet for a specific time bin
    %
    % Parameters:
    %   subjectData - Array of subject data structures
    %   conditionName - Name of the condition (e.g., 'Baseline', 'Ext_D1')
    %   filePath - Path to Excel file
    %   binIndex - Index of the time bin
    %   csvMode - Optional flag, if true writes to CSV instead of Excel sheet
    
    if nargin < 5
        csvMode = false;
    end
    
    % Bin field name in the structure
    binField = sprintf('Bin%d', binIndex);
    
    % Check if all subjects have data for this bin
    validSubjects = 0;
    for i = 1:length(subjectData)
        if isfield(subjectData(i), binField)
            validSubjects = validSubjects + 1;
        end
    end
    
    if validSubjects == 0 || isempty(subjectData)
        warning('No valid subjects with data for bin %d', binIndex);
        return;
    end
    
    % Initialize arrays for table columns (use cell arrays for more reliable Excel writing)
    serialNumbers = cell(length(subjectData), 1);
    treatments = cell(length(subjectData), 1);
    recordingTime = cell(length(subjectData), 1);
    
    wakePercent = cell(length(subjectData), 1);
    swsPercent = cell(length(subjectData), 1);
    remPercent = cell(length(subjectData), 1);
    
    wakeTotalTime = cell(length(subjectData), 1);
    swsTotalTime = cell(length(subjectData), 1);
    remTotalTime = cell(length(subjectData), 1);
    
    wakeBouts = cell(length(subjectData), 1);
    swsBouts = cell(length(subjectData), 1);
    remBouts = cell(length(subjectData), 1);
    
    wakeBoutLength = cell(length(subjectData), 1);
    swsBoutLength = cell(length(subjectData), 1);
    remBoutLength = cell(length(subjectData), 1);
    
    wakeLatency = cell(length(subjectData), 1);
    swsLatency = cell(length(subjectData), 1);
    remLatency = cell(length(subjectData), 1);
    
    wakeLat2Min = cell(length(subjectData), 1);
    swsLat2Min = cell(length(subjectData), 1);
    remLat2Min = cell(length(subjectData), 1);
    
    transCount = cell(length(subjectData), 1);
    
    % Extract data from each subject for the specific bin
    for i = 1:length(subjectData)
        serialNumbers{i} = subjectData(i).SerialNumber;
        treatments{i} = subjectData(i).Treatment;
        
        if ~isfield(subjectData(i), binField)
            continue;
        end
        
        binData = subjectData(i).(binField);
        
        % Recording time
        if isfield(binData, 'DurationMin')
            recordingTime{i} = binData.DurationMin;
        end
        
        % Extract PercentTime values
        if isfield(binData, 'PercentTime')
            if isfield(binData.PercentTime, 'Wake')
                wakePercent{i} = binData.PercentTime.Wake;
            end
            
            if isfield(binData.PercentTime, 'SWS')
                swsPercent{i} = binData.PercentTime.SWS;
            end
            
            if isfield(binData.PercentTime, 'REM')
                remPercent{i} = binData.PercentTime.REM;
            end
        end
        
        % Extract TotalTime values
        if isfield(binData, 'TotalTime')
            if isfield(binData.TotalTime, 'Wake')
                wakeTotalTime{i} = binData.TotalTime.Wake;
            end
            
            if isfield(binData.TotalTime, 'SWS')
                swsTotalTime{i} = binData.TotalTime.SWS;
            end
            
            if isfield(binData.TotalTime, 'REM')
                remTotalTime{i} = binData.TotalTime.REM;
            end
        end
        
        % Extract Bouts values
        if isfield(binData, 'Bouts')
            if isfield(binData.Bouts, 'Wake')
                wakeBouts{i} = binData.Bouts.Wake;
            end
            
            if isfield(binData.Bouts, 'SWS')
                swsBouts{i} = binData.Bouts.SWS;
            end
            
            if isfield(binData.Bouts, 'REM')
                remBouts{i} = binData.Bouts.REM;
            end
        end
        
        % Extract BoutLength values
        if isfield(binData, 'BoutLength')
            if isfield(binData.BoutLength, 'Wake')
                wakeBoutLength{i} = binData.BoutLength.Wake;
            end
            
            if isfield(binData.BoutLength, 'SWS')
                swsBoutLength{i} = binData.BoutLength.SWS;
            end
            
            if isfield(binData.BoutLength, 'REM')
                remBoutLength{i} = binData.BoutLength.REM;
            end
        end
        
        % Extract Latencies values
        if isfield(binData, 'Latencies')
            if isfield(binData.Latencies, 'Wake')
                wakeLatency{i} = binData.Latencies.Wake;
            end
            
            if isfield(binData.Latencies, 'SWS')
                swsLatency{i} = binData.Latencies.SWS;
            end
            
            if isfield(binData.Latencies, 'REM')
                remLatency{i} = binData.Latencies.REM;
            end
        end
        
        % Extract Latency2Min values
        if isfield(binData, 'Latency2Min')
            if isfield(binData.Latency2Min, 'Wake')
                wakeLat2Min{i} = binData.Latency2Min.Wake;
            end
            
            if isfield(binData.Latency2Min, 'SWS')
                swsLat2Min{i} = binData.Latency2Min.SWS;
            end
            
            if isfield(binData.Latency2Min, 'REM')
                remLat2Min{i} = binData.Latency2Min.REM;
            end
        end
        
        % Extract Transitions count
        if isfield(binData, 'Transitions') && isfield(binData.Transitions, 'Count')
            transCount{i} = binData.Transitions.Count;
        end
    end
    
    % Create a raw data cell array with all the data
    rawData = [serialNumbers, treatments, recordingTime, ...
               wakePercent, swsPercent, remPercent, ...
               wakeTotalTime, swsTotalTime, remTotalTime, ...
               wakeBouts, swsBouts, remBouts, ...
               wakeBoutLength, swsBoutLength, remBoutLength, ...
               wakeLatency, swsLatency, remLatency, ...
               wakeLat2Min, swsLat2Min, remLat2Min, ...
               transCount];
    
    % Create headers
    headers = {'SerialNumber', 'Treatment', 'RecordingTime_min', ...
              'WakePercent', 'SWSPercent', 'REMPercent', ...
              'WakeTotalTime_min', 'SWSTotalTime_min', 'REMTotalTime_min', ...
              'WakeBouts', 'SWSBouts', 'REMBouts', ...
              'WakeBoutLength_min', 'SWSBoutLength_min', 'REMBoutLength_min', ...
              'WakeLatency_min', 'SWSLatency_min', 'REMLatency_min', ...
              'WakeLatency2Min_min', 'SWSLatency2Min_min', 'REMLatency2Min_min', ...
              'TransitionCount'};
    
    % Write to file using direct methods
    try
        if csvMode
            % Create CSV
            fid = fopen(filePath, 'w');
            
            % Write headers
            fprintf(fid, '%s,', headers{1:end-1});
            fprintf(fid, '%s\n', headers{end});
            
            % Write data rows
            for row = 1:size(rawData, 1)
                for col = 1:size(rawData, 2)-1
                    if isnumeric(rawData{row, col})
                        fprintf(fid, '%.6f,', rawData{row, col});
                    else
                        fprintf(fid, '%s,', char(rawData{row, col}));
                    end
                end
                
                % Last column
                if isnumeric(rawData{row, end})
                    fprintf(fid, '%.6f\n', rawData{row, end});
                else
                    fprintf(fid, '%s\n', char(rawData{row, end}));
                end
            end
            
            fclose(fid);
            disp(['Wrote individual data for ' conditionName ', Bin ' num2str(binIndex) ' to CSV file']);
        else
            % Write to Excel using xlswrite
            sheetName = sprintf('Individual_%s_Bin%d', conditionName, binIndex);
            sheetName = validateSheetName(sheetName);
            
            % Make sure file exists
            if ~isfile(filePath)
                xlswrite(filePath, {'Placeholder'}, sheetName);
            end
            
            % Write headers
            xlswrite(filePath, headers, sheetName, 'A1');
            
            % Write data
            xlswrite(filePath, rawData, sheetName, 'A2');
            
            disp(['Wrote individual data for ' conditionName ', Bin ' num2str(binIndex) ' to Excel']);
        end
    catch e
        warning('Error writing individual data for %s, Bin %d: %s', conditionName, binIndex, e.message);
        
        % Try writing to CSV as fallback
        if ~csvMode
            try
                csvFilePath = strrep(filePath, '.xlsx', sprintf('_Individual_%s_Bin%d.csv', conditionName, binIndex));
                writeIndividualDataByBin(subjectData, conditionName, csvFilePath, binIndex, true);
            catch
                warning('Both Excel and CSV writing methods failed for individual data.');
            end
        end
    end
end