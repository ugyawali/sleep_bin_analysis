function subjectData = processSubjectFileWithTimeBins(filePath, binSizeHours, numBins)
    % Process a single subject data file with support for time bins
    %
    % Parameters:
    %   filePath - Path to Excel file with sleep data
    %   binSizeHours - Size of each time bin in hours
    %   numBins - Number of time bins to create
    %
    % Returns:
    %   subjectData - Structure with processed metrics for each time bin

    % Read the Excel file
    try
        [~, ~, ext] = fileparts(filePath);
        if strcmpi(ext, '.xlsx')
            data = readtable(filePath, 'Sheet', 1);
        else
            % For older .xls files, sometimes need different approach
            [~, ~, raw] = xlsread(filePath);
            headers = raw(1,:);
            data = cell2table(raw(2:end,:), 'VariableNames', headers);
        end
    catch e
        warning('Error reading file %s: %s. Trying alternative method...', filePath, e.message);
        try
            [num, txt, raw] = xlsread(filePath);
            % Create table from raw data
            headers = raw(1,:);
            headers = cellfun(@(x) strrep(strrep(x, ' ', '_'), '.', ''), headers, 'UniformOutput', false);
            numRows = size(raw, 1);
            
            % Prepare data for table
            tableData = cell(numRows-1, length(headers));
            for col = 1:length(headers)
                for row = 2:numRows
                    tableData{row-1, col} = raw{row, col};
                end
            end
            
            data = cell2table(tableData, 'VariableNames', headers);
        catch e2
            error('Cannot read file %s: %s', filePath, e2.message);
        end
    end
    
    % Initialize structure for results with necessary fields
    subjectData = struct('Treatment', '', 'SerialNumber', '');
    
    % Find the stage column
    stageCol = findStageColumn(data);
    fprintf('Processing file: %s\n', filePath);
    fprintf('Found stage column at index: %d\n', stageCol);
    
    % Calculate total recording time in minutes (assuming 10s bins)
    recordingTime = size(data, 1) * 10 / 60; % Convert 10s bins to minutes
    subjectData.RecordingTime_min = recordingTime;
    
    % Prepare stage data
    if istable(data)
        stageData = data{:, stageCol};
        if isnumeric(stageData)
            % Convert numeric stages to strings
            stageMap = containers.Map({1, 2, 3}, {'Wake', 'SWS', 'REM'});
            stageDataCell = cell(size(stageData));
            for i = 1:length(stageData)
                if isKey(stageMap, stageData(i))
                    stageDataCell{i} = stageMap(stageData(i));
                else
                    stageDataCell{i} = 'Unknown';
                end
            end
            stageData = stageDataCell;
        elseif ~iscell(stageData)
            stageData = cellstr(stageData);
        end
    else
        error('Cannot process data: not a valid table');
    end
    
    % Define bin boundaries in rows (assuming 10s per row)
    binSizeRows = binSizeHours * 60 * 6; % 6 rows per minute, 60 minutes per hour
    
    % Process each time bin
    for bin = 1:numBins
        binStart = (bin-1) * binSizeRows + 1;
        binEnd = min(bin * binSizeRows, length(stageData));
        
        % Handle case where recording is shorter than expected
        if binStart > length(stageData)
            warning('Bin %d (start: %d) exceeds recording length (%d). Skipping bin.', bin, binStart, length(stageData));
            continue;
        end
        
        % Get data for this bin
        binStageData = stageData(binStart:binEnd);
        
        % Calculate bin metrics (adjusting for 10s bins)
        binMetrics = calculateBinMetrics(binStageData, binSizeHours);
        
        % Store bin metrics in subject data
        fieldName = sprintf('Bin%d', bin);
        subjectData.(fieldName) = binMetrics;
    end
    
    % Also calculate overall metrics across the whole recording
    overallMetrics = calculateBinMetrics(stageData, recordingTime / 60);
    subjectData.Overall = overallMetrics;
    
    return;
end
