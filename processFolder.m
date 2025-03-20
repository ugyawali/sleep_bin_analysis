function subjects = processFolder(folderPath, suvoSerials, binSizeHours, numBins)
    % Process all subject files in a folder and categorize by treatment
    % With added support for time bin analysis
    %
    % Parameters:
    %   folderPath - Path to folder containing Excel files
    %   suvoSerials - List of serial numbers assigned to Suvorexant group
    %   binSizeHours - Size of each time bin in hours (1, 3, or 6)
    %   numBins - Number of time bins to create (12 / binSizeHours)

    % Get all XLS/XLSX files in the folder
    files = dir(fullfile(folderPath, '*.xls*'));
    
    if isempty(files)
        warning('No Excel files found in folder: %s', folderPath);
        subjects = struct('Treatment', {}, 'SerialNumber', {});  % Initialize with fields
        return;
    end
    
    % Print debug info
    fprintf('Found %d files in folder: %s\n', length(files), folderPath);
    fprintf('Suvorexant serial numbers: %s\n', strjoin(suvoSerials, ', '));
    fprintf('Using time bins of %d hours (%d bins total)\n', binSizeHours, numBins);
    
    % Initialize array to store subject data
    subjects = struct('Treatment', {}, 'SerialNumber', {});  % Initialize with fields
    
    % Track assigned groups for reporting
    suvoCount = 0;
    vehCount = 0;
    
    % Process each file
    for i = 1:length(files)
        % Extract serial number from filename
        [~, filename, ~] = fileparts(files(i).name);
        serialNumber = extractSerialNumber(filename);
        
        if isempty(serialNumber)
            warning('Could not extract serial number from filename: %s', files(i).name);
            continue;
        end
        
        % Determine treatment group
        if ismember(serialNumber, suvoSerials)
            treatment = 'Suvorexant';
            suvoCount = suvoCount + 1;
        else
            treatment = 'Vehicle';
            vehCount = vehCount + 1;
        end
        
        % Read and process the data file
        filePath = fullfile(folderPath, files(i).name);
        try
            % Process the file with time bin support
            subjectData = processSubjectFileWithTimeBins(filePath, binSizeHours, numBins);
            
            % Add metadata
            subjectData.SerialNumber = serialNumber;
            subjectData.Treatment = treatment;
            subjectData.Filename = files(i).name;
            
            % Add to array of subjects
            if isempty(subjects)
                subjects = subjectData;
            else
                subjects(end+1) = subjectData;
            end
            
            fprintf('Successfully processed file: %s (SerialNumber: %s, Treatment: %s)\n', files(i).name, serialNumber, treatment);
            
        catch e
            warning('Error processing file %s: %s', files(i).name, e.message);
            fprintf('Stack trace: %s\n', getReport(e, 'extended'));
        end
    end
    
    % Print summary of group assignments
    fprintf('Processed %d files: %d assigned to Suvorexant, %d assigned to Vehicle\n', length(subjects), suvoCount, vehCount);
    
    % Check if we have any processed subjects
    if isempty(subjects) || numel(subjects) == 0
        warning('No files were successfully processed in folder: %s', folderPath);
        % Ensure structure array has the required fields even if empty
        subjects = struct('Treatment', {}, 'SerialNumber', {});
    end
end
