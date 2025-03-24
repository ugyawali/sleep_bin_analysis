function writeDetailedTransitionsByBin(subjectData, conditionName, filePath, binIndex, csvMode)
    % Write detailed transition counts for a specific time bin
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
    
    % Define the possible transitions
    fromStages = {'Wake', 'SWS', 'REM'};
    toStages = {'Wake', 'SWS', 'REM'};
    
    % Initialize columns for the table (use cell arrays for more reliable Excel writing)
    serialNumbers = cell(length(subjectData), 1);
    treatments = cell(length(subjectData), 1);
    
    % Create empty matrix for transition counts
    transMatrix = cell(length(subjectData), length(fromStages) * length(toStages));
    
    % Create column names for the transitions
    transColNames = cell(1, length(fromStages) * length(toStages));
    colIdx = 1;
    for i = 1:length(fromStages)
        for j = 1:length(toStages)
            transColNames{colIdx} = [fromStages{i} '_to_' toStages{j}];
            colIdx = colIdx + 1;
        end
    end
    
    % Extract the actual transition counts for each subject from DetailedTransitions
    for s = 1:length(subjectData)
        serialNumbers{s} = subjectData(s).SerialNumber;
        treatments{s} = subjectData(s).Treatment;
        
        if ~isfield(subjectData(s), binField)
            continue;
        end
        
        binData = subjectData(s).(binField);
        
        if isfield(binData, 'DetailedTransitions') && isfield(binData.DetailedTransitions, 'Counts')
            % Get the transition matrix for this subject
            subjTransMatrix = binData.DetailedTransitions.Counts;
            
            % Convert to row vector
            colIdx = 1;
            for i = 1:length(fromStages)
                for j = 1:length(toStages)
                    transMatrix{s, colIdx} = subjTransMatrix(i, j);
                    colIdx = colIdx + 1;
                end
            end
        else
            warning('Detailed transitions not available for subject %s in bin %d', serialNumbers{s}, binIndex);
        end
    end
    
    % Create raw data array
    rawData = [serialNumbers, treatments];
    
    % Add transition data
    rawData = [rawData, transMatrix];
    
    % Create headers
    headers = {'SerialNumber', 'Treatment'};
    headers = [headers, transColNames];
    
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
                        fprintf(fid, '%.2f,', rawData{row, col});
                    else
                        fprintf(fid, '%s,', char(rawData{row, col}));
                    end
                end
                
                % Last column
                if isnumeric(rawData{row, end})
                    fprintf(fid, '%.2f\n', rawData{row, end});
                else
                    fprintf(fid, '%s\n', char(rawData{row, end}));
                end
            end
            
            fclose(fid);
            disp(['Wrote transition data for ' conditionName ', Bin ' num2str(binIndex) ' to CSV file']);
        else
            % Write to Excel using xlswrite
            sheetName = sprintf('Transitions_%s_Bin%d', conditionName, binIndex);
            sheetName = validateSheetName(sheetName);
            
            % Make sure file exists
            if ~isfile(filePath)
                xlswrite(filePath, {'Placeholder'}, sheetName);
            end
            
            % Write headers
            xlswrite(filePath, headers, sheetName, 'A1');
            
            % Write data
            xlswrite(filePath, rawData, sheetName, 'A2');
            
            disp(['Wrote transition data for ' conditionName ', Bin ' num2str(binIndex) ' to Excel']);
        end
    catch e
        warning('Error writing transition data for %s, Bin %d: %s', conditionName, binIndex, e.message);
        
        % Try writing to CSV as fallback
        if ~csvMode
            try
                csvFilePath = strrep(filePath, '.xlsx', sprintf('_Transitions_%s_Bin%d.csv', conditionName, binIndex));
                writeDetailedTransitionsByBin(subjectData, conditionName, csvFilePath, binIndex, true);
            catch
                warning('Both Excel and CSV writing methods failed for transitions data.');
            end
        end
    end
end
