function writeTableToExcel(dataTable, filePath, sheetName)
    % Writes a table to Excel, handling various error conditions
    
    % If the file has a .csv extension, use writetable instead
    [~, ~, ext] = fileparts(filePath);
    if strcmpi(ext, '.csv')
        writetable(dataTable, filePath);
        disp(['Successfully wrote table to CSV file: ' filePath]);
        return;
    end
    
    try
        % First try standard writetable
        writetable(dataTable, filePath, 'Sheet', sheetName);
        disp(['Successfully wrote table to sheet: ' sheetName]);
    catch e1
        warning('Standard writetable failed for %s: %s', sheetName, e1.message);
        
        try
            % Try alternative approach: convert to cell array and write directly
            headers = dataTable.Properties.VariableNames;
            data = table2cell(dataTable);
            
            % Create sheet if it doesn't exist
            if ~isfile(filePath)
                xlswrite(filePath, {'Placeholder'}, sheetName);
            else
                try
                    % Check if sheet exists
                    [~, sheets] = xlsfinfo(filePath);
                    if ~any(strcmp(sheets, sheetName))
                        xlswrite(filePath, {'Placeholder'}, sheetName);
                    end
                catch
                    % If xlsfinfo fails, just try to create the sheet
                    try
                        xlswrite(filePath, {'Placeholder'}, sheetName);
                    catch
                        % Ignore errors here
                    end
                end
            end
            
            % Write headers
            xlswrite(filePath, headers, sheetName, 'A1');
            
            % Write data if not empty
            if ~isempty(data)
                xlswrite(filePath, data, sheetName, 'A2');
            end
            
            disp(['Successfully wrote data to sheet using alternative method: ' sheetName]);
        catch e2
            warning('All Excel writing methods failed for %s: %s', sheetName, e2.message);
        end
    end
end