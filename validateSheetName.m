function validName = validateSheetName(sheetName)
    % Ensure sheet name is valid for Excel
    
    % Replace invalid characters
    invalid = {':', '\\', '/', '?', '*', '[', ']'};
    replacement = {'_', '_', '_', '_', '_', '(', ')'};
    
    validName = sheetName;
    for i = 1:length(invalid)
        validName = strrep(validName, invalid{i}, replacement{i});
    end
    
    % Truncate to 31 characters (Excel limit)
    if length(validName) > 31
        validName = validName(1:31);
    end
end
