function serialNumber = extractSerialNumber(filename)
    % Extract serial number from filename
    % Examples: "CEO Baseline 164a Scoring.xlsx", "JR D5 1676 Scoring.xlsx"
    
    % Look for patterns like 164a, 1676, etc.
    pattern = '\d{3,4}[a-zA-Z]?';
    match = regexp(filename, pattern, 'match');
    
    if ~isempty(match)
        serialNumber = match{1};
        fprintf('Extracted serial number "%s" from filename "%s"\n', serialNumber, filename);
    else
        warning('Failed to extract serial number from filename: %s', filename);
        serialNumber = '';
    end
end
