function saveTimeBinnedResultsToExcel(baselineVsVehTables, suvoVsVehTables, conditionNames, filePath, baselineData, extDayData, binSizeHours, numBins)
    % Save time-binned results to Excel file with multiple sheets
    %
    % Parameters:
    %   baselineVsVehTables - Cell array of comparison tables for Baseline vs Vehicle
    %   suvoVsVehTables - Cell array of comparison tables for Suvorexant vs Vehicle
    %   conditionNames - Cell array of condition names
    %   filePath - Full path to Excel file to create
    %   baselineData - Array of baseline subject data structures
    %   extDayData - Cell array of extinction day subject data structures
    %   binSizeHours - Size of time bins in hours
    %   numBins - Number of time bins
    
    % Single try-catch block for all Excel operations
    try
        % First check if file exists and delete it if it does
        if exist(filePath, 'file')
            delete(filePath);
        end
        
        % Create a more detailed README with actual condition names and bin numbers
        readme = {'This Excel file contains time-binned sleep analysis data from a cocaine CPP study with suvorexant treatment.'};
        readme = [readme; {''}];
        readme = [readme; {sprintf('TIME BINS: Data analyzed in %d-hour bins (%d bins total)', binSizeHours, numBins)}];
        readme = [readme; {''}];
        readme = [readme; {'SHEETS:'}];
        readme = [readme; {'- README: This information'}];
        readme = [readme; {'- Summary: Key findings across all time bins'}];
        readme = [readme; {'- BL_vs_Veh_Summary: Effect of cocaine across all time bins'}];
        readme = [readme; {'- Suvo_vs_Veh_Summary: Effect of suvorexant across all time bins'}];
        readme = [readme; {''}];

        % Add condition and bin specific descriptions
        for i = 1:length(conditionNames)
            for bin = 1:numBins
                readme = [readme; {sprintf('- %s_Bin%d_BL_vs_Veh: Effect of cocaine on %s during bin %d', conditionNames{i}, bin, conditionNames{i}, bin)}];
                readme = [readme; {sprintf('- %s_Bin%d_Suvo_vs_Veh: Effect of suvorexant on %s during bin %d', conditionNames{i}, bin, conditionNames{i}, bin)}];
            end
        end

        % Add information about individual data sheets
        readme = [readme; {''}];
        for bin = 1:numBins
            readme = [readme; {sprintf('- Individual_Baseline_Bin%d: Individual subject data for baseline in bin %d', bin, bin)}];
            for i = 1:length(conditionNames)
                readme = [readme; {sprintf('- Individual_%s_Bin%d: Individual data for %s in bin %d', conditionNames{i}, bin, conditionNames{i}, bin)}];
            end
        end

        % Add information about transition sheets
        readme = [readme; {''}];
        for bin = 1:numBins
            readme = [readme; {sprintf('- Transitions_Baseline_Bin%d: Stage transitions for baseline in bin %d', bin, bin)}];
            for i = 1:length(conditionNames)
                readme = [readme; {sprintf('- Transitions_%s_Bin%d: Stage transitions for %s in bin %d', conditionNames{i}, bin, conditionNames{i}, bin)}];
            end
        end

        % Add information about Prism-formatted sheets
        readme = [readme; {''}];
        for bin = 1:numBins
            readme = [readme; {sprintf('- Prism_Bin%d: Data formatted for GraphPad Prism for bin %d', bin, bin)}];
        end

        % Add metrics information
        readme = [readme; {''}];
        readme = [readme; {'METRICS:'}];
        readme = [readme; {'- TotalTime_min: Total time spent in each sleep stage (minutes)'}];
        readme = [readme; {'- PercentTime: Percentage of recording time in each stage'}];
        readme = [readme; {'- Bouts: Number of episodes of each stage'}];
        readme = [readme; {'- BoutLength_min: Average duration of bouts for each stage (minutes)'}];
        readme = [readme; {'- Trans_Count: Number of transitions between stages'}];
        readme = [readme; {'- Latencies_min: Time to first occurrence of each stage (minutes)'}];
        readme = [readme; {'- Latency2Min_min: Time to first 2-minute bout of each stage (minutes)'}];
        readme = [readme; {''}];
        readme = [readme; {sprintf('GENERATED: %s', string(datetime('now')))}];
        
        % Write the README sheet
        xlswrite(filePath, readme, 'README');
        disp('Created README sheet.');
        
        % Create a time bin summary sheet
        disp('Creating time bin summary sheet...');
        createTimeBinSummarySheet(baselineVsVehTables, suvoVsVehTables, conditionNames, numBins, filePath);
        
        % Add individual data sheets for Prism
        disp('Creating individual data sheets for Prism...');
        writeIndividualDataForPrism(baselineData, extDayData, conditionNames, filePath, binSizeHours, numBins);
        
        % For each bin, write detailed comparison tables
        for bin = 1:numBins
            % Write individual subject data for baseline in this bin
            writeIndividualDataByBin(baselineData, 'Baseline', filePath, bin);
            
            % Write detailed transition data for baseline in this bin
            writeDetailedTransitionsByBin(baselineData, 'Baseline', filePath, bin);
            
            % For each condition, write individual subject data and detailed transitions
            for i = 1:length(conditionNames)
                writeIndividualDataByBin(extDayData{i}, conditionNames{i}, filePath, bin);
                writeDetailedTransitionsByBin(extDayData{i}, conditionNames{i}, filePath, bin);
                
                % Baseline vs Vehicle tables for this condition and bin
                % Create valid sheet name (max 31 chars)
                sheetName = sprintf('%s_Bin%d_BL_vs_Veh', conditionNames{i}, bin);
                sheetName = strrep(sheetName, ' ', '_'); % Replace spaces
                sheetName = validateSheetName(sheetName);
                
                disp(['Writing sheet: ' sheetName]);
                % Make sure the table is not empty
                if ~isempty(baselineVsVehTables{i, bin})
                    % Try multiple methods to write the table
                    try
                        % Method 1: writetable
                        writetable(baselineVsVehTables{i, bin}, filePath, 'Sheet', sheetName);
                    catch e1
                        disp(['Error with method 1: ' e1.message]);
                        try
                            % Method 2: xlswrite directly
                            headers = baselineVsVehTables{i, bin}.Properties.VariableNames;
                            data = table2cell(baselineVsVehTables{i, bin});
                            
                            % Write headers and data
                            xlswrite(filePath, headers, sheetName, 'A1');
                            xlswrite(filePath, data, sheetName, 'A2');
                        catch e2
                            disp(['Error with method 2: ' e2.message]);
                            % Additional fallback options would go here
                        end
                    end
                else
                    warning('Baseline vs Vehicle table for %s, Bin %d is empty', conditionNames{i}, bin);
                end
                
                % Suvorexant vs Vehicle tables for this condition and bin
                sheetName = sprintf('%s_Bin%d_Suvo_vs_Veh', conditionNames{i}, bin);
                sheetName = strrep(sheetName, ' ', '_'); % Replace spaces
                sheetName = validateSheetName(sheetName);
                
                disp(['Writing sheet: ' sheetName]);
                % Make sure the table is not empty
                if ~isempty(suvoVsVehTables{i, bin})
                    % Try multiple methods to write the table
                    try
                        % Method 1: writetable
                        writetable(suvoVsVehTables{i, bin}, filePath, 'Sheet', sheetName);
                    catch e1
                        disp(['Error with method 1: ' e1.message]);
                        try
                            % Method 2: xlswrite directly
                            headers = suvoVsVehTables{i, bin}.Properties.VariableNames;
                            data = table2cell(suvoVsVehTables{i, bin});
                            
                            % Write headers and data
                            xlswrite(filePath, headers, sheetName, 'A1');
                            xlswrite(filePath, data, sheetName, 'A2');
                        catch e2
                            disp(['Error with method 2: ' e2.message]);
                            % Additional fallback options would go here
                        end
                    end
                else
                    warning('Suvorexant vs Vehicle table for %s, Bin %d is empty', conditionNames{i}, bin);
                end
            end
        end
        
        disp('Excel file writing completed successfully!');
        
    catch mainError
        warning('Error saving Excel file: %s', mainError.message);
        disp('Trying alternative method to save data as separate CSV files...');
        
        % Alternative method: Save as separate CSV files
        [pathName, nameOnly, ~] = fileparts(filePath);
        baseFileName = fullfile(pathName, nameOnly);
        
        % Save README
        writecell(readme, [baseFileName '_README.csv']);
        
        % Save Summary (create it first)
        summaryTable = createTimeBinSummaryTableOnly(baselineVsVehTables, suvoVsVehTables, conditionNames, numBins);
        writetable(summaryTable, [baseFileName '_Summary.csv']);
        
        % For each bin, save separate tables
        for bin = 1:numBins
            % Save individual subject data for baseline in this bin
            writeIndividualDataByBin(baselineData, 'Baseline', [baseFileName sprintf('_Individual_Baseline_Bin%d.csv', bin)], bin, true);
            
            % For each condition, save separate tables
            for i = 1:length(conditionNames)
                % Save individual data
                writeIndividualDataByBin(extDayData{i}, conditionNames{i}, [baseFileName sprintf('_Individual_%s_Bin%d.csv', conditionNames{i}, bin)], bin, true);
                
                % Save detailed transitions
                writeDetailedTransitionsByBin(extDayData{i}, conditionNames{i}, [baseFileName sprintf('_Transitions_%s_Bin%d.csv', conditionNames{i}, bin)], bin, true);
                
                % Save comparison tables
                baseName = sprintf('%s_Bin%d_BL_vs_Veh.csv', conditionNames{i}, bin);
                if ~isempty(baselineVsVehTables{i, bin})
                    writetable(baselineVsVehTables{i, bin}, [baseFileName '_' baseName]);
                end
                
                suvoName = sprintf('%s_Bin%d_Suvo_vs_Veh.csv', conditionNames{i}, bin);
                if ~isempty(suvoVsVehTables{i, bin})
                    writetable(suvoVsVehTables{i, bin}, [baseFileName '_' suvoName]);
                end
            end
        end
        disp(['Data saved as separate CSV files in ' pathName]);
    end
end