function binMetrics = calculateBinMetrics(stageData, binDurationHours)
    % Calculate sleep metrics for a given time bin
    %
    % Parameters:
    %   stageData - Cell array of stage labels for this bin
    %   binDurationHours - Duration of the bin in hours
    %
    % Returns:
    %   binMetrics - Structure with all calculated metrics
    
    % Define stages
    stages = {'Wake', 'SWS', 'REM'};
    
    % Initialize metrics structure
    binMetrics = struct();
    
    % Calculate bin duration in minutes
    binDurationMin = binDurationHours * 60;
    binMetrics.DurationMin = binDurationMin;
    
    % Initialize metrics
    totalTime = zeros(length(stages), 1);
    boutCounts = zeros(length(stages), 1);
    latencies = inf(length(stages), 1);
    latencies2Min = inf(length(stages), 1);
    
    % Analyze each stage
    for s = 1:length(stages)
        stage = stages{s};
        % Handle different possible labels for the same stage
        if strcmpi(stage, 'Wake')
            stageIndices = strcmpi(stageData, 'Wake') | strcmpi(stageData, 'W');
        elseif strcmpi(stage, 'SWS')
            stageIndices = strcmpi(stageData, 'SWS') | strcmpi(stageData, 'NREM') | strcmpi(stageData, 'N');
        elseif strcmpi(stage, 'REM')
            stageIndices = strcmpi(stageData, 'REM') | strcmpi(stageData, 'R');
        else
            stageIndices = strcmpi(stageData, stage);
        end
        
        % Calculate total time (adjusting for 10s bins)
        totalTime(s) = sum(stageIndices) * 10 / 60; % Convert from 10s bins to minutes
        
        % Find first occurrence for latency (adjusting for 10s bins)
        firstIdx = find(stageIndices, 1);
        if ~isempty(firstIdx)
            latencies(s) = firstIdx * 10 / 60; % Convert from 10s bins to minutes
        end
        
        % Find all bout start/end times
        boutStarts = [];
        boutEnds = [];
        boutLengths = [];
        
        % Find transitions from off→on (bout starts) and on→off (bout ends)
        startTransitions = find(diff([0; stageIndices]) == 1);
        endTransitions = find(diff([stageIndices; 0]) == -1);
        
        if ~isempty(startTransitions)
            boutStarts = startTransitions;
            boutEnds = endTransitions;
            boutLengths = (boutEnds - boutStarts + 1) * 10 / 60; % Convert from 10s bins to minutes
            boutCounts(s) = length(boutStarts);
            
            % Find first bout that's at least 2 minutes long
            longBoutIdx = find(boutLengths >= 2, 1);
            if ~isempty(longBoutIdx)
                latencies2Min(s) = boutStarts(longBoutIdx) * 10 / 60; % Convert from 10s bins to minutes
            end
        end
    end
   
% Calculate mean bout lengths for each stage separately
boutLengths = zeros(length(stages), 1);
for s = 1:length(stages)
    stage = stages{s};
    
    % Handle different possible labels for the same stage
    if strcmpi(stage, 'Wake')
        stageIndices = strcmpi(stageData, 'Wake') | strcmpi(stageData, 'W');
    elseif strcmpi(stage, 'SWS')
        stageIndices = strcmpi(stageData, 'SWS') | strcmpi(stageData, 'NREM') | strcmpi(stageData, 'N');
    elseif strcmpi(stage, 'REM')
        stageIndices = strcmpi(stageData, 'REM') | strcmpi(stageData, 'R');
    else
        stageIndices = strcmpi(stageData, stage);
    end
    
    % Find transitions for this specific stage
    % Find transitions from off→on (bout starts) and on→off (bout ends)
    stageBoutStarts = find(diff([0; stageIndices]) == 1);
    stageBoutEnds = find(diff([stageIndices; 0]) == -1);
    
    if ~isempty(stageBoutStarts)
        % Calculate bout lengths for this specific stage
        stageBoutLengths = (stageBoutEnds - stageBoutStarts + 1) * 10 / 60; % Convert from 10s bins to minutes
        boutLengths(s) = mean(stageBoutLengths);
    else
        boutLengths(s) = NaN;
    end
end



    % Calculate detailed stage transitions
    prevStage = '';
    transCount = 0;
    
    % Initialize transition matrix for detailed analysis
    % Wake=1, SWS=2, REM=3
    transMatrix = zeros(3, 3);
    
    for i = 1:length(stageData)
        currStage = stageData{i};
        
        % Skip unknown stages like 'X'
        if ~any(strcmpi(currStage, {'Wake', 'W', 'SWS', 'NREM', 'N', 'REM', 'R'}))
            continue;
        end
        
        % Map to standard stage names
        if strcmpi(currStage, 'W')
            currStage = 'Wake';
        elseif any(strcmpi(currStage, {'NREM', 'N'}))
            currStage = 'SWS';
        elseif strcmpi(currStage, 'R')
            currStage = 'REM';
        end
        
        % Count transitions
        if ~isempty(prevStage) && ~strcmp(currStage, prevStage)
            transCount = transCount + 1;
            
            % Record detailed transition
            fromIdx = find(strcmpi(stages, prevStage));
            toIdx = find(strcmpi(stages, currStage));
            
            if ~isempty(fromIdx) && ~isempty(toIdx)
                transMatrix(fromIdx, toIdx) = transMatrix(fromIdx, toIdx) + 1;
            end
        end
        
        prevStage = currStage;
    end
    
    % Store results in structure - TotalTime
    binMetrics.TotalTime = struct();
    binMetrics.TotalTime.Category = 'TotalTime';
    binMetrics.TotalTime.Metric = 'TotalTime_min';
    for s = 1:length(stages)
        binMetrics.TotalTime.(stages{s}) = totalTime(s);
    end
    
    % PercentTime
    binMetrics.PercentTime = struct();
    binMetrics.PercentTime.Category = 'PercentTime';
    binMetrics.PercentTime.Metric = 'PercentTime';
    for s = 1:length(stages)
        binMetrics.PercentTime.(stages{s}) = (totalTime(s) / binDurationMin) * 100;
    end
    
    % Bouts
    binMetrics.Bouts = struct();
    binMetrics.Bouts.Category = 'Bouts';
    binMetrics.Bouts.Metric = 'Bouts';
    for s = 1:length(stages)
        binMetrics.Bouts.(stages{s}) = boutCounts(s);
    end
    % Store in results structure
binMetrics.BoutLength = struct();
binMetrics.BoutLength.Category = 'BoutLength';
binMetrics.BoutLength.Metric = 'BoutLength_min';
for s = 1:length(stages)
    binMetrics.BoutLength.(stages{s}) = boutLengths(s);
end
    % Latencies
    binMetrics.Latencies = struct();
    binMetrics.Latencies.Category = 'Latencies';
    binMetrics.Latencies.Metric = 'Latencies_min';
    for s = 1:length(stages)
        if isinf(latencies(s))
            binMetrics.Latencies.(stages{s}) = NaN;
        else
            binMetrics.Latencies.(stages{s}) = latencies(s);
        end
    end
    
    % Latency2Min
    binMetrics.Latency2Min = struct();
    binMetrics.Latency2Min.Category = 'Latency2Min';
    binMetrics.Latency2Min.Metric = 'Latency2Min_min';
    for s = 1:length(stages)
        if isinf(latencies2Min(s))
            binMetrics.Latency2Min.(stages{s}) = NaN;
        else
            binMetrics.Latency2Min.(stages{s}) = latencies2Min(s);
        end
    end
    
    % Transitions
    binMetrics.Transitions = struct();
    binMetrics.Transitions.Category = 'Transitions';
    binMetrics.Transitions.Metric = 'Trans_Count';
    binMetrics.Transitions.Count = transCount;
    
    % Detailed Transitions with counts between stage pairs
    binMetrics.DetailedTransitions = struct();
    binMetrics.DetailedTransitions.StageNames = stages;
    binMetrics.DetailedTransitions.Counts = transMatrix;
    binMetrics.DetailedTransitions.Wake_to_SWS = transMatrix(1,2);
    binMetrics.DetailedTransitions.Wake_to_REM = transMatrix(1,3);
    binMetrics.DetailedTransitions.SWS_to_Wake = transMatrix(2,1);
    binMetrics.DetailedTransitions.SWS_to_REM = transMatrix(2,3);
    binMetrics.DetailedTransitions.REM_to_Wake = transMatrix(3,1);
    binMetrics.DetailedTransitions.REM_to_SWS = transMatrix(3,2);
end
