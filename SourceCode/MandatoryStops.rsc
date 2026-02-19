Macro "Mandatory Stops Frequency"(Args)
    vwT = OpenTable("Tours", "FFB", {Args.MandatoryTours})
    pbar = CreateObject("G30 Progress Bar", "Mandatory Stops Frequency", true, 2)

    models = {'Work', 'Univ'}
    seeds = {3599969, 3699961}
    for i = 1 to models.length do
        model = models[i]
        spec = {Purpose: model,
                ToursView: vwT,
                Utility: Args.(model + "StopsFreqUtility"),
                ChoicesField: "StopsChoice",
                RandomSeed: seeds[i]}
        RunMacro("Mandatory Stops Choice", Args, spec)
        if pbar.Step() then
            Return()
    end

    // Post processing
    // Fill number of forward and return stops based on the StopsChoice field above
    v = GetDataVector(vwT + "|", "StopsChoice",)
    vecsSet.NForwardStops = if v = null then null else s2i(Left(v,1))
    vecsSet.NReturnStops = if v = null then null else s2i(Right(v,1))
    SetDataVectors(vwT + "|", vecsSet,)
    pbar.Destroy()

    // For adults on Work tours who make a forward stop and also have school dropoffs, figure out if stop is before or after dropoffs
    SetView(vwT)
    n = SelectByQuery('Selection', 'several', 'Select * where NForwardStops > 0 and NumDropoffs > 0',)
    SetRandomSeed(3799973)
    vRand = RandSamples(n, 'Uniform',)
    prob = Args.StopBeforeDropoffsProb
    v = if vRand < prob then 1 else 0
    SetDataVector(vwT + "|Selection", "IsStopBeforeDropoffs", v, )

    // For adults on Work tours who make a return stop and also have school pickups, figure out if stop is before or after pickups
    SetView(vwT)
    n = SelectByQuery('Selection', 'several', 'Select * where NReturnStops > 0 and NumPickups > 0',)
    SetRandomSeed(3899989)
    vRand = RandSamples(n, 'Uniform',)
    prob = Args.StopBeforePickupsProb
    v = if vRand < prob then 1 else 0
    SetDataVector(vwT + "|Selection", "IsStopBeforePickups", v, )
    CloseView(vwT)
    Return(true)
endMacro


Macro "Mandatory Stops Choice"(Args, spec)
    purpose = spec.Purpose
    filter = printf("TourPurpose = '%s'", {purpose})

    // Join Tours to PersonHH
    abm = RunMacro("Get ABM Manager", Args)
    vwJ = JoinViews("ToursData", GetFieldFullSpec(spec.ToursView, "PerID"), GetFieldFullSpec(abm.PersonHHView, abm.PersonID), )

    tag = purpose + "MandatoryStops"
    obj = CreateObject("PMEChoiceModel", {ModelName: purpose + " Mandatory Stops"})
    obj.OutputModelFile = Args.[Output Folder] + "\\Intermediate\\" + tag + ".mdl"
    obj.AddTableSource({SourceName: "TourData", View: vwJ, IDField: "TourID"})
    obj.AddMatrixSource({SourceName: "AutoSkim", File: Args.HighwaySkimAM, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
    obj.AddPrimarySpec({Name: "TourData", Filter: filter, OField: "Origin", DField: "Destination"})
    obj.AddUtility({UtilityFunction: spec.Utility})
    obj.AddOutputSpec({ChoicesField: spec.ChoicesField})
    obj.ReportShares = 1
    obj.RandomSeed = spec.RandomSeed
    ret = obj.Evaluate()
    if !ret then
        Throw("Running '" + tag + "' choice model failed.")
    Args.(purpose + " Stops Spec") = CopyArray(ret)

    if spec.LeaveDataOpen = null then 
        CloseView(vwJ)
endMacro



Macro "MandatoryStops Destination"(Args)
    // Run Destination Choice
    dirs = {"Forward", "Return"}
    types = {"Work", "Univ"}
    periods = {"AM", "PM", "OP"}
    tourFile = Args.MandatoryTours
    objT = CreateObject("Table", tourFile)
    spec = {ToursView: objT.GetView()}

    // Compute size variable field. Add to TAZDemographics output table
    objD = CreateObject("Table", Args.DemographicOutputs)
    obj4D = CreateObject("Table", Args.AccessibilitiesOutputs)
    outFld = "MandatoryStops_DCAttr"
    newFlds = {{FieldName: outFld, Type: "real", Width: 12, Decimals: 2}}
    objD.AddFields({Fields: newFlds})
    objJ = objD.Join({Table: obj4D, LeftFields: {"TAZ"}, RightFields: {"TAZID"}})
    
    opt = null
    opt.TableObject = objJ
    opt.Equation = Args.MandStopSizeVar
    opt.FillField = outFld
    opt.ExponentiateCoeffs = 1
    RunMacro("Compute Size Variable", opt)
    objJ = null
    objD = null
    obj4D = null
    
    pbar = CreateObject("G30 Progress Bar", "Intermediate Stops Destinations, 12 segments: (Forward, Return), (AM, PM, OP) and (Work, Univ)", false, 15)
    for dir in dirs do
        // Get Stop Formula Fields
        spec.Direction = dir
        ODInfo = RunMacro("Create Stop Formula Fields", spec)
        spec.ODInfo = ODInfo

        for period in periods do
            spec.Period  = period
            spec.StopFilter = printf("N%sStops >= 1", {dir}) // e.g. NForwardStops >= 1
            spec.PeriodFilter = printf("TOD%s = '%s'", {dir, period})
            skimFile = Args.("HighwaySkim" + period)

            // Calculate delta TT matrix
            deltaTT = GetTempPath() + "DeltaTT_" + dir + "_" + period + ".mtx"
            spec.DeltaSkim = deltaTT
            spec.MatrixSpec = {File: skimFile, Core: "Time", RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"}
            spec.OutputCoreName = "DeltaTT"
            ret = RunMacro("Calculate Delta TT", Args, spec)
            if ret = 2 then continue // No records for delta TT calculation. Move on to next period.
            spec.DeltaTT = deltaTT

            // Calculate delta Dist matrix
            deltaDist = GetTempPath() + "DeltaDist_" + dir + "_" + period + ".mtx"
            spec.DeltaSkim = deltaDist
            spec.MatrixSpec = {File: skimFile, Core: "Distance", RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"}
            spec.OutputCoreName = "DeltaDist"
            ret = RunMacro("Calculate Delta TT", Args, spec)
            if ret = 2 then continue // No records for delta TT calculation. Move on to next period.
            spec.DeltaDist = deltaDist
            
            // Purposes Loop
            for type in types do
                spec.Type = type
                spec.RandomSeed = 3999971 + 100*dirs.position(dir) + 10*periods.position(period) + types.position(type)
                RunMacro("Intermediate Stop DC", Args, spec)
                pbar.Step()
            end
        end     // period loop

        // Now that destinations have been computed, fill in the realized detour travel times
        if dir = "Forward" then 
            depFld = 'TourStartTime'
        else 
            depFld = 'DestDepTime'
        
        opt = {ToursObj: objT, Filter: printf("N%sStops > 0", {dir}), ODInfo: ODInfo, 
                StopTAZField: "Stop" + dir + "TAZ", ModeField: dir + "Mode", DepTimeField: depFld, 
                OutputField: dir + "StopDeltaTT"}
        RunMacro("Calculate Detour TT", Args, opt)

        // Remove infeasible stops (where stops delta TT is null or more than 45 min)
        filter = printf("N%sStops > 0 and (%sStopDeltaTT = null or %sStopDeltaTT > 45)", {dir, dir, dir})
        opt = {TableObject: objT, Filter: filter, Direction: dir}
        RunMacro("Remove Infeasible Stops", opt)
        
        // Destroy temp formula fields
        RunMacro("Destroy Stop Formula Fields", spec)
        pbar.Step()

    end     // dirs loop
    pbar.Destroy()
    objT = null
    Return(true)
endMacro


/*
    Create the origin and destination given the stop direction and the stop number
        Stop #1, Forward direction
        - No drop offs:                     O: 'Origin', D: 'Destination'
        - With drop offs:   
            - IS Stop before dropoffs:      O: 'Origin', D: FirstDropoffTAZ
            - IS Stop after dropoffs:       O: LastDropoffTAZ, D: 'Destination'
        
        Note: Model does not allow for 2 intermediate stops on mandatory tours
        Stop #2, Forward direction
        - No drop offs:                     O: Stop#1TAZ, D: 'Destination'
        - With drop offs:                   2 intermediate stops not allowed with dropoffs

        Stop #1, Return direction
        - No pick ups:                      O: 'Destination', D: 'Origin'
        - With pick ups:            
            - IS Stop before pickups:       O: 'Destination', D: FirstPickupTAZ
            - IS Stop after pickups:        O: LastPickupTAZ, D: 'Origin'
        
        Note: Model does not allow for 2 intermediate stops on mandatory tours
        Stop #2, Return direction
        - No pick ups:                      O: Stop#1TAZ, D: 'Origin'
        - With pick ups:                    2 intermediate stops not allowed with pickups
*/
Macro "Create Stop Formula Fields"(spec)
    vwT = spec.ToursView
    dir = spec.Direction

    ODInfo = null
    if dir = 'Forward' then do
        // Get last drop off stop
        expr = "if DropoffTAZ2 <> null then DropoffTAZ2 else DropoffTAZ1"
        lastDO = CreateExpression(vwT, "LastDO", expr,)
        ODInfo.LastStop = lastDO
        
        expr1 = "if nz(NumDropoffs) = 0 then Origin " +
                "else if IsStopBeforeDropoffs = 1 then Origin " + 
                "else " + lastDO
        ODInfo.Origin = CreateExpression(vwT, "DeltaTTOrig", expr1, )

        expr2 = "if nz(NumDropoffs) = 0 then Destination " +
                "else if IsStopBeforeDropoffs = 1 then DropoffTAZ1 " + 
                "else Destination"
        ODInfo.Destination = CreateExpression(vwT, "DeltaTTDest", expr2, )
    end
    else do
        // Get last pickup stop
        expr = "if PickupTAZ2 <> null then PickupTAZ2 else PickupTAZ1"
        lastPU = CreateExpression(vwT, "LastPU", expr,)
        ODInfo.LastStop = lastPU

        expr1 = "if nz(NumPickups) = 0 then Destination " +
                "else if IsStopBeforePickups = 1 then Destination " + 
                "else " + lastPU
        ODInfo.Origin = CreateExpression(vwT, "DeltaTTOrig", expr1, )

        expr2 = "if nz(NumPickups) = 0 then Origin " +
                "else if IsStopBeforePickups = 1 then PickupTAZ1 " + 
                "else Origin"
        ODInfo.Destination = CreateExpression(vwT, "DeltaTTDest", expr2, )
    end
    Return(ODInfo)
endMacro



Macro "Destroy Stop Formula Fields"(spec)
    vwT = spec.ToursView
    ODInfo = spec.ODInfo
    DestroyExpression(GetFieldFullSpec(vwT, ODInfo.Origin))
    DestroyExpression(GetFieldFullSpec(vwT, ODInfo.Destination))
    DestroyExpression(GetFieldFullSpec(vwT, ODInfo.LastStop))
endMacro


/*
    Given the tours view and the direction, stop # and period:
        - Select records from tour view for direction, stop#, period
        - Obtain appropriate auto skim
        - Compute a deltaTT matrix of TourID by Destinations that
            - For each row (TourID) contains the deltaTT incurred by visiting each destination j on route from tour Origin to tour Destination
*/
Macro "Calculate Delta TT"(Args, spec)
    vwT = spec.ToursView
    dir = spec.Direction
    ODInfo = spec.ODInfo

    stopfilter = spec.StopFilter
    periodfilter = spec.PeriodFilter
    if periodfilter <> null then
        deltaCalcFilter = printf("(%s) and (%s)", {stopfilter, periodfilter})
    else
        deltaCalcFilter = stopfilter

    // Export relevant records to a table
    SetView(vwT)
    n = SelectByQuery("DeltaCalc", "several", "Select * where " + deltaCalcFilter,)
    if n = 0 then 
        Return(2)
    vwTmp = ExportView(vwT + "|DeltaCalc", "MEM", "TempForDeltaCalc", {"TourID", ODInfo.Origin, ODInfo.Destination},)

    // Need to copy matrix with only the relevant indices
    skimMat = spec.MatrixSpec.File
    core = spec.MatrixSpec.Core
    ridx = spec.MatrixSpec.RowIndex
    cidx = spec.MatrixSpec.ColIndex
    m = OpenMatrix(skimMat,)
    mc = CreateMatrixCurrency(m, core, ridx, cidx,)
    mtxTemp = GetTempPath() + "TempSkim.mtx"
    mNew = CopyMatrix(mc, {"File Name": mtxTemp, Label: "TempSkim", Indices: "Current"})
    mNew = null
    mc = null
    m = null

    // Copy transpose matrix (prior to transpose) if provided
    if spec.MatrixSpecR <> null then do
        skimMat = spec.MatrixSpecR.File
        core = spec.MatrixSpecR.Core
        ridx = spec.MatrixSpecR.RowIndex
        cidx = spec.MatrixSpecR.ColIndex
        m = OpenMatrix(skimMat,)
        mc = CreateMatrixCurrency(m, core, ridx, cidx,)
        mtxTempR = GetTempPath() + "TempSkimR.mtx"
        mNew = CopyMatrix(mc, {"File Name": mtxTempR, Label: "TempSkimR", Indices: "Current"})
        mNew = null
        mc = null
        m = null
    end
    
    // Calculate Delta TT Matrix
    Opts = null
    Opts.SourceData = {ViewName: vwTmp, ID: "TourID", Row: ODInfo.Origin, Column: ODInfo.Destination}
    Opts.SkimMatrixFile = mtxTemp
    Opts.SkimMatrixCore = spec.MatrixSpec.Core
    if spec.MatrixSpecR <> null then do
        Opts.SkimMatrixFile2 = mtxTempR
        Opts.SkimMatrixCore2 = spec.MatrixSpecR.Core
    end
    Opts.DeltaSkim = spec.DeltaSkim
    obj = CreateObject("TransCAD.ABM")
    ret = obj.IntermediateStopsDeltaTT(Opts)
    mc_deltaTime = ret.Matrix
    if spec.OutputCoreName <> null then do
        m = mc_deltaTime.Matrix
        SetMatrixCoreNames(m, {spec.OutputCoreName})
        RenameMatrix(m, spec.OutputCoreName)
    end

    CloseView(vwTmp)
    Return(1)
endMacro


/*
    Run Intermediate stops destination choice model given trip direction, stop number, period and purpose 
*/
Macro "Intermediate Stop DC"(Args, spec)
    vwT = spec.ToursView
    dir = spec.Direction
    period = spec.Period
    deltaTT = spec.DeltaTT
    deltaDist = spec.DeltaDist
    type = spec.Type
    ODInfo = spec.ODInfo

    // Filters
    stopfilter = printf("N%sStops >= 1", {dir})             // e.g. NForwardStops >= 1
    typefilter = printf("TourPurpose = '%s'", {type})       // e.g. TourPurpose = 'Work'
    periodfilter = printf("TOD%s = '%s'", {dir, period})    // e.g. TODForward = 'AM'

    filter = printf("(%s) and (%s) and (%s)", {stopfilter, periodfilter, typefilter})
    
    SetView(vwT)
    n = SelectByQuery("__Selection", "several", "Select * where " + filter,)
    if n > 0 then do
        util_fn = Args.MandStopDestUtility
        choiceFld = printf("Stop%sTAZ", {dir})
        
        // Run Destination Choice
        tag = type + "_" + dir + "_" + period + "Stop"
        obj = CreateObject("PMEChoiceModel", {ModelName: tag})
        obj.OutputModelFile = Args.[Output Folder] + "\\Intermediate\\" + tag + ".dcm"
        obj.AddTableSource({SourceName: "TourData", View: vwT, IDField: "TourID"})
        obj.AddTableSource({SourceName: "TAZData", File: Args.DemographicOutputs, IDField: "TAZ"})
        obj.AddTableSource({SourceName: "TAZ4Ds", File: Args.AccessibilitiesOutputs, IDField: "TAZID"})
        obj.AddMatrixSource({SourceName: "AutoSkim", File: Args.("HighwaySkim" + period), RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
        obj.AddMatrixSource({SourceName: "Intrazonal", File: Args.IZMatrix, RowIndex: "TAZ", ColIndex: "TAZ"})
        obj.AddMatrixSource({SourceName: "DeltaTT", File: deltaTT, PersonBased: 1})
        obj.AddMatrixSource({SourceName: "DeltaDist", File: deltaDist, PersonBased: 1})
        obj.AddPrimarySpec({Name: "TourData", Filter: filter, OField: ODInfo.Origin})
        obj.AddUtility({UtilityFunction: util_fn, SubstituteStrings: {{"<Dir>", dir}}})
        obj.AddDestinations({DestinationsSource: "AutoSkim", DestinationsIndex: "InternalTAZ"})
        obj.AddSizeVariable({Name: "TAZData", Field: 'MandatoryStops_DCAttr'})
        obj.AddOutputSpec({ChoicesField: choiceFld})
        obj.RandomSeed = spec.RandomSeed
        ret = obj.Evaluate()
        if !ret then
            Throw("Running 'Intermediate Stop Location' model failed for: " + tag)
    end
endMacro



Macro "MandatoryStops Duration"(Args)
    types = {"Work", "Univ"}
    dirs = {"Forward", "Return"}
    pbar = CreateObject("G30 Progress Bar", "Running Intemediate Stops Duration Choice for combination of (Work, Univ) and (Forward, Return)", false, 4)
    
    objT = CreateObject("Table", Args.MandatoryTours)
    abm = RunMacro("Get ABM Manager", Args)
    for type in types do
        for dir in dirs do
            spec = null
            spec.Type = type
            spec.Direction = dir
            spec.ToursView = objT.GetView()
            spec.RandomSeed = 4099981 + 10*types.position(type) + dirs.position(dir)
            RunMacro("Mandatory Duration Choice", Args, spec)
            pbar.Step()
        end
    end
    pbar.Destroy()
    Return(true)
endMacro


Macro "Mandatory Duration Choice"(Args, spec)
    // Get dir specific data
    dir = spec.Direction
    type = spec.Type
    filter = printf("N%sStops > 0", {dir})
    
    abm = RunMacro("Get ABM Manager", Args)
    vw = JoinViews("ToursData", GetFieldFullSpec(spec.ToursView, "PerID"), GetFieldFullSpec(abm.PersonHHView, abm.PersonID), )

    // Run Duration choice model
    tag = "Stops_" + type + "_" + dir
    obj = CreateObject("PMEChoiceModel", {SourcesObject: Args.SourcesObject, ModelName: tag + " Duration"})
    obj.OutputModelFile = Args.[Output Folder] + "\\Intermediate\\" + tag + "_Duration.mdl"
    obj.AddTableSource({SourceName: "TourData", View: vw, IDField: "TourID"})
    obj.AddPrimarySpec({Name: "TourData", Filter: filter})
    obj.AddUtility({UtilityFunction: Args.(type + "StopDurUtility"),
                    AvailabilityExpressions: Args.StopDurAvail, 
                    SubstituteStrings: {{"<dir>", dir}}})
    obj.AddOutputSpec({ChoicesField: dir + "StopDurChoice"})
    obj.RandomSeed = spec.RandomSeed
    obj.ReportShares = 1
    ret = obj.Evaluate()
    if !ret then
        Throw("Running 'Mandatory Stop Duration' model failed for: " + type + "_" + dir)
    Args.(tag + " Spec") = CopyArray(ret)
    obj = null

    // Simulate Time based on choice interval
    SetView(vw)
    n = SelectByQuery("__Selection", "several", "Select * where " + filter,)
    opt = {ViewSet: vw + "|__Selection", InputField: dir + "StopDurChoice", OutputField: dir + "StopDuration", AlternativeIntervalInMin: 1}
    RunMacro("Simulate Time", opt)

    if spec.LeaveDataOpen = null then 
        CloseView(vw)
endMacro


// Mandatory stops scheduling
// Modify departure time from home/work and arrival time at work/home
// Fields modified in the tours table are:
// - TourStartTime, DestArrTime, TourEndTime and DestDepTime
// Do not concern (yet) with tours potentially encroaching previous or subsequent mandatory tours
Macro "MandatoryStops Scheduling"(Args)
    objT = CreateObject("Table", Args.MandatoryTours)
    spec = {ToursObj: objT}
    pbar = CreateObject("G30 Progress Bar", "Running Intermediate Stops Scheduling for forward and return stops", false, 8)

    maxTours = 2 // Max two tours that have stops

    for i = 1 to maxTours do
        spec.TourNo = i
        // Add temporary fields
        RunMacro("Mandatory Stops Prep", spec)
        pbar.Step()

        RunMacro("Stop Forward Schedule", spec)
        pbar.Step()
        
        RunMacro("Stop Return Schedule", spec)
        pbar.Step()

        RunMacro("Mandatory Stops Postprocess", spec)
        pbar.Step()
    end
    pbar.Destroy()

    RunMacro("Fill Stop Travel Times", Args, spec)

    Return(true)
endMacro


Macro "Mandatory Stops Prep"(spec)
    // Add fields
    obj = spec.ToursObj
    flds = {{FieldName: "PrevPID", Type: "Integer"},
            {FieldName: "NextPID", Type: "Integer"},
            {FieldName: "PrevTourEnd", Type: "Integer"},
            {FieldName: "NextTourStart", Type: "Integer"},
            {FieldName: "RemoveFStop", Type: "Short"},
            {FieldName: "RemoveRStop", Type: "Short"}
           }
    obj.AddFields({Fields: flds})

    // Fill fields
    vwT = obj.GetView()
    order = {{"PerID", "Ascending"}, {"ActivityStartTime", "Ascending"}}
    obj.Sort({FieldArray: order})
    vecs = obj.GetDataVectors({FieldNames: {"PerID", "TourEndTime", "TourStartTime"}})
    
    vecsSet = null
    vecsSet.PrevPID = RunMacro("Shift Vector", {Vector: vecs.PerID, Method: 'Prev'})
    vecsSet.NextPID = RunMacro("Shift Vector", {Vector: vecs.PerID, Method: 'Next'})
    vecsSet.PrevTourEnd = RunMacro("Shift Vector", {Vector: vecs.TourEndTime, Method: 'Prev'})
    vecsSet.NextTourStart = RunMacro("Shift Vector", {Vector: vecs.TourStartTime, Method: 'Next'})
    obj.SetDataVectors({FieldData: vecsSet})
endMacro


// ===== Forward scheduing
Macro "Stop Forward Schedule"(spec)
    toursObj = spec.ToursObj
    tourQry = printf("MandTourNo = %u", {spec.TourNo})

    // ===== Forward scheduing
    // Case 1: No school dropoffs
    // Leave early to accomodate the stop
    // If early departure encroaches on previous tour, then leave 5 min after the previous tour ends (thereby incurring a delay)
    // If delay results in a modified main activity duration of less than 30 min, drop the stop
    qry = printf("NForwardStops > 0 and nz(NumDropoffs) = 0 and %s", {tourQry})
    nF = toursObj.SelectByQuery({SetName: "FStopNoPUDO", Query: qry})
    if nF > 0 then do
        flds = {"ForwardStopDeltaTT", "ForwardStopDuration", "TourStartTime", "DestArrTime", 
                "PrevPID", "PrevTourEnd", "PerID", "ActivityDuration", "RemoveFStop"}
        vecs = toursObj.GetDataVectors({FieldNames: flds})
        vStopTime = vecs.ForwardStopDeltaTT + vecs.ForwardStopDuration
        vDesiredStart = vecs.TourStartTime - vStopTime  // Desired start time if there is no overlap onto previous tour
        
        // Check if early start overlaps with previous tour
        vOverlapPrevTour = if (vecs.PerID = vecs.PrevPID) and (vDesiredStart < vecs.PrevTourEnd) then 1 else 0
        
        // And set delay to diff of new constrained start time and the desired start time
        vDelay = if vOverlapPrevTour = 1 then (vecs.PrevTourEnd + 5) - vDesiredStart else 0
        vNewActDur = vecs.ActivityDuration - vDelay // Realized activity duration due to the delay
        vRemoveStop = if (vNewActDur < 30) then 1 else vecs.RemoveFStop

        vecsSet = null
        vecsSet.TourStartTime = if vRemoveStop <> 1 then vDesiredStart + vDelay else vecs.TourStartTime
        vecsSet.DestArrTime = if vRemoveStop <> 1 then vecs.DestArrTime + vDelay else vecs.DestArrTime
        vecsSet.RemoveFStop = vRemoveStop
        toursObj.SetDataVectors({FieldData: vecsSet})
    end

    // Case 2: Stop before school dropffs
    // School dropoff times not negotiable
    // Leave home early by (stop duration + excess delta TT)
    // Remove stop if it encroaches on previous tour
    // Note stop duration was constrained to be in the 5-15 min range by the duration model
    qry = printf("NForwardStops > 0 and NumDropoffs > 0 and IsStopBeforeDropoffs = 1 and %s", {tourQry})
    nF = toursObj.SelectByQuery({SetName: "FStopPUDO1", Query: qry})
    if nF > 0 then do
        flds = {"ForwardStopDeltaTT", "ForwardStopDuration", "TourStartTime", 
                "PrevPID", "PrevTourEnd", "PerID", "RemoveFStop"}
        vecs = toursObj.GetDataVectors({FieldNames: flds})
        vStopTime = vecs.ForwardStopDeltaTT + vecs.ForwardStopDuration
        vDesiredStart = vecs.TourStartTime - vStopTime
        vRemoveStop = if (vecs.PerID = vecs.PrevPID) and (vDesiredStart < vecs.PrevTourEnd) then 1 else vecs.RemoveFStop

        vecsSet = null
        vecsSet.TourStartTime = if vRemoveStop <> 1 then vDesiredStart else vecs.TourStartTime
        vecsSet.RemoveFStop = vRemoveStop
        toursObj.SetDataVectors({FieldData: vecsSet})
    end

    // Case 3: Stop after school dropffs
    // School dropoff times not negotiable 
    // Arrive late to the mandatory dest
    // If arriving late reduces main activity duration to less than 30 min, drop the stop
    qry = printf("NForwardStops > 0 and NumDropoffs > 0 and IsStopBeforeDropoffs = 0 and %s", {tourQry})
    nF = toursObj.SelectByQuery({SetName: "FStopPUDO2", Query: qry})
    if nF > 0 then do
        flds = {"ForwardStopDeltaTT", "ForwardStopDuration", "DestArrTime", "ActivityDuration", "RemoveFStop"}
        vecs = toursObj.GetDataVectors({FieldNames: flds})
        vStopTime = vecs.ForwardStopDeltaTT + vecs.ForwardStopDuration
        vNewActDur = vecs.ActivityDuration - vStopTime
        vRemoveStop = if (vNewActDur < 30) then 1 else vecs.RemoveFStop
        
        vecsSet = null
        vecsSet.DestArrTime = if vRemoveStop <> 1 then vecs.DestArrTime + vStopTime else vecs.DestArrTime
        vecsSet.RemoveFStop = vRemoveStop
        toursObj.SetDataVectors({FieldData: vecsSet})
    end
endMacro


// ===== Return scheduing
Macro "Stop Return Schedule"(spec)
    toursObj = spec.ToursObj
    tourQry = printf("MandTourNo = %u", {spec.TourNo})

    // ===== Return scheduing
    // Case 1: No school pickups
    // Arrive late at home
    // If arrival late at home encroaches on next tour, leave work early
    // if leaving work early reduces main activity duration to under 30 min, drop the stop
    qry = printf("NReturnStops > 0 and nz(NumPickups) = 0 and %s", {tourQry})
    nR = toursObj.SelectByQuery({SetName: "RStopNoPUDO", Query: qry})
    if nR > 0 then do
        flds = {"ReturnStopDeltaTT", "ReturnStopDuration", "TourEndTime", "DestDepTime", 
                "NextPID", "NextTourStart", "PerID", "ActivityDuration", "RemoveRStop"}
        vecs = toursObj.GetDataVectors({FieldNames: flds})
        vStopTime = vecs.ReturnStopDeltaTT + vecs.ReturnStopDuration
        vLateArr = vecs.TourEndTime + vStopTime
        vOverlapNextTour = if (vecs.PerID = vecs.NextPID) and (vLateArr > vecs.NextTourStart) then 1 else 0
        vDelay = if vOverlapNextTour = 1 then vLateArr - vecs.NextTourStart else 0
        vNewActDur = vecs.ActivityDuration - vDelay
        vRemoveStop = if (vNewActDur < 30) then 1 else vecs.RemoveRStop
        
        vecsSet = null
        vecsSet.DestDepTime = if vRemoveStop <> 1 then vecs.DestDepTime - vDelay else vecs.DestDepTime
        vecsSet.TourEndTime = if vRemoveStop <> 1 then vLateArr - vDelay else vecs.TourEndTime
        vecsSet.RemoveRStop = vRemoveStop
        toursObj.SetDataVectors({FieldData: vecsSet})
    end

    // Case 2: Stop before school pickups
    // School pickup times not negotiable
    // Leave work early by (stop duration + excess delta TT)
    // If leaving work early reduces the main activity duration to less than 30 min, drop the stop
    // Note stop duration was constrained to be in the 5-15 min range by the duration model
    qry = printf("NReturnStops > 0 and NumPickups > 0 and IsStopBeforePickups = 1 and %s", {tourQry})
    nR = toursObj.SelectByQuery({SetName: "RStopPUDO1", Query: qry})
    if nR > 0 then do
        flds = {"ReturnStopDeltaTT", "ReturnStopDuration", "DestDepTime", "ActivityDuration", "RemoveRStop"}
        vecs = toursObj.GetDataVectors({FieldNames: flds})
        vStopTime = vecs.ReturnStopDeltaTT + vecs.ReturnStopDuration
        vDesiredDep = vecs.DestDepTime - vStopTime
        vNewActDur = vecs.ActivityDuration - vStopTime
        vRemoveStop = if (vNewActDur < 30) then 1 else vecs.RemoveRStop

        vecsSet = null
        vecsSet.DestDepTime = if vRemoveStop <> 1 then vDesiredDep else vecs.DestDepTime
        vecsSet.RemoveRStop = vRemoveStop
        toursObj.SetDataVectors({FieldData: vecsSet})
    end

    // Case 3: Stop after school pickups
    // School pickup times not negotiable 
    // Arrive late at home
    // If late arrival encroaches on next tour, srop stop
    qry = printf("NReturnStops > 0 and NumPickups > 0 and IsStopBeforePickups = 0 and %s", {tourQry})
    nR = toursObj.SelectByQuery({SetName: "RStopPUDO2", Query: qry})
    if nR > 0 then do
        flds = {"ReturnStopDeltaTT", "ReturnStopDuration", "TourEndTime", "NextPID", "NextTourStart", "PerID", "RemoveRStop"}
        vecs = toursObj.GetDataVectors({FieldNames: flds})
        vStopTime = vecs.ReturnStopDeltaTT + vecs.ReturnStopDuration
        vNewArr = vecs.TourEndTime + vStopTime
        vRemoveStop = if (vecs.PerID = vecs.NextPID) and (vNewArr > vecs.NextTourStart) then 1 else vecs.RemoveRStop

        vecsSet = null
        vecsSet.TourEndTime = if vRemoveStop <> 1 then vNewArr else vecs.TourEndTime
        vecsSet.RemoveRStop = vRemoveStop
        toursObj.SetDataVectors({FieldData: vecsSet})
    end
endMacro


Macro "Mandatory Stops Postprocess"(spec)
    obj = spec.ToursObj

    // Remove infeasible stops
    opt = {TableObject: obj, Filter: "RemoveFStop = 1", Direction: "Forward"}
    RunMacro("Remove Infeasible Stops", opt)

    opt = {TableObject: obj, Filter: "RemoveRStop = 1", Direction: "Return"}
    RunMacro("Remove Infeasible Stops", opt)

    // Delete additional fields from the tours table
    obj.DropFields({FieldNames: {"PrevPID", "NextPID", "PrevTourEnd", "NextTourStart", "RemoveFStop", "RemoveRStop"}})
endMacro



Macro "Fill Stop Travel Times"(Args, spec)
    toursObj = spec.ToursObj
    vwT = toursObj.GetView()
    
    // **************************** Forward Stops ***********************************
    qry = "NForwardStops > 0"
    
    //****************** Travel Time to Stop ***********************
    // Travel time from the previous origin to the intermediate stop
    // Create expressions for origin and departure time fields
    expr = "if IsStopBeforeDropoffs = 1 or nz(NumDropoffs) = 0 then Origin " +
           "else if NumDropoffs = 2 then DropoffTAZ2 else DropoffTAZ1"
    orig = CreateExpression(vwT, "OFld", expr,)
    
    expr = "if IsStopBeforeDropoffs = 1 or nz(NumDropoffs) = 0 then TourStartTime " + 
           "else if NumDropoffs = 2 then DepDropoff2 else DepDropoff1"
    depTime = CreateExpression(vwT, "DepFld", expr,)
    
    fillSpec = {View: vwT, OField: orig, DField: "StopForwardTAZ", FillField: "TimeToStopF", 
                Filter: qry, ModeField: "ForwardMode", DepTimeField: depTime}
    RunMacro("Fill Travel Times", Args, fillSpec)
    DestroyExpression(GetFieldFullSpec(vwT, orig))
    DestroyExpression(GetFieldFullSpec(vwT, depTime))

     //****************** Travel Time from Stop ***********************
    // Travel time from intermediate stop to next destination
    // Create expressions for origin and departure time fields
    expr = "if IsStopBeforeDropoffs = 1 then DropoffTAZ1 else Destination"
    dest = CreateExpression(vwT, "DFld", expr,)
    
    expr = "if IsStopBeforeDropoffs = 1 or nz(NumDropoffs) = 0 then TourStartTime + TimeToStopF + ForwardStopDuration " + 
           "else if NumDropoffs = 2 then DepDropoff2 + TimeToStopF + ForwardStopDuration " + 
           "else DepDropoff1 + TimeToStopF + ForwardStopDuration"
    depTime = CreateExpression(vwT, "DepFld", expr,)
    
    fillSpec = {View: vwT, OField: "StopForwardTAZ", DField: dest, FillField: "TimeFromStopF", 
                Filter: qry, ModeField: "ForwardMode", DepTimeField: depTime}
    RunMacro("Fill Travel Times", Args, fillSpec)
    DestroyExpression(GetFieldFullSpec(vwT, dest))
    DestroyExpression(GetFieldFullSpec(vwT, depTime))

    // **************************** Return Stops ***********************************
    qry = "NReturnStops > 0"

    //****************** Travel Time to Stop ***********************
    // Travel time from the previous origin to the intermediate stop
    // Create expressions for origin and departure time fields
    expr = "if IsStopBeforePickups = 1 or nz(NumPickups) = 0 then Destination " +
           "else if NumPickups = 2 then PickupTAZ2 else PickupTAZ1"
    orig = CreateExpression(vwT, "OFld", expr,)
    
    expr = "if IsStopBeforePickups = 1 or nz(NumPickups) = 0 then DestDepTime " +
           "else if NumPickups = 2 then DepPickup2 else DepPickup1"
    depTime = CreateExpression(vwT, "DepFld", expr,)
    
    fillSpec = {View: vwT, OField: orig, DField: "StopReturnTAZ", FillField: "TimeToStopR", 
                Filter: qry, ModeField: "ReturnMode", DepTimeField: depTime}
    RunMacro("Fill Travel Times", Args, fillSpec)
    DestroyExpression(GetFieldFullSpec(vwT, orig))
    DestroyExpression(GetFieldFullSpec(vwT, depTime))

     //****************** Travel Time from Stop ***********************
    // Travel time from intermediate stop to next destination
    // Create expressions for origin and departure time fields
    expr = "if IsStopBeforePickups = 1 then PickupTAZ1 else Origin"
    dest = CreateExpression(vwT, "DFld", expr,)
    
    expr = "if IsStopBeforePickups = 1 or nz(NumPickups) = 0 then DestDepTime + TimeToStopR + ReturnStopDuration " +
           "else if NumPickups = 2 then DepPickup2 + TimeToStopR + ReturnStopDuration " +
           "else DepPickup1 + TimeToStopR + ReturnStopDuration"
    depTime = CreateExpression(vwT, "DepFld", expr,)
    
    fillSpec = {View: vwT, OField: "StopReturnTAZ", DField: dest, FillField: "TimeFromStopR", 
                Filter: qry, ModeField: "ReturnMode", DepTimeField: depTime}
    RunMacro("Fill Travel Times", Args, fillSpec)
    DestroyExpression(GetFieldFullSpec(vwT, dest))
    DestroyExpression(GetFieldFullSpec(vwT, depTime))
endMacro


Macro "Get Time Vector"(v)
    vOut = CopyVector(v)
    vOut = if vOut < 0 then 1440 + vOut
           else if vOut > 1440 then vOut - 1440
           else vOut
    Return(vOut)
endMacro


// Macro that removes the intermediate stop (choice, duration, destination info) if not feasible.
// This happens because the stop may not be feasible given the mode (such as "Walk")
Macro "Remove Infeasible Stops"(opt)
    obj = opt.TableObject
    dir = opt.Direction
    n = obj.SelectByQuery({Query: opt.Filter, SetName: "__Remove"})

    if n > 0 then do
        AppendToLogFile(1, opt.Message + String(n) + " Intermediate " + dir + " stops removed due to schedule constraints.")
        v = Vector(n, "Long",)
        vS = Vector(n, "String",)
        vecsSet = null
        vecsSet.("N" + dir + "Stops") = nz(v)
        vecsSet.("Stop" + dir + "TAZ") = v
        vecsSet.(dir + "StopDurChoice") = vS
        vecsSet.(dir + "StopDuration") = v
        vecsSet.(dir + "StopDeltaTT") = v
        vStopsChoice = obj.StopsChoice
        if dir = 'Forward' then do
            vecsSet.StopsChoice = "0_" + Right(vStopsChoice, 1)
            vecsSet.IsStopBeforeDropoffs = v
        end
        else do
            vecsSet.StopsChoice = Left(vStopsChoice, 1) + "_0"
            vecsSet.IsStopBeforePickups = v
        end
        obj.SetDataVectors({FieldData: vecsSet})        
    end
    obj.ChangeSet()
endMacro
