Macro "Visitor Stops Setup"(Args)
    on error do
        ShowMessage(GetLastError())
        return(0)
    end

    Args.ABMFlag = 2

    objT = CreateObject("Table", Args.VisitorTours)
    flds = {{FieldName: "StopsChoice", Type: "String", Width: 5}}
    dirs = {"Forward", "Return"}
    stopNos = {"1", "2"}
    for dir in dirs do
        flds = flds + {{FieldName: "N" + dir + "Stops", Type: "Short"},
                       {FieldName: "Remove" + dir + "Stop", Type: "Short"},
                       {FieldName: dir + "Stop1to2Time", Type: "Real"}}
        
        for s in stopNos do
            flds = flds + {{FieldName: "Purpose" + dir + "Stop" + s, Type: "String", Width: 5},
                           {FieldName: "Stop" + dir + "TAZ" + s, Type: "Integer"},
                           {FieldName: dir + "StopDurChoice" + s, Type: "String", Width: 15},
                           {FieldName: dir + "StopDuration" + s, Type: "Real"},
                           {FieldName: dir + "StopDeltaTT" + s, Type: "Real"},
                           {FieldName: "TimeToStop" + dir + s, Type: "Real"}}
        end
    end
    flds = flds + {{FieldName: "StopsChoice", Type: "String", Width: 5},
                    {FieldName: "NForwardStops", Type: "Short"},
                    {FieldName: "NReturnStops", Type: "Short"}}
    objT.AddFields({Fields: flds})
    objT = null
    
    return(1)
endMacro


Macro "Visitor Stops Frequency"(Args)
    on error do
        ShowMessage(GetLastError())
        return(0)
    end
    Args.ABMFlag = 2
    
    objTours = CreateObject("Table", Args.VisitorTours)
    objA = CreateObject("Table", Args.AccessibilitiesOutputs)

    // Run Model and populate results
    obj = CreateObject("PMEChoiceModel", {ModelName: "Visitor Stops Frequency"})
    obj.OutputModelFile = printf("%s\\Intermediate\\VisitorStopsFreq.mdl", {Args.[Output Folder]})
    obj.AddTableSource({SourceName: "VisitorData", View: objTours.GetView(), IDField: "TourID"})
    obj.AddTableSource({SourceName: "TAZAccessibilities", View: objA.GetView(), IDField: "TAZID"})
    obj.AddMatrixSource({SourceName: "AutoSkim", File: Args.HighwaySkimAM, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
    obj.AddPrimarySpec({Name: "VisitorData", OField: "Origin", DField: "Destination"})
    obj.AddUtility({UtilityFunction: Args.VisitorStopsFreqUtility})
    obj.AddOutputSpec({ChoicesField: "StopsChoice"})
    obj.ReportShares = 1
    obj.RandomSeed = 989991
    ret = obj.Evaluate()
    if !ret then
        Throw("Visitor stops frequency model failed")
    Args.("VisitorStopsFreq Spec") = CopyArray(ret) // For calibration purposes
    obj = null

    // Fill N Forward and Return Stops
    v = objTours.StopsChoice
    vecsSet = null
    vecsSet.NForwardStops = if v = null then null else s2i(Left(v,1))
    vecsSet.NReturnStops = if v = null then null else s2i(Right(v,1))
    objTours.SetDataVectors({FieldData: vecsSet})

    objA = null
    objTours = null
    return(1)
endMacro


Macro "Visitor Stops Purpose"(Args)
    on error do
        ShowMessage(GetLastError())
        return(0)
    end
    
    Args.ABMFlag = 2
    objT = CreateObject("Table", Args.VisitorTours)
    stopPurps = {"Rec", "Shop", "Other"}
    stopsProb = {0.22, 0.20, 0.58} // Probabilities for Rec, Shop, Other stops
    params = {population: stopPurps, weight: stopsProb}

    dirs = {"Forward", "Return"}
    stopsArr = {"1", "2"}
    for dir in dirs do
        for s in stopsArr do
            filter = printf("N%sStops >= %s", {dir, s}) // e.g. NForwardStops >= 1
            n = objT.SelectByQuery({Query: filter, SetName: "__PurposeSet"})
            if n > 0 then do
                objT.ChangeSet("__PurposeSet")
                SetRandomSeed(4200 + 10*ASCII(Left(dir, 1)) + s2i(s))
                v = RandSamples(n, "Discrete", params)
                vecsSet = null
                vecsSet.("Purpose" + dir + "Stop" + s) = v
                objT.SetDataVectors({FieldData: vecsSet})
            end
        end
    end
    objT = null
    return(1)
endMacro


Macro "Visitor Stops Destination"(Args)
    on error do
        ShowMessage(GetLastError())
        return(0)
    end
    Args.ABMFlag = 2
    
    // Run Destination Choice
    dirs = {"Forward", "Return"}
    types = {"Rec", "Shop", "Other"}
    stopsArr = {"1", "2"}
    
    tourFile = Args.VisitorTours
    objT = CreateObject("Table", tourFile)
    spec = {ToursObject: objT, ToursView: objT.GetView(), SkimFile: Args.HighwaySkimAM}
    pbar = CreateObject("G30 Progress Bar", "Intermediate Stops Destinations: (Forward, Return) and (Rec, Shop, Other) and (Stop1, Stop2)", false, 15)
    for dir in dirs do
        // Get Stop Formula Fields
        spec.Direction = dir
        if dir = "Forward" then
            spec.ODInfo = {Origin: "Origin", Destination: "Destination"}
        else
            spec.ODInfo = {Origin: "Destination", Destination: "Origin"}

        spec.StopFilter = printf("N%sStops >= 1", {dir}) // e.g. NForwardStops >= 1

        // Calculate delta TT matrix
        deltaTT = GetTempPath() + "DeltaTT_" + dir + ".mtx"
        spec.DeltaSkim = deltaTT
        spec.MatrixSpec = {File: spec.SkimFile, Core: "Time", RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"}
        spec.OutputCoreName = "DeltaTT"
        ret = RunMacro("Calculate Delta TT", Args, spec)
        if ret = 2 then continue // No records for delta TT calculation. Move on to next direction.
        spec.DeltaTT = deltaTT

        // Purposes Loop
        for type in types do
            spec.Type = type
            for s in stopsArr do
                spec.StopNo = s
                spec.RandomSeed = 3999971 + 10*types.position(type) + s2i(s)
                ret = RunMacro("Vis Intermediate Stop DC", Args, spec)
                pbar.Step()
            end
        end

        // Now that destinations have been computed, fill in the realized detour travel times
        if dir = "Forward" then 
            depFld = 'TourStartTime'
        else 
            depFld = 'ActivityEndTime'
        
        for stopNo in stopsArr do
            opt = {ToursObj: objT, Filter: printf("N%sStops >= %s", {dir, stopNo}), ODInfo: spec.ODInfo, 
                    StopTAZField: "Stop" + dir + "TAZ" + stopNo, ModeField: "Mode", DepTimeField: depFld, 
                    OutputField: dir + "StopDeltaTT" + stopNo}
            RunMacro("Calculate Detour TT", Args, opt)

            // Remove infeasible stops (where stops delta TT is null or more than 75 min)
            filter = printf("Stop%sTAZ%s <> null and (%sStopDeltaTT%s = null or %sStopDeltaTT%s > 75)", {dir, stopNo, dir, stopNo, dir, stopNo})
            opt = {TableObject: objT, Filter: filter, Direction: dir, StopNo: stopNo}
            RunMacro("Remove Infeasible Vis Stops", opt)
        end 

        filter = printf("(N%sStops = 1 and Stop%sTAZ2 <> null)", {dir, dir})
        n = objT.SelectByQuery({Query: filter, SetName: "__Swap"})
        if n > 0 then do
            vNull = Vector(n, "Long", )
            objT.ChangeSet("__Swap")
            objT.(dir + "StopDeltaTT1") = objT.(dir + "StopDeltaTT2")
            objT.("Stop" + dir + "TAZ1") = objT.("Stop" + dir + "TAZ2")
            objT.("Purpose" + dir + "Stop1") = objT.("Purpose" + dir + "Stop2")
            objT.(dir + "StopDeltaTT2") = vNull
            objT.("Stop" + dir + "TAZ2") = vNull
            objT.("Purpose" + dir + "Stop2") = vNull
            objT.ChangeSet()
        end

        pbar.Step()
    end     // dirs loop
    pbar.Destroy()
    objT = null
    obj4D = null
    Return(ret)
endMacro


/*
    Run Intermediate stops destination choice model given trip direction, stop number, period and purpose 
*/
Macro "Vis Intermediate Stop DC"(Args, spec)
    vwT = spec.ToursView
    dir = spec.Direction
    deltaTT = spec.DeltaTT
    deltaDist = spec.DeltaDist
    type = spec.Type
    ODInfo = spec.ODInfo
    stopNo = spec.StopNo
    skimFile = spec.SkimFile
    seed = spec.RandomSeed

    availExpressions = null
    availExpressions.Alternative = {"Destinations"}
    availExpressions.Expression = {"DeltaTT.DeltaTT <= 45"}

    // Filters
    stopfilter = printf("N%sStops >= %s", {dir, stopNo})             // e.g. NForwardStops >= 1
    typefilter = printf("Purpose%sStop%s = '%s'", {dir, stopNo, type})   // e.g. PurposeForwardStop1 = 'Rec'
    filter = printf("(%s) and (%s)", {stopfilter, typefilter})
    
    SetView(vwT)
    n = SelectByQuery("__Selection", "several", "Select * where " + filter,)
    if n > 0 then do
        if type = "Rec" then 
            clusterField = "VisitorClusterRec"
        else
            clusterField = "VisitorClusterOther"
        outFld = "Stop" + dir + "TAZ" + stopNo

        util = Args.(type + "LocUtilityVis")
        clusters = Args.(type + "LocClustersVis")
        
        TAZDB = Args.TAZGeography
        TAZBin = Substitute(TAZDB, ".dbd", ".bin",) 

        // Called Nested DC procedure
        Opts = null
        Opts.output_dir = Args.OutputFolder + "\\Intermediate"
        Opts.trip_type = "VisitorStopsLocChoice" + type
        Opts.zone_utils = util
        Opts.cluster_data = clusters
        Opts.primary_spec = {Name: "VisitorData", Filter: filter, OField: "Origin"}
        Opts.dc_spec = {DestinationsSource: "AutoSkim", DestinationsIndex: "InternalTAZ"}
        Opts.cluster_equiv_spec = {File: TAZBin, ZoneIDField: "TAZID", ClusterIDField: clusterField}
        Opts.tables = {TAZData: {File: Args.DemographicOutputs, IDField: "TAZ"},
                        TAZBin: {File: TAZBin, IDField: "TAZID"},
                        TAZAccessibilities: {File: Args.AccessibilitiesOutputs, IDField: "TAZID"},
                        VisitorData: {View: vwT, IDField: "TourID"}}
        Opts.matrices = {AutoSkim: {File: skimFile, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"},
                         Intrazonal: {File: Args.IZMatrix, RowIndex: "TAZ", ColIndex: "TAZ"},
                         DeltaTT: {File: deltaTT, PersonBased: 1}}
        Opts.zone_availabilities = availExpressions
        Opts.random_seed = seed
        objNested = CreateObject("NestedDC", Opts)
        objNested.Run()

        // Copy DC choices
        choiceFile = printf("%s\\Intermediate\\choices\\VisitorStopsLocChoice%s_DC_Choices.bin", {Args.OutputFolder, type})
        fopts = {PrimaryView: vwT,
                PrimaryViewID: "TourID",
                PersonChoicesField: outFld,
                ChoicesFile: choiceFile,
                IDField: "[_PersonID]",
                ChoiceField: "[_DestZoneID]"
                }
        RunMacro("Copy DC Choices", fopts)
    end

    Return(1)
endmacro


// Macro that removes the intermediate stop (choice, duration, destination info) if not feasible.
// This happens because the stop may not be feasible given the mode (such as "Walk")
Macro "Remove Infeasible Vis Stops"(opt)
    obj = opt.TableObject
    dir = opt.Direction
    stopNo = opt.StopNo
    n = obj.SelectByQuery({Query: opt.Filter, SetName: "__Remove"})

    if n > 0 then do
        obj.ChangeSet("__Remove")
        vNum = obj.("N" + dir + "Stops")
        AppendToLogFile(1, opt.Message + String(n) + " Intermediate " + dir + " stops removed due to schedule constraints.")
        v = Vector(n, "Long",)
        vS = Vector(n, "String",)
        vecsSet = null
        vecsSet.("N" + dir + "Stops") = vNum - 1
        vecsSet.("Stop" + dir + "TAZ" + stopNo) = v
        vecsSet.(dir + "StopDurChoice" + stopNo) = vS
        vecsSet.(dir + "StopDuration" + stopNo) = v
        vecsSet.(dir + "StopDeltaTT" + stopNo) = v
        vecsSet.("Purpose" + dir + "Stop" + stopNo) = vS
        vStopsChoice = obj.StopsChoice
        if dir = 'Forward' then
            vecsSet.StopsChoice =  i2s(s2i(Left(vStopsChoice, 1)) - 1) + "_" + Right(vStopsChoice, 1)
        else
            vecsSet.StopsChoice = Left(vStopsChoice, 1) + "_" + i2s(s2i(Right(vStopsChoice, 1)) - 1)
        obj.SetDataVectors({FieldData: vecsSet})        
    end
    obj.ChangeSet()
endMacro



Macro "Visitor Stop Scheduling"(Args)
    on error do
        ShowMessage(GetLastError())
        return(0)
    end
    Args.ABMFlag = 2

    objT = CreateObject("Table", Args.VisitorTours)
    
    // Update prev and next tour info
    RunMacro("Update Prev and Next Tour Info", objT)

    // Determine tour order
    RunMacro("Determine Tour Order", objT)

    // Determine stop order and exchange stops 1 and 2 info as desired
    // Stop closest to trip origin is stop 1, next closest is stop 2
    RunMacro("Determine Stop Order", objT, Args.HighwaySkimAM)

    maxTours = 1
    dirs = {"Forward", "Return"}
    maxStops = 2
    for i = 1 to maxTours do
        for dir in dirs do
            for s = 1 to maxStops do
                // Schedule intermediate stops
                spec = {Tours: objT, Direction: dir, StopNo: s, TourBuffer: 15}
                RunMacro("Schedule Visitor Stops", Args, spec)
            end
        end
        // Update prev and next tour info
        //RunMacro("Update Prev and Next Tour Info", objT)
    end

    Return(1)
endMacro


Macro "Determine Tour Order"(objT)
    // Add tour order field
    flds = {{FieldName: "TourOrder", Type: "Short"}}
    objT.AddFields({Fields: flds})
    
    // Get # tours by HHID
    agg = objT.Aggregate({GroupBy: {"HHID"}, 
                          FieldStats: {HHID: {"count"}}
                          })

    SortOrder = {{"HHID", "Ascending"}}
    agg.Sort({FieldArray: SortOrder})
    arrTours = v2a(agg.count_HHID)

    mapArray = {{1}, {1,2}, {1,2,3}, {1,2,3,4}} // Max 4 tours
    // e.g. given {3, 4, 1, 2, 3}
    // return {{1,2,3}, {1,2,3,4}, {1}, {1,2}, {1,2,3}}
    a = arrTours.Map(do (f) Return(mapArray[f]) end)
    vOrder = a2v(a.flatten())

    // Set tour order
    SortOrder = {{"HHID", "Ascending"}, {"TourStartTime", "Ascending"}}
    objT.Sort({FieldArray: SortOrder})
    objT.TourOrder = vOrder
endMacro


// Updates and fills previous and next tour info
Macro "Update Prev and Next Tour Info"(objT)
    objT.ChangeSet()
    SortOrder = {{"HHID", "Ascending"}, {"TourStartTime", "Ascending"}}
    objT.Sort({FieldArray: SortOrder})
    vecs = objT.GetDataVectors({FieldNames: {"HHID", "TourStartTime", "TourEndTime"}})
    
    vPrevHHID = RunMacro("Shift Vector", {Vector: vecs.HHID, Method: "Prev"})
    vPrevTourEndTime = RunMacro("Shift Vector", {Vector: vecs.TourEndTime, Method: "Prev"})
    vNextHHID = RunMacro("Shift Vector", {Vector: vecs.HHID, Method: "Next"})
    vNextTourStartTime = RunMacro("Shift Vector", {Vector: vecs.TourStartTime, Method: "Next"})
    vPrevTourEndTime = if vPrevHHID = vecs.HHID then vPrevTourEndTime else null
    vNextTourStartTime = if vNextHHID = vecs.HHID then vNextTourStartTime else null

    vecsSet = null
    vecsSet.PrevHHID = vPrevHHID
    vecsSet.PrevTourEndTime = vPrevTourEndTime
    vecsSet.NextHHID = vNextHHID
    vecsSet.NextTourStartTime = vNextTourStartTime
    objT.SetDataVectors({FieldData: vecsSet})
endMacro


Macro "Determine Stop Order"(objT, skimFile)
    directions = {"Forward", "Return"}
    for dir in directions do
        if dir = "Forward" then
            anchor  = "Origin"
        else
            anchor = "Destination"
        
        baseFilter = printf("N%sStops > 1", {dir})
        n = objT.SelectByQuery({Filter: baseFilter, SetName: "__TwoStops"})
        if n = 0 then
            continue

        // Fill travel times from anchor to each of the two stops
        mObj = CreateObject('Matrix', skimFile)
        mObj.SetIndex({RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
        mc = mObj.Time
        vw = objT.GetView()
        for stopNo in {"1", "2"} do
            dField = printf("Stop%sTAZ%s", {dir, stopNo})
            fillField = printf("TimeToStop%s%s", {dir, stopNo})
            oSpec = GetFieldFullSpec(vw, anchor)
            dSpec = GetFieldFullSpec(vw, dField)
            fSpec = GetFieldFullSpec(vw, fillField)
            FillViewFromMatrix(vw + "|__TwoStops", oSpec, dSpec, {{fSpec, mc}})
        end
        mObj = null

        // Swap stops info as needed
        swapFilter = printf("TimeToStop%s%s > TimeToStop%s%s", {dir, "1", dir, "2"})
        n = objT.SelectByQuery({Filter: baseFilter + " and " + swapFilter, SetName: "__Swap"})
        if n = 0 then
            continue

        // Swap stop info
        objT.ChangeSet("__Swap")
        flds = {"Stop" + dir + "TAZ1", "Stop" + dir + "TAZ2", 
                "TimeToStop" + dir + "1", "TimeToStop" + dir + "2", dir + "StopDeltaTT1", dir + "StopDeltaTT2", 
                dir + "StopDurChoice1", dir + "StopDurChoice2", dir + "StopDuration1", dir + "StopDuration2",
                "Purpose" + dir + "Stop1", "Purpose" + dir + "Stop2"
                }
        vecs = objT.GetDataVectors({FieldNames: flds})
        
        vecsSet = null
        vecsSet.("Stop" + dir + "TAZ1") = vecs.("Stop" + dir + "TAZ2")
        vecsSet.("Stop" + dir + "TAZ2") = vecs.("Stop" + dir + "TAZ1")
        vecsSet.("TimeToStop" + dir + "1") = vecs.("TimeToStop" + dir + "2")
        vecsSet.("TimeToStop" + dir + "2") = vecs.("TimeToStop" + dir + "1")
        vecsSet.(dir + "StopDeltaTT1") = vecs.(dir + "StopDeltaTT2")
        vecsSet.(dir + "StopDeltaTT2") = vecs.(dir + "StopDeltaTT1")
        vecsSet.(dir + "StopDurChoice1") = vecs.(dir + "StopDurChoice2")
        vecsSet.(dir + "StopDurChoice2") = vecs.(dir + "StopDurChoice1")
        vecsSet.(dir + "StopDuration1") = vecs.(dir + "StopDuration2")
        vecsSet.(dir + "StopDuration2") = vecs.(dir + "StopDuration1")
        vecsSet.("Purpose" + dir + "Stop1") = vecs.("Purpose" + dir + "Stop2")
        vecsSet.("Purpose" + dir + "Stop2") = vecs.("Purpose" + dir + "Stop1")
        objT.SetDataVectors({FieldData: vecsSet})
    end // of dir loop
endMacro


Macro "Schedule Visitor Stops"(Args, spec)
    if stopNo = "1" then
        RunMacro("Visitor Single Stop Scheduling", Args, spec)
    else
        RunMacro("Visitor Multi Stop Scheduling", Args, spec)
endMacro


Macro "Visitor Single Stop Scheduling"(Args, spec)
    tourBuffer = spec.TourBuffer
    dir = spec.Direction
    objT = spec.Tours

    qry = printf("N%sStops = 1", {dir})
    n = objT.SelectByQuery({Filter: qry, SetName: "__OneStop"})

    flds = {"PrevTourEndTime", "NextTourStartTime", "TourStartTime", "TourEndTime", 
            dir + "StopDeltaTT" + stopNo, dir + "StopDuration" + stopNo, "Purpose" + dir + "Stop" + stopNo}
    vecs = objT.GetDataVectors({FieldNames: flds})
    
    // Determine makeup time as the sum of the stop duration and the computed detour travel time
    vMakeUp = vecs.("StopDuration" + stopNo) + vecs.(dir + "StopDeltaTT" + stopNo)
    if dir = "Forward" then do
        anchor  = "Origin"
        vNewDep = vecs.TourStartTime - vMakeUp
        vLost = if (vecs.PrevTourEndTime = null) or (vecs.PrevTourEndTime + tourBuffer <= vNewDep) then 0 else (vecs.PrevTourEndTime - vNewDep + 15)
    end
    else do // Return: Check with subsequent tour
        anchor = "Destination"
        vNewArr = vecs.TourEndTime + vMakeUp
        vLost = if (vecs.NextTourStartTime = null) or (vecs.NextTourStartTime >= vNewArr + tourBuffer) then 0 else vNewArr - vecs.NextTourStartTime + 15
    end
    vNewStopDur = vecs.(dir + "StopDuration" + stopNo) - vLost
    vPurp = vecs.("Purpose" + dir + "Stop" + stopNo)
    vMinDur = if vPurp = "Rec" then 30 else 15
    vRemove = if vNewStopDur < vMinDur then 1 else 0

    // Finalize
    vecsSet = null
    if dir = "Forward" then
        vecsSet.TourStartTime = vNewDep
    else
        vecsSet.TourEndTime = vNewArr
    vecsSet.(dir + "StopDuration" + stopNo) = vNewStopDur
    vecsSet.("Remove" + dir + "Stop") = vRemove
    objT.SetDataVectors({FieldData: vecsSet})

    // Remove infeasible stops
    filter = "Remove" + dir + "Stop = 1"
    opt = {TableObject: objT, Filter: filter, Direction: dir, StopNo: "1"}
    RunMacro("Remove Infeasible Vis Stops", opt)
endMacro


Macro "Visitor Multi Stop Scheduling"(Args, spec)
    tourBuffer = spec.TourBuffer
    dir = spec.Direction
    objT = spec.Tours

    qry = printf("N%sStops = 2", {dir})
    n = objT.SelectByQuery({Filter: qry, SetName: "__TwoStops"})

    // Fill realized travel time between stops 1 and 2
    optF = {View: objT.GetView(), 
            OField: vecsSet.("Stop" + dir + "TAZ1"), DField: vecsSet.("Stop" + dir + "TAZ2"), 
            DepTimeField: "ActivityStartTime", ModeField: "Mode", 
            FillField: dir + "Stop1to2Time"}
    RunMacro("Fill Travel Times", Args, optF)

    flds = {"PrevTourEndTime", "NextTourStartTime", "TourStartTime", "TourEndTime", 
            dir + "StopDeltaTT1", dir + "StopDeltaTT2",
            dir + "StopDuration1", dir + "StopDuration2",
            "Purpose" + dir + "Stop1", "Purpose" + dir + "Stop2", dir + "Stop1to2Time"}
    vecs = objT.GetDataVectors({FieldNames: flds})
    
    // Determine makeup time as the sum of the stop duration and the computed detour travel time
    vMakeUp = vecs.("StopDuration1") + vecs.(dir + "StopDeltaTT1") + vecs.("StopDuration2") + vecs.(dir + "Stop1to2Time")
    if dir = "Forward" then do
        vNewDep = vecs.TourStartTime - vMakeUp
        vLost = if (vecs.PrevTourEndTime = null) or (vecs.PrevTourEndTime + tourBuffer <= vNewDep) then 0 else (vecs.PrevTourEndTime - vNewDep + 15)
    end
    else do // Return: Check with subsequent tour
        vNewArr = vecs.TourEndTime + vMakeUp
        vLost = if (vecs.NextTourStartTime = null) or (vecs.NextTourStartTime >= vNewArr + tourBuffer) then 0 else vNewArr - vecs.NextTourStartTime + 15
    end
    vNewStopDur = vecs.(dir + "StopDuration" + stopNo) - vLost
    vPurp = vecs.("Purpose" + dir + "Stop" + stopNo)
    vMinDur = if vPurp = "Rec" then 30 else 15
    vRemove = if vNewStopDur < vMinDur then 1 else 0

    // Finalize
    vecsSet = null
    if dir = "Forward" then
        vecsSet.TourStartTime = vNewDep
    else
        vecsSet.TourEndTime = vNewArr
    vecsSet.(dir + "StopDuration" + stopNo) = vNewStopDur
    vecsSet.("Remove" + dir + "Stop") = vRemove
    objT.SetDataVectors({FieldData: vecsSet})

    // Remove infeasible stops
    filter = "Remove" + dir + "Stop = 1"
    opt = {TableObject: objT, Filter: filter, Direction: dir, StopNo: "1"}
    RunMacro("Remove Infeasible Vis Stops", opt)
endMacro
