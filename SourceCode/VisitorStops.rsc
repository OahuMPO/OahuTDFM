Macro "Visitor Stops Setup"(Args)
    on error do
        ShowMessage(GetLastError())
        return(0)
    end

    Args.ABMFlag = 2

    objT = CreateObject("Table", Args.VisitorTours)
    flds = {{FieldName: "StopsChoice", Type: "String", Width: 5},
            {FieldName: "NForwardStops", Type: "Short"},
            {FieldName: "NReturnStops", Type: "Short"}}
    dirs = {"Forward", "Return"}
    tourNos = {"1", "2"}
    for dir in dirs do
        for t in tourNos do
            flds = flds + {{FieldName: "Stop" + dir + "TAZ" + t, Type: "Integer"},
                           {FieldName: dir + "StopDurChoice" + t, Type: "String", Width: 15},
                           {FieldName: dir + "StopDuration" + t, Type: "Real"},
                           {FieldName: dir + "StopDeltaTT" + t, Type: "Real"}}
        end
    end
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


Macro "Visitor Stops Destination"(Args)
    // Run Destination Choice
    dirs = {"Forward", "Return"}
    types = {"Rec", "Shop", "Other"}
    stopsArr = {"1", "2"}
    
    tourFile = Args.VisitorTours
    objT = CreateObject("Table", tourFile)
    spec = {ToursView: objT.GetView()}

    spec = {ToursObject: objT, ToursView: objT.GetView()}
    pbar = CreateObject("G30 Progress Bar", "Intermediate Stops Destinations: (Forward, Return) and (Rec, Shop, Other)", false, 7)
    for dir in dirs do
        // Get Stop Formula Fields
        spec.Direction = dir
        if dir = "Forward" then
            spec.ODInfo = {Origin: "Origin", Destination: "Destination"}
        else
            spec.ODInfo = {Origin: "Destination", Destination: "Origin"}

        spec.StopFilter = printf("N%sStops >= 1", {dir}) // e.g. NForwardStops >= 1
        spec.SkimFile = Args.HighwaySkimAM

        // Calculate delta TT matrix
        deltaTT = GetTempPath() + "DeltaTT_" + dir + ".mtx"
        spec.DeltaSkim = deltaTT
        spec.MatrixSpec = {File: skimFile, Core: "Time", RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"}
        spec.OutputCoreName = "DeltaTT"
        ret = RunMacro("Calculate Delta TT", Args, spec)
        if ret = 2 then continue // No records for delta TT calculation. Move on to next direction.
        spec.DeltaTT = deltaTT

        // Calculate delta Dist matrix
        /*
        deltaDist = GetTempPath() + "DeltaDist_" + dir + ".mtx"
        spec.DeltaSkim = deltaDist
        spec.MatrixSpec = {File: skimFile, Core: "Distance", RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"}
        spec.OutputCoreName = "DeltaDist"
        ret = RunMacro("Calculate Delta TT", Args, spec)
        if ret = 2 then continue // No records for delta TT calculation. Move on to next direction.
        spec.DeltaDist = deltaDist
        */
        
        // Purposes Loop
        for type in types do
            spec.Type = type
            for s in stopsArr.(type) do
                spec.StopNo = s
                spec.RandomSeed = 3999971 + 100*periods.position(period) + 10*types.position(type) + s2i(s)
                RunMacro("Vis Intermediate Stop DC", Args, spec)
                pbar.Step()
            end
        end

        // Now that destinations have been computed, fill in the realized detour travel times
        if dir = "Forward" then 
            depFld = 'TourStartTime'
        else 
            depFld = 'ActivityEndTime'
        
        for stopNo in stopsArr.Work do
            opt = {ToursObj: objT, Filter: printf("N%sStops >= %s", {dir, stopNo}), ODInfo: ODInfo, 
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
            objT.(dir + "StopDeltaTT2") = vNull
            objT.("Stop" + dir + "TAZ2") = vNull
        end

        pbar.Step()
    end     // dirs loop
    pbar.Destroy()
    objT = null
    obj4D = null
    Return(true)
endMacro


/*
    Run Intermediate stops destination choice model given trip direction, stop number, period and purpose 
*/
Macro "Vis Intermediate Stop DC"(Args, spec)
    vwT = spec.ToursView
    dir = spec.Direction
    period = spec.Period
    deltaTT = spec.DeltaTT
    deltaDist = spec.DeltaDist
    type = spec.Type
    ODInfo = spec.ODInfo
    stopNo = spec.StopNo
    skimFile = spec.SkimFile
    seed = spec.RandomSeed

    availExpressions = null
    availExpressions.Alternative = {"Destinations"}
    availExpressions.Expression = {"deltaTT.deltaTT <= 45"}

    // Filters
    stopfilter = printf("N%sStops >= %s", {dir, stopNo})             // e.g. NForwardStops >= 1
    typefilter = printf("TourPurpose = '%s'", {type})               // e.g. TourPurpose = 'Work'
    filter = printf("(%s) and (%s)", {stopfilter, typefilter})
    
    SetView(vwT)
    n = SelectByQuery("__Selection", "several", "Select * where " + filter,)
    if n > 0 then do
        filter = spec.Filter
        if type = "Rec" then 
            clusterField = "VisitorClusterRec"
        else
            clusterField = "VisitorClusterOther"
        outFld = "Stops" + dir + "TAZ" + stopNo

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
        Opts.primary_spec = {Name: "VisitorData", Filter: filter, OField: "LodgingTAZ"}
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
        vStopsChoice = obj.StopsChoice
        if dir = 'Forward' then
            vecsSet.StopsChoice =  i2s(s2i(Left(vStopsChoice, 1)) - 1) + "_" + Right(vStopsChoice, 1)
        else
            vecsSet.StopsChoice = Left(vStopsChoice, 1) + "_" + i2s(s2i(Right(vStopsChoice, 1)) - 1)
        obj.SetDataVectors({FieldData: vecsSet})        
    end
    obj.ChangeSet()
endMacro
