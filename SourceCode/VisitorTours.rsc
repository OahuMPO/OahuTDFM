Macro "Visitor Tour Frequencies"(Args)
    on error do
        ShowMessage(GetLastError())
        return(0)
    end

    Args.ABMFlag = 2    // Indicate that the visitor abm step is active

    // Work Tour frequency
    spec = {Purpose: "Work",
            Filter: "PurposeCat = 2"}
    ret = RunMacro("Run Visitor Tour Freq", Args, spec)

    // Rec Tour frequency
    spec = {Purpose: "Rec",
            Filter: "HouseholdID > 0"}
    ret = RunMacro("Run Visitor Tour Freq", Args, spec)

    // Other Tour frequency
    spec = {Purpose: "Other",
            Filter: "HouseholdID > 0"}
    ret = RunMacro("Run Visitor Tour Freq", Args, spec)

    // Shop Tour frequency
    spec = {Purpose: "Shop",
            Filter: "HouseholdID > 0"}
    ret = RunMacro("Run Visitor Tour Freq", Args, spec)

    Return(ret)
endMacro


Macro "Run Visitor Tour Freq"(Args, spec)
    visabm = RunMacro("Get Visitor ABM Manager", Args)
    purp = spec.Purpose
    filter = spec.Filter
    
    utilFile = Args.("Visitor" + purp + "TourFreqUtility")
    outFld = "Number" + purp + "Tours"

    TAZDB = Args.TAZGeography
    TAZBin = Substitute(TAZDB, ".dbd", ".bin",) 
    objT = CreateObject("Table", TAZBin)
    objA = CreateObject("Table", Args.AccessibilitiesOutputs)

    // Run Model and populate results
    obj = CreateObject("PMEChoiceModel", {ModelName: purp + " Visitor Tour Frequency"})
    obj.OutputModelFile = printf("%s\\Intermediate\\Visitor%sTourFreq.mdl", {Args.[Output Folder], purp})
    obj.AddTableSource({SourceName: "VisitorData", View: visabm.HHView, IDField: visabm.HHID})
    obj.AddTableSource({SourceName: "TAZBin", View: objT.GetView(), IDField: "TAZID"})
    obj.AddTableSource({SourceName: "TAZAccessibilities", View: objA.GetView(), IDField: "TAZID"})
    obj.AddPrimarySpec({Name: "VisitorData", Filter: filter, OField: "LodgingTAZ"})
    obj.AddUtility({UtilityFunction: utilFile})
    obj.AddOutputSpec({ChoicesField: outFld})
    obj.ReportShares = 1
    obj.RandomSeed = 99991
    ret = obj.Evaluate()
    if !ret then
        Throw("Visitor tour frequency model failed for purpose: " + purp)
    Args.(purp + "VisitorTourFreq Spec") = CopyArray(ret) // For calibration purposes
    
    // Subtract 1 to convert alternatives to number of tours
    outFld = "HH." + outFld
    visabm.(outFld) = visabm.(outFld) - 1

    objT = null
    objA = null
    Return(1)
endMacro


Macro "Visitor Tour Destinations"(Args)
    on error do
        ShowMessage(GetLastError())
        return(0)
    end

    Args.ABMFlag = 2    // Indicate that the visitor abm step is active

    availExpressions = null
    availExpressions.Alternative = {"Destinations"}
    availExpressions.Expression = {"TAZData.TotalEmployment.D > 0"}

    // Work Tour
    spec = {Purpose: "Work",
            Filter: "NumberWorkTours = 1",
            ClusterField: "VisitorClusterOther",
            OutputField: "WorkTAZ1",
            ZoneAvailabilities: availExpressions,
            Seed: 899981}
    ret = RunMacro("Run Visitor Tour DC", Args, spec)

    nMaxTours = 2    // Two tours possible for Rec, Other, and Shop
    for i = 1 to nMaxTours do
        // Recreation Tour DC
        availExpressions.Expression = {"(nz(TAZData.VisArrivalsRec.D) + TAZData.Emp_Hotel.D + TAZData.Emp_Retail.D + TAZData.Emp_Services.D > 0) or (TAZBin.VisitorClusterRec.D <= 5)"}
        spec = {Purpose: "Rec",
                Filter: "NumberRecTours >= " + String(i),
                ClusterField: "VisitorClusterRec",
                OutputField: "RecTAZ" + String(i),
                ZoneAvailabilities: availExpressions,
                Seed: 899981 + 1000 + i}
        ret = RunMacro("Run Visitor Tour DC", Args, spec)

        // Other Tour DC
        availExpressions.Expression = {"TAZData.Emp_Public.D + TAZData.Emp_Hotel.D + TAZData.Emp_Services.D > 0"}
        spec = {Purpose: "Other",
                Filter: "NumberOtherTours >= " + String(i),
                ClusterField: "VisitorClusterOther",
                OutputField: "OtherTAZ" + String(i),
                ZoneAvailabilities: availExpressions,
                Seed: 899981 + 2000 + i}
        ret = RunMacro("Run Visitor Tour DC", Args, spec)

        // Shop Tour DC
        availExpressions.Expression = {"TAZData.Emp_Retail.D + TAZData.Emp_Hotel.D + TAZData.Emp_Services.D > 0"}
        spec = {Purpose: "Shop",
                Filter: "NumberShopTours >= " + String(i),
                ClusterField: "VisitorClusterOther",
                OutputField: "ShopTAZ" + String(i),
                ZoneAvailabilities: availExpressions,
                Seed: 899981 + 3000 + i}
        ret = RunMacro("Run Visitor Tour DC", Args, spec)
    end

    Return(ret)
endMacro


Macro "Run Visitor Tour DC"(Args, spec)
    purp = spec.Purpose
    filter = spec.Filter
    clusterField = spec.ClusterField
    outFld = spec.OutputField
    avails = spec.ZoneAvailabilities
    seed = spec.Seed

    util = Args.(purp + "LocUtilityVis")
    clusters = Args.(purp + "LocClustersVis")
    
    visabm = RunMacro("Get Visitor ABM Manager", Args)
    TAZDB = Args.TAZGeography
    TAZBin = Substitute(TAZDB, ".dbd", ".bin",) 

    // Called Nested DC procedure
    Opts = null
    Opts.output_dir = Args.OutputFolder + "\\Intermediate"
    Opts.trip_type = "VisitorLocChoice" + purp
    Opts.zone_utils = util
    Opts.cluster_data = clusters
    Opts.primary_spec = {Name: "VisitorData", Filter: filter, OField: "LodgingTAZ"}
    Opts.dc_spec = {DestinationsSource: "AutoSkim", DestinationsIndex: "InternalTAZ"}
    Opts.cluster_equiv_spec = {File: TAZBin, ZoneIDField: "TAZID", ClusterIDField: clusterField}
    Opts.tables = {TAZData: {File: Args.DemographicOutputs, IDField: "TAZ"},
                    TAZBin: {File: TAZBin, IDField: "TAZID"},
                    TAZAccessibilities: {File: Args.AccessibilitiesOutputs, IDField: "TAZID"},
                    VisitorData: {View: visabm.HHView, IDField: visabm.HHID}}
    Opts.matrices = {AutoSkim: {File: Args.HighwaySkimAM, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"},
                     Intrazonal: {File: Args.IZMatrix, RowIndex: "TAZ", ColIndex: "TAZ"}}
    Opts.zone_availabilities = avails
    Opts.random_seed = seed
    objNested = CreateObject("NestedDC", Opts)
    objNested.Run()

    // Copy DC choices
    choiceFile = Args.OutputFolder + "\\Intermediate\\choices\\VisitorLodgingChoice_DC_Choices.bin"
    choiceFile = printf("%s\\Intermediate\\choices\\VisitorLocChoice%s_DC_Choices.bin", {Args.OutputFolder, purp})
    fopts = {PrimaryView: visabm.HHView,
            PrimaryViewID: visabm.HHID,
            PersonChoicesField: outFld,
            ChoicesFile: choiceFile,
            IDField: "[_PersonID]",
            ChoiceField: "[_DestZoneID]"
            }
    RunMacro("Copy DC Choices", fopts)

    Return(1)
endmacro


Macro "Visitor Tour Modes"(Args)
    on error do
        ShowMessage(GetLastError())
        return(0)
    end

    Args.ABMFlag = 2    // Indicate that the visitor abm step is active

    // Work Tour
    spec = {Purpose: "Work",
            Filter: "NumberWorkTours = 1",
            OutputField: "WorkMode1",
            DestField: "WorkTAZ1",
            Seed: 899981}
    ret = RunMacro("Run Visitor Tour MC", Args, spec)

    nMaxTours = 2    // Two tours possible for Rec, Other, and Shop
    for i = 1 to nMaxTours do
        // Recreation Tour MC
        spec = {Purpose: "Rec",
                Filter: "NumberRecTours >= " + String(i),
                OutputField: "RecMode" + String(i),
                DestField: "RecTAZ" + String(i),
                Seed: 899981 + 1000 + i}
        ret = RunMacro("Run Visitor Tour MC", Args, spec)

        // Other Tour MC
        spec = {Purpose: "Other",
                Filter: "NumberOtherTours >= " + String(i),
                OutputField: "OtherMode" + String(i),
                DestField: "OtherTAZ" + String(i),
                Seed: 899981 + 2000 + i}
        ret = RunMacro("Run Visitor Tour MC", Args, spec)

        // Shop Tour MC
        spec = {Purpose: "Shop",
                Filter: "NumberShopTours >= " + String(i),
                OutputField: "ShopMode" + String(i),
                DestField: "ShopTAZ" + String(i),
                Seed: 899981 + 3000 + i}
        ret = RunMacro("Run Visitor Tour MC", Args, spec)
    end

    Return(ret)
endMacro


Macro "Run Visitor Tour MC"(Args, spec)
    purp = spec.Purpose
    filter = spec.Filter
    outFld = spec.OutputField
    seed = spec.Seed
    dFld = spec.DestField

    util = Args.(purp + "VisitorMCUtility")
    
    availExpressions = null
    availExpressions.Alternative = {"Walk"}
    availExpressions.Expression = {"WalkSkim.Distance < 1.5"}

    w_t_skim = Args.[Output Folder] + "\\skims\\transit\\AM_w_bus.mtx"

    visabm = RunMacro("Get Visitor ABM Manager", Args)
    TAZDB = Args.TAZGeography
    TAZBin = Substitute(TAZDB, ".dbd", ".bin",)
    
    objT = CreateObject("Table", TAZBin)
    objA = CreateObject("Table", Args.AccessibilitiesOutputs)
    objD = CreateObject("Table", Args.DemographicOutputs)

    // Run Model and populate results
    obj = CreateObject("PMEChoiceModel", {ModelName: purp + " Visitor Tour MC"})
    obj.OutputModelFile = printf("%s\\Intermediate\\Visitor%sTourMC.mdl", {Args.[Output Folder], purp})
    obj.AddTableSource({SourceName: "VisitorData", View: visabm.HHView, IDField: visabm.HHID})
    obj.AddTableSource({SourceName: "TAZBin", View: objT.GetView(), IDField: "TAZID"})
    obj.AddTableSource({SourceName: "TAZAccessibilities", View: objA.GetView(), IDField: "TAZID"})
    obj.AddTableSource({SourceName: "TAZData", View: objD.GetView(), IDField: "TAZ"})
    obj.AddMatrixSource({SourceName: "AutoSkim", File: Args.HighwaySkimAM, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
    obj.AddMatrixSource({SourceName: "WalkSkim", File: Args.WalkSkim, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
    obj.AddMatrixSource({SourceName: "BusSkim", File: w_t_skim, RowIndex: "RCIndex", ColIndex: "RCIndex"})
    obj.AddPrimarySpec({Name: "VisitorData", Filter: filter, OField: "LodgingTAZ", DField: dFld})
    obj.AddUtility({UtilityFunction: util, AvailabilityExpressions: availExpressions})
    obj.AddOutputSpec({ChoicesField: outFld})
    obj.ReportShares = 1
    obj.RandomSeed = seed
    ret = obj.Evaluate()
    if !ret then
        Throw("Visitor tour mode choice model failed for purpose: " + purp)
    Args.(purp + "VisitorTourMC Spec") = CopyArray(ret) // For calibration purposes
    
    objD = null
    objT = null
    objA = null
    Return(1)
endMacro


Macro "Visitor Tour TOD"(Args)
    on error do
        ShowMessage(GetLastError())
        return(0)
    end

    Args.ABMFlag = 2    // Indicate that the visitor abm step is active
    visabm = RunMacro("Get Visitor ABM Manager", Args)
    
    availObj = RunMacro("Create TOD Availability Table", visabm.HHView)
    vwJ = JoinViews("VisitorDataPlusAvail", GetFieldFullSpec(visabm.HHView, "HouseholdID"), GetFieldFullSpec(availObj.GetView(), "HHID"),)
    
    purps = {"Work1", "Rec1", "Shop1", "Other1", "Rec2", "Other2", "Shop2"}
    pbar = CreateObject("G30 Progress Bar", "Running TOD Model", true, purps.length)
    for val in purps do
        tourNo = Right(val, 1)
        p = Left(val, StringLength(val) - 1)
        choiceFld = p + "TOD" + tourNo
        utilTable = Args.(p + "TourTODUtility")
        filter = "Number" + p + "Tours >= " + tourNo

        availOpt = {Utility: utilTable, SourceName: "VisitorData"}
        availExpressions = RunMacro("Get TOD Availability Expressions", availOpt)
        
        // Run TOD
        Opts = {abmManager: visabm,
                PrimaryView: vwJ,
                Type: p,
                ModelName: p + "TourTOD",
                ModelFile: p + "TourTOD.mdl",
                Filter: filter,
                DestField: p + "TAZ" + tourNo,
                ModeField: p + "Mode" + tourNo, 
                Utility: utilTable,
                AvailabilityExpressions: availExpressions,
                ChoiceField: choiceFld,
                RandomSeed: 1499977 + 10000*s2i(tourNo),
                SimulateTimeFields: {StartTime: p + "_StartTime" + tourNo, EndTime: p + "_EndTime" + tourNo, Duration: p + "_Duration" + tourNo},
                MinimumDuration: 10}
        ret = RunMacro("Run Visitor Tour TOD", Args, Opts)

        // Update Avail Table
        spec = {PrimaryView: vwJ, ChoiceField: choiceFld, Filter: filter}
        RunMacro("Update TOD Availability Table", spec)

        if pbar.Step() then
            Return(0)
    end
    CloseView(vwJ)
    availObj = null
    pbar.Destroy()

    Return(ret)
endMacro


// Macro that runs the activity TOD model.
// Each alternative is a combination of activity start time period and activity end time period.
// Called for work tour, univ tour or school tour choices
Macro "Run Visitor Tour TOD"(Args, Opts)
    primaryView = Opts.PrimaryView
    filter = Opts.Filter

    TAZDB = Args.TAZGeography
    TAZBin = Substitute(TAZDB, ".dbd", ".bin",)
    objT = CreateObject("Table", TAZBin)

    // Basic Check
    if Opts.Utility = null or Opts.ModelName = null or Opts.ModelFile = null or filter = null
        or Opts.DestField or Opts.ChoiceField = null then
            Throw("Invalid inputs to macro 'Run Visitor Tour TOD'")

    // Get Utility Options
    utilOpts = {UtilityFunction: Opts.Utility}
    availExpressions = Opts.AvailabilityExpressions
    if availExpressions <> null then
        utilOpts = utilOpts + {AvailabilityExpressions: availExpressions}
    
    utilOpts.SubstituteStrings = {{"<ModeField>", Opts.ModeField}}

    // Run Model and populate results
    obj = CreateObject("PMEChoiceModel", {ModelName: Opts.ModelName})
    obj.OutputModelFile = Args.OutputFolder + "\\Intermediate\\" + Opts.ModelFile
    obj.AddTableSource({SourceName: "VisitorData", View: primaryView, IDField: "HouseholdID"})
    obj.AddTableSource({SourceName: "TAZBin", View: objT.GetView(), IDField: "TAZID"})
    obj.AddMatrixSource({SourceName: "AutoSkim", File: Args.HighwaySkimAM, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
    obj.AddPrimarySpec({Name: "VisitorData", Filter: filter, OField: "LodgingTAZ", DField: Opts.DestField})
    obj.AddUtility(utilOpts)
    obj.AddOutputSpec({ChoicesField: Opts.ChoiceField})
    obj.ReportShares = Args.ReportShares
    obj.RandomSeed = Opts.RandomSeed
    ret = obj.Evaluate()
    if !ret then
        Throw("Model 'Visitor Tour TOD Choice' failed for " + Opts.ModelName)
    Args.(Opts.ModelName +  " Spec") = CopyArray(ret)

    if Opts.SimulateTimeFields <> null then
        RunMacro("Simulate Time Fields", Args, Opts)

    objT = null
    Return(ret)
endMacro


Macro "Simulate Time Fields"(Args, Opts)
    visabm = Opts.abmManager
    type = Opts.Type
    stTimeFld = Opts.SimulateTimeFields.StartTime
    enTimeFld = Opts.SimulateTimeFields.EndTime
    durFld = Opts.SimulateTimeFields.Duration
    todFld = Opts.ChoiceField
    periods = {"EA", "AM", "MD", "PM", "EV", "NT"}

    // Activity Start Time
    for period in periods do
        profileTable = Args.("TODProfile" + period)

        // Select records for this period and using master filter
        filter = printf("(%s) and (Upper(Left(%s,2)) = '%s')", {Opts.Filter, todFld, period})
        set = visabm.CreateHHSet({Filter: filter, Activate: 1, SetName: "__MasterSetPlus" + period})
        if set.Size = 0 then continue

        intervals = profileTable.Interval
        weights = profileTable.(type + "Start")
        opt = {SampleSize: set.Size, Intervals: intervals, Weights: weights, RandomSeed: Opts.RandomSeed*10 + 1}
        vOut = RunMacro("Simulate Time From Interval", opt)
        
        vecsSet = null
        vecsSet.(stTimeFld) = vOut
        visabm.SetHHVectors(vecsSet) 
    end

    // Activity End Time
    for period in periods do
        profileTable = Args.("TODProfile" + period)

        // Select records for this period and using master filter
        filter = printf("(%s) and (Upper(Right(%s,2)) = '%s')", {Opts.Filter, todFld, period})
        set = visabm.CreateHHSet({Filter: filter, Activate: 1, SetName: "__MasterSetPlus" + period})
        if set.Size = 0 then continue

        intervals = profileTable.Interval
        weights = profileTable.(type + "End")
        opt = {SampleSize: set.Size, Intervals: intervals, Weights: weights, RandomSeed: Opts.RandomSeed*10 + 2}
        vOut = RunMacro("Simulate Time From Interval", opt)
        
        vecsSet = null
        vecsSet.(enTimeFld) = vOut
        visabm.SetHHVectors(vecsSet)
    end

    // Fill duration field
    set = visabm.CreateHHSet({Filter: Opts.Filter, Activate: 1, SetName: "__MasterSet"})
    vecs = visabm.GetHHVectors({stTimeFld, enTimeFld, todFld})
    vecsSet = null
    vecsSet.(stTimeFld) = Min(vecs.(stTimeFld), vecs.(enTimeFld))   // Swap Start and End if end time is less than start time
    vecsSet.(enTimeFld) = Max(vecs.(stTimeFld), vecs.(enTimeFld))
    vecsSet.(durFld) = vecsSet.(enTimeFld) - vecsSet.(stTimeFld)
    visabm.SetHHVectors(vecsSet)

    // Impose minimum duration but do not allow activity to encroach other time periods
    minDur = Opts.MinimumDuration
    if minDur > 0 then do
        qry = printf("%s and %s < %u", {Opts.Filter, durFld, minDur})
        set = visabm.CreateHHSet({Filter: qry, Activate: 1, SetName: "__LessThanMinDur"})
        if set.Size = 0 then Return()

        vecs = visabm.GetHHVectors({stTimeFld, enTimeFld, todFld, durFld})
        
        timePeriods = Args.TimePeriods
        vTOD = vecs.(todFld)
        vStartTOD = Left(vTOD, 2)
        vEndTOD = Right(vTOD, 2)
        vStartLB = a2v(v2a(vStartTOD).Map(do (f) Return(timePeriods.(f).StartTime) end))
        vEndUB = a2v(v2a(vEndTOD).Map(do (f) Return(timePeriods.(f).EndTime) end))
        
        vDeltaDur = Max(0, minDur - vecs.(durFld)) // Extra duration needed to meet minimum duration
        
        // Allocate half of the delta duration to start time and half to end time to get new desired start and end times
        vOrigStart = vecs.(stTimeFld)
        vOrigEnd = vecs.(enTimeFld)
        vNewStart = vOrigStart - Floor(vDeltaDur/2)
        vNewEnd = vOrigEnd + Ceil(vDeltaDur/2)
        
        // Make sure the new start/end do not encroach other time periods. If so, allocate the delta duration to the other end
        // Note that due to the minimum duration being less than length of any period, only either start or end will encroach other period
        vecsSet = null
        vecsSet.(stTimeFld) = if vNewStart < vStartLB then vOrigStart
                                else if vNewEnd > vEndUB then vOrigStart - vDeltaDur
                                    else vNewStart
        vecsSet.(enTimeFld) = if vNewStart < vStartLB then vOrigEnd + vDeltaDur 
                                else if vNewEnd > vEndUB then vOrigEnd
                                    else vNewEnd
        vecsSet.(durFld) = vecsSet.(enTimeFld) - vecsSet.(stTimeFld)
        visabm.SetHHVectors(vecsSet)
    end
endMacro


// Create Person by TOD combination availability table
Macro "Create TOD Availability Table"(vw)
    // Create an In-Memory availability table
    vwM = ExportView(vw + "|", "MEM", "AvailTable", {"HouseholdID"},)
    objAv = CreateObject("Table", vwM)
    objAv.RenameField({FieldName: "HouseholdID", NewName: "HHID"})
    
    // Add fields and fill with ones
    periods = {"EA", "AM", "MD", "PM", "EV", "NT"}
    flds = null
    for i = 1 to periods.length do
        for j = i to periods.length do
            flds = flds + {{FieldName: "Avail_" + periods[i] + "_" + periods[j], Type: "Short"}}
        end
    end
    objAv.AddFields({Fields: flds})

    vOne = Vector(objAv.HHID.length, "Short", {Constant: 1})
    vecsSet = null
    for fld in flds do
        vecsSet.(fld.FieldName) = vOne
    end
    objAv.SetDataVectors({FieldData: vecsSet})
    Return(objAv)
endMacro


// Create the availability expressions to use in the PME class
// One expression for each alternative
// Works only for an MNL model with the alternative names in the PME table
Macro "Get TOD Availability Expressions"(availOpt)
    util = availOpt.Utility
    srcName = availOpt.SourceName
    colNames = util.Map(do (f) Return(f[1]) end)
    excludeCols = {"description", "expression", "coefficient", "segment", "filter"}
    availExpressions = null
    for col in colNames do
        if excludeCols.position(Lower(col)) = 0 then do// Add expression
            st = left(col,2)
            rt = right(col,2)
            expr = srcName + ".Avail_" + st + "_" + rt
            availExpressions.Alternative = availExpressions.Alternative + {col}
            availExpressions.Expression = availExpressions.Expression + {expr}
        end
    end
    Return(availExpressions)
endMacro


// Update the availability table based on choice
Macro "Update TOD Availability Table"(spec)
    vw = spec.PrimaryView
    choiceField = spec.ChoiceField
    filter = spec.Filter

    SetView(vw)
    n = SelectByQuery("__Selection", "several", "Select * where " + filter,)
    if n = 0 then 
        Return()
    v = GetDataVector(vw + "|__Selection", choiceField,)
    
    periods = {"EA", "AM", "MD", "PM", "EV", "NT"}
    st = v2a(v).Map(do (f) Return(periods.position(Left(f,2))) end)
    en = v2a(v).Map(do (f) Return(periods.position(Right(f,2))) end)
    vSt = a2v(st)
    vEn = a2v(en)
    
    vecsSet = null
    for i = 1 to periods.length do
        for j = i to periods.length do
            updateFld = "Avail_" + periods[i] + "_" + periods[j]
            vCurrAvail = GetDataVector(vw + "|__Selection", updateFld,)
            vecsSet.(updateFld) = if (vEn < i) or (vSt > j) then vCurrAvail else 0
        end
    end
    SetDataVectors(vw + "|__Selection", vecsSet,)
endMacro


/*
    Adjust tour start and end times if necessary when there are multiple tours using a mimimum buffer between tours
    Find out how much time needs to be made up between tours and adjust the start and end times accordingly
    Change TOD choice if necessary
*/
Macro "Adjust TOD Schedules"(Args)
    abm = RunMacro("Get ABM Manager", Args)

    // Two work tours
    spec = {abmManager: abm, 
            Filter: "NumberWorkTours > 1", 
            Tours: {"Work1", "Work2"},
            PeriodInfo: Args.TimePeriods}
    RunMacro("Adjust TOD Schedule", spec)

    // Two univ tours
    spec = {abmManager: abm, 
            Filter: "NumberUnivTours > 1", 
            Tours: {"Univ1", "Univ2"},
            PeriodInfo: Args.TimePeriods}
    RunMacro("Adjust TOD Schedule", spec)

    // Work and school tours
    spec = {abmManager: abm, 
            Filter: "NumberWorkTours > 0 and AttendSchool = 1 and SchoolTAZ <> null", 
            Tours: {"Work1", "School"}, 
            PeriodInfo: Args.TimePeriods}
    RunMacro("Adjust TOD Schedule", spec)

    // Work and univ tours
    spec = {abmManager: abm, 
            Filter: "NumberWorkTours > 0 and NumberUnivTours > 0", 
            Tours: {"Work1", "Univ1"}, 
            PeriodInfo: Args.TimePeriods}
    RunMacro("Adjust TOD Schedule", spec)

    Return(true)
endMacro


Macro "Adjust TOD Schedule"(spec)
    abm = spec.abmManager
    filter = spec.Filter
    tourA = spec.Tours[1]
    tourB = spec.Tours[2]
    buffer = 15
    typeA = if tourA = "School" then "School" else Left(tourA, 4)
    typeB = if tourB = "School" then "School" else Left(tourB, 4)

    // Get the fields for the tours
    ActStartFldTourA = tourA + "_StartTime"
    ActEndFldTourA = tourA + "_EndTime"
    ActTODTourA = tourA + "_ActivityTOD"
    TourAToHomeFld = typeA + "ToHomeTime"
    HomeToTourAFld = "HomeTo" + typeA + "Time"

    ActStartFldTourB = tourB + "_StartTime"
    ActEndFldTourB = tourB + "_EndTime"
    ActTODTourB = tourB + "_ActivityTOD"
    TourBToHomeFld = typeB + "ToHomeTime"
    HomeToTourBFld = "HomeTo" + typeB + "Time"

    // Get the appropriate vectors
    set = abm.CreatePersonSet({Filter: filter, SetName: "__MultipleTours", Activate: 1})
    if set.Size = 0 then
        Return()
    
    flds = {ActStartFldTourA, ActEndFldTourA, ActStartFldTourB, ActEndFldTourB, 
            TourAToHomeFld, HomeToTourAFld, TourBToHomeFld, HomeToTourBFld}
    vecs = abm.GetPersonVectors(flds)
    
    vStA = vecs.(ActStartFldTourA)
    vStB = vecs.(ActStartFldTourB)
    vEnA = vecs.(ActEndFldTourA)
    vEnB = vecs.(ActEndFldTourB)

    // Safety check to ensure that there are no missing return time values
    vH2ATime = vecs.(HomeToTourAFld)
    vA2HTime = vecs.(TourAToHomeFld)
    vA2HTime = if vA2HTime = null then vH2ATime else vA2HTime

    vH2BTime = vecs.(HomeToTourBFld)
    vB2HTime = vecs.(TourBToHomeFld)
    vB2HTime = if vB2HTime = null then vH2BTime else vB2HTime
    
    // Find earliest start time for tour A/B and the makeup time
    vEarliestSt = if vStA < vStB then 
                    vEnA + vA2HTime + vH2BTime + buffer
                  else 
                    vEnB + vB2HTime + vH2ATime + buffer

    vMakeUp = if vStA < vStB then Ceil(Max(vEarliestSt - vStB, 0)) else Ceil(Max(vEarliestSt - vStA, 0))
    vMakeUp1 = Floor(vMakeUp/2)
    vMakeUp2 = Ceil(vMakeUp/2)

    // Adjust the start and end times for tours A and B
    vStA =  if vStA < vStB then vStA - vMakeUp1 else vStA + vMakeUp2
    vEnA =  if vStA < vStB then vEnA - vMakeUp1 else vEnA + vMakeUp2
    vStB =  if vStA < vStB then vStB + vMakeUp2 else vStB - vMakeUp1
    vEnB =  if vStA < vStB then vEnB + vMakeUp2 else vEnB - vMakeUp1
    
    // Compute updated Activity TOD in the case the trip has spilled over to the next period
    periodInfo = spec.PeriodInfo
    vTODStA = RunMacro("Get Time Period Vector", vStA, periodInfo)
    vTODEnA = RunMacro("Get Time Period Vector", vEnA, periodInfo)
    vTODStB = RunMacro("Get Time Period Vector", vStB, periodInfo)
    vTODEnB = RunMacro("Get Time Period Vector", vEnB, periodInfo)
    
    // Set the appropriate vectors
    vecsSet = null
    vecsSet.(ActStartFldTourA) = vStA
    vecsSet.(ActEndFldTourA) = vEnA
    vecsSet.(ActStartFldTourB) = vStB
    vecsSet.(ActEndFldTourB) = vEnB
    vecsSet.(ActTODTourA) = vTODStA + "-" + vTODEnA
    vecsSet.(ActTODTourB) = vTODStB + "-" + vTODEnB
    abm.SetPersonVectors(vecsSet)
endMacro
