Macro "SubTour Setup"(Args)
    // Add fields to mandatory tours table
    obj = CreateObject('Table', Args.MandatoryTours)
    flds = {{FieldName: "SubTour", Type: "Short"},
            {FieldName: "SubTourTAZ", Type: "Integer"},
            {FieldName: "SubTourActStartInt", Type: "String", Width: 10},
            {FieldName: "SubTourActStartTime", Type: "Integer"},
            {FieldName: "SubTourActDurInt", Type: "String", Width: 12},
            {FieldName: "SubTourActDuration", Type: "Integer"},
            {FieldName: "SubTourMode", Type: "String", Width: 15},
            {FieldName: "SubTourModeCode", Type: "Short"},
            {FieldName: "SubTourStartTime", Type: "Integer"},
            {FieldName: "SubTourForwardTT", Type: "Real"},
            {FieldName: "SubTourReturnTT", Type: "Real"},
            {FieldName: "SubTourEndTime", Type: "Integer"}}
    fldNames = flds.Map(do (f) Return(f.FieldName) end)
    obj.DropFields({FieldNames: fldNames})
    obj.AddFields({Fields: flds})
    obj = null
    Return(true)
endMacro


Macro "SubTour Frequency"(Args)
    objT = CreateObject("Table", Args.MandatoryTours)
    spec = {ToursView: objT.GetView()}
    RunMacro("Eval Sub Tour Choice", Args, spec)
    Return(true)
endMacro


Macro "Eval Sub Tour Choice"(Args, spec)
    // Join Tours to PersonHH
    abm = RunMacro("Get ABM Manager", Args)
    vwT = spec.ToursView
    vwJ = JoinViews("TourData", GetFieldFullSpec(vwT, "PerID"), GetFieldFullSpec(abm.PersonHHView, abm.PersonID), )

    filter = "TourPurpose = 'Work' or TourPurpose = 'Univ'"

    // Run Model for workers and populate results
    obj = CreateObject("PMEChoiceModel", {ModelName: "SubTour Choice"})
    obj.AddTableSource({SourceName: "TourData", View: vwJ, IDField: "TourID"})
    obj.AddMatrixSource({SourceName: "AutoSkim", File: Args.HighwaySkimOP, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
    obj.OutputModelFile = Args.[Output Folder] + "\\Intermediate\\SubTourChoice.mdl"
    obj.AddPrimarySpec({Name: "TourData", OField: "Origin", DField: "Destination", Filter: filter})
    obj.AddUtility({UtilityFunction: Args.SubTourChoiceUtility, AvailabilityExpressions: Args.SubTourChoiceAvail})
    obj.AddOutputSpec({ChoicesField: "SubTour"})
    obj.ReportShares = 1
    obj.RandomSeed = 4199989
    ret = obj.Evaluate()
    if !ret then
        Throw("Running SubTour Choice model failed.")
    Args.[SubTour Choice Spec] = CopyArray(ret)
    obj = null

    if spec.LeaveDataOpen = null then 
        CloseView(vwJ)
endMacro


Macro "SubTour Destination"(Args)
    // Compute size variable field. Add to TAZDemographics output table
    objD = CreateObject("Table", Args.DemographicOutputs)
    obj4D = CreateObject("Table", Args.AccessibilitiesOutputs)
    outFld = "SubTourSizeVar"
    newFlds = {{FieldName: outFld, Type: "real", Width: 12, Decimals: 2}}
    objD.AddFields({Fields: newFlds})
    objJ = objD.Join({Table: obj4D, LeftFields: {"TAZ"}, RightFields: {"TAZID"}})
    
    opt = null
    opt.TableObject = objJ
    opt.Equation = Args.SubTourSizeVar
    opt.FillField = outFld
    opt.ExponentiateCoeffs = 1
    RunMacro("Compute Size Variable", opt)
    objJ = null
    objD = null
    obj4D = null
    
    // Run Destination Choice model
    tag = "SubTour Destination"
    obj = CreateObject("PMEChoiceModel", {ModelName: tag})
    obj.AddTableSource({SourceName: "TourData", File: Args.MandatoryTours, IDField: "TourID"})
    obj.AddMatrixSource({SourceName: "AutoSkim", File: Args.HighwaySkimOP, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
    obj.AddMatrixSource({SourceName: "Intrazonal", File: Args.IZMatrix, RowIndex: "TAZ", ColIndex: "TAZ"})
    obj.AddTableSource({SourceName: "TAZData", File: Args.DemographicOutputs, IDField: "TAZ"})
    obj.AddTableSource({SourceName: "TAZ4Ds", File: Args.AccessibilitiesOutputs, IDField: "TAZID"})
    obj.OutputModelFile = Args.[Output Folder] + "\\Intermediate\\" + tag + ".dcm"
    obj.AddPrimarySpec({Name: "TourData", Filter: "SubTour = 1", OField: "Destination"})
    obj.AddUtility({UtilityFunction: Args.SubTourDestUtility})
    obj.AddDestinations({DestinationsSource: "AutoSkim", DestinationsIndex: "InternalTAZ"})
    obj.AddSizeVariable({Name: "TAZData", Field: "SubTourSizeVar"})
    obj.AddOutputSpec({ChoicesField: "SubTourTAZ"})
    obj.RandomSeed = 4299961
    ret = obj.Evaluate()
    if !ret then
        Throw("Running 'WorkBased Tours Destination' model failed")

    Return(true)
endMacro


Macro "SubTour Duration"(Args)
    objT = CreateObject("Table", Args.MandatoryTours)
    Opts = {ModelName: "SubTourDuration",
            ModelFile: "SubTourDuration.mdl",
            ToursView: objT.GetView(),
            Filter: "SubTour = 1",
            OrigField: "Destination",
            DestField: "SubTourTAZ",
            Availabilities: Args.SubTourDurAvail,
            Utility: Args.SubTourDurUtility,
            ChoiceTable: Args.MandatoryTours,
            ChoiceField: "SubTourActDurInt",
            SimulatedTimeField: "SubTourActDuration",
            AlternativeIntervalInMin: 1,
            RandomSeed: 4399987}
    RunMacro("Subtour Activity Time", Args, Opts)

    Return(true)
endMacro


Macro "SubTour StartTime"(Args)
    // Get Availability based on duration
    altSpec = Args.SubTourStartAlts
    availSpec = RunMacro("Get SubTour St Avails", altSpec)

    objT = CreateObject("Table", Args.MandatoryTours)
    // Run Start Time model
    Opts = {ModelName: "SubTourStartTime",
            ModelFile: "SubTourStartTime.mdl",
            ToursView: objT.GetView(),
            Filter: "SubTour = 1",
            OrigField: "Destination",
            DestField: "SubTourTAZ",
            Availabilities: availSpec,
            Alternatives: altSpec,
            Utility: Args.SubTourStartUtility,
            ChoiceTable: Args.MandatoryTours,
            ChoiceField: "SubTourActStartInt",
            SimulatedTimeField: "SubTourActStartTime",
            RandomSeed: 4499969}
    RunMacro("Subtour Activity Time", Args, Opts)
    Return(true)
endMacro


// Macro that runs the activity time choice model to generate activity duration.
// Called for work tour, univ tour or school tour choices
Macro "Subtour Activity Time"(Args, Opts)
    filter = Opts.Filter

    // Basic Check
    if Opts.Utility = null or Opts.ModelName = null or Opts.ModelFile = null or filter = null
        or Opts.DestField or Opts.ChoiceField = null then
            Throw("Invalid inputs to macro 'Subtour Activity Time'")

    // Join Tours to PersonHH
    abm = RunMacro("Get ABM Manager", Args)
    vwT = Opts.ToursView
    vwJ = JoinViews("TourData", GetFieldFullSpec(vwT, "PerID"), GetFieldFullSpec(abm.PersonHHView, abm.PersonID), )

    // Get Utility Options
    utilOpts = null
    utilOpts.UtilityFunction = Opts.Utility
    if Opts.SubstituteStrings <> null then
        utilOpts.SubstituteStrings = Opts.SubstituteStrings
    if Opts.Availabilities <> null then
        utilOpts.AvailabilityExpressions = Opts.Availabilities

    // Run Model and populate results
    obj = CreateObject("PMEChoiceModel", {ModelName: Opts.ModelName})
    obj.OutputModelFile = Args.[Output Folder] + "\\Intermediate\\" + Opts.ModelFile
    obj.AddTableSource({SourceName: "TourData", View: vwJ, IDField: "TourID"})
    obj.AddMatrixSource({SourceName: "AutoSkim", File: Args.HighwaySkimOP, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
    obj.AddPrimarySpec({Name: "TourData", Filter: filter, OField: Opts.OrigField, DField: Opts.DestField})
    if Opts.Alternatives <> null then
        obj.AddAlternatives({AlternativesList: Opts.Alternatives})
    obj.AddUtility(utilOpts)
    obj.AddOutputSpec({ChoicesField: Opts.ChoiceField})
    obj.ReportShares = 1
    obj.RandomSeed = Opts.RandomSeed
    ret = obj.Evaluate()
    if !ret then
        Throw("Model 'SubTour Activity Time Choice' failed")
    Args.(Opts.ModelName +  " Spec") = CopyArray(ret)

    // Simulate time after choice of interval is made
    simFld = Opts.SimulatedTimeField
    if simFld <> null then do
        objT = CreateObject("Table", Opts.ChoiceTable)
        n = objT.SelectByQuery({Query: filter, SetName: "_Sel"})
        if n > 0 then do
            // Simulate duration in minutes for duration choice predicted above
            opt = null
            opt.ViewSet = vwT + "|_Sel"
            opt.InputField = Opts.ChoiceField
            opt.OutputField = simFld
            opt.HourlyProfile = Opts.HourlyProfile
            opt.AlternativeIntervalInMin = Opts.AlternativeIntervalInMin
            RunMacro("Simulate Time", opt)
        end
    end

    if Opts.LeaveDataOpen = null then
        CloseView(vwJ)
endMacro


Macro "SubTour Mode"(Args)
    objT = CreateObject("Table", Args.MandatoryTours)

    // Filter out rail modes if rail is not present
    ret = RunMacro("Filter Mode Utility Spec", {
        util: Args.SubTourModeUtility,
        Args: Args
    })

    obj = CreateObject("PMEChoiceModel", {SourcesObject: Args.SourcesObject, ModelName: "Work Based Tour Mode Choice"})
    obj.OutputModelFile = Args.[Output Folder] + "\\Intermediate\\SubTourMode.mdl"
    obj.AddTableSource({SourceName: "TourData", View: objT.GetView(), IDField: "TourID"})
    obj.AddMatrixSource({SourceName: "WalkSkim", File: Args.WalkSkim, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
    obj.AddPrimarySpec({Name: "TourData", Filter: "SubTour = 1", OField: "Destination", DField: "SubTourTAZ"})
    obj.AddUtility({UtilityFunction: ret.Utility})
    obj.AddOutputSpec({ChoicesField: "SubTourMode"})
    obj.ReportShares = 1
    obj.RandomSeed = 4599989
    ret = obj.Evaluate()
    if !ret then
        Throw("Running Work Based Tour Choice model failed.")
    Args.[SubTour Mode Spec] = CopyArray(ret)
    obj = null

    // For people who chose 'AutoDriver', determine how many are drive alone vs carpool 
    n = objT.SelectByQuery({SetName: "_SubToursAutoDriver", Query: "SubTour = 1 and Lower(SubTourMode) = 'autodriver'"})
    if n > 0 then do
        SetRandomSeed(4600003)
        params = null
        params.population = {"drivealone", "carpool"}
        params.weight = {Args.SubTourDAPct, 100 - Args.SubTourDAPct}
        vSamples = RandSamples(n, "Discrete", params)
        objT.SubTourMode = vSamples
    end

    // Attach Mode code
    codeMap = {DriveAlone: 1, Carpool: 2, Walk: 3, Bike: 4, Other: 7}
    objT.SelectByQuery({SetName: "_SubTours", Query: "SubTour = 1"})
    v = objT.SubTourMode
    arrModeCode = v2a(v).Map(do (f) Return(codeMap.(f)) end)
    objT.SubTourModeCode = a2v(arrModeCode)

    objT = null
    Return(true)
endMacro


Macro "SubTour PostProcess"(Args)
    // Compute departure time from work/univ and return time for the sub tour
    objT = CreateObject("Table", Args.MandatoryTours)
    vw = objT.GetView()
    filter = "SubTour = 1 and SubTourActStartTime > 0"
    n = objT.SelectByQuery({SetName: "SubTourRecords", Query: filter})
    if n = 0 then
        Throw("No valid SubTours found. Please check the mandatory tours table and check sub-tour models.")

    // Fill destination to subtour stop time. Use Activity Start Time field for the approx departure field
    fillSpec = {View: vw, OField: "Destination", DField: "SubTourTAZ", FillField: "SubTourForwardTT", 
                Filter: filter, ModeField: "SubTourMode", DepTimeField: "SubTourActStartTime"}
    RunMacro("Fill Travel Times", Args, fillSpec)

    // Compute departure time from mandatory destination for the subtour
    // First compute SubTourActStartTime - SubTourForwardTT
    // If this time is greater than the mandatory destination arrival time (i.e. DestArrTime), then this is fine
    // Else set this time to 5 minutes past the DestArrTime
    objT.ChangeSet({SetName: "SubTourRecords"})
    vecs = objT.GetDataVectors({FieldNames: {"DestArrTime", "SubTourActStartTime", "SubTourForwardTT", "SubTourActDuration", "DestDepTime"}})
    vSt = vecs.SubTourActStartTime - vecs.SubTourForwardTT
    vSt = Max(vSt, vecs.DestArrTime + 5)
    objT.SubTourStartTime = vSt

    // Calculate return travel time
    expr = CreateExpression(vw, "SubTourActEnd", "SubTourStartTime + SubTourForwardTT + SubTourActDuration",)
    fillSpec = {View: vw, OField: "SubTourTAZ", DField: "Destination", FillField: "SubTourReturnTT", 
                Filter: filter, ModeField: "SubTourMode", DepTimeField: "SubTourActEnd"}
    RunMacro("Fill Travel Times", Args, fillSpec)
    DestroyExpression(GetFieldFullSpec(vw, expr))
    
    // Calculate sub tour end time
    vDep = vSt + vecs.SubTourForwardTT + vecs.SubTourActDuration
    vEn = vDep + objT.SubTourReturnTT
    objT.SubTourEndTime = vEn

    // Erase data from records where sub tours were not feasible and log values
    filter = "(SubTour = 1) and (SubTourActStartTime = null or SubTourEndTime > DestDepTime)"
    n = objT.SelectByQuery({SetName: "RemoveSubTourRecs", Query: filter})
    if n > 0 then do
        AppendToLogFile(2, String(n) + " records had infeasible SubTour schedules and sub tour information was erased.")
        vecsSet = null
        vNullLong = Vector(n, "Long",)
        vNullString = Vector(n, "String",)
        
        vecsSet = null
        vecsSet.SubTour = Vector(n, "Short", {Constant: 0})
        vecsSet.SubTourTAZ = vNullLong
        vecsSet.SubTourActStartTime = vNullLong
        vecsSet.SubTourActDuration = vNullLong
        vecsSet.SubTourModeCode = vNullLong
        vecsSet.SubTourStartTime = vNullLong
        vecsSet.SubTourEndTime = vNullLong
        vecsSet.SubTourActStartInt = vNullString
        vecsSet.SubTourActDurInt = vNullString
        vecsSet.SubTourMode = vNullString
        vecsSet.SubTourForwardTT = vNullLong
        vecsSet.SubTourReturnTT = vNullLong
        objT.SetDataVectors({FieldData: vecsSet})
    end
    objT = null
    Return(true)
endMacro


// Determines availability expressions for each of the start time alternatives.
// Approximates earliest time a person can start the next activity
// Determined by adding the arrival time after the first tour to twice the free flow travel time from home to new activity
Macro "Get SubTour St Avails"(altSpec)
    availArr = null

    alts = altSpec.Alternative
    availArr.Alternative = CopyArray(alts)

    for alt in alts do
        tmpArr = ParseString(alt, "- ")
        //altStart = s2i(tmpArr[1]) * 60
        //altEnd = s2i(tmpArr[2]) * 60
        //expr = "if (TourData.SubTourActDuration + AutoSkim.Time + " + String(altEnd) + " < TourData.DestDepTime)"
        //expr = expr + " and (TourData.DestArrTime + AutoSkim.Time < " + String(altStart) + ") then 1 else 0"

        altStHr = tmpArr[1]
        altEndHr = tmpArr[2]
        exprSt = printf("Floor((TourData.DestArrTime + AutoSkim.Time)/60.0) <= %s", {altStHr})
        exprEnd = printf("Floor((TourData.DestDepTime - TourData.SubTourActDuration - AutoSkim.Time)/60.0) + 1 >= %s", {altEndHr})
        expr = printf("if (%s) and (%s) then 1 else 0", {exprSt, exprEnd})
        availArr.Expression =  availArr.Expression + {expr}
    end

    Return(availArr)
endMacro
