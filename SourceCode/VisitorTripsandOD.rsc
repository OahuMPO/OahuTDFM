/************************** Trip file creation macros ****************************************
/**********************************************************************************
/*
    Creates visitor trip file from visitor tour file.
    Process 4 sets of tours and two directions (Forward or Return) for each set.
    - Tours without intermediate stops
    - Tours with intermediate stops
*/
Macro "Create Visitor Trip File"(Args)
    // This requires a mandatory tour file
    tourFile = Args.VisitorTours
    if !GetFileInfo(tourFile) then
        Throw("Please run the step to create the visitor tour file before creating a trip file")

    vw = OpenTable("Tours", "FFB", {tourFile})
    vwT = ExportView(vw + "|", "MEM", "ToursMem",,)
    CloseView(vw)

    dirs = {"Forward", "Return"}
    tfArr = {0, 1}
    arrsOut = null
    for dir in dirs do
        spec = {ToursView: vwT, Direction: dir}
        for y in tfArr do
            spec.HasStops = y
            arrs = RunMacro("Extract Visitor Trips", spec)
            RunMacro("Append Arrays", arrs, &arrsOut) 
        end
    end

    vwTemp = RunMacro("Write Vis Trips Table", {Data: arrsOut, StartingTripID: 1, PeriodInfo: Args.TimePeriods})

    // Run mode switch model
    RunMacro("Visitor Mode Switch Model", Args, vwTemp)

    // Export to final table
    exportOpts = {"Row Order": {{"TripID", "Ascending"}} }
    ExportView(vwTemp + "|", "FFB", Args.VisitorTrips,, exportOpts)
    CloseView(vwTemp)
    CloseView(vwT)

    // Convert time fields
    objTrips = CreateObject("Table", Args.VisitorTrips)
    modify = CreateObject("CC.ModifyTableOperation", objTrips.GetView())
    outFlds = {"OrigDepTime", "DestArrTime", "DestDepTime"}
    for fld in outFlds do
        modify.ChangeField(fld, {Type: "Time", Format: "hh:mm tt"})
    end
    //modify.Apply()

    Return(1)
endMacro


/*
    The code that processes the trips from the tours file.
    Records from tour file selected by:
     - Direction
     - Whether or not tours have PUDO stops
     - Whether or not tours have intermediate stops
*/
Macro "Extract Visitor Trips"(spec)
    hasStops = spec.HasStops
    opts = {ToursView: spec.ToursView, Direction: spec.Direction}
    arrs = null
    if !hasStops then
        arrs = RunMacro("Process Direct Vis Tours", opts)
    else
        arrs = RunMacro("Process Vis Tours with Stops", opts)
    
    Return(arrs)
endMacro


/* 
    Extract trips on direct tours
    Since dir is specified, generate only one trip per tour
    # trips added = # selected tours
*/
Macro "Process Direct Vis Tours"(opts)
    vwT = opts.ToursView
    dir = opts.Direction
    stopsFld = "N" + dir + "Stops"
    
    // Select tours and get vectors
    filter = printf("nz(%s) = 0", {stopsFld})
    SetView(vwT)
    n = SelectByQuery("__Selection", "several", "Select * where " + filter,)
    {fields, specs} = GetFields(vwT,)
    vecs = GetDataVectors(vwT + "|__Selection", fields, {OptArray: 1})
    
    // Get output arrays
    ret = RunMacro("Get Direct Vis Leg", vecs, dir)
    Return(CopyArray(ret))
endMacro


Macro "Process Vis Tours with Stops"(opts)
    vwT = opts.ToursView
    dir = opts.Direction
    stopsFld = "N" + dir + "Stops"
    
    filter = printf("%s > 0", {stopsFld})
    SetView(vwT)
    n = SelectByQuery("__Selection", "several", "Select * where " + filter,)
    {fields, specs} = GetFields(vwT,)
    vecs = GetDataVectors(vwT + "|__Selection", fields, {OptArray: 1})
    
    // Trip records for Origin/Destination to first stop
    ret = RunMacro("Get Vis First Stop Leg", vecs, dir)

    // Trip records for last stop leg to origin/destination
    ret1 = RunMacro("Get Vis Last Stop Leg", vecs, dir)
    RunMacro("Append Arrays", ret1, &ret)

    // Trips for intermediate stops
    filter = printf("%s >= 2", {stopsFld})
    n = SelectByQuery("__Selection", "several", "Select * where " + filter,)
    if n > 0 then do
        vecs = GetDataVectors(vwT + "|__Selection", fields, {OptArray: 1})
        ret2 = RunMacro("Get Vis Intermediate Stop Leg", vecs, dir)
        RunMacro("Append Arrays", ret2, &ret)
    end
    
    Return(CopyArray(ret))
endMacro


// Trip info for direct tour leg eithe between hotel to main tour destination or back from main tour destination to hotel
Macro "Get Direct Vis Leg"(vecs, dir)
    n = vecs.TourID.length
    vHome = Vector(n, "String", {Constant: 'Hotel'})
    vLegNo = Vector(n, "Short", {Constant: 1})
    if dir = 'Forward' then do
        vOPurp = vHome
        vDPurp = vecs.TourType
        vOrig = vecs.Origin
        vDest = vecs.Destination
        vOrigDep = vecs.TourStartTime
        vDestArr = vecs.TourStartTime + vecs.ForwardTT
        vDestDep = vecs.ActivityEndTime
    end
    else do
        vOPurp = vecs.TourType
        vDPurp = vHome
        vOrig = vecs.Destination
        vDest = vecs.Origin
        vOrigDep = vecs.ActivityEndTime
        vDestArr = vecs.TourEndTime
        vDestDep = Vector(n, "Long",)
    end
    vMode = vecs.Mode

    // Get output arrays
    retArr = RunMacro("Get Basic Vis Trip Info", vecs, dir)
    retArr.Origin = v2a(vOrig)
    retArr.Destination = v2a(vDest)
    retArr.Mode = v2a(vMode)
    retArr.OrigPurpose = v2a(vOPurp)
    retArr.DestPurpose = v2a(vDPurp)
    retArr.LegNo = v2a(vLegNo)
    retArr.OrigDep = v2a(vOrigDep)
    retArr.DestArr = v2a(vDestArr)
    retArr.DestDep = v2a(vDestDep)
    Return(CopyArray(retArr))
endMacro


// Trip info for tour leg eithe between hotel to first stop on the forward tour leg  or from main tour destination to first stop on return tour leg
Macro "Get Vis First Stop Leg"(vecs, dir)
    n = vecs.TourID.length
    stopTAZFld = "Stop" + dir + "TAZ1"
    vLegNo = Vector(n, "Short", {Constant: 1})
    if dir = "Forward" then do
        vOrig = vecs.Origin
        vOPurp = Vector(n, "String", {Constant: "Hotel"})
        vDPurp = vecs.PurposeForwardStop1
        vOrigDep = vecs.TourStartTime
        vDestArr = vOrigDep + vecs.TimeToStopForward1
        vDestDep = vDestArr + vecs.ForwardStopDuration1
    end
    else do
        vOrig = vecs.Destination
        vOPurp = vecs.TourType
        vDPurp = vecs.PurposeReturnStop1
        vOrigDep = vecs.ActivityEndTime
        vDestArr = vOrigDep + vecs.TimeToStopReturn1
        vDestDep = vDestArr + vecs.ReturnStopDuration1
    end
    vMode = vecs.Mode

    retArr = RunMacro("Get Basic Vis Trip Info", vecs, dir)
    retArr.Origin = v2a(vOrig)
    retArr.Destination = v2a(vecs.(stopTAZFld))
    retArr.Mode = v2a(vMode)
    retArr.OrigPurpose = v2a(vOPurp)
    retArr.DestPurpose = v2a(vDPurp)
    retArr.LegNo = v2a(vLegNo)
    retArr.OrigDep = v2a(vOrigDep)
    retArr.DestArr = v2a(vDestArr)
    retArr.DestDep = v2a(vDestDep)
    Return(CopyArray(retArr))
endMacro


Macro "Get Vis Intermediate Stop Leg"(vecs, dir)
    n = vecs.TourID.length
    modeFld = dir + "Mode"
    vLegNo = Vector(n, "Short", {Constant: 2})
    if dir = "Forward" then do
        vOrigDep = vecs.TourStartTime + vecs.TimeToStopForward1 + vecs.ForwardStopDuration1
        vDestArr = vOrigDep + nz(vecs.ForwardStop1to2Time)
        vDestDep = vDestArr + vecs.ForwardStopDuration2
        vOrig = vecs.StopForwardTAZ1
        vDest = vecs.StopForwardTAZ2
        vOPurp = vecs.PurposeForwardStop1
        vDPurp = vecs.PurposeForwardStop2
    end
    else do
        vOrigDep = vecs.ActivityEndTime + vecs.TimeToStopReturn1 + vecs.ReturnStopDuration1
        vDestArr = vOrigDep + nz(vecs.ReturnStop1to2Time)
        vDestDep = vDestArr + vecs.ReturnStopDuration2
        vOrig = vecs.StopReturnTAZ1
        vDest = vecs.StopReturnTAZ2
        vOPurp = vecs.PurposeReturnStop1
        vDPurp = vecs.PurposeReturnStop2
    end
    vMode = vecs.Mode
    retArr = RunMacro("Get Basic Vis Trip Info", vecs, dir)
    retArr.Origin = v2a(vOrig)
    retArr.Destination = v2a(vDest)
    retArr.Mode = v2a(vMode)
    retArr.OrigPurpose = v2a(vOPurp)
    retArr.DestPurpose = v2a(vDPurp)
    retArr.LegNo = v2a(vLegNo)
    retArr.OrigDep = v2a(vOrigDep)
    retArr.DestArr = v2a(vDestArr)
    retArr.DestDep = v2a(vDestDep)
    Return(CopyArray(retArr))
endMacro



Macro "Get Vis Last Stop Leg"(vecs, dir)
    n = vecs.TourID.length
    if dir = "Forward" then do
        vStops = nz(vecs.NForwardStops)
        vOrig = if vStops = 2 then vecs.StopForwardTAZ2 else vecs.StopForwardTAZ1
        vDest = vecs.Destination
        vOPurp = if vStops = 2 then vecs.PurposeForwardStop2 else vecs.PurposeForwardStop1
        vDPurp = vecs.TourType
        vOrigDep = vecs.TourStartTime + vecs.TimeToStopForward1 + vecs.ForwardStopDuration1
                     + nz(vecs.ForwardStop1to2Time) + nz(vecs.ForwardStopDuration2)
        vDestArr = vecs.ActivityStartTime
        vDestDep = vecs.ActivityEndTime
    end
    else do
        vStops = nz(vecs.NReturnStops)
        vOrig = if vStops = 2 then vecs.StopReturnTAZ2 else vecs.StopReturnTAZ1
        vDest = vecs.Origin
        vOPurp = if vStops = 2 then vecs.PurposeReturnStop2 else vecs.PurposeReturnStop1
        vDPurp = Vector(n, "String", {Constant: "Hotel"})
        vOrigDep = vecs.ActivityEndTime + vecs.TimeToStopReturn1 + vecs.ReturnStopDuration1
                     + nz(vecs.ReturnStop1to2Time) + nz(vecs.ReturnStopDuration2)
        vDestArr = vecs.TourEndTime
        vDestDep = Vector(n, "Long",)
    end

    vMode = vecs.Mode
    retArr = RunMacro("Get Basic Vis Trip Info", vecs, dir)
    retArr.Origin = v2a(vOrig)
    retArr.Destination = v2a(vDest)
    retArr.Mode = v2a(vMode)
    retArr.OrigPurpose = v2a(vOPurp)
    retArr.DestPurpose = v2a(vDPurp)
    retArr.LegNo = v2a(vStops + 1)
    retArr.OrigDep = v2a(vOrigDep)
    retArr.DestArr = v2a(vDestArr)
    retArr.DestDep = v2a(vDestDep)
    Return(CopyArray(retArr))
endMacro


Macro "Get Basic Vis Trip Info"(vecs, dir)
    n = vecs.TourID.length
    ret = null
    ret.TourID = v2a(vecs.TourID)
    ret.HHID = v2a(vecs.HHID)
    ret.TourType = v2a(vecs.TourType)
    ret.Direction = v2a(Vector(n, "String", {Constant: left(dir, 1)}))
    Return(ret)
endMacro


Macro "Write Vis Trips Table"(spec)
    data = spec.Data
    nRecs = data[1][2].length
    periodInfo = spec.PeriodInfo
    vwOut = RunMacro("Create Empty Vis Trip File", {ViewName: "TempTrips", NRecords: nRecs})
    
    vecsOut = null
    for item in data do
        fld = item[1]
        vecsOut.(fld) = a2v(data.(fld))
    end
    vecsOut.One = Vector(nRecs, "Short", {Constant: 1})
    vecsOut.Period = RunMacro("Get TOD Vector", (vecsOut.OrigDep + vecsOut.DestArr)/2, periodInfo)
    vecsOut.Mode = Lower(vecsOut.Mode)
    SetDataVectors(vwOut + "|", vecsOut,)

    // Write Trip ID field
    startingTripID = spec.StartingTripID
    vTripID = Vector(nRecs, "Long", {{"Sequence", startingTripID, 1}})
    order = {TourID: "Ascending", Direction: "Ascending", LegNo: "Ascending"}
    SetDataVector(vwOut + "|", "TripID", vTripID, {SortOrder: order})

    // Convert time fields
    inFlds = {"OrigDep", "DestArr", "DestDep"}
    outFlds = inFlds.Map(do (f) Return(f + "Time") end)
    objT = CreateObject("Table", vwOut)
    RunMacro("Create Time Fields", {TripsObj: objT, InputFields: inFlds, OutputFields: outFlds})
    Return(vwOut)
endMacro


Macro "Create Empty Vis Trip File"(spec)
    flds = {{"TripID", "Integer", 12, null, "Yes"},
            {"TourID", "Integer", 12, null, "Yes"},
            {"HHID", "Integer", 12, null, "Yes"},
            {"TourType", "String", 12, null, "No"},
            {"Direction", "String", 1, null, "No"},
            {"LegNo", "Integer", 8, null, "No"},
            {"OrigPurpose", "String", 18, null, "No"},
            {"DestPurpose", "String", 18, null, "No"},
            {"Origin", "Integer", 12, null, "Yes"},
            {"Destination", "Integer", 12, null, "Yes"},
            {"OrigDep", "Integer", 12, null, "No"},
            {"DestArr", "Integer", 12, null, "No"},
            {"DestDep", "Integer", 12, null, "No"},
            {"Period", "String", 2, null, "No"},
            {"Mode", "String", 15, null, "No"},
            {"One", "Tiny", 1, null, "No"}}
    vwOut = CreateTable(spec.ViewName,, "MEM", flds)
    if spec.NRecords > 0 then
        AddRecords(vwOut,,, {{"Empty Records", spec.NRecords}})    
    Return(vwOut)
endMacro


// Create Visitor AM, PM and OP matrices with modes as the cores
Macro "Write Visitor OD"(Args)
    on error do
        ShowMessage(GetLastError())
        return(0)
    end
    Args.ABMFlag = 0

    mSkimObj = CreateObject("Matrix", Args.HighwaySkimAM)
    mSkimObj.SetIndex("TAZ")
    mcSkim = mSkimObj.Time

    objT = CreateObject("Table", Args.VisitorTrips)
    vM = objT.Mode
    modes = SortArray(v2a(vM), {Unique: "True"})
    
    periods = {'AM', 'PM', 'OP'}
    for p in periods do
        outFile = Args.(p + "_Visitor_OD")
        label = printf("%s_VisitorOD", {p})

        mOpts = {FileName: outFile, Tables: modes, Label: label}
        mat = CopyMatrixStructure({mcSkim}, mOpts)
        mObj = CreateObject("Matrix", mat)
        
        filter = printf("Period = '%s'", {p})
        n = objT.SelectByQuery({SetName: "__Selection", Query: filter})
        UpdateMatrixFromView(mat, objT.GetView() + "|__Selection", "Origin", "Destination", 
                                GetFieldFullSpec(vwTrips, "Mode"),          
                                {GetFieldFullSpec(vwTrips, "One")}, 
                                "Add", {"Missing is zero": "Yes"})

        mObj = null
    end

    mSkimObj = null
    objT = null
    return(1)
endMacro


Macro "Visitor Mode Switch Model"(Args, vwT)
    // Add temporary walk time and walk distance fields
    objT = CreateObject("Table", vwT)
    flds = {{FieldName: "SkimTime", Type: "Real"},
            {FieldName: "SkimDist", Type: "Real"}}
    objT.AddFields({Fields: flds})

    // Determine mode switch between 4 combinations
    // 1. Auto to Walk (IZ trips only)
    opts = {VisitorTripsObj: objT, 
            Skim: Args.WalkSkim, ModeTo: "walk",
            Filter: "(Mode = 'sov' or Mode = 'hov2' or Mode = 'hov3') and (Origin = Destination) and !(Direction = 'F' and LegNo = 1)",
            SwitchPct: Args.AutoToWalkShiftPct
            }
    RunMacro("Run Visitor Mode Switch", opts)

    // 2. TNC to Walk
    opts = {VisitorTripsObj: objT, 
            Skim: Args.WalkSkim, ModeTo: "walk",
            Filter: "(Mode = 'tnc') and !(Direction = 'F' and LegNo = 1)",
            SwitchPct: Args.TNCToWalkShiftPct
            }
    RunMacro("Run Visitor Mode Switch", opts)

    // 3. Bus to Walk
    opts = {VisitorTripsObj: objT, 
            Skim: Args.WalkSkim, ModeTo: "walk",
            Filter: "(Mode = 'w_bus') and !(Direction = 'F' and LegNo = 1)",
            SwitchPct: Args.BusToWalkShiftPct
            }
    RunMacro("Run Visitor Mode Switch", opts)

    // 4. Bus to TNC
    opts = {VisitorTripsObj: objT, 
            Skim: Args.HighwaySkimOP, ModeTo: "tnc",
            Filter: "(Mode = 'w_bus') and !(Direction = 'F' and LegNo = 1)",
            SwitchPct: Args.BusToTNCShiftPct
            }
    RunMacro("Run Visitor Mode Switch", opts)

    // Remove temporary walk time and walk distance fields
    objT.DropFields({FieldNames: {"SkimTime", "SkimDist"}})
    objT = null
endMacro


Macro "Run Visitor Mode Switch"(opts)
    objT = opts.VisitorTripsObj
    modeT = opts.ModeTo
    filter = opts.Filter
    target = opts.SwitchPct/100.0

    n = objT.SelectByQuery({Query: filter, SetName: "_MasterSet"})
    if n = 0 then 
        Return(1)

    // Fill WalkTime and WalkDistance for the master set records
    mObj = CreateObject('Matrix', opts.Skim)
    mObj.SetIndex({RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
    mcT = mObj.Time
    mcD = mObj.Distance
    
    vw = objT.GetView()
    oSpec = GetFieldFullSpec(vw, "Origin")
    dSpec = GetFieldFullSpec(vw, "Destination")
    
    fSpec = GetFieldFullSpec(vw, "SkimTime")
    FillViewFromMatrix(vw + "|_MasterSet", oSpec, dSpec, {{fSpec, mcT}})

    fSpec = GetFieldFullSpec(vw, "SkimDist")
    FillViewFromMatrix(vw + "|_MasterSet", oSpec, dSpec, {{fSpec, mcD}})

    mcT = null
    mcD = null
    mObj = null

    // Identify feasible set
    minDurQry = "(DestDep = null) or (DestDep - OrigDep - SkimTime >= 10)" // 10 is the min activity/stop duration
    if modeT = "walk" then
        qry = printf("(%s) and (SkimDist <= 1.5) and (%s)", {filter, minDurQry})
    else
        qry = printf("(%s) and (%s)", {filter, minDurQry})
    n1 = objT.SelectByQuery({Query: qry, SetName: "_FeasibleSet"})
    if n1 = 0 then 
        Return(1)

    // Update target so that it can be achieved using the feasible set
    newTarget = Min(target*n/n1, 1.0)

    // Get a uniform dist random vector
    SetRandomSeed(r2i(target*10000))
    v = RandSamples(n1, "Uniform",)
    vCurrMode = objT.Mode
    vNewDestArr = objT.OrigDep + objT.SkimTime
    vNewMode = if v <= newTarget then modeT else vCurrMode
    vDestArr = if v <= newTarget then vNewDestArr else objT.DestArr
    objT.Mode = vNewMode
    objT.DestArr = vDestArr
endMacro
