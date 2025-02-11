/************************** Tour file creation macros ****************************************
/* 
    Create master tour file
    1. Create tours from person decisions
    2. Modify tours based on PUDO stops
    3. Create separate Dropoff/Pickup tours table as necessary
    4. Concatenate adjusted tours with dropoff and pickup tours tables
*/
Macro "Create Mandatory Tour File"(Args)
    abm = RunMacro("Get ABM Manager", Args)

    // Process tours
    arrsOut = null
    purposes = {'Work1', 'Work2', 'Univ1', 'Univ2', 'School'}
    for p in purposes do
        // Process direct tours
        spec = {Label: p, abmManager: abm, Periods: Args.TimePeriods}
        arrs = RunMacro("Synthesize Tour Data", spec)
        RunMacro("Append Arrays", arrs, &arrsOut)
    end

    // Write Tours Table (from the direct tours above)
    vwTours = RunMacro("Write Tours Table", {Data: arrsOut})

    // Modify tours that also had PUDO activities. This will update the appropriate records of the tour file.
    periods = {"AM", "PM", "OP"}
    skimArgs = {TimePeriods: Args.TimePeriods}
    for p in periods do
        mObj.(p) = CreateObject("Matrix", Args.("HighwaySkim" + p))
        mObj.(p).SetIndex({RowIndex: "TAZ", ColIndex: "TAZ"})
        skimArgs.(p) = mObj.(p).Time
    end
    RunMacro("Modify Tours for PUDO", {ToursView: vwTours, abmManager: abm, SkimArgs: skimArgs, Periods: Args.TimePeriods})
    
    // Add Separate PUDO Tours. This will create two new tables that have to be concatenated with the tours file.
    puFile = GetTempPath() + "PUTours.bin"
    doFile = GetTempPath() + "DOTours.bin"
    opts = {abmManager: abm, SkimArgs: skimArgs, Periods: Args.TimePeriods, DropoffTourFile: doFile, PickupTourFile: puFile}
    RunMacro("Add Separate PUDO tours", opts)
    
    // Export tours table to temporary file
    tempToursFile = GetTempPath() + "TempTours.bin"
    ExportView(vwTours + "|", "FFB", tempToursFile,,)
    CloseView(vwTours)
    
    // Concatenate PU, DO and ToursFile
    ConcatenateFiles({tempToursFile, doFile, puFile}, Args.MandatoryTours)
    dictFile = Substitute(Args.MandatoryTours, ".bin", ".dcb",)
    CopyFile(GetTempPath() + "TempTours.dcb", dictFile)

    // Postprocess: Fill Tour ID field
    vwT = OpenTable("Tours", "FFB", {Args.MandatoryTours})
    v = GetDataVector(vwT + "|", "TourID",)
    v1 = Vector(v.length, "Long", {{"Sequence", 1, 1}})
    sortOrder = {{'HID', 'Ascending'}, {'PerID', 'Ascending'}, {'ActivityStartTime', 'Ascending'}}
    SetDataVector(vwT + "|", "TourID", v1, {SortOrder: sortOrder})
    RunMacro("Handle MT Access Mode Issues", vwT)
        
    CloseView(vwT)
    Return(true)
endMacro


/*
    Generate tours from person decision for a given purpose: 'Work1', 'Work2', 'HigherEd1', 'HigherEd2', 'School'
    Returns an option array. The option names are the field names and the values are vectors containing one record for each tour.
*/
Macro "Synthesize Tour Data"(spec)
    label = spec.Label
    pInfo = RunMacro("Get Purpose Info", label)
    purpose = pInfo.Purpose
    tourNumber = pInfo.TourNo
    abm = spec.abmManager
    PeriodInfo = spec.Periods

    destFld = purpose + "TAZ" // e.g. 'WorkTAZ'
    actStartFld = label + "_StartTime"
    actDurFld = label + "_Duration"
    ttForwardFld = "HomeTo" + purpose + "Time"
    ttReturnFld = purpose + "ToHomeTime"
    if purpose = "School" then do
        modeFldF = "SchoolForwardMode"
        modeFldR = "SchoolReturnMode"
        purpFilter = "(AttendSchool = 1 or AttendDaycare = 1)"
    end
    else do
        modeFldF = purpose + "Mode"
        modeFldR = purpose + "Mode"
        purpFilter = printf("(Number%sTours >= %s)", {purpose, tourNumber})
    end
    modeCodeFldF = modeFldF + "Code"
    modeCodeFldR = modeFldR + "Code"
    
    // Get Filter: Example 'NumberWorkTours = 1 and WorkTAZ <> null'
    filter = printf("%s and %s <> null and %s_StartTime <> null", {purpFilter, destFld, label})
    set = abm.CreatePersonSet({Filter: filter, Activate: 1})
    nRecs = set.Size
    
    hhIDSpec = GetFieldFullSpec(abm.HHView, "HouseholdID")
    flds = {"PersonID", hhIDSpec, "TAZID", destFld, actStartFld, actDurFld, ttForwardFld,
            modeFldF, modeFldR, modeCodeFldF, modeCodeFldR, 
            "DropoffTourFlag", "PickupTourFlag", "NDropOffsEnRoute", "NPickupsEnRoute"}
    vecs = abm.GetPersonHHVectors(flds)

    // Tours
    arrs = null
    arrs.PerID = v2a(vecs.PersonID)
    arrs.HID = v2a(vecs.(hhIDSpec))
    arrs.TourType = v2a(Vector(nRecs, "String", {Constant: "Mandatory"}))
    arrs.TourPurpose = v2a(Vector(nRecs, "String", {Constant: purpose}))
    arrs.MandTourNo = arrTourNo
    arrs.HTAZ = v2a(vecs.TAZID)
    arrs.Origin = v2a(vecs.TAZID)
    arrs.Destination = v2a(vecs.(destFld))
    arrs.ForwardMode = v2a(vecs.(modeFldF))
    arrs.ReturnMode = v2a(vecs.(modeFldR))
    arrs.ForwardModeCode = v2a(vecs.(modeCodeFldF))
    arrs.ReturnModeCode = v2a(vecs.(modeCodeFldR))
    
    vTourStart = vecs.(actStartFld) - vecs.(ttForwardFld)
    arrs.TourStartTime = v2a(vTourStart)
    arrs.DestArrTime = v2a(vecs.(actStartFld))
    arrs.ActivityStartTime = v2a(vecs.(actStartFld))
    arrs.ActivityDuration = v2a(vecs.(actDurFld))
    
    vActEnd = vecs.(actStartFld) + vecs.(actDurFld)
    arrs.ActivityEndTime = v2a(vActEnd)
    arrs.DestDepTime = v2a(vActEnd)
    
    //vTourEnd = vActEnd + vecs.(ttReturnFld)
    vTourEnd = vActEnd + vecs.(ttForwardFld)
    arrs.TourEndTime = v2a(vTourEnd)
    
    vTODF = RunMacro("Get TOD Vector", vTourStart, PeriodInfo)
    vTODR = RunMacro("Get TOD Vector", vActEnd, PeriodInfo)
    arrs.TODForward = v2a(vTODF)
    arrs.TODReturn = v2a(vTODR)

    // Identify tour records that need to be modified for PUDO
    vModifyDO = if (vecs.NDropOffsEnRoute > 0) or (vecs.DropoffTourFlag = 'W1' and vecs.(modeCodeFldF) = 2) then 1
    arrs.ModifyRecforDropoff = v2a(vModifyDO)

    vModifyPU = if (vecs.NPickupsEnRoute > 0) or (vecs.PickupTourFlag = 'W1' and vecs.(modeCodeFldR) = 2) then 1
    arrs.ModifyRecforPickup = v2a(vModifyPU)

    Return(arrs)
endMacro


Macro "Get Purpose Info"(str)
    validStrs = {'Work1', 'Work2', 'Univ1', 'Univ2', 'School'}
    if validStrs.position(str) = 0 then
        Throw("Invalid input to macro 'Get Purpose Info'")

    if lower(str) = "school" or lower(str) = "daycare" then do
        purpose = "School"
        tourNo = "1"
    end
    else do
        purpose = Left(str, StringLength(str) - 1)
        tourNo = Right(str, 1)
    end
    Return({Purpose: purpose, TourNo: tourNo})
endMacro


/*
    Write data from option array into the final tours table
*/
Macro "Write Tours Table"(spec)
    data = spec.Data
    nRecs = data.Origin.length
    vwOut = RunMacro("Create Empty Tour File", {ViewName: "Tours", NRecords: nRecs})
    
    vecsOut = null
    for item in data do
        fld = item[1]
        vecsOut.(fld) = a2v(data.(fld))
    end
    vecsOut.AssignForwardHalf = if vecsOut.ForwardMode = 'Carpool' and Lower(vecsOut.TourPurpose) = 'school' then 0 else 1
    vecsOut.AssignReturnHalf = if vecsOut.ReturnMode = 'Carpool' and Lower(vecsOut.TourPurpose) = 'school' then 0 else 1
    SetDataVectors(vwOut + "|", vecsOut,)

    // Fill Mandatory Tour No field
    obj = CreateObject("Table", vwOut)
    order = {{"PerID", "Ascending"}, {"ActivityStartTime", "Ascending"}}
    obj.Sort({FieldArray: order})
    
    vP = obj.PerID
    vPrevP = RunMacro("Shift Vector", {Vector: vP, Method: 'Prev'})
    v = if (vP = vPrevP) then 2 else 1
    obj.MandTourNo = v

    Return(vwOut)
endMacro


// Create a temporary in-memory tour table and return the view
Macro "Create Empty Tour File"(spec)
    flds = {{"TourID", "Integer", 12, null, "Yes"},
            {"HID", "Integer", 12, null, "Yes"},
            {"PerID", "Integer", 12, null, "Yes"},
            {"TourType", "String", 12, null, "No"},
            {"TourPurpose", "String", 8, null, "Yes"},
            {"MandTourNo", "Short", 2, null, "No"},
            {"HTAZ", "Integer", 12, null, "Yes"},
            {"Origin", "Integer", 12, null, "Yes"},
            {"Destination", "Integer", 12, null, "Yes"},
            {"ForwardModeCode", "Integer", 2, null, "No"},
            {"ReturnModeCode", "Integer", 2, null, "No"},
            {"ForwardMode", "String", 15, null, "No"},
            {"ReturnMode", "String", 15, null, "No"},
            {"TourStartTime", "Integer", 12, null, "No"},
            {"DestArrTime", "Integer", 12, null, "No"},
            {"ActivityStartTime", "Integer", 12, null, "No"},
            {"ActivityDuration", "Integer", 12, null, "No"},
            {"ActivityEndTime", "Integer", 12, null, "No"},
            {"DestDepTime", "Integer", 12, null, "No"},
            {"TourEndTime", "Integer", 12, null, "No"},
            {"TODForward", "String", 2, null, "No"},
            {"TODReturn", "String", 2, null, "No"},
            {"NumDropoffs", "Short", 2, null, "No"},
            {"DropoffKidID1", "Integer", 12, null, "No"},
            {"DropoffTAZ1", "Integer", 12, null, "No"},
            {"ArrDropoff1", "Integer", 12, null, "No"},
            {"DepDropoff1", "Integer", 12, null, "No"},
            {"DropoffKidID2", "Integer", 12, null, "No"},
            {"DropoffTAZ2", "Integer", 12, null, "No"},
            {"ArrDropoff2", "Integer", 12, null, "No"},
            {"DepDropoff2", "Integer", 12, null, "No"},
            {"NumPickups", "Short", 2, null, "No"},
            {"PickupKidID1", "Integer", 12, null, "No"},
            {"PickupTAZ1", "Integer", 12, null, "No"},
            {"ArrPickup1", "Integer", 12, null, "No"},
            {"DepPickup1", "Integer", 12, null, "No"},
            {"PickupKidID2", "Integer", 12, null, "No"},
            {"PickupTAZ2", "Integer", 12, null, "No"},
            {"ArrPickup2", "Integer", 12, null, "No"},
            {"DepPickup2", "Integer", 12, null, "No"},
            {"StopsChoice", "String", 3, null, "No"},
            {"NForwardStops", "Short", 2, null, "No"},
            {"StopForwardTAZ", "Integer", 12, null, "No"},
            {"IsStopBeforeDropoffs", "Short", 2, null, "No"},
            {"ForwardStopDeltaTT", "Real", 12, 2, "No"},
            {"ForwardStopDurChoice", "String", 12,, "No"},
            {"ForwardStopDuration", "Real", 12, 2, "No"},
            {"TimeToStopF", "Real", 12, 2, "No"},
            {"TimeFromStopF", "Real", 12, 2, "No"},
            {"NReturnStops", "Short", 2, null, "No"},
            {"StopReturnTAZ", "Integer", 12, null, "No"},
            {"IsStopBeforePickups", "Short", 2, null, "No"},
            {"ReturnStopDeltaTT", "Real", 12, 2, "No"},
            {"ReturnStopDurChoice", "String", 12,, "No"},
            {"ReturnStopDuration", "Real", 12, 2, "No"},
            {"TimeToStopR", "Real", 12, 2, "No"},
            {"TimeFromStopR", "Real", 12, 2, "No"},
            {"NumSubTours", "Short", 2, null, "No"},
            {"ModifyRecforDropoff", "Short", 2, null, "No"},
            {"ModifyRecforPickup", "Short", 2, null, "No"},
            {"AssignForwardHalf", "Short", 2, null, "No"},
            {"AssignReturnHalf", "Short", 2, null, "No"}}
    vwOut = CreateTable(spec.ViewName,, "MEM", flds)
    if spec.NRecords > 0 then
        AddRecords(vwOut,,, {{"Empty Records", spec.NRecords}})    
    Return(vwOut)
endMacro


Macro "Append Arrays"(arrs, arrsOut)
    if arrsOut = null then
        arrsOut = CopyArray(arrs)
    else do
        inputFlds = arrs.Map(do (f) Return(f[1]) end) // Get input option names
        inputLength = arrs[1][2].length
        
        outputFlds = arrsOut.Map(do (f) Return(f[1]) end) // Get output option names
        outputLength = arrsOut[1][2].length
        
        // Loop over each output field and append appropriate vector
        for fld in outputFlds do
            if inputFlds.position(fld) > 0 then // Append given input array
                arrsOut.(fld) = arrsOut.(fld) + CopyArray(arrs.(fld))
            else do                             // Not present in input arrays. Append null records.
                dim nullArray[inputLength]
                arrsOut.(fld) = CopyArray(arrsOut.(fld)) + nullArray
            end
        end

        // Check for any new input fields not already present in arrsOut
        for fld in inputFlds do
            if outputFlds.position(fld) = 0 then do // New field
                dim nullArray[outputLength]
                arrsOut.(fld) = nullArray + CopyArray(arrs.(fld))
            end
        end        
    end
endMacro


/************************** Trip file creation macros ****************************************
/**********************************************************************************
/*
    Creates mandatory trip file from mandatory tour file.
    Mandatory trips are created from mandatory tours.
    Process 4 sets of tours and two directions (Forward or Return) for each set.
    - Tours without PUDO and without intermediate stops
    - Tours with PUDO but no intermediate stops
    - Tours without PUDO but with intermediate stops
    - Tours with PUDO and intermediate stops with two sub-cases
        - Stop before PUDO
        - Stop after PUDO
*/
Macro "Create Mandatory Trip File"(Args)
    // This requires a mandatory tour file
    tourFile = Args.MandatoryTours
    if !GetFileInfo(tourFile) then
        Throw("Please run the step to create the mandatory tour file before creating a trip file")

    vw = OpenTable("Tours", "FFB", {tourFile})
    vwT = ExportView(vw + "|", "MEM", "ToursMem",,)
    CloseView(vw)

    dirs = {"Forward", "Return"}
    tfArr = {0, 1}
    arrsOut = null
    for dir in dirs do
        spec = {ToursView: vwT, Direction: dir}
        for x in tfArr do
            spec.HasPUDO = x
            for y in tfArr do
                spec.HasStops = y
                arrs = RunMacro("Extract Trips", spec)
                RunMacro("Append Arrays", arrs, &arrsOut) 
            end
        end
        arrs = RunMacro("Extract Subtour Trip", Args, spec)
        RunMacro("Append Arrays", arrs, &arrsOut)
    end

    opts = {Data: arrsOut, CarpoolOccupancy: Args.WorkCarpoolOccupancy, NonHHAutoOccupancy: Args.NonHHAutoOccupancy, StartingTripID: 1}
    vwTemp = RunMacro("Write Trips Table", opts)

    // Export to final table
    exportOpts = {"Row Order": {{"TripID", "Ascending"}} }
    ExportView(vwTemp + "|", "FFB", Args.MandatoryTrips,, exportOpts)
    CloseView(vwTemp)
    CloseView(vwT)
    Return(1)
endMacro


/*
    The code that processes the trips from the tours file.
    Records from tour file selected by:
     - Direction
     - Whether or not tours have PUDO stops
     - Whether or not tours have intermediate stops
*/
Macro "Extract Trips"(spec)
    hasPudo = spec.HasPUDO
    hasStops = spec.HasStops

    opts = {ToursView: spec.ToursView, Direction: spec.Direction}
    arrs = null
    // No PUDO and no stops
    if !hasPudo and !hasStops then                      // No PUDO and no stops
        arrs = RunMacro("Process Direct Tours", opts)
    else if hasPudo and !hasStops then                  // PUDO but no stops.
        arrs = RunMacro("Process Tours with Pudo", opts)
    else if !hasPudo and hasStops then                  // Stops but no PUDO.
        arrs = RunMacro("Process Tours with Stops", opts)
    else                                                // Both PUDO and Stops. Will have only one stop and upto two PUDO stops.
        arrs = RunMacro("Process Tours with Pudo and Stops", opts)
    
    Return(arrs)
endMacro


/* 
    Extract trips on direct tours
    Since dir is specified, generate only one trip per tour
    # trips added = # selected tours
*/
Macro "Process Direct Tours"(opts)
    vwT = opts.ToursView
    dir = opts.Direction
    if dir = 'Forward' then
        pudoFld = "NumDropoffs"
    else
        pudoFld = "NumPickups"
    stopsFld = "N" + dir + "Stops"
    
    // Select tours and get vectors
    filter = printf("nz(%s) = 0 and nz(%s) = 0", {pudoFld, stopsFld})
    SetView(vwT)
    n = SelectByQuery("__Selection", "several", "Select * where " + filter,)
    {fields, specs} = GetFields(vwT,)
    vecs = GetDataVectors(vwT + "|__Selection", fields, {OptArray: 1})
    
    // Get output arrays
    ret = RunMacro("Get Direct Leg", vecs, dir)
    Return(CopyArray(ret))
endMacro


/* 
    Extract trips from tours with PUDO stops but no intermediate stops
    - Maximum of three PUDO stops on forward or return leg
    - Maximum of 4 trips per forward/return leg
      # Trips = NToursWithPUDO * (Number PUDO stops + 1)
*/
Macro "Process Tours with Pudo"(opts)
    vwT = opts.ToursView
    dir = opts.Direction
    if dir = 'Forward' then
        pudoFld = "NumDropoffs"
    else
        pudoFld = "NumPickups"
    stopsFld = "N" + dir + "Stops"
    
    // Trip records for Home/Dest to first PUDO stop
    filter = printf("%s > 0 and nz(%s) = 0 and Lower(TourPurpose) <> 'pickup'", {pudoFld, stopsFld})
    SetView(vwT)
    n = SelectByQuery("__Selection", "several", "Select * where " + filter,)
    {fields, specs} = GetFields(vwT,)
    vecs = GetDataVectors(vwT + "|__Selection", fields, {OptArray: 1})
    
    ret = RunMacro("Get First PUDO Leg", vecs, dir)

    // Trip records for last PUDO stop to Dest/Home
    filter = printf("%s > 0 and nz(%s) = 0 and Lower(TourPurpose) <> 'dropoff'", {pudoFld, stopsFld})
    SetView(vwT)
    n = SelectByQuery("__Selection", "several", "Select * where " + filter,)
    {fields, specs} = GetFields(vwT,)
    vecs = GetDataVectors(vwT + "|__Selection", fields, {OptArray: 1})

    ret1 = RunMacro("Get Last PUDO Leg", vecs, dir)
    RunMacro("Append Arrays", ret1, &ret)

    // Trips for intermediate PUDO stops
    stops = {2}
    for stopNo in stops do
        filter = printf("%s >= %lu and nz(%s) = 0", {pudoFld, stopNo, stopsFld})
        n = SelectByQuery("__Selection", "several", "Select * where " + filter,)
        if n > 0 then do
            vecs = GetDataVectors(vwT + "|__Selection", fields, {OptArray: 1})
            ret2 = RunMacro("Get Intermediate PUDO Leg", vecs, dir, stopNo)
            RunMacro("Append Arrays", ret2, &ret)
        end
    end
    Return(CopyArray(ret))
endMacro


Macro "Process Tours with Stops"(opts)
    vwT = opts.ToursView
    dir = opts.Direction
    if dir = 'Forward' then
        pudoFld = "NumDropoffs"
    else
        pudoFld = "NumPickups"
    stopsFld = "N" + dir + "Stops"
    
    filter = printf("nz(%s) = 0 and %s > 0", {pudoFld, stopsFld})
    SetView(vwT)
    n = SelectByQuery("__Selection", "several", "Select * where " + filter,)
    {fields, specs} = GetFields(vwT,)
    vecs = GetDataVectors(vwT + "|__Selection", fields, {OptArray: 1})
    
    // Trip records for Home/Dest to first stop
    ret = RunMacro("Get First Stop Leg", vecs, dir)

    // Trip records for last PUDO stop to last stop
    ret1 = RunMacro("Get Last Stop Leg", vecs, dir)
    RunMacro("Append Arrays", ret1, &ret)

    // Trips for intermediate stops
    /*stopNo = 2
    filter = printf("nz(%s) = 0 and %s >= %lu", {pudoFld, stopsFld, stopNo})
    n = SelectByQuery("__Selection", "several", "Select * where " + filter,)
    if n > 0 then do
        vecs = GetDataVectors(vwT + "|__Selection", fields, {OptArray: 1})
        ret2 = RunMacro("Get Intermediate Stop Leg", vecs, dir, stopNo)
        RunMacro("Append Arrays", ret2, &ret)
    end*/
    
    Return(CopyArray(ret))
endMacro


Macro "Process Tours with Pudo and Stops"(opts)
    vwT = opts.ToursView
    {fields, specs} = GetFields(vwT,)
    dir = opts.Direction
    if dir = 'Forward' then do
        pudoFld = "NumDropoffs"
        flagFld = "IsStopBeforeDropoffs"
    end
    else do
        pudoFld = "NumPickups"
        flagFld = "IsStopBeforePickups"
    end
    stopsFld = "N" + dir + "Stops"
    
    // ********** Case 1: Stops before PUDO
    filter = printf("%s > 0 and %s > 0 and %s = 1", {pudoFld, stopsFld, flagFld})
    SetView(vwT)
    n = SelectByQuery("__Selection", "several", "Select * where " + filter,)
    if n > 0 then do
        vecs = GetDataVectors(vwT + "|__Selection", fields, {OptArray: 1})
        // Trip from origin to stop
        ret = RunMacro("Get First Stop Leg", vecs, dir)

        // Trip from stop to pudo
        ret1 = RunMacro("Get Stop to PUDO Leg", vecs, dir)
        RunMacro("Append Arrays", ret1, &ret)

        // Trip from pudo to dest
        ret2 = RunMacro("Get Last PUDO Leg", vecs, dir)
        RunMacro("Append Arrays", ret2, &ret)
    end

    // ********** Case 2: Stops after PUDO
    filter = printf("%s > 0 and %s > 0 and %s = 0", {pudoFld, stopsFld, flagFld})
    SetView(vwT)
    n = SelectByQuery("__Selection", "several", "Select * where " + filter,)
    if n > 0 then do
        vecs = GetDataVectors(vwT + "|__Selection", fields, {OptArray: 1})
        // Trip from origin to pudo
        ret3 = RunMacro("Get First PUDO Leg", vecs, dir)
        RunMacro("Append Arrays", ret3, &ret)

        // Trip from pudo to stop
        ret4 = RunMacro("Get PUDO to Stop Leg", vecs, dir)
        RunMacro("Append Arrays", ret4, &ret)

        // Trip from stop to dest
        ret5 = RunMacro("Get Last Stop Leg", vecs, dir)
        RunMacro("Append Arrays", ret5, &ret)
    end
    
    // Trips for intermediate PUDO stops (for case where there are two PUDO stops)
    stopNo = 2
    filter = printf("%s >= %lu and %s > 0", {pudoFld, stopNo, stopsFld})
    n = SelectByQuery("__Selection", "several", "Select * where " + filter,)
    if n > 0 then do
        vecs = GetDataVectors(vwT + "|__Selection", fields, {OptArray: 1})
        ret6 = RunMacro("Get Intermediate PUDO Leg", vecs, dir, stopNo)
        RunMacro("Append Arrays", ret6, &ret)
    end

    Return(CopyArray(ret))
endMacro


Macro "Get Basic Trip Info"(vecs, dir)
    n = vecs.TourID.length
    ret = null
    ret.TourID = v2a(vecs.TourID)
    ret.PerID = v2a(vecs.PerID)
    ret.HID = v2a(vecs.HID)
    ret.HTAZ = v2a(vecs.HTAZ)
    ret.TourType = v2a(vecs.TourType)
    ret.TourPurpose = v2a(vecs.TourPurpose)
    ret.Direction = v2a(Vector(n, "String", {Constant: left(dir, 1)}))
    ret.Assign = v2a(vecs.("Assign" + dir + "Half"))
    ret.TOD = v2a(vecs.("TOD" + dir))
    Return(ret)
endMacro


Macro "Get Direct Leg"(vecs, dir)
    n = vecs.TourID.length
    vHome = Vector(n, "String", {Constant: 'Home'})
    vLegNo = Vector(n, "Short", {Constant: 1})
    if dir = 'Forward' then do
        vOPurp = vHome
        vDPurp = vecs.TourPurpose
        vOrig = vecs.Origin
        vDest = vecs.Destination
        vOrigDep = vecs.TourStartTime
        vDestArr = vecs.DestArrTime
        vDestDep = if Lower(vecs.TourPurpose) = 'pickup' then vecs.DepPickup1 else vecs.DestDepTime // First leg of a pickup only tour
    end
    else do
        vOPurp = vecs.TourPurpose
        vDPurp = vHome
        vOrig = vecs.Destination
        vDest = vecs.Origin
        vOrigDep = vecs.DestDepTime
        vDestArr = vecs.TourEndTime
        vDestDep = Vector(n, "Long",)
    end

    vTripPurp = if Lower(vecs.TourPurpose) = 'dropoff' or Lower(vecs.TourPurpose) = 'pickup' then
                    'HBO'
                else
                    'HB' + vecs.TourPurpose
    
    // Mode Info
    if TypeOf(vecs.Mode) <> 'vector' then do
        vMode = vecs.(dir + 'Mode')
        vModeCode = vecs.(dir + 'ModeCode')
    end
    else do
        vMode = vecs.Mode
        vModeCode = vecs.ModeCode
    end

    // Get output arrays
    retArr = RunMacro("Get Basic Trip Info", vecs, dir)
    retArr.Origin = v2a(vOrig)
    retArr.Destination = v2a(vDest)
    retArr.ModeCode = v2a(vModeCode)
    retArr.Mode = v2a(vMode)
    retArr.TripPurpose = v2a(vTripPurp) // e.g. HBWork
    retArr.OrigPurpose = v2a(vOPurp)
    retArr.DestPurpose = v2a(vDPurp)
    retArr.LegNo = v2a(vLegNo)
    retArr.OrigDep = v2a(vOrigDep)
    retArr.DestArr = v2a(vDestArr)
    retArr.DestDep = v2a(vDestDep)
    Return(CopyArray(retArr))
endMacro


Macro "Get First PUDO Leg"(vecs, dir)
    n = vecs.TourID.length
    vLegNo = Vector(n, "Short", {Constant: 1})
    if dir = "Forward" then do
        vOrig = vecs.Origin
        vModeCode = Vector(n, "Short", {Constant: 2})
        vMode = Vector(n, "String", {Constant: "Carpool"})
        vOPurp = Vector(n, "String", {Constant: "Home"})
        vDPurp = Vector(n, "String", {Constant: "DropOff"})
        vTPurp = Vector(n, "String", {Constant: "HBO"})
        vPudoTAZ = vecs.DropoffTAZ1
        vOrigDep = vecs.TourStartTime
        vDestArr = vecs.ArrDropOff1
        vDestDep = vecs.DepDropOff1
    end
    else do
        vOrig = vecs.Destination
        vModeCode = Vector(n, "Short", {Constant: 1})
        vMode = Vector(n, "String", {Constant: "DriveAlone"})
        vOPurp = vecs.TourPurpose
        vDPurp = Vector(n, "String", {Constant: "PickUp"})
        vTPurp = Vector(n, "String", {Constant: "NHBWork"})
        vPudoTAZ = vecs.PickupTAZ1
        vOrigDep = vecs.DestDepTime
        vDestArr = vecs.ArrPickUp1
        vDestDep = vecs.DepPickUp1
    end
    retArr = RunMacro("Get Basic Trip Info", vecs, dir)
    retArr.Origin = v2a(vOrig)
    retArr.Destination = v2a(vPudoTAZ)
    retArr.ModeCode = v2a(vModeCode)
    retArr.Mode = v2a(vMode)
    retArr.OrigPurpose = v2a(vOPurp)
    retArr.DestPurpose = v2a(vDPurp)
    retArr.TripPurpose = v2a(vTPurp)
    retArr.LegNo = v2a(vLegNo)
    retArr.OrigDep = v2a(vOrigDep)
    retArr.DestArr = v2a(vDestArr)
    retArr.DestDep = v2a(vDestDep)
    Return(CopyArray(retArr))
endMacro


Macro "Get Last PUDO Leg"(vecs, dir)
    n = vecs.TourID.length
    if dir = "Forward" then do
        vDest = vecs.Destination
        vModeCode = Vector(n, "Short", {Constant: 1})
        vMode = Vector(n, "String", {Constant: "DriveAlone"})
        vOPurp = Vector(n, "String", {Constant: "DropOff"})
        vDPurp = vecs.TourPurpose
        vTPurp = Vector(n, "String", {Constant: "NHBWork"})
        vPudoTAZ1 = vecs.DropoffTAZ1
        vPudoTAZ2 = vecs.DropoffTAZ2
        vPudo = nz(vecs.NumDropOffs)
        vStops = nz(vecs.NForwardStops)
        vOrigDep = if vPudoTAZ2 <> null then vecs.DepDropoff2 else vecs.DepDropoff1
        vDestArr = vecs.DestArrTime
        vDestDep = vecs.DestDepTime
    end
    else do
        vDest = vecs.Origin
        vModeCode = Vector(n, "Short", {Constant: 2})
        vMode = Vector(n, "String", {Constant: "Carpool"})
        vOPurp = Vector(n, "String", {Constant: "PickUp"})
        vDPurp = Vector(n, "String", {Constant: "Home"})
        vTPurp = Vector(n, "String", {Constant: "HBO"})
        vPudoTAZ1 = vecs.PickupTAZ1
        vPudoTAZ2 = vecs.PickupTAZ2
        vPudo = nz(vecs.NumPickups)
        vStops = nz(vecs.NReturnStops)
        vOrigDep = if vPudoTAZ2 <> null then vecs.DepPickUp2 else vecs.DepPickUp1
        vDestArr = vecs.TourEndTime
        vDestDep = Vector(n, "Long",)
    end
    vLastPUDO = if vPudoTAZ2 <> null then vPudoTAZ2 else vPudoTAZ1
    vLegNo = if Lower(vecs.TourPurpose) = 'pickup' then vPudo else vPudo + vStops + 1 // Because the first pickup stop is already the main destination

    retArr = RunMacro("Get Basic Trip Info", vecs, dir)
    retArr.Origin = v2a(vLastPUDO)
    retArr.Destination = v2a(vDest)
    retArr.ModeCode = v2a(vModeCode)
    retArr.Mode = v2a(vMode)
    retArr.OrigPurpose = v2a(vOPurp)
    retArr.DestPurpose = v2a(vDPurp)
    retArr.TripPurpose = v2a(vTPurp)
    retArr.LegNo = v2a(vLegNo)
    retArr.OrigDep = v2a(vOrigDep)
    retArr.DestArr = v2a(vDestArr)
    retArr.DestDep = v2a(vDestDep)
    Return(CopyArray(retArr))
endMacro


Macro "Get Intermediate PUDO Leg"(vecs, dir, stopNo)
    n = vecs.TourID.length
    vModeCode = Vector(n, "Short", {Constant: 2})
    vMode = Vector(n, "String", {Constant: "Carpool"})
    vTPurp = Vector(n, "String", {Constant: "NHBO"})
    if dir = "Forward" then do
        oFld = "DropoffTAZ" + String(stopNo - 1)
        dFld = "DropoffTAZ" + String(stopNo)
        vPurp = Vector(n, "String", {Constant: "DropOff"})
        vFlag = nz(vecs.IsStopBeforeDropoffs)
        vOrigDep = vecs.("DepDropOff" + String(stopNo - 1))
        vDestArr = vecs.("ArrDropOff" + String(stopNo))
        vDestDep = vecs.("DepDropOff" + String(stopNo))
    end
    else do
        oFld = "PickupTAZ" + String(stopNo - 1)
        dFld = "PickupTAZ" + String(stopNo)
        vPurp = Vector(n, "String", {Constant: "PickUp"})
        vFlag = nz(vecs.IsStopBeforePickups)
        vOrigDep = vecs.("DepPickUp" + String(stopNo - 1))
        vDestArr = vecs.("ArrPickUp" + String(stopNo))
        vDestDep = vecs.("DepPickUp" + String(stopNo))
    end
    vLegNo = if Lower(vecs.TourPurpose) = 'pickup' then stopNo - 1    // Because the first pickup stop is already the main destination
             else vFlag + stopNo
    
    retArr = RunMacro("Get Basic Trip Info", vecs, dir)
    retArr.Origin = v2a(vecs.(oFld))
    retArr.Destination = v2a(vecs.(dFld))
    retArr.ModeCode = v2a(vModeCode)
    retArr.Mode = v2a(vMode)
    retArr.OrigPurpose = v2a(vPurp)
    retArr.DestPurpose = v2a(vPurp)
    retArr.TripPurpose = v2a(vTPurp)
    retArr.LegNo = v2a(vLegNo)
    retArr.OrigDep = v2a(vOrigDep)
    retArr.DestArr = v2a(vDestArr)
    retArr.DestDep = v2a(vDestDep)
    Return(CopyArray(retArr))
endMacro


Macro "Get First Stop Leg"(vecs, dir)
    n = vecs.TourID.length
    stopTAZFld = "Stop" + dir + "TAZ"
    vLegNo = Vector(n, "Short", {Constant: 1})
    if dir = "Forward" then do
        vOrig = vecs.Origin
        vOPurp = Vector(n, "String", {Constant: "Home"})
        vDPurp = Vector(n, "String", {Constant: "IntermediateStop"})
        vTPurp = Vector(n, "String", {Constant: "HBOther"})
        vOrigDep = vecs.TourStartTime
        vDestArr = vOrigDep + vecs.TimeToStopF
        vDestDep = vDestArr + vecs.ForwardStopDuration
    end
    else do
        vOrig = vecs.Destination
        vOPurp = vecs.TourPurpose
        vDPurp = Vector(n, "String", {Constant: "IntermediateStop"})
        vTPurp = Vector(n, "String", {Constant: "NHB"})
        vOrigDep = vecs.DestDepTime
        vDestArr = vOrigDep + vecs.TimeToStopR
        vDestDep = vDestArr + vecs.ReturnStopDuration
    end

    // Mode Info
    if TypeOf(vecs.Mode) <> 'vector' then do
        vMode = vecs.(dir + 'Mode')
        vModeCode = vecs.(dir + 'ModeCode')
    end
    else do
        vMode = vecs.Mode
        vModeCode = vecs.ModeCode
    end

    retArr = RunMacro("Get Basic Trip Info", vecs, dir)
    retArr.Origin = v2a(vOrig)
    retArr.Destination = v2a(vecs.(stopTAZFld))
    retArr.ModeCode = v2a(vModeCode)
    retArr.Mode = v2a(vMode)
    retArr.OrigPurpose = v2a(vOPurp)
    retArr.DestPurpose = v2a(vDPurp)
    retArr.TripPurpose = v2a(vTPurp)
    retArr.LegNo = v2a(vLegNo)
    retArr.OrigDep = v2a(vOrigDep)
    retArr.DestArr = v2a(vDestArr)
    retArr.DestDep = v2a(vDestDep)
    Return(CopyArray(retArr))
endMacro


Macro "Get Last Stop Leg"(vecs, dir)
    n = vecs.TourID.length
    
    // Mode Info
    if TypeOf(vecs.Mode) <> 'vector' then do
        vMode = vecs.(dir + 'Mode')
        vModeCode = vecs.(dir + 'ModeCode')
    end
    else do
        vMode = vecs.Mode
        vModeCode = vecs.ModeCode
    end
    
    if dir = "Forward" then do
        vDest = vecs.Destination
        vOPurp = Vector(n, "String", {Constant: "IntermediateStop"})
        vDPurp = vecs.TourPurpose
        vTPurp = Vector(n, "String", {Constant: "NHB"})
        vLastStop = vecs.StopForwardTAZ
        if vecs.NumDropoffs <> null then do
            vPudo = nz(vecs.NumDropoffs)
            vStops = nz(vecs.NForwardStops)
        end
        else do
            vPudo = Vector(n, "Short", {Constant: 0})
            vStops = nz(vecs.ForwardStop)
        end
        vOrigDep = vecs.DestArrTime - vecs.TimeFromStopF
        vDestArr = vecs.DestArrTime
        vDestDep = vecs.DestDepTime
    end
    else do
        vDest = vecs.Origin
        vOPurp = Vector(n, "String", {Constant: "IntermediateStop"})
        vDPurp = Vector(n, "String", {Constant: "Home"})
        vTPurp = Vector(n, "String", {Constant: "HBOther"})
        vLastStop = vecs.StopReturnTAZ
        if vecs.NumPickups <> null then do
            vPudo = nz(vecs.NumPickups)
            vStops = nz(vecs.NReturnStops)
        end
        else do
            vPudo = Vector(n, "Short", {Constant: 0})
            vStops = nz(vecs.ReturnStop)
        end
        vMode = if vPudo > 0 then "Carpool" else vMode // Last stop leg with kids picked up from school
        vModeCode = if vPudo > 0 then 2 else vModeCode
        vOrigDep = vecs.TourEndTime - vecs.TimeFromStopR
        vDestArr = vecs.TourEndTime
        vDestDep = Vector(n, "Long",)
    end

    retArr = RunMacro("Get Basic Trip Info", vecs, dir)
    retArr.Origin = v2a(vLastStop)
    retArr.Destination = v2a(vDest)
    retArr.ModeCode = v2a(vModeCode)
    retArr.Mode = v2a(vMode)
    retArr.OrigPurpose = v2a(vOPurp)
    retArr.DestPurpose = v2a(vDPurp)
    retArr.TripPurpose = v2a(vTPurp)
    retArr.LegNo = v2a(vPudo + vStops + 1)
    retArr.OrigDep = v2a(vOrigDep)
    retArr.DestArr = v2a(vDestArr)
    retArr.DestDep = v2a(vDestDep)
    Return(CopyArray(retArr))
endMacro

/*
Macro "Get Intermediate Stop Leg"(vecs, dir, stopNo)
    n = vecs.TourID.length
    modeFld = dir + "Mode"
    modeCodeFld = dir + "ModeCode"
    vLegNo = Vector(n, "Short", {Constant: 2})
    if dir = "Forward" then do
        oFld = "StopForwardTAZ" + String(stopNo - 1)
        dFld = "StopForwardTAZ" + String(stopNo)
    end
    else do
        oFld = "StopReturnTAZ" + String(stopNo - 1)
        dFld = "StopReturnTAZ" + String(stopNo)
    end
    vPurp = Vector(n, "String", {Constant: "IntermediateStop"})
    vTPurp = Vector(n, "String", {Constant: "NHBO"})
    
    retArr = RunMacro("Get Basic Trip Info", vecs, dir)
    retArr.Origin = v2a(vecs.(oFld))
    retArr.Destination = v2a(vecs.(dFld))
    retArr.ModeCode = v2a(vecs.(modeCodeFld))
    retArr.Mode = v2a(vecs.(modeFld))
    retArr.OrigPurpose = v2a(vPurp)
    retArr.DestPurpose = v2a(vPurp)
    retArr.TripPurpose = v2a(vTPurp)
    retArr.LegNo = v2a(vLegNo)
    Return(CopyArray(retArr))
endMacro
*/

Macro "Get Stop to PUDO Leg"(vecs, dir, stopNo)
    n = vecs.TourID.length
    modeFld = dir + "Mode"
    modeCodeFld = dir + "ModeCode"
    vLegNo = Vector(n, "Short", {Constant: 2})
    if dir = "Forward" then do
        vDPurp = Vector(n, "String", {Constant: "DropOff"})
        vMode = Vector(n, "String", {Constant: "Carpool"})
        vModeCode = Vector(n, "Short", {Constant: 2})
        vPudoTAZ = vecs.DropoffTAZ1
        vOrigDep = vecs.ArrDropoff1 - vecs.TimeFromStopF
        vDestArr = vecs.ArrDropoff1
        vDestDep = vecs.DepDropoff1
    end
    else do
        vDPurp = Vector(n, "String", {Constant: "PickUp"})
        vMode = vecs.(modeFld)
        vModeCode = vecs.(modeCodeFld)
        vPudoTAZ = vecs.PickupTAZ1
        vOrigDep = vecs.ArrPickup1 - vecs.TimeFromStopR
        vDestArr = vecs.ArrPickup1
        vDestDep = vecs.DepPickup1
    end
    vPudo = if vPudoTAZ2 <> null then vPudoTAZ2 else vPudoTAZ1
    vOPurp = Vector(n, "String", {Constant: "IntermediateStop"})
    vTPurp = Vector(n, "String", {Constant: "NHBO"})
    vStopTAZ = vecs.("Stop" + dir + "TAZ")
    
    retArr = RunMacro("Get Basic Trip Info", vecs, dir)
    retArr.Origin = v2a(vStopTAZ)
    retArr.Destination = v2a(vPudoTAZ)
    retArr.ModeCode = v2a(vModeCode)
    retArr.Mode = v2a(vMode)
    retArr.OrigPurpose = v2a(vOPurp)
    retArr.DestPurpose = v2a(vDPurp)
    retArr.TripPurpose = v2a(vTPurp)
    retArr.LegNo = v2a(vLegNo)
    retArr.OrigDep = v2a(vOrigDep)
    retArr.DestArr = v2a(vDestArr)
    retArr.DestDep = v2a(vDestDep)
    Return(CopyArray(retArr))
endMacro


Macro "Get PUDO to Stop Leg"(vecs, dir, stopNo)
    n = vecs.TourID.length
    modeFld = dir + "Mode"
    modeCodeFld = dir + "ModeCode"
    if dir = "Forward" then do
        vOPurp = Vector(n, "String", {Constant: "DropOff"})
        vMode = vecs.(modeFld)
        vModeCode = vecs.(modeCodeFld)
        vPudoTAZ1 = vecs.DropoffTAZ1
        vPudoTAZ2 = vecs.DropoffTAZ2
        vPudo = nz(vecs.NumDropoffs)
        vOrigDep = if vecs.NumDropoffs = 2 then vecs.DepDropoff2 else vecs.DepDropoff1
        vDestArr = vOrigDep + vecs.TimeToStopF
        vDestDep = vDestArr + vecs.ForwardStopDuration
    end
    else do
        vOPurp = Vector(n, "String", {Constant: "PickUp"})
        vMode = Vector(n, "String", {Constant: "Carpool"})
        vModeCode = Vector(n, "Short", {Constant: 2})
        vPudoTAZ1 = vecs.PickupTAZ1
        vPudoTAZ2 = vecs.PickupTAZ2
        vPudo = nz(vecs.NumPickups)
        vOrigDep = if vecs.NumPickups = 2 then vecs.DepPickup2 else vecs.DepPickup1
        vDestArr = vOrigDep + vecs.TimeToStopR
        vDestDep = vDestArr + vecs.ReturnStopDuration
    end
    vPudoTAZ = if vPudoTAZ2 <> null then vPudoTAZ2 else vPudoTAZ1
    vDPurp = Vector(n, "String", {Constant: "IntermediateStop"})
    vTPurp = Vector(n, "String", {Constant: "NHBO"})
    vStopTAZ = vecs.("Stop" + dir + "TAZ")
    
    retArr = RunMacro("Get Basic Trip Info", vecs, dir)
    retArr.Origin = v2a(vPudoTAZ)
    retArr.Destination = v2a(vStopTAZ)
    retArr.ModeCode = v2a(vModeCode)
    retArr.Mode = v2a(vMode)
    retArr.OrigPurpose = v2a(vOPurp)
    retArr.DestPurpose = v2a(vDPurp)
    retArr.TripPurpose = v2a(vTPurp)
    retArr.LegNo = v2a(vPudo + 1)
    retArr.OrigDep = v2a(vOrigDep)
    retArr.DestArr = v2a(vDestArr)
    retArr.DestDep = v2a(vDestDep)
    Return(CopyArray(retArr))
endMacro


Macro "Extract Subtour Trip"(Args, spec)
    vwT = spec.ToursView
    dir = spec.Direction
    filter = "Subtour = 1"
    SetView(vwT)
    n = SelectByQuery("__Selection", "several", "Select * where " + filter,)
    {fields, specs} = GetFields(vwT,)
    vecs = GetDataVectors(vwT + "|__Selection", fields, {OptArray: 1})
    n = vecs.TourID.length
    if n = 0 then
        Throw("No SubTours Found. Check output tours data file.")
    
    if dir = 'Forward' then do
        vOrig = vecs.Destination
        vDest = vecs.SubTourTAZ
        vOPurp = Vector(n, "String", {Constant: "Work"})
        vDPurp = Vector(n, "String", {Constant: "Other"})
        vOrigDep = vecs.SubTourStartTime
        vDestArr = vOrigDep + vecs.SubTourForwardTT
        vDestDep = vDestArr + vecs.SubTourActDuration
    end
    else do
        vOrig = vecs.SubTourTAZ
        vDest = vecs.Destination
        vOPurp = Vector(n, "String", {Constant: "Other"})
        vDPurp = Vector(n, "String", {Constant: "Work"})
        vOrigDep = vecs.SubTourEndTime - vecs.SubTourReturnTT
        vDestArr = vecs.SubTourEndTime
        vDestDep = Vector(n, "Long",)
    end

    // Write subtour data
    retArr = null
    retArr.TourID = v2a(10000000 + vecs.TourID)
    retArr.PerID = v2a(vecs.PerID)
    retArr.HID = v2a(vecs.HID)
    retArr.HTAZ = v2a(vecs.HTAZ)
    retArr.TourType = v2a(Vector(n, "String", {Constant: "Mandatory"}))
    retArr.TourPurpose = v2a(Vector(n, "String", {Constant: "SubTour"}))
    retArr.Direction = v2a(Vector(n, "String", {Constant: left(dir, 1)}))
    retArr.Assign = v2a(vecs.("Assign" + dir + "Half"))
    retArr.Origin = v2a(vOrig)
    retArr.Destination = v2a(vDest)
    retArr.ModeCode = v2a(vecs.SubTourModeCode)
    retArr.Mode = v2a(vecs.SubTourMode)
    retArr.OrigPurpose = v2a(vOPurp)
    retArr.DestPurpose = v2a(vDPurp)
    retArr.TripPurpose = v2a(Vector(n, "String", {Constant: "NHBW"}))
    retArr.LegNo = v2a(Vector(n, "Short", {"Constant": 1}))
    retArr.OrigDep = v2a(vOrigDep)
    retArr.DestArr = v2a(vDestArr)
    retArr.DestDep = v2a(vDestDep)
    
    PeriodInfo = Args.TimePeriod
    retArr.TOD = v2a(RunMacro("Get TOD Vector", vOrigDep, PeriodInfo))
    Return(CopyArray(retArr))
endMacro


Macro "Write Trips Table"(spec)
    data = spec.Data
    nRecs = data[1][2].length
    vwOut = RunMacro("Create Empty Trip File", {ViewName: "TempTrips", NRecords: nRecs})
    
    vecsOut = null
    for item in data do
        fld = item[1]
        vecsOut.(fld) = a2v(data.(fld))
    end
    vecsOut.One = Vector(nRecs, "Short", {Constant: 1})
    vecsOut.TripCount = if (Lower(vecsOut.Mode) = 'carpool' and vecsOut.OrigPurpose <> "DropOff" and vecsOut.OrigPurpose <> "PickUp" 
                                and vecsOut.DestPurpose <> "DropOff" and vecsOut.DestPurpose <> "PickUp") then 
                            vecsOut.Assign/spec.CarpoolOccupancy
                        else if Lower(vecsOut.Mode) = 'nonhhauto' then
                            vecsOut.Assign/spec.NonHHAutoOccupancy
                        else
                            vecsOut.Assign
    vecsOut.Period = if vecsOut.TOD = "AM" or vecsOut.TOD = "PM" then vecsOut.TOD else "OP"
    vecsOut.Mode = Lower(vecsOut.Mode)
    SetDataVectors(vwOut + "|", vecsOut,)

    // Write Trip ID field
    startingTripID = spec.StartingTripID
    vTripID = Vector(nRecs, "Long", {{"Sequence", startingTripID, 1}})
    order = {TourID: "Ascending", Direction: "Ascending", LegNo: "Ascending"}
    SetDataVector(vwOut + "|", "TripID", vTripID, {SortOrder: order})
    Return(vwOut)
endMacro


Macro "Create Empty Trip File"(spec)
    flds = {{"TripID", "Integer", 12, null, "Yes"},
            {"TourID", "Integer", 12, null, "Yes"},
            {"HID", "Integer", 12, null, "Yes"},
            {"PerID", "Integer", 12, null, "Yes"},
            {"TourType", "String", 12, null, "No"},
            {"TourPurpose", "String", 8, null, "Yes"},
            {"HTAZ", "Integer", 12, null, "Yes"},
            {"Direction", "String", 1, null, "No"},
            {"LegNo", "Integer", 8, null, "No"},
            {"OrigPurpose", "String", 18, null, "No"},
            {"DestPurpose", "String", 18, null, "No"},
            {"TripPurpose", "String", 12, null, "No"},
            {"Origin", "Integer", 12, null, "Yes"},
            {"Destination", "Integer", 12, null, "Yes"},
            {"OrigDep", "Integer", 12, null, "No"},
            {"DestArr", "Integer", 12, null, "No"},
            {"DestDep", "Integer", 12, null, "No"},
            {"TOD", "String", 2, null, "No"},
            {"Period", "String", 2, null, "No"},
            {"ModeCode", "Short", 2, null, "No"},
            {"Mode", "String", 15, null, "No"},
            {"Assign", "Short", 1, null, "No"},
            {"One", "Tiny", 1, null, "No"},
            {"TripCount", "Float", 12, 2, "No"}}
    vwOut = CreateTable(spec.ViewName,, "MEM", flds)
    if spec.NRecords > 0 then
        AddRecords(vwOut,,, {{"Empty Records", spec.NRecords}})    
    Return(vwOut)
endMacro


// ========================================================================================
/*
    Non Mandatory Tour Creation
    - Joint and Solo Tour and Trip files
*/
// =========================================================================================
// Main macro to write the joint tour file
Macro "Write Joint Tour File"(Args)
    abm = RunMacro("Get ABM Manager", Args)

    // Create temporary empty output tour file and export to in memory view
    vwT = RunMacro("Create InMemory Tour File", 'Joint')

    arrsOut = null
    purps = {"Other1", "Other2", "Shop1"}
    for p in purps do
        spec = {abmManager: abm, ModelType: p, Periods: Args.TimePeriods}
        arrs = RunMacro("Get Joint Tour Segment", spec)
        RunMacro("Append Arrays", arrs, &arrsOut)
    end

    // Write Vectors and export to final table
    nRecs = arrsOut[1][2].length
    AddRecords(vwT,,,{"Empty Records": nRecs})

    vecsOut = null
    vecsOut.TourID = Vector(nRecs, "Long", {{"Sequence", 20000000, 1}})
    for arr in arrsOut do
        vecsOut.(arr[1]) = a2v(arr[2])
    end
    SetDataVectors(vwT + "|", vecsOut,)
    RunMacro("Handle MT Access Mode Issues", vwT)
    exportOpts = {"Row Order": {{"HID", "Ascending"}, {"TourStartTime", "Ascending"}} }
    ExportView(vwT + "|", "FFB", Args.NonMandatoryJointTours,, exportOpts)
    CloseView(vwT)

    Return(true)
endMacro


// =========================================================================================
// Main macro to write the joint tour file
Macro "Write Solo Tour File"(Args)
    abm = RunMacro("Get ABM Manager", Args)

    // Create temporary empty output tour file and export to in memory view
    vwT = RunMacro("Create InMemory Tour File", 'Solo')

    arrsOut = null
    purps = {"Other1", "Other2", "Other3", "Shop1", "Shop2"}
    for p in purps do
        spec = {abmManager: abm, ModelType: p, Periods: Args.TimePeriods}
        arrs = RunMacro("Get Solo Tour Segment", spec)
        RunMacro("Append Arrays", arrs, &arrsOut)
    end

    // Write Vectors and export to final table
    nRecs = arrsOut[1][2].length
    AddRecords(vwT,,,{"Empty Records": nRecs})

    vecsOut = null
    vecsOut.TourID = Vector(nRecs, "Long", {{"Sequence", 30000000, 1}})
    for arr in arrsOut do
        vecsOut.(arr[1]) = a2v(arr[2])
    end
    SetDataVectors(vwT + "|", vecsOut,)
    RunMacro("Handle MT Access Mode Issues", vwT)
    exportOpts = {"Row Order": {{"PerID", "Ascending"}, {"TourStartTime", "Ascending"}} }
    ExportView(vwT + "|", "FFB", Args.NonMandatorySoloTours,, exportOpts)
    CloseView(vwT)

    Return(true)
endMacro


//=========================================================================================================
Macro "Create InMemory Tour File"(type)
    flds = {{"HID", "Integer", 12, null, "No"},
            {"PerID", "Integer", 12, null, "No"},
            {"TourID", "Integer", 12, null, "No", "Tour ID"},
            {"TourType", "String", 13, null, "No"},
            {"TourPurpose", "String", 13, null, "No"},
            {"HTAZ", "Integer", 12, null, "No"},
            {"Origin", "Integer", 12, null, "No"},
            {"Destination", "Integer", 12, null, "No"},
            {"TourStartTime", "Integer", 12, null, "No"},
            {"DestArrTime", "Integer", 12, null, "No"},
            {"ActivityStartTime", "Integer", 12, null, "No"},
            {"ActivityDuration", "Integer", 12, null, "No"},
            {"ActivityEndTime", "Integer", 12, null, "No"},
            {"DestDepTime", "Integer", 12, null, "No"},
            {"TourEndTime", "Integer", 12, null, "No"},
            {"ForwardMode", "String", 15, null, "No"},
            {"ReturnMode", "String", 15, null, "No"},
            {"ForwardModeCode", "Integer", 10, null, "No"},
            {"ReturnModeCode", "Integer", 10, null, "No"},
            {"TODForward", "String", 2, null, "No"},
            {"TODReturn", "String", 2, null, "No"},
            {"IsAutoTour", "Integer", 2, null, "No"},
            {"ActiveTour", "Integer", 1, null, "No"},
            {"TourTag", "String", 6, null, "No"},
            {"TimeToStopF", "Real", 10, 1, "No"},
            {"TimeFromStopF", "Real", 10, 1, "No"},
            {"TimeToStopR", "Real", 10, 1, "No"},
            {"TimeFromStopR", "Real", 10, 1, "No"},
            {"AssignForwardHalf", "Short", 2, null, "No"},
            {"AssignReturnHalf", "Short", 2, null, "No"}}
    
    if Lower(type) = 'joint' then
        flds = flds + {{"TourComposition", "String", 6, null, "No"},
                        {"nAdultsOnTour", "Integer", 12, null, "No"},
                        {"nKidsOnTour", "Integer", 12, null, "No"},
                        {"HasPreSchKids", "Integer", 1, null, "No"},
                        {"HasSeniors", "Integer", 1, null, "No"},
                        {"HasWorkers", "Integer", 1, null, "No"}}
    
    vw = CreateTable("TourDiary", , "MEM", flds)
    Return(vw)
endMacro


//==========================================================================================================
Macro "Get Joint Tour Segment"(spec)
    abm = spec.abmManager
    p = spec.ModelType
    purpose = SubString(p, 1, StringLength(p) - 1) // "Other" or "Shop"
    tag = "Joint_" + p

    filter = printf("%s_Composition <> null and %s_StartTime <> null", {tag, tag})
    setInfo = abm.CreateHHSet({Filter: filter, Activate: 1})
    n = setInfo.Size
    if n = 0 then
        Throw(printf("No Valid Joint Tours of Type %s detected. Check output", {p}))

    modeMap = {Carpool: 2, Walk: 3, Bike: 4, W_Bus: 21, Other: 7}

    {flds, specs} = GetFields(abm.HHView,)
    vecs = abm.GetHHVectors(flds)
    vMode = vecs.(tag + "_Mode")
    vAutoTour = if Lower(vMode) = 'carpool' then 1 else 0
    vActiveTour = if Lower(vMode) = 'walk' or Lower(vMode) = 'bike' then 1 else 0

    hhIDFld = abm.HHID
    vTourStart = vecs.("DepTimeTo" + tag)
    vActEnd = vecs.(tag + "_StartTime") + vecs.(tag + "_Duration")
    
    arrs = null
    arrs.HID = v2a(vecs.(hhIDFld))
    arrs.PerID = v2a(vecs.(tag + "_DesignatedPerson"))
    arrs.HTAZ = v2a(vecs.TAZID)
    arrs.Origin = v2a(vecs.TAZID)
    arrs.Destination = v2a(vecs.(tag + "_Destination"))
    arrs.TourPurpose = v2a(Vector(n, "String", {Constant: purpose}))
    arrs.TourType = v2a(Vector(n, "String", {Constant: "NonMandatory"}))
    arrs.TourTag = v2a(Vector(n, "String", {Constant: p}))  
    arrs.TourStartTime = v2a(vTourStart)
    arrs.DestArrTime = v2a(vecs.("ArrTimeAt" + tag)) 
    arrs.ActivityStartTime = v2a(vecs.(tag + "_StartTime"))
    arrs.ActivityDuration = v2a(vecs.(tag + "_Duration"))
    arrs.ActivityEndTime = v2a(vActEnd) 
    arrs.DestDepTime = v2a(vecs.("DepTimeFrom" + tag))
    arrs.TourEndTime = v2a(vecs.("ArrTimeFrom" + tag)) 
    arrs.ForwardMode = v2a(vMode)
    arrs.ReturnMode = v2a(vMode)
    arrs.ForwardModeCode = arrs.ForwardMode.Map(do (f) Return(modeMap.(f)) end)
    arrs.ReturnModeCode = arrs.ReturnMode.Map(do (f) Return(modeMap.(f)) end)
    arrs.IsAutoTour = v2a(vAutoTour) 
    arrs.AssignForwardHalf = v2a(Vector(n, "Short", {Constant: 1}))
    arrs.AssignReturnHalf = v2a(Vector(n, "Short", {Constant: 1}))  
    arrs.ActiveTour = v2a(vActiveTour)
    arrs.TourComposition = v2a(vecs.(tag + "_Composition")) 
    arrs.nAdultsOnTour = v2a(vecs.(tag + "_nAdults")) 
    arrs.nKidsOnTour = v2a(vecs.(tag + "_nKids"))
    arrs.HasPreSchKids = v2a(vecs.(tag + "_HasPreSchKids")) 
    arrs.HasSeniors = v2a(vecs.(tag + "_HasSeniors")) 
    arrs.HasWorkers = v2a(vecs.(tag + "_HasWorkers"))

    PeriodInfo = spec.Periods
    vTODF = RunMacro("Get TOD Vector", vTourStart, PeriodInfo)
    vTODR = RunMacro("Get TOD Vector", vActEnd, PeriodInfo)
    arrs.TODForward = v2a(vTODF)
    arrs.TODReturn = v2a(vTODR)
    Return(CopyArray(arrs))
endMacro



//==========================================================================================================
Macro "Get Solo Tour Segment"(spec)
    abm = spec.abmManager
    p = spec.ModelType
    purpose = SubString(p, 1, StringLength(p) - 1) // "Other" or "Shop"
    tag = "Solo_" + p

    filter = printf("%s_StartTime <> null", {tag})
    setInfo = abm.CreatePersonSet({Filter: filter, Activate: 1})
    n = setInfo.Size
    if n = 0 then
        Throw(printf("No Valid Solo Tours of Type %s detected. Check output", {p}))

    modeMap = {DriveAlone: 1, Carpool: 2, Walk: 3, Bike: 4, W_Bus: 21, Other: 7}

    {flds, specs} = GetFields(abm.PersonView,)
    flds = ArrayExclude(flds, {abm.HHIDinPersonView})
    hhIDSpec = GetFieldFullSpec(abm.PersonView, abm.HHIDinPersonView)
    vecs = abm.GetPersonHHVectors(flds + {"TAZID", hhIDSpec})
    vMode = vecs.(tag + "_Mode")
    vAutoTour = if Lower(vMode) = 'drivealone' or Lower(vMode) = 'carpool' then 1 else 0
    vActiveTour = if Lower(vMode) = 'walk' or Lower(vMode) = 'bike' then 1 else 0

    personIDFld = abm.PersonID
    vTourStart = vecs.("DepTimeTo" + tag)
    vActEnd = vecs.(tag + "_StartTime") + vecs.(tag + "_Duration")

    arrs = null
    arrs.HID = v2a(vecs.(hhIDSpec))
    arrs.PerID = v2a(vecs.(personIDFld))
    arrs.HTAZ = v2a(vecs.TAZID)
    arrs.Origin = v2a(vecs.TAZID)
    arrs.Destination = v2a(vecs.(tag + "_Destination"))
    arrs.TourPurpose = v2a(Vector(n, "String", {Constant: purpose}))
    arrs.TourType = v2a(Vector(n, "String", {Constant: "NonMandatory"}))
    arrs.TourTag = v2a(Vector(n, "String", {Constant: p}))  
    arrs.TourStartTime = v2a(vTourStart)
    arrs.DestArrTime = v2a(vecs.("ArrTimeAt" + tag)) 
    arrs.ActivityStartTime = v2a(vecs.(tag + "_StartTime")) 
    arrs.ActivityDuration = v2a(vecs.(tag + "_Duration"))
    arrs.ActivityEndTime = v2a(vActEnd) 
    arrs.DestDepTime = v2a(vecs.("DepTimeFrom" + tag))
    arrs.TourEndTime = v2a(vecs.("ArrTimeFrom" + tag)) 
    arrs.ForwardMode = v2a(vMode)
    arrs.ReturnMode = v2a(vMode)
    arrs.ForwardModeCode = arrs.ForwardMode.Map(do (f) Return(modeMap.(f)) end)
    arrs.ReturnModeCode = arrs.ReturnMode.Map(do (f) Return(modeMap.(f)) end)
    arrs.IsAutoTour = v2a(vAutoTour) 
    arrs.AssignForwardHalf = v2a(Vector(n, "Short", {Constant: 1}))
    arrs.AssignReturnHalf = v2a(Vector(n, "Short", {Constant: 1}))  
    arrs.ActiveTour = v2a(vActiveTour)

    PeriodInfo = spec.Periods
    vTODF = RunMacro("Get TOD Vector", vTourStart, PeriodInfo)
    vTODR = RunMacro("Get TOD Vector", vActEnd, PeriodInfo)
    arrs.TODForward = v2a(vTODF)
    arrs.TODReturn = v2a(vTODR)
    Return(CopyArray(arrs))
endMacro


/*
    Joint and Solo NonMandatroy Trip Creation
*/
Macro "Create Joint Trip File"(Args) 
    RunMacro("Create NonMandatory Trip File", Args, {Type: 'Joint'})
    Return(true)
endMacro


Macro "Create Solo Trip File"(Args) 
    RunMacro("Create NonMandatory Trip File", Args, {Type: 'Solo'})
    Return(true)
endMacro


/*
    Creates nonmandatory trip file from nonmandatory tour file.
    Process either 'Joint' or 'Solo' tours
    Process 2 sets of tours and two directions (Forward or Return) for each set.
    - Tours without intermediate stops
    - Tours with intermediate stops
*/
Macro "Create NonMandatory Trip File"(Args, spec)
    // This requires a mandatory tour file
    type = spec.Type
    tourFile = Args.("NonMandatory" + type + "Tours")
    if !GetFileInfo(tourFile) then
        Throw("Please run the step to create the nonmandatory tour file before creating a trip file")

    objT = CreateObject("Table", tourFile)
    vw = objT.GetView()
    vwT = ExportView(vw + "|", "MEM", "ToursMem",,)
    CloseView(vw)
    objT = null

    dirs = {"Forward", "Return"}
    tfArr = {0, 1}
    arrsOut = null
    for dir in dirs do
        spec = {ToursView: vwT, Direction: dir}
            for y in tfArr do
                spec.HasStops = y
                arrs = RunMacro("Extract NM Trips", spec)
                RunMacro("Append Arrays", arrs, &arrsOut) 
            end
    end

    if Lower(type) = 'joint' then
        startID = 20000000
    else
        startID = 30000000

    vwTemp = RunMacro("Write Trips Table", {Data: arrsOut, CarpoolOccupancy: Args.NMCarpoolOccupancy, StartingTripID: startID})

    // Export to final table
    exportOpts = {"Row Order": {{"TripID", "Ascending"}} }
    ExportView(vwTemp + "|", "FFB", Args.("NonMandatory" + type + "Trips"),, exportOpts)
    CloseView(vwTemp)
    CloseView(vwT)
    Return(1)
endMacro


/*
    The code that processes the nm trips from the tours file.
    Records from tour file selected by:
     - Direction
     - Whether or not tours have intermediate stops
*/
Macro "Extract NM Trips"(spec)
    hasPudo = spec.HasPUDO
    hasStops = spec.HasStops

    opts = {ToursView: spec.ToursView, Direction: spec.Direction}
    arrs = null
    // No PUDO and no stops
    if !hasStops then                      // No stops
        arrs = RunMacro("Process NM Direct Tours", opts)
    else                                   // Stops.
        arrs = RunMacro("Process NM Tours with Stops", opts)
    
    Return(arrs)
endMacro


/* 
    Extract trips on direct tours
    Since dir is specified, generate only one trip per tour
    # trips added = # selected tours
*/
Macro "Process NM Direct Tours"(opts)
    vwT = opts.ToursView
    dir = opts.Direction
    
    // Select tours and get vectors
    filter = printf("nz(%sStop) = 0", {dir})
    SetView(vwT)
    n = SelectByQuery("__Selection", "several", "Select * where " + filter,)
    {fields, specs} = GetFields(vwT,)
    vecs = GetDataVectors(vwT + "|__Selection", fields, {OptArray: 1})
    
    // Get output arrays
    ret = RunMacro("Get Direct Leg", vecs, dir)

    // Process trip count field
    Return(CopyArray(ret))
endMacro


Macro "Process NM Tours with Stops"(opts)
    vwT = opts.ToursView
    dir = opts.Direction
    
    filter = printf("%sStop > 0", {dir})
    SetView(vwT)
    n = SelectByQuery("__Selection", "several", "Select * where " + filter,)
    {fields, specs} = GetFields(vwT,)
    vecs = GetDataVectors(vwT + "|__Selection", fields, {OptArray: 1})
    
    // Trip records for Home/Dest to first stop
    ret = RunMacro("Get First Stop Leg", vecs, dir)

    // Trip records for last PUDO stop to last stop
    ret1 = RunMacro("Get Last Stop Leg", vecs, dir)
    RunMacro("Append Arrays", ret1, &ret)

    Return(CopyArray(ret))
endMacro


/*
    Create Master trip file combinining mandatory, joint non-mandatory and solo non-mandatory trips
*/
Macro "Generate ABM Trip File"(Args)
    inputFiles = {Args.MandatoryTrips, Args.NonMandatoryJointTrips, Args.NonMandatorySoloTrips}
    
    // Concatenate all input .bin files to a temp .bin file
    tempFile = GetTempPath() + "ABMTrips.bin"
    ConcatenateFiles(inputFiles, tempFile)
    
    inDictFile = Substitute(Args.NonMandatoryJointTrips, ".bin", ".dcb",)
    outDictFile = Substitute(tempFile, ".bin", ".dcb",)
    CopyFile(inDictFile, outDictFile)

    objT1 = CreateObject("Table", tempFile)
    vwM = ExportView(objT1.GetView() + "|", "MEM", "TempABMTrips",,)
    objT = CreateObject("Table", vwM)
    
    // Create time fields (for readability)
    inFlds = {"OrigDep", "DestArr", "DestDep"}
    outFlds = inFlds.Map(do (f) Return(f + "Time") end)
    RunMacro("Create Time Fields", {TripsObj: objT, InputFields: inFlds, OutputFields: outFlds})

    // Export file to ABM_Trips
    expOpts.[Row Order] = {{"HID", "Ascending"}, {"PerID", "Ascending"}, {"OrigDep", "Ascending"}}
    ExportView(objT.GetView() + "|", "FFB", Args.ABM_Trips,, expOpts)
    objT = null

    // Change time fields to "hh:mm tt" format
    // Have to use the CC.ModifyTableObject
    obj = CreateObject("Table", Args.ABM_Trips)
    modify = CreateObject("CC.ModifyTableOperation", obj.GetView())
    for fld in outFlds do
        modify.ChangeField(fld, {Type: "Time", Format: "hh:mm tt"})
    end
    modify.Apply()

    objT = null
    CloseView(vwM)
    
    Return(true)
endMacro


Macro "Create Time Fields"(opt)
    obj = opt.TripsObj
    inFlds = opt.InputFields
    outFlds = opt.OutputFields
    flds = outFlds.Map(do (f) Return({FieldName: f, Type: "Integer", Width: 12}) end)
    obj.AddFields({Fields: flds})
    vecs = obj.GetDataVectors({FieldNames: inFlds})
    
    vecsSet = null
    for val in inFlds do
        vecsSet.(val + "Time") = if vecs.(val) >= 1440 then (vecs.(val) - 1440)*60 // Next Day
                                    else if vecs.(val) < 0 then (1440 + vecs.(val))*60  // Prev Day
                                        else vecs.(val)*60
    end
    obj.SetDataVectors({FieldData: vecsSet})
endMacro

/*
Flip forward/return modes for microtransit acces by TOD
E.g. AM network files (tnw) are drive access. This means that any AM
return trips trying to get back into the MT district won't get assigned.
No MT is available to access on the return.
Change these to walk access and walk egress. Not perfect, but this will make
sure they are assigned.

TODO: when we move to bi-directional transit networks by TOD, we can improve
this.
*/

Macro "Handle MT Access Mode Issues" (vwT)
    v_fwd_mode = GetDataVector(vwT + "|", "ForwardMode",)
    v_ret_mode = GetDataVector(vwT + "|", "ReturnMode",)
    v_fwd_tod = GetDataVector(vwT + "|", "TODForward",)
    v_ret_tod = GetDataVector(vwT + "|", "TODReturn",)
    // PM
    v_fwd_mode = if v_fwd_tod = "PM" and Left(Lower(v_fwd_mode), 3) = "mt_"
        then Substitute(v_fwd_mode, "MT_", "W_", 1)
        else v_fwd_mode
    // AM/MD/NT
    v_ret_mode = if (v_fwd_tod <> "PM") and Left(Lower(v_ret_mode), 3) = "mt_"
        then Substitute(v_ret_mode, "MT_", "W_", 1) 
        else v_ret_mode
    SetDataVector(vwT + "|", "ForwardMode", v_fwd_mode, )
    SetDataVector(vwT + "|", "ReturnMode", v_ret_mode, )
endmacro