/*
    ABM Manager Utilities:
    'Get ABM Manager' returns the ABM manager object
    If the object does not exist, the object is created
    If the views pertaining to the object are closed (For e.g. the flowchart automatically closes all view when you start), the views are added
*/
Macro "Get ABM Manager"(Args)
    abm = Args.[ABM Manager]
    if abm = null then do
        abm = CreateObject("ABM_Manager")
        Args.[ABM Manager] = abm
    end
    
    if !abm.IsHHDataLoaded() then
        abm.SetHouseholdData({File: Args.Households, ID: "HouseholdID"})

    if !abm.IsPersonDataLoaded() then
        abm.SetPersonData({File: Args.Persons, ID: "PersonID", HHID: "HouseholdID"})

    if !abm.IsPersonHHViewOpen() then
        abm.CreatePersonHHView()

    abm.ClearHHSets()
    abm.ClearPersonSets()
    Return(abm)
endMacro


/*
    ABM Manager Utilities:
    'Export ABM Data' exports the in-memory tables from the abm manager object back to the population and household files
*/
Macro "Export ABM Data"(Args, opts)
    abm = Args.[ABM Manager]
    if abm = null or TypeOf(abm) <> "object" then
        Return()
    
    // Export HH Data
    if abm.IsHHDataLoaded() then do
        hhOpts = {File: Args.Households}
        if opts.HHFields <> null then
            hhOpts = hhOpts + {Fields: opts.HHFields}
        if opts.HHFilter <> null then
            hhOpts = hhOpts + {Filter: opts.HHFilter}
        if opts.UseActiveHHSet = 1 then
            hhOpts = hhOpts + {UseActiveSet: 1}
        abm.ExportHHView(hhOpts)
    end

    // Export Person Data
    if abm.IsPersonDataLoaded() then do
        pOpts = {File: Args.Persons}
        if opts.PersonFields <> null then
            pOpts = pOpts + {Fields: opts.PersonFields}
        if opts.PersonFilter <> null then
            pOpts = pOpts + {Filter: opts.PersonFilter}
        if opts.UseActivePersonSet = 1 then
            pOpts = pOpts + {UseActiveSet: 1}
        abm.ExportPersonView(pOpts)
    end
endMacro


/*
    ABM Manager Utilities:
    Macro to create an empty ABM object and return it. Will be called by the flowchart plugin macros.
*/
Macro "Create ABM Object"
    abm = CreateObject("ABM_Manager")
    Return(abm)
endMacro


/*
    ABM Preprocessor. 
    Remove and add all ABM related fields to the In-Memory Person and HH tables.
    Ideally called in each feedback loop.
*/
Macro "ABM Preprocess"(Args)
    // Person File
    abm = RunMacro("Get ABM Manager", Args)
    flds = {{Name: "AttendDaycare", Type: "Short", Width: 2, Description: "Does child attend daycare?|1: Yes|2: No.|Filled for Age < 5"},
            {Name: "AttendSchool", Type: "Short", Width: 2, Description: "Does child attend school on given day?|1: Yes|2: No.|Filled for Age >= 5 and Age <= 18"},
            {Name: "WorkTAZ", Type: "Integer", Width: 12, Description: "WorkTAZ|Filled if WorkIndustry <= 10"},
            {Name: "SchoolTAZ", Type: "Integer", Width: 12, Description: "SchoolTAZ|Filled for school age [5, 18] students OR if AttendDaycare = 1"},
            {Name: "HometoWorkTime", Type: "Real", Width: 12, Decimals: 2, Description: "Home to work travel time"},
            {Name: "HometoUnivTime", Type: "Real", Width: 12, Decimals: 2, Description: "Home to univ travel time"},
            {Name: "HometoSchoolTime", Type: "Real", Width: 12, Decimals: 2, Description: "Home to school/daycare travel time"},
            {Name: "WorktoHomeTime", Type: "Real", Width: 12, Decimals: 2, Description: "Work to home travel time"},
            {Name: "UnivtoHomeTime", Type: "Real", Width: 12, Decimals: 2, Description: "Univ to home travel time"},
            {Name: "SchooltoHomeTime", Type: "Real", Width: 12, Decimals: 2, Description: "School/daycare to home travel time"},
            {Name: "NumberWorkTours", Type: "Integer", Width: 12, Description: "Filled if TravelToWork = 1"},
            {Name: "NumberUnivTours", Type: "Integer", Width: 12, Description: "Filled if AttendUniv = 1"}
           }
    
    purps = {"Work", "Univ", "School"}
    for p in purps do
        if p = "School" then
            freqArr = {""}
        else
            freqArr = {"1", "2"}
        
        for i in freqArr do
            flds = flds + {{Name: p + i + "_DurChoice", Type: "String", Width: 12, Description: "Activity duration choice for tour" + i + " in hours"},
                            {Name: p + i + "_StartInt", Type: "String", Width: 12, Description: "Activity start interval for tour" + i + ": Format StHr - EndHr"},
                            {Name: p + i + "_Duration", Type: "Long", Width: 12, Description: "Activity duration for tour" + i + " in minutes"},
                            {Name: p + i + "_StartTime", Type: "Long", Width: 12, Description: "Activity start time for tour" + i + " in minutes from midnight"}
                          }
        end
    end

    flds = flds + {{Name: "WorkMode", Type: "String", Width: 12},
                    {Name: "UnivMode", Type: "String", Width: 12},
                    {Name: "SchoolForwardMode", Type: "String", Width: 12},
                    {Name: "SchoolReturnMode", Type: "String", Width: 12},
                    {Name: "WorkModeCode", Type: "Short", Width: 2, Description: "1. DriveAlone|2. Carpool|3. Walk|4. Bike|5. PTWalk|6. PTDrive|7. Other|8. SchoolBus"},
                    {Name: "UnivModeCode", Type: "Short", Width: 2, Description: "1. DriveAlone|2. Carpool|3. Walk|4. Bike|5. PTWalk|6. PTDrive|7. Other|8. SchoolBus"},
                    {Name: "SchoolForwardModeCode", Type: "Short", Width: 2, Description: "1. DriveAlone|2. Carpool|3. Walk|4. Bike|5. PTWalk|8. SchoolBus"},
                    {Name: "SchoolReturnModeCode", Type: "Short", Width: 2, Description: "1. DriveAlone|2. Carpool|3. Walk|4. Bike|5. PTWalk|8. SchoolBus"},
                    {Name: "VehiclePriority", Type: "Short", Description: "A lower value indicates a higher priority of being allocated a vehicle"},
                    {Name: "VehicleAvail", Type: "Short", Description: "Temporary field that determines if a vehicle is available prior to mode choice"},
                    {Name: "VehicleUsed", Type: "Short", Description: "Temporary field that determines if a vehicle was used by this person for mode choice"},
                    {Name: "NDropoffs", Type: "Short", Width: 2,  Description: "Number of kids dropped off at school by this person"},
                    {Name: "NDropoffsEnRoute", Type: "Short", Width: 2,  Description: "Number of kids dropped off at school by this person on their way to work"},
                    {Name: "DropoffPersonID", Type: "Long", Width: 12,  Description: "Dropoff person ID for school carpool"},
                    {Name: "DropoffTourFlag", Type: "String", Width: 2,  Description: "Dropoff Tour Flag|'S': Person makes separate school drop off tour|'W1': Person drops off a kid on the way to the first work tour|'W2': Person drops off a kid on the way to the second work tour"},
                    {Name: "NPickups", Type: "Short", Width: 2,  Description: "Number of kids picked up at school by this person"},
                    {Name: "PickupPersonID", Type: "Long", Width: 12,  Description: "Pickup person ID for school carpool"},
                    {Name: "PickupTourFlag", Type: "String", Width: 2,  Description: "Pickup Tour Flag|'S': Person makes separate school pick up tour|'W1': Person picks up a kid on the way from the first work tour|'W2': Person picks up a kid on the way from the second work tour"},
                    {Name: "NPickupsEnRoute", Type: "Short", Width: 2,  Description: "Number of kids picked up at school by this person on their way back from work"}
                }


    fldNames = flds.Map(do (f) Return(f.Name) end)
    abm.DropPersonFields(fldNames)
    abm.AddPersonFields(flds)

    // HH File
    flds = {{Name: "NLicensed", Type: "Short", Description: "Number of people in the HH with driving license. Outcome of 'Driver License' model"},
            {Name: "NSchoolKids", Type: "Integer", Width: 12, Description: "Number of kids in the HH who go to school on the given day (AttendSchool = 1)"},
            {Name: "VehiclesUsed", Type: "Short", Description: "Temporary field to indicate vehicles used up by prior mode choice allocations"},
            {Name: "VehiclesRem", Type: "Short", Description: "Temporary field to indicate vehicles left after mode choice allocations"},
            {Name: "NSchoolDropoffs", Type: "Integer", Width: 12, Description: "Number of kids in the HH who choose school drop offs"},
            {Name: "NSchoolPickups", Type: "Integer", Width: 12, Description: "Number of kids in the HH who choose school pick ups"}}
    fldNames = flds.Map(do (f) Return(f.Name) end)
    abm.DropHHFields(fldNames)
    abm.AddHHFields(flds)
    
    Return(true)
endmacro


/*
    Macro to compute size variable term.
    Input: A PME table (column format) with columns 'Variable' and 'Coefficient' or an eqvivalent options array.
           Also a table object which has the variables
           An output field to write into 
    Output: The output field name is added to the table and filled by the equation.
*/
Macro "Compute Size Variable"(opt)
    obj = opt.TableObject
    vw = obj.GetView()

    if opt.FillField = null and opt.NewOutputField = null then
        Throw("Please specify either the 'FillField' or 'NewOutputField' option for 'Compute Size Variable' macro")

    if opt.FillField <> null and opt.NewOutputField <> null then
        Throw("Please specify only one of the 'FillField' or 'NewOutputField' option for 'Compute Size Variable' macro")

    // Add new field if required
    if opt.NewOutputField <> null then do
        outFld = opt.NewOutputField
        newFlds = {{FieldName: outFld, Type: "real", Width: 12, Decimals: 2}}
        obj.AddFields({Fields: newFlds})
    end
    else do
        flds = obj.GetFieldNames()
        if flds.position(opt.FillField) = 0 then
            Throw("Field '" + opt.FillField + "' does not exist in the view sent to 'Compute Size Variable'")
        outFld = opt.FillField
    end

    // Get the list of fields and the coefficients
    eqn = opt.Equation
    flds = eqn.Variable
    coeffs = eqn.Coefficient
    
    // Compute
    vecs = GetDataVectors(vw + "|", flds,)
    vOut = Vector(vecs[1].Length, "Real", {{"Constant", 0.0}})
    for i = 1 to coeffs.length do
        coeff = coeffs[i]
        
        if opt.ExponentiateCoeffs = 1 then
            coeff = exp(coeff)
        
        if coeff <> null then
            vOut = vOut + nz(vecs[i])*coeff
    end
    SetDataVector(vw + "|", outFld, vOut,)
    obj = null
endMacro


// Given the choice of hourly time intervals, simulate a time within that interval
// Additionally use information from an hourly profile table if provided.
// For example the hourly profile can state: 30% probability for '0-5min', 40% probability for "5-55 min" and 30% for "55-60 min"
Macro "Simulate Time"(opt)

    viewSet = opt.ViewSet
    minuteFlag = opt.AlternativeIntervalInMin

    // Parse hourly profile if provided and create categories and wgts arrays. Do basic checks.
    hrlyProfile = opt.HourlyProfile
    if hrlyProfile <> null and minuteFlag = 1  then
        Throw("Do not specify both options 'AlternativeIntervalInMin' and 'HourlyProfile' to 'Simulate Time' macro")

    if hrlyProfile <> null then do
        for i = 1 to hrlyProfile.Length do
            categories = categories + {hrlyProfile[i][1]}
            wgts = wgts + {hrlyProfile[i][2]}
            arr = ParseString(hrlyProfile[i][1], " -")
            if arr[1] = null or s2i(arr[1]) < 0 or s2i(arr[1]) >= 60 then
                Throw("Error in hourly profile specification")
            if arr[2] = null or ( s2i(arr[2]) <= s2i(arr[1]) ) then
                Throw("Error in hourly profile specification")
        end
    end

    vInput = GetDataVector(viewSet, opt.InputField,)
    vUniform = RandSamples(vInput.Length, "Uniform",)
    temp = v2a(vInput).Map(do (f) Return(ParseString(f, " -")) end)
    // vL and vR are vectors containing the main choice interval start and end values
    vL = a2v(temp.Map(do (f) Return(if f = null then null else s2i(f[1])) end))
    vR = a2v(temp.Map(do (f) Return(if f = null then null else s2i(f[2])) end))

    if hrlyProfile = null and minuteFlag <> 1 then do
        vMinute = vUniform*(vR-vL)*60
        vChoice = Round(vL*60 + vMinute, 0)
    end
    else if hrlyProfile = null and minuteFlag = 1 then do
        vMinute = vUniform*(vR-vL)
        vChoice = Round(vL + vMinute, 0)
    end
    else do
        params = null
        params.population = categories
        params.weight = wgts
        vSamples = RandSamples(vInput.Length, "Discrete", params) // The chosen hourly profile categories
        
        temp = v2a(vSamples).Map(do (f) Return(ParseString(f, " -")) end)
        // vL1 and vR1 are vectors containing the chosen hourly profile choice interval start and end values
        vL1 = a2v(temp.Map(do (f) Return(if f = null then null else s2i(f[1])) end))
        vR1 = a2v(temp.Map(do (f) Return(if f = null then null else s2i(f[2])) end))
        vMinute = vL1 + vUniform*(vR1-vL1)
        vChoice = Round(vL*60 + vMinute, 0) // Minutes from midnight
    end
    SetDataVector(viewSet, opt.OutputField, vChoice,)
endMacro


Macro "Fill from matrix"(spec)
    abm = spec.abmManager
    set.Name = null
    if spec.Filter <> null then do
        set = abm.CreatePersonSet({Filter: spec.Filter, Activate: 1})
        if set.Size = 0 then
            Return()
    end
    
    mObj = CreateObject('Matrix', spec.Matrix.Name)
    mObj.SetIndex({RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
    mc = mObj.(spec.matrix.Core)
    vw = abm.PersonHHView
    oSpec = GetFieldFullSpec(vw, spec.OField)
    dSpec = GetFieldFullSpec(vw, spec.DField)
    fSpec = GetFieldFullSpec(vw, spec.FillField)
    FillViewFromMatrix(vw + "|" + set.Name, oSpec, dSpec, {{fSpec, mc}})
endMacro



// Fill travel times in the view based on origin, destination and optionally using departure time and mode.
// If mode field is not supplied, the auto skims are used
// If departure time field is not supplied, the OP skims are used
Macro "Fill Travel Times"(Args, spec)
    if spec.View = null or spec.OField = null or spec.DField = null or spec.FillField = null then
        Throw("Please specify all of 'View', 'OField', 'DField' and 'FillField' options to macro 'Fill Travel Times'")

    if TypeOf(spec.OField) <> "string" or TypeOf(spec.DField) <> "string" or TypeOf(spec.FillField) <> "string" then
        Throw("Options 'OField', 'DField' and 'FillField' to macro 'Fill Travel Times' need to be strings")

    // Define skim files: Change this spec after skims by mode are produced by the model.
    skimSpec = null
    skimSpec.Auto.AM = {File: Args.HighwaySkim, Core: "Time"}
    skimSpec.Auto.PM = {File: Args.HighwaySkim, Core: "Time"}
    skimSpec.Auto.OP = {File: Args.HighwaySkim, Core: "Time"}
    skimSpec.PTWalk.AM = {File: Args.TransitWalkSkim, Core: "Total Time"}
    skimSpec.PTWalk.PM = {File: Args.TransitWalkSkim, Core: "Total Time"}
    skimSpec.PTWalk.OP = {File: Args.TransitWalkSkim, Core: "Total Time"}
    skimSpec.PTDrive.AM = {File: Args.TransitDriveSkim, Core: "Total Time"}
    skimSpec.PTDrive.PM = {File: Args.TransitDriveSkim, Core: "Total Time"}
    skimSpec.PTDrive.OP = {File: Args.TransitDriveSkim, Core: "Total Time"}
    skimSpec.Walk = {File: Args.WalkSkim, Core: "Time"}
    skimSpec.Bike = {File: Args.BikeSkim, Core: "Time"}

    // Clear values in field first
    baseFilter = printf("(%s <> null and %s <> null)", {spec.OField, spec.DField})
    if spec.Filter <> null then
        baseFilter = printf("%s and (%s)", {baseFilter, spec.Filter})

    vw = spec.View
    SetView(vw)
    n = SelectByQuery("__BaseSet", "several", "Select * where " + baseFilter,)
    if n = 0 then
        Throw("No valid records in macro 'Fill Travel Times'.\nCheck base filer or valid values of 'OField' and 'DField'.")

    // Check for mode field. If no mode field specified, then fill field with auto times
    modeFld = spec.ModeField
    if modeFld = null then
        mainModes = {"auto"}
    else do
        mainModes = skimSpec.Map(do(f) Return(Lower(f[1])) end) // {"auto", "walk", "bike", "ptwalk", "ptdrive"}
        exprStr = printf("if Lower(%s) <> 'walk' and Lower(%s) <> 'bike' and Lower(%s) <> 'ptwalk' and Lower(%s) <> 'ptdrive' then 'auto' else Lower(%s)", 
                         {modeFld, modeFld, modeFld, modeFld, modeFld})
        modeExpr = CreateExpression(vw, "ModeGroup", exprStr,)
    end

    // Check for dep time field. If not present, no need of a mode sub-loop. Fill with OP skim values.
    timePeriods = Args.TimePeriods
    amStart = timePeriods.AM.StartTime
    amEnd = timePeriods.AM.EndTime
    pmStart = timePeriods.PM.StartTime
    pmEnd = timePeriods.PM.EndTime
    depTimeFld = spec.DepTimeField
    if depTimeFld <> null then do
        timeAllowance = '0'
        amQry = printf("(%s - %s >= %s and %s - %s < %s)", {depTimeFld, timeAllowance, String(amStart), depTimeFld, timeAllowance, String(amEnd)})
        pmQry = printf("(%s - %s >= %s and %s - %s < %s)", {depTimeFld, timeAllowance, String(pmStart), depTimeFld, timeAllowance, String(pmEnd)})
        exprStr = printf("if %s then 'AM' else if %s then 'PM' else 'OP'", {amQry, pmQry})
        depPeriod = CreateExpression(vw, "DepPeriod", exprStr,)
    end

    // Loop over the modes in the mode field
    for mode in mainModes do
        if Lower(mode) = "walk" or Lower(mode) = "bike" then
            nmMode = 1
        else
            nmMode = 0

        baseModeFilter = baseFilter
        if modeFld <> null then do
            modeFilter = printf("(%s = '%s')", {modeExpr, mode})
            baseModeFilter = printf("%s and %s", {baseFilter, modeFilter})
        end  

        if nmMode or depTimeFld = null then
            periods = {"OP"}
        else
            periods = {"AM", "PM", "OP"}
        
        for period in periods do
            if nmMode then
                skimData = skimSpec.(mode)
            else
                skimData = skimSpec.(mode).(period)
            
            // Departure time query
            finalFilter = baseModeFilter
            if depTimeFld <> null and !nmMode then do // No period skims for walk/bike. Can skip dep time query.
                depFilter = printf("(%s = '%s')", {depPeriod, period})
                finalFilter = printf("%s and %s", {baseModeFilter, depFilter})
            end

            SetView(vw)    
            qry = "Select * where " + finalFilter
            n = SelectByQuery("__FillSet", "several", qry,)
            if n = 0 then continue
            
            mObj = CreateObject("Matrix", skimData.File)
            mObj.SetIndex({RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
            mc = mObj.(skimData.Core)
            oSpec = GetFieldFullSpec(vw, spec.OField)
            dSpec = GetFieldFullSpec(vw, spec.DField)
            fSpec = GetFieldFullSpec(vw, spec.FillField)
            FillViewFromMatrix(vw + "|__FillSet", oSpec, dSpec, {{fSpec, mc}})
            mc = null
        end
    end

    if depPeriod <> null then
        DestroyExpression(GetFieldFullSpec(vw, depPeriod))
    if modeExpr <> null then
        DestroyExpression(GetFieldFullSpec(vw, modeExpr))
endMacro


Macro "Get TOD Vector"(v, PeriodInfo)
    amStart = PeriodInfo.AM.StartTime
    amEnd = PeriodInfo.AM.EndTime
    pmStart = PeriodInfo.PM.StartTime
    pmEnd = PeriodInfo.PM.EndTime
    vRet = if (v >= amStart and v < amEnd) then 'AM'
           else if (v >= pmStart and v < pmEnd) then 'PM'
           else 'OP'
    Return(CopyVector(vRet))
endMacro


// Macro that returns approximate auto travel time given dep/arr time (one OD pair only). 
// Note that exact time profile of the trip can sometimes be ignored. Two examples below.
// Ex. 1: TT for a trip arriving at 7:01 AM will be determined from AM skims although most of travel would have been in previous OP period.
// Ex. 2: TT for a trip departing at 6:59 AM will be determined from OP times although most of travel will be during the peak AM period.
Macro "Get Auto TT"(skimArgs, opt)
    orig = i2s(opt.Orig)
    dest = i2s(opt.Dest)
    time = opt.DepTime
    
    timePeriods = skimArgs.TimePeriods
    amStart = timePeriods.AM.StartTime
    amEnd = timePeriods.AM.EndTime
    pmStart = timePeriods.PM.StartTime
    pmEnd = timePeriods.PM.EndTime

    if time >= amStart and time < amEnd then       // 7AM to 9AM
        period = "AM"
    else if time >= pmStart and time < pmEnd then  // 4PM to 6PM
        period = "PM"
    else
        period = "OP"

    val = GetMatrixValue(skimArgs.(period), orig, dest)
    Return(val)
endMacro



// Given orig, stop and dest TAZ fields, compute the excess mode specific travel time as a result of making the stop.
// i.e. compute OriginToStopTime + StopToDestTime - OrigToStopTime
Macro "Calculate Detour TT"(Args, spec)
    // Open Table or use view passed
    vw = spec.ToursView
    ODInfo = spec.ODInfo
    dir = spec.Direction
    if dir = "Forward" then 
        depFld = 'TourStartTime'
    else 
        depFld = 'DestDepTime'
    
    modeFld = dir + "Mode"
    stopFld = "Stop" + dir + "TAZ"
    filter = printf("N%sStops > 0", {dir})

    // Add temporary fields
    toursObj = CreateObject("Table", vw)
    flds = {{FieldName: "OrigToStopTT"}, 
            {FieldName: "StopToDestTT"}, 
            {FieldName: "OrigToDestTT"}}
    toursObj.AddFields({Fields: flds})

    // Fill travel times from origin to stop, stop to dest and orig to dest
    opt = {View: vw, Filter: filter, OField: ODInfo.Origin, DField: stopFld, DepTimeField: depFld, ModeField: modeFld, FillField: "OrigToStopTT"}
    RunMacro("Fill Travel Times", Args, opt)

    opt = {View: vw, Filter: filter, OField: stopFld, DField: ODInfo.Destination, DepTimeField: depFld, ModeField: modeFld, FillField: "StopToDestTT"}
    RunMacro("Fill Travel Times", Args, opt)

    opt = {View: vw, Filter: filter, OField: ODInfo.Origin, DField: ODInfo.Destination, DepTimeField: depFld, ModeField: modeFld, FillField: "OrigToDestTT"}
    RunMacro("Fill Travel Times", Args, opt)
    
    flds = flds.Map(do (f) Return(f.FieldName) end) // {"OrigToStopTT", "StopToDestTT", "OrigToDestTT"}
    fillField = dir + "StopDeltaTT"
    n = toursObj.CreateSet({SetName: "BaseSet", Filter: filter})
    if n = 0 then Return()
    vecs = toursObj.GetDataVectors({FieldNames: flds})
    v = vecs.OrigToStopTT + vecs.StopToDestTT - vecs.OrigToDestTT
    toursObj.(fillField) = v

    // Drop temporary fields
    toursObj.DropFields({FieldNames: flds})
endMacro



// This macro tries to replicate the GetPrevRecord() and GetNext Record() on a vector
// This is because the formula fields or filling data using GetPrevRecord() does not work
// Assume a vector v = {1,2,3,4,56}
// Given this vector, the 'Prev' option returns {,1,2,3,4}
// Given this vector, the 'Next' option returns {2,3,4,56,}
Macro "Shift Vector"(spec)
    v = spec.Vector
    method = Lower(spec.Method)

    if Lower(TypeOf(v)) <> 'vector' then
        Throw("Option 'Vector' sent to 'Shift Vector' is not of type vector")

    if method <> 'prev' and method <> 'next' then
        Throw("Option 'Method' sent to 'Shift Vector' is neither 'Prev' nor 'Next'")

    a = v2a(v)
    if method = 'prev' then
        a2 = {} + SubArray(a, 1, a.length - 1)
    else
        a2 = SubArray(a, 2, a.length - 1) + {}

    Return(a2v(a2))
endMacro