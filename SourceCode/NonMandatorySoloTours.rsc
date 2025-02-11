// This file contains the macros for Solo Discretionary Tours. 
// The macros are all called through "SoloDiscretionaryTours.model" file
//=================================================================================================
Macro "SoloTours Setup"(Args)
    abm = RunMacro("Get ABM Manager", Args)
    flds = {{Name: "TimeAvailForSoloTours_09to18", Type: "Real", Width: 12, Decimals: 1, Description: "Available time from 9 AM to 6 PM for discretionary solo tours"},
            {Name: "MaximumFreeTime", Type: "Real", Width: 12, Decimals: 1, Description: "Max Free Time in hrs (temp field used for avail)"},
            {Name: "SoloTourPattern", Type: "String", Width: 12,   Description: "Chosen Solo Tour Pattern"},
            {Name: "NumberSoloShopTours", Type: "Short", Width: 10},
            {Name: "NumberSoloOtherTours", Type: "Short", Width: 10}
            }
    tourList = null
    tourList.Other = {1,2,3}
    tourList.Shop = {1,2}
    for spec in tourList do
        p = spec[1]
        numTourList = spec[2]
        for i in numTourList do
            tag = p + String(i)
            flds = flds + { {Name: "Solo_" + tag + "_Destination", Type: "Long", Width: 12, Description: "Activity location TAZ choice for Solo " + tag},
                            {Name: "Solo_" + tag + "_DurChoice", Type: "String", Width: 12, Description: "Activity duration choice for Solo " + tag + " in hours"},
                            {Name: "Solo_" + tag + "_Duration", Type: "Long", Width: 12, Description: "Activity duration for Solo " + tag + " in minutes"},
                            {Name: "Solo_" + tag + "_StartInt", Type: "String", Width: 12, Description: "Activity start interval for Solo " + tag + ": Format StHr - EndHr"},
                            {Name: "Solo_" + tag + "_StartTime", Type: "Long", Width: 12, Description: "Activity start time for Solo " + tag + " in minutes from midnight"},
                            {Name: "HomeToSolo_" + tag + "_TT", Type: "Real", Width: 12, Description: "Travel time home to activity for Solo " + tag},
                            {Name: "DepTimeToSolo_" + tag, Type: "Long", Width: 12, Description: "Departure time for Solo " + tag + " in minutes from midnight"},
                            {Name: "ArrTimeAtSolo_" + tag, Type: "Long", Width: 12, Description: "Arrival time at work for Solo " + tag + " in minutes from midnight"},
                            {Name: "Solo_" + tag + "_EndTime", Type: "Long", Width: 12, Description: "Activity end time for Solo " + tag + " in minutes from midnight"},
                            {Name: "Solo_" + tag + "ToHome_TT", Type: "Real", Width: 12, Description: "Travel time back home for Solo " + tag + " in minutes from midnight"},
                            {Name: "DepTimeFromSolo_" + tag, Type: "Long", Width: 12, Description: "Departure time from work after Solo " + tag + " in minutes from midnight"},
                            {Name: "ArrTimeFromSolo_" + tag, Type: "Long", Width: 12, Description: "Arrival time back home after Solo " + tag + " in minutes from midnight"},  
                            {Name: "Solo_" + tag + "_Mode", Type: "String", Width: 15, Description: "Activity mode choice for Solo " + tag}
                            }
        end
    end
    fldNames = flds.Map(do (f) Return(f.Name) end)

    abm.AddPersonFields(flds)

    // Time avail for solo tours
    //---Initilizing the time-use matrix from data after mandatory tours and joint discretionary tours
    TimeManager = RunMacro("Get Time Manager", abm)

    TimeManager.LoadTimeUseMatrix({MatrixFile: Args.JointTimeUseMatrix})
    
    RunMacro("Get Time Avail for Solo Tours", Args, abm, TimeManager)

    // Compute MC/DC Logsums
    RunMacro("NonMandatory Solo Accessibility", Args)

    Return(true)
endMacro


//==================================================================================================
Macro "Get Time Avail for Solo Tours"(Args, abm, TimeManager)
    perSpec = {ViewName: abm.PersonView, PersonID: abm.PersonID}
    opts = {PersonSpec: perSpec, PersonFillField: "TimeAvailForSoloTours_09to18", Metric: "FreeTime", StartTime: 540, EndTime: 1080}
    TimeManager.FillPersonTimeField(opts)
endMacro


//==================================================================================================
Macro "SoloTours Frequency"(Args)
    abm = RunMacro("Get ABM Manager", Args)
    objDC = CreateObject("Table", Args.NonMandatoryDestAccessibility)

    // Run Model for all HH whose SubPattern contains J and populate output fields
    obj = CreateObject("PMEChoiceModel", {ModelName: "Joint Tours Frequency"})
    obj.OutputModelFile = Args.[Output Folder] + "\\Intermediate\\SoloToursFrequency.mdl"
    obj.AddTableSource({SourceName: "PersonHH", View: abm.PersonHHView, IDField: abm.PersonID})
    obj.AddTableSource({SourceName: "DCLogsums", View: objDC.GetView(), IDField: "TAZID"})
    obj.AddPrimarySpec({Name: "PersonHH", Filter: "Lower(SubPattern) contains 'i'", OField: "TAZID"})
    obj.AddUtility({UtilityFunction: Args.SoloTourFrequencyUtility, AvailabilityExpressions: Args.SoloTourFreqAvailability})
    obj.AddOutputSpec({ChoicesField: "SoloTourPattern"})
    obj.ReportShares = 1
    obj.RandomSeed = 5699989
    ret = obj.Evaluate()
    if !ret then
        Throw("Model Run failed for Joint Tours Frequency")
    Args.[SoloTours Frequency Spec] = CopyArray(ret)
    obj = null
    
    // Writing the choice from SoloTourPattern into number of other and shop tours
    otherTour_map = { 'O1': 1, 'O2': 2, 'O3': 3, 'S1': 0, 'S2' : 0, 'O1S1': 1, 'O2S1': 2, 'No Tours': 0}
    shopTour_map =  { 'O1': 0, 'O2': 0, 'O3': 0, 'S1': 1, 'S2' : 2, 'O1S1': 1, 'O2S1': 1, 'No Tours': 0}
    vST = abm.[Person.SoloTourPattern]
    arrO = v2a(vST).Map(do (f) if f = null then Return(0) else Return(otherTour_map.(f)) end)
    arrS = v2a(vST).Map(do (f) if f = null then Return(0) else Return(shopTour_map.(f)) end)
    vecsSet = null
    vecsSet.NumberSoloOtherTours =  a2v(arrO)
    vecsSet.NumberSoloShopTours =  a2v(arrS)
    abm.SetPersonVectors(vecsSet)    
    Return(true)
endMacro


//=================================================================================================
Macro "SoloTours Destination Other"(Args)
    abm = RunMacro("Get ABM Manager", Args)
    
    //---looping for each Other tour
    pbar = CreateObject("G30 Progress Bar", "Running destination choice for three sets of solo discretionary Other tours", false, 3)
    for tourno in {1, 2, 3} do
        obj = CreateObject("PMEChoiceModel", {ModelName: "Solo Tours Destination Other"})
        obj.OutputModelFile = Args.[Output Folder] + "\\Intermediate\\SoloToursDestinationOther.dcm"
        obj.AddTableSource({SourceName: "PersonHH", View: abm.PersonHHView, IDField: abm.PersonID})
        obj.AddTableSource({SourceName: "TAZData", File: Args.DemographicOutputs, IDField: "TAZ"})
        obj.AddTableSource({SourceName: "TAZ4Ds", File: Args.AccessibilitiesOutputs, IDField: "TAZID"})
        obj.AddMatrixSource({SourceName: "AutoSkim", File: Args.HighwaySkimOP, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
        obj.AddMatrixSource({SourceName: "WalkSkim", File: Args.WalkSkim, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
        obj.AddMatrixSource({SourceName: "Intrazonal", File: Args.IZMatrix, RowIndex: "TAZ", ColIndex: "TAZ"})
        obj.AddMatrixSource({SourceName: "ModeAccessibility", File: Args.NonMandSoloModeAccessOther, RowIndex: "TAZ", ColIndex: "TAZ"})
        obj.AddPrimarySpec({Name: "PersonHH", Filter: "NumberSoloOtherTours >= " + i2s(tourno), OField: "TAZID"})
        obj.AddUtility({UtilityFunction: Args.SoloTourDestOtherUtility})
        obj.AddDestinations({DestinationsSource: "AutoSkim", DestinationsIndex: "InternalTAZ"})
        obj.AddSizeVariable({Name: "TAZData", Field: "SoloOtherSize"})
        obj.AddOutputSpec({ChoicesField: "Solo_Other" + i2s(tourno) + "_Destination"})
        obj.RandomSeed = 5798621 + tourno*(tourno - 1)
        ret = obj.Evaluate()
        if !ret then
            Throw("Model Run failed for Solo Tours Destination Other")
        pbar.Step()
    end
    pbar.Destroy()
    Return(true)
endMacro


//==========================================================================================================
Macro "SoloTours Destination Shop"(Args)
    abm = RunMacro("Get ABM Manager", Args)

    //---Looping for each Shop tour
    pbar = CreateObject("G30 Progress Bar", "Running destination choice for two sets of solo discretionary Shop tours", false, 2)
    for tourno in {1, 2} do
        obj = CreateObject("PMEChoiceModel", {ModelName: "Solo Tours Destination Shop"})
        obj.OutputModelFile = Args.[Output Folder] + "\\Intermediate\\SoloToursDestinationShop.dcm"
        obj.AddTableSource({SourceName: "PersonHH", View: abm.PersonHHView, IDField: abm.PersonID})
        obj.AddTableSource({SourceName: "TAZData", File: Args.DemographicOutputs, IDField: "TAZ"})
        obj.AddTableSource({SourceName: "TAZ4Ds", File: Args.AccessibilitiesOutputs, IDField: "TAZID"})
        obj.AddMatrixSource({SourceName: "AutoSkim", File: Args.HighwaySkimOP, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
        obj.AddMatrixSource({SourceName: "WalkSkim", File: Args.WalkSkim, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
        obj.AddMatrixSource({SourceName: "Intrazonal", File: Args.IZMatrix, RowIndex: "TAZ", ColIndex: "TAZ"})
        obj.AddMatrixSource({SourceName: "ModeAccessibility", File: Args.NonMandSoloModeAccessShop, RowIndex: "TAZ", ColIndex: "TAZ"})
        obj.AddPrimarySpec({Name: "PersonHH", Filter: "NumberSoloShopTours >= " + i2s(tourno), OField: "TAZID"})
        obj.AddUtility({UtilityFunction: Args.SoloTourDestShopUtility})
        obj.AddDestinations({DestinationsSource: "AutoSkim", DestinationsIndex: "InternalTAZ"})
        obj.AddSizeVariable({Name: "TAZData", Field: "SoloShopSize"})
        obj.AddOutputSpec({ChoicesField: "Solo_Shop" + i2s(tourno) + "_Destination"})
        obj.RandomSeed = 5899985 + tourno*6
        ret = obj.Evaluate()
        if !ret then
            Throw("Model Run failed for Solo Tours Destination Shop")
        pbar.Step()
    end
    pbar.Destroy()
    Return(true)
endMacro


//==========================================================================================================
Macro "SoloTours Scheduling"(Args)
    // Initilizing the time-use matrix from data after mandatory and joint tours
    abm = RunMacro("Get ABM Manager", Args)
    TimeManager = RunMacro("Get Time Manager", abm)
    TimeManager.LoadTimeUseMatrix({MatrixFile: Args.JointTimeUseMatrix}) // loading the existing file

    purps = {"Other1", "Shop1", "Other2", "Other3", "Shop2"}
    pbar = CreateObject("G30 Progress Bar", "Solo Tour Scheduling for Other1, Shop1, Other2, Other3 and Shop2 Tours", false, 5)
    for p in purps do
        pbar1 = CreateObject("G30 Progress Bar", "Solo Tour Scheduling (Duration, StartTime, Mode and TimeManagerUpdate) for " + p + " Tours", true, 4)

        spec = {ModelType: p, abmManager: abm, TimeManager: TimeManager}
        // Duration Model
        RunMacro("SoloTours Duration", Args, spec)
        if pbar1.Step() then
            Return()

        // Start Time Model
        RunMacro("SoloTours StartTime", Args, spec)
        if pbar1.Step() then
            Return()

        // Mode Choice Model
        RunMacro("SoloTours Mode", Args, spec)
        if pbar1.Step() then
            Return()

        // Update Time Manager
        RunMacro("Solo Update TimeManager", Args, spec)
        if pbar1.Step() then
            Return()
        
        pbar1.Destroy()
        
        if pbar.Step() then
            Return()
    end
    pbar.Destroy()

    // Write out time manager matrix
    TimeManager.ExportTimeUseMatrix(Args.SoloTimeUseMatrix)
    Return(true)
endMacro


//==========================================================================================================
Macro "SoloTours Duration"(Args, spec)
    p = spec.ModelType
    abm = spec.abmManager
    TimeManager = spec.TimeManager

    purpose = SubString(p, 1, StringLength(p) - 1) //"Other" or "Shop"
    tourNo = Right(p,1)
    tourFilter = printf("NumberSolo%sTours >= %s", {purpose, tourNo})

    personSpec = {ViewName: abm.PersonView, PersonID: abm.PersonID, Filter: tourFilter}   
    opts = {PersonSpec: personSpec, PersonFillField: "MaximumFreeTime", Metric: "MaxAvailTime", StartTime: 360, EndTime: 1380} // 0700 to 2300
    TimeManager.FillPersonTimeField(opts)

    // Get Duration Availabilities
    utilFunction = Args.("SoloTourDur" + purpose + "Utility")
    fldSpec = 'PersonHH.MaximumFreeTime'
    durAvailArray = RunMacro("Get Duration Avail", utilFunction, fldSpec)

    // Run Duration Model
    modelName = "Solo Tours Duration " + purpose
    Opts = {abmManager: abm,
            ModelName: modelName,
            ModelFile: "SoloToursDuration" + purpose + ".mdl",
            PrimarySpec: {Name: 'PersonHH', View: abm.PersonHHView, ID: abm.PersonID},
            Filter: tourFilter,
            DestField: "Solo_" + p + "_Destination", 
            Utility: utilFunction,
            Availabilities: durAvailArray,
            ChoiceField: "Solo_" + p + "_DurChoice",
            SimulatedTimeField: "Solo_" + p + "_Duration",
            AlternativeIntervalInMin: 1,
            MinimumDuration: 10,
            RandomSeed: 5999993 + 10*StringLength(p) + s2i(tourNo)
            }
    RunMacro("NonMandatory Activity Time", Args, Opts)
endMacro


//==========================================================================================================
Macro "SoloTours StartTime"(Args, spec)
    p = spec.ModelType
    abm = spec.abmManager
    TimeManager = spec.TimeManager

    purpose = SubString(p, 1, StringLength(p) - 1) // "Other" or "Shop"
    tourNo = Right(p,1)

    // Get start time availabilities based on duration and availability of persons on the tour
    soloAltTable = Args.("SoloTourStart" + purpose + "Alts")
    startTimeAlts = soloAltTable.Alternative
    tourFilter = "NumberSolo" + purpose + "Tours >= " + tourNo
    
    // Get start time availability table
    personSpec = {ViewName: abm.PersonView, PersonID: abm.PersonID, Filter: tourFilter}
    opts = {PersonSpec: personSpec, 
            DurationField: "Solo_" + p + "_Duration", 
            StartTimeAlts: startTimeAlts, 
            OutputAvailFile: Args.SoloStartAvails}
    TimeManager.GetStartTimeAvailabilities(opts)

    objA = CreateObject("Table", Args.SoloStartAvails)
    vwJ = JoinViews("SoloToursPersonData", GetFieldFullSpec(abm.PersonHHView, abm.PersonID), GetFieldFullSpec(objA.GetView(), "RecordID"),)

    // Create start time availability expressions
    stAvailArray = RunMacro("Get StartTime Avail", soloAltTable, "SoloToursPersonData")

    // Run Start Time Model
    modelName = "Solo Tours StartTime " + purpose
    StOpts = {abmManager: abm,
                ModelName: modelName,
                ModelFile: "SoloToursStart" + purpose + ".mdl",
                PrimarySpec: {Name: 'SoloToursPersonData', View: vwJ, ID: abm.PersonID},
                Filter: tourFilter,
                DestField: "Solo_" + p + "_Destination",
                Alternatives: soloAltTable,
                Utility: Args.("SoloTourStart" + purpose + "Utility"),
                Availabilities: stAvailArray,
                ChoiceField: "Solo_" + p + "_StartInt",
                SimulatedTimeField: "Solo_" + p + "_StartTime",
                SubstituteStrings: {{"<tourno>", tourNo}},
                MinimumDuration: 10,
                RandomSeed: 6099983 + StringLength(p) + s2i(tourNo)}
    RunMacro("NonMandatory Activity Time", Args, StOpts)
    if !spec.LeaveDataOpen then do
        CloseView(vwJ)
        objA = null
    end

    // Fill Activity end time (used in setting avails)
    setInfo = abm.CreatePersonSet({Filter: tourFilter, Activate: 1})
    if setInfo.Size = 0 then
        Throw("No records to run 'Activity Start Model' for Solo " + p)

    durFld = "Solo_" + p + "_Duration"
    actStartTimeFld = StOpts.SimulatedTimeField
    actEndTimeFld = "Solo_" + p + "_EndTime"
    vecs = abm.GetPersonVectors({durFld, actStartTimeFld})
    vecsSet = null
    vecsSet.(actEndTimeFld) = vecs.(durFld) + vecs.(actStartTimeFld)
    abm.SetPersonVectors(vecsSet)
endMacro


//================================================================================================
Macro "SoloTours Mode"(Args, spec)
    p = spec.ModelType
    abm = spec.abmManager
    purpose = SubString(p, 1, StringLength(p) - 1) // "Other" or "Shop"
    tourNo = Right(p,1)
    baseFilter = "Solo_" + p + "_Destination <> null"
    
    // Compute time period field
    periodDefs = Args.TimePeriods
    amStart = periodDefs.AM.StartTime
    amEnd = periodDefs.AM.EndTime
    pmStart = periodDefs.PM.StartTime
    pmEnd = periodDefs.PM.EndTime
    depTime = "Solo_" + p + "_StartTime"
    amQry = printf("(%s >= %s and %s < %s)", {depTime, String(amStart), depTime, String(amEnd)})
    pmQry = printf("(%s >= %s and %s < %s)", {depTime, String(pmStart), depTime, String(pmEnd)})
    exprStr = printf("if %s then 'AM' else if %s then 'PM' else 'OP'", {amQry, pmQry})
    depPeriod = CreateExpression(abm.PersonView, "DepPeriod", exprStr,)

    // Run Mode Choice for each time period
    tpers = {"AM", "PM", "OP"}
    for tper in tpers do
        periodFilter = printf("(DepPeriod = '%s')", {tper})
        filterStr = baseFilter + " and " + periodFilter

        MCOpts = {abmManager: abm, TourTag: p, Purpose: purpose, TimePeriod: tper, Filter: filterStr, 
                    RandomSeed: 6199987 + 100*StringLength(p) + 10*s2i(tourNo) + tpers.position(tper)}
        RunMacro("SoloTours Mode Eval", Args, MCOpts)
    end

    // Filling in the departure and arrival time based on mode taken and TOD
    RunMacro("SoloTours Mode PostProcess", Args, spec)
    DestroyExpression(GetFieldFullSpec(abm.PersonView, "DepPeriod"))
endMacro


//================================================================================================
Macro "SoloTours Mode Eval"(Args, MCOpts)
    // Inputs
    p = MCOpts.TourTag
    purpose = MCOpts.Purpose
    tod = MCOpts.TimePeriod
    abm = MCOpts.abmManager
    tourNo = Right(p,1)
    modelName = "Solo_" + purpose + "_" + tod + "_Mode"
    ptSkimFile = printf("%s\\output\\skims\\transit\\%s_w_bus.mtx", {Args.[Scenario Folder], tod})
    objTAZ = CreateObject("Table", Args.DemographicOutputs)

    activeTransitModes = RunMacro("Get Active Transit Modes", Args)
    railPresent = RunMacro("Is value in array", activeTransitModes, "Rail")
    if railPresent then do
        WalkRailSkimFile = printf("%s\\output\\skims\\transit\\%s_w_rail.mtx", {Args.[Scenario Folder], tod})
    end
    MTBusSkimFile = printf("%s\\output\\skims\\transit\\%s_mt_bus.mtx", {Args.[Scenario Folder], tod})
    MTRailSkimFile = printf("%s\\output\\skims\\transit\\%s_mt_rail.mtx", {Args.[Scenario Folder], tod})

    obj = null
    obj = CreateObject("PMEChoiceModel", {ModelName: modelName})
    obj.OutputModelFile = Args.[Output Folder] + "\\Intermediate\\SoloToursMode" + purpose + tod + ".mdl"
    obj.AddMatrixSource({SourceName: "AutoSkim", File: Args.("HighwaySkim" + tod), RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
    obj.AddMatrixSource({SourceName: "W_BusSkim", File: ptSkimFile, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
    mtPresent = RunMacro("MT Districts Exist?", Args)
    if mtPresent then obj.AddMatrixSource({SourceName: "MT_BusSkim", File: MTBusSkimFile, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
    if railPresent then do
        obj.AddMatrixSource({SourceName: "W_RailSkim", File: WalkRailSkimFile, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
        if mtPresent then obj.AddMatrixSource({SourceName: "MT_RailSkim", File: MTRailSkimFile, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
    end
    obj.AddMatrixSource({SourceName: "WalkSkim", File: Args.WalkSkim, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
    obj.AddMatrixSource({SourceName: "BikeSkim", File: Args.BikeSkim, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
    obj.AddTableSource({SourceName: "TAZData", View: objTAZ.GetView(), IDField: "TAZ"})
    obj.AddTableSource({SourceName: "PersonHH", View: abm.PersonHHView, IDField: abm.PersonID})
    obj.AddPrimarySpec({Name: "PersonHH", Filter: MCOpts.Filter, OField: "TAZID", DField: "Solo_" + p + "_Destination"})

    // Filter out modes that aren't present
    ret = RunMacro("Filter Mode Utility Spec", {
        util: Args.("SoloTourMode" + purpose + "Utility"),
        avail: Args.("SoloTourMode" + purpose + "Avail"),
        Args: Args
    })
    util = ret.Utility
    avail = ret.Availability

    utilOpts = null
    utilOpts.UtilityFunction = util
    utilOpts.SubstituteStrings = {{"<tourno>", tourNo}}
    utilOpts.AvailabilityExpressions = avail
    obj.AddUtility(utilOpts)

    obj.AddOutputSpec({ChoicesField: "Solo_" + p + "_Mode"})
    obj.ReportShares = 1
    obj.RandomSeed = MCOpts.RandomSeed
    ret = obj.Evaluate()
    if !ret then
        Throw("Model Run failed for Solo Tours Mode " + purpose + tourNo)

    Args.(modelName + " Spec") = CopyArray(ret)
endMacro


//==========================================================================================================
// The macro calculates travel times to and from the home to the activity.
// Finally it establishes departure time from home and arrival time back home to inform future model decisions.
Macro "SoloTours Mode PostProcess"(Args, spec)
    p = spec.ModelType
    abm = spec.abmManager
    purpose = SubString(p, 1, StringLength(p) - 1) // "Other" or "Shop"
    tourNo = Right(p,1)

    filter = "NumberSolo" + purpose + "Tours >= " + tourNo
    setInfo = abm.CreatePersonSet({Filter: filter, Activate: 1})
    if setInfo.Size = 0 then
        Return()

    // Get relevant time fields
    tag = "Solo_" + p
    durFld = tag + "_Duration"
    destFld = tag + "_Destination"
    actStartTimeFld = tag + "_StartTime"
    actEndTimeFld = tag + "_EndTime"
    homeToActTTFld = "HomeTo" + tag + "_TT"
    actToHomeTTFld = tag + "ToHome_TT"
    modeFld = tag + "_Mode"
    
    // Fill travel times from home to activity and activity to home
    fillSpec = {View: abm.PersonHHView, OField: "TAZID", DField: destFld, FillField: homeToActTTFld, 
                Filter: filter, ModeField: modeFld, DepTimeField: actStartTimeFld}
    RunMacro("Fill Travel Times", Args, fillSpec)

    fillSpec = {View: abm.PersonHHView, OField: destFld, DField: "TAZID", FillField: actToHomeTTFld, 
                Filter: filter, ModeField: modeFld, DepTimeField: actEndTimeFld}
    RunMacro("Fill Travel Times", Args, fillSpec)
    
    // Fill departure time from home and arrival time after activity
    arrTimeBackHomeFld = "ArrTimeFrom" + tag
    depTimeFromHomeFld = "DepTimeTo" + tag
    arrTimeAtDestFld = "ArrTimeAt" + tag
    depTimeFromDestFld = "DepTimeFrom" + tag
    
    vecs = abm.GetPersonVectors({actStartTimeFld, actEndTimeFld, homeToActTTFld, actToHomeTTFld})
    vecsSet = null
    vecsSet.(arrTimeBackHomeFld) = vecs.(actEndTimeFld) + Round(vecs.(actToHomeTTFld),0)
    vecsSet.(depTimeFromHomeFld) = vecs.(actStartTimeFld) - Round(vecs.(homeToActTTFld), 0)
    vecsSet.(arrTimeAtDestFld) = vecs.(actStartTimeFld)
    vecsSet.(depTimeFromDestFld) = vecs.(actEndTimeFld)
    abm.SetPersonVectors(vecsSet)
endMacro


//==========================================================================================================
Macro "Solo Update TimeManager"(Args, spec)
    p = spec.ModelType
    abm = spec.abmManager
    TimeManager = spec.TimeManager

    purpose = SubString(p, 1, StringLength(p) - 1) // "Other" or "Shop"
    tourNo = Right(p,1)
    
    freqFld = "NumberSolo" + purpose + "Tours"
    fldsToExport = {abm.PersonID, freqFld, "DepTimeToSolo_" + p, "ArrTimeFromSolo_" + p}
    vwTemp = ExportView(abm.PersonView + "|", "MEM", "TempPerson", fldsToExport,)
    tourOpts = {ViewName: vwTemp,
                Filter: freqFld + " >= " + tourNo,
                PersonID: abm.PersonID,
                Departure: "DepTimeToSolo_" + p,
                Arrival: "ArrTimeFromSolo_" + p}
    TimeManager.UpdateMatrixFromTours(tourOpts)

    CloseView(vwTemp)
endMacro
