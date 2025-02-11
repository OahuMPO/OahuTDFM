// This file contains the macros for Joint Discretionary Tours. 
//=================================================================================================
Macro "JointTours Setup"(Args)
    abm = RunMacro("Get ABM Manager", Args)

    flds = {{Name: "MaxFreeTime", Type: "Real", Description: "Max Free Time in hrs (temp field used for avail)"},
            {Name: "HHAvgFreeTime_09to18", Type: "Real", Description: "Average free time across members in the HH in hrs after mandatory tours between 9 AM and 6 PM"},
            {Name: "JointTourPattern", Type: "String", Width: 12,   Description: "Chosen Joint Tour Pattern"},
            {Name: "NumberJointShopTours", Type: "Short", Width: 10},
            {Name: "NumberJointOtherTours", Type: "Short", Width: 10}
            }
    purps = {"Other1", "Other2", "Shop1"}
    for p in purps do 
        flds = flds + {{Name: "Joint_" + p + "_Composition", Type: "String", Width: 12, Description: "Activity adult participation for Joint " + p},
                       {Name: "Joint_" + p + "_DesignatedPerson", Type: "Long", Width: 12, Description: "A designated person on Joint " + p},
                       {Name: "Joint_" + p + "_nAdults", Type: "Short", Description: "Number adults on joint tour"},
                       {Name: "Joint_" + p + "_nKids", Type: "Short", Description: "Number kids on joint tour"},
                       {Name: "Joint_" + p + "_HasPreSchKids", Type: "Short", Description: "Does joint tour have pre-school kids?"},
                       {Name: "Joint_" + p + "_HasSeniors", Type: "Short", Description: "Does joint tour have seniors?"},
                       {Name: "Joint_" + p + "_HasWorkers", Type: "Short", Description: "Does joint tour have workers?"},
                       {Name: "Joint_" + p + "_Destination", Type: "Long", Width: 12, Description: "Activity location TAZ choice for Joint " + p},
                       {Name: "Joint_" + p + "_DurChoice", Type: "String", Width: 12, Description: "Activity duration choice for Joint " + p + " in hours"},
                       {Name: "Joint_" + p + "_Duration", Type: "Long", Width: 12, Description: "Activity duration for Joint " + p + " in minutes"},
                       {Name: "Joint_" + p + "_StartInt", Type: "String", Width: 12, Description: "Activity start interval for Joint " + p + ": Format StHr - EndHr"},
                       {Name: "Joint_" + p + "_StartTime", Type: "Long", Width: 12, Description: "Activity start time for Joint " + p + " in minutes from midnight"},
                       {Name: "HomeTo" + "Joint_" + p + "_TT", Type: "Real", Description: "Travel time home to activity for Joint " + p},
                       {Name: "DepTimeTo" + "Joint_" + p, Type: "Long", Width: 12, Description: "Departure time for Joint " + p + " in minutes from midnight"},
                       {Name: "ArrTimeAt" + "Joint_" + p, Type: "Long", Width: 12, Description: "Arrival time at work for Joint " + p + " No." + " in minutes from midnight"},
                       {Name: "Joint_" + p + "_EndTime", Type: "Long", Width: 12, Description: "Activity end time for Joint " + p + " in minutes from midnight"},
                       {Name: "Joint_" + p + "ToHome_TT", Type: "Real", Description: "Travel time back home for Joint " + p + " in minutes from midnight"},
                       {Name: "DepTimeFrom" + "Joint_" + p, Type: "Long", Width: 12, Description: "Departure time from work after Joint " + p + " in minutes from midnight"},
                       {Name: "ArrTimeFrom" + "Joint_" + p, Type: "Long", Width: 12, Description: "Arrival time back home after Joint " + p + " in minutes from midnight"},  
                       {Name: "Joint_" + p + "_Mode", Type: "String", Width: 15, Description: "Activity mode choice for Joint " + p}
                      }
    end
    fldNames = flds.Map(do (f) Return(f.Name) end)
    //abm.DropHHFields(fldNames)
    abm.AddHHFields(flds) 
    
    // Add several person fields
    flds = {{Name: "Probability_Other1_Adult", Type: "Real", Decimals: 2, Description: "Probability of participation in a joint discretionary tour"},
            {Name: "Probability_Other2_Adult", Type: "Real", Decimals: 2, Description: "Probability of participation in a joint discretionary tour"},
            {Name: "Probability_Other_Child", Type: "Real", Decimals: 2, Description: "Probability of participation in a joint discretionary tour"},
            {Name: "Probability_Shop1_Adult", Type: "Real", Decimals: 2, Description: "Probability of participation in a joint discretionary tour"},
            {Name: "Probability_Shop_Child", Type: "Real", Decimals: 2, Description: "Probability of participation in a joint discretionary tour"},
            {Name: "FreeTime_07to20", Type: "Real", Decimals: 2, Description: "Travel time back home for Joint " + p + " in minutes from midnight"},
            {Name: "InJoint_Other1_Tour", Type: "Short", Description: "Is person on the first Other joint discretionary tour?"},
            {Name: "InJoint_Other2_Tour", Type: "Short", Description: "Is person on the second Other joint discretionary tour?"},
            {Name: "InJoint_Shop1_Tour", Type: "Short", Description: "Is person on the first Shop joint discretionary tour?"}}
    //abm.DropPersonFields(fldNames)
    abm.AddPersonFields(flds) 
    
    // Compute Person and HH variables for pattern choice
    TimeManager = RunMacro("Get Time Manager", abm)
    TimeManager.LoadTimeUseMatrix({MatrixFile: Args.MandTimeUseMatrix})
    RunMacro("Get Time Avail for Joint Tours", Args, abm, TimeManager)

    // Compute NonMandatory Accessibilities 
    RunMacro("NonMandatory Joint Accessibility", Args)

    Return(true)
endMacro


//=================================================================================================
// Get Master composition set
// Returns arr as follows:
// arr[1] = {{1}}
// arr[2] = {{1}, {2}, {1,2}}
// arr[3] = {{1}, {2}, {1,2}, {1,3}, {2,3}, {1,2,3}, {3}}
// ...
// arr[9] = {...} // 511 elements
Macro "Get Master Composition Set"(n)
    dim arr[n]
    arr[1] = {{1}}
    for i = 2 to n do
        newArr = arr[i-1].Map(do (f) Return(f + {i}) end) // Add latest element to each sub-array in the set
        arr[i] = CopyArray(arr[i-1]) + newArr + {{i}}     // Append new array to previous array and add {i} element to the array
    end
    Return(arr)
endMacro


//=================================================================================================
Macro "JointTours Frequency"(Args)
    abm = RunMacro("Get ABM Manager", Args)
    objDC = CreateObject("Table", Args.NonMandatoryDestAccessibility)

    // Run Model for all HH whose SubPattern contains J and populate output fields
    obj = CreateObject("PMEChoiceModel", {ModelName: "Joint Tours Frequency"})
    obj.OutputModelFile = Args.[Output Folder] + "\\Intermediate\\JointToursFrequency.mdl"
    obj.AddTableSource({SourceName: "HH", View: abm.HHView, IDField: abm.HHID})
    obj.AddTableSource({SourceName: "DCLogsums", View: objDC.GetView(), IDField: "TAZID"})
    obj.AddPrimarySpec({Name: "HH", Filter: "Lower(SubPattern) contains 'j'", OField: "TAZID"})
    obj.AddUtility({UtilityFunction: Args.JointTourFrequencyUtility})
    obj.AddOutputSpec({ChoicesField: "JointTourPattern"})
    obj.ReportShares = 1
    obj.RandomSeed = 4799999
    ret = obj.Evaluate()
    if !ret then
        Throw("Model Run failed for Joint Tours Frequency")
    Args.[JointTours Frequency Spec] = CopyArray(ret)
    obj = null

    // Writing the choice from JointTourPattern into number of other and shop tours
    otherTour_map = { 'O1': 1, 'O2': 2, 'S1': 0, 'S2' : 0, 'O1S1': 1, 'O2S1': 2, 'O1S2': 1}
    shopTour_map =  { 'O1': 0, 'O2': 0, 'S1': 1, 'S2' : 2, 'O1S1': 1, 'O2S1': 1, 'O1S2': 2}
    vJT = abm.[HH.JointTourPattern]
    arrO = v2a(vJT).Map(do (f) if f = null then Return(0) else Return(otherTour_map.(f)) end)
    arrS = v2a(vJT).Map(do (f) if f = null then Return(0) else Return(shopTour_map.(f)) end)
    vecsSet = null
    vecsSet.NumberJointOtherTours =  a2v(arrO)
    vecsSet.NumberJointShopTours =  a2v(arrS)
    abm.SetHHVectors(vecsSet)    
    Return(true)
endMacro


//=================================================================================================
Macro "JointTours Destination Other"(Args)
    // Loop over each Other tour
    abm = RunMacro("Get ABM Manager", Args)
    for tourno in {1, 2} do
        // Find Joint Tour Destination Other Location
        obj = CreateObject("PMEChoiceModel", {SourcesObject: Args.SourcesObject, ModelName: "Joint Tours Destination Other"})
        obj.OutputModelFile = Args.[Output Folder] + "\\Intermediate\\JointToursDestinationOther.dcm"
        obj.AddTableSource({SourceName: "HH", View: abm.HHView, IDField: abm.HHID})
        obj.AddTableSource({SourceName: "TAZData", File: Args.DemographicOutputs, IDField: "TAZ"})
        obj.AddTableSource({SourceName: "TAZ4Ds", File: Args.AccessibilitiesOutputs, IDField: "TAZID"})
        obj.AddMatrixSource({SourceName: "AutoSkim", File: Args.HighwaySkimOP, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
        obj.AddMatrixSource({SourceName: "WalkSkim", File: Args.WalkSkim, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
        obj.AddMatrixSource({SourceName: "Intrazonal", File: Args.IZMatrix, RowIndex: "TAZ", ColIndex: "TAZ"})
        obj.AddMatrixSource({SourceName: "ModeAccessibility", File: Args.NonMandJointModeAccessOther, RowIndex: "TAZ", ColIndex: "TAZ"})
        obj.AddPrimarySpec({Name: "HH", Filter: "NumberJointOtherTours >= " + i2s(tourno), OField: "TAZID"})
        obj.AddUtility({UtilityFunction: Args.JointTourDestOtherUtility})
        obj.AddDestinations({DestinationsSource: "AutoSkim", DestinationsIndex: "InternalTAZ"})
        obj.AddSizeVariable({Name: "TAZData", Field: "JointOtherSize"})
        obj.AddOutputSpec({ChoicesField: "Joint_Other" + i2s(tourno) + "_Destination"})
        obj.RandomSeed = 4899993 + 4*tourno
        ret = obj.Evaluate()
        if !ret then
            Throw("Model Run failed for Joint Tours Destination Other")
    end
    Return(true)
endMacro



//==========================================================================================================
Macro "JointTours Destination Shop"(Args)
    // Loop over each Shop tour
    abm = RunMacro("Get ABM Manager", Args)
    for tourno in {1} do
        // Find Joint Tour Destination Shop Location
        obj = CreateObject("PMEChoiceModel", {SourcesObject: Args.SourcesObject, ModelName: "Joint Tours Destination Shop"})
        obj.OutputModelFile = Args.[Output Folder] + "\\Intermediate\\JointToursDestinationShop.dcm"
        obj.AddTableSource({SourceName: "HH", View: abm.HHView, IDField: abm.HHID})
        obj.AddTableSource({SourceName: "TAZData", File: Args.DemographicOutputs, IDField: "TAZ"})
        obj.AddTableSource({SourceName: "TAZ4Ds", File: Args.AccessibilitiesOutputs, IDField: "TAZID"})
        obj.AddMatrixSource({SourceName: "AutoSkim", File: Args.HighwaySkimOP, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
        obj.AddMatrixSource({SourceName: "Intrazonal", File: Args.IZMatrix, RowIndex: "TAZ", ColIndex: "TAZ"})
        obj.AddPrimarySpec({Name: "HH", Filter: "NumberJointShopTours >= " + i2s(tourno), OField: "TAZID"})
        obj.AddUtility({UtilityFunction: Args.JointTourDestShopUtility})
        obj.AddDestinations({DestinationsSource: "AutoSkim", DestinationsIndex: "InternalTAZ"})
        obj.AddSizeVariable({Name: "TAZData", Field: "JointShopSize"})
        obj.AddOutputSpec({ChoicesField: "Joint_Shop" + i2s(tourno) + "_Destination"})
        obj.RandomSeed = 4999999
        ret = obj.Evaluate()
        if !ret then
            Throw("Model Run failed for Joint Tours Destination Shop")
    end

    Return(true)
endMacro


//==========================================================================================================
Macro "JointTours Scheduling"(Args)
    // Initialize Time Manager
    abm = RunMacro("Get ABM Manager", Args)
    TimeManager = RunMacro("Get Time Manager", abm)
    TimeManager.LoadTimeUseMatrix({MatrixFile: Args.MandTimeUseMatrix}) // loading the existing file

    // Get master set of person combinations
    maxHHSize = 9
    Args.MasterCompositionSet = RunMacro("Get Master Composition Set", maxHHSize)

    purps = {"Other1", "Shop1", "Other2"}
    pbar = CreateObject("G30 Progress Bar", "Joint Tour Scheduling for Other1, Shop1 and Other2 Tours", true, 3)
    for p in purps do
        pbar1 = CreateObject("G30 Progress Bar", "Joint Tour Scheduling (Composition, Participation, Duration, StartTime, Mode and TimeManagerUpdate) for " + p + " Tours", true, 6)

        // Composition Model
        spec = {ModelType: p, abmManager: abm, TimeManager: TimeManager}
        RunMacro("JointTours Composition", Args, spec)
        if pbar1.Step() then
            Return()

        // Participation Model
        RunMacro("JointTours Participation", Args, spec)
        if pbar1.Step() then
            Return()          

        // Duration Model
        RunMacro("JointTours Duration", Args, spec)
        if pbar1.Step() then
            Return()

        // Start Time Model
        RunMacro("JointTours StartTime", Args, spec)
        if pbar1.Step() then
            Return()

        // Mode Choice Model
        RunMacro("JointTours Mode", Args, spec)
        if pbar1.Step() then
            Return()

        // Update Time Manager
        RunMacro("JT Update TimeManager", Args, spec)
        if pbar1.Step() then
            Return()
        pbar1.Destroy()
        
        if pbar.Step() then
            Return()
    end
    pbar.Destroy()

    // Write out time manager matrix
    TimeManager.ExportTimeUseMatrix(Args.JointTimeUseMatrix)
    ret_value = 1
  
    Return(true)
endMacro


//==================================================================================================
Macro "JointTours Composition"(Args, spec)
    // Run Choice Model
    RunMacro("Composition Choice Model", Args, spec)
    
    // Post Process
    RunMacro("Fill Composition Fields", Args, spec)
endMacro


//==================================================================================================
Macro "Composition Choice Model"(Args, spec)
    p = spec.ModelType
    abm = spec.abmManager
    purp = Left(p, Stringlength(p) - 1)
    TourNo = s2i(Right(p, 1))
    id = Stringlength(p) - 5 + TourNo // Yields 1, 2 or 3 for Shop1, Other1 and Other2 respectively
    
    // Compute availabilities based on alternative names and HH member composition
    {availSpec, fieldList} = RunMacro("Compute Composition Availabilities", Args, spec)

    // Run Choice Model
    outputField = "Joint_" + p + "_Composition"
    utilFn = Args.("JointTourComposition" + purp + "Utility")
    filter = "NumberJoint" + purp + "Tours >= " + i2s(TourNo)

    obj = CreateObject("PMEChoiceModel", {ModelName: "Joint Tour Composition " + p})
    obj.OutputModelFile = Args.[Output Folder] + "\\Intermediate\\JointTourComposition" + p + ".mdl"
    obj.AddTableSource({SourceName: "HH", View: abm.HHView, IDField: abm.HHID})
    obj.AddPrimarySpec({Name: "HH", Filter: filter})
    obj.AddUtility({UtilityFunction: utilFn, AvailabilityExpressions: availSpec}) 
    obj.AddOutputSpec({ChoicesField: outputField})
    obj.ReportShares = 1
    obj.RandomSeed = 5100067 + id*30
    ret = obj.Evaluate()
    if !ret then
        Throw("Model Run failed for Joint Tour Composition " + p)
    Args.("JointTours " + purp + " Composition Spec") = CopyArray(ret)
    obj = null
endMacro


//==================================================================================================
Macro "Fill Composition Fields"(Args, spec)
    p = spec.ModelType
    abm = spec.abmManager
    outputField = "Joint_" + p + "_Composition"
    participationField = "InJoint_" + p + "_Tour"
    
    // Fill in number of adults and number of kids on tour for HHs that did not choose 'All'
    filter = printf("%s <> 'All' and %s <> null", {outputField, outputField})
    abm.CreateHHSet({Filter: filter, Activate: 1})
    
    fld = 'HH.' + outputField
    v = abm.(fld)
    vAdults = v2a(v).Map(do (f) Return(s2i(SubString(f, 2, 1))) end)
    vKids = v2a(v).Map(do (f) Return(s2i(SubString(f, 4, 1))) end)
        
    vecsSet = null
    vecsSet.("Joint_" + p + "_nAdults") = a2v(vAdults)
    vecsSet.("Joint_" + p + "_nKids") = a2v(vKids)
    abm.SetHHVectors(vecsSet)

    // Fill in number of adults and number of kids on tour for HHs that chose 'All'
    // Restrict tour party for large HHs. Max of 6 people on any tour with at least 1 adult and at most 4 kids
    filter = printf("%s = 'All'", {outputField})
    abm.CreateHHSet({Filter: filter, Activate: 1})
    vecs = abm.GetHHVectors({"Adults", "Kids"})
    vA = if vecs.Adults > 6 then 6 else vecs.Adults
    vK = if vecs.Kids > 4 then 4 else vecs.Kids
    vExcess = vA + vK - 6
    vA = if vExcess > 0 then vA - vExcess else vA

    vecsSet = null
    vecsSet.("Joint_" + p + "_nAdults") = vA
    vecsSet.("Joint_" + p + "_nKids") = vK
    abm.SetHHVectors(vecsSet)
    
    // For cases where all members are on the tour, fill joint info fields
    types = {"Adults", "Kids"}
    baseQrys = {"Age >= 18", "Age < 18"}
    for i = 1 to types.length do
        t = types[i]
        tourFld = "Joint_" + p + "_n" + t // e.g. Joint_Other1_nAdults
        filter = printf("%s = %s and %s", {tourFld, t, baseQrys[i]})
        setInfo = abm.CreatePersonSet({Filter: filter, UsePersonHHView: 1, Activate:1})
        if setInfo.Size > 0 then do
            v = Vector(setInfo.Size, "Short", {Constant: 1})
            outFld = 'Person.' + participationField
            abm.(outFld) = v
        end
    end
endMacro


//=================================================================================================
Macro "JointTours Participation"(Args, spec)
    p = spec.ModelType
    abm = spec.abmManager
    purpose = SubString(p, 1, StringLength(p) - 1) //"Other" or "Shop"
    tourNo = s2i(Right(p, 1))
    kidsFreqFld = "Joint_" + p + "_nKids"
    adultsFreqFld = "Joint_" + p + "_nAdults"
    id = Stringlength(p) - 5 + tourNo // Yields 1, 2 or 3 for Shop1, Other1 and Other2 respectively
    pbar = CreateObject("G30 Progress Bar", "Joint Tour Participation, for Adults and Kids", true, 2)

    // Adults ===============
    // Probability model
    filter = "Age >= 18 and " + adultsFreqFld + " < Adults"
    utilFn = Args.("Participation" + purpose + "AdultsUtility")
    probFld = "Probability_" + p + "_Adult"
    Opts = {Purpose: purpose, PersonType: 'Adult', Filter: filter, UtilityFunction: utilFn, ProbabilityField: probFld, TourNo: tourNo, abmManager: spec.abmManager}
    RunMacro("Participation Probability Model", Args, Opts)

    // Pick Participants
    flagFld = "InJoint_" + p + "_Tour"
    seed = 5200421 + (id-1)*id
    SetRandomSeed(seed)
    Opts = {FrequencyField: adultsFreqFld, Filter: filter, ProbabilityField: probFld, OutputField: flagFld, abmManager: spec.abmManager}
    RunMacro("Participation Decision Model", Args, Opts)

    if pbar.Step() then
        Return()

    // Kids ==================
    // Probability model
    filter = "Age < 18 and " + kidsFreqFld + " > 0 and " + kidsFreqFld + " < Kids"
    utilFn = Args.("Participation" + purpose + "KidsUtility")
    probFld = "Probability_" + purpose + "_Child"
    Opts = {Purpose: purpose, PersonType: 'Child', Filter: filter, UtilityFunction: utilFn, ProbabilityField: probFld, abmManager: spec.abmManager}
    RunMacro("Participation Probability Model", Args, Opts)

    // Pick participants on the tour
    seed = 5302121 + 6*id
    SetRandomSeed(seed)
    Opts = {FrequencyField: kidsFreqFld, Filter: filter, ProbabilityField: probFld, OutputField: flagFld, abmManager: spec.abmManager}
    RunMacro("Participation Decision Model", Args, Opts)

    // Postprocess
    RunMacro("Participation PostProcess", Args, spec)

    pbar.Destroy()
endMacro


//=================================================================================================
Macro "Participation Probability Model"(Args, Opts)
    purp = Opts.Purpose
    pType = Opts.PersonType
    abm = Opts.abmManager
    probFile = Args.[Output Folder] + "\\Intermediate\\ParticipationProb_" + pType + "_" + purp + ".bin"
    
    utilOpts = null
    utilOpts.UtilityFunction = Opts.UtilityFunction
    if Opts.TourNo > 0 then
        utilOpts.SubstituteStrings = {{"<TourNo>", String(Opts.TourNo)}}

    // Run Model
    obj = CreateObject("PMEChoiceModel", {ModelName: "Joint Tour Participation - " + purp + " - " + pType})
    obj.AddTableSource({SourceName: "PersonHH", View: abm.PersonHHView, IDField: abm.PersonID})
    obj.OutputModelFile = Args.[Output Folder] + "\\Intermediate\\Participation_" + pType + "_" + purp + ".mdl"
    obj.AddPrimarySpec({Name: "PersonHH", Filter: Opts.Filter})
    obj.AddUtility(utilOpts) 
    obj.AddOutputSpec({ProbabilityTable: probFile})
    ret = obj.Evaluate()
    if !ret then
        Throw("Model Run failed for Joint Tour Participation - " + pType + " " + purp)
    Args.("JT " + pType + " " + purp + " Participation Spec") = CopyArray(ret)
    
    // Join probability table to above PersonHH and copy values
    objP = CreateObject("Table", probFile)
    vwProb = objP.GetView()
    {flds, specs} = GetFields(vwProb,)
    vwJ1 = JoinViews("ProbPersonHH", specs[1], GetFieldFullSpec(abm.PersonHHView, abm.PersonID),)
    v = GetDataVector(vwJ1 + "|", "[Yes Probability]",)
    SetDataVector(vwJ1 + "|", Opts.ProbabilityField, v,)
    CloseView(vwJ1)
    objP = null
endMacro


//=================================================================================================
Macro "Participation Decision Model"(Args, Opts)
    outField = Opts.OutputField
    abm = Opts.abmManager
    abm.CreatePersonSet({Filter: Opts.Filter, Activate: 1, UsePersonHHView: 1})

    mr = CreateObject("Model.Runtime")
    codeUI = mr.GetModelCodeUI()
    iterOpts = {UIName: codeUI,
                MacroName: "Pick Tour Participants",
                MacroArgs: {MasterSet: Args.MasterCompositionSet, ProbabilityField: Opts.ProbabilityField, FrequencyField: Opts.FrequencyField, OutputField: outField},
                InputFields: {abm.PersonID, Opts.ProbabilityField, Opts.FrequencyField},
                OutputFields: {outField}
                }
    abm.Iterate(iterOpts)
endMacro


//=================================================================================================
Macro "Pick Tour Participants"(opt)
    vecs = opt.InputVecs
    vecsOut = opt.OutputVecs
    startIdx = opt.StartIndex
    endIdx = opt.EndIndex
    nTotal = endIdx - startIdx + 1
    
    MacroArgs = opt.MacroArgs
    masterSet = MacroArgs.MasterSet
    probFld = MacroArgs.ProbabilityField
    freqFld = MacroArgs.FrequencyField
    outFld = MacroArgs.OutputField
    
    dim probArr[nTotal]
    v = vecs.(probFld)
    for i = 1 to nTotal do
        probArr[i] = v[startIdx - 1 + i]
    end
    
    // Given array of probabilities and how many to choose, draw choice assuming binomial poisson distribution
    vF = vecs.(freqFld)
    nPick = vF[startIdx]    // Number of records to pick out of nTotal
    choice = RunMacro("Draw Binomial Poisson", probArr, nPick, masterSet[nTotal])

    // Fill appropriate output vector based on choice index
    vOut = vecsOut.(outFld)
    for ch in choice do
        vOut[startIdx + ch - 1] = 1
    end
endMacro


//=================================================================================================
Macro "Draw Binomial Poisson"(probArr, nReq, combinationSet)
    choices = null
    weights = null
    
    // Generate valid sets and their corresponding probabilities
    for set in combinationSet do
        if set.length = nReq then do    // Set valid if its cardinal number equals the number of participants on the tour
            choices = choices + {set}
            prob = 1
            prevElem = 0
            
            for elem in set do
                for j = prevElem + 1 to elem - 1 do
                    prob = prob*(1 - probArr[j])
                end
                prob = prob*probArr[elem]
                prevElem = elem
            end
            
            for j = prevElem + 1 to probArr.length do
                prob = prob*(1 - probArr[j])    
            end

            if prob < 1e-6 then
                prob = 1e-6

            weights = weights + {prob*1000000} // To avoid small numbers that seem to trip up RandSamples() below
        end
    end

    // Choose
    samples = RandSamples(1, "Discrete", {weight: weights}) 

    // Return chosen set
    Return(choices[samples[1]])
endMacro


//=================================================================================================
Macro "Participation PostProcess"(Args, spec)
    p = spec.ModelType
    abm = spec.abmManager
    vwP = abm.PersonView
    fld = "InJoint_" + p + "_Tour"

    // Specify array
    optA = null
    optA.("Joint_" + p + "_HasPreSchKids") = printf("(Age <= 5).%s.Max", {fld})
    optA.DefaultValue = 0
    aggSpec = {optA}

    optA = null
    optA.("Joint_" + p + "_HasWorkers") = printf("(WorkerCategory <= 2).%s.Max", {fld})
    optA.DefaultValue = 0
    aggSpec = aggSpec + {optA}

    optA = null
    optA.("Joint_" + p + "_HasSeniors") = printf("(Age >= 65).%s.Max", {fld})
    optA.DefaultValue = 0
    aggSpec = aggSpec + {optA}

    // Fill field in HH layer with one of the person IDs 
    optA = null
    optA.("Joint_" + p + "_DesignatedPerson") = printf("(%s = 1).%s.Min", {fld, abm.PersonID})
    optA.DefaultValue = null
    aggSpec = aggSpec + {optA}

    abm.AggregatePersonData({Spec: aggSpec})
endMacro


//==========================================================================================================
Macro "JointTours Duration"(Args, spec)
    p = spec.ModelType
    abm = spec.abmManager
    TimeManager = spec.TimeManager

    purpose = SubString(p, 1, StringLength(p) - 1) //"Other" or "Shop"
    TourNo = s2i(Right(p, 1))
    id = Stringlength(p) - 5 + TourNo // Yields 1, 2 or 3 for Shop1, Other1 and Other2 respectively

    // Get maximum contiguous free time for persons making the joint tour
    HHFilter = "Joint_" + p + "_Composition <> null"
    PersonFilter = "InJoint_" + p + "_Tour = 1"

    HHSpec = {ViewName: abm.HHView, HHID: abm.HHID, Filter: HHFilter}
    PersonSpec = {ViewName: abm.PersonView, PersonID: abm.PersonID, Filter: PersonFilter}   
    opts = {PersonSpec: PersonSpec, HHSpec: HHSpec, HHFillField: "MaxFreeTime", Metric: "MaxAvailTime", StartTime: 360, EndTime: 1380} // 0700 to 2300
    TimeManager.FillHHTimeField(opts)

    // Get Duration Availabilities
    utilFunction = Args.("JointTourDur" + purpose + "Utility")
    FldSpec = 'HH.MaxFreeTime'
    durAvailArray = RunMacro("Get Duration Avail", utilFunction, FldSpec)

    // Run Duration Model
    modelName = "Joint Tours Duration " + purpose
    Opts = {abmManager: abm,
            ModelName: modelName,
            ModelFile: "JointToursDuration" + purpose + ".mdl",
            PrimarySpec: {Name: 'HH', View: abm.HHView, ID: abm.HHID},
            Filter: HHFilter,
            DestField: "Joint_" + p + "_Destination", 
            Utility: utilFunction,
            Availabilities: durAvailArray,
            ChoiceField: "Joint_" + p + "_DurChoice",
            SimulatedTimeField: "Joint_" + p + "_Duration",
            AlternativeIntervalInMin: 1,
            SubstituteStrings: {{"<purp>", p}},
            MinimumDuration: 10,
            RandomSeed: 5400821 + 6*id
            }
    RunMacro("NonMandatory Activity Time", Args, Opts)
endMacro


//==========================================================================================================
Macro "JointTours StartTime"(Args, spec)
    p = spec.ModelType
    abm = spec.abmManager
    TimeManager = spec.TimeManager

    purpose = SubString(p, 1, StringLength(p) - 1) // "Other" or "Shop"
    TourNo = s2i(Right(p, 1))
    id = Stringlength(p) - 5 + TourNo // Yields 1, 2 or 3 for Shop1, Other1 and Other2 respectively

    // Get start time availabilities based on duration and availability of persons on the tour
    jointAltTable = Args.("JointTourStart" + purpose + "Alts")
    startTimeAlts = jointAltTable.Alternative
    
    // Get start time availability table
    HHFilter = "Joint_" + p + "_Composition <> null"
    PersonFilter = "InJoint_" + p + "_Tour = 1"
    HHSpec = {ViewName: abm.HHView, HHID: abm.HHID, Filter: HHFilter}
    PersonSpec = {ViewName: abm.PersonView, PersonID: abm.PersonID, Filter: PersonFilter}  
    
    opts = {HHSpec: HHSpec, 
            PersonSpec: PersonSpec, 
            DurationField: "Joint_" + p + "_Duration", 
            StartTimeAlts: startTimeAlts, 
            OutputAvailFile: Args.JointStartAvails}
    TimeManager.GetJointStartTimeAvailabilities(opts)

    objA = CreateObject("Table", Args.JointStartAvails)
    vwJ = JoinViews("JointToursHHData", GetFieldFullSpec(abm.HHView, abm.HHID), GetFieldFullSpec(objA.GetView(), "RecordID"),)

    // Create start time availability expressions
    stAvailArray = RunMacro("Get StartTime Avail", jointAltTable, "JointToursHHData")

    // Run Start Time Model
    modelName = "Joint Tours StartTime " + purpose
    JtOpts =   {abmManager: abm,
                ModelName: modelName,
                ModelFile: "JointToursStart" + purpose + ".mdl",
                PrimarySpec: {Name: 'JointToursHHData', View: vwJ, ID: abm.HHID},
                Filter: HHSpec.Filter,
                DestField: "Joint_" + p + "_Destination",
                Alternatives: jointAltTable,
                Utility: Args.("JointTourStart" + purpose + "Utility"),
                Availabilities: stAvailArray,
                ChoiceField: "Joint_" + p + "_StartInt",
                SimulatedTimeField: "Joint_" + p + "_StartTime",
                SubstituteStrings: {{"<purp>", p}},
                MinimumDuration: 10,
                RandomSeed: 5502175 + 6*id}
    RunMacro("NonMandatory Activity Time", Args, JtOpts)
    if !spec.LeaveDataOpen then do
        CloseView(vwJ)
        objA = null
    end

    // Fill Activity end time (used in setting avails)
    setInfo = abm.CreateHHSet({Filter: HHFilter, Activate: 1})
    if setInfo.Size = 0 then
        Throw("No records to run 'Activity Start Model' for Joint " + p)

    durFld = "Joint_" + p + "_Duration"
    actStartTimeFld = JtOpts.SimulatedTimeField
    actEndTimeFld = "Joint_" + p + "_EndTime"
    vecs = abm.GetHHVectors({durFld, actStartTimeFld})
    vecsSet = null
    vecsSet.(actEndTimeFld) = vecs.(durFld) + vecs.(actStartTimeFld)
    abm.SetHHVectors(vecsSet)
endMacro


//-================================================================================================
Macro "JointTours Mode"(Args, spec)
    p = spec.ModelType
    abm = spec.abmManager
    purpose = SubString(p, 1, StringLength(p) - 1) // "Other" or "Shop"
    baseFilter = "Joint_" + p + "_Composition <> null and Joint_" + p + "_Destination <> null"
    TourNo = s2i(Right(p, 1))
    id = Stringlength(p) - 5 + TourNo // Yields 1, 2 or 3 for Shop1, Other1 and Other2 respectively
    
    // Compute time period field
    periodDefs = Args.TimePeriods
    amStart = periodDefs.AM.StartTime
    amEnd = periodDefs.AM.EndTime
    pmStart = periodDefs.PM.StartTime
    pmEnd = periodDefs.PM.EndTime
    depTime = "Joint_" + p + "_StartTime"
    amQry = printf("(%s >= %s and %s < %s)", {depTime, String(amStart), depTime, String(amEnd)})
    pmQry = printf("(%s >= %s and %s < %s)", {depTime, String(pmStart), depTime, String(pmEnd)})
    exprStr = printf("if %s then 'AM' else if %s then 'PM' else 'OP'", {amQry, pmQry})
    depPeriod = CreateExpression(abm.HHView, "DepPeriod", exprStr,)

    timePeriods = {"AM", "PM", "OP"}
    for tper in timePeriods do
        periodFilter = printf("(DepPeriod = '%s')", {tper})
        filterStr = baseFilter + " and " + periodFilter
        MCOpts = {abmManager: abm, TourTag: p, Purpose: purpose, TimePeriod: tper, Filter: filterStr,
                    RandomSeed: 5601355 + 6*id + timePeriods.position(tper)}
        RunMacro("JointTours Mode Eval", Args, MCOpts)
    end
    
    // Fill departure from home/arrival back home fields after mode choice
    RunMacro("JointTours Mode PostProcess", Args, spec)

    DestroyExpression(GetFieldFullSpec(abm.HHView, "DepPeriod"))
endMacro


//-================================================================================================
Macro "JointTours Mode Eval"(Args, MCOpts)
    p = MCOpts.TourTag
    purpose = MCOpts.Purpose
    tod = MCOpts.TimePeriod
    abm = MCOpts.abmManager
    modelName = "Joint_" + purpose + "_" + tod + "_Mode"
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
    obj.OutputModelFile = Args.[Output Folder] + "\\Intermediate\\JointToursMode" + purpose + tod + ".mdl"
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
    obj.AddTableSource({SourceName: "HH", View: abm.HHView, IDField: abm.HHID})
    obj.AddPrimarySpec({Name: "HH", Filter: MCOpts.Filter, OField: "TAZID", DField: "Joint_" + p + "_Destination"})
    
    // Filter out modes that aren't present
    ret = RunMacro("Filter Mode Utility Spec", {
        util: Args.("JointTourMode" + purpose + "Utility"),
        avail: Args.("JointTourMode" + purpose + "Avail"),
        Args: Args
    })
    util = ret.Utility
    avail = ret.Availability

    utilOpts = null
    utilOpts.UtilityFunction = util
    utilOpts.SubstituteStrings = {{"<purp>", p}}
    utilOpts.AvailabilityExpressions = avail
    obj.AddUtility(utilOpts)
    
    obj.AddOutputSpec({ChoicesField: "Joint_" + p + "_Mode"})
    obj.ReportShares = Args.ReportShares
    obj.RandomSeed = MCOpts.RandomSeed
    ret = obj.Evaluate()
    if !ret then
        Throw("Model Run failed for Joint Tours Mode " + p)
    Args.(modelName + " Spec") = CopyArray(ret)
endMacro


//==========================================================================================================
// The macro calculates travel times to and from the home to the activity.
// Finally it establishes departure time from home and arrival time back home to inform future model decisions.
Macro "JointTours Mode PostProcess"(Args, spec)
    p = spec.ModelType
    abm = spec.abmManager
    purpose = SubString(p, 1, StringLength(p) - 1) // "Other" or "Shop"

    filter = "Joint_" + p + "_Composition <> null"
    setInfo = abm.CreateHHSet({Filter: filter, Activate: 1})
    if setInfo.Size = 0 then
        Return()

    durFld = "Joint_" + p + "_Duration"
    destFld = "Joint_" + p + "_Destination"
    actStartTimeFld = "Joint_" + p + "_StartTime"
    actEndTimeFld = "Joint_" + p + "_EndTime"
    homeToActTTFld = "HomeTo" + "Joint_" + p + "_TT"
    actToHomeTTFld = "Joint_" + p + "ToHome_TT"
    modeFld = "Joint_" + p + "_Mode"

    fillSpec = {View: abm.HHView, OField: "TAZID", DField: destFld, FillField: homeToActTTFld, 
                Filter: filter, ModeField: modeFld, DepTimeField: actStartTimeFld}
    RunMacro("Fill Travel Times", Args, fillSpec)

    fillSpec = {View: abm.HHView, OField: destFld, DField: "TAZID", FillField: actToHomeTTFld, 
                Filter: filter, ModeField: modeFld, DepTimeField: actEndTimeFld}
    RunMacro("Fill Travel Times", Args, fillSpec)

    // Fill departure time from home and arrival time after activity
    arrTimeBackHomeFld = "ArrTimeFrom" + "Joint_" + p
    depTimeFromHomeFld = "DepTimeTo" + "Joint_" + p
    arrTimeAtDestFld = "ArrTimeAt" + "Joint_" + p
    depTimeFromDestFld = "DepTimeFrom" + "Joint_" + p
    
    vecs = abm.GetHHVectors({actStartTimeFld, actEndTimeFld, homeToActTTFld, actToHomeTTFld})
    vecsSet = null
    vecsSet.(arrTimeBackHomeFld) = vecs.(actEndTimeFld) + Round(vecs.(actToHomeTTFld),0)
    vecsSet.(depTimeFromHomeFld) = vecs.(actStartTimeFld) - Round(vecs.(homeToActTTFld), 0)
    vecsSet.(arrTimeAtDestFld) = vecs.(actStartTimeFld)
    vecsSet.(depTimeFromDestFld) = vecs.(actEndTimeFld)
    abm.SetHHVectors(vecsSet)
endMacro


//==========================================================================================================
Macro "JT Update TimeManager"(Args, spec)
    p = spec.ModelType
    abm = spec.abmManager
    TimeManager = spec.TimeManager
    
    fldsToExport = {abm.PersonID, "InJoint_" + p + "_Tour", "DepTimeTo" + "Joint_" + p, "ArrTimeFrom" + "Joint_" + p}
    vwTemp = ExportView(abm.PersonHHView + "|", "MEM", "TempPersonHH", fldsToExport,)
    tourOpts = {ViewName: vwTemp,
                Filter: "InJoint_" + p + "_Tour = 1",
                PersonID: abm.PersonID,
                Departure: "DepTimeTo" + "Joint_" + p,
                Arrival: "ArrTimeFrom" + "Joint_" + p}
    TimeManager.UpdateMatrixFromTours(tourOpts)

    // Refill time fields
    RunMacro("Get Time Avail for Joint Tours", Args, abm, TimeManager)

    CloseView(vwTemp)
endMacro


//=================================================================================================
// Calculate time available by person needed by discretionary joint tour models 
Macro "Get Time Avail for Joint Tours"(Args, abm, TimeManager)
    hhSpec = {ViewName: abm.HHView, HHID: abm.HHID}
    opts = {HHSpec: hhSpec, Metric: "AverageFreeTime", HHFillField: "HHAvgFreeTime_09to18", StartTime: 540, EndTime: 1080}
    TimeManager.FillHHTimeField(opts)

    perSpec = {ViewName: abm.PersonView, PersonID: abm.PersonID}
    opts = {PersonSpec: perSpec, Metric: "FreeTime", PersonFillField: "FreeTime_07to20", StartTime: 420, EndTime: 1200}
    TimeManager.FillPersonTimeField(opts)
endMacro


//============================================================================================================
// Macro that parses the non-mandatory tour composition alternative names
// and sets full-household alternatives to 0, since they are captured under the All alternative.
// The macro returns two arrays, one with the alternative names and one with their corresponding availability field names.
Macro "Compute Composition Availabilities" (Args, spec)
    specAvail = null
    alts = null
    exprs = null

    p = spec.ModelType
    if Lower(p) contains "other" then do
        utils = Args.JointTourCompositionOtherUtility
        qry = "NumberJointOtherTours > 0"
    end
    else if Lower(p) contains "shop" then do
        utils = Args.JointTourCompositionShopUtility
        qry = "NumberJointShopTours > 0"
    end
    else
        Throw("Invalid non-mandatory tour purpose.")

    // Generate the list of availability fields to add
    for i = 1 to utils.length do
        if utils[i][1] <> "Description" and utils[i][1] <> "Expression" and utils[i][1] <> "All" then do
            alts = alts + {utils[i][1]}
            exprs = exprs + {"HH" + ".Avail_" + utils[i][1]}
        end
    end

    specAvail.Alternative = alts
    specAvail.Expression = exprs

    // Delete the fields from the primary source table, in case they already exist.
    // This list is also passed back to the calling macro to aid in the deletion of the fields after the step runs.
    flds = alts.Map(do (f) Return("Avail_" + f) end)
    abm = spec.abmManager
    abm.DropHHFields(flds)
    fldsToAdd = flds.Map(do (f) Return ({Name: f, Type: "Short"}) end)
    abm.AddHHFields(fldsToAdd)
    
    // Populate the fields with 1's and 0's
    abm.CreateHHSet({Filter: qry, Activate: 1})
    vecs = abm.GetHHVectors({"Adults", "Kids"})

    vecsSet = null
    for alt in alts do
        altAdults = s2i(Substring(alt, 2, 1))   // Number of adults in the alternative
        altKids = s2i(Substring(alt, 4, 1))     // Minimum number of kids in the alternative

        v_Avail = vecs.Adults >= altAdults and vecs.Kids >= altKids and not(vecs.Adults = altAdults and vecs.Kids = altKids)
        vecsSet.("Avail_" + alt) = v_Avail
    end
    abm.SetHHVectors(vecsSet)
    abm.ActivateHHSet()

    // Return the availability specAvail to be used while writing the model file
    Return({specAvail, flds})
endMacro


//=========================================================================================================
// Determines availability expressions for each of the duration alternatives (solo and joint tours)
Macro "Get Duration Avail"(utilSpec, fieldSpec)
    availArr = null
    alts = null
    generalCols = {"expression", "description", "segment", "filter", "coefficient"}
    for elem in utilSpec do
        colName = elem[1]
        if generalCols.position(colName) = 0 then 
            alts = alts + {colName}
    end

    availArr.Alternative = CopyArray(alts)
    for alt in alts do
        tmpArr = ParseString(alt, "- ")
        altDur = s2i(tmpArr[1]) // Minutes
        expr = "if " + fieldSpec + " * 60 >= " + String(altDur) + " then 1 else 0" // free time in hrs
        availArr.Expression =  availArr.Expression + {expr}
    end

    Return(availArr)
endMacro


// Macro that runs the activity time choice model to generate activity duration.
// Called for joint and solo discretionary tours
Macro "NonMandatory Activity Time"(Args, Opts)
    abm = Opts.abmManager
    filter = Opts.Filter
    primarySpec = Opts.PrimarySpec

    // Basic Check
    if Opts.Utility = null or Opts.ModelName = null or Opts.ModelFile = null or filter = null
        or Opts.DestField or Opts.ChoiceField = null or Opts.PrimarySpec = null then
            Throw("Invalid inputs to macro 'NonMandatory Activity Time'")

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
    obj.AddTableSource({SourceName: primarySpec.Name, View: primarySpec.View, IDField: primarySpec.ID})
    obj.AddMatrixSource({SourceName: "AutoSkim", File: Args.HighwaySkimOP, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
    obj.AddPrimarySpec({Name: primarySpec.Name, Filter: filter, OField: "TAZID", DField: Opts.DestField})
    if Opts.Alternatives <> null then
        obj.AddAlternatives({AlternativesList: Opts.Alternatives})
    obj.AddUtility(utilOpts)
    obj.AddOutputSpec({ChoicesField: Opts.ChoiceField})
    obj.ReportShares = Args.ReportShares
    obj.RandomSeed = Opts.RandomSeed
    ret = obj.Evaluate()
    if !ret then
        Throw("Model 'Mandatory Activity Choice' failed for " + Opts.ModelName)
    Args.(Opts.ModelName +  " Spec") = CopyArray(ret)

    vw = primarySpec.View
    // Simulate time after choice of interval is made
    simFld = Opts.SimulatedTimeField
    if simFld <> null then do
        if primarySpec.Name contains "Person" then
            set = abm.CreatePersonSet({Filter: filter, Activate: 1, UsePersonHHView: 1})
        else
            set = abm.CreateHHSet({Filter: filter, Activate: 1})

        if set.Size > 0 then do
            // Simulate duration in minutes for duration choice predicted above
            opt = null
            opt.ViewSet = vw + "|" + set.Name
            opt.InputField = Opts.ChoiceField
            opt.OutputField = simFld
            opt.AlternativeIntervalInMin = Opts.AlternativeIntervalInMin
            RunMacro("Simulate Time", opt)
        end
    end

    // Set minimum duration if specified
    minDur = Opts.MinimumDuration
    if minDur > 0 then do
        qry = printf("(%s) and (%s < %s)", {filter, Opts.SimulatedTimeField, string(minDur)})
        if primarySpec.Name contains "Person" then
            set = abm.CreatePersonSet({Filter: qry, Activate: 1, UsePersonHHView: 1})
        else
            set = abm.CreateHHSet({Filter: qry, Activate: 1})
        
        if set.Size > 0 then do
            v = Vector(set.Size, "Long", {{"Constant", minDur}})
            vecsSet = null
            vecsSet.(simFld) = v
            
            if primarySpec.Name = "HH" then
                abm.SetHHVectors(vecsSet)
            else
                abm.SetPersonVectors(vecsSet)
        end
    end
endMacro


//==========================================================================================================
// Determines availability expressions for each of the start-time alternatives
Macro "Get StartTime Avail"(altSpec, srcName)
    availArr = null
    alts = altSpec.Alternative
    availArr.Alternative = CopyArray(alts)

    for alt in alts do
        expr = srcName + ".[" + alt + "]"
        availArr.Expression =  availArr.Expression + {expr}        
    end

    Return(availArr)
endMacro
