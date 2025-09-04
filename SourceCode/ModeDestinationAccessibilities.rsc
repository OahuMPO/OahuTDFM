/*
    - Generate MC accessibiltiies (modified logsums) by mode group using an aggregate mode choice specification
      Generate accessibility matrices by auto, non motorized and transit mode groups
*/
Macro "Mandatory Accessibility"(Args)
    Args.ABMFlag = 0
    modeGroups = null
    modeGroups.Auto = {"DriveAlone", "Carpool", "Other"}
    modeGroups.NM = {"Bike", "Walk"}
    modeGroups.PT = {"W_Bus"}

    // Generate MC accessibilities from the aggregate mode choice spec
    spec = {Type: "Mandatory",
            ModeGroups: modeGroups,
            OutputFile: Args.MandatoryModeAccessibility,
            AlternativesTree: Args.MandatoryAggMCModes,
            UtilityFunction: Args.MandatoryAggMCUtility,
            Availability: Args.MandatoryAggMCAvail, 
            VOT: Args.VOT, 
            AOC: Args.AOC}
    RunMacro("MC Accessibility", Args, spec)

    // Generate DC accessibilities from the aggregate destination choice spec
    RunMacro("Mand DC Accessibility", Args)

    Return(true)
endMacro


Macro "NonMandatory Joint Accessibility"(Args)
    // These will be used as utility terms in the aggregate DC models (for accessibility).
    // Note no mode logsum term for Shop, since DC spec does not have MC accessibility.
    modeGroups = null
    modeGroups.Auto = {"Carpool"}
    modeGroups.NM = {"Bike", "Walk"}
    modeGroups.PT = {"W_Bus"}
    spec = {Type: "NonMandatory",
            ModeGroups: modeGroups,
            OutputFile: Args.NonMandJointModeAccessOther,
            UtilityFunction: Args.JointOtherAggMCUtility,
            Availability: Args.AggModeAvail
            }
    RunMacro("MC Accessibility", Args, spec)

    // Compute DC TAZ Logsums (also computes size variable for destination choice)
    purps = {"Other", "Shop"}
    for p in purps do
        opt =  {Purpose: p, 
                Type: "Joint", 
                UtilityFunction: Args.("JointTourDest" + p + "Utility"),
                SizeVarEquation: Args.("JointTour" + p + "SizeVar")
                }
        RunMacro("NonMand DC Accessibility", Args, opt)
    end
endMacro


/*
    Accessibility calculation for solo tours
*/
Macro "NonMandatory Solo Accessibility"(Args)
    // MC accessibilities first
    // Solo Other Tours
    modeGroups = null
    modeGroups.Auto = {"DriveAlone", "Carpool"}
    modeGroups.NM = {"Bike", "Walk"}
    modeGroups.PT = {"W_Bus"}
    modeGroups.NoDA = {"Carpool", "Bike", "Walk", "W_Bus"}
    
    purps = {"Other", "Shop"}
    for p in purps do
        // MC Accessibility
        spec = {Type: "NonMandatory",
                ModeGroups: modeGroups,
                OutputFile: Args.("NonMandSoloModeAccess" + p),
                UtilityFunction: Args.("Solo" + p + "AggMCUtility"),
                Availability: Args.AggModeAvail
                }
        RunMacro("MC Accessibility", Args, spec)

        // DC Accessibility for market segment of people with driver license
        opt =  {Purpose: p, 
                Type: "Solo", 
                UtilityFunction: Args.("SoloTourDest" + p + "LSUtility"),
                SizeVarEquation: Args.("SoloTour" + p + "SizeVar")
                }
        RunMacro("NonMand DC Accessibility", Args, opt)

        // DC Accessibility for market segment of people without driver license
        opt =  {Purpose: p, 
                Type: "Solo", 
                UtilityFunction: Args.("SoloTourDest" + p + "LSUtility"),
                SizeVarEquation: Args.("SoloTour" + p + "SizeVar"),
                Segment: "NoLicense"
                }
        RunMacro("NonMand DC Accessibility", Args, opt)
    end
endMacro


/*
    Given the aggregate mode choice spec, generate accessibility matrices by mode group

    - For a given mode group, disable modes not part of the group, run model and obtain the logsum matrix
    - Compute the accessibilty core as ln(1 + exp(Root_Logsum))
*/
Macro "MC Accessibility"(Args, spec)
    type = spec.Type
    modeGroups = spec.ModeGroups
    outFile = spec.OutputFile
    inputAvail = spec.Availability
    altTree = spec.AlternativesTree
    outCores = modeGroups.Map(do (f) Return(f[1]) end)

    // Create empty output MC logsum matrix
    mSpec = {TAZFile: Args.TAZGeography, OutputFile: outFile, DataType: "Double", 
             Cores: outCores, Label: type + " Mode Accessibility"}
    RunMacro("Create Empty Matrix", mSpec)
    mOutObj = CreateObject("Matrix", outFile)

    // Define mode groups
    allModes = null
    for m in modeGroups do
        allModes = allModes + m[2]
    end
    allModes = SortArray(allModes, {Unique: "True"})

    for modeGroup in modeGroups do
        mainMode = modeGroup[1]
        subModes = modeGroup[2]
        
        // Get availability parameter and modify it
        avail = null
        if inputAvail <> null then
            avail = CopyArray(inputAvail)
        
        availAlts = avail.Alternative // Array of alternatives in current availability specification
        for mode in allModes do
            pos = subModes.position(mode)
            if pos > 0 then 
                continue

            altPos = availAlts.position(mode)
            if altPos > 0 then  // replace avail
                avail.Expression[altPos] = "TAZ4Ds.TAZID.O < 0"    
            else do             // add new avail term
                avail.Alternative = avail.Alternative + {mode}
                avail.Expression = avail.Expression + {"TAZ4Ds.TAZID.O < 0"}  // Always unavailable
            end
        end

        // Run PME model
        util = null
        util.UtilityFunction = spec.UtilityFunction
        if spec.VOT <> null then
            util.SubstituteStrings = {{"<VOT>", String(spec.VOT)}}
        if spec.AOC <> null then
            util.SubstituteStrings = util.SubstituteStrings + {{"<AOC>", String(spec.AOC)}}
        util.AvailabilityExpressions = avail

        tag = type + mainMode + "AggregateMC"
        logsumFile = printf("%s\\Intermediate\\MCLogsum_%s.mtx", {Args.[Output Folder], tag})
        probFile = printf("%s\\Intermediate\\MCProb_%s.mtx", {Args.[Output Folder], tag})
        utilFile = printf("%s\\Intermediate\\MCUtil_%s.mtx", {Args.[Output Folder], tag})
        
        obj = CreateObject("PMEChoiceModel", {ModelName: tag})
        obj.OutputModelFile = printf("%s\\Intermediate\\%s.mdl", {Args.[Output Folder], tag})
        obj.AddTableSource({SourceName: "TAZ4Ds", File: Args.AccessibilitiesOutputs, IDField: "TAZID"})
        obj.AddTableSource({SourceName: "TAZData", File: Args.DemographicOutputs, IDField: "TAZ"})
        obj.AddMatrixSource({SourceName: "AutoSkim", File: Args.HighwaySkimAM, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
        w_t_skim = Args.[Output Folder] + "\\skims\\transit\\AM_w_bus.mtx"
        obj.AddMatrixSource({SourceName: "W_BusSkim", File: w_t_skim, RowIndex: "RCIndex", ColIndex: "RCIndex"})
        obj.AddMatrixSource({SourceName: "WalkSkim", File: Args.WalkSkim, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
        obj.AddMatrixSource({SourceName: "BikeSkim", File: Args.BikeSkim, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
        obj.AddMatrixSource({SourceName: "Intrazonal", File: Args.IZMatrix, RowIndex: "TAZ", ColIndex: "TAZ"})
        obj.AddPrimarySpec({Name: "AutoSkim"})
        obj.AddUtility(util)
        if altTree <> null then
            obj.AddAlternatives({AlternativesTree: altTree})
        obj.AddOutputSpec({Probability: probFile, Logsum: logsumFile, Utility: utilFile})
        ret = obj.Evaluate()

        // Get logsums, write into new matrix
        mObj = CreateObject("Matrix", logsumFile)
        mOutObj.(mainMode) := log(1 + nz(exp(mObj.Root)))
        mObj = null
    end
    mOutObj = null
    mat = null
endMacro


/*
    Generate DC accessibility vectors by mode type (Auto, NM, PT)
    Each vector represents destinations accessible by that mode
    - Compute the accessibilty vector as the log as marginal row sum of the exponentiated utility matrix
*/
Macro "Mand DC Accessibility"(Args)
    // Create empty output table
    flds = {"DestAccessibilityAuto", "DestAccessibilityNM", "DestAccessibilityPT"}
    spec = {ReferenceFile: Args.TAZGeography, OutputFile: Args.MandatoryDestAccessibility, Fields: flds}
    objOut = RunMacro("Create TAZ DC Logsum Table", spec)
    sortOrder = {{"TAZID", "Ascending"}}
    objOut.Sort({FieldArray: sortOrder})

    // Run DC model for each mode, and fill the column in the table
    modes = {"Auto", "NM", "PT"}
    for mode in modes do
        tag = "DC_" + mode

        // Compute Size Variable and fill field in output TAZ demographics table
        sizeFld = tag + "_SizeVar"
        opt = null
        opt.TableObject = CreateObject("Table", Args.DemographicOutputs)
        opt.Equation = Args.(sizeFld)   // e.g. Args.DC_Auto_SizeVar
        opt.NewOutputField = sizeFld
        RunMacro("Compute Size Variable", opt)
        opt.TableObject = null

        // Run DC Model
        tmpUtilFile = printf("%s\\Intermediate\\%s_Util.mtx", {Args.[Output Folder], tag})
        obj = CreateObject("PMEChoiceModel", {ModelName: tag})
        obj.OutputModelFile = printf("%s\\Intermediate\\%s.dcm", {Args.[Output Folder], tag})
        obj.AddTableSource({SourceName: "TAZ4Ds", File: Args.AccessibilitiesOutputs, IDField: "TAZID"})
        obj.AddTableSource({SourceName: "TAZData", File: Args.DemographicOutputs, IDField: "TAZ"})
        obj.AddMatrixSource({SourceName: "AutoSkim", File: Args.HighwaySkimAM, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
        obj.AddMatrixSource({SourceName: "Intrazonal", File: Args.IZMatrix, RowIndex: "TAZ", ColIndex: "TAZ"})
        obj.AddMatrixSource({SourceName: "ModeAccessibility", File: Args.MandatoryModeAccessibility, RowIndex: "TAZ", ColIndex: "TAZ"})
        obj.AddPrimarySpec({Name: "AutoSkim"})
        obj.AddUtility({UtilityFunction: Args.(tag + "_Utility")})
        obj.AddDestinations({DestinationsSource: "AutoSkim", DestinationsIndex: "InternalTAZ"})
        obj.AddSizeVariable({Name: "TAZData", Field: sizeFld})
        obj.AddOutputSpec({Probability: GetTempPath() + "Prob.mtx", Utility: tmpUtilFile})
        ret = obj.Evaluate()
        if !ret then
            Throw("Model Run failed while computing TAZ destination logsums for " + tag)

        // Open Utility Matrix, add exp(u) core, get row marginal sums, apply ln() and write to output table
        mObj = CreateObject("Matrix", tmpUtilFile)
        mObj.AddCores({"ExpUtil"})
        coreNames = mObj.GetCoreNames()
        mObj.ExpUtil := exp(mObj.(coreNames[1]))
        vLS = mObj.GetVector({Core: "ExpUtil", Marginal: "Row Sum"})
        vLS = if vLS = null then null else log(vLS)
        mObj = null

        outFld = "DestAccessibility" + mode
        objOut.(outFld) = vLS
    end
    objOut = null
endMacro


/*
    Generate DC accessibility vectors for non mandatory tour
    Aggregate DC model assumed
    - Compute the accessibilty vector as the log as marginal row sum of the exponentiated utility matrix
*/
Macro "NonMand DC Accessibility"(Args, spec)
    type = spec.Type             // One of 'Solo', 'Joint'
    purp = spec.Purpose          // One of 'Other', 'Shop'
    utilEqn = spec.UtilityFunction
    segment = spec.Segment
    if segment <> null then do
        tag = type + purp + "_" + segment
        segment = Lower(segment)
    end
    else
        tag = type + purp
    mcLSMtx = Args.("NonMand" + type + "ModeAccess" + purp)

    // Compute size variable field. Add to TAZDemographics output table
    objD = CreateObject("Table", Args.DemographicOutputs)
    obj4D = CreateObject("Table", Args.AccessibilitiesOutputs)
    sizeFld = type + purp + "Size"
    newFlds = {{FieldName: sizeFld, Type: "Real", Width: 12, Decimals: 2}}
    objD.AddFields({Fields: newFlds})
    objJ = objD.Join({Table: obj4D, LeftFields: {"TAZ"}, RightFields: {"TAZID"}})
    
    opt = null
    opt.TableObject = objJ
    opt.Equation = spec.SizeVarEquation
    opt.FillField = sizeFld
    opt.ExponentiateCoeffs = 1
    RunMacro("Compute Size Variable", opt)
    objJ = null
    objD = null
    obj4D = null

    // Run DC Model
    tmpUtilFile = printf("%s\\Intermediate\\%s_Util.mtx", {Args.[Output Folder], tag})
    obj = CreateObject("PMEChoiceModel", {ModelName: tag})
    if segment <> null then
        obj.Segment = segment
    obj.OutputModelFile = printf("%s\\Intermediate\\%s.dcm", {Args.[Output Folder], tag})
    obj.AddTableSource({SourceName: "TAZData", File: Args.DemographicOutputs, IDField: "TAZ"})
    obj.AddTableSource({SourceName: "TAZ4Ds", File: Args.AccessibilitiesOutputs, IDField: "TAZID"})
    obj.AddMatrixSource({SourceName: "AutoSkim", File: Args.HighwaySkimOP, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
    obj.AddMatrixSource({SourceName: "WalkSkim", File: Args.WalkSkim, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
    obj.AddMatrixSource({SourceName: "Intrazonal", File: Args.IZMatrix, RowIndex: "TAZ", ColIndex: "TAZ"})
    if mcLSMtx <> null then
        obj.AddMatrixSource({SourceName: "ModeAccessibility", File: mcLSMtx, RowIndex: "TAZ", ColIndex: "TAZ"})
    obj.AddPrimarySpec({Name: "AutoSkim"})
    obj.AddUtility({UtilityFunction: utilEqn})
    obj.AddDestinations({DestinationsSource: "AutoSkim", DestinationsIndex: "InternalTAZ"})
    obj.AddSizeVariable({Name: "TAZData", Field: sizeFld})
    obj.AddOutputSpec({Probability: GetTempPath() + "Prob.mtx", Utility: tmpUtilFile})
    ret = obj.Evaluate()
    if !ret then
        Throw("Model Run failed while computing TAZ destination logsums for " + tag)

    // Open Utility Matrix, add exp(u) core, get row marginal sums, apply ln() and write to output table
    mObj = CreateObject("Matrix", tmpUtilFile)
    mObj.AddCores({"ExpUtil"})
    coreNames = mObj.GetCoreNames()
    mObj.ExpUtil := exp(mObj.(coreNames[1]))
    vLS = mObj.GetVector({Core: "ExpUtil", Marginal: "Row Sum"})
    vLS = if vLS = null then null else log(vLS)
    mObj = null

    objOut = CreateObject("Table", Args.NonMandatoryDestAccessibility)
    sortOrder = {{"TAZID", "Ascending"}}
    objOut.Sort({FieldArray: sortOrder})
    outFld = "Accessibility_" + tag
    objOut.(outFld) = vLS
endMacro


// Create empty DC Logsum Table
Macro "Create TAZ DC Logsum Table"(spec)
    // Create empty output table
    objTAZ = CreateObject("Table", spec.ReferenceFile)
    objOut = objTAZ.Export({FileName: spec.OutputFile, FieldNames: {"TAZID"}})
    flds = spec.Fields
    fldsAdd = flds.Map(do(f) Return({FieldName: f, Type: "real", Width: 12, Decimals: 2}) end)
    objOut.AddFields({Fields: fldsAdd})
    sortOrder = {{"TAZID", "Ascending"}}
    objOut.Sort({FieldArray: sortOrder})
    objTAZ = null
    Return(objOut)
endMacro
