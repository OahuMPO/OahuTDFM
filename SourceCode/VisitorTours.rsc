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
