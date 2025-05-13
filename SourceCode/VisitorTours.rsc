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
