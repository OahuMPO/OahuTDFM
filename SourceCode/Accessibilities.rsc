
macro "CalculateAccessibilities" (Args, Results)
    ret_value = 1
    Args.ABMFlag = 0
    ret_value = RunMacro("Determine Intersections", Args)
    if !ret_value then goto quit
    // input data files
    TAZDB = Args.TAZGeography
    TAZData = Args.DemographicOutputs
    LineDB = Args.HighwayDatabase
    RouteSystem = Args.TransitRoutes
    skim_dir = Args.[Output Folder] + "\\skims"

    Line = CreateObject("Table", {FileName: LineDB, LayerType: "Line"})
    Node = CreateObject("Table", {FileName: LineDB, LayerType: "Node"})
    Node.CreateSet({SetName: "Int", Filter: "NumLinks > 2"})
    NodeLayer = Node.GetView()
    TAZGeo = CreateObject("Table", TAZDB)
    TAZGeo.Sort({FieldArray: {{"TAZID", "Ascending"}}})
    TAZBin = CreateObject("Table", TAZData)
    TAZO = TAZGeo.Join({Table: TAZBin, LeftFields: "TAZID", RightFields: "TAZ"})
    TAZLayer = TAZGeo.GetView() 
    JoinTAZView = TAZO.GetView() 
    TAZFlds = TAZO.GetFieldSpecs()
    TAZPointFile = GetTempFileName("*.dbd") 
    ExportGeography(TAZLayer + "|", TAZPointFile, {"Field Spec": TAZFlds, Centroid: "True"})
    TAZP = CreateObject("Table", TAZPointFile)
    TAZP.Sort({FieldArray: {{"TAZID", "Ascending"}}})
    ptlayer = TAZP.GetView() 
    
    fields = {
   {FieldName: "NIntersections"}, 
    {FieldName: "IntersectionDensity"}, 
    {FieldName: "AttractionDensity"}, 
    {FieldName: "RetailEmp"},
    {FieldName: "ServiceEmp"},
    {FieldName: "HH"},
    {FieldName: "Pop"}, 
    {FieldName: "TotalEmp"}, 
    {FieldName: "AgricultureEmp"},
    {FieldName: "ManufactureEmp"},
    {FieldName: "WholesaleEmp"},
    {FieldName: "TransportConstructionEmp"},
    {FieldName: "BasEmpDensity"}, 
    {FieldName: "IndEmpDensity"}, 
    {FieldName: "ZScoreIntersection"}, 
    {FieldName: "ZScoreAttraction"}, 
    {FieldName: "ZScoreBasic"}, 
    {FieldName: "ZScoreInd"},
    {FieldName: "ConnectivityIndex"},
    {FieldName: "NumLocalIntersections"},
    {FieldName: "ThreewayIntersections"},
    {FieldName: "FourWayIntersections"},
    {FieldName: "HouseholdDensity"},
    {FieldName: "RetailEmploymentDensity"},
    {FieldName: "RetailAndServiceEmploymentDensity"},
    {FieldName: "TotalEmploymentDensity"},
    {FieldName: "RetailEmploymentAndHouseholdDiversity"},
    {FieldName: "RetailServiceEmpHHDiversity"},
    {FieldName: "JobsHousingDiversity"},
    {FieldName: "JobMixDiversity"},
    {FieldName: "MixRetailIntersectionDensity"},
    {FieldName: "TransitStopDensity"},
    {FieldName: "WalkIndex"}, 
    {FieldName: "WalkIndexBasIndLowered"}, 
    {FieldName: "TransitAccessibilityToJobsAM"},
    {FieldName: "TransitAccessibilityToRetailAM"},
    {FieldName: "AutoAccessibilityToJobsAM"},
    {FieldName: "AutoAccessibilityToRetailAM"},
    {FieldName: "NonMotorizedAccessibility"}
    }
    TAZP.AddFields({Fields: fields})
     
    aggsz = {{"AgricultureEmp", "SUM", "EMP_Agriculture",},
            {"ManufactureEmp", "SUM", "EMP_Manufacturing",},
            {"WholesaleEmp", "SUM", "EMP_Wholesale",},
            {"RetailEmp", "SUM", "Emp_Retail", }, 
            {"TransportConstructionEmp", "SUM", "EMP_TransportConstruction",},
            {"RetailEmp", "SUM", "Emp_Retail", }, 
            {"Pop", "SUM", "Population", }, 
            {"HH", "SUM", "OccupiedHH",},
            {"ServiceEmp", "SUM", "Emp_Services",},
            {"TotalEmp", "SUM", "TotalEmployment",}}

    aggsn =  {{"NIntersections", "SUM", "Numlinks", },
              {"ThreewayIntersections", "SUM", "IntersectionThreeway",},
              {"FourWayIntersections", "SUM", "IntersectionFourway",},
              {"NumLocalIntersections", "SUM", "LocalIntersection",}}

    ColumnAggregate(ptlayer+"|", 0.5, NodeLayer+"|Int",aggsn, )
    ColumnAggregate(ptlayer+"|", 0.5, TAZLayer+"|", aggsz, )
    area = 3.14159 * 0.5 * 0.5
    TAZP.IntersectionDensity = (TAZP.ThreewayIntersections + TAZP.FourWayIntersections) / area
    TAZP.AttractionDensity = (4.1 * TAZP.RetailEmp + TAZP.Pop) / area
    TAZP.BasEmpDensity = (TAZP.AgricultureEmp + TAZP.ManufactureEmp + TAZP.WholesaleEmp + TAZP.TransportConstructionEmp) / area
    TAZP.IndEmpDensity = (TAZP.AgricultureEmp + TAZP.ManufactureEmp) / area
    avgintden = TAZP.Statistics({FieldName: "IntersectionDensity", Method: "Mean"})
    stdintden = TAZP.Statistics({FieldName: "IntersectionDensity", Method: "Sdev"})
    avgattden = TAZP.Statistics({FieldName: "AttractionDensity", Method: "Mean"})
    stdattden = TAZP.Statistics({FieldName: "AttractionDensity", Method: "Sdev"})
    avgbasden = TAZP.Statistics({FieldName: "BasEmpDensity", Method: "Mean"})
    stdbasden = TAZP.Statistics({FieldName: "BasEmpDensity", Method: "Sdev"})
    avgindden = TAZP.Statistics({FieldName: "IndEmpDensity", Method: "Mean"})
    stdindden = TAZP.Statistics({FieldName: "IndEmpDensity", Method: "Sdev"})
    TAZP.ZScoreIntersection = (TAZP.IntersectionDensity - avgintden) / stdintden
    TAZP.ZScoreAttraction = (TAZP.AttractionDensity - avgattden) / stdattden
    TAZP.ZScoreBasic = (TAZP.BasEmpDensity - avgbasden) / stdbasden
    TAZP.ZScoreInd = (TAZP.IndEmpDensity - avgindden) / stdindden
    TAZP.WalkIndex = 1.0 / (1 + exp(-(TAZP.ZScoreIntersection + TAZP.ZScoreAttraction - TAZP.ZScoreBasic - TAZP.ZScoreInd)))   
    TAZP.WalkIndexBasIndLowered = 1.0 / (1 + exp(-(TAZP.ZScoreIntersection + TAZP.ZScoreAttraction - 0.25*TAZP.ZScoreBasic - 0.25*TAZP.ZScoreInd)))   

    RegionalEmp = TAZO.Statistics({FieldName: "TotalEmployment", Method: "Sum"})
    AvgEmp = TAZO.Statistics({FieldName: "TotalEmployment", Method: "Mean"})
    RegionalHH = TAZO.Statistics({FieldName: "OccupiedHH", Method: "Sum"})
    AvgHH = TAZO.Statistics({FieldName: "OccupiedHH", Method: "Mean"})
    RegionalRetEmp = TAZO.Statistics({FieldName: "Emp_Retail", Method: "Sum"})
    AvgLocInt = TAZP.Statistics({FieldName: "NumLocalIntersections", Method: "Mean"})

    TAZP.ConnectivityIndex = if TAZP.NumLocalIntersections > 0 then TAZP.FourWayIntersections / TAZP.NumLocalIntersections else 0

    TAZP.HouseholdDensity = TAZP.HH / area
    TAZP.RetailEmploymentDensity = TAZP.RetailEmp / area

    TAZP.RetailAndServiceEmploymentDensity = (TAZP.ServiceEmp + TAZP.RetailEmp)/ area
    TAZP.TotalEmploymentDensity = TAZP.TotalEmp / area

    TAZP.RetailServiceEmpHHDiversity = (0.001 * (TAZP.ServiceEmp + TAZP.RetailEmp)) * TAZP.HH / (TAZP.ServiceEmp + TAZP.RetailEmp + TAZP.HH)
    TAZP.RetailEmploymentAndHouseholdDiversity = 0.001 * TAZP.RetailEmp * TAZP.HH / (TAZP.RetailEmp + TAZP.HH)

    B = RegionalEmp / RegionalHH
    TAZP.JobsHousingDiversity = 1 - (Abs(B * TAZP.HH - TAZP.TotalEmp) / (B * TAZP.HH + TAZP.TotalEmp))
    B = if RegionalRetEmp > 0 then (RegionalEmp - RegionalRetEmp) / RegionalRetEmp else 0

    TAZP.JobMixDiversity = 1 - (Abs(B * TAZP.RetailEmp - (TAZP.TotalEmp - TAZP.RetailEmp)) / (B * TAZP.RetailEmp + (TAZP.TotalEmp - TAZP.RetailEmp)))
    a = AvgLocInt / AvgEmp
    b = AvgLocInt / AvgHH

    TAZP.MixRetailIntersectionDensity = if TAZP.NumLocalIntersections * (TAZP.TotalEmp * a)  * (TAZP.HH * b ) = 0 then null else Log(TAZP.NumLocalIntersections * (TAZP.TotalEmp * a) * (TAZP.HH * b)) / (TAZP.NumLocalIntersections + (TAZP.TotalEmp * a) + (TAZP.HH * b ))

    times = {"AM"}
    
    for time in times do

        // open transit skim matrix
        PTSkim = skim_dir + "\\transit\\AM_w_bus.mtx"
        AutoSkim = Args.HighwaySkimAM
        // make employment and retail employment vectors row based
        PTMat = CreateObject("Matrix", PTSkim)
        PTMat.SetRowIndex("RCIndex")
        PTMat.SetColIndex("RCIndex")
 
        AUTOMat = CreateObject("Matrix", AutoSkim) 
        AUTOMat.SetRowIndex("InternalTAZ")
        AUTOMat.SetColIndex("InternalTAZ")
       
        PTMat.AddCores({"Scratch", "Decay"}) 
        AUTOMat.AddCores({"Scratch", "Decay"}) 
        emp_Jobs = CopyVector(TAZO.TotalEmployment)
        emp_Jobs.RowBased = true
        emp_Retail = CopyVector(TAZO.Emp_Retail)
        emp_Retail.RowBased = true

        PTMat.Scratch := 1.0
        AUTOMat.Scratch := 1.0
        PTMat.Scratch := if PTMat.[Total Time] <> null and PTMat.[Total Time] < 30 then emp_Jobs * PTMat.Scratch else null      
        AUTOMat.Scratch := if AUTOMat.Time <> null and AUTOMat.Time < 30 then emp_Jobs * AUTOMat.Scratch else null  
        // calculate total employment accessible to zone
        origtotemploy = PTMat.GetVector({Core: "Scratch" , Marginal: "Row Sum"}) 
        origtotemploya = AUTOMat.GetVector({Core: "Scratch" , Marginal: "Row Sum"}) 
//        PTMat.Decay := if PTMat.[Total Time] < DecayTransitStartTime and PTMat.[Total Time] <> null then 1.0 else if PTMat.[Total Time] <> null then DecayTransitA * exp(DecayTransitC * PTMat.[Total Time]) * pow(PTMat.[Total Time], DecayTransitB) else null
//        AUTOMat.Decay := if AUTOMat.Time < DecayAutoStartTime and AUTOMat.Time <> null then 1.0 else if AUTOMat.Time <> null then DecayAutoA * exp(DecayAutoC * AUTOMat.Time) * pow(AUTOMat.Time, DecayAutoB) else null
        PTMat.Scratch := if PTMat.[Total Time] <> null then emp_Jobs * PTMat.Decay else null      
        AUTOMat.Scratch := if AUTOMat.Time <> null then emp_Jobs * AUTOMat.Decay else null  
//        origtotemployDecay = PTMat.GetVector({Core: "Scratch" , Marginal: "Row Sum"}) 
//        origtotemployaDecay = AUTOMat.GetVector({Core: "Scratch" , Marginal: "Row Sum"}) 

        PTMat.Scratch := 1.0
        AUTOMat.Scratch := 1.0
        
        
        // if transit total time is less than 30 minutes then all retail employment is accessible to the zone, otherwise no retail employment is accessible
        PTMat.Scratch := if PTMat.[Total Time] <> null and PTMat.[Total Time] < 30 then emp_Retail * PTMat.Scratch else null      
        // if auto total time is less than 30 minutes then all retail employment is accessible to the zone, otherwise no retail employment is accessible
        AUTOMat.Scratch := if AUTOMat.Time <> null and AUTOMat.Time < 30 then emp_Retail * AUTOMat.Time else null  
        // calculate total employment accessible to zone
        origretemploy = PTMat.GetVector({Core: "Scratch" , Marginal: "Row Sum"}) 
        origretemploya = AUTOMat.GetVector({Core: "Scratch" , Marginal: "Row Sum"}) 
//        PTMat.Scratch := if PTMat.[Total Time] <> null then emp_Retail * PTMat.Decay else null      
//        AUTOMat.Scratch := if AUTOMat.Time <> null then emp_Retail * AUTOMat.Decay else null  
//        origretemployDecay = PTMat.GetVector({Core: "Scratch" , Marginal: "Row Sum"}) 
//        origretemployaDecay = AUTOMat.GetVector({Core: "Scratch" , Marginal: "Row Sum"}) 

        WalkSkim = Args.WalkSkim
        WALKMat = CreateObject("Matrix", WalkSkim) 
        WALKMat.SetRowIndex("InternalTAZ")
        WALKMat.SetColIndex("InternalTAZ")

        PTMat.Scratch := 1.0

        // if walk time time is less than 30 minutes then all  employment is accessible to the zone, otherwise no  employment is accessible
        PTMat.Scratch := if WALKMat.Time <> null and WALKMat.Time < 30 then emp_Jobs else null
         // calculate total employment accessible to zone
        origwalkemploy = PTMat.GetVector({Core: "Scratch" , Marginal: "Row Sum"}) 
//        PTMat.Decay := if WALKMat.WalkTime < DecayWalkStartTime and WALKMat.WalkTime <> null then 1.0 else if WALKMat.WalkTime <> null then DecayWalkA * exp(DecayWalkC * WALKMat.WalkTime) * pow(WALKMat.WalkTime, DecayWalkB) else null
//        PTMat.Scratch := if WALKMat.WalkTime <> null then emp_Jobs * PTMat.Decay else null
//        origwalkemployDecay = PTMat.GetVector({Core: "Scratch" , Marginal: "Row Sum"}) 
        // delete temporary matrix if it sill exists
        PTMat.DropCores({"Scratch", "Decay"})
        AUTOMat.DropCores({"Scratch", "Decay"})

      
        // write total employment accessibility, retail employment accessibility, and non-motorized accessibiilty values
        TAZP.SetDataVectors({
            FieldData: {{"TransitAccessibilityToJobs"+time, origtotemploy}, 
                        {"TransitAccessibilityToRetail"+time, origretemploy}, 
                        {"AutoAccessibilityToJobs"+time, origtotemploya},
                        {"AutoAccessibilityToRetail"+time, origretemploya}, 
                        {"NonMotorizedAccessibility", origwalkemploy}}, 
            Options: {SortOrder: {{"TAZ", "Ascending"}}}} )
        end
 // calculate all operations
    // add route system, stops, line, node, and physical stops layer
    // add taz layer
    STOPS = CreateObject("Table", {FileName: RouteSystem, LayerType: "Stop"})
    StopsLayer = STOPS.GetView()
    
    // put in a field of 1's in the physical stop layer
    ones = CreateExpression(StopsLayer, "ONE", "1", )
    // determine number of physical stops in each TAZ
    ColumnAggregate(ptlayer+"|", 0.5, StopsLayer+"|", {{"TransitStopDensity", "Sum", ones, }}, null)

    // calculate transit stop density (number of stops / buffer area)
    TAZP.TransitStopDensity = TAZP.TransitStopDensity / area
    AccessExportTable = Args.AccessibilitiesOutputs

    keepflds = {"TAZID", "NIntersections", 
    "IntersectionDensity", 
    "AttractionDensity", 
    "BasEmpDensity", 
    "IndEmpDensity", 
    "ZScoreIntersection", 
    "ZScoreAttraction", 
    "ZScoreBasic", 
    "ZScoreInd",
    "ConnectivityIndex",
    "NumLocalIntersections",
    "ThreewayIntersections",
    "FourWayIntersections",
    "HouseholdDensity",
    "RetailEmploymentDensity",
    "RetailAndServiceEmploymentDensity",
    "TotalEmploymentDensity",
    "RetailEmploymentAndHouseholdDiversity",
    "RetailServiceEmpHHDiversity",
    "JobsHousingDiversity",
    "JobMixDiversity",
    "MixRetailIntersectionDensity",
    "TransitStopDensity",
    "WalkIndex", 
    "WalkIndexBasIndLowered", 
    "TransitAccessibilityToJobsAM",
    "TransitAccessibilityToRetailAM",
    "AutoAccessibilityToJobsAM",
    "AutoAccessibilityToRetailAM",
    "NonMotorizedAccessibility"}
    TAZPExp = TAZP.Export({FileName: AccessExportTable, FieldNames: keepflds})
    quit:
    Return(ret_value)
endmacro


macro "Determine Intersections" (Args)
    ret_value = 1
    LineDB = Args.HighwayDatabase
    Line = CreateObject("Table", {FileName: LineDB, LayerType: "Line"})
    Node = CreateObject("Table", {FileName: LineDB, LayerType: "Node"})
    fields = {
            {FieldName: "NumLinks", Type: "integer"}, 
            {FieldName: "IntersectionThreeway", Type: "integer"}, 
            {FieldName: "IntersectionFourway", Type: "integer"}, 
            {FieldName: "LocalIntersection", Type: "integer"}
            }
    Node.AddFields({Fields: fields})
    nlyr = Node.GetView()
    llyr = Line.GetView()
    SetLayer(nlyr)
    rec = GetFirstRecord(nlyr + "|",)
    while rec <> null do
        nodeid = nlyr.ID
        SetLayer(nlyr)
        links = GetNodeLinks(nodeid)
        nlinks = 0
        for link in links do
            linkrec = ID2RH(link)
            SetRecord(llyr,linkrec)
            cls = llyr.HCMType
            if cls <> 'Freeway' and cls <> 'Expressway' and cls <> 'Ramp' and cls <> 'CC' then
                nlinks = nlinks + 1
            end
        nlyr.NumLinks = nlinks
        if nlinks  = 3 then nlyr.IntersectionThreeway = 1
        if nlinks >= 4 then nlyr.IntersectionFourway = 1
        if nlinks > 1 then nlyr.LocalIntersection = 1
        rec = GetNextRecord(nlyr + "|", rec, )
        end
    quit:
    Return(ret_value)
EndMacro

