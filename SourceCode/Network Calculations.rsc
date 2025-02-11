/*

*/

Macro "Network Calculations" (Args)
    RunMacro("CopyDataToOutputFolder", Args)
    RunMacro("Filter Transit Modes", Args)
    RunMacro("Expand DTWB", Args)
    RunMacro("Calculate Toll Cost", Args)
    RunMacro("Mark KNR Nodes", Args)
    RunMacro("Determine Area Type", Args)
    RunMacro("Add Cluster Info to SE Data", Args)
    RunMacro("Create Visitor Clusters", Args)
    RunMacro("Speeds and Capacities", Args)
    RunMacro("CalculateTransitSpeeds Oahu", Args)
    RunMacro("Compute Intrazonal Matrix", Args)
    RunMacro("Create Intra Cluster Matrix", Args)

    return(1)
endmacro

// copies highway network, demographics, and transit routes to output folder for further processing
macro "CopyDataToOutputFolder" (Args)
    ret_value = 1
    hwyinputdb = Args.HighwayInputDatabase 
    hwyoutputdb = Args.HighwayDatabase 
    DemoMaster = Args.Demographics
    DemoOut = Args.DemographicOutputs
    rs_filemaster = Args.TransitRouteInputs
    rs_file = Args.TransitRoutes
    CopyDatabase(hwyinputdb, hwyoutputdb)   // copy master network to copy
    CopyTableFiles(null, "FFB", DemoMaster, , DemoOut, )    // copy master demographics to copy
    tab = CreateObject("Table", DemoOut)
    dropflds = {"EMP_NAICS_11", "EMP_NAICS_21", "EMP_NAICS_22", "EMP_NAICS_23", "EMP_NAICS_31-33", "EMP_NAICS_42", "EMP_NAICS_44-45", "EMP_NAICS_48-49", "EMP_NAICS_51", 
                "EMP_NAICS_52", "EMP_NAICS_53", "EMP_NAICS_54", "EMP_NAICS_55", "EMP_NAICS_56", "EMP_NAICS_61", "EMP_NAICS_62", "EMP_NAICS_71", "EMP_NAICS_72", 
                "EMP_NAICS_81", "EMP_NAICS_92"}
    tab.DropFields({FieldNames: dropflds})
        // add master transit network

    o = CreateObject("DataManager")
    o.AddDataSource("RS", {FileName: rs_filemaster, DataType: "RS"})
    routeLayers = o.GetRouteLayers("RS")

    RouteLayer = routeLayers.RouteLayer
    StopLayer = routeLayers.StopLayer
    LineLayer = routeLayers.LineLayer
    NodeLayer = routeLayers.NodeLayer

    SetLayer(RouteLayer)
    
    args = null
    args.RouteLayer       = RouteLayer
    args.RouteSet         = null
    args.StopSet          = null
    args.TaggedField      = "NodeID"
    args.RepeatedSetName  = "Repeated Stop Tags"
    args.NotTaggedSetName = "Not Tagged"
    args.CreateLog        = true
    // check for repeated tagged nodes and flag if found
    TagInfo = RunMacro("Select Repeated Tagged Nodes", args)
    
    if TagInfo.errorMessage <> null then Throw(GetLastError())
    // create copy of route system

    o.CopyRouteSystem("RS", {TargetRS: rs_file, Settings: {Geography: hwyoutputdb}})
    Return(ret_value)
endmacro

/*
Remove network columns from the mode table that if that mode doesn't exist in 
the scenario. This will in turn control which networks (tnw) get created.
*/

Macro "Filter Transit Modes" (Args)
    rts_file = Args.TransitRoutes
    mode_file = Args.TransitModeTable

    // Open a map and determine if rail is present
    map = CreateObject("Map", rts_file)
    {nlyr, llyr, rlyr, slyr} = map.GetLayerNames()
    n = map.SelectByQuery({
        SetName: "rail",
        Query: "Mode = 7"
    })
    map = null

    if n = 0 then do
        // Can't edit CSVs directly, so convert to a MEM table
        temp = CreateObject("Table", mode_file)
        tbl = temp.Export({ViewName: "temp"})
        tbl.DropFields("rail")
        temp = null
        tbl.Export({FileName: mode_file})
    end

    // Remove modes from MC parameter files
    RunMacro("Filter Visitor Transit Modes", Args)
endmacro

/*
Removes modes from the visitor csv parameter files if they don't exist
in the scenario.
*/

Macro "Filter Visitor Transit Modes" (Args)
    
    mode_table = Args.TransitModeTable
    // access_modes = Args.access_modes
    mc_dir = Args.[Input Folder] + "/visitors/mc"
    
    transit_modes = RunMacro("Get Transit Net Def Col Names", mode_table)

    // get visitor trip purposes
    factor_file = Args.VisOccupancyFactors
    fac_tbl = CreateObject("Table", factor_file)
    trip_types = fac_tbl.trip_type
    trip_types = SortVector(trip_types, {Unique: "true"})

    for trip_type in trip_types do
        coef_file = mc_dir + "/" + trip_type + ".csv"
        coef_tbl = CreateObject("Table", coef_file)

        // Start by selecting all non-transit modes
        coef_tbl.SelectByQuery({
            SetName: "export",
            Query: "Alternative = 'auto' or Alternative = 'tnc' or Alternative = 'walk' or Alternative = 'bike'"
        })
        // Now add transit modes that exist to the selection
        for mode in transit_modes do
            coef_tbl.SelectByQuery({
                SetName: "export",
                Operation: "more",
                Query: "Alternative = '" + mode + "'"
            })  
        end
        temp = coef_tbl.Export({ViewName: "temp"})
        coef_tbl = null
        temp.Export({FileName: coef_file})
    end
endmacro

/*
The DTWB field has a string that says whether

Drive
Transit
Walk
Bike

is available. This macro converts that into one-hot fields.
*/

Macro "Expand DTWB" (Args)
    
    hwy_dbd = Args.HighwayDatabase

    tbl = CreateObject("Table", {FileName: hwy_dbd, Layer: 2})
    tbl.AddField("D")
    tbl.AddField("T")
    tbl.AddField("W")
    tbl.AddField("B")
    
    v_dtwb = tbl.DTWB
    set = null
    set.D = if Position(v_dtwb, "D") <> 0 then 1 else 0
    set.T = if Position(v_dtwb, "T") <> 0 then 1 else 0
    set.W = if Position(v_dtwb, "W") <> 0 then 1 else 0
    set.B = if Position(v_dtwb, "B") <> 0 then 1 else 0
    tbl.SetDataVectors({FieldData: set})
endmacro

/*
Converts toll rates ($/mi) into the toll cost for using a link. These
total cost fields are what is used in skimming.
*/

Macro "Calculate Toll Cost" (Args)

    hwy_dbd = Args.HighwayDatabase
    periods = {"AM", "PM", "OP"}

    tbl = CreateObject("Table", {FileName: hwy_dbd, Layer: 2})
    for period in periods do
        tbl.AddField("TollCostSOV" + period)
        tbl.AddField("TollCostHOV" + period)
        v_tolltype = tbl.TollType
        v_tollrate = tbl.("TollRate" + period)
        v_length = tbl.Length
        v_tollcost = nz(v_tollrate) * v_length
        tbl.("TollCostSOV" + period) = if v_tolltype = "Free" then 0 else v_tollcost
        tbl.("TollCostHOV" + period) = if v_tolltype = "Free" or v_tolltype = "HOT" 
            then 0 else v_tollcost
    end
endmacro

/*
Marks nodes with bus stops as KNR. Also marks KNR nodes within microtransit
districts as MT nodes.
*/

Macro "Mark KNR Nodes" (Args)

    rts_file = Args.TransitRoutes

    map = CreateObject("Map", rts_file)
    {nlyr, llyr, rlyr, slyr} = map.GetLayerNames()

    node = CreateObject("Table", nlyr)
    node.AddField({
        FieldName: "KNR",
        Description: "If node can be considered for KNR|(If a bus stop is at the node)"
    })
    links = CreateObject("Table", llyr)
    links.SelectByQuery({
        SetName: "drive_links",
        Query: "Select * where D  = 1 and HCMType <> 'Freeway'"
    })
    SetLayer(nlyr)
    SelectByLinks("drive nodes", "several", "drive_links", )
    n = SelectByVicinity ("knr", "several", slyr + "|", 10/5280, {"Source And": "drive nodes"})
    v = Vector(n, "Long", {Constant: 1})
    node.ChangeSet("knr")
    node.KNR = v

    // Transfer MT dist info to the TAZ layer
    mt_exists = RunMacro("MT Districts Exist?", Args)
    taz = CreateObject("Table", Args.TAZGeography)
    taz.AddField({
        FieldName: "MTDist",
        Description: "The microtransit district number"
    })
    taz.AddField({
        FieldName: "MTFare",
        Description: "The microtransit district fare"
    })
    taz.AddField({
        FieldName: "MTHeadway",
        Description: "The microtransit district headway"
    })
    taz_specs = taz.GetFieldSpecs({NamedArray: TRUE})
    se = CreateObject("Table", Args.DemographicOutputs)
    se_specs = se.GetFieldSpecs({NamedArray: TRUE})
    join = taz.Join({
        Table: se,
        LeftFields: "TAZID",
        RightFields: "TAZ"
    })
    join.(taz_specs.MTDist) = join.(se_specs.MTDist)
    join.(taz_specs.MTFare) = join.(se_specs.MTFare)
    join.(taz_specs.MTHeadway) = join.(se_specs.MTHeadway)
    join = null
    node.AddField({
        FieldName: "MTDist",
        Description: "The microtransit district number"
    })
    node.AddField({
        FieldName: "MTFare",
        Description: "The microtransit district fare"
    })
    node.AddField({
        FieldName: "MTHeadway",
        Description: "The microtransit district headway"
    })
    // If there are no MT districts, don't bother with the rest of this macro
    if !mt_exists then return()
    
    // Mark nodes within MT districts as MT nodes
    {tlyr} = map.AddLayer({FileName: Args.TAZGeography})
    tazs = CreateObject("Table", tlyr)
    tazs.SelectByQuery({
        SetName: "mt_tazs",
        Query: "Select * where MTDist > 0"
    })
    SetLayer(nlyr)
    n = SelectByVicinity ("mt_knr", "several", tlyr + "|mt_tazs", 0, {"Source And": "knr"})
    if n = 0 then Throw("No knr nodes found in MT districts")
    TagLayer("Value", nlyr + "|mt_knr", nlyr + ".MTDist", tlyr + "|mt_tazs", tlyr + ".MTDist")
    TagLayer("Value", nlyr + "|mt_knr", nlyr + ".MTFare", tlyr + "|mt_tazs", tlyr + ".MTFare")
    TagLayer("Value", nlyr + "|mt_knr", nlyr + ".MTHeadway", tlyr + "|mt_tazs", tlyr + ".MTHeadway")
endmacro

/*
Prepares input options for the AreaType.rsc library of tools, which
tags TAZs and Links with area types.
*/

Macro "Determine Area Type" (Args)

    scen_dir = Args.[Scenario Folder]
    taz_dbd = Args.TAZGeography
    se_bin = Args.DemographicOutputs
    hwy_dbd = Args.HighwayDatabase
    area_tbl = Args.AreaTypes

    // Get area from TAZ layer
    {map, {taz_lyr}} = RunMacro("Create Map", {file: taz_dbd})

    // Calculate total employment and density
    se_vw = OpenTable("se", "FFB", {se_bin, })
    a_fields =  {
        {"TotalEmployment", "Integer", 10, ,,,, "Total employment"},
        {"Density", "Real", 10, 2,,,, "Density used in area type calculation.|Considers HH and Emp."},
        {"AreaType", "Character", 10,,,,, "Area Type"},
        {"ATSmoothed", "Integer", 10,,,,, "Whether or not the area type was smoothed"},
        {"EmpDensity", "Real", 10, 2,,,, "Employment density. Used in some DC models.|TotalEmp / Area."}
    }
    RunMacro("Add Fields", {view: se_vw, a_fields: a_fields})

    // Join the se to TAZ
    jv = JoinViews("jv", taz_lyr + ".TAZID", se_vw + ".TAZ", )

    data = GetDataVectors(
        jv + "|",
        {
            "Area",
            "Population",
            "GroupQuarterPopulation",
            "Emp_Agriculture",
            "Emp_Manufacturing",
            "Emp_Wholesale",
            "Emp_Retail",
            "Emp_TransportConstruction",
            "Emp_FinanceRealEstate",
            "Emp_Education",
            "Emp_HealthCare",
            "Emp_Services",
            "Emp_Public",
            "Emp_Hotel",
            "Emp_Military"
        },
        {OptArray: TRUE, "Missing as Zero": TRUE}
    )
    tot_emp = data.Emp_Agriculture + 
        data.Emp_Manufacturing + 
        data.Emp_Wholesale + 
        data.Emp_Retail + 
        data.Emp_TransportConstruction + 
        data.Emp_FinanceRealEstate + 
        data.Emp_Education + 
        data.Emp_HealthCare + 
        data.Emp_Services + 
        data.Emp_Public + 
        data.Emp_Hotel + 
        data.Emp_Military
    data.HH_POP = data.Population - data.GroupQuarterPopulation
    factor = data.HH_POP.sum() / tot_emp.sum()
    density = (data.HH_POP + tot_emp * factor) / data.area
    emp_density = tot_emp / data.area
    areatype = Vector(density.length, "String", )
    for i = 1 to area_tbl.length do
        name = area_tbl[i].AreaType
        cutoff = area_tbl[i].Density
        areatype = if density >= cutoff then name else areatype
    end
    // SetDataVector(jv + "|", "TotalEmp", tot_emp, )
    SetDataVector(jv + "|", se_vw + ".TotalEmployment", tot_emp, )
    SetDataVector(jv + "|", se_vw + ".Density", density, )
    SetDataVector(jv + "|", se_vw + ".AreaType", areatype, )
    SetDataVector(jv + "|", se_vw + ".EmpDensity", emp_density, )

    views.se_vw = se_vw
    views.jv = jv
    views.taz_lyr = taz_lyr
    RunMacro("Smooth Area Type", Args, map, views)
    RunMacro("Tag Highway with Area Type", Args, map, views)

    CloseView(jv)
    CloseView(se_vw)
    CloseMap(map)
EndMacro

/*
Uses buffers to smooth the boundaries between the different area types.
*/

Macro "Smooth Area Type" (Args, map, views)
    
    taz_dbd = Args.TAZGeography
    area_tbl = Args.AreaTypes
    se_vw = views.se_vw
    jv = views.jv
    taz_lyr = views.taz_lyr

    // This smoothing operation uses Enclosed inclusion
    if GetSelectInclusion() = "Intersecting" then do
        reset_inclusion = TRUE
        SetSelectInclusion("Enclosed")
    end

    // Loop over the area types in reverse order (e.g. Urban to Rural)
    // Skip the last (least dense) area type (usually "Rural") as those do
    // not require buffering.
    for t = area_tbl.length to 2 step -1 do
        type = area_tbl[t].AreaType
        buffer = area_tbl[t].Buffer

        // Select TAZs of current type
        SetView(jv)
        query = "Select * where " + se_vw + ".AreaType = '" + type + "'"
        n = SelectByQuery("selection", "Several", query)

        if n > 0 then do
            // Create a temporary buffer (deleted at end of macro)
            // and add to map.
            a_path = SplitPath(taz_dbd)
            bufferDBD = a_path[1] + a_path[2] + "ATbuffer.dbd"
            CreateBuffers(bufferDBD, "buffer", {"selection"}, "Value", {buffer},)
            bLyr = AddLayer(map,"buffer",bufferDBD,"buffer")

            // Select zones within the 1 mile buffer that have not already
            // been smoothed.
            SetLayer(taz_lyr)
            n2 = SelectByVicinity("in_buffer", "several", "buffer|", , )
            qry = "Select * where ATSmoothed = 1"
            n2 = SelectByQuery("in_buffer", "Less", qry)

            if n2 > 0 then do
            // Set those zones' area type to the current type and mark
            // them as smoothed
            opts = null
            opts.Constant = type
            v_atype = Vector(n2, "String", opts)
            opts = null
            opts.Constant = 1
            v_smoothed = Vector(n2, "Long", opts)
            SetDataVector(
                jv + "|in_buffer", se_vw + "." + "AreaType", v_atype,
            )
            SetDataVector(
                jv + "|in_buffer", se_vw + "." + "ATSmoothed", v_smoothed,
            )
            end

            DropLayer(map, bLyr)
            DeleteDatabase(bufferDBD)
        end
    end

    if reset_inclusion then SetSelectInclusion("Intersecting")
EndMacro

/*
Tags highway links with the area type of the TAZ they are nearest to.
*/

Macro "Tag Highway with Area Type" (Args, map, views)

    hwy_dbd = Args.HighwayDatabase
    area_tbl = Args.AreaTypes
    se_vw = views.se_vw
    jv = views.jv
    taz_lyr = views.taz_lyr

    // This smoothing operation uses intersecting inclusion.
    // This prevents links inbetween urban and surban from remaining rural.
    if GetSelectInclusion() = "Enclosed" then do
        reset_inclusion = "true"
        SetSelectInclusion("Intersecting")
    end

    // Add highway links to map and add AreaType field
    hwy_dbd = hwy_dbd
    {nLayer, llyr} = GetDBLayers(hwy_dbd)
    llyr = AddLayer(map, llyr, hwy_dbd, llyr)
    a_fields = {{"AreaType", "Character", 10, }}
    RunMacro("Add Fields", {view: llyr, a_fields: a_fields})
    SetLayer(llyr)
    SelectByQuery("primary", "several", "Select * where DTWB contains 'D' or DTWB contains 'T'")

    // Loop over each area type starting with most dense.  Skip the first.
    // All remaining links after this loop will be tagged with the lowest
    // area type. Secondary links (walk network) not tagged.
    for t = area_tbl.length to 2 step -1 do
        type = area_tbl[t].AreaType

        // Select TAZs of current type
        SetView(jv)
        query = "Select * where " + se_vw + ".AreaType = '" + type + "'"
        n = SelectByQuery("selection", "Several", query)

        if n > 0 then do
            // Create buffer and add it to the map
            buffer_dbd = GetTempFileName(".dbd")
            opts = null
            opts.Exterior = "Merged"
            opts.Interior = "Merged"
            CreateBuffers(buffer_dbd, "buffer", {"selection"}, "Value", {100/5280}, )
            bLyr = AddLayer(map, "buffer", buffer_dbd, "buffer")

            // Select links within the buffer that haven't been updated already
            SetLayer(llyr)
            n2 = SelectByVicinity(
                "links", "several", taz_lyr + "|selection", 0, 
                {"Source And": "primary"}
            )
            query = "Select * where AreaType <> null"
            n2 = SelectByQuery("links", "Less", query)

            // Remove buffer from map
            DropLayer(map, bLyr)

            if n2 > 0 then do
                // For these links, update their area type
                v_at = Vector(n2, "String", {{"Constant", type}})
                SetDataVector(llyr + "|links", "AreaType", v_at, )
            end
        end
    end

    // Select all remaining links and assign them to the
    // first (lowest density) area type.
    SetLayer(llyr)
    query = "Select * where AreaType = null and (DTWB contains 'D' or DTWB contains 'T')"
    n = SelectByQuery("links", "Several", query)
    if n > 0 then do
        type = area_tbl[1].AreaType
        v_at = Vector(n, "String", {{"Constant", type}})
        SetDataVector(llyr + "|links", "AreaType", v_at, )
    end

    // If this script modified the user setting for inclusion, change it back.
    if reset_inclusion = "true" then SetSelectInclusion("Enclosed")
EndMacro

/*
DC models for residents and visitors need the cluster information on the se
data table.
*/

Macro "Add Cluster Info to SE Data" (Args)

    se_file = Args.DemographicOutputs
    taz_file = Args.TAZGeography

    se = CreateObject("Table", se_file)
    se.AddField({FieldName: "Cluster", type: "integer"})
    se.AddField({FieldName: "ClusterName", type: "string"})
    taz = CreateObject("Table", taz_file)
    join = se.Join({
        Table: taz,
        LeftFields: "TAZ",
        RightFields: "TAZID"
    })
    join.Cluster = join.District7
    join.ClusterName = "c" + String(join.District7)
endmacro

/*
The visitor model uses the Cluster field from the se data, but
collapses district 3 into 1 due to a lack of observations in the
survey.
*/

Macro "Create Visitor Clusters" (Args)
    se_file = Args.DemographicOutputs
    se = CreateObject("Table", se_file)
    se.AddField({FieldName: "VisCluster", type: "integer"})
    se.AddField({FieldName: "VisClusterName", type: "string"})
    se.VisCluster = if se.Cluster = 3
        then 1
        else se.Cluster
    se.VisClusterName = "c" + String(se.VisCluster)
endmacro

macro "Speeds and Capacities" (Args, Result)
    ret_value = 1
    // input data files
    LineDB = Args.HighwayDatabase
    Line = CreateObject("Table", {FileName: LineDB, LayerType: "Line"})
    Node = CreateObject("Table", {FileName: LineDB, LayerType: "Node"})
    fields = {
    {FieldName: "ABSpeedLimit"}, 
    {FieldName: "BASpeedLimit"}, 
    {FieldName: "ABFreeFlowSpeed"}, 
    {FieldName: "BAFreeFlowSpeed"}, 
    {FieldName: "ABFreeFlowTime"}, 
    {FieldName: "BAFreeFlowTime"}, 
    {FieldName: "ABAMTime"}, 
    {FieldName: "BAAMTime"}, 
    {FieldName: "ABPMTime"}, 
    {FieldName: "BAPMTime"}, 
    {FieldName: "ABOPTime"}, 
    {FieldName: "BAOPTime"}, 
    {FieldName: "cap_phpl"}, 
    {FieldName: "ABHourlyCapacityAM"}, 
    {FieldName: "BAHourlyCapacityAM"},  
    {FieldName: "ABHourlyCapacityPM"}, 
    {FieldName: "BAHourlyCapacityPM"},  
    {FieldName: "ABHourlyCapacityOP"}, 
    {FieldName: "BAHourlyCapacityOP"},  
    {FieldName: "ABCapacityAM"}, 
    {FieldName: "BACapacityAM"},  
    {FieldName: "ABCapacityPM"}, 
    {FieldName: "BACapacityPM"},  
    {FieldName: "ABCapacityOP"}, 
    {FieldName: "BACapacityOP"},  
    {FieldName: "ABAlpha"}, 
    {FieldName: "BAAlpha"}, 
    {FieldName: "ABBeta"}, 
    {FieldName: "BABeta"},
    {FieldName: "Mode"}
   }
    Line.AddFields({Fields: fields})

    Line.HCMMedian = if Line.HCMType = "MLHighway" or Line.HCMType = "Freeway"
        then "Restrictive"
        else Line.HCMMedian

    SpeedCap = Args.SpeedCapacityLookup
    SC = CreateObject("Table", SpeedCap)
    join = Line.Join({Table: SC, LeftFields: {"HCMType", "AreaType", "HCMMedian"}, RightFields: {"HCMType", "AreaType", "HCMMedian"}})
    join.cap_phpl = join.cape_phpl
    periods = {"AM", "PM", "OP"}
    for period in periods do
        // hourly capacities
        join.("ABHourlyCapacity" + period) = if join.("AB_LANE" + period) <> null and join.Dir >= 0 
            then join.cape_phpl * join.("AB_LANE" + period) 
            else if join.("AB_LANE" + period) = null and join.Dir = -1 
                then null 
                else join.cape_phpl
        join.("BAHourlyCapacity" + period) = if join.("BA_LANE" + period) <> null and join.Dir <= 0 
            then join.cape_phpl * join.("BA_LANE" + period) 
            else if join.("BA_LANE" + period) = null and join.Dir = 1 
                then null 
                else join.cape_phpl

        // period capacities
        join.("ABCapacity" + period) = join.("ABHourlyCapacity" + period) * Args.(period + "CapFactor")
        join.("BACapacity" + period) = join.("BAHourlyCapacity" + period) * Args.(period + "CapFactor")
    end
    join.ABAlpha = join.Alpha
    join.BAAlpha = join.Alpha
    join.ABBeta = join.Beta
    join.BABeta = join.Beta
    join.ABSpeedLimit = join.PostedSpeed
    join.BASpeedLimit = join.PostedSpeed
    join.ABFreeFlowSpeed = join.PostedSpeed + join.ModifyPosted
    join.BAFreeFlowSpeed = join.PostedSpeed + join.ModifyPosted
    join.ABFreeFlowTime = join.Length / join.ABFreeFlowSpeed * 60   
    join.BAFreeFlowTime = join.Length / join.BAFreeFlowSpeed * 60   
    join.ABAMTime = join.Length / join.ABFreeFlowSpeed * 60   
    join.BAAMTime = join.Length / join.BAFreeFlowSpeed * 60   
    join.ABPMTime = join.Length / join.ABFreeFlowSpeed * 60   
    join.BAPMTime = join.Length / join.BAFreeFlowSpeed * 60   
    join.ABOPTime = join.Length / join.ABFreeFlowSpeed * 60   
    join.BAOPTime = join.Length / join.BAFreeFlowSpeed * 60   
    join.Mode = 1
    join = null

    // Use congested times if the file exists
    CongestedTimeFile = Args.CongestedTimeFile
    if GetFileInfo(CongestedTimeFile) <> null then do
        targetfields = {"ABAMTime", "BAAMTime", "ABPMTime", "BAPMTime", "ABOPTime", "BAOPTime"}
        sourcefields = {"ABAMCongestedTime", "BAAMCongestedTime", "ABPMCongestedTime", "BAPMCongestedTime", "ABOPCongestedTime", "BAOPCongestedTime"}

        ConSpd = CreateObject("Table", CongestedTimeFile)
        jvcg = Line.Join({Table: ConSpd, LeftFields: "ID", RightFields: "ID"})
        for i = 1 to sourcefields.length do
            jvcg.(targetfields[i]) = if jvcg.(sourcefields[i]) = null 
                then jvcg.(targetfields[i]) 
                else jvcg.(sourcefields[i])
        end
        jvcg = null
    end


    quit:
    Return(ret_value)


endmacro

/*

*/

macro "CalculateTransitSpeeds Oahu" (Args, Result)
    ret_value = 1
    // input data files
    LineDB = Args.HighwayDatabase
    RouteSystem = Args.TransitRoutes
    
    Line = CreateObject("Table", {FileName: LineDB, LayerType: "Line"})
    Node = CreateObject("Table", {FileName: LineDB, LayerType: "Node"})
    fields = {
    {FieldName: "ABWalkTime"}, 
    {FieldName: "BAWalkTime"}, 
    {FieldName: "ABBikeTime"}, 
    {FieldName: "BABikeTime"}, 
    {FieldName: "ABTransitFactor"}, 
    {FieldName: "BATransitFactor"}, 
    // {FieldName: "ABTransitSpeed"}, 
    // {FieldName: "BATransitSpeed"}, 
    // {FieldName: "ABTransitTime"}, 
    // {FieldName: "BATransitTime"}, 
    // {FieldName: "ABTransitSpeedAM"}, 
    // {FieldName: "BATransitSpeedAM"}, 
    {FieldName: "ABTransitTimeAM"}, 
    {FieldName: "BATransitTimeAM"}, 
    // {FieldName: "ABTransitSpeedPM"}, 
    // {FieldName: "BATransitSpeedPM"}, 
    {FieldName: "ABTransitTimePM"}, 
    {FieldName: "BATransitTimePM"}, 
    // {FieldName: "ABTransitSpeedOP"}, 
    // {FieldName: "BATransitSpeedOP"}, 
    {FieldName: "ABTransitTimeOP"}, 
    {FieldName: "BATransitTimeOP"} 
   }
    Line.AddFields({Fields: fields})

    // Bike/Walk times
    Line.WalkSpeed = if Line.WalkSpeed = null then Args.WalkSpeed else Line.WalkSpeed
    Line.BikeSpeed = if Line.BikeSpeed = null then Args.BikeSpeed else Line.BikeSpeed
    Line.ABWalkTime = Line.Length / Line.WalkSpeed * 60
    Line.BAWalkTime = Line.Length / Line.WalkSpeed * 60
    Line.ABBikeTime = Line.Length / Line.BikeSpeed * 60
    Line.BABikeTime = Line.Length / Line.BikeSpeed * 60

    // Transit times
    SpeedCap = Args.SpeedCapacityLookup
    SC = CreateObject("Table", SpeedCap)
    join = Line.Join({Table: SC, LeftFields: {"HCMType", "AreaType", "HCMMedian"}, RightFields: {"HCMType", "AreaType", "HCMMedian"}})
    join.ABTransitFactor = join.TransitFactor
    join.BATransitFactor = join.TransitFactor
    periods = {"AM", "PM", "OP"}
    dirs = {"AB", "BA"}
    for period in periods do
        for dir in dirs do
            join.(dir + "TransitTime" + period) = join.(dir + period + "Time") * join.(dir + "TransitFactor")
            // Some transit routes run on links without times in a given period
            // (e.g. hov links). In these cases, use the free flow time.
            join.(dir + "TransitTime" + period) = if join.(dir + "TransitTime" + period) = null
                then join.ABFreeFlowTime * join.(dir + "TransitFactor")
                else join.(dir + "TransitTime" + period)
        end
    end

    
    // Calculate for non-drivable links (e.g. transit only)
    join.SelectByQuery({
        SetName: "transit_only",
        Query: "T = 1 and D = 0"
    })
    for period in periods do
        for dir in dirs do
            posted_speed = if join.PostedSpeed = null
                then 25
                else join.PostedSpeed
            // join.(dir + "TransitSpeed" + period) = posted_speed
            join.(dir + "TransitTime" + period) = join.Length / posted_speed * 60
            // Fill in the non-period fields just to avoid confusion when
            // looking at the link layer.
            // join.(dir + "TransitSpeed") = posted_speed
            // join.(dir + "TransitTime") = join.Length / posted_speed * 60
        end
    end

    RS = CreateObject("Table", {FileName: RouteSystem, LayerType: "Route"})
    RouteLayer = RS.GetView()
    STOP = CreateObject("Table", {FileName: RouteSystem, LayerType: "Stop"})
    fields = {
    {FieldName: "NodeID", Type: "integer"}}
    STOP.AddFields({Fields: fields})
    n = TagRouteStopsWithNode(RouteLayer, , "NodeID", 10) // fill stop layer with node ID for each stop
    RS = null
    STOP = null

       quit:
    Return(ret_value)

endmacro

/*

*/

macro "BuildNetworks Oahu" (Args, Result)

    RunMacro("BuildHighwayNetwork Oahu", Args)
    RunMacro("Check Highway Network", Args)
    if RunMacro("MT Districts Exist?", Args) then do
        RunMacro("Create Microtransit Access Matrix", Args)
    end
    RunMacro("Create Transit Networks", Args)
    return(1)
EndMacro

/*

*/

macro "BuildHighwayNetwork Oahu" (Args)

    periods = {"AM", "PM", "OP"}
    out_dir = Args.[Output Folder] + "/skims"

    ret_value = 1
    LineDB = Args.HighwayDatabase
    // TurnPenaltyFile = Args.TurnPenalties
    
    for period in periods do
        netfile = out_dir + "/highwaynet_" + period + ".net"
    
        netObj = CreateObject("Network.Create")
        netObj.LayerDB = LineDB
        netObj.Filter =  "D = 1 and (nz(AB_LANE" + period + ") + nz(BA_LANE" + period + ") > 0)" 
        netObj.AddLinkField({Name: "FreeFlowTime", Field: {"ABFreeFlowTime", "BAFreeFlowTime"}})
        netObj.AddLinkField({Name: "Time", Field: {"AB" + period + "Time", "BA" + period + "Time"}})
        netObj.AddLinkField({Name: "Capacity", Field: {"ABCapacity" + period, "BACapacity" + period}})
        netObj.AddLinkField({Name: "WalkTime", Field: {"ABWalkTime", "BAWalkTime"}})
        netObj.AddLinkField({Name: "BikeTime", Field: {"ABBikeTime", "BABikeTime"}})
        netObj.AddLinkField({Name: "Alpha", Field: {"ABAlpha", "BAAlpha"}})
        netObj.AddLinkField({Name: "Beta", Field: {"ABBeta", "BABeta"}})
        netObj.AddLinkField({Name: "TollCostSOV", Field: {"TollCostSOV" + period, "TollCostSOV" + period}})
        netObj.AddLinkField({Name: "TollCostHOV", Field: {"TollCostHOV" + period, "TollCostHOV" + period}})
        netObj.OutNetworkName = netfile
        netObj.Run()
        
        netSetObj = null
        netSetObj = CreateObject("Network.Settings")
        netSetObj.LayerDB = LineDB
        netSetObj.LoadNetwork(netfile)
        netSetObj.CentroidFilter = "Centroid = 1"
        // netSetObj.SetPenalties({LinkPenaltyTable: TurnPenaltyFile, PenaltyField: "Penalty"})
        netSetObj.SetPenalties({UTurn: -1})
        netSetObj.Run()

    end

    quit:
    Return(ret_value)

endmacro

/*
Runs a test assignment for each period using a dummy matrix to check that all 
links have proper capacity/speed values.
*/

Macro "Check Highway Network" (Args)

    if Args.Iteration > 1 then return()

    out_dir = Args.[Output Folder]
    skim_dir = out_dir + "/skims"
    se_file = Args.DemographicOutputs
    hwy_dbd = Args.HighwayDatabase
    periods = {"AM", "PM", "OP"}

    se = CreateObject("Table", se_file)
    mtx_file = GetTempFileName(".mtx")
    mh = CreateMatrixFromView("temp", se.GetView() + "|", "TAZ", "TAZ", {"TAZ"}, {"File Name": mtx_file})
    mtx = CreateObject("Matrix", mh)
    mtx.AddCores({"SOV"})
    mtx.SOV := mtx.TAZ
    mtx = null

    obj = CreateObject("Network.Assignment")
    obj.LayerDB = hwy_dbd
    obj.ResetClasses()
    obj.Iterations = 1
    obj.Convergence = .01
    obj.DemandMatrix ({MatrixFile: mtx_file})
    obj.AddClass({Demand: "SOV"})
    obj.FlowTable = GetRandFileName("*.bin")
    for period in periods do
        obj.Network = skim_dir + "/highwaynet_" + period + ".net"
        obj.DelayFunction = {Function: "bpr.vdf", Fields : {"FreeFlowTime",
            "Capacity", "Alpha", "Beta", "None"}}
        ret_value = obj.Run()
        results = obj.GetResults()
    end
endmacro

Macro "Create Transit Networks" (Args)

    rsFile = Args.TransitRoutes
    Periods = {"AM", "PM", "OP"}
    AccessModes = Args.AccessModes
    skim_dir = Args.OutputSkims

    // Remove mt access modes if MT districts aren't defined
    if !RunMacro("MT Districts Exist?", Args)
        then AccessModes = ExcludeArrayElements(AccessModes, AccessModes.position("mt"), 1)

    objLyrs = CreateObject("AddRSLayers", {FileName: rsFile})
    rtLyr = objLyrs.RouteLayer

    // Retag stops to nodes. While this step is done by the route manager
    // during scenario creation, a user might create a new route to test after
    // creating the scenario. This makes sure it 'just works'.
    TagRouteStopsWithNode(rtLyr,, "Node_ID", 0.2)

    for period in Periods do
        for acceMode in AccessModes do
            tnwFile = skim_dir + "\\transit\\" + period + "_" + acceMode + ".tnw"
            o = CreateObject("Network.CreateTransit")
            o.LayerRS = rsFile
            o.OutNetworkName = tnwFile
            o.UseModes({TransitModeField: "Mode", NonTransitModeField: "Mode"})

            // route attributes
            o.RouteFilter = period + "Headway > 0 & Mode > 0"
            o.AddRouteField({Name: period + "Headway", Field: period + "Headway"})
            o.AddRouteField({Name: "Fare", Field: "Fare"})

            // stop attributes
            o.StopToNodeTagField = "Node_ID"
            o.AddStopField({Name: "dwell_on", Field: "dwell_on"})
            o.AddStopField({Name: "dwell_off", Field: "dwell_off"})

            // link attributes
            o.AddLinkField({Name: "bus_time", TransitFields: {"ABTransitTime" + period, "BATransitTime" + period},
                                            NonTransitFields: {"ABWalkTime", "BAWalkTime"}})
            // rail time never changes so just use AM
            o.AddLinkField({Name: "rail_time", TransitFields: {"ABTransitTimeAM", "ABTransitTimeAM"},
                                            NonTransitFields: {"ABWalkTime", "BAWalkTime"}})

            // drive attributes
            o.IncludeDriveLinks = true
            o.DriveLinkFilter = "D = 1"
            o.AddLinkField({Name: "DriveTime",
                            TransitFields:    {"AB" + period + "Time", "BA" + period + "Time"}, 
                            NonTransitFields: {"AB" + period + "Time", "BA" + period + "Time"}})

            // walk attributes
            o.IncludeWalkLinks = true
            o.WalkLinkFilter = "W = 1"
            o.AddLinkField({Name:             "WalkTime",
                            TransitFields:    {"ABWalkTime", "BAWalkTime"}, 
                            NonTransitFields: {"ABWalkTime", "BAWalkTime"}})

            ok = o.Run()

            RunMacro("Set Transit Network", Args, period, acceMode,)
          end // for acceMode
      end // for period
endMacro

Macro "Set Transit Network" (Args, period, acceMode, currTransMode)
    rsFile = Args.TransitRoutes
    modeTable = Args.TransitModeTable
    skim_dir = Args.OutputSkims
    tnwFile = skim_dir + "\\transit\\" + period + "_" + acceMode + ".tnw"

    // If this is microtransit access, open the parking/access matrix file
    if acceMode = "mt" then do
        mt_access_mtx = Args.("MTAccessMatrix" + period)
        mt_park = CreateObject("Matrix", mt_access_mtx)
    end

    o = CreateObject("Network.SetPublicPathFinder", {RS: rsFile, NetworkName: tnwFile})

    // define user classes
    UserClasses = null
    ModeUseFld = null
    DrvTimeFld = null
    DrvInUse   = null
    PermitAllW = null
    AllowWacc = null
    ParkFilter = null

    // build class name list and class-specific PnR/KnR option array
    transit_modes = RunMacro("Get Transit Net Def Col Names", modeTable)
    if acceMode = "w" 
        then TransModes = transit_modes
        // if "pnr", "knr", "mt" remove 'all'
        else TransModes = ExcludeArrayElements(transit_modes, transit_modes.position("all"), 1)

    for transMode in TransModes do
        UserClasses = UserClasses + {period + "-" + acceMode + "-" + transMode}
        ModeUseFld = ModeUseFld + {transMode}
        PermitAllW = PermitAllW + {false}

        if acceMode = "w" then do
            DrvTimeFld = DrvTimeFld + {}
            DrvInUse = DrvInUse + {false}
            AllowWacc = AllowWacc + {true}
            ParkFilter = ParkFilter + {}
        end else do
            DrvTimeFld = DrvTimeFld + {"DriveTime"}
            DrvInUse = DrvInUse + {true}
            AllowWacc = AllowWacc + {false}

            if acceMode = "knr" 
                then ParkFilter = ParkFilter + {"KNR = 1"}
            if acceMode = "pnr" 
                then ParkFilter = ParkFilter + {"PNR = 1"}
            if acceMode = "mt" then do
                // ParkFilter = ParkFilter + {"MTDist <> null"}
                ParkTimeMatrix = ParkTimeMatrix + {mt_park.TotalTime}
                ParkCostMatrix = ParkCostMatrix + {mt_park.Fare}
                ParkDistanceMatrix = ParkDistanceMatrix + {mt_park.Distance}
            end
        end // else (if acceMode)
    end // for transMode

    o.UserClasses = UserClasses

    o.DriveTime = DrvTimeFld
    DrvOpts = null
    DrvOpts.InUse = DrvInUse
    DrvOpts.PermitAllWalk = PermitAllW
    DrvOpts.AllowWalkAccess = AllowWacc
    DrvOpts.ParkingNodes = ParkFilter
    DrvOpts.ParkTimeMatrix = ParkTimeMatrix
    DrvOpts.ParkCostMatrix = ParkCostMatrix
    DrvOpts.ParkDistanceMatrix = ParkDistanceMatrix
    if period = "PM" then
        o.DriveEgress(DrvOpts)
    else
        o.DriveAccess(DrvOpts)

    o.CentroidFilter = "Centroid = 1"
    o.LinkImpedance = "bus_time" // default

    o.Parameters(
        {MaxTripCost : 240,
         MaxTransfers: 2,
         VOT         : 0.1984 // $/min (40% of the median wage)
        })

    o.AccessControl(
        {PermitWalkOnly:     false,
         MaxWalkAccessPaths: 10
        })

    o.Combination(
        {CombinationFactor: .1
        })
    o.StopTimeFields({
        InitialPenalty: null,
        //TransferPenalty: "xfer_pen",
        DwellOn: "dwell_on",
        DwellOff: "dwell_off"
    })
    o.TimeGlobals(
        {Headway:         14,
         InitialPenalty:  0,
         TransferPenalty: 5,
         MaxInitialWait:  30,
         MaxTransferWait: 10,
         MinInitialWait:  2,
         MinTransferWait: 5,
         Layover:         5, 
         MaxAccessWalk:   45,
         MaxEgressWalk:   45,
         MaxModalTotal:   120
        })

    o.RouteTimeFields(
        {Headway: period + "Headway"
        })

    o.ModeTable(
        {TableName: modeTable,
        // A field in the mode table that contains a list of
        // link network field names. These network field names
        // in turn point to the AB/BA fields on the link layer.
         TimeByMode:          "IVTT",
         ModesUsedField:      ModeUseFld,
         OnlyCombineSameMode: "true"
        })
    
    o.ModeTimeFields({
        DwellOn: "DwellOn",
        DwellOff: "DwellOff"
    })

    o.RouteWeights(
        {
         Fare: null,
         Time: null,
         InitialPenalty: null,
         TransferPenalty: null,
         InitialWait: null,
         TransferWeight: null,
         Dwelling: null
        })

    o.GlobalWeights(
        {Fare:            1.0,
         Time:            1.0,
         InitialPenalty:  1.0,
         TransferPenalty: 3.0,
         InitialWait:     3.0,
         TransferWait:    3.0,
         Dwelling:        2.0,
         WalkTimeFactor:  3.0,
         DriveTimeFactor: 1.0
        })

    o.Fare(
        {Type:              "Flat",
         RouteFareField:    "Fare",
        //  RouteXFareField:   "XferFare",
         FareValue:         0.0, // default value
         TransferFareValue: 0.0 // defaul value
        })

    if currTransMode <> null then
        o.CurrentClass = period + "-" + acceMode + "-" + currTransMode 

    ok = o.Run()
    if !ok then goto quit

  quit:
    return(ok)
endMacro

/*
The microtransit network use an origin-to-parking matrix to restrict
access to only those nodes within the same MT district. This creates that
matrix.
*/

Macro "Create Microtransit Access Matrix" (Args)

    LineDB = Args.HighwayDatabase
    net_dir = Args.[Output Folder] + "/skims"
    periods = {"AM", "PM", "OP"}

    map = CreateObject("Map", LineDB)
    {nlyr, llyr} = map.GetLayerNames()
    
    se = CreateObject("Table", Args.DemographicOutputs)
    centroid_mt_ids = se.MTDist
    se = null
    nodes = CreateObject("Table", nlyr)

    for period in periods do
        netfile = net_dir + "/highwaynet_" + period + ".net"

        // MT nodes may not be reachable in all periods due to lane changes
        // by time of day. Avoid errors by marking which MT nodes are reachable
        // in the current period. This is then used to filter the origin-to-parking
        // matrix.
        nodes.AddField("in_network_" + period)
        nh = ReadNetwork(netfile)
        SetLayer(nlyr)
        SelectFromNetwork("mt_in_net", "several", nh, )
        nodes.ChangeSet("mt_in_net")
        nodes.("in_network_" + period) = 1
        nodes.SelectByQuery({
            SetName: "mt",
            Query: "MTDist <> null and in_network_" + period + " = 1"
        })
        node_mt_ids = nodes.MTDist
        node_mt_fare = nodes.MTFare
        node_mt_headway = nodes.MTHeadway

        out_file = Args.("MTAccessMatrix" + period)
        skimvar = "Time"
        obj = CreateObject("Network.Skims")
        obj.LoadNetwork (netfile)
        obj.LayerDB = LineDB
        obj.Origins ="Centroid <> null"
        obj.Destinations = "MTDist <> null and in_network_" + period + " = 1"
        obj.Minimize = skimvar
        obj.AddSkimField({"Length", "All"})
        obj.OutputMatrix({MatrixFile: out_file, Matrix: "Microtransit Access Matrix"})
        ok = obj.Run()
        m = CreateObject("Matrix", out_file)
        m.RenameCores({CurrentNames: {"Length (Skim)"}, NewNames: {"Distance"}})
        m.AddCores({"Fare", "Headway", "TotalTime", "OrigDist", "DestDist", "IntraDist"})
        m.Fare := node_mt_fare
        m.Headway := node_mt_headway
        m.TotalTime := m.Time + m.Headway
        centroid_mt_ids.rowbased = "false"
        m.OrigDist := centroid_mt_ids
        m.DestDist := node_mt_ids
        m.IntraDist := if m.OrigDist = m.DestDist then 1 else 0
        m.IntraDist := if m.OrigDist = null then null else m.IntraDist
        m.IntraDist := if m.DestDist = null then null else m.IntraDist
        core_names = m.GetCoreNames()
        for core in core_names do
            m.(core) := m.(core) * m.IntraDist
        end
        
        // Transpose PM matrix. The result of the above skim is a drive accesss
        // matrix, but the PM network is set to drive egress. 
        if period = "PM" then do
            t_file = Substitute(out_file, ".mtx", "_transposed.mtx", )
            t = m.Transpose({OutputFile: t_file})
            t = null
            m = null
            obj = null
            DeleteFile(out_file)
            RenameFile(t_file, out_file)
        end
    end
endmacro
