
Macro "Transit Assignment" (Args)
    Args.ABMFlag = 0
    RunMacro("GenerateTransitOD", Args)
    RunMacro("PTAssign", Args)
    return(1)
endmacro

/*

*/
Macro "GenerateTransitOD" (Args)
    ret_value = 1
    periods = {"AM", "PM", "OP"}
    accessModes = Args.AccessModes
    // Remove mt access modes if MT districts aren't defined
    if !RunMacro("MT Districts Exist?", Args)
        then ExcludeArrayElements(accessModes, accessModes.position("mt"), 1)
        
    transit_modes = RunMacro("Get Transit Net Def Col Names", Args.TransitModeTable)
    transit_modes = ExcludeArrayElements(transit_modes, transit_modes.position("all"), 1)
    
    // Create output cores
    cores = null
    for access in accessModes do
        for mode in transit_modes do
            for per in periods do
                core = per + "_" + access + "_" + mode + "_Trips"   // e.g. AM_pnr_bus_Trips
                cores = cores + {core}
            end
            cores = cores + {"DAY_" + access + "_" + mode + "_Trips"}    // e.g. DAY_pnr_bus_Trips
        end
    end
    
    transitod = Args.Transit_OD
    for i = 1 to periods.length do
        per = periods[i]
        odmtx = Args.(per + "_OD")
        mODT = CreateObject("Matrix", odmtx)
        mODT.SetRowIndex("Rows")
        mODT.SetColIndex("Columns")
        mODT_cores = mODT.GetCoreNames()
        if i = 1 then do
            o = CreateObject("Matrix", {Empty: TRUE})
            mOut = o.CloneMatrixStructure({MatrixLabel: "TransitTrips", CloneSource: mODT.w_bus, MatrixFile: transitod, Matrices: cores })
            mo = CreateObject("Matrix", mOut)
        end

        for access in accessModes do
            for mode in transit_modes do
                acc_mode = access + "_" + mode
                if mODT_cores.position(acc_mode) = 0 then continue
                mc = mODT.(acc_mode)
                mo.(per + "_" + acc_mode + "_Trips") := mc
                mo.("DAY_" + acc_mode + "_Trips") := nz(mo.("DAY_" + acc_mode + "_Trips")) + nz(mc)
            end
        end
    end
    quit:
    Return(ret_value)
endmacro

/*

*/
macro "PTAssign" (Args)
    ret_value = 1
    LineDB = Args.HighwayDatabase
    RouteSystem = Args.TransitRoutes
    net_dir = Args.[Output Folder] + "/skims/transit"
    assn_dir = Args.[Output Folder] + "/Assignment/Transit"
    RunMacro("Create Directory", assn_dir)

    // Get the main transit modes from the mode table. Exclude the all-transit network from assignment.
    transit_modes = RunMacro("Get Transit Net Def Col Names", Args.TransitModeTable)
    transit_modes = ExcludeArrayElements(transit_modes, transit_modes.position("all"), 1)
    access_modes = Args.AccessModes
    // Remove mt access modes if MT districts aren't defined
    if !RunMacro("MT Districts Exist?", Args)
        then access_modes = ExcludeArrayElements(access_modes, access_modes.position("mt"), 1)
    periods = {"AM", "PM", "OP"}

    for period in periods do
        for access in access_modes do
            tnet_file = net_dir + "/" + period + "_" + access + ".tnw"
            for transit_mode in transit_modes do
                obj = null
                obj = CreateObject("Network.PublicTransportAssignment", {RS: RouteSystem, NetworkName: tnet_file})
                obj.ODLayerType = "Node"
                obj.Method = "PFE"
                obj.Iterations = 1
                obj.FlowTable = assn_dir + "/" + access + "_" + transit_mode + "_" + period + "_flows.bin"
                obj.WalkFlowTable = assn_dir + "/" + access + "_" + transit_mode + "_" + period + "_walkflows.bin"
                obj.OnOffTable = assn_dir + "/" + access + "_" + transit_mode + "_" + period + "_onoff.bin"
                obj.TransitLinkFlowsTable = assn_dir + "/" + access + "_" + transit_mode + "_" + period + "_agg.bin"
                class_name = period + "-" + access + "-" + transit_mode
                mopts = {MatrixFile: Args.Transit_OD, Matrix: period + "_" + access + "_" + transit_mode + "_Trips"}
                obj.AddDemandMatrix({Class: class_name, Matrix: mopts})
                ok = obj.Run()
                results = obj.GetResults()
            end
        end
    end

    quit:
    Return(ret_value)
endmacro

