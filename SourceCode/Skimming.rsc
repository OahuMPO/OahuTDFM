/*

*/

macro "HighwayAndTransitSkim Oahu" (Args, Result)
    Args.ABMFlag = 0
    RunMacro("HighwayNetworkSkim Oahu", Args)
    RunMacro("transit skim", Args)
    return(1)
endmacro

macro "HighwayNetworkSkim Oahu" (Args)
    ret_value = 1
    Args.ABMFlag = 0
    LineDB = Args.HighwayDatabase
    AMhwyskimfile = Args.HighwaySkimAM
    PMhwyskimfile = Args.HighwaySkimPM
    OPhwyskimfile = Args.HighwaySkimOP
    walkskimfile = Args.WalkSkim
    bikeskimfile = Args.BikeSkim
    net_dir = Args.[Output Folder] + "/skims"
    
    Line = CreateObject("Table", {FileName: LineDB, LayerType: "Line"})
    Node = CreateObject("Table", {FileName: LineDB, LayerType: "Node"})
    NodeLayer = Node.GetView()

    TAZData = Args.DemographicOutputs
    dem = CreateObject("Table", TAZData)

    periods = {"AM", "PM", "OP"}
    for period in periods do
        netfile = net_dir + "/highwaynet_" + period + ".net"
        hwyskimfile = Args.("HighwaySkim" + period)
        skimvar = "Time"
        obj = CreateObject("Network.Skims")
        obj.LoadNetwork (netfile)
        obj.LayerDB = LineDB
        obj.Origins ="Centroid <> null"
        obj.Destinations = "Centroid <> null"
        obj.Minimize = skimvar
        obj.AddSkimField({"Length", "All"})
        obj.AddSkimField({"FreeFlowTime", "All"})
        obj.AddSkimField({"TollCostSOV", "All"})
        obj.AddSkimField({"TollCostHOV", "All"})
        obj.OutputMatrix({MatrixFile: hwyskimfile, Matrix: "HighwaySkim"})
        ok = obj.Run()

        obj = null
        obj = CreateObject("Distribution.Intrazonal")
        obj.SetMatrix({MatrixFile:hwyskimfile, Matrix: skimvar})
        obj.OperationType = "Replace"
        obj.TreatMissingAsZero = false
        obj.Neighbours = 3
        obj.Factor = 0.5
        ret_value = obj.Run()
        if !ret_value then goto quit

        obj = null
        obj = CreateObject("Distribution.Intrazonal")
        obj.SetMatrix({MatrixFile:hwyskimfile, Matrix: "Length (Skim)"})
        obj.OperationType = "Replace"
        obj.TreatMissingAsZero = false
        obj.Neighbours = 3
        obj.Factor = 0.5
        ok = obj.Run()

        obj = null
        obj = CreateObject("Distribution.Intrazonal")
        obj.SetMatrix({MatrixFile:hwyskimfile, Matrix: "FreeFlowTime (Skim)"})
        obj.OperationType = "Replace"
        obj.TreatMissingAsZero = false
        obj.Neighbours = 3
        obj.Factor = 0.5
        ok = obj.Run()

        m = CreateObject("Matrix", hwyskimfile)
        m.RenameCores({CurrentNames: "Length (Skim)", NewNames: "Distance"})
        m.RenameCores({CurrentNames: "TollCostSOV (Skim)", NewNames: "TollCostSOV"})
        m.RenameCores({CurrentNames: "TollCostHOV (Skim)", NewNames: "TollCostHOV"})
        m.RenameCores({CurrentNames: "FreeFlowTime (Skim)", NewNames: "FreeFlowTime"})
        idx = m.AddIndex({IndexName: "TAZ",
                    ViewName: NodeLayer, Dimension: "Both",
                    OriginalID: "ID", NewID: "ID", Filter: "Centroid = 1"})
        idxint = m.AddIndex({IndexName: "InternalTAZ",
                    ViewName: NodeLayer, Dimension: "Both",
                    // OriginalID: "ID", NewID: "Centroid", Filter: "Centroid <> null and CentroidType = 'Internal'"})
                    OriginalID: "ID", NewID: "ID", Filter: "Centroid = 1"})

        // Fill the diagonals of the toll cores with zeroes
        v = m.GetVector({Core: "TollCostSOV", Diagonal: "Row"})
        m.SetVector({Core: "TollCostSOV", Vector: nz(v), Diagonal: 1})
        
        v = m.GetVector({Core: "TollCostHOV", Diagonal: "Row"})
        m.SetVector({Core: "TollCostHOV", Vector: nz(v), Diagonal: 1})
    end


    obj = CreateObject("Network.Skims")
    obj.LoadNetwork (netfile)
    obj.LayerDB = LineDB
    obj.Origins ="Centroid <> null"
    obj.Destinations = "Centroid <> null"
    obj.Minimize = "WalkTime"
    obj.AddSkimField({"Length", "All"})
    obj.OutputMatrix({MatrixFile: walkskimfile, Matrix: "WalkSkim"})
    ok = obj.Run()

    obj = CreateObject("Network.Skims")
    obj.LoadNetwork (netfile)
    obj.LayerDB = LineDB
    obj.Origins ="Centroid <> null"
    obj.Destinations = "Centroid <> null"
    obj.Minimize = "BikeTime"
    obj.AddSkimField({"Length", "All"})
    obj.OutputMatrix({MatrixFile: bikeskimfile, Matrix: "BikeSkim"})
    ok = obj.Run()

    obj = null
    obj = CreateObject("Distribution.Intrazonal")
    obj.SetMatrix({MatrixFile:walkskimfile, Matrix: "WalkTime"})
    obj.OperationType = "Replace"
    obj.TreatMissingAsZero = false
    obj.Neighbours = 3
    obj.Factor = 0.5
    ret_value = obj.Run()
    if !ret_value then goto quit

    obj = null
    obj = CreateObject("Distribution.Intrazonal")
    obj.SetMatrix({MatrixFile:walkskimfile, Matrix: "Length (Skim)"})
    obj.OperationType = "Replace"
    obj.TreatMissingAsZero = false
    obj.Neighbours = 3
    obj.Factor = 0.5
    ret_value = obj.Run()
    if !ret_value then goto quit

    obj = null
    obj = CreateObject("Distribution.Intrazonal")
    obj.SetMatrix({MatrixFile:bikeskimfile, Matrix: "BikeTime"})
    obj.OperationType = "Replace"
    obj.TreatMissingAsZero = false
    obj.Neighbours = 3
    obj.Factor = 0.5
    ret_value = obj.Run()
    if !ret_value then goto quit

    obj = null
    obj = CreateObject("Distribution.Intrazonal")
    obj.SetMatrix({MatrixFile:bikeskimfile, Matrix: "Length (Skim)"})
    obj.OperationType = "Replace"
    obj.TreatMissingAsZero = false
    obj.Neighbours = 3
    obj.Factor = 0.5
    ret_value = obj.Run()
    if !ret_value then goto quit

    m = CreateObject("Matrix", walkskimfile)
    m.RenameCores({CurrentNames: {"WalkTime", "Length (Skim)"}, NewNames: {"Time", "Distance"}})
    idx = m.AddIndex({IndexName: "TAZ",
                ViewName: NodeLayer, Dimension: "Both",
                OriginalID: "ID", NewID: "ID", Filter: "Centroid = 1"})
    idxint = m.AddIndex({IndexName: "InternalTAZ",
                ViewName: NodeLayer, Dimension: "Both",
                // OriginalID: "ID", NewID: "Centroid", Filter: "Centroid <> null and CentroidType = 'Internal'"})
                OriginalID: "ID", NewID: "ID", Filter: "Centroid = 1"})

    m = CreateObject("Matrix", bikeskimfile)
    m.RenameCores({CurrentNames: {"BikeTime", "Length (Skim)"}, NewNames: {"Time", "Distance"}})
    idx = m.AddIndex({IndexName: "TAZ",
                ViewName: NodeLayer, Dimension: "Both",
                OriginalID: "ID", NewID: "ID", Filter: "Centroid = 1"})
    idxint = m.AddIndex({IndexName: "InternalTAZ",
                ViewName: NodeLayer, Dimension: "Both",
                // OriginalID: "ID", NewID: "Centroid", Filter: "Centroid <> null and CentroidType = 'Internal'"})
                OriginalID: "ID", NewID: "ID", Filter: "Centroid = 1"})

    quit:
    Return(ret_value)

endmacro

/*

*/

Macro "transit skim" (Args)
    
    // if in a feedback loop and it is the second loop or higher
    if Args.Iteration > 1 then do 
        ret_value = RunMacro("CalculateTransitSpeeds Oahu", Args)
        if !ret_value then goto quit
        ret_value = RunMacro("Create Transit Networks", Args)
        if !ret_value then goto quit
    end

    periods = {"AM", "PM", "OP"}
    access_modes = Args.AccessModes
    modeTable = Args.TransitModeTable
    rsFile = Args.TransitRoutes
    skim_dir = Args.OutputSkims

    // Remove mt access modes if MT districts aren't defined
    if !RunMacro("MT Districts Exist?", Args)
        then access_modes = ExcludeArrayElements(access_modes, access_modes.position("mt"), 1)

    Line = CreateObject("Table", {FileName: Args.HighwayDatabase, LayerType: "Line"})
    Node = CreateObject("Table", {FileName: Args.HighwayDatabase, LayerType: "Node"})
    NodeLayer = Node.GetView()

    transit_modes = RunMacro("Get Transit Net Def Col Names", modeTable)

    for period in periods do
        for acceMode in access_modes do

        tnwFile = skim_dir + "\\transit\\" + period + "_" + acceMode + ".tnw"
        if acceMode = "w" 
            then TransModes = transit_modes
            // if "pnr" or "knr" or "mt" remove 'all'
            else TransModes = ExcludeArrayElements(transit_modes, transit_modes.position("all"), 1)

            for transMode in TransModes do
                label = period + " " + acceMode + " " + transMode + " Skim Matrix"
                outFile = skim_dir + "\\transit\\" + period + "_" + acceMode + "_" + transMode + ".mtx"

                ok = RunMacro("Set Transit Network", Args, period, acceMode, transMode)
                if !ok then goto quit

                // do skim
                obj = null
                obj = CreateObject("Network.PublicTransportSkims")

                obj.Network = tnwFile
                obj.LayerRS = rsFile
                obj.Method = "PF"
                obj.SkimByNodes = True
                obj.OriginFilter = "Centroid = 1"
                obj.DestinationFilter = "Centroid = 1"

                obj.SkimVariables = {"Generalized Cost", "Fare",
                                    "In-Vehicle Time",
                                    "Initial Wait Time",
                                    "Transfer Wait Time",
                                    "Initial Penalty Time",
                                    "Transfer Penalty Time",
                                    "Transfer Walk Time",
                                    "Access Walk Time",
                                    "Egress Walk Time",
                                    "Access Drive Time",
                                    "Egress Drive Time",
                                    "Dwelling Time",
                                    "Total Time",
                                    "In-Vehicle Cost",
                                    "Initial Wait Cost",
                                    "Transfer Wait Cost",
                                    "Initial Penalty Cost",
                                    "Transfer Penalty Cost",
                                    "Transfer Walk Cost",
                                    "Access Walk Cost",
                                    "Egress Walk Cost",
                                    "Access Drive Cost",
                                    "Egress Drive Cost",
                                    "Dwelling Cost",
                                    "Number of Transfers",
                                    "In-Vehicle Distance",
                                    "Access Drive Distance",
                                    "Egress Drive Distance",
                                    "Length",
                                    "DriveTime",
                                    "WalkTime"
                                    }
                obj.OutputMatrix({MatrixFile: outFile, MatrixLabel: label, Compression: True})
                if acceMode <> "w" then do
                    park_mtx = Substitute(outFile, ".mtx", "_park.mtx", )
                    if period = "PM"
                        then obj.ReportEgressParkMatrix({MatrixFile: park_mtx})
                        else obj.ReportAccessParkMatrix({MatrixFile: park_mtx})
                end
                ok = obj.Run()
                if !ok then goto quit

                m = CreateObject("Matrix", outFile)
                idx = m.AddIndex({IndexName: "TAZ",
                                    ViewName: NodeLayer, Dimension: "Both",
                                    OriginalID: "ID", NewID: "ID", Filter: "Centroid = 1"})
                idxint = m.AddIndex({IndexName: "InternalTAZ",
                                        ViewName: NodeLayer, Dimension: "Both",
                                        // OriginalID: "ID", NewID: "Centroid", Filter: "Centroid <> null and CentroidType = 'Internal'"})
                                        OriginalID: "ID", NewID: "ID", Filter: "Centroid = 1"})
            end // for transMode
        end
    end

  quit:
    return(ok)
endmacro