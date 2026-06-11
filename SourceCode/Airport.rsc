/*
Called by flowchart
*/

Macro "Airport Model" (Args)
    Args.ABMFlag = 0
    RunMacro("Airport Generation", Args)
    RunMacro("Airport TOD", Args)
    RunMacro("Airport Gravity", Args)
    RunMacro("Airport ModeChoice", Args)
    RunMacro("Airport Directionality and Occupancy", Args)
    return(1)
endmacro

/*
Airport generation. Productions are specified as inputs in the AirportTrips
field in the SE data.
*/

Macro "Airport Generation" (Args)

    se_file = Args.DemographicOutputs
    rate_file = Args.[Airport Attraction Rates]
    vis_ratio = Args.[Airport Vis Ratio]

    se = CreateObject("Table", se_file)
    se_vw = se.GetView()

    // Productions
    se.AddField("air_vis_p")
    se.air_vis_p = se.AirportTrips * vis_ratio
    se.AddField("air_res_p")
    se.air_res_p = se.AirportTrips * (1 - vis_ratio)

    // Attractions    
    {drive, folder, name, ext} = SplitPath(rate_file)
    RunMacro("Create Sum Product Fields", {
        view: se_vw, factor_file: rate_file,
        field_desc: "Airport Attractions|See " + name + ext + " for details."
    })
EndMacro

/*
Split airport productions and attractions into time periods
*/

Macro "Airport TOD" (Args)

    se_file = Args.DemographicOutputs
    rate_file = Args.[Airport TOD Rates]

    se = CreateObject("Table", se_file)
    se_vw = se.GetView()

    {drive, folder, name, ext} = SplitPath(rate_file)
    RunMacro("Create Sum Product Fields", {
        view: se_vw, factor_file: rate_file,
        field_desc: "Airport Productions and Attractions by Time of Day|See " + name + ext + " for details."
    })

endmacro


/*
Prepares arguments for the "Gravity" macro in the utils.rsc library.
*/

Macro "Airport Gravity" (Args)

    se_file = Args.DemographicOutputs
    periods = {"AM", "PM", "OP"}
    param_dir = Args.[Input Folder] + "/airport"
    out_dir = Args.[Output Folder]
    air_dir = out_dir + "/airport"
    RunMacro("Create Directory", air_dir)
    
    for period in periods do
        RunMacro("Gravity", {
            se_file: se_file,
            skim_file: out_dir + "/skims/HighwaySkim" + period + ".mtx",
            param_file: param_dir + "/air_gravity_" + period + ".csv",
            output_matrix: air_dir + "/air_gravity_pa_" + period + ".mtx"
        })
    end
EndMacro


/*
Airport mode choice model
Same model for residents and visitors.
Apply by time period using period-specific skims
One output matrix for each time period that contains res and vis PA trips split by mode (air_res_auto, air_res_bus etc.)
*/

Macro "Airport ModeChoice" (Args)

    se_file = Args.DemographicOutputs
    periods = {"AM", "PM", "OP"}
    param_dir = Args.[Input Folder] + "/airport"
    out_dir = Args.[Output Folder]
    air_dir = out_dir + "/airport"
    tazGeo = Args.TAZGeography
    tazBin = Substitute(tazGeo, ".dbd", ".bin",)
    RunMacro("Create Directory", air_dir)

    // Remove Rail modes if scenario does not have rail
    activeTransitModes = RunMacro("Get Active Transit Modes", Args)
    railPresent = RunMacro("Is value in array", activeTransitModes, "Rail")
    util = RunMacro("Filter Mode Utility Spec", {Args: Args, util: Args.AirportMCUtility})
    
    for period in periods do
        RunMacro("AirportMC", {
            period: period,
            se_file: tazBin,
            mc_utility: util.Utility,
            rail_present: railPresent,
            bus_skim_file: out_dir + "/skims/transit/" + period + "_w_bus.mtx",
            rail_skim_file: out_dir + "/skims/transit/" + period + "_w_rail.mtx",
            input_matrix: air_dir + "/air_gravity_pa_" + period + ".mtx",
            output_matrix: air_dir + "/air_mc_pa_" + period + ".mtx"
        })
    end
EndMacro

/*

*/

Macro "Airport Directionality and Occupancy" (Args)
    
    out_dir = Args.[Output Folder]
    air_dir = out_dir + "/airport"
    periods = {"AM", "PM", "OP"}
    vis_occ = Args.[Airport Vis Occupancy]
    res_occ = Args.[Airport Res Occupancy]

    for period in periods do
        pa_mtx_file = air_dir + "/air_mc_pa_" + period + ".mtx"
        od_mtx_file = air_dir + "/air_trips_od_veh_" + period + ".mtx"

        pa_mtx = CreateObject("Matrix", pa_mtx_file)
        od_mtx = pa_mtx.Transpose({OutputFile: od_mtx_file})
        for core_name in pa_mtx.GetCoreNames() do
            od_mtx.(core_name) := od_mtx.(core_name) * .5 + pa_mtx.(core_name) * .5

            if Lower(core_name) contains "bus" or Lower(core_name) contains "rail" or Lower(core_name) contains "shuttle" then
                occ = 1.0
            else if Lower(core_name) contains "_vis" then
                occ = vis_occ
            else 
                occ = res_occ
            od_mtx.(core_name) := od_mtx.(core_name) / occ
        end
    end
endmacro


Macro "AirportMC"(spec)

    period = spec.period
    output_matrix = spec.output_matrix

    // Run MC model and get probability matrix
    pth = SplitPath(output_matrix)
    mdlfile = pth[1] + pth[2] + "/air_mc_" + period + ".mdl"
    probFile = pth[1] + pth[2] + "/air_mc_prob_" + period + ".mtx"
    tag = "Airport_MC_" + period
    
    obj = CreateObject("PMEChoiceModel", {ModelName: tag})
    obj.OutputModelFile = mdlfile
    obj.AddTableSource({SourceName: "SEData", File: spec.se_file, IDField: "TAZID"})
    obj.AddMatrixSource({SourceName: "BusSkim", File: spec.bus_skim_file})
    if spec.rail_present then
        obj.AddMatrixSource({SourceName: "RailSkim", File: spec.rail_skim_file})
    obj.AddPrimarySpec({Name: "BusSkim"})
    obj.AddUtility({UtilityFunction: spec.mc_utility})
    obj.AddOutputSpec({Probability: probFile})
    obj.MarketSegment = period
    ret = obj.Evaluate()
    if !ret then
        Throw("Running " + tag + " MC model failed.")
    //Args.(tag + " Spec") = CopyArray(ret)
    obj = null

    // Create output file. Multiply probability by airport gravity model trips (res and vis) to get mode-specific trip estimates
    CopyFile(spec.input_matrix, output_matrix)
    m = CreateObject("Matrix", output_matrix)
    purps = m.GetCoreNames()
    mP = CreateObject("Matrix", probFile)
    modes = mP.GetCoreNames()
    for purp in purps do
        for mode in modes do
            m.AddCores({purp + "_" + mode})
            m.(purp + "_" + mode) := m.(purp) * mP.(mode)
        end
    end
    m.DropCores(purps)
    mP = null
    m.Rename("Visitor PA MC trips")
    m = null
endMacro
