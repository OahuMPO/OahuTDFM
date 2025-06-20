/*
Called by flowchart
*/

Macro "Airport Model" (Args)
    Args.ABMFlag = 0
    RunMacro("Airport Generation", Args)
    RunMacro("Airport TOD", Args)
    RunMacro("Airport Gravity", Args)
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

*/

Macro "Airport Directionality and Occupancy" (Args)
    
    out_dir = Args.[Output Folder]
    air_dir = out_dir + "/airport"
    periods = {"AM", "PM", "OP"}
    vis_occ = Args.[Airport Vis Occupancy]
    res_occ = Args.[Airport Res Occupancy]

    for period in periods do
        pa_mtx_file = air_dir + "/air_gravity_pa_" + period + ".mtx"
        od_mtx_file = air_dir + "/air_trips_od_veh_" + period + ".mtx"

        pa_mtx = CreateObject("Matrix", pa_mtx_file)
        od_mtx = pa_mtx.Transpose({OutputFile: od_mtx_file})
        for core_name in pa_mtx.GetCoreNames() do
            od_mtx.(core_name) := od_mtx.(core_name) * .5 + pa_mtx.(core_name) * .5

            if Position(core_name, "_vis") > 0
                then occ = vis_occ
                else occ = res_occ
            od_mtx.(core_name) := od_mtx.(core_name) / occ
        end
    end
endmacro