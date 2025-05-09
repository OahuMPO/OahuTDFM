/*
Called by flowchart
*/

Macro "Commercial Vehicles" (Args)
    Args.ABMFlag = 0
    RunMacro("CV Productions/Attractions", Args)
    RunMacro("CV TOD", Args)
    RunMacro("CV Gravity", Args)
    return(1)
endmacro

/*
CV productions
Attractions are the same as productions
*/

Macro "CV Productions/Attractions" (Args)

    se_file = Args.DemographicOutputs
    rate_file = Args.[CV Trip Rates]

    se_vw = OpenTable("se", "FFB", {se_file})
    {drive, folder, name, ext} = SplitPath(rate_file)
    RunMacro("Create Sum Product Fields", {
        view: se_vw, factor_file: rate_file,
        field_desc: "CV Productions and Attractions|See " + name + ext + " for details."
    })

    CloseView(se_vw)
EndMacro

/*
Split CV productions and attractions into time periods
*/

Macro "CV TOD" (Args)

    se_file = Args.DemographicOutputs
    rate_file = Args.[CV TOD Rates]

    se_vw = OpenTable("se", "FFB", {se_file})
    {drive, folder, name, ext} = SplitPath(rate_file)
    RunMacro("Create Sum Product Fields", {
        view: se_vw, factor_file: rate_file,
        field_desc: "CV Productions and Attractions by Time of Day|See " + name + ext + " for details."
    })

    CloseView(se_vw)
endmacro

/*
Prepares arguments for the "Gravity" macro in the utils.rsc library.
*/

Macro "CV Gravity" (Args)

    out_dir = Args.[Output Folder]
    cv_dir = Args.[Input Folder] + "/cv"
    periods = {"AM", "PM", "OP"}

    for period in periods do

        RunMacro("Gravity", {
            se_file: Args.DemographicOutputs,
            skim_file: out_dir + "/skims/HighwaySkim" + period + ".mtx",
            param_file: cv_dir + "/cv_gravity_" + period + ".csv",
            output_matrix: out_dir + "/cv/cv_gravity_" + period + ".mtx"
        })
    end
EndMacro