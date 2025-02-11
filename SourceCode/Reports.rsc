Macro "Reports" (Args)
    RunMacro("Load Link Layer", Args)
    RunMacro("Calculate Daily Fields", Args)
    RunMacro("Create Count Difference Map", Args)
    RunMacro("VOC Maps", Args)
    RunMacro("Speed Maps", Args)
    RunMacro("Count PRMSEs", Args)
    RunMacro("Summarize Links", Args)
    RunMacro("Transit Summary", Args)
    return(1)
endmacro

/*
This loads the final assignment results onto the link layer.
*/

Macro "Load Link Layer" (Args)

    hwy_dbd = Args.HighwayDatabase
    assn_dir = Args.[Output Folder] + "\\assignment\\roadway\\"
    periods = {"AM", "PM", "OP"}

    {nlyr, llyr} = GetDBLayers(hwy_dbd)

    for period in periods do
        assn_file = assn_dir + "\\" + period + "Flows.bin"

        temp = CreateObject("Table", assn_file)
        field_names = temp.GetFieldNames()
        temp = null
        RunMacro("Join Table To Layer", hwy_dbd, "ID", assn_file, "ID1")
        // {map, {nlyr, llyr}} = RunMacro("Create Map", {file: hwy_dbd})
        tbl = CreateObject("Table", {FileName: hwy_dbd, Layer: llyr})
        for field_name in field_names do
            if field_name = "ID1" then continue
            // Remove the field if it already exists before renaming
            tbl.DropFields(field_name + "_" + period)
            tbl.RenameField({FieldName: field_name, NewName: field_name + "_" + period})
        end

        // Calculate delay by time period and direction
        a_dirs = {"AB", "BA"}
        for dir in a_dirs do

            // Add delay field
            delay_field = dir + "_Delay_" + period
            tbl.AddField({
                FieldName: delay_field,
                Description: "Hours of Delay|(CongTime - FFTime) * Flow / 60"
            })

            // Get data vectors
            v_fft = nz(tbl.(dir + "FreeFlowTime"))
            v_ct = nz(tbl.(dir + "_Time_" + period))
            v_vol = nz(tbl.(dir + "_Flow_" + period))

            // Calculate delay
            v_delay = (v_ct - v_fft) * v_vol / 60
            v_delay = max(v_delay, 0)
            tbl.(delay_field) = v_delay
        end
        tbl = null
    end
endmacro

/*
This macro summarize fields across time period and direction.

The loaded network table will have a volume field for each class that looks like
"AB_Flow_auto_AM". It will also have fields aggregated across classes that look
like "BA_Flow_PM". Direction (AB/BA) and time period (e.g. AM)
will be looped over. Create an array of the rest of the field names to
summarize. e.g. {"Flow_auto", "Flow", "VMT"}.
*/

Macro "Calculate Daily Fields" (Args)

  a_periods = {"AM", "PM", "OP"}
  loaded_dbd = Args.HighwayDatabase
  a_dir = {"AB", "BA"}
  modes = {"drivealone", "carpool", "LTRK", "MTRK", "HTRK"}

  // Add link layer to workspace
  {nlyr, llyr} = GetDBLayers(loaded_dbd)
  llyr = AddLayerToWorkspace(llyr, loaded_dbd, llyr)

  // Calculate non-additive daily fields
  fields_to_add = {
    {"AB_Speed_Daily", "Real", 10, 2,,,, "Slowest speed throughout day"},
    {"BA_Speed_Daily", "Real", 10, 2,,,, "Slowest speed throughout day"},
    {"AB_Time_Daily", "Real", 10, 2,,,, "Highest time throughout day"},
    {"BA_Time_Daily", "Real", 10, 2,,,, "Highest time throughout day"},
    {"AB_VOC_Daily", "Real", 10, 2,,,, "Highest v/c throughout day"},
    {"BA_VOC_Daily", "Real", 10, 2,,,, "Highest v/c throughout day"}
  }
  RunMacro("Add Fields", {view: llyr, a_fields: fields_to_add})
  fields_to_add = null

  for d = 1 to a_dir.length do
    dir = a_dir[d]

    v_min_speed = GetDataVector(llyr + "|", dir + "_Speed_Daily", )
    v_min_speed = if (v_min_speed = null) then 9999 else v_min_speed
    v_max_time = GetDataVector(llyr + "|", dir + "_Time_Daily", )
    v_max_time = if (v_max_time = null) then 0 else v_max_time
    v_max_voc = nz(GetDataVector(llyr + "|", dir + "_VOC_Daily", ))

    for p = 1 to a_periods.length do
      period = a_periods[p]

      v_speed = GetDataVector(llyr + "|", dir + "_Speed_" + period, )
      v_time = GetDataVector(llyr + "|", dir + "_Time_" + period, )
      v_voc = GetDataVector(llyr + "|", dir + "_VOC_" + period, )

      v_min_speed = min(v_min_speed, v_speed)
      v_max_time = max(v_max_time, v_time)
      v_max_voc = max(v_max_voc, v_voc)
    end

    SetDataVector(llyr + "|", dir + "_Speed_Daily", v_min_speed, )
    SetDataVector(llyr + "|", dir + "_Time_Daily", v_max_time, )
    SetDataVector(llyr + "|", dir + "_VOC_Daily", v_max_voc, )
  end

  // Sum up the flow fields
  for mode in modes do

    for dir in a_dir do
      out_field = dir + "_" + mode + "_Flow_Daily"
      fields_to_add = fields_to_add + {{out_field, "Real", 10, 2,,,,"Daily " + dir + " " + mode + " Flow"}}
      v_output = null

      // For this direction and mode, sum every period
      for period in a_periods do
        input_field = dir + "_Flow_" + mode + "_" + period
        v_add = GetDataVector(llyr + "|", input_field, )
        v_output = nz(v_output) + nz(v_add)
      end

      output.(out_field) = v_output
      output.(dir + "_Flow_Daily") = nz(output.(dir + "_Flow_Daily")) + v_output
      output.Total_Flow_Daily = nz(output.Total_Flow_Daily) + v_output
    end

    output.("Total_" + mode + "_Flow_Daily") = output.("AB_" + mode + "_Flow_Daily") + output.("BA_" + mode + "_Flow_Daily")
  end
  fields_to_add = fields_to_add + {
    {"AB_Flow_Daily", "Real", 10, 2,,,,"AB Daily Flow"},
    {"BA_Flow_Daily", "Real", 10, 2,,,,"BA Daily Flow"},
    {"Total_Flow_Daily", "Real", 10, 2,,,,"Daily Flow in both direction"}
  }
  for mode in modes do
    fields_to_add = fields_to_add + {
        {"Total_" + mode + "_Flow_Daily", "Real", 10, 2,,,,"Daily " + mode + " Flow in both direction"}
    }
  end

  // Other fields to sum
  a_fields = {"VMT", "VHT", "Delay"}
  for field in a_fields do
    for dir in a_dir do
      v_output = null
      out_field = dir + "_" + field + "_Daily"
      fields_to_add = fields_to_add + {{out_field, "Real", 10, 2,,,,"Daily " + dir + " " + field}}
      for period in a_periods do
        input_field = dir + "_" + field + "_" + period
        v_add = GetDataVector(llyr + "|", input_field, )
        v_output = nz(v_output) + nz(v_add)
      end
      output.(out_field) = v_output
      output.("Total_" + field + "_Daily") = nz(output.("Total_" + field + "_Daily")) + v_output
    end

    description = "Daily " + field + " in both directions"
    if field = "Delay" then description = description + " (hours)"
    fields_to_add = fields_to_add + {{"Total_" + field + "_Daily", "Real", 10, 2,,,, description}}
  end

  // The assignment files don't have total delay by period. Create those.
  for period in a_periods do
    out_field = "Tot_Delay_" + period
    fields_to_add = fields_to_add + {{out_field, "Real", 10, 2,,,, period + " Total Delay"}}
    {v_ab, v_ba} = GetDataVectors(llyr + "|", {"AB_Delay_" + period, "BA_Delay_" + period}, )
    v_output = nz(v_ab) + nz(v_ba)
    output.(out_field) = v_output
  end

  RunMacro("Add Fields", {view: llyr, a_fields: fields_to_add})
  SetDataVectors(llyr + "|", output, )
  DropLayerFromWorkspace(llyr)
EndMacro


/*
Create maps that compare model volumes to counts.
Produced for all scenarios, but only a valid comparison for
the base year scenario.
*/

Macro "Create Count Difference Map" (Args)
  
  output_dir = Args.[Output Folder]
  hwy_dbd = Args.HighwayDatabase
  map_dir = output_dir + "/_reports/maps"
  RunMacro("Create Directory", map_dir)

  // Create total count diff map
  opts = null
  opts.output_file = map_dir + "/Count Difference - Total.map"
  opts.hwy_dbd = hwy_dbd
  opts.count_id_field = "CountID"
  opts.count_field = "DailyCount"
  opts.vol_field = "Total_Flow_Daily"
  opts.field_suffix = "All"
  RunMacro("Count Difference Map", opts)
EndMacro


/*
Creates V/C maps for each time period
*/

Macro "VOC Maps" (Args)

  hwy_dbd = Args.HighwayDatabase
  periods = {"AM", "PM", "OP", "Daily"}
  output_dir = Args.[Output Folder] + "/_reports/maps"
  RunMacro("Create Directory", output_dir)
  
  levels = {"E"}

  // The first set of colors are the traditional green-to-red ramp. The second
  // set of colors are yellow-to-blue, which is color-blind friendly.
  a_line_colors =	{
    {
      ColorRGB(10794, 52428, 17733),
      ColorRGB(63736, 63736, 3084),
      ColorRGB(65535, 32896, 0),
      ColorRGB(65535, 0, 0)
    },
    {
      ColorRGB(65535, 65535, 54248),
      ColorRGB(41377, 56026, 46260),
      ColorRGB(16705, 46774, 50372),
      ColorRGB(8738, 24158, 43176)
    }
  }
  color_suffixes = {"GrRd", "YlBu"}

  for j = 1 to 2 do
    line_colors = a_line_colors[j]
    color_suffix = color_suffixes[j]

    for period in periods do
      for los in levels do

        mapFile = output_dir + "/voc_" + period + "_LOS" + los + "_" + color_suffix + ".map"

        //Create a new, blank map
        {map, {nlyr, llyr}} = RunMacro("Create Map", {file: hwy_dbd})
        SetLayerVisibility(map + "|" + nlyr, "false")
        SetLayer(llyr)

        // Dualized Scaled Symbol Theme
        flds = {llyr+".AB_Flow_" + period}
        opts = null
        opts.Title = period + " Flow"
        opts.[Data Source] = "All"
        opts.[Minimum Size] = 1
        opts.[Maximum Size] = 10
        theme_name = CreateContinuousTheme("Flows", flds, opts)
        // Set color to white to make it disappear in legend
        dual_colors = {ColorRGB(65535,65535,65535)}
        dual_linestyles = {LineStyle({{{2, -1, 0},{0,0,1},{0,0,-1}}})}
        dual_linesizes = {0}
        SetThemeLineStyles(theme_name , dual_linestyles)
        SetThemeLineColors(theme_name , dual_colors)
        SetThemeLineWidths(theme_name , dual_linesizes)
        ShowTheme(, theme_name)

        // Apply color theme based on the V/C
        num_classes = 4
        theme_title = if period = "Daily"
          then "Max V/C (LOS " + los + ")"
          else period + " V/C (LOS " + los + ")"
        cTheme = CreateTheme(
          theme_title, llyr+".AB_VOC_" + period, "Manual",
          num_classes,
          {
            {"Values",{
              {0.0,"True",0.6,"False"},
              {0.6,"True",0.75,"False"},
              {0.75,"True",0.9,"False"},
              {0.9,"True",100,"False"}
              }},
            {"Other", "False"}
          }
        )

        dualline = LineStyle({{{2, -1, 0},{0,0,1},{0,0,-1}}})

        for i = 1 to num_classes do
            class_id = llyr +"|" + cTheme + "|" + String(i)
            SetLineStyle(class_id, dualline)
            SetLineColor(class_id, line_colors[i])
            SetLineWidth(class_id, 2)
        end

        // Change the labels of the classes for legend
        labels = {
          "Congestion Free (VC < .6)",
          "Moderate Traffic (VC .60 to .75)",
          "Heavy Traffic (VC .75 to .90)",
          "Stop and Go (VC > .90)"
        }
        SetThemeClassLabels(cTheme, labels)
        ShowTheme(,cTheme)

        // Hide centroid connectors
        SetLayer(llyr)
        ccquery = "Select * where HCMType = 'CC'"
        n1 = SelectByQuery ("CCs", "Several", ccquery,)
        if n1 > 0 then SetDisplayStatus(llyr + "|CCs", "Invisible")

        // Configure Legend
        SetLegendDisplayStatus(llyr + "|", "False")
        RunMacro("G30 create legend", "Theme")
        subtitle = if period = "Daily"
          then "Daily Flow + Max V/C"
          else period + " Period"
        SetLegendSettings (
          GetMap(),
          {
            "Automatic",
            {0, 1, 0, 0, 1, 4, 0},
            {1, 1, 1},
            {"Arial|Bold|16", "Arial|9", "Arial|Bold|16", "Arial|12"},
            {"", subtitle}
          }
        )
        str1 = "XXXXXXXX"
        solid = FillStyle({str1, str1, str1, str1, str1, str1, str1, str1})
        SetLegendOptions (GetMap(), {{"Background Style", solid}})

        // Save map
        RedrawMap(map)
        windows = GetWindows("Map")
        window = windows[1][1]
        RestoreWindow(window)
        SaveMap(map, mapFile)
        CloseMap(map)
      end
    end
  end
EndMacro


/*
Creates a map showing speed reductions (similar to Google) for each period
*/

Macro "Speed Maps" (Args)

  hwy_dbd = Args.HighwayDatabase
  periods = {"AM", "PM", "OP"}
  output_dir = Args.[Output Folder] + "/_reports/maps"
  if GetDirectoryInfo(output_dir, "All") = null then CreateDirectory(output_dir)

  for period in periods do

    mapFile = output_dir + "/speed_" + period + ".map"

    //Create a new, blank map
    {map, {nlyr, llyr}} = RunMacro("Create Map", {file: hwy_dbd})
    SetLayerVisibility(map + "|" + nlyr, "false")
    SetLayer(llyr)

    // Dualized Scaled Symbol Theme
    flds = {llyr+".AB_Flow_" + period}
    opts = null
    opts.Title = period + " Flow"
    opts.[Data Source] = "All"
    opts.[Minimum Size] = 1
    opts.[Maximum Size] = 10
    theme_name = CreateContinuousTheme("Flows", flds, opts)
    // Set color to white to make it disappear in legend
    SetThemeLineColors(theme_name , {ColorRGB(65535,65535,65535)})
    dual_linestyles = {LineStyle({{{1, -1, 0}}})}
    SetThemeLineStyles(theme_name , dual_linestyles)
    ShowTheme(, theme_name)

    // Apply color theme based on the % speed reduction
    tbl = null
    tbl = CreateObject("Table", {View: llyr})
    dirs = {"AB", "BA"}
    for dir in dirs do
      tbl.AddField({FieldName: dir + period + "SpeedRedux", Type: "integer"})
      v_cong_speed = tbl.(dir + "_Speed_" + period)
      v_posted_speed = tbl.PostedSpeed
      tbl.(dir + period + "SpeedRedux") = min((v_cong_speed - v_posted_speed) / v_posted_speed * 100, 0)
    end
    num_classes = 5
    theme_title = period + " Speed Reduction %"
    ab_field = "AB" + period + "SpeedRedux"
    cTheme = CreateTheme(
      theme_title, llyr + "." + ab_field, "Manual",
      num_classes,
      {
        {"Values",{
          {-10,"True", 100,"True"},
          {-20,"True", -10,"False"},
          {-35,"True", -20,"False"},
          {-50,"True", -35,"False"},
          {-100,"True", -50,"False"}
          }}
      }
    )
    line_colors =	{
      ColorRGB(6682, 38550, 16705),
      ColorRGB(42662, 55769, 27242),
      ColorRGB(65535, 65535, 49087),
      ColorRGB(65021, 44718, 24929),
      ColorRGB(55255, 6425, 7196)
    }
    // dualline = LineStyle({{{2, -1, 0},{0,0,1},{0,0,-1}}})
    dualline = LineStyle({{{1, -1, 0}}})

    for i = 1 to num_classes do
        class_id = llyr +"|" + cTheme + "|" + String(i + 1) // 1 is the other class
        SetLineStyle(class_id, dualline)
        SetLineColor(class_id, line_colors[i])
        SetLineWidth(class_id, 2)
    end

    // Change the labels of the classes for legend
    labels = {
      "Other",
      "Reduction < 10%",
      "Reduction < 20%",
      "Reduction < 35%",
      "Reduction < 50%",
      "Reduction > 50%"
    }
    SetThemeClassLabels(cTheme, labels)
    ShowTheme(,cTheme)

    // Hide centroid connectors
    SetLayer(llyr)
    ccquery = "Select * where HCMType = 'CC'"
    n1 = SelectByQuery ("CCs", "Several", ccquery,)
    if n1 > 0 then SetDisplayStatus(llyr + "|CCs", "Invisible")

    // Configure Legend
    SetLegendDisplayStatus(llyr + "|", "False")
    RunMacro("G30 create legend", "Theme")
    subtitle = period + " Period"
    SetLegendSettings (
      GetMap(),
      {
        "Automatic",
        {0, 1, 0, 0, 1, 4, 0},
        {1, 1, 1},
        {"Arial|Bold|16", "Arial|9", "Arial|Bold|16", "Arial|12"},
        {"", subtitle}
      }
    )
    str1 = "XXXXXXXX"
    solid = FillStyle({str1, str1, str1, str1, str1, str1, str1, str1})
    SetLegendOptions (GetMap(), {{"Background Style", solid}})

    // Save map
    RedrawMap(map)
    windows = GetWindows("Map")
    window = windows[1][1]
    RestoreWindow(window)
    SaveMap(map, mapFile)
    CloseMap(map)
  end
EndMacro


/*
Creates tables with %RMSE and volume % diff by facility type and volume group
*/

Macro "Count PRMSEs" (Args)
  hwy_dbd = Args.HighwayDatabase

  opts.hwy_bin = Substitute(hwy_dbd, ".dbd", ".bin", )
  opts.volume_field = "Volume_All"
  opts.count_id_field = "CountID"
  opts.count_field = "Count_All"
  opts.class_field = "HCMType"
  opts.area_field = "AreaType"
  opts.median_field = "HCMMedian"
  opts.screenline_field = "Screenline"
  opts.volume_breaks = {10000, 25000, 50000, 100000}
  opts.out_dir = Args.[Output Folder] + "/_reports/roadway_tables"
  RunMacro("Roadway Count Comparison Tables", opts)

//   // Rename screenline to cutline
//   in_file = opts.out_dir + "/count_comparison_by_screenline.csv"
//   out_file = opts.out_dir + "/count_comparison_by_cutline.csv"
//   if GetFileInfo(out_file) <> null then DeleteFile(out_file)
//   RenameFile(in_file, out_file)

//   // Run it again to generate the screenline table
//   opts.screenline_field = "Screenline"
//   RunMacro("Roadway Count Comparison Tables", opts)
endmacro

/*
Summarize highway stats like VMT and VHT
*/

Macro "Summarize Links" (Args)

  hwy_dbd = Args.HighwayDatabase
  periods = {"AM", "PM", "OP", "Daily"}

  opts.hwy_dbd = hwy_dbd
  out_dir = Args.[Output Folder]
  for period in periods do
    
    if period = "Daily"
      then total = "Total"
      else total = "Tot"
    
    opts.summary_fields = {total + "_Flow_" + period, total + "_VMT_" + period, total + "_VHT_" + period, total + "_Delay_" + period}
    grouping_fields = {"AreaType"}
    for grouping_field in grouping_fields do
      opts.output_file = out_dir + "/_reports/roadway_tables/Link_Summary_by_FT_and_" + grouping_field + "_" + period + ".bin"
      opts.grouping_fields = {"HCMType", grouping_field}
      RunMacro("Link Summary", opts)
      // Calculate space-mean-speed
      tbl = CreateObject("Table", opts.output_file)
      v_vmt = tbl.("sum_" + total + "_VMT_" + period)
      v_vht = tbl.("sum_" + total + "_VHT_" + period)
      tbl.AddField({
        FieldName: "SpaceMeanSpeed",
        Description: "Space-mean speed (VMT / VHT)"
      })
      tbl.SpaceMeanSpeed = v_vmt / v_vht
      tbl = null
    end
  end
EndMacro

/*
Summarizes transit assignment.
*/

Macro "Transit Summary" (Args)
  
  out_dir  = Args.[Output Folder]
  assn_dir = out_dir + "/assignment/transit"
  
  opts = null
  RunMacro("Summarize Transit", {
    transit_asn_dir: assn_dir,
    TransModeTable: Args.TransitModeTable,
    output_dir: out_dir + "/_reports/transit",
    loaded_network: Args.HighwayDatabase,
    scen_rts: Args.TransitRoutes
  })
EndMacro